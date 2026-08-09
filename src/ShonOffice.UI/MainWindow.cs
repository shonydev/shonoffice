using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using ShonOffice.Application.UseCases;
using ShonOffice.Domain.Documents;
using ShonOffice.Infra.OpenXml;

namespace ShonOffice.UI;

/// <summary>
/// Main window. It's just another input adapter (as the README's hexagonal
/// architecture states): it calls <see cref="OpenDocumentUseCase"/> from
/// <c>ShonOffice.Application</c>, never touching Open XML SDK directly. It
/// renders <see cref="WordDocument.FormattedElements"/> — paragraphs with
/// bold, italic, underline, size, color, alignment, indentation and
/// headings, interleaved with tables — instead of just plain text, which is
/// the visual difference from the original Rust/egui GUI (<c>src/main.rs</c>),
/// which only showed <c>ui.label(paragraph)</c> for each paragraph with no
/// styling or notion of a table.
/// </summary>
public sealed class MainWindow : Window
{
    private const string DefaultTitle = "ShonOffice";

    private static readonly Color DefaultHeadingColor = Color.Parse("#2E74B5"); // typical blue of Heading1/2 in Word templates
    private static readonly Color TableBorderColor = Color.Parse("#BFBFBF"); // typical Word table border gray

    private readonly OpenDocumentUseCase _openDocument = new(
        new WordOpenXmlReader(),
        new NotImplementedExcelReader(),
        new NotImplementedPowerPointReader());

    private readonly TextBlock _filePathText = new() { VerticalAlignment = VerticalAlignment.Center, Opacity = 0.7 };
    private readonly TextBlock _errorText = new() { Foreground = Brushes.Firebrick, Margin = new Thickness(12, 0, 12, 8), IsVisible = false };
    private readonly StackPanel _content = new() { Spacing = 6, Margin = new Thickness(16) };
    private readonly StackPanel _topBar;

    public MainWindow()
    {
        Title = DefaultTitle;
        Width = 900;
        Height = 700;

        var openButton = new Button { Content = "\U0001F4C2 Open Word..." };
        openButton.Click += async (_, _) => await OpenFileAsync();

        _topBar = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 10,
            Margin = new Thickness(12, 10),
            Children = { openButton, _filePathText },
        };

        var scroll = new ScrollViewer { Content = _content };

        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,*") };
        Grid.SetRow(_topBar, 0);
        Grid.SetRow(_errorText, 1);
        Grid.SetRow(scroll, 2);
        root.Children.Add(_topBar);
        root.Children.Add(_errorText);
        root.Children.Add(scroll);

        Content = root;

        ShowInitialMessage();
    }

    private void ShowInitialMessage()
    {
        _content.Children.Clear();
        _content.Children.Add(new TextBlock
        {
            Text = "Open a .docx file to see its content.",
            Opacity = 0.7,
        });
    }

    private async Task OpenFileAsync()
    {
        _errorText.IsVisible = false;

        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Open Word document",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("Word (*.docx)") { Patterns = new[] { "*.docx" } },
            },
        });

        var file = files.FirstOrDefault();
        if (file is null)
        {
            return; // the user cancelled the dialog
        }

        var filePath = file.Path.LocalPath;

        try
        {
            var document = (WordDocument)_openDocument.Execute(filePath);
            _filePathText.Text = filePath;
            ShowDocument(document, filePath);
        }
        catch (Exception ex)
        {
            _content.Children.Clear();
            _errorText.Text = $"Could not open the document: {ex.Message}";
            _errorText.IsVisible = true;
        }
    }

    private void ShowDocument(WordDocument document, string filePath)
    {
        _content.Children.Clear();

        // Once a document is open, the "Open Word..." row and the file
        // path disappear from view: the reading area takes over the whole
        // window and the window title takes on the role of showing which
        // document is open, just like Word does.
        _topBar.IsVisible = false;
        Title = Path.GetFileName(filePath);

        var elements = document.FormattedElements;
        if (elements is null || elements.Count == 0)
        {
            // Adapter without rich formatting (e.g. a WordDocument
            // rebuilt from a PDF via ConvertPdfToWordUseCase): plain text
            // is shown, just like the Rust GUI used to do.
            foreach (var text in document.Paragraphs)
            {
                _content.Children.Add(new TextBlock { Text = text, TextWrapping = TextWrapping.Wrap });
            }

            return;
        }

        foreach (var element in elements)
        {
            Control block = element switch
            {
                Paragraph paragraph => CreateParagraphBlock(paragraph),
                Table table => CreateTableBlock(table),
                _ => new TextBlock(),
            };

            _content.Children.Add(block);
        }
    }

    private static TextBlock CreateParagraphBlock(Paragraph paragraph)
    {
        var (baseFontSize, baseBold) = HeadingSizeAndBold(paragraph.HeadingLevel);

        var block = new TextBlock
        {
            TextWrapping = TextWrapping.Wrap,
            TextAlignment = ConvertAlignment(paragraph.Alignment),
            Margin = new Thickness(PointsToPixels(paragraph.LeftIndentPoints), 0, 0, 0),
            Inlines = new InlineCollection(),
        };

        var hasVisibleText = false;

        foreach (var textRun in paragraph.Runs)
        {
            if (textRun.Content.Length == 0)
            {
                continue;
            }

            hasVisibleText |= textRun.Content.Trim().Length > 0;

            var run = new Run(textRun.Content)
            {
                FontWeight = textRun.Bold || baseBold ? FontWeight.Bold : FontWeight.Normal,
                FontStyle = textRun.Italic ? FontStyle.Italic : FontStyle.Normal,
                TextDecorations = textRun.Underline ? TextDecorations.Underline : null,
                FontSize = textRun.FontSizePoints ?? baseFontSize,
            };

            if (textRun.ColorHex is not null && TryParseColorHex(textRun.ColorHex, out var color))
            {
                run.Foreground = new SolidColorBrush(color);
            }
            else if (paragraph.HeadingLevel is not null)
            {
                run.Foreground = new SolidColorBrush(DefaultHeadingColor);
            }

            block.Inlines!.Add(run);
        }

        if (!hasVisibleText)
        {
            // Empty paragraph in the .docx: it has no text to show, but it
            // does represent a deliberate line break (e.g. the blank line
            // separating one section from the next). Without giving it a
            // minimum height, an empty TextBlock collapses to 0px and that
            // line break disappears from view, even though real Word
            // leaves a visible blank space.
            block.MinHeight = baseFontSize * 1.3;
        }

        return block;
    }

    /// <summary>
    /// Renders a <see cref="Table"/> as a bordered grid, the same way Word
    /// shows it: each cell is a <see cref="Border"/> with its own
    /// paragraph(s) inside. The background color of each cell comes from
    /// <see cref="TableCell.BackgroundColorHex"/> — the real <c>w:shd</c>
    /// carried by the <c>.docx</c> — instead of assuming by position that
    /// "the first row is the header and goes blue": that rule would paint
    /// blue with white text on tables that, in the original document, have
    /// no header row at all, just bold text at the start of each cell
    /// (which is already resolved by <see cref="CreateParagraphBlock"/>
    /// from the run's real formatting).
    /// </summary>
    private static Border CreateTableBlock(Table table)
    {
        var columnCount = table.ColumnCount;

        var grid = new Grid();
        for (var c = 0; c < columnCount; c++)
        {
            grid.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));
        }

        for (var r = 0; r < table.Rows.Count; r++)
        {
            grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

            var row = table.Rows[r];

            for (var c = 0; c < row.Cells.Count; c++)
            {
                var cell = CreateCellBlock(row.Cells[c]);
                Grid.SetRow(cell, r);
                Grid.SetColumn(cell, c);
                grid.Children.Add(cell);
            }
        }

        return new Border
        {
            BorderBrush = new SolidColorBrush(TableBorderColor),
            BorderThickness = new Thickness(1, 1, 0, 0),
            Child = grid,
            Margin = new Thickness(0, 4),
        };
    }

    private static Border CreateCellBlock(TableCell cell)
    {
        var content = new StackPanel { Spacing = 2 };

        foreach (var paragraph in cell.Paragraphs)
        {
            content.Children.Add(CreateParagraphBlock(paragraph));
        }

        IBrush background = cell.BackgroundColorHex is not null && TryParseColorHex(cell.BackgroundColorHex, out var backgroundColor)
            ? new SolidColorBrush(backgroundColor)
            : Brushes.Transparent;

        return new Border
        {
            BorderBrush = new SolidColorBrush(TableBorderColor),
            BorderThickness = new Thickness(0, 0, 1, 1),
            Background = background,
            Padding = new Thickness(8, 6),
            Child = content,
        };
    }

    /// <summary>
    /// Default size and boldness by heading level, for when the style
    /// didn't bring an explicit resolved size (some templates define it
    /// via theme instead of on the style itself). Approximates Word's
    /// default sizes for Heading1-4+ and normal text (11pt).
    /// </summary>
    private static (double FontSizePoints, bool Bold) HeadingSizeAndBold(int? level) => level switch
    {
        1 => (28.0, true),
        2 => (22.0, true),
        3 => (18.0, true),
        4 => (15.0, true),
        >= 5 => (13.0, true),
        _ => (11.0, false),
    };

    private static TextAlignment ConvertAlignment(ParagraphAlignment alignment) => alignment switch
    {
        ParagraphAlignment.Center => TextAlignment.Center,
        ParagraphAlignment.Right => TextAlignment.Right,
        ParagraphAlignment.Justified => TextAlignment.Justify,
        _ => TextAlignment.Left,
    };

    private static double PointsToPixels(double points) => points * 96.0 / 72.0;

    private static bool TryParseColorHex(string colorHex, out Color color)
    {
        try
        {
            color = Color.Parse(colorHex.StartsWith('#') ? colorHex : $"#{colorHex}");
            return true;
        }
        catch
        {
            color = default;
            return false;
        }
    }
}
