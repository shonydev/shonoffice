using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ShonOffice.Domain.Documents;
using Xunit;
using OpenXmlParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using OpenXmlTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using OpenXmlTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using OpenXmlTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using DomainParagraph = ShonOffice.Domain.Documents.Paragraph;
using DomainTable = ShonOffice.Domain.Documents.Table;

namespace ShonOffice.Infra.OpenXml.Tests;

/// <summary>
/// Doesn't use any sample .docx on disk: it builds a minimal one in memory
/// with Open XML SDK (a "Heading1" style with its own bold/color/size, a
/// paragraph that uses it with no direct formatting, and a normal paragraph
/// with direct formatting that must override the inherited one) to verify
/// that <see cref="StyleResolver"/> resolves style inheritance the same way
/// Word does, not just each run's direct formatting.
/// </summary>
public class WordOpenXmlReaderTests
{
    [Fact]
    public void Read_ResolvesFormatInheritedFromStyleAndDirectFormat()
    {
        var filePath = Path.Combine(Path.GetTempPath(), $"shonoffice-test-{Guid.NewGuid():N}.docx");
        try
        {
            CreateTestDocument(filePath);

            var document = new WordOpenXmlReader().Read(filePath);

            Assert.NotNull(document.FormattedParagraphs);
            var paragraphs = document.FormattedParagraphs!;
            Assert.Equal(2, paragraphs.Count);

            var title = paragraphs[0];
            Assert.Equal(1, title.HeadingLevel);
            Assert.Equal(ParagraphAlignment.Center, title.Alignment);
            var titleRun = Assert.Single(title.Runs);
            Assert.True(titleRun.Bold); // inherited from the Heading1 style, the run has no direct formatting
            Assert.Equal(28, titleRun.FontSizePoints);
            Assert.Equal("2E74B5", titleRun.ColorHex);

            var body = paragraphs[1];
            Assert.Null(body.HeadingLevel);
            Assert.Equal(ParagraphAlignment.Left, body.Alignment);
            var bodyRun = Assert.Single(body.Runs);
            Assert.True(bodyRun.Bold); // the run's direct formatting
            Assert.Equal("FF0000", bodyRun.ColorHex); // direct formatting overrides the inherited one (which was blue)
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Read_KeepsTablesAndEmptyParagraphsInDocumentOrder()
    {
        var filePath = Path.Combine(Path.GetTempPath(), $"shonoffice-test-{Guid.NewGuid():N}.docx");
        try
        {
            CreateDocumentWithTableAndBlankLine(filePath);

            var document = new WordOpenXmlReader().Read(filePath);

            Assert.NotNull(document.FormattedElements);
            var elements = document.FormattedElements!;

            // Paragraph, blank line, table, paragraph: the same order as in the .docx.
            Assert.Equal(4, elements.Count);

            var intro = Assert.IsType<DomainParagraph>(elements[0]);
            Assert.Equal("Before the table", intro.PlainText);

            var blankLine = Assert.IsType<DomainParagraph>(elements[1]);
            Assert.Equal(string.Empty, blankLine.PlainText);

            var table = Assert.IsType<DomainTable>(elements[2]);
            Assert.Equal(2, table.Rows.Count);
            Assert.Equal(2, table.ColumnCount);
            Assert.Equal("Header 1", table.Rows[0].Cells[0].PlainText);
            Assert.Equal("Header 2", table.Rows[0].Cells[1].PlainText);
            Assert.Equal("Data 1", table.Rows[1].Cells[0].PlainText);
            Assert.Equal("Data 2", table.Rows[1].Cells[1].PlainText);

            // Only the first cell has a w:shd with a real fill in the test
            // .docx: the background color is read per cell from the
            // document itself, it isn't assumed by row position (see
            // MainWindow.cs).
            Assert.Equal("2E74B5", table.Rows[0].Cells[0].BackgroundColorHex);
            Assert.Null(table.Rows[0].Cells[1].BackgroundColorHex); // fill="auto": no explicit fill
            Assert.Null(table.Rows[1].Cells[0].BackgroundColorHex); // no w:shd

            var closing = Assert.IsType<DomainParagraph>(elements[3]);
            Assert.Equal("After the table", closing.PlainText);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    private static void CreateDocumentWithTableAndBlankLine(string filePath)
    {
        using var document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);

        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        body.Append(new OpenXmlParagraph(new Run(new Text("Before the table"))));
        body.Append(new OpenXmlParagraph()); // deliberate blank line

        var table = new OpenXmlTable(
            new OpenXmlTableRow(
                new OpenXmlTableCell(
                    new TableCellProperties(new Shading { Fill = "2E74B5" }),
                    new OpenXmlParagraph(new Run(new Text("Header 1")))),
                new OpenXmlTableCell(
                    new TableCellProperties(new Shading { Fill = "auto" }), // "auto" = no explicit fill, not a color
                    new OpenXmlParagraph(new Run(new Text("Header 2"))))),
            new OpenXmlTableRow(
                new OpenXmlTableCell(new OpenXmlParagraph(new Run(new Text("Data 1")))), // no w:shd
                new OpenXmlTableCell(new OpenXmlParagraph(new Run(new Text("Data 2"))))));
        body.Append(table);

        body.Append(new OpenXmlParagraph(new Run(new Text("After the table"))));

        mainPart.Document.Save();
    }

    private static void CreateTestDocument(string filePath)
    {
        using var document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);

        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(new FontSize { Val = "22" }))), // 11pt by default, like Word
            new Style(
                new StyleParagraphProperties(
                    new Justification { Val = JustificationValues.Center },
                    new OutlineLevel { Val = 0 }),
                new StyleRunProperties(
                    new Bold(),
                    new Color { Val = "2E74B5" },
                    new FontSize { Val = "56" })) // 28pt
            {
                Type = StyleValues.Paragraph,
                StyleId = "Heading1",
            });
        stylesPart.Styles.Save();

        var body = mainPart.Document.Body!;

        body.Append(new OpenXmlParagraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new Run(new Text("Test title"))));

        body.Append(new OpenXmlParagraph(
            new Run(
                new RunProperties(new Bold(), new Color { Val = "FF0000" }),
                new Text("Bold and red text"))));

        mainPart.Document.Save();
    }
}
