using System.Text;
using DocumentFormat.OpenXml.Packaging;
using ShonOffice.Domain.Documents;
using ShonOffice.Domain.Ports;
using OpenXmlParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using OpenXmlTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using OpenXmlTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using OpenXmlTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using OpenXmlText = DocumentFormat.OpenXml.Wordprocessing.Text;
using OpenXmlRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using OpenXmlBody = DocumentFormat.OpenXml.Wordprocessing.Body;
using TabChar = DocumentFormat.OpenXml.Wordprocessing.TabChar;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;

namespace ShonOffice.Infra.OpenXml;

/// <summary>
/// Implements <see cref="IDocxReader"/> with Microsoft's official Open XML
/// SDK. Unlike the original Rust reader (<c>read_docx_text</c> in
/// <c>src/lib.rs</c>), which only concatenated the text of each run, this
/// adapter resolves the <b>effective</b> formatting of every paragraph —
/// bold, italic, underline, size, color, font, alignment, indentation and
/// heading level — through <see cref="StyleResolver"/>, which walks the
/// document's style chain the same way Word does internally. That's the
/// difference you see compared to real Word: titles show up large, bold
/// and colored; text respects alignment and indentation, instead of
/// showing up as unstyled plain-text lines.
///
/// It also walks <c>w:body</c> element by element (not just
/// <c>Elements&lt;Paragraph&gt;()</c>) so it doesn't miss tables
/// (<c>w:tbl</c>), which are siblings of paragraphs there, and it keeps
/// empty paragraphs instead of discarding them: in a real <c>.docx</c>
/// those empty paragraphs are usually deliberate line breaks between
/// sections (e.g. the blank line before a title), and discarding them
/// flattens that visual separation.
/// </summary>
public sealed class WordOpenXmlReader : IDocxReader
{
    public WordDocument Read(string filePath)
    {
        using var document = WordprocessingDocument.Open(filePath, isEditable: false);

        var mainPart = document.MainDocumentPart
            ?? throw new InvalidOperationException($"'{filePath}' has no main content (MainDocumentPart is null).");

        var body = mainPart.Document.Body
            ?? throw new InvalidOperationException($"'{filePath}' has no <w:body>.");

        var resolver = new StyleResolver(mainPart.StyleDefinitionsPart);

        var elements = ConvertElements(body, resolver);

        return new WordDocument(filePath, elements);
    }

    /// <summary>
    /// Walks the direct children of <c>w:body</c> in their real order,
    /// converting each top-level paragraph and table. Using
    /// <c>body.ChildElements</c> (instead of <c>Elements&lt;Paragraph&gt;()</c>,
    /// which would filter out tables) is what allows interleaving
    /// <see cref="Table"/> and <see cref="Paragraph"/> respecting the
    /// position in which they appear in the document.
    /// </summary>
    private static List<IContentElement> ConvertElements(OpenXmlBody body, StyleResolver resolver)
    {
        var elements = new List<IContentElement>();

        foreach (var child in body.ChildElements)
        {
            switch (child)
            {
                case OpenXmlParagraph xmlParagraph:
                    elements.Add(ConvertParagraph(xmlParagraph, resolver));
                    break;
                case OpenXmlTable xmlTable:
                    elements.Add(ConvertTable(xmlTable, resolver));
                    break;
                // SectionProperties and other top-level elements (section
                // breaks, revision marks, etc.) have no text content to
                // show and are intentionally ignored.
            }
        }

        return elements;
    }

    private static Table ConvertTable(OpenXmlTable xmlTable, StyleResolver resolver)
    {
        var rows = xmlTable.Elements<OpenXmlTableRow>()
            .Select(xmlRow => ConvertRow(xmlRow, resolver))
            .ToList();

        return new Table(rows);
    }

    private static Domain.Documents.TableRow ConvertRow(OpenXmlTableRow xmlRow, StyleResolver resolver)
    {
        var cells = xmlRow.Elements<OpenXmlTableCell>()
            .Select(xmlCell => ConvertCell(xmlCell, resolver))
            .ToList();

        return new Domain.Documents.TableRow(cells);
    }

    private static Domain.Documents.TableCell ConvertCell(OpenXmlTableCell xmlCell, StyleResolver resolver)
    {
        // A cell almost always has a single paragraph, but OOXML allows
        // several (e.g. real line breaks inside the cell), so all of them
        // are kept instead of concatenating them into one.
        var paragraphs = xmlCell.Elements<OpenXmlParagraph>()
            .Select(p => ConvertParagraph(p, resolver))
            .ToList();

        return new Domain.Documents.TableCell(paragraphs, CellBackgroundColor(xmlCell));
    }

    /// <summary>
    /// Real background color of the cell (<c>w:shd w:fill="..."</c> inside
    /// <c>w:tcPr</c>), if it has one. "auto" and "" are not a color, they
    /// mean "no explicit fill" — that's how Word uses them, and treating
    /// them as a hex value would accidentally paint cells black.
    /// </summary>
    private static string? CellBackgroundColor(OpenXmlTableCell xmlCell)
    {
        var fill = xmlCell.TableCellProperties?.Shading?.Fill?.Value;
        if (string.IsNullOrEmpty(fill) || string.Equals(fill, "auto", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return fill;
    }

    private static Paragraph ConvertParagraph(OpenXmlParagraph xmlParagraph, StyleResolver resolver)
    {
        var styleId = xmlParagraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

        var alignment = resolver.Alignment(styleId, xmlParagraph.ParagraphProperties?.Justification?.Val);
        var indent = resolver.LeftIndentPoints(styleId, xmlParagraph.ParagraphProperties?.Indentation);
        var headingLevel = resolver.HeadingLevel(styleId);

        var runs = new List<TextRun>();
        foreach (var run in xmlParagraph.Elements<OpenXmlRun>())
        {
            var content = TextOfRun(run);
            if (content.Length == 0)
            {
                continue;
            }

            var format = resolver.ResolveRunFormat(run.RunProperties, styleId);
            runs.Add(new TextRun(
                content,
                bold: format.Bold,
                italic: format.Italic,
                underline: format.Underline,
                fontSizePoints: format.FontSizePoints,
                colorHex: format.ColorHex,
                fontName: format.FontName));
        }

        // A paragraph with no text runs (or with empty runs) is, almost
        // always, a deliberate blank line between sections of the
        // document: it's kept as-is, with an empty run list, so the UI can
        // reserve its space instead of the line break disappearing.
        return new Paragraph(runs, alignment, headingLevel, indent, styleId);
    }

    /// <summary>
    /// Text of a run including manual tabs and line breaks (<c>w:tab</c>,
    /// <c>w:br</c>), not just the <c>w:t</c> nodes like the Rust version did.
    /// </summary>
    private static string TextOfRun(OpenXmlRun run)
    {
        var content = new StringBuilder();

        foreach (var child in run.ChildElements)
        {
            switch (child)
            {
                case OpenXmlText textNode:
                    content.Append(textNode.Text);
                    break;
                case TabChar:
                    content.Append('\t');
                    break;
                case Break:
                    content.Append('\n');
                    break;
            }
        }

        return content.ToString();
    }
}
