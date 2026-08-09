namespace ShonOffice.Domain.Documents;

/// <summary>
/// A paragraph of a <see cref="WordDocument"/> with its formatting:
/// alignment, indentation, heading level (if it's a title) and the list of
/// <see cref="TextRun"/> (runs) that make it up. This is what the original
/// plain-text model was missing to be able to display itself the way Word
/// really does, instead of a list of unstyled lines.
/// </summary>
/// <remarks>
/// Named <c>Paragraph</c> here in the domain; <c>ShonOffice.Infra.OpenXml</c>
/// aliases the OOXML type (<c>DocumentFormat.OpenXml.Wordprocessing.Paragraph</c>,
/// which represents <c>w:p</c>) to avoid the name clash.
/// </remarks>
public sealed class Paragraph : IContentElement
{
    public IReadOnlyList<TextRun> Runs { get; }

    public ParagraphAlignment Alignment { get; }

    /// <summary>Heading level 1-9 (styles "Heading1".."Heading9"/"Title"), or null for normal text.</summary>
    public int? HeadingLevel { get; }

    public double LeftIndentPoints { get; }

    /// <summary>Word paragraph style id (e.g. "Heading1"), for diagnostics/debugging.</summary>
    public string? StyleName { get; }

    public Paragraph(
        IReadOnlyList<TextRun> runs,
        ParagraphAlignment alignment = ParagraphAlignment.Left,
        int? headingLevel = null,
        double leftIndentPoints = 0,
        string? styleName = null)
    {
        Runs = runs;
        Alignment = alignment;
        HeadingLevel = headingLevel;
        LeftIndentPoints = leftIndentPoints;
        StyleName = styleName;
    }

    /// <summary>Plain text of all runs concatenated, for compatibility with the simple model (<see cref="WordDocument.Paragraphs"/>).</summary>
    public string PlainText => string.Concat(Runs.Select(r => r.Content));
}
