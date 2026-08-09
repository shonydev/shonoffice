namespace ShonOffice.Domain.Documents;

/// <summary>
/// A Word document in memory: its content, one element per top-level
/// paragraph or table. <see cref="Paragraphs"/> is plain text and matches
/// what the Rust phase already read (<c>read_docx_text</c>);
/// <see cref="FormattedParagraphs"/> is the enriched model (bold, color,
/// size, alignment, headings...) delivered by an adapter capable of
/// resolving it, such as <c>ShonOffice.Infra.OpenXml</c>; and
/// <see cref="FormattedElements"/> is the same but also interleaving the
/// <see cref="Table"/> elements in the real order they appear in the
/// document — <see cref="FormattedParagraphs"/> alone isn't enough for that
/// because, being only paragraphs, it has no way to say "a table goes here"
/// or to keep its position relative to the rest of the text. All three are
/// kept because not every adapter can produce formatting: reconstructing a
/// <see cref="WordDocument"/> from a PDF (<see cref="Ports.IPdfEngine"/>),
/// for example, only has text.
/// </summary>
public sealed class WordDocument : OfficeDocument
{
    public override DocumentType Type => DocumentType.Word;

    public IReadOnlyList<string> Paragraphs { get; }

    /// <summary>
    /// Full formatting per paragraph, when the adapter that produced this
    /// document supports it. Null when only plain text is available.
    /// Doesn't include tables: if the document has any, use
    /// <see cref="FormattedElements"/> to avoid losing them.
    /// </summary>
    public IReadOnlyList<Paragraph>? FormattedParagraphs { get; }

    /// <summary>
    /// Top-level paragraphs and tables, in the same order they appear in
    /// the document. Null when only plain text is available (same case as
    /// <see cref="FormattedParagraphs"/> being null).
    /// </summary>
    public IReadOnlyList<IContentElement>? FormattedElements { get; }

    public WordDocument(string filePath, IReadOnlyList<string> paragraphs)
        : base(filePath)
    {
        Paragraphs = paragraphs;
        FormattedParagraphs = null;
        FormattedElements = null;
    }

    public WordDocument(string filePath, IReadOnlyList<Paragraph> formattedParagraphs)
        : base(filePath)
    {
        FormattedParagraphs = formattedParagraphs;
        FormattedElements = formattedParagraphs;
        Paragraphs = formattedParagraphs.Select(p => p.PlainText).ToArray();
    }

    public WordDocument(string filePath, IReadOnlyList<IContentElement> formattedElements)
        : base(filePath)
    {
        FormattedElements = formattedElements;
        FormattedParagraphs = formattedElements.OfType<Paragraph>().ToArray();
        Paragraphs = formattedElements.Select(PlainTextOf).ToArray();
    }

    private static string PlainTextOf(IContentElement element) => element switch
    {
        Paragraph paragraph => paragraph.PlainText,
        Table table => string.Join(
            "\n",
            table.Rows.Select(row => string.Join(" | ", row.Cells.Select(c => c.PlainText)))),
        _ => string.Empty,
    };
}
