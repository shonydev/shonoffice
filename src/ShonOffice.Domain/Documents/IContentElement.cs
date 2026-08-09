namespace ShonOffice.Domain.Documents;

/// <summary>
/// Marks the types that can appear, in any order, as a top-level element of
/// the body of a <see cref="WordDocument"/>: a <see cref="Paragraph"/> or a
/// <see cref="Table"/>. Before this the model only accounted for paragraphs
/// (<c>WordDocument.FormattedParagraphs</c>), so a reading adapter that only
/// walked top-level paragraphs (<c>body.Elements&lt;Paragraph&gt;()</c>)
/// completely ignored tables: in OOXML a table (<c>w:tbl</c>) is a sibling
/// of the paragraph inside <c>w:body</c>, not a paragraph, so its rows never
/// showed up. <see cref="WordDocument.FormattedElements"/> preserves the
/// real order of the document between paragraphs and tables so the UI can
/// reconstruct it exactly as Word shows it.
/// </summary>
public interface IContentElement
{
}
