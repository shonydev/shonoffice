namespace ShonOffice.Domain.Documents;

/// <summary>
/// Horizontal alignment of a <see cref="Paragraph"/>, as defined by Word
/// (<c>w:jc</c>): unlike the plain text read by the original Rust phase,
/// this is necessary to be able to reconstruct, for example, a centered
/// title exactly as Word shows it.
/// </summary>
public enum ParagraphAlignment
{
    Left,
    Center,
    Right,
    Justified,
}
