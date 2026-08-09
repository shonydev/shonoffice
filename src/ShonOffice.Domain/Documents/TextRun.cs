namespace ShonOffice.Domain.Documents;

/// <summary>
/// A Word "run": a stretch of text inside a <see cref="Paragraph"/> that
/// shares the same character formatting (bold, italic, color, size, font).
/// A paragraph almost always has more than one <see cref="TextRun"/> when,
/// for example, only one word is bold.
/// </summary>
/// <remarks>
/// Named <c>TextRun</c> rather than the more obvious <c>Text</c> on
/// purpose: <c>DocumentFormat.OpenXml.Wordprocessing</c> already defines a
/// <c>Text</c> type (the OOXML <c>w:t</c> element), and both namespaces are
/// imported side by side in <c>ShonOffice.Infra.OpenXml</c>.
/// </remarks>
public sealed class TextRun
{
    public string Content { get; }

    public bool Bold { get; }

    public bool Italic { get; }

    public bool Underline { get; }

    /// <summary>Font size in points, or null if it couldn't be resolved.</summary>
    public double? FontSizePoints { get; }

    /// <summary>Color in "RRGGBB" hex format (same as OOXML), or null if it's the default color.</summary>
    public string? ColorHex { get; }

    public string? FontName { get; }

    public TextRun(
        string content,
        bool bold = false,
        bool italic = false,
        bool underline = false,
        double? fontSizePoints = null,
        string? colorHex = null,
        string? fontName = null)
    {
        Content = content;
        Bold = bold;
        Italic = italic;
        Underline = underline;
        FontSizePoints = fontSizePoints;
        ColorHex = colorHex;
        FontName = fontName;
    }
}
