using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ShonOffice.Domain.Documents;

namespace ShonOffice.Infra.OpenXml;

/// <summary>
/// Already-resolved run formatting (after applying the whole inheritance
/// chain), ready to become a domain <see cref="TextRun"/>.
/// </summary>
internal readonly record struct ResolvedRunFormat(
    bool Bold,
    bool Italic,
    bool Underline,
    double? FontSizePoints,
    string? ColorHex,
    string? FontName);

/// <summary>
/// Resolves the <b>effective</b> formatting of a paragraph or a run by
/// walking Word's style chain — <c>docDefaults</c> → base style → ... →
/// paragraph style, via <c>w:basedOn</c> — before applying direct
/// formatting.
///
/// This is exactly what the original Rust reader (<c>read_docx_text</c>)
/// was missing: that code only concatenated the text of each run and
/// ignored styles and direct formatting, which is why the GUI showed
/// everything as plain text regardless of whether the .docx had bold,
/// blue, larger titles — Word shows them that way because it inherits that
/// formatting from the "Heading1" style (or similar), not because each run
/// has it written explicitly.
/// </summary>
internal sealed class StyleResolver
{
    private readonly Dictionary<string, Style> _stylesById = new(StringComparer.OrdinalIgnoreCase);
    private readonly RunPropertiesDefault? _defaultRunProperties;

    public StyleResolver(StyleDefinitionsPart? stylesPart)
    {
        var styles = stylesPart?.Styles;
        if (styles is null)
        {
            return;
        }

        foreach (var style in styles.Elements<Style>())
        {
            var id = style.StyleId?.Value;
            if (id is not null)
            {
                _stylesById[id] = style;
            }
        }

        _defaultRunProperties = styles.Elements<DocDefaults>().FirstOrDefault()?.RunPropertiesDefault;
    }

    /// <summary>Heading level (1-9) of the paragraph style, or null if it isn't a heading.</summary>
    public int? HeadingLevel(string? paragraphStyleId)
    {
        foreach (var style in Ancestors(paragraphStyleId))
        {
            var outline = style.StyleParagraphProperties?.GetFirstChild<OutlineLevel>();
            if (outline?.Val?.Value is int zeroBasedLevel && zeroBasedLevel is >= 0 and <= 8)
            {
                return zeroBasedLevel + 1;
            }

            var id = style.StyleId?.Value;
            if (id is not null
                && id.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
                && int.TryParse(id.AsSpan("Heading".Length), out var level))
            {
                return level;
            }

            if (string.Equals(id, "Title", StringComparison.OrdinalIgnoreCase))
            {
                return 1;
            }
        }

        return null;
    }

    /// <summary>Effective alignment: the paragraph's direct one if present, otherwise inherited from the style.</summary>
    public ParagraphAlignment Alignment(string? paragraphStyleId, EnumValue<JustificationValues>? directAlignment)
    {
        if (directAlignment?.Value is { } directValue)
        {
            return Convert(directValue);
        }

        foreach (var style in Ancestors(paragraphStyleId))
        {
            var val = style.StyleParagraphProperties?.GetFirstChild<Justification>()?.Val;
            if (val?.Value is { } inheritedValue)
            {
                return Convert(inheritedValue);
            }
        }

        return ParagraphAlignment.Left;
    }

    /// <summary>Effective left indentation in points: the paragraph's direct one if present, otherwise inherited from the style.</summary>
    public double LeftIndentPoints(string? paragraphStyleId, Indentation? directIndentation)
    {
        if (TwipsToPoints(directIndentation?.Left) is { } directPoints)
        {
            return directPoints;
        }

        foreach (var style in Ancestors(paragraphStyleId))
        {
            var indentation = style.StyleParagraphProperties?.GetFirstChild<Indentation>();
            if (TwipsToPoints(indentation?.Left) is { } inheritedPoints)
            {
                return inheritedPoints;
            }
        }

        return 0;
    }

    /// <summary>
    /// Effective formatting of a run, applying in order: document
    /// defaults, the paragraph style chain (root to leaf), the character
    /// style referenced by the run (if any) and, finally, the run's direct
    /// formatting — which always wins.
    /// </summary>
    public ResolvedRunFormat ResolveRunFormat(RunProperties? directFormat, string? paragraphStyleId)
    {
        var accumulator = new RunFormatAccumulator();

        accumulator.Apply(_defaultRunProperties?.GetFirstChild<RunPropertiesBaseStyle>());

        foreach (var style in Ancestors(paragraphStyleId).Reverse())
        {
            accumulator.Apply(style.StyleRunProperties);
        }

        var characterStyleId = directFormat?.GetFirstChild<RunStyle>()?.Val?.Value;
        foreach (var style in Ancestors(characterStyleId).Reverse())
        {
            accumulator.Apply(style.StyleRunProperties);
        }

        accumulator.Apply(directFormat);

        return accumulator.Result();
    }

    /// <summary>
    /// The requested style and all its ancestors via <c>w:basedOn</c>, from
    /// most specific (the style itself) to most general (the root). Stops
    /// if it detects a cycle (malformed documents).
    /// </summary>
    private IEnumerable<Style> Ancestors(string? styleId)
    {
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var current = FindStyle(styleId);

        while (current is not null)
        {
            var id = current.StyleId?.Value;
            if (id is not null && !visited.Add(id))
            {
                yield break;
            }

            yield return current;
            current = FindStyle(current.BasedOn?.Val?.Value);
        }
    }

    private Style? FindStyle(string? styleId) =>
        styleId is not null && _stylesById.TryGetValue(styleId, out var style) ? style : null;

    private static double? TwipsToPoints(StringValue? twips) =>
        twips?.Value is { } text && double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var value)
            ? value / 20.0
            : null;

    private static ParagraphAlignment Convert(JustificationValues value)
    {
        if (value == JustificationValues.Center) return ParagraphAlignment.Center;
        if (value == JustificationValues.Right) return ParagraphAlignment.Right;
        if (value == JustificationValues.End) return ParagraphAlignment.Right;
        if (value == JustificationValues.Both) return ParagraphAlignment.Justified;
        if (value == JustificationValues.Distribute) return ParagraphAlignment.Justified;
        return ParagraphAlignment.Left;
    }

    /// <summary>Accumulates run formatting by layering: each <see cref="Apply"/> call only overrides what it explicitly carries.</summary>
    private sealed class RunFormatAccumulator
    {
        private bool? _bold;
        private bool? _italic;
        private bool? _underline;
        private double? _fontSizePoints;
        private string? _colorHex;
        private string? _fontName;

        public void Apply(OpenXmlCompositeElement? properties)
        {
            if (properties is null)
            {
                return;
            }

            var bold = properties.GetFirstChild<Bold>();
            if (bold is not null)
            {
                _bold = bold.Val is null || bold.Val.Value;
            }

            var italic = properties.GetFirstChild<Italic>();
            if (italic is not null)
            {
                _italic = italic.Val is null || italic.Val.Value;
            }

            var underline = properties.GetFirstChild<Underline>();
            if (underline is not null)
            {
                _underline = underline.Val is not null && underline.Val.Value != UnderlineValues.None;
            }

            var size = properties.GetFirstChild<FontSize>();
            if (size?.Val?.Value is { } sizeText
                && double.TryParse(sizeText, NumberStyles.Any, CultureInfo.InvariantCulture, out var halfPoints))
            {
                _fontSizePoints = halfPoints / 2.0;
            }

            var color = properties.GetFirstChild<Color>();
            if (color?.Val?.Value is { } colorText && !string.Equals(colorText, "auto", StringComparison.OrdinalIgnoreCase))
            {
                _colorHex = colorText;
            }

            var font = properties.GetFirstChild<RunFonts>();
            if (font?.Ascii?.Value is { } fontName)
            {
                _fontName = fontName;
            }
        }

        public ResolvedRunFormat Result() => new(
            Bold: _bold ?? false,
            Italic: _italic ?? false,
            Underline: _underline ?? false,
            FontSizePoints: _fontSizePoints,
            ColorHex: _colorHex,
            FontName: _fontName);
    }
}
