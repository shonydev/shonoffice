using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ShonOffice.Domain.Documentos;

namespace ShonOffice.Infra.OpenXml;

/// <summary>
/// Formato de run ya resuelto (después de aplicar toda la cadena de
/// herencia), listo para convertirse en un <see cref="Texto"/> del dominio.
/// </summary>
internal readonly record struct FormatoRunResuelto(
    bool Negrita,
    bool Cursiva,
    bool Subrayado,
    double? TamanoPunto,
    string? ColorHex,
    string? NombreFuente);

/// <summary>
/// Resuelve el formato <b>efectivo</b> de un párrafo o un run recorriendo la
/// cadena de estilos de Word — <c>docDefaults</c> → estilo base → ... →
/// estilo del párrafo, vía <c>w:basedOn</c> — antes de aplicar el formato
/// directo.
///
/// Esto es justamente lo que le faltaba a la lectura original en Rust
/// (<c>read_docx_text</c>): ese código solo concatenaba el texto de cada
/// run e ignoraba estilos y formato directo, por eso la GUI mostraba todo
/// como texto plano sin importar que el .docx tuviera títulos en negrita,
/// azules y más grandes — Word los muestra así porque hereda ese formato
/// del estilo "Heading1" (o similar), no porque cada run lo tenga escrito.
/// </summary>
internal sealed class ResolutorDeEstilos
{
    private readonly Dictionary<string, Style> _estilosPorId = new(StringComparer.OrdinalIgnoreCase);
    private readonly RunPropertiesDefault? _rPrPorDefecto;

    public ResolutorDeEstilos(StyleDefinitionsPart? stylesPart)
    {
        var styles = stylesPart?.Styles;
        if (styles is null)
        {
            return;
        }

        foreach (var estilo in styles.Elements<Style>())
        {
            var id = estilo.StyleId?.Value;
            if (id is not null)
            {
                _estilosPorId[id] = estilo;
            }
        }

        _rPrPorDefecto = styles.Elements<DocDefaults>().FirstOrDefault()?.RunPropertiesDefault;
    }

    /// <summary>Nivel de encabezado (1-9) del estilo de párrafo, o null si no es un encabezado.</summary>
    public int? NivelEncabezado(string? estiloDeParrafoId)
    {
        foreach (var estilo in Ascendientes(estiloDeParrafoId))
        {
            var outline = estilo.StyleParagraphProperties?.GetFirstChild<OutlineLevel>();
            if (outline?.Val?.Value is int nivelBase0 && nivelBase0 is >= 0 and <= 8)
            {
                return nivelBase0 + 1;
            }

            var id = estilo.StyleId?.Value;
            if (id is not null
                && id.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
                && int.TryParse(id.AsSpan("Heading".Length), out var nivel))
            {
                return nivel;
            }

            if (string.Equals(id, "Title", StringComparison.OrdinalIgnoreCase))
            {
                return 1;
            }
        }

        return null;
    }

    /// <summary>Alineación efectiva: la directa del párrafo si existe, si no la heredada del estilo.</summary>
    public AlineacionTexto Alineacion(string? estiloDeParrafoId, EnumValue<JustificationValues>? alineacionDirecta)
    {
        if (alineacionDirecta?.Value is { } valorDirecto)
        {
            return Convertir(valorDirecto);
        }

        foreach (var estilo in Ascendientes(estiloDeParrafoId))
        {
            var val = estilo.StyleParagraphProperties?.GetFirstChild<Justification>()?.Val;
            if (val?.Value is { } valorHeredado)
            {
                return Convertir(valorHeredado);
            }
        }

        return AlineacionTexto.Izquierda;
    }

    /// <summary>Sangría izquierda efectiva en puntos: la directa del párrafo si existe, si no la heredada del estilo.</summary>
    public double SangriaIzquierdaPuntos(string? estiloDeParrafoId, Indentation? sangriaDirecta)
    {
        if (TwipsAPuntos(sangriaDirecta?.Left) is { } puntosDirectos)
        {
            return puntosDirectos;
        }

        foreach (var estilo in Ascendientes(estiloDeParrafoId))
        {
            var indentacion = estilo.StyleParagraphProperties?.GetFirstChild<Indentation>();
            if (TwipsAPuntos(indentacion?.Left) is { } puntosHeredados)
            {
                return puntosHeredados;
            }
        }

        return 0;
    }

    /// <summary>
    /// Formato efectivo de un run, aplicando en orden: valores por defecto
    /// del documento, cadena del estilo del párrafo (de raíz a hoja),
    /// estilo de caracter referenciado por el run (si tiene) y, por último,
    /// el formato directo del run — que siempre gana.
    /// </summary>
    public FormatoRunResuelto ResolverFormatoDeRun(RunProperties? formatoDirecto, string? estiloDeParrafoId)
    {
        var acumulador = new AcumuladorFormatoRun();

        acumulador.Aplicar(_rPrPorDefecto?.GetFirstChild<RunPropertiesBaseStyle>());

        foreach (var estilo in Ascendientes(estiloDeParrafoId).Reverse())
        {
            acumulador.Aplicar(estilo.StyleRunProperties);
        }

        var estiloDeCaracterId = formatoDirecto?.GetFirstChild<RunStyle>()?.Val?.Value;
        foreach (var estilo in Ascendientes(estiloDeCaracterId).Reverse())
        {
            acumulador.Aplicar(estilo.StyleRunProperties);
        }

        acumulador.Aplicar(formatoDirecto);

        return acumulador.Resultado();
    }

    /// <summary>
    /// Estilo pedido y todos sus ancestros vía <c>w:basedOn</c>, de más
    /// específico (el propio estilo) a más general (la raíz). Corta si
    /// detecta un ciclo (documentos malformados).
    /// </summary>
    private IEnumerable<Style> Ascendientes(string? estiloId)
    {
        var visitados = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var actual = BuscarEstilo(estiloId);

        while (actual is not null)
        {
            var id = actual.StyleId?.Value;
            if (id is not null && !visitados.Add(id))
            {
                yield break;
            }

            yield return actual;
            actual = BuscarEstilo(actual.BasedOn?.Val?.Value);
        }
    }

    private Style? BuscarEstilo(string? estiloId) =>
        estiloId is not null && _estilosPorId.TryGetValue(estiloId, out var estilo) ? estilo : null;

    private static double? TwipsAPuntos(StringValue? twips) =>
        twips?.Value is { } texto && double.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor)
            ? valor / 20.0
            : null;

    private static AlineacionTexto Convertir(JustificationValues valor)
    {
        if (valor == JustificationValues.Center) return AlineacionTexto.Centro;
        if (valor == JustificationValues.Right) return AlineacionTexto.Derecha;
        if (valor == JustificationValues.End) return AlineacionTexto.Derecha;
        if (valor == JustificationValues.Both) return AlineacionTexto.Justificado;
        if (valor == JustificationValues.Distribute) return AlineacionTexto.Justificado;
        return AlineacionTexto.Izquierda;
    }

    /// <summary>Acumula formato de run superponiendo capas: cada <see cref="Aplicar"/> solo pisa lo que trae explícito.</summary>
    private sealed class AcumuladorFormatoRun
    {
        private bool? _negrita;
        private bool? _cursiva;
        private bool? _subrayado;
        private double? _tamanoPunto;
        private string? _colorHex;
        private string? _nombreFuente;

        public void Aplicar(OpenXmlCompositeElement? propiedades)
        {
            if (propiedades is null)
            {
                return;
            }

            var negrita = propiedades.GetFirstChild<Bold>();
            if (negrita is not null)
            {
                _negrita = negrita.Val is null || negrita.Val.Value;
            }

            var cursiva = propiedades.GetFirstChild<Italic>();
            if (cursiva is not null)
            {
                _cursiva = cursiva.Val is null || cursiva.Val.Value;
            }

            var subrayado = propiedades.GetFirstChild<Underline>();
            if (subrayado is not null)
            {
                _subrayado = subrayado.Val is not null && subrayado.Val.Value != UnderlineValues.None;
            }

            var tamano = propiedades.GetFirstChild<FontSize>();
            if (tamano?.Val?.Value is { } textoTamano
                && double.TryParse(textoTamano, NumberStyles.Any, CultureInfo.InvariantCulture, out var medioPuntos))
            {
                _tamanoPunto = medioPuntos / 2.0;
            }

            var color = propiedades.GetFirstChild<Color>();
            if (color?.Val?.Value is { } textoColor && !string.Equals(textoColor, "auto", StringComparison.OrdinalIgnoreCase))
            {
                _colorHex = textoColor;
            }

            var fuente = propiedades.GetFirstChild<RunFonts>();
            if (fuente?.Ascii?.Value is { } nombreFuente)
            {
                _nombreFuente = nombreFuente;
            }
        }

        public FormatoRunResuelto Resultado() => new(
            Negrita: _negrita ?? false,
            Cursiva: _cursiva ?? false,
            Subrayado: _subrayado ?? false,
            TamanoPunto: _tamanoPunto,
            ColorHex: _colorHex,
            NombreFuente: _nombreFuente);
    }
}
