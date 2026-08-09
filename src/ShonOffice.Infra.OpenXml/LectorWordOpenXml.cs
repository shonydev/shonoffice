using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ShonOffice.Domain.Documentos;
using ShonOffice.Domain.Puertos;

namespace ShonOffice.Infra.OpenXml;

/// <summary>
/// Implementa <see cref="IDocxReader"/> con el Open XML SDK oficial de
/// Microsoft. A diferencia de la lectura original en Rust
/// (<c>read_docx_text</c> en <c>src/lib.rs</c>), que solo concatenaba el
/// texto de cada run, este adaptador resuelve el formato <b>efectivo</b> de
/// cada párrafo — negrita, cursiva, subrayado, tamaño, color, fuente,
/// alineación, sangría y nivel de encabezado — a través de
/// <see cref="ResolutorDeEstilos"/>, que recorre la cadena de estilos del
/// documento igual que lo hace Word internamente. Esa es la diferencia que
/// se ve en la comparación con Word real: los títulos aparecen grandes,
/// negrita y en color; el texto respeta alineación y sangría, en vez de
/// mostrarse todo como líneas de texto plano sin estilo.
///
/// También recorre <c>w:body</c> elemento por elemento (no solo
/// <c>Elements&lt;Paragraph&gt;()</c>) para no perderse las tablas
/// (<c>w:tbl</c>), que son hermanas de los párrafos ahí, y conserva los
/// párrafos vacíos en vez de descartarlos: en un <c>.docx</c> real esos
/// párrafos vacíos suelen ser saltos de línea deliberados entre secciones
/// (una línea en blanco antes de un título, por ejemplo), y descartarlos
/// aplana esa separación visual.
/// </summary>
public sealed class LectorWordOpenXml : IDocxReader
{
    public DocumentoWord Leer(string rutaArchivo)
    {
        using var documento = WordprocessingDocument.Open(rutaArchivo, isEditable: false);

        var partePrincipal = documento.MainDocumentPart
            ?? throw new InvalidOperationException($"'{rutaArchivo}' no tiene contenido principal (MainDocumentPart nulo).");

        var cuerpo = partePrincipal.Document.Body
            ?? throw new InvalidOperationException($"'{rutaArchivo}' no tiene <w:body>.");

        var resolutor = new ResolutorDeEstilos(partePrincipal.StyleDefinitionsPart);

        var elementos = ConvertirElementos(cuerpo, resolutor);

        return new DocumentoWord(rutaArchivo, elementos);
    }

    /// <summary>
    /// Recorre los hijos directos de <c>w:body</c> en su orden real,
    /// convirtiendo cada párrafo y cada tabla de nivel superior. Usar
    /// <c>cuerpo.ChildElements</c> (en vez de <c>Elements&lt;Paragraph&gt;()</c>,
    /// que filtraría las tablas) es lo que permite intercalar
    /// <see cref="Tabla"/> y <see cref="Parrafo"/> respetando la posición en
    /// la que aparecen en el documento.
    /// </summary>
    private static List<IElementoDeContenido> ConvertirElementos(Body cuerpo, ResolutorDeEstilos resolutor)
    {
        var elementos = new List<IElementoDeContenido>();

        foreach (var hijo in cuerpo.ChildElements)
        {
            switch (hijo)
            {
                case Paragraph parrafoXml:
                    elementos.Add(ConvertirParrafo(parrafoXml, resolutor));
                    break;
                case Table tablaXml:
                    elementos.Add(ConvertirTabla(tablaXml, resolutor));
                    break;
                // SectionProperties y otros elementos de nivel superior (saltos de
                // sección, marcas de revisión, etc.) no tienen contenido de texto
                // que mostrar y se ignoran a propósito.
            }
        }

        return elementos;
    }

    private static Tabla ConvertirTabla(Table tablaXml, ResolutorDeEstilos resolutor)
    {
        var filas = tablaXml.Elements<TableRow>()
            .Select(filaXml => ConvertirFila(filaXml, resolutor))
            .ToList();

        return new Tabla(filas);
    }

    private static FilaTabla ConvertirFila(TableRow filaXml, ResolutorDeEstilos resolutor)
    {
        var celdas = filaXml.Elements<TableCell>()
            .Select(celdaXml => ConvertirCelda(celdaXml, resolutor))
            .ToList();

        return new FilaTabla(celdas);
    }

    private static CeldaTabla ConvertirCelda(TableCell celdaXml, ResolutorDeEstilos resolutor)
    {
        // Una celda casi siempre tiene un único párrafo, pero OOXML permite
        // varios (p. ej. saltos de línea reales dentro de la celda), así que
        // se conservan todos en vez de concatenarlos en uno solo.
        var parrafos = celdaXml.Elements<Paragraph>()
            .Select(p => ConvertirParrafo(p, resolutor))
            .ToList();

        return new CeldaTabla(parrafos, ColorDeFondoDeCelda(celdaXml));
    }

    /// <summary>
    /// Color de fondo real de la celda (<c>w:shd w:fill="..."</c> dentro de
    /// <c>w:tcPr</c>), si lo tiene. "auto" y "" no son un color, son "sin
    /// relleno explícito" — Word los usa así, y tratarlos como un hex
    /// pintaría celdas de negro por accidente.
    /// </summary>
    private static string? ColorDeFondoDeCelda(TableCell celdaXml)
    {
        var relleno = celdaXml.TableCellProperties?.Shading?.Fill?.Value;
        if (string.IsNullOrEmpty(relleno) || string.Equals(relleno, "auto", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return relleno;
    }

    private static Parrafo ConvertirParrafo(Paragraph parrafoXml, ResolutorDeEstilos resolutor)
    {
        var estiloId = parrafoXml.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

        var alineacion = resolutor.Alineacion(estiloId, parrafoXml.ParagraphProperties?.Justification?.Val);
        var sangria = resolutor.SangriaIzquierdaPuntos(estiloId, parrafoXml.ParagraphProperties?.Indentation);
        var nivelEncabezado = resolutor.NivelEncabezado(estiloId);

        var textos = new List<Texto>();
        foreach (var run in parrafoXml.Elements<Run>())
        {
            var contenido = TextoDeRun(run);
            if (contenido.Length == 0)
            {
                continue;
            }

            var formato = resolutor.ResolverFormatoDeRun(run.RunProperties, estiloId);
            textos.Add(new Texto(
                contenido,
                negrita: formato.Negrita,
                cursiva: formato.Cursiva,
                subrayado: formato.Subrayado,
                tamanoPunto: formato.TamanoPunto,
                colorHex: formato.ColorHex,
                nombreFuente: formato.NombreFuente));
        }

        // Un párrafo sin runs de texto (o con runs vacíos) es, casi siempre,
        // una línea en blanco deliberada entre secciones del documento: se
        // conserva igual, con la lista de textos vacía, para que la UI pueda
        // reservarle su espacio en vez de que el salto de línea desaparezca.
        return new Parrafo(textos, alineacion, nivelEncabezado, sangria, estiloId);
    }

    /// <summary>
    /// Texto de un run incluyendo tabs y saltos de línea manuales
    /// (<c>w:tab</c>, <c>w:br</c>), no solo los nodos <c>w:t</c> como hacía
    /// la versión en Rust.
    /// </summary>
    private static string TextoDeRun(Run run)
    {
        var contenido = new StringBuilder();

        foreach (var hijo in run.ChildElements)
        {
            switch (hijo)
            {
                case Text nodoDeTexto:
                    contenido.Append(nodoDeTexto.Text);
                    break;
                case TabChar:
                    contenido.Append('\t');
                    break;
                case Break:
                    contenido.Append('\n');
                    break;
            }
        }

        return contenido.ToString();
    }
}
