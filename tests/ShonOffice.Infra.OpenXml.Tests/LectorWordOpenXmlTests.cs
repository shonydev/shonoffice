using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ShonOffice.Domain.Documentos;
using Xunit;

namespace ShonOffice.Infra.OpenXml.Tests;

/// <summary>
/// No usa ningún .docx de ejemplo en disco: construye uno mínimo en memoria
/// con Open XML SDK (un estilo "Heading1" con negrita/color/tamaño propios,
/// un párrafo que lo usa sin formato directo, y un párrafo normal con
/// formato directo que debe pisar al heredado) para validar que
/// <see cref="ResolutorDeEstilos"/> resuelve la herencia de estilos igual
/// que Word, no solo el formato directo de cada run.
/// </summary>
public class LectorWordOpenXmlTests
{
    [Fact]
    public void Leer_ResuelveFormatoHeredadoDeEstiloYFormatoDirecto()
    {
        var rutaArchivo = Path.Combine(Path.GetTempPath(), $"shonoffice-test-{Guid.NewGuid():N}.docx");
        try
        {
            CrearDocumentoDePrueba(rutaArchivo);

            var documento = new LectorWordOpenXml().Leer(rutaArchivo);

            Assert.NotNull(documento.ParrafosConFormato);
            var parrafos = documento.ParrafosConFormato!;
            Assert.Equal(2, parrafos.Count);

            var titulo = parrafos[0];
            Assert.Equal(1, titulo.NivelEncabezado);
            Assert.Equal(AlineacionTexto.Centro, titulo.Alineacion);
            var textoTitulo = Assert.Single(titulo.Textos);
            Assert.True(textoTitulo.Negrita); // heredado del estilo Heading1, el run no tiene formato directo
            Assert.Equal(28, textoTitulo.TamanoPunto);
            Assert.Equal("2E74B5", textoTitulo.ColorHex);

            var cuerpo = parrafos[1];
            Assert.Null(cuerpo.NivelEncabezado);
            Assert.Equal(AlineacionTexto.Izquierda, cuerpo.Alineacion);
            var textoCuerpo = Assert.Single(cuerpo.Textos);
            Assert.True(textoCuerpo.Negrita); // formato directo del run
            Assert.Equal("FF0000", textoCuerpo.ColorHex); // formato directo pisa al heredado (que era azul)
        }
        finally
        {
            File.Delete(rutaArchivo);
        }
    }

    [Fact]
    public void Leer_ConservaTablasYParrafosVaciosEnElOrdenDelDocumento()
    {
        var rutaArchivo = Path.Combine(Path.GetTempPath(), $"shonoffice-test-{Guid.NewGuid():N}.docx");
        try
        {
            CrearDocumentoConTablaYLineaEnBlanco(rutaArchivo);

            var documento = new LectorWordOpenXml().Leer(rutaArchivo);

            Assert.NotNull(documento.ElementosConFormato);
            var elementos = documento.ElementosConFormato!;

            // Párrafo, línea en blanco, tabla, párrafo: el mismo orden que en el .docx.
            Assert.Equal(4, elementos.Count);

            var introduccion = Assert.IsType<Parrafo>(elementos[0]);
            Assert.Equal("Antes de la tabla", introduccion.TextoPlano);

            var lineaEnBlanco = Assert.IsType<Parrafo>(elementos[1]);
            Assert.Equal(string.Empty, lineaEnBlanco.TextoPlano);

            var tabla = Assert.IsType<Tabla>(elementos[2]);
            Assert.Equal(2, tabla.Filas.Count);
            Assert.Equal(2, tabla.CantidadDeColumnas);
            Assert.Equal("Encabezado 1", tabla.Filas[0].Celdas[0].TextoPlano);
            Assert.Equal("Encabezado 2", tabla.Filas[0].Celdas[1].TextoPlano);
            Assert.Equal("Dato 1", tabla.Filas[1].Celdas[0].TextoPlano);
            Assert.Equal("Dato 2", tabla.Filas[1].Celdas[1].TextoPlano);

            // Solo la primera celda tiene w:shd con relleno real en el .docx de
            // prueba: el color de fondo se lee por celda desde el propio
            // documento, no se asume por posición de fila (ver MainWindow.cs).
            Assert.Equal("2E74B5", tabla.Filas[0].Celdas[0].ColorDeFondoHex);
            Assert.Null(tabla.Filas[0].Celdas[1].ColorDeFondoHex); // fill="auto": sin relleno explícito
            Assert.Null(tabla.Filas[1].Celdas[0].ColorDeFondoHex); // sin w:shd

            var cierre = Assert.IsType<Parrafo>(elementos[3]);
            Assert.Equal("Después de la tabla", cierre.TextoPlano);
        }
        finally
        {
            File.Delete(rutaArchivo);
        }
    }

    private static void CrearDocumentoConTablaYLineaEnBlanco(string rutaArchivo)
    {
        using var documento = WordprocessingDocument.Create(rutaArchivo, WordprocessingDocumentType.Document);

        var partePrincipal = documento.AddMainDocumentPart();
        partePrincipal.Document = new Document(new Body());
        var cuerpo = partePrincipal.Document.Body!;

        cuerpo.Append(new Paragraph(new Run(new Text("Antes de la tabla"))));
        cuerpo.Append(new Paragraph()); // línea en blanco deliberada

        var tabla = new Table(
            new TableRow(
                new TableCell(
                    new TableCellProperties(new Shading { Fill = "2E74B5" }),
                    new Paragraph(new Run(new Text("Encabezado 1")))),
                new TableCell(
                    new TableCellProperties(new Shading { Fill = "auto" }), // "auto" = sin relleno explícito, no es un color
                    new Paragraph(new Run(new Text("Encabezado 2"))))),
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text("Dato 1")))), // sin w:shd
                new TableCell(new Paragraph(new Run(new Text("Dato 2"))))));
        cuerpo.Append(tabla);

        cuerpo.Append(new Paragraph(new Run(new Text("Después de la tabla"))));

        partePrincipal.Document.Save();
    }

    private static void CrearDocumentoDePrueba(string rutaArchivo)
    {
        using var documento = WordprocessingDocument.Create(rutaArchivo, WordprocessingDocumentType.Document);

        var partePrincipal = documento.AddMainDocumentPart();
        partePrincipal.Document = new Document(new Body());

        var parteDeEstilos = partePrincipal.AddNewPart<StyleDefinitionsPart>();
        parteDeEstilos.Styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(new FontSize { Val = "22" }))), // 11pt por defecto, como Word
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
        parteDeEstilos.Styles.Save();

        var cuerpo = partePrincipal.Document.Body!;

        cuerpo.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new Run(new Text("Título de prueba"))));

        cuerpo.Append(new Paragraph(
            new Run(
                new RunProperties(new Bold(), new Color { Val = "FF0000" }),
                new Text("Texto en negrita y rojo"))));

        partePrincipal.Document.Save();
    }
}
