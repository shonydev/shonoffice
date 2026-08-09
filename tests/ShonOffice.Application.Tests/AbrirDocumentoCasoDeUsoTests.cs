using ShonOffice.Application.CasosDeUso;
using ShonOffice.Application.Tests.Falsos;
using ShonOffice.Domain.Documentos;
using ShonOffice.Domain.Excepciones;
using Xunit;

namespace ShonOffice.Application.Tests;

public class AbrirDocumentoCasoDeUsoTests
{
    private static AbrirDocumentoCasoDeUso CrearCasoDeUso() =>
        new(new LectorWordFalso(), new LectorExcelFalso(), new LectorPowerPointFalso());

    [Fact]
    public void Ejecutar_ConArchivoDocx_DevuelveDocumentoWord()
    {
        var casoDeUso = CrearCasoDeUso();

        var resultado = casoDeUso.Ejecutar("informe.docx");

        Assert.IsType<DocumentoWord>(resultado);
    }

    [Fact]
    public void Ejecutar_ConArchivoXlsx_DevuelveDocumentoExcel()
    {
        var casoDeUso = CrearCasoDeUso();

        var resultado = casoDeUso.Ejecutar("planilla.xlsx");

        Assert.IsType<DocumentoExcel>(resultado);
    }

    [Fact]
    public void Ejecutar_ConArchivoPptx_DevuelveDocumentoPowerPoint()
    {
        var casoDeUso = CrearCasoDeUso();

        var resultado = casoDeUso.Ejecutar("presentacion.pptx");

        Assert.IsType<DocumentoPowerPoint>(resultado);
    }

    [Fact]
    public void Ejecutar_ConExtensionNoSoportada_LanzaFormatoNoSoportadoException()
    {
        var casoDeUso = CrearCasoDeUso();

        Assert.Throws<FormatoNoSoportadoException>(() => casoDeUso.Ejecutar("archivo.txt"));
    }

    [Fact]
    public void Ejecutar_EsInsensibleAMayusculas()
    {
        var casoDeUso = CrearCasoDeUso();

        var resultado = casoDeUso.Ejecutar("INFORME.DOCX");

        Assert.IsType<DocumentoWord>(resultado);
    }
}
