using ShonOffice.Application.CasosDeUso;
using ShonOffice.Application.Tests.Falsos;
using ShonOffice.Domain.Documentos;
using Xunit;

namespace ShonOffice.Application.Tests;

public class GuardarDocumentoCasoDeUsoTests
{
    [Fact]
    public void Ejecutar_ConDocumentoWord_LlamaAlEscritorDeWord()
    {
        var escritorWord = new EscritorWordFalso();
        var casoDeUso = new GuardarDocumentoCasoDeUso(
            escritorWord, new EscritorExcelFalso(), new EscritorPowerPointFalso());
        var documento = new DocumentoWord("informe.docx", new[] { "hola" });

        casoDeUso.Ejecutar(documento, "salida.docx");

        var escritura = Assert.Single(escritorWord.Escrituras);
        Assert.Equal("salida.docx", escritura.Ruta);
        Assert.Same(documento, escritura.Documento);
    }

    [Fact]
    public void Ejecutar_ConDocumentoExcel_LlamaAlEscritorDeExcel()
    {
        var escritorExcel = new EscritorExcelFalso();
        var casoDeUso = new GuardarDocumentoCasoDeUso(
            new EscritorWordFalso(), escritorExcel, new EscritorPowerPointFalso());
        var documento = new DocumentoExcel("planilla.xlsx", Array.Empty<Hoja>());

        casoDeUso.Ejecutar(documento, "salida.xlsx");

        Assert.Single(escritorExcel.Escrituras);
    }
}
