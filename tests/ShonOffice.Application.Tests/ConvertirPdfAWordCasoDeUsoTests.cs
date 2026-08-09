using ShonOffice.Application.CasosDeUso;
using ShonOffice.Application.Tests.Falsos;
using Xunit;

namespace ShonOffice.Application.Tests;

public class ConvertirPdfAWordCasoDeUsoTests
{
    [Fact]
    public void Ejecutar_ExtraeTextoDelMotorPdfYLoGuardaComoWord()
    {
        var motorPdf = new MotorPdfFalso(new[] { "primer parrafo", "segundo parrafo" });
        var escritorWord = new EscritorWordFalso();
        var casoDeUso = new ConvertirPdfAWordCasoDeUso(motorPdf, escritorWord);

        var resultado = casoDeUso.Ejecutar("origen.pdf", "destino.docx");

        Assert.Equal(new[] { "primer parrafo", "segundo parrafo" }, resultado.Parrafos);
        var escritura = Assert.Single(escritorWord.Escrituras);
        Assert.Equal("destino.docx", escritura.Ruta);
        Assert.Same(resultado, escritura.Documento);
    }
}
