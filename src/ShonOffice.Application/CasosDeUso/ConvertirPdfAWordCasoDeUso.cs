using ShonOffice.Domain.Documentos;
using ShonOffice.Domain.Puertos;

namespace ShonOffice.Application.CasosDeUso;

/// <summary>
/// Convierte un PDF a un documento Word. La extracción pesada de texto queda
/// a cargo del motor Rust (<see cref="IPdfEngine"/>, vía FFI); este caso de
/// uso solo orquesta: extraer texto y reconstruirlo como
/// <see cref="DocumentoWord"/> antes de guardarlo con <see cref="IDocxWriter"/>.
/// </summary>
public sealed class ConvertirPdfAWordCasoDeUso
{
    private readonly IPdfEngine _motorPdf;
    private readonly IDocxWriter _escritorWord;

    public ConvertirPdfAWordCasoDeUso(IPdfEngine motorPdf, IDocxWriter escritorWord)
    {
        _motorPdf = motorPdf;
        _escritorWord = escritorWord;
    }

    public DocumentoWord Ejecutar(string rutaPdfOrigen, string rutaWordDestino)
    {
        var parrafos = _motorPdf.ExtraerTexto(rutaPdfOrigen);
        var documento = new DocumentoWord(rutaWordDestino, parrafos);

        _escritorWord.Escribir(documento, rutaWordDestino);

        return documento;
    }
}
