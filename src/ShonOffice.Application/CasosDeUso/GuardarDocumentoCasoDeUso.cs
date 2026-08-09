using ShonOffice.Domain.Documentos;
using ShonOffice.Domain.Excepciones;
using ShonOffice.Domain.Puertos;

namespace ShonOffice.Application.CasosDeUso;

/// <summary>
/// Guarda un <see cref="Documento"/> ya en memoria en disco, delegando al
/// puerto de escritura correspondiente a su tipo concreto.
/// </summary>
public sealed class GuardarDocumentoCasoDeUso
{
    private readonly IDocxWriter _escritorWord;
    private readonly IXlsxWriter _escritorExcel;
    private readonly IPptxWriter _escritorPowerPoint;

    public GuardarDocumentoCasoDeUso(
        IDocxWriter escritorWord,
        IXlsxWriter escritorExcel,
        IPptxWriter escritorPowerPoint)
    {
        _escritorWord = escritorWord;
        _escritorExcel = escritorExcel;
        _escritorPowerPoint = escritorPowerPoint;
    }

    public void Ejecutar(Documento documento, string rutaDestino)
    {
        switch (documento)
        {
            case DocumentoWord word:
                _escritorWord.Escribir(word, rutaDestino);
                break;
            case DocumentoExcel excel:
                _escritorExcel.Escribir(excel, rutaDestino);
                break;
            case DocumentoPowerPoint powerPoint:
                _escritorPowerPoint.Escribir(powerPoint, rutaDestino);
                break;
            default:
                throw new FormatoNoSoportadoException(documento.Tipo.ToString());
        }
    }
}
