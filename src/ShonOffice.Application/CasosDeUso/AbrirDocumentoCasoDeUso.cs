using ShonOffice.Domain.Documentos;
using ShonOffice.Domain.Excepciones;
using ShonOffice.Domain.Puertos;

namespace ShonOffice.Application.CasosDeUso;

/// <summary>
/// Abre un documento de Office detectando su tipo por la extensión del
/// archivo y delegando la lectura al puerto correspondiente. Es el primer
/// adaptador de entrada (UI, CLI) el que llama a este caso de uso; nunca al
/// revés.
/// </summary>
public sealed class AbrirDocumentoCasoDeUso
{
    private readonly IDocxReader _lectorWord;
    private readonly IXlsxReader _lectorExcel;
    private readonly IPptxReader _lectorPowerPoint;

    public AbrirDocumentoCasoDeUso(
        IDocxReader lectorWord,
        IXlsxReader lectorExcel,
        IPptxReader lectorPowerPoint)
    {
        _lectorWord = lectorWord;
        _lectorExcel = lectorExcel;
        _lectorPowerPoint = lectorPowerPoint;
    }

    public Documento Ejecutar(string rutaArchivo)
    {
        var extension = Path.GetExtension(rutaArchivo).ToLowerInvariant();

        return extension switch
        {
            ".docx" => _lectorWord.Leer(rutaArchivo),
            ".xlsx" => _lectorExcel.Leer(rutaArchivo),
            ".pptx" => _lectorPowerPoint.Leer(rutaArchivo),
            _ => throw new FormatoNoSoportadoException(extension),
        };
    }
}
