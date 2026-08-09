using ShonOffice.Domain.Documentos;

namespace ShonOffice.Domain.Puertos;

/// <summary>
/// Puerto para guardar un <see cref="DocumentoExcel"/> como archivo <c>.xlsx</c>.
/// </summary>
public interface IXlsxWriter
{
    void Escribir(DocumentoExcel documento, string rutaDestino);
}
