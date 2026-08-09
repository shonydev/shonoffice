using ShonOffice.Domain.Documentos;

namespace ShonOffice.Domain.Puertos;

/// <summary>
/// Puerto para leer un archivo <c>.xlsx</c> desde disco.
/// </summary>
public interface IXlsxReader
{
    DocumentoExcel Leer(string rutaArchivo);
}
