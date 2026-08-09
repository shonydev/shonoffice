using ShonOffice.Domain.Documentos;

namespace ShonOffice.Domain.Puertos;

/// <summary>
/// Puerto para leer un archivo <c>.pptx</c> desde disco.
/// </summary>
public interface IPptxReader
{
    DocumentoPowerPoint Leer(string rutaArchivo);
}
