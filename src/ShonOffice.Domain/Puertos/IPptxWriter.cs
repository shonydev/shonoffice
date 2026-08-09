using ShonOffice.Domain.Documentos;

namespace ShonOffice.Domain.Puertos;

/// <summary>
/// Puerto para guardar un <see cref="DocumentoPowerPoint"/> como archivo <c>.pptx</c>.
/// </summary>
public interface IPptxWriter
{
    void Escribir(DocumentoPowerPoint documento, string rutaDestino);
}
