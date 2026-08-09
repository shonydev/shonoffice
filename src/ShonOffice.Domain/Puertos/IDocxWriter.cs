using ShonOffice.Domain.Documentos;

namespace ShonOffice.Domain.Puertos;

/// <summary>
/// Puerto para guardar un <see cref="DocumentoWord"/> como archivo <c>.docx</c>.
/// </summary>
public interface IDocxWriter
{
    void Escribir(DocumentoWord documento, string rutaDestino);
}
