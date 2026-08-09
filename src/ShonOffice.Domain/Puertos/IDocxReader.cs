using ShonOffice.Domain.Documentos;

namespace ShonOffice.Domain.Puertos;

/// <summary>
/// Puerto para leer un archivo <c>.docx</c> desde disco. La implementación
/// concreta (hoy en Rust, mañana probablemente Open XML SDK en
/// <c>ShonOffice.Infra.OpenXml</c>) es un detalle de infraestructura.
/// </summary>
public interface IDocxReader
{
    DocumentoWord Leer(string rutaArchivo);
}
