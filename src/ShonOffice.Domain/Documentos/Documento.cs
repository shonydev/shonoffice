namespace ShonOffice.Domain.Documentos;

/// <summary>
/// Representa un documento de Office ya abierto en memoria, independientemente
/// de qué adaptador de infraestructura lo haya leído (Open XML SDK, motor Rust,
/// etc.). Es el núcleo del dominio: no conoce ninguna librería externa.
/// </summary>
public abstract class Documento
{
    /// <summary>Ruta del archivo de origen (o destino, al guardar).</summary>
    public string RutaArchivo { get; }

    /// <summary>Formato concreto de este documento.</summary>
    public abstract TipoDocumento Tipo { get; }

    protected Documento(string rutaArchivo)
    {
        RutaArchivo = rutaArchivo;
    }
}
