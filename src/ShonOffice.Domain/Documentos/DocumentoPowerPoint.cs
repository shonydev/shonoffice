namespace ShonOffice.Domain.Documentos;

/// <summary>
/// Documento PowerPoint en memoria: una colección de <see cref="Diapositiva"/>.
/// </summary>
public sealed class DocumentoPowerPoint : Documento
{
    public override TipoDocumento Tipo => TipoDocumento.PowerPoint;

    public IReadOnlyList<Diapositiva> Diapositivas { get; }

    public DocumentoPowerPoint(string rutaArchivo, IReadOnlyList<Diapositiva> diapositivas)
        : base(rutaArchivo)
    {
        Diapositivas = diapositivas;
    }
}
