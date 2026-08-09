namespace ShonOffice.Domain.Documentos;

/// <summary>
/// Documento Excel en memoria: una colección de <see cref="Hoja"/>.
/// </summary>
public sealed class DocumentoExcel : Documento
{
    public override TipoDocumento Tipo => TipoDocumento.Excel;

    public IReadOnlyList<Hoja> Hojas { get; }

    public DocumentoExcel(string rutaArchivo, IReadOnlyList<Hoja> hojas)
        : base(rutaArchivo)
    {
        Hojas = hojas;
    }
}
