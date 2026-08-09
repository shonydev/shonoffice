namespace ShonOffice.Domain.Documentos;

/// <summary>
/// Una hoja dentro de un <see cref="DocumentoExcel"/>: su nombre y sus filas,
/// ya convertidas a texto celda por celda.
/// </summary>
public sealed class Hoja
{
    public string Nombre { get; }

    public IReadOnlyList<IReadOnlyList<string>> Filas { get; }

    public Hoja(string nombre, IReadOnlyList<IReadOnlyList<string>> filas)
    {
        Nombre = nombre;
        Filas = filas;
    }
}
