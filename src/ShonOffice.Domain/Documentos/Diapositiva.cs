namespace ShonOffice.Domain.Documentos;

/// <summary>
/// Una diapositiva dentro de un <see cref="DocumentoPowerPoint"/>.
/// </summary>
public sealed class Diapositiva
{
    public int Numero { get; }

    public string? Titulo { get; }

    public IReadOnlyList<string> TextosContenido { get; }

    public Diapositiva(int numero, string? titulo, IReadOnlyList<string> textosContenido)
    {
        Numero = numero;
        Titulo = titulo;
        TextosContenido = textosContenido;
    }
}
