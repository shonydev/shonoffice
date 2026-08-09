namespace ShonOffice.Domain.Excepciones;

/// <summary>
/// Se lanza cuando se pide abrir o guardar un archivo cuya extensión no
/// corresponde a ningún puerto conocido (<c>.docx</c>, <c>.xlsx</c>, <c>.pptx</c>).
/// </summary>
public sealed class FormatoNoSoportadoException : Exception
{
    public FormatoNoSoportadoException(string extensionODescripcion)
        : base($"Formato no soportado: '{extensionODescripcion}'")
    {
    }
}
