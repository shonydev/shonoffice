namespace ShonOffice.Domain.Documentos;

/// <summary>
/// Alineación horizontal de un <see cref="Parrafo"/>, tal como la define
/// Word (<c>w:jc</c>): a diferencia del texto plano que leía la fase Rust,
/// esto es necesario para poder reconstruir, por ejemplo, un título
/// centrado igual que lo muestra Word.
/// </summary>
public enum AlineacionTexto
{
    Izquierda,
    Centro,
    Derecha,
    Justificado,
}
