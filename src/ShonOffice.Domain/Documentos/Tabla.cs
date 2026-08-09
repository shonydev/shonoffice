namespace ShonOffice.Domain.Documentos;

/// <summary>
/// Una celda de una <see cref="Tabla"/>. En OOXML una celda (<c>w:tc</c>)
/// contiene, igual que el cuerpo del documento, una lista de párrafos (casi
/// siempre uno solo, pero puede tener más), así que se reutiliza
/// <see cref="Parrafo"/> tal cual en vez de duplicar un modelo de texto
/// distinto para tablas.
/// </summary>
public sealed class CeldaTabla
{
    public IReadOnlyList<Parrafo> Parrafos { get; }

    /// <summary>
    /// Color de fondo de la celda en hexadecimal "RRGGBB" (el <c>w:shd</c>
    /// de la celda), o null si no tiene relleno explícito. No todas las
    /// filas de una tabla son "header": esto refleja el formato real que
    /// trae el <c>.docx</c> en vez de asumir, por posición, que la primera
    /// fila siempre lleva un color de fondo distinto.
    /// </summary>
    public string? ColorDeFondoHex { get; }

    public CeldaTabla(IReadOnlyList<Parrafo> parrafos, string? colorDeFondoHex = null)
    {
        Parrafos = parrafos;
        ColorDeFondoHex = colorDeFondoHex;
    }

    /// <summary>Texto plano de la celda, uniendo sus párrafos con salto de línea.</summary>
    public string TextoPlano => string.Join("\n", Parrafos.Select(p => p.TextoPlano));
}

/// <summary>Una fila de una <see cref="Tabla"/>: la lista de celdas que la componen, en orden.</summary>
public sealed class FilaTabla
{
    public IReadOnlyList<CeldaTabla> Celdas { get; }

    public FilaTabla(IReadOnlyList<CeldaTabla> celdas)
    {
        Celdas = celdas;
    }
}

/// <summary>
/// Una tabla de un <see cref="DocumentoWord"/> (<c>w:tbl</c> en OOXML): una
/// lista de <see cref="FilaTabla"/>, cada una con sus <see cref="CeldaTabla"/>.
/// No existía ningún equivalente en el modelo de dominio antes de esto — el
/// primer adaptador de <c>IDocxReader</c> solo sabía leer párrafos de nivel
/// superior del cuerpo del documento, así que cualquier tabla del <c>.docx</c>
/// se perdía por completo al leerlo (ver <see cref="IElementoDeContenido"/>).
/// </summary>
public sealed class Tabla : IElementoDeContenido
{
    public IReadOnlyList<FilaTabla> Filas { get; }

    public Tabla(IReadOnlyList<FilaTabla> filas)
    {
        Filas = filas;
    }

    /// <summary>Cantidad máxima de columnas entre todas las filas (las filas pueden no ser todas del mismo largo, p. ej. por celdas combinadas).</summary>
    public int CantidadDeColumnas => Filas.Count == 0 ? 0 : Filas.Max(f => f.Celdas.Count);
}
