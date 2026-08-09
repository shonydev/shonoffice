namespace ShonOffice.Domain.Documentos;

/// <summary>
/// Un "run" de Word: un tramo de texto dentro de un <see cref="Parrafo"/>
/// que comparte el mismo formato de caracter (negrita, cursiva, color,
/// tamaño, fuente). Un párrafo casi siempre tiene más de un <see cref="Texto"/>
/// cuando, por ejemplo, solo una palabra está en negrita.
/// </summary>
public sealed class Texto
{
    public string Contenido { get; }

    public bool Negrita { get; }

    public bool Cursiva { get; }

    public bool Subrayado { get; }

    /// <summary>Tamaño de fuente en puntos, o null si no se pudo resolver.</summary>
    public double? TamanoPunto { get; }

    /// <summary>Color en formato hexadecimal "RRGGBB" (igual que OOXML), o null si es el color por defecto.</summary>
    public string? ColorHex { get; }

    public string? NombreFuente { get; }

    public Texto(
        string contenido,
        bool negrita = false,
        bool cursiva = false,
        bool subrayado = false,
        double? tamanoPunto = null,
        string? colorHex = null,
        string? nombreFuente = null)
    {
        Contenido = contenido;
        Negrita = negrita;
        Cursiva = cursiva;
        Subrayado = subrayado;
        TamanoPunto = tamanoPunto;
        ColorHex = colorHex;
        NombreFuente = nombreFuente;
    }
}
