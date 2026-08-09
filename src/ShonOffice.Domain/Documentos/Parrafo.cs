namespace ShonOffice.Domain.Documentos;

/// <summary>
/// Un párrafo de un <see cref="DocumentoWord"/> con su formato: alineación,
/// sangría, nivel de encabezado (si es un título) y la lista de
/// <see cref="Texto"/> (runs) que lo componen. Es lo que le falta al modelo
/// de texto plano original para poder mostrarse como lo hace Word de verdad,
/// en vez de una lista de líneas sin estilo.
/// </summary>
public sealed class Parrafo : IElementoDeContenido
{
    public IReadOnlyList<Texto> Textos { get; }

    public AlineacionTexto Alineacion { get; }

    /// <summary>Nivel de encabezado 1-9 (estilos "Heading1".."Heading9"/"Title"), o null si es texto normal.</summary>
    public int? NivelEncabezado { get; }

    public double SangriaIzquierdaPuntos { get; }

    /// <summary>Id del estilo de párrafo de Word (p. ej. "Heading1"), para diagnóstico/depuración.</summary>
    public string? NombreEstilo { get; }

    public Parrafo(
        IReadOnlyList<Texto> textos,
        AlineacionTexto alineacion = AlineacionTexto.Izquierda,
        int? nivelEncabezado = null,
        double sangriaIzquierdaPuntos = 0,
        string? nombreEstilo = null)
    {
        Textos = textos;
        Alineacion = alineacion;
        NivelEncabezado = nivelEncabezado;
        SangriaIzquierdaPuntos = sangriaIzquierdaPuntos;
        NombreEstilo = nombreEstilo;
    }

    /// <summary>Texto plano concatenado de todos los runs, para compatibilidad con el modelo simple (<see cref="DocumentoWord.Parrafos"/>).</summary>
    public string TextoPlano => string.Concat(Textos.Select(t => t.Contenido));
}
