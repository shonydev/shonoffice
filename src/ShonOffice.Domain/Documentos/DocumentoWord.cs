namespace ShonOffice.Domain.Documentos;

/// <summary>
/// Documento Word en memoria: su contenido, un elemento por cada párrafo o
/// tabla de nivel superior. <see cref="Parrafos"/> es texto plano y
/// equivale a lo que ya leía la fase Rust (<c>read_docx_text</c>);
/// <see cref="ParrafosConFormato"/> es el modelo enriquecido (negrita,
/// color, tamaño, alineación, encabezados...) que entrega un adaptador
/// capaz de resolverlo, como <c>ShonOffice.Infra.OpenXml</c>; y
/// <see cref="ElementosConFormato"/> es lo mismo pero intercalando también
/// las <see cref="Tabla"/> en el orden real en que aparecen en el
/// documento — <see cref="ParrafosConFormato"/> por sí solo no alcanza para
/// eso porque, al ser solo párrafos, no tiene forma de decir "acá va una
/// tabla" ni de mantener su posición relativa al resto del texto. Se
/// mantienen los tres porque no todos los adaptadores pueden producir
/// formato: la reconstrucción de un <see cref="DocumentoWord"/> a partir de
/// un PDF (<see cref="Puertos.IPdfEngine"/>), por ejemplo, solo tiene texto.
/// </summary>
public sealed class DocumentoWord : Documento
{
    public override TipoDocumento Tipo => TipoDocumento.Word;

    public IReadOnlyList<string> Parrafos { get; }

    /// <summary>
    /// Formato completo por párrafo, cuando el adaptador que generó este
    /// documento lo soporta. Null cuando solo hay texto plano disponible.
    /// No incluye tablas: si el documento tiene alguna, usar
    /// <see cref="ElementosConFormato"/> para no perderlas.
    /// </summary>
    public IReadOnlyList<Parrafo>? ParrafosConFormato { get; }

    /// <summary>
    /// Párrafos y tablas de nivel superior, en el mismo orden en que
    /// aparecen en el documento. Null cuando solo hay texto plano
    /// disponible (mismo caso que <see cref="ParrafosConFormato"/> null).
    /// </summary>
    public IReadOnlyList<IElementoDeContenido>? ElementosConFormato { get; }

    public DocumentoWord(string rutaArchivo, IReadOnlyList<string> parrafos)
        : base(rutaArchivo)
    {
        Parrafos = parrafos;
        ParrafosConFormato = null;
        ElementosConFormato = null;
    }

    public DocumentoWord(string rutaArchivo, IReadOnlyList<Parrafo> parrafosConFormato)
        : base(rutaArchivo)
    {
        ParrafosConFormato = parrafosConFormato;
        ElementosConFormato = parrafosConFormato;
        Parrafos = parrafosConFormato.Select(p => p.TextoPlano).ToArray();
    }

    public DocumentoWord(string rutaArchivo, IReadOnlyList<IElementoDeContenido> elementosConFormato)
        : base(rutaArchivo)
    {
        ElementosConFormato = elementosConFormato;
        ParrafosConFormato = elementosConFormato.OfType<Parrafo>().ToArray();
        Parrafos = elementosConFormato.Select(TextoPlanoDe).ToArray();
    }

    private static string TextoPlanoDe(IElementoDeContenido elemento) => elemento switch
    {
        Parrafo parrafo => parrafo.TextoPlano,
        Tabla tabla => string.Join(
            "\n",
            tabla.Filas.Select(fila => string.Join(" | ", fila.Celdas.Select(c => c.TextoPlano)))),
        _ => string.Empty,
    };
}
