namespace ShonOffice.Domain.Puertos;

/// <summary>
/// Puerto hacia el motor de procesamiento de PDF. La implementación pesada
/// (parsing/extracción de layout) vive en Rust y se expone como <c>cdylib</c>
/// vía FFI en <c>ShonOffice.Infra.Native</c>; el dominio solo conoce esta
/// interfaz, nunca los bindings generados por <c>csbindgen</c>.
/// </summary>
public interface IPdfEngine
{
    /// <summary>
    /// Extrae el texto de un PDF, un bloque/párrafo por elemento, como paso
    /// previo a reconstruirlo como <see cref="Documentos.DocumentoWord"/> en
    /// la capa de aplicación.
    /// </summary>
    IReadOnlyList<string> ExtraerTexto(string rutaArchivoPdf);
}
