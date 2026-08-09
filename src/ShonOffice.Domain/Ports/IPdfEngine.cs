namespace ShonOffice.Domain.Ports;

/// <summary>
/// Port to the PDF processing engine (parsing/layout extraction). Per the
/// "Nuevo rumbo arquitectónico" described in the README, the domain
/// doesn't assume which technology implements this: the first
/// implementation is expected to be pure .NET
/// (<c>ShonOffice.Infra.Pdf.NotImplementedManagedPdfEngine</c>), and a
/// Rust engine exposed via FFI is only a candidate for later, behind this
/// same interface, if a concrete technical need justifies it.
/// </summary>
public interface IPdfEngine
{
    /// <summary>
    /// Extracts the text of a PDF, one block/paragraph per element, as a
    /// step prior to rebuilding it as a <see cref="Documents.WordDocument"/>
    /// in the application layer.
    /// </summary>
    IReadOnlyList<string> ExtractText(string pdfFilePath);
}
