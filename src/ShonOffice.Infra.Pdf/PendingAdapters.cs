using ShonOffice.Domain.Ports;

namespace ShonOffice.Infra.Pdf;

/// <summary>
/// Placeholder for the first <see cref="IPdfEngine"/> implementation:
/// pure .NET, no Rust/FFI involved. This is the "C# by default" adapter —
/// per the README's "Nuevo rumbo arquitectónico", PDF support should start
/// here, using a managed library (e.g. PdfPig, iText, QuestPDF), and only
/// move to a Rust engine behind this same port if a concrete technical
/// need justifies the extra complexity (FFI, cross-platform native
/// builds, a separate CI pipeline, etc.) — see "Próximos pasos sugeridos"
/// in the main README.
/// </summary>
public sealed class NotImplementedManagedPdfEngine : IPdfEngine
{
    public IReadOnlyList<string> ExtractText(string pdfFilePath) =>
        throw new NotImplementedException("PDF reading not implemented yet.");
}
