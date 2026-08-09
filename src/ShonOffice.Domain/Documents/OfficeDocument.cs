namespace ShonOffice.Domain.Documents;

/// <summary>
/// Represents an Office document already opened in memory, regardless of
/// which infrastructure adapter read it (Open XML SDK, the Rust engine,
/// etc.). This is the core of the domain: it doesn't know about any
/// external library.
/// </summary>
/// <remarks>
/// Named <c>OfficeDocument</c> rather than the more obvious <c>Document</c>
/// on purpose: <c>DocumentFormat.OpenXml.Wordprocessing</c> already defines
/// a <c>Document</c> type (the OOXML <c>w:document</c> root element), and
/// both namespaces are imported side by side in <c>ShonOffice.Infra.OpenXml</c>.
/// </remarks>
public abstract class OfficeDocument
{
    /// <summary>Path of the source file (or destination, when saving).</summary>
    public string FilePath { get; }

    /// <summary>Concrete format of this document.</summary>
    public abstract DocumentType Type { get; }

    protected OfficeDocument(string filePath)
    {
        FilePath = filePath;
    }
}
