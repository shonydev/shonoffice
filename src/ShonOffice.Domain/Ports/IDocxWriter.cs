using ShonOffice.Domain.Documents;

namespace ShonOffice.Domain.Ports;

/// <summary>
/// Port for saving a <see cref="WordDocument"/> as a <c>.docx</c> file.
/// </summary>
public interface IDocxWriter
{
    void Write(WordDocument document, string destinationPath);
}
