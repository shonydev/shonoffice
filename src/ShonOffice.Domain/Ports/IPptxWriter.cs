using ShonOffice.Domain.Documents;

namespace ShonOffice.Domain.Ports;

/// <summary>
/// Port for saving a <see cref="PowerPointDocument"/> as a <c>.pptx</c> file.
/// </summary>
public interface IPptxWriter
{
    void Write(PowerPointDocument document, string destinationPath);
}
