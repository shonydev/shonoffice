using ShonOffice.Domain.Documents;

namespace ShonOffice.Domain.Ports;

/// <summary>
/// Port for reading a <c>.pptx</c> file from disk.
/// </summary>
public interface IPptxReader
{
    PowerPointDocument Read(string filePath);
}
