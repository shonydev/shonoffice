using ShonOffice.Domain.Documents;

namespace ShonOffice.Domain.Ports;

/// <summary>
/// Port for reading a <c>.xlsx</c> file from disk.
/// </summary>
public interface IXlsxReader
{
    ExcelDocument Read(string filePath);
}
