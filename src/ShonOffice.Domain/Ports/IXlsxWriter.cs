using ShonOffice.Domain.Documents;

namespace ShonOffice.Domain.Ports;

/// <summary>
/// Port for saving an <see cref="ExcelDocument"/> as a <c>.xlsx</c> file.
/// </summary>
public interface IXlsxWriter
{
    void Write(ExcelDocument document, string destinationPath);
}
