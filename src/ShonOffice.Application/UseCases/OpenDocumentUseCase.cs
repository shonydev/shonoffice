using ShonOffice.Domain.Documents;
using ShonOffice.Domain.Exceptions;
using ShonOffice.Domain.Ports;

namespace ShonOffice.Application.UseCases;

/// <summary>
/// Opens an Office document, detecting its type from the file extension
/// and delegating the reading to the corresponding port. It's the input
/// adapter (UI, CLI) that calls this use case; never the other way around.
/// </summary>
public sealed class OpenDocumentUseCase
{
    private readonly IDocxReader _wordReader;
    private readonly IXlsxReader _excelReader;
    private readonly IPptxReader _powerPointReader;

    public OpenDocumentUseCase(
        IDocxReader wordReader,
        IXlsxReader excelReader,
        IPptxReader powerPointReader)
    {
        _wordReader = wordReader;
        _excelReader = excelReader;
        _powerPointReader = powerPointReader;
    }

    public OfficeDocument Execute(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();

        return extension switch
        {
            ".docx" => _wordReader.Read(filePath),
            ".xlsx" => _excelReader.Read(filePath),
            ".pptx" => _powerPointReader.Read(filePath),
            _ => throw new UnsupportedFormatException(extension),
        };
    }
}
