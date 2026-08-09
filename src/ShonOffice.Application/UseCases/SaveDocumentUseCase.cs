using ShonOffice.Domain.Documents;
using ShonOffice.Domain.Exceptions;
using ShonOffice.Domain.Ports;

namespace ShonOffice.Application.UseCases;

/// <summary>
/// Saves an <see cref="OfficeDocument"/> already in memory to disk,
/// delegating to the writer port corresponding to its concrete type.
/// </summary>
public sealed class SaveDocumentUseCase
{
    private readonly IDocxWriter _wordWriter;
    private readonly IXlsxWriter _excelWriter;
    private readonly IPptxWriter _powerPointWriter;

    public SaveDocumentUseCase(
        IDocxWriter wordWriter,
        IXlsxWriter excelWriter,
        IPptxWriter powerPointWriter)
    {
        _wordWriter = wordWriter;
        _excelWriter = excelWriter;
        _powerPointWriter = powerPointWriter;
    }

    public void Execute(OfficeDocument document, string destinationPath)
    {
        switch (document)
        {
            case WordDocument word:
                _wordWriter.Write(word, destinationPath);
                break;
            case ExcelDocument excel:
                _excelWriter.Write(excel, destinationPath);
                break;
            case PowerPointDocument powerPoint:
                _powerPointWriter.Write(powerPoint, destinationPath);
                break;
            default:
                throw new UnsupportedFormatException(document.Type.ToString());
        }
    }
}
