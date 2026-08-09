using ShonOffice.Domain.Documents;
using ShonOffice.Domain.Ports;

namespace ShonOffice.Application.UseCases;

/// <summary>
/// Converts a PDF into a Word document. Text extraction is delegated to
/// whatever <see cref="IPdfEngine"/> implementation is wired in — a plain
/// .NET one today, or a Rust engine later if that becomes justified (see
/// the port's docs); this use case only orchestrates: extract text and
/// rebuild it as a <see cref="WordDocument"/> before saving it with
/// <see cref="IDocxWriter"/>.
/// </summary>
public sealed class ConvertPdfToWordUseCase
{
    private readonly IPdfEngine _pdfEngine;
    private readonly IDocxWriter _wordWriter;

    public ConvertPdfToWordUseCase(IPdfEngine pdfEngine, IDocxWriter wordWriter)
    {
        _pdfEngine = pdfEngine;
        _wordWriter = wordWriter;
    }

    public WordDocument Execute(string sourcePdfPath, string destinationWordPath)
    {
        var paragraphs = _pdfEngine.ExtractText(sourcePdfPath);
        var document = new WordDocument(destinationWordPath, paragraphs);

        _wordWriter.Write(document, destinationWordPath);

        return document;
    }
}
