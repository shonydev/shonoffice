using ShonOffice.Domain.Documents;
using ShonOffice.Domain.Ports;

namespace ShonOffice.Application.Tests.Fakes;

/// <summary>
/// In-memory fake adapters for the domain's ports. They allow testing the
/// use cases without depending on Open XML SDK or the Rust engine.
/// </summary>
internal sealed class FakeWordReader : IDocxReader
{
    public WordDocument Read(string filePath) => new(filePath, new[] { "fake content" });
}

internal sealed class FakeExcelReader : IXlsxReader
{
    public ExcelDocument Read(string filePath) => new(filePath, Array.Empty<Sheet>());
}

internal sealed class FakePowerPointReader : IPptxReader
{
    public PowerPointDocument Read(string filePath) => new(filePath, Array.Empty<Slide>());
}

internal sealed class FakeWordWriter : IDocxWriter
{
    public List<(WordDocument Document, string Path)> Writes { get; } = new();

    public void Write(WordDocument document, string destinationPath) =>
        Writes.Add((document, destinationPath));
}

internal sealed class FakeExcelWriter : IXlsxWriter
{
    public List<(ExcelDocument Document, string Path)> Writes { get; } = new();

    public void Write(ExcelDocument document, string destinationPath) =>
        Writes.Add((document, destinationPath));
}

internal sealed class FakePowerPointWriter : IPptxWriter
{
    public List<(PowerPointDocument Document, string Path)> Writes { get; } = new();

    public void Write(PowerPointDocument document, string destinationPath) =>
        Writes.Add((document, destinationPath));
}

internal sealed class FakePdfEngine : IPdfEngine
{
    private readonly IReadOnlyList<string> _textToReturn;

    public FakePdfEngine(IReadOnlyList<string>? textToReturn = null) =>
        _textToReturn = textToReturn ?? new[] { "text extracted from the fake pdf" };

    public IReadOnlyList<string> ExtractText(string pdfFilePath) => _textToReturn;
}
