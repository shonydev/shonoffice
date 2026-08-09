using ShonOffice.Application.Tests.Fakes;
using ShonOffice.Application.UseCases;
using Xunit;

namespace ShonOffice.Application.Tests;

public class ConvertPdfToWordUseCaseTests
{
    [Fact]
    public void Execute_ExtractsTextFromThePdfEngineAndSavesItAsWord()
    {
        var pdfEngine = new FakePdfEngine(new[] { "first paragraph", "second paragraph" });
        var wordWriter = new FakeWordWriter();
        var useCase = new ConvertPdfToWordUseCase(pdfEngine, wordWriter);

        var result = useCase.Execute("source.pdf", "destination.docx");

        Assert.Equal(new[] { "first paragraph", "second paragraph" }, result.Paragraphs);
        var write = Assert.Single(wordWriter.Writes);
        Assert.Equal("destination.docx", write.Path);
        Assert.Same(result, write.Document);
    }
}
