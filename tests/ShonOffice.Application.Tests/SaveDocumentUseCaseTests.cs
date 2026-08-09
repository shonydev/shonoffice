using ShonOffice.Application.Tests.Fakes;
using ShonOffice.Application.UseCases;
using ShonOffice.Domain.Documents;
using Xunit;

namespace ShonOffice.Application.Tests;

public class SaveDocumentUseCaseTests
{
    [Fact]
    public void Execute_WithWordDocument_CallsTheWordWriter()
    {
        var wordWriter = new FakeWordWriter();
        var useCase = new SaveDocumentUseCase(
            wordWriter, new FakeExcelWriter(), new FakePowerPointWriter());
        var document = new WordDocument("report.docx", new[] { "hello" });

        useCase.Execute(document, "output.docx");

        var write = Assert.Single(wordWriter.Writes);
        Assert.Equal("output.docx", write.Path);
        Assert.Same(document, write.Document);
    }

    [Fact]
    public void Execute_WithExcelDocument_CallsTheExcelWriter()
    {
        var excelWriter = new FakeExcelWriter();
        var useCase = new SaveDocumentUseCase(
            new FakeWordWriter(), excelWriter, new FakePowerPointWriter());
        var document = new ExcelDocument("spreadsheet.xlsx", Array.Empty<Sheet>());

        useCase.Execute(document, "output.xlsx");

        Assert.Single(excelWriter.Writes);
    }
}
