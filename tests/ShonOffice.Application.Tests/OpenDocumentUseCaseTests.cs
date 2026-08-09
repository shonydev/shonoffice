using ShonOffice.Application.Tests.Fakes;
using ShonOffice.Application.UseCases;
using ShonOffice.Domain.Documents;
using ShonOffice.Domain.Exceptions;
using Xunit;

namespace ShonOffice.Application.Tests;

public class OpenDocumentUseCaseTests
{
    private static OpenDocumentUseCase CreateUseCase() =>
        new(new FakeWordReader(), new FakeExcelReader(), new FakePowerPointReader());

    [Fact]
    public void Execute_WithDocxFile_ReturnsWordDocument()
    {
        var useCase = CreateUseCase();

        var result = useCase.Execute("report.docx");

        Assert.IsType<WordDocument>(result);
    }

    [Fact]
    public void Execute_WithXlsxFile_ReturnsExcelDocument()
    {
        var useCase = CreateUseCase();

        var result = useCase.Execute("spreadsheet.xlsx");

        Assert.IsType<ExcelDocument>(result);
    }

    [Fact]
    public void Execute_WithPptxFile_ReturnsPowerPointDocument()
    {
        var useCase = CreateUseCase();

        var result = useCase.Execute("presentation.pptx");

        Assert.IsType<PowerPointDocument>(result);
    }

    [Fact]
    public void Execute_WithUnsupportedExtension_ThrowsUnsupportedFormatException()
    {
        var useCase = CreateUseCase();

        Assert.Throws<UnsupportedFormatException>(() => useCase.Execute("file.txt"));
    }

    [Fact]
    public void Execute_IsCaseInsensitive()
    {
        var useCase = CreateUseCase();

        var result = useCase.Execute("REPORT.DOCX");

        Assert.IsType<WordDocument>(result);
    }
}
