using ShonOffice.Domain.Documents;
using ShonOffice.Domain.Ports;

namespace ShonOffice.Infra.OpenXml;

/// <summary>
/// Placeholder implementation of <see cref="IXlsxReader"/>: lets us finish
/// wiring up <c>OpenDocumentUseCase</c> (which needs all three readers)
/// before real Excel reading exists, without blocking the
/// <c>ShonOffice</c> UI for <c>.docx</c>. See "Next steps" in the README:
/// "Read Excel (.xlsx) — via Open XML SDK in Infra.OpenXml".
/// </summary>
public sealed class NotImplementedExcelReader : IXlsxReader
{
    public ExcelDocument Read(string filePath) =>
        throw new NotImplementedException(".xlsx reading not implemented yet.");
}

/// <summary>
/// Placeholder implementation of <see cref="IPptxReader"/>, analogous to
/// <see cref="NotImplementedExcelReader"/> but for PowerPoint. See "Next
/// steps" in the README: "Read PowerPoint (.pptx) — via Open XML SDK in
/// Infra.OpenXml".
/// </summary>
public sealed class NotImplementedPowerPointReader : IPptxReader
{
    public PowerPointDocument Read(string filePath) =>
        throw new NotImplementedException(".pptx reading not implemented yet.");
}
