namespace ShonOffice.Domain.Documents;

/// <summary>
/// An Excel document in memory: a collection of <see cref="Sheet"/>.
/// </summary>
public sealed class ExcelDocument : OfficeDocument
{
    public override DocumentType Type => DocumentType.Excel;

    public IReadOnlyList<Sheet> Sheets { get; }

    public ExcelDocument(string filePath, IReadOnlyList<Sheet> sheets)
        : base(filePath)
    {
        Sheets = sheets;
    }
}
