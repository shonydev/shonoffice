namespace ShonOffice.Domain.Documents;

/// <summary>
/// A sheet within an <see cref="ExcelDocument"/>: its name and its rows,
/// already converted to text cell by cell.
/// </summary>
public sealed class Sheet
{
    public string Name { get; }

    public IReadOnlyList<IReadOnlyList<string>> Rows { get; }

    public Sheet(string name, IReadOnlyList<IReadOnlyList<string>> rows)
    {
        Name = name;
        Rows = rows;
    }
}
