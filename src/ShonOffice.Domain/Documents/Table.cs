namespace ShonOffice.Domain.Documents;

/// <summary>
/// A cell of a <see cref="Table"/>. In OOXML a cell (<c>w:tc</c>) contains,
/// just like the document body, a list of paragraphs (almost always just
/// one, but it can have more), so <see cref="Paragraph"/> is reused as-is
/// instead of duplicating a separate text model for tables.
/// </summary>
public sealed class TableCell
{
    public IReadOnlyList<Paragraph> Paragraphs { get; }

    /// <summary>
    /// Background color of the cell in "RRGGBB" hex (the cell's <c>w:shd</c>),
    /// or null if it has no explicit fill. Not every row of a table is a
    /// "header": this reflects the real formatting carried by the
    /// <c>.docx</c> instead of assuming, by position, that the first row
    /// always has a different background color.
    /// </summary>
    public string? BackgroundColorHex { get; }

    public TableCell(IReadOnlyList<Paragraph> paragraphs, string? backgroundColorHex = null)
    {
        Paragraphs = paragraphs;
        BackgroundColorHex = backgroundColorHex;
    }

    /// <summary>Plain text of the cell, joining its paragraphs with a line break.</summary>
    public string PlainText => string.Join("\n", Paragraphs.Select(p => p.PlainText));
}

/// <summary>A row of a <see cref="Table"/>: the list of cells that make it up, in order.</summary>
public sealed class TableRow
{
    public IReadOnlyList<TableCell> Cells { get; }

    public TableRow(IReadOnlyList<TableCell> cells)
    {
        Cells = cells;
    }
}

/// <summary>
/// A table of a <see cref="WordDocument"/> (<c>w:tbl</c> in OOXML): a list
/// of <see cref="TableRow"/>, each with its <see cref="TableCell"/>. There
/// was no equivalent in the domain model before this — the first
/// <c>IDocxReader</c> adapter could only read top-level paragraphs of the
/// document body, so any table in the <c>.docx</c> was completely lost when
/// reading it (see <see cref="IContentElement"/>).
/// </summary>
/// <remarks>
/// Named <c>Table</c>/<c>TableRow</c>/<c>TableCell</c> here in the domain;
/// <c>ShonOffice.Infra.OpenXml</c> aliases the equivalent OOXML types to
/// avoid the name clash.
/// </remarks>
public sealed class Table : IContentElement
{
    public IReadOnlyList<TableRow> Rows { get; }

    public Table(IReadOnlyList<TableRow> rows)
    {
        Rows = rows;
    }

    /// <summary>Maximum number of columns across all rows (rows may not all be the same length, e.g. due to merged cells).</summary>
    public int ColumnCount => Rows.Count == 0 ? 0 : Rows.Max(r => r.Cells.Count);
}
