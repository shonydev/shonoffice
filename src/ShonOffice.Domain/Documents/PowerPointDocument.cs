namespace ShonOffice.Domain.Documents;

/// <summary>
/// A PowerPoint document in memory: a collection of <see cref="Slide"/>.
/// </summary>
public sealed class PowerPointDocument : OfficeDocument
{
    public override DocumentType Type => DocumentType.PowerPoint;

    public IReadOnlyList<Slide> Slides { get; }

    public PowerPointDocument(string filePath, IReadOnlyList<Slide> slides)
        : base(filePath)
    {
        Slides = slides;
    }
}
