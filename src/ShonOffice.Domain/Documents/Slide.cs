namespace ShonOffice.Domain.Documents;

/// <summary>
/// A slide within a <see cref="PowerPointDocument"/>.
/// </summary>
public sealed class Slide
{
    public int Number { get; }

    public string? Title { get; }

    public IReadOnlyList<string> ContentTexts { get; }

    public Slide(int number, string? title, IReadOnlyList<string> contentTexts)
    {
        Number = number;
        Title = title;
        ContentTexts = contentTexts;
    }
}
