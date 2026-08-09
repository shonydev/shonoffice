namespace ShonOffice.Domain.Exceptions;

/// <summary>
/// Thrown when asked to open or save a file whose extension doesn't match
/// any known port (<c>.docx</c>, <c>.xlsx</c>, <c>.pptx</c>).
/// </summary>
public sealed class UnsupportedFormatException : Exception
{
    public UnsupportedFormatException(string extensionOrDescription)
        : base($"Unsupported format: '{extensionOrDescription}'")
    {
    }
}
