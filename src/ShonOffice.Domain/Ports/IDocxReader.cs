using ShonOffice.Domain.Documents;

namespace ShonOffice.Domain.Ports;

/// <summary>
/// Port for reading a <c>.docx</c> file from disk. The concrete
/// implementation (today Rust, likely Open XML SDK tomorrow in
/// <c>ShonOffice.Infra.OpenXml</c>) is an infrastructure detail.
/// </summary>
public interface IDocxReader
{
    WordDocument Read(string filePath);
}
