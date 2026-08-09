using ShonOffice.Domain.Documentos;
using ShonOffice.Domain.Puertos;

namespace ShonOffice.Application.Tests.Falsos;

/// <summary>
/// Adaptadores falsos in-memory para los puertos del dominio. Permiten
/// testear los casos de uso sin depender de Open XML SDK ni del motor Rust.
/// </summary>
internal sealed class LectorWordFalso : IDocxReader
{
    public DocumentoWord Leer(string rutaArchivo) => new(rutaArchivo, new[] { "contenido falso" });
}

internal sealed class LectorExcelFalso : IXlsxReader
{
    public DocumentoExcel Leer(string rutaArchivo) => new(rutaArchivo, Array.Empty<Hoja>());
}

internal sealed class LectorPowerPointFalso : IPptxReader
{
    public DocumentoPowerPoint Leer(string rutaArchivo) => new(rutaArchivo, Array.Empty<Diapositiva>());
}

internal sealed class EscritorWordFalso : IDocxWriter
{
    public List<(DocumentoWord Documento, string Ruta)> Escrituras { get; } = new();

    public void Escribir(DocumentoWord documento, string rutaDestino) =>
        Escrituras.Add((documento, rutaDestino));
}

internal sealed class EscritorExcelFalso : IXlsxWriter
{
    public List<(DocumentoExcel Documento, string Ruta)> Escrituras { get; } = new();

    public void Escribir(DocumentoExcel documento, string rutaDestino) =>
        Escrituras.Add((documento, rutaDestino));
}

internal sealed class EscritorPowerPointFalso : IPptxWriter
{
    public List<(DocumentoPowerPoint Documento, string Ruta)> Escrituras { get; } = new();

    public void Escribir(DocumentoPowerPoint documento, string rutaDestino) =>
        Escrituras.Add((documento, rutaDestino));
}

internal sealed class MotorPdfFalso : IPdfEngine
{
    private readonly IReadOnlyList<string> _textoAEntregar;

    public MotorPdfFalso(IReadOnlyList<string>? textoAEntregar = null) =>
        _textoAEntregar = textoAEntregar ?? new[] { "texto extraido del pdf falso" };

    public IReadOnlyList<string> ExtraerTexto(string rutaArchivoPdf) => _textoAEntregar;
}
