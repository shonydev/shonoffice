using ShonOffice.Domain.Documentos;
using ShonOffice.Domain.Puertos;

namespace ShonOffice.Infra.OpenXml;

/// <summary>
/// Implementación placeholder de <see cref="IXlsxReader"/>: permite terminar
/// de inyectar <c>AbrirDocumentoCasoDeUso</c> (que necesita los tres
/// lectores) antes de que exista la lectura real de Excel, sin bloquear la
/// UI de <c>ShonOffice.docx</c>. Ver "Próximos pasos" en el README:
/// "Leer Excel (.xlsx) — vía Open XML SDK en Infra.OpenXml".
/// </summary>
public sealed class LectorExcelNoImplementado : IXlsxReader
{
    public DocumentoExcel Leer(string rutaArchivo) =>
        throw new NotImplementedException("Lectura de .xlsx todavía no implementada.");
}

/// <summary>
/// Implementación placeholder de <see cref="IPptxReader"/>, análoga a
/// <see cref="LectorExcelNoImplementado"/> pero para PowerPoint. Ver
/// "Próximos pasos" en el README: "Leer PowerPoint (.pptx) — vía Open XML
/// SDK en Infra.OpenXml".
/// </summary>
public sealed class LectorPowerPointNoImplementado : IPptxReader
{
    public DocumentoPowerPoint Leer(string rutaArchivo) =>
        throw new NotImplementedException("Lectura de .pptx todavía no implementada.");
}
