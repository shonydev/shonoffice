using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using ShonOffice.Application.CasosDeUso;
using ShonOffice.Domain.Documentos;
using ShonOffice.Infra.OpenXml;

namespace ShonOffice.UI;

/// <summary>
/// Ventana principal. Es un adaptador de entrada más (como marca la
/// arquitectura hexagonal del README): llama a
/// <see cref="AbrirDocumentoCasoDeUso"/> de <c>ShonOffice.Application</c>,
/// nunca toca Open XML SDK directamente. Renderiza
/// <see cref="DocumentoWord.ElementosConFormato"/> — párrafos con negrita,
/// cursiva, subrayado, tamaño, color, alineación, sangría y encabezados,
/// intercalados con tablas — en vez de solo texto plano, que es la
/// diferencia visual con la GUI original en Rust/egui (<c>src/main.rs</c>),
/// que solo mostraba <c>ui.label(parrafo)</c> por cada párrafo sin ningún
/// estilo ni noción de tabla.
/// </summary>
public sealed class MainWindow : Window
{
    private static readonly Color ColorEncabezadoPorDefecto = Color.Parse("#2E74B5"); // azul típico de Heading1/2 en las plantillas de Word
    private static readonly Color ColorBordeTabla = Color.Parse("#BFBFBF"); // gris de borde de tabla típico de Word

    private readonly AbrirDocumentoCasoDeUso _abrirDocumento = new(
        new LectorWordOpenXml(),
        new LectorExcelNoImplementado(),
        new LectorPowerPointNoImplementado());

    private readonly TextBlock _rutaArchivoTexto = new() { VerticalAlignment = VerticalAlignment.Center, Opacity = 0.7 };
    private readonly TextBlock _errorTexto = new() { Foreground = Brushes.Firebrick, Margin = new Thickness(12, 0, 12, 8), IsVisible = false };
    private readonly StackPanel _contenido = new() { Spacing = 6, Margin = new Thickness(16) };

    public MainWindow()
    {
        Title = "ShonOffice";
        Width = 900;
        Height = 700;

        var botonAbrir = new Button { Content = "\U0001F4C2 Abrir Word..." };
        botonAbrir.Click += async (_, _) => await AbrirArchivoAsync();

        var barraSuperior = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 10,
            Margin = new Thickness(12, 10),
            Children = { botonAbrir, _rutaArchivoTexto },
        };

        var scroll = new ScrollViewer { Content = _contenido };

        var raiz = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,*") };
        Grid.SetRow(barraSuperior, 0);
        Grid.SetRow(_errorTexto, 1);
        Grid.SetRow(scroll, 2);
        raiz.Children.Add(barraSuperior);
        raiz.Children.Add(_errorTexto);
        raiz.Children.Add(scroll);

        Content = raiz;

        MostrarMensajeInicial();
    }

    private void MostrarMensajeInicial()
    {
        _contenido.Children.Clear();
        _contenido.Children.Add(new TextBlock
        {
            Text = "Abrí un archivo .docx para ver su contenido.",
            Opacity = 0.7,
        });
    }

    private async Task AbrirArchivoAsync()
    {
        _errorTexto.IsVisible = false;

        var archivos = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Abrir documento Word",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("Word (*.docx)") { Patterns = new[] { "*.docx" } },
            },
        });

        var archivo = archivos.FirstOrDefault();
        if (archivo is null)
        {
            return; // el usuario canceló el diálogo
        }

        var rutaArchivo = archivo.Path.LocalPath;

        try
        {
            var documento = (DocumentoWord)_abrirDocumento.Ejecutar(rutaArchivo);
            _rutaArchivoTexto.Text = rutaArchivo;
            MostrarDocumento(documento);
        }
        catch (Exception ex)
        {
            _contenido.Children.Clear();
            _errorTexto.Text = $"No se pudo abrir el documento: {ex.Message}";
            _errorTexto.IsVisible = true;
        }
    }

    private void MostrarDocumento(DocumentoWord documento)
    {
        _contenido.Children.Clear();

        var elementos = documento.ElementosConFormato;
        if (elementos is null || elementos.Count == 0)
        {
            // Adaptador sin formato enriquecido (p. ej. un DocumentoWord
            // reconstruido desde PDF vía ConvertirPdfAWordCasoDeUso): se
            // muestra el texto plano, igual que hacía la GUI en Rust.
            foreach (var texto in documento.Parrafos)
            {
                _contenido.Children.Add(new TextBlock { Text = texto, TextWrapping = TextWrapping.Wrap });
            }

            return;
        }

        foreach (var elemento in elementos)
        {
            Control bloque = elemento switch
            {
                Parrafo parrafo => CrearBloqueDeParrafo(parrafo),
                Tabla tabla => CrearBloqueDeTabla(tabla),
                _ => new TextBlock(),
            };

            _contenido.Children.Add(bloque);
        }
    }

    private static TextBlock CrearBloqueDeParrafo(Parrafo parrafo)
    {
        var (tamanoBase, negritaBase) = TamanoYNegritaPorEncabezado(parrafo.NivelEncabezado);

        var bloque = new TextBlock
        {
            TextWrapping = TextWrapping.Wrap,
            TextAlignment = ConvertirAlineacion(parrafo.Alineacion),
            Margin = new Thickness(PuntosAPixeles(parrafo.SangriaIzquierdaPuntos), 0, 0, 0),
            Inlines = new InlineCollection(),
        };

        var tieneTextoVisible = false;

        foreach (var texto in parrafo.Textos)
        {
            if (texto.Contenido.Length == 0)
            {
                continue;
            }

            tieneTextoVisible |= texto.Contenido.Trim().Length > 0;

            var run = new Run(texto.Contenido)
            {
                FontWeight = texto.Negrita || negritaBase ? FontWeight.Bold : FontWeight.Normal,
                FontStyle = texto.Cursiva ? FontStyle.Italic : FontStyle.Normal,
                TextDecorations = texto.Subrayado ? TextDecorations.Underline : null,
                FontSize = texto.TamanoPunto ?? tamanoBase,
            };

            if (texto.ColorHex is not null && TryParseColorHex(texto.ColorHex, out var color))
            {
                run.Foreground = new SolidColorBrush(color);
            }
            else if (parrafo.NivelEncabezado is not null)
            {
                run.Foreground = new SolidColorBrush(ColorEncabezadoPorDefecto);
            }

            bloque.Inlines!.Add(run);
        }

        if (!tieneTextoVisible)
        {
            // Párrafo vacío en el .docx: no tiene texto que mostrar, pero sí
            // representa un salto de línea deliberado (p. ej. la línea en
            // blanco que separa una sección de la siguiente). Sin darle una
            // altura mínima, un TextBlock sin contenido colapsa a 0px y ese
            // salto de línea desaparece de la vista, aunque en Word real
            // deja un espacio en blanco visible.
            bloque.MinHeight = tamanoBase * 1.3;
        }

        return bloque;
    }

    /// <summary>
    /// Renderiza una <see cref="Tabla"/> como una grilla con bordes, igual
    /// que la muestra Word: cada celda es un <see cref="Border"/> con su
    /// propio párrafo (o párrafos) adentro. El color de fondo de cada
    /// celda sale de <see cref="CeldaTabla.ColorDeFondoHex"/> — el
    /// <c>w:shd</c> real que trae el <c>.docx</c> — en vez de asumir por
    /// posición que "la primera fila es el encabezado y va en azul": esa
    /// regla pintaba de azul con texto blanco tablas que en el documento
    /// original no tienen ninguna fila de encabezado, solo texto en negrita
    /// al inicio de cada celda (que ya viene resuelto por
    /// <see cref="CrearBloqueDeParrafo"/> desde el formato real del run).
    /// </summary>
    private static Border CrearBloqueDeTabla(Tabla tabla)
    {
        var cantidadDeColumnas = tabla.CantidadDeColumnas;

        var grilla = new Grid();
        for (var c = 0; c < cantidadDeColumnas; c++)
        {
            grilla.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));
        }

        for (var f = 0; f < tabla.Filas.Count; f++)
        {
            grilla.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

            var fila = tabla.Filas[f];

            for (var c = 0; c < fila.Celdas.Count; c++)
            {
                var celda = CrearBloqueDeCelda(fila.Celdas[c]);
                Grid.SetRow(celda, f);
                Grid.SetColumn(celda, c);
                grilla.Children.Add(celda);
            }
        }

        return new Border
        {
            BorderBrush = new SolidColorBrush(ColorBordeTabla),
            BorderThickness = new Thickness(1, 1, 0, 0),
            Child = grilla,
            Margin = new Thickness(0, 4),
        };
    }

    private static Border CrearBloqueDeCelda(CeldaTabla celda)
    {
        var contenido = new StackPanel { Spacing = 2 };

        foreach (var parrafo in celda.Parrafos)
        {
            contenido.Children.Add(CrearBloqueDeParrafo(parrafo));
        }

        IBrush fondo = celda.ColorDeFondoHex is not null && TryParseColorHex(celda.ColorDeFondoHex, out var colorDeFondo)
            ? new SolidColorBrush(colorDeFondo)
            : Brushes.Transparent;

        return new Border
        {
            BorderBrush = new SolidColorBrush(ColorBordeTabla),
            BorderThickness = new Thickness(0, 0, 1, 1),
            Background = fondo,
            Padding = new Thickness(8, 6),
            Child = contenido,
        };
    }

    /// <summary>
    /// Tamaño y negrita por defecto según el nivel de encabezado, para
    /// cuando el estilo no trajo un tamaño explícito resuelto (algunas
    /// plantillas lo definen vía tema en vez de en el propio estilo).
    /// Aproxima los tamaños por defecto de Word para Heading1-4+ y texto
    /// normal (11pt) — el fallback de texto normal decía "11pt" en este
    /// mismo comentario pero el valor usado era 14pt: por eso el cuerpo del
    /// texto (y las celdas de tabla, que casi nunca traen tamaño propio)
    /// se veían más grandes de lo que Word realmente muestra.
    /// </summary>
    private static (double TamanoPunto, bool Negrita) TamanoYNegritaPorEncabezado(int? nivel) => nivel switch
    {
        1 => (28.0, true),
        2 => (22.0, true),
        3 => (18.0, true),
        4 => (15.0, true),
        >= 5 => (13.0, true),
        _ => (11.0, false),
    };

    private static TextAlignment ConvertirAlineacion(AlineacionTexto alineacion) => alineacion switch
    {
        AlineacionTexto.Centro => TextAlignment.Center,
        AlineacionTexto.Derecha => TextAlignment.Right,
        AlineacionTexto.Justificado => TextAlignment.Justify,
        _ => TextAlignment.Left,
    };

    private static double PuntosAPixeles(double puntos) => puntos * 96.0 / 72.0;

    private static bool TryParseColorHex(string colorHex, out Color color)
    {
        try
        {
            color = Color.Parse(colorHex.StartsWith('#') ? colorHex : $"#{colorHex}");
            return true;
        }
        catch
        {
            color = default;
            return false;
        }
    }
}
