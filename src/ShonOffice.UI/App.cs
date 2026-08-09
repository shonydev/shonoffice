using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Themes.Fluent;
using AvaloniaApplication = Avalonia.Application;

namespace ShonOffice.UI;

/// <summary>
/// Se define todo en C# (sin archivos <c>.axaml</c>) a propósito: evita que
/// el build dependa del compilador de markup de Avalonia, que no se puede
/// validar en este entorno sin conexión a NuGet. Funcionalmente es
/// equivalente a un <c>App.axaml</c> con <c>&lt;FluentTheme/&gt;</c>.
/// </summary>
public sealed class App : AvaloniaApplication
{
    public override void Initialize()
    {
        Styles.Add(new FluentTheme());
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime escritorio)
        {
            escritorio.MainWindow = new MainWindow();
        }

        base.OnFrameworkInitializationCompleted();
    }
}
