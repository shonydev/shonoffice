using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Themes.Fluent;
using AvaloniaApplication = Avalonia.Application;

namespace ShonOffice.UI;

/// <summary>
/// Everything is defined in C# (no <c>.axaml</c> files) on purpose: it
/// avoids the build depending on Avalonia's markup compiler, which can't be
/// validated in this environment without a NuGet connection. Functionally
/// it's equivalent to an <c>App.axaml</c> with <c>&lt;FluentTheme/&gt;</c>.
/// </summary>
public sealed class App : AvaloniaApplication
{
    public override void Initialize()
    {
        Styles.Add(new FluentTheme());
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            desktop.MainWindow = new MainWindow();
        }

        base.OnFrameworkInitializationCompleted();
    }
}
