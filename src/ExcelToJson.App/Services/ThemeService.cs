using ExcelToJson.Core.Models;
using Microsoft.FluentUI.AspNetCore.Components;

namespace ExcelToJson.App.Services;

/// <summary>Tema seçimi (Sistem / Açık / Koyu) ve etkin koyu mod durumu.</summary>
public sealed class ThemeService
{
    private readonly SettingsService settings;

    public ThemeService(SettingsService settings)
    {
        this.settings = settings;
        IsDark = ResolveIsDark(Mode);
    }

    public ThemeMode Mode => settings.Current.Theme;

    /// <summary>Fluent tasarım temasının gerçekte koyu olup olmadığı (sistem modunda işletim sistemine bağlıdır).</summary>
    public bool IsDark { get; private set; }

    /// <summary>Windows'un uygulama teması koyu mu? (HKCU\...\Personalize\AppsUseLightTheme = 0)</summary>
    public static bool IsSystemDark()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            return key?.GetValue("AppsUseLightTheme") is int v && v == 0;
        }
        catch
        {
            return false;
        }
    }

    private static bool ResolveIsDark(ThemeMode mode) => mode switch
    {
        ThemeMode.Dark => true,
        ThemeMode.Light => false,
        _ => IsSystemDark()
    };

    public event Action? Changed;

    public DesignThemeModes FluentMode => Mode switch
    {
        ThemeMode.Light => DesignThemeModes.Light,
        ThemeMode.Dark => DesignThemeModes.Dark,
        _ => DesignThemeModes.System
    };

    public void SetMode(ThemeMode mode)
    {
        if (mode == Mode)
            return;
        settings.Update(s => s.Theme = mode);
        IsDark = ResolveIsDark(mode);
        Changed?.Invoke();
    }

    public void ReportLuminance(bool isDark)
    {
        if (IsDark == isDark)
            return;
        IsDark = isDark;
        Changed?.Invoke();
    }
}
