using System.Text;
using System.Text.Json;
using ExcelToJson.Core.Models;
using ExcelToJson.Core.Profiles;
using Microsoft.Extensions.Logging;

namespace ExcelToJson.App.Services;

/// <summary>%AppData%\ExcelToJson\settings.json içinde tutulan uygulama ayarları.</summary>
public sealed class AppSettings
{
    public ThemeMode Theme { get; set; } = ThemeMode.System;
    public NamingStyle DefaultNamingStyle { get; set; } = NamingStyle.CamelCase;
    public bool DefaultTransliterateTurkish { get; set; } = true;
    public string DefaultShape { get; set; } = OutputShapes.ArrayOfObjects;
    public int PreviewRecordCount { get; set; } = 10;
    public bool PreviewPanelOpen { get; set; } = true;
    public List<string> RecentFiles { get; set; } = [];
    public string? LastOutputFolder { get; set; }
    public WindowPlacement? Window { get; set; }
}

public sealed class WindowPlacement
{
    public int X { get; set; }
    public int Y { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public bool Maximized { get; set; }
}

public sealed class SettingsService
{
    public const int MaxRecentFiles = 10;
    private readonly ILogger<SettingsService> _logger;
    private readonly string _path;

    public SettingsService(ILogger<SettingsService> logger) : this(logger, AppPaths.SettingsFile) { }

    public SettingsService(ILogger<SettingsService> logger, string path)
    {
        _logger = logger;
        _path = path;
        Current = Load();
    }

    public AppSettings Current { get; private set; }

    public event Action? Changed;

    private AppSettings Load()
    {
        try
        {
            if (File.Exists(_path))
                return JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(_path), ProfileSerializer.Options) ?? new AppSettings();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Ayarlar okunamadı, varsayılanlar kullanılıyor");
        }
        return new AppSettings();
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
            var temp = _path + ".tmp";
            File.WriteAllText(temp, JsonSerializer.Serialize(Current, ProfileSerializer.Options), new UTF8Encoding(false));
            File.Move(temp, _path, overwrite: true);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Ayarlar kaydedilemedi");
        }
        Changed?.Invoke();
    }

    public void Update(Action<AppSettings> change)
    {
        change(Current);
        Save();
    }

    /// <summary>Son açılan dosyalar; artık var olmayanlar listeden düşürülür.</summary>
    public IReadOnlyList<string> GetRecentFiles()
    {
        var existing = Current.RecentFiles.Where(File.Exists).ToList();
        if (existing.Count != Current.RecentFiles.Count)
        {
            Current.RecentFiles = existing;
            Save();
        }
        return existing;
    }

    public void AddRecentFile(string path)
    {
        Current.RecentFiles.RemoveAll(p => string.Equals(p, path, StringComparison.OrdinalIgnoreCase));
        Current.RecentFiles.Insert(0, path);
        if (Current.RecentFiles.Count > MaxRecentFiles)
            Current.RecentFiles.RemoveRange(MaxRecentFiles, Current.RecentFiles.Count - MaxRecentFiles);
        Save();
    }

    public void RemoveRecentFile(string path)
    {
        Current.RecentFiles.RemoveAll(p => string.Equals(p, path, StringComparison.OrdinalIgnoreCase));
        Save();
    }
}
