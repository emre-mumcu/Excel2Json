namespace ExcelToJson.App.Services;

/// <summary>Uygulamanın kullandığı klasörler.</summary>
public static class AppPaths
{
    public static string RoamingDirectory { get; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelToJson");

    public static string LocalDirectory { get; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelToJson");

    public static string SettingsFile => Path.Combine(RoamingDirectory, "settings.json");
    public static string ProfileDirectory => Path.Combine(RoamingDirectory, "profiles");
    public static string LogDirectory => Path.Combine(LocalDirectory, "logs");
}
