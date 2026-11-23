using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelToJson.Core.Export;

/// <summary>
/// Dosya adı şablonu. Desteklenen placeholder'lar:
/// {file} (uzantısız kaynak dosya adı), {sheet}, {profile}, {date} (yyyy-MM-dd), {time} (HHmmss)
/// ve diğer tüm {…} ifadeleri tarih formatı olarak yorumlanır: {yyyyMMdd}, {yyyyMMdd_HHmmss}…
/// </summary>
public static partial class FileNameTemplate
{
    public const string Default = "{sheet}_{yyyyMMdd_HHmmss}.json";

    [GeneratedRegex(@"\{([^{}]+)\}")]
    private static partial Regex PlaceholderRegex();

    public static string Resolve(string? template, string sourceFile, string sheet, DateTime now, string? profileName = null)
    {
        template = string.IsNullOrWhiteSpace(template) ? Default : template.Trim();
        var result = PlaceholderRegex().Replace(template, m =>
        {
            var key = m.Groups[1].Value;
            return key.ToLowerInvariant() switch
            {
                "file" => Path.GetFileNameWithoutExtension(sourceFile),
                "sheet" => sheet,
                "profile" => profileName ?? string.Empty,
                "date" => now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                "time" => now.ToString("HHmmss", CultureInfo.InvariantCulture),
                _ => FormatDate(now, key, m.Value)
            };
        });

        result = Sanitize(result);
        if (!result.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
            result += ".json";
        return result;
    }

    private static string FormatDate(DateTime now, string format, string fallback)
    {
        // Yalnızca tarih/saat format karakterlerinden oluşan ifadeler tarih olarak yorumlanır.
        if (!format.All(ch => "yMdHhmsf_-. ".Contains(ch)))
            return fallback.Trim('{', '}');
        try
        {
            return now.ToString(format, CultureInfo.InvariantCulture);
        }
        catch (FormatException)
        {
            return fallback;
        }
    }

    /// <summary>Geçersiz dosya adı karakterlerini '_' ile değiştirir.</summary>
    public static string Sanitize(string fileName)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var sb = new StringBuilder(fileName.Length);
        foreach (var ch in fileName)
            sb.Append(invalid.Contains(ch) || ch is ':' or '*' or '?' or '"' or '<' or '>' or '|' or '/' or '\\' ? '_' : ch);
        var s = sb.ToString().Trim().TrimEnd('.');
        return s.Length == 0 ? "export" : s;
    }

    /// <summary>Dosya varsa "ad (2).json", "ad (3).json"… şeklinde boş bir ad bulur.</summary>
    public static string GetAvailablePath(string path)
    {
        if (!File.Exists(path))
            return path;
        var dir = Path.GetDirectoryName(path) ?? string.Empty;
        var name = Path.GetFileNameWithoutExtension(path);
        var ext = Path.GetExtension(path);
        for (var i = 2; ; i++)
        {
            var candidate = Path.Combine(dir, $"{name} ({i}){ext}");
            if (!File.Exists(candidate))
                return candidate;
        }
    }
}
