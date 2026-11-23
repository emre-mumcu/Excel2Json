using ExcelToJson.Core.Export;

namespace ExcelToJson.App.Services;

/// <summary>Çoklu sayfa export'unda bir sayfanın seçimi ve kullanılacak profil.</summary>
public sealed class MultiSheetItem
{
    public const string CurrentChoice = "current";
    public const string DefaultChoice = "default";
    public const string ProfilePrefix = "profile:";

    public bool Selected { get; set; }

    /// <summary>"current" (açık ayarlar), "default" (tüm sütunlar) veya "profile:&lt;ad&gt;".</summary>
    public string Choice { get; set; } = DefaultChoice;

    public string? ProfileName => Choice.StartsWith(ProfilePrefix, StringComparison.Ordinal) ? Choice[ProfilePrefix.Length..] : null;
}

/// <summary>Export adımındaki çoklu sayfa seçimlerinin durumu (adımlar arasında korunur).</summary>
public sealed class MultiSheetSelection
{
    public Dictionary<string, MultiSheetItem> Items { get; } = new(StringComparer.OrdinalIgnoreCase);
    public MultiSheetMode Mode { get; set; } = MultiSheetMode.SeparateFiles;
    public string CombinedFileNameTemplate { get; set; } = "{file}_{yyyyMMdd_HHmmss}.json";
    public MultiSheetExportResult? LastResult { get; set; }

    public int SelectedCount => Items.Values.Count(i => i.Selected);
}
