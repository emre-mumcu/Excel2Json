using System.Text;
using System.Text.Json;
using ExcelToJson.Core.Export;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Profiles;

/// <summary>Profilleri bir klasörde "&lt;ad&gt;.profile.json" dosyaları olarak saklar.</summary>
public sealed class FileProfileStore : IProfileStore
{
    private const string Extension = ".profile.json";
    private static readonly UTF8Encoding Utf8NoBom = new(false);

    public FileProfileStore(string directory) => Directory = directory;

    public string Directory { get; }

    public static string DefaultDirectory =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelToJson", "profiles");

    private string PathFor(string name) => Path.Combine(Directory, FileNameTemplate.Sanitize(name) + Extension);

    public async Task<IReadOnlyList<ProfileSummary>> ListAsync(CancellationToken cancellationToken = default)
    {
        if (!System.IO.Directory.Exists(Directory))
            return [];

        var result = new List<ProfileSummary>();
        foreach (var file in System.IO.Directory.EnumerateFiles(Directory, "*" + Extension))
        {
            try
            {
                await using var fs = File.OpenRead(file);
                using var doc = await JsonDocument.ParseAsync(fs, cancellationToken: cancellationToken);
                var root = doc.RootElement;
                string? Str(string p) => root.TryGetProperty(p, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;
                var name = Str("name") ?? Path.GetFileName(file)[..^Extension.Length];
                var updated = root.TryGetProperty("updatedAt", out var u) && u.TryGetDateTime(out var d) ? d : File.GetLastWriteTimeUtc(file);
                result.Add(new ProfileSummary(name, file, Str("sourceFileName"), Str("sheetName"), updated));
            }
            catch (Exception ex) when (ex is JsonException or IOException)
            {
                // Bozuk profil listede gösterilmez.
            }
        }
        return result.OrderByDescending(p => p.UpdatedAt).ToList();
    }

    public async Task<MappingProfile> LoadAsync(string name, CancellationToken cancellationToken = default)
    {
        var path = PathFor(name);
        if (!File.Exists(path))
            throw new FileNotFoundException($"\"{name}\" profili bulunamadı.", path);
        return await ImportFromFileAsync(path, cancellationToken);
    }

    public async Task SaveAsync(MappingProfile profile, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(profile.Name))
            throw new ArgumentException("Profil adı boş olamaz.", nameof(profile));
        System.IO.Directory.CreateDirectory(Directory);
        profile.UpdatedAt = DateTime.UtcNow;
        await ExportToFileAsync(profile, PathFor(profile.Name), cancellationToken);
    }

    public Task DeleteAsync(string name, CancellationToken cancellationToken = default)
    {
        var path = PathFor(name);
        if (File.Exists(path))
            File.Delete(path);
        return Task.CompletedTask;
    }

    public async Task<ProfileSummary?> FindForSourceAsync(string sourceFileName, string sheetName, CancellationToken cancellationToken = default)
    {
        var all = await ListAsync(cancellationToken);
        return all.FirstOrDefault(p =>
            string.Equals(p.SourceFileName, sourceFileName, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(p.SheetName, sheetName, StringComparison.OrdinalIgnoreCase));
    }

    public async Task ExportToFileAsync(MappingProfile profile, string path, CancellationToken cancellationToken = default)
    {
        var json = ProfileSerializer.Serialize(profile);
        var temp = path + ".tmp";
        await File.WriteAllTextAsync(temp, json, Utf8NoBom, cancellationToken);
        File.Move(temp, path, overwrite: true);
    }

    public async Task<MappingProfile> ImportFromFileAsync(string path, CancellationToken cancellationToken = default)
    {
        var json = await File.ReadAllTextAsync(path, cancellationToken);
        return ProfileSerializer.Deserialize(json);
    }
}
