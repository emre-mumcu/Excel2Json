using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Profiles;

public sealed record ProfileSummary(string Name, string FilePath, string? SourceFileName, string? SheetName, DateTime UpdatedAt);

/// <summary>Profil (mapping şablonu) deposu.</summary>
public interface IProfileStore
{
    string Directory { get; }

    Task<IReadOnlyList<ProfileSummary>> ListAsync(CancellationToken cancellationToken = default);
    Task<MappingProfile> LoadAsync(string name, CancellationToken cancellationToken = default);
    Task SaveAsync(MappingProfile profile, CancellationToken cancellationToken = default);
    Task DeleteAsync(string name, CancellationToken cancellationToken = default);

    /// <summary>Aynı dosya + sayfa için en son kaydedilen profil (otomatik öneri).</summary>
    Task<ProfileSummary?> FindForSourceAsync(string sourceFileName, string sheetName, CancellationToken cancellationToken = default);

    Task ExportToFileAsync(MappingProfile profile, string path, CancellationToken cancellationToken = default);
    Task<MappingProfile> ImportFromFileAsync(string path, CancellationToken cancellationToken = default);
}
