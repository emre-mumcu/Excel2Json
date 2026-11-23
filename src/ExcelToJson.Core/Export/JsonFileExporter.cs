using ExcelToJson.Core.Excel;
using ExcelToJson.Core.Json;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Export;

public sealed class FileExportResult
{
    public required ExportRunResult Run { get; init; }
    /// <summary>Başarılıysa yazılan dosyanın yolu.</summary>
    public string? FilePath { get; init; }
    public long FileSize { get; init; }
    public TimeSpan Duration { get; init; }
    public bool Succeeded => FilePath is not null && Run.Succeeded;
}

/// <summary>
/// JSON'u dosyaya atomik olarak yazar: önce aynı klasörde geçici bir dosyaya yazılır,
/// başarıyla tamamlanınca hedefe taşınır. Hata/iptal durumunda hedef dosyaya dokunulmaz.
/// </summary>
public sealed class JsonFileExporter(JsonExportEngine engine)
{
    public Task<FileExportResult> ExportAsync(IWorkbookSession session, MappingProfile profile, string targetPath,
        IProgress<ExportProgress>? progress = null, CancellationToken cancellationToken = default) =>
        Task.Run(() => Export(session, profile, targetPath, progress, cancellationToken), cancellationToken);

    public FileExportResult Export(IWorkbookSession session, MappingProfile profile, string targetPath,
        IProgress<ExportProgress>? progress = null, CancellationToken cancellationToken = default)
    {
        var started = DateTime.UtcNow;
        var directory = Path.GetDirectoryName(Path.GetFullPath(targetPath))!;
        Directory.CreateDirectory(directory);
        var tempPath = Path.Combine(directory, $".{Path.GetFileName(targetPath)}.{Guid.NewGuid():N}.tmp");

        ExportRunResult run;
        try
        {
            using (var stream = new FileStream(tempPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 1 << 16))
            {
                run = engine.Run(stream, session, profile, new ExportRunOptions { Progress = progress }, cancellationToken);
            }

            if (!run.Succeeded)
            {
                TryDelete(tempPath);
                return new FileExportResult { Run = run, Duration = DateTime.UtcNow - started };
            }

            File.Move(tempPath, targetPath, overwrite: true);
        }
        catch
        {
            TryDelete(tempPath);
            throw;
        }

        return new FileExportResult
        {
            Run = run,
            FilePath = targetPath,
            FileSize = new FileInfo(targetPath).Length,
            Duration = DateTime.UtcNow - started
        };
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch
        {
            // Geçici dosya silinemezse yapılacak bir şey yok.
        }
    }
}
