using System.Text;
using System.Text.Json;
using ExcelToJson.Core.Excel;
using ExcelToJson.Core.Json;
using ExcelToJson.Core.Mapping;
using ExcelToJson.Core.Models;
using ExcelToJson.Core.Profiles;

namespace ExcelToJson.Core.Export;

public enum MultiSheetMode
{
    /// <summary>Her sayfa kendi dosyasına (kendi profilinin dosya adı şablonu ve format ayarlarıyla).</summary>
    SeparateFiles,
    /// <summary>Tek dosya: { "Sayfa1": …, "Sayfa2": … }</summary>
    SingleFile
}

/// <summary>Bir sayfanın, kendi profiliyle export işi.</summary>
public sealed record SheetExportJob(string SheetName, MappingProfile Profile, IReadOnlyList<string> Warnings);

public sealed record MultiSheetProgress(int SheetIndex, int SheetCount, string SheetName, ExportProgress Sheet)
{
    public double Fraction => SheetCount <= 0 ? 0 : Math.Clamp((SheetIndex + Sheet.Fraction) / SheetCount, 0, 1);
}

public sealed class SheetExportOutcome
{
    public required string SheetName { get; init; }
    public required ExportRunResult Run { get; init; }
    /// <summary>Ayrı dosya modunda yazılan dosya (başarısızsa null).</summary>
    public string? FilePath { get; init; }
    public long FileSize { get; init; }
    public bool Succeeded => Run.Succeeded;
}

public sealed class MultiSheetExportResult
{
    public required MultiSheetMode Mode { get; init; }
    public List<SheetExportOutcome> Sheets { get; } = [];
    /// <summary>Tek dosya modunda yazılan dosya (başarısızsa null).</summary>
    public string? CombinedFilePath { get; set; }
    public long CombinedFileSize { get; set; }
    public TimeSpan Duration { get; set; }

    public bool Succeeded => Sheets.Count > 0 && Sheets.All(s => s.Succeeded) && (Mode == MultiSheetMode.SeparateFiles || CombinedFilePath is not null);
    public int SucceededCount => Sheets.Count(s => s.Succeeded);
    public int TotalRecords => Sheets.Where(s => s.Succeeded).Sum(s => s.Run.RecordCount);
    public IEnumerable<string> WrittenFiles => Mode == MultiSheetMode.SingleFile
        ? CombinedFilePath is null ? [] : [CombinedFilePath]
        : Sheets.Where(s => s.FilePath is not null).Select(s => s.FilePath!);
}

/// <summary>
/// Birden fazla sayfayı, her biri kendi profiliyle, ayrı dosyalara veya tek bir dosyada
/// <c>{ "&lt;sayfa&gt;": … }</c> olarak export eder. Her iki modda da yazım atomiktir.
/// </summary>
public sealed class MultiSheetExporter(JsonExportEngine engine, JsonFileExporter fileExporter)
{
    /// <summary>
    /// Bir sayfa için export işi hazırlar. Profil verilmezse <paramref name="defaults"/> ayarlarıyla tüm sütunlar
    /// dahil edilir; verilirse sütunlar sayfanın şemasına başlık adına göre yeniden eşleştirilir.
    /// </summary>
    public static SheetExportJob PrepareJob(IWorkbookSession session, string sheetName, MappingProfile? profile, MappingProfile defaults)
    {
        var sheet = session.GetSheet(sheetName).Name;
        if (profile is null)
        {
            var p = defaults.Clone();
            p.SheetName = sheet;
            p.Fields = [];
            var schema = session.GetSchema(sheet, p.ReadOptions);
            p.Fields = FieldFactory.CreateColumnFields(schema, p.NamingStyle, p.TransliterateTurkish);
            // Varsayılan profilde alan Id'leri yeni olduğu için alan seçimi gerektiren şekiller kullanılamaz.
            if (p.Output.Shape != OutputShapes.ArrayOfObjects && p.Output.Shape != OutputShapes.ArrayOfArrays)
                p.Output.Shape = OutputShapes.ArrayOfObjects;
            return new SheetExportJob(sheet, p, []);
        }

        var source = profile.Clone();
        var reconciled = ProfileReconciler.Reconcile(source, session.GetSchema(sheet, source.ReadOptions));
        return new SheetExportJob(sheet, reconciled.Profile, reconciled.Warnings);
    }

    /// <summary>
    /// Ayrı dosya modunda her işin hedef yolunu kendi profilinin şablonuyla çözer.
    /// Aynı toplu işlemde iki sayfa aynı ada çözülürse sonrakilere numara eklenir.
    /// </summary>
    public static List<string> ResolveSeparatePaths(string folder, string sourceFileName, IEnumerable<SheetExportJob> jobs,
        DateTime now, string? profileName = null)
    {
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var paths = new List<string>();
        foreach (var job in jobs)
        {
            var name = FileNameTemplate.Resolve(job.Profile.Format.FileNameTemplate, sourceFileName, job.SheetName, now, profileName ?? job.Profile.Name);
            var path = Path.Combine(folder, name);
            if (!used.Add(path))
            {
                var stem = Path.GetFileNameWithoutExtension(name);
                for (var i = 2; !used.Add(path = Path.Combine(folder, $"{stem} ({i}).json")); i++) { }
            }
            paths.Add(path);
        }
        return paths;
    }

    public Task<MultiSheetExportResult> ExportSeparateAsync(IWorkbookSession session, IReadOnlyList<SheetExportJob> jobs,
        IReadOnlyList<string> targetPaths, IProgress<MultiSheetProgress>? progress = null, CancellationToken cancellationToken = default) =>
        Task.Run(() => ExportSeparate(session, jobs, targetPaths, progress, cancellationToken), cancellationToken);

    public MultiSheetExportResult ExportSeparate(IWorkbookSession session, IReadOnlyList<SheetExportJob> jobs,
        IReadOnlyList<string> targetPaths, IProgress<MultiSheetProgress>? progress = null, CancellationToken cancellationToken = default)
    {
        if (jobs.Count != targetPaths.Count)
            throw new ArgumentException("Her sayfa için bir hedef yol verilmelidir.", nameof(targetPaths));

        var started = DateTime.UtcNow;
        var result = new MultiSheetExportResult { Mode = MultiSheetMode.SeparateFiles };
        for (var i = 0; i < jobs.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var job = jobs[i];
            var index = i;
            var sheetProgress = progress is null ? null : new InlineProgress(p => progress.Report(new MultiSheetProgress(index, jobs.Count, job.SheetName, p)));

            // Bir sayfanın başarısız olması diğerlerini durdurmaz; her dosya kendi içinde atomiktir.
            var r = fileExporter.Export(session, job.Profile, targetPaths[i], sheetProgress, cancellationToken);
            result.Sheets.Add(new SheetExportOutcome { SheetName = job.SheetName, Run = r.Run, FilePath = r.FilePath, FileSize = r.FileSize });
        }
        result.Duration = DateTime.UtcNow - started;
        return result;
    }

    public Task<MultiSheetExportResult> ExportCombinedAsync(IWorkbookSession session, IReadOnlyList<SheetExportJob> jobs, string targetPath,
        ExportFormatOptions format, IProgress<MultiSheetProgress>? progress = null, CancellationToken cancellationToken = default) =>
        Task.Run(() => ExportCombined(session, jobs, targetPath, format, progress, cancellationToken), cancellationToken);

    /// <summary>
    /// Tüm sayfaları tek dosyaya <c>{ "&lt;sayfa&gt;": &lt;sayfanın çıktısı&gt; }</c> olarak yazar. Her sayfanın değeri kendi
    /// profilinin çıktı şekli ve wrapper ayarlarıyla üretilir; girinti/kodlama <paramref name="format"/>'tan gelir.
    /// Herhangi bir sayfa başarısız olursa dosya yazılmaz.
    /// </summary>
    public MultiSheetExportResult ExportCombined(IWorkbookSession session, IReadOnlyList<SheetExportJob> jobs, string targetPath,
        ExportFormatOptions format, IProgress<MultiSheetProgress>? progress = null, CancellationToken cancellationToken = default)
    {
        var started = DateTime.UtcNow;
        var result = new MultiSheetExportResult { Mode = MultiSheetMode.SingleFile };
        var directory = Path.GetDirectoryName(Path.GetFullPath(targetPath))!;
        Directory.CreateDirectory(directory);
        var tempPath = Path.Combine(directory, $".{Path.GetFileName(targetPath)}.{Guid.NewGuid():N}.tmp");
        var now = DateTime.Now;
        var ok = true;

        try
        {
            using (var stream = new FileStream(tempPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 1 << 16))
            {
                if (format.WriteBom)
                    stream.Write(Encoding.UTF8.Preamble);
                using var writer = new Utf8JsonWriter(stream, JsonExportEngine.CreateWriterOptions(format));
                writer.WriteStartObject();
                for (var i = 0; i < jobs.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var job = jobs[i];
                    var index = i;
                    var options = new ExportRunOptions
                    {
                        Now = now,
                        Progress = progress is null ? null : new InlineProgress(p => progress.Report(new MultiSheetProgress(index, jobs.Count, job.SheetName, p)))
                    };
                    writer.WritePropertyName(job.SheetName);
                    var run = engine.WriteValue(writer, session, job.Profile, options, cancellationToken);
                    result.Sheets.Add(new SheetExportOutcome { SheetName = job.SheetName, Run = run });
                    if (!run.Succeeded)
                    {
                        ok = false;
                        break;
                    }
                }
                if (ok)
                {
                    writer.WriteEndObject();
                    writer.Flush();
                }
            }

            if (!ok)
            {
                TryDelete(tempPath);
            }
            else
            {
                File.Move(tempPath, targetPath, overwrite: true);
                result.CombinedFilePath = targetPath;
                result.CombinedFileSize = new FileInfo(targetPath).Length;
            }
        }
        catch
        {
            TryDelete(tempPath);
            throw;
        }

        result.Duration = DateTime.UtcNow - started;
        return result;
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
            // yok say
        }
    }

    /// <summary>Senkron IProgress (Progress&lt;T&gt; gibi senkronizasyon bağlamına post etmez).</summary>
    private sealed class InlineProgress(Action<ExportProgress> report) : IProgress<ExportProgress>
    {
        public void Report(ExportProgress value) => report(value);
    }
}
