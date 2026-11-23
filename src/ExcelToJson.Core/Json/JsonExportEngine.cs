using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelToJson.Core.Conversion;
using ExcelToJson.Core.Excel;
using ExcelToJson.Core.Json.Shapes;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Json;

public readonly record struct ExportProgress(int Processed, int Total)
{
    public double Fraction => Total <= 0 ? 0 : Math.Clamp((double)Processed / Total, 0, 1);
}

public sealed class ExportRunOptions
{
    /// <summary>Önizleme için en fazla üretilecek kayıt sayısı. Null ise tümü.</summary>
    public int? MaxRecords { get; init; }
    public IProgress<ExportProgress>? Progress { get; init; }
    public DateTime? Now { get; init; }
    /// <summary>Format ayarlarını geçersiz kılmak için (önizleme her zaman girintili gösterilir).</summary>
    public bool? ForceIndented { get; init; }
}

/// <summary>Bir JSON üretim çalışmasının sonucu.</summary>
public sealed class ExportRunResult
{
    public int RecordCount { get; set; }
    public int TotalRows { get; set; }
    public int SkippedEmptyRows { get; set; }
    public int SkippedHiddenRows { get; set; }
    public int MissingFormulaValues { get; set; }
    public IReadOnlyList<ConversionIssue> Issues { get; set; } = [];
    public int IssueCount { get; set; }
    public List<string> Warnings { get; } = [];
    /// <summary>"Export'u durdur" politikasında dönüşüm hatası oluştu.</summary>
    public bool Aborted { get; set; }
    /// <summary>Şekil hatası gibi JSON üretimini engelleyen hata.</summary>
    public string? FatalError { get; set; }
    public bool Succeeded => !Aborted && FatalError is null;
}

/// <summary>
/// Excel satırlarını profile göre JSON'a dönüştürür ve <see cref="Utf8JsonWriter"/> ile doğrudan stream'e yazar.
/// Tüm JSON hiçbir zaman string olarak bellekte tutulmaz (gruplama gibi tampon gerektiren şekiller hariç).
/// </summary>
public sealed class JsonExportEngine(IValueConverter converter, OutputShapeRegistry shapes)
{
    public OutputShapeRegistry Shapes { get; } = shapes;

    public static JsonExportEngine CreateDefault() => new(new ValueConverter(), OutputShapeRegistry.CreateDefault());

    public static JsonWriterOptions CreateWriterOptions(ExportFormatOptions format, bool? forceIndented = null) => new()
    {
        Indented = forceIndented ?? format.Indented,
        IndentSize = Math.Clamp(format.IndentSize, 1, 8),
        Encoder = format.UnescapedUnicode ? JavaScriptEncoder.UnsafeRelaxedJsonEscaping : JavaScriptEncoder.Default
    };

    /// <summary>Profilin JSON çıktısını (BOM dahil) stream'e yazar.</summary>
    public ExportRunResult Run(Stream output, IWorkbookSession session, MappingProfile profile, ExportRunOptions options,
        CancellationToken cancellationToken = default)
    {
        if (profile.Format.WriteBom)
            output.Write(Encoding.UTF8.Preamble);

        using var writer = new Utf8JsonWriter(output, CreateWriterOptions(profile.Format, options.ForceIndented));
        var result = WriteValue(writer, session, profile, options, cancellationToken);
        writer.Flush();
        return result;
    }

    /// <summary>
    /// Profilin çıktısını tek bir JSON değeri olarak verilen writer'a yazar. Birden çok sayfayı
    /// tek dosyada birleştirirken (<c>{ "sayfa": … }</c>) kullanılır.
    /// </summary>
    public ExportRunResult WriteValue(Utf8JsonWriter writer, IWorkbookSession session, MappingProfile profile, ExportRunOptions options,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(profile.SheetName, nameof(profile.SheetName));

        var result = new ExportRunResult();
        var context = new ExportContext(options.Now ?? DateTime.Now, session.Info.FileName, profile.SheetName);
        var plan = ConversionPlan.Create(profile, converter, context);
        var strategy = Shapes.Get(profile.Output.Shape);
        var issues = new IssueCollector();
        var stats = new RowReadStats();
        var shapeContext = new ShapeContext(profile, plan, issues, cancellationToken);
        var wrapper = profile.Output.Wrapper;

        // Toplam satır sayısı (ilerleme, meta.count ve önizleme özeti için). Cache üzerinden hızlıdır; dönüşüm yapılmaz.
        var needTotal = options.Progress is not null || options.MaxRecords is not null
                        || (wrapper.Enabled && (wrapper.IncludeMeta || wrapper.RootFields.Count > 0));
        RowReadStats? countStats = null;
        if (needTotal)
        {
            countStats = new RowReadStats();
            result.TotalRows = session.ReadRows(profile.SheetName, profile.ReadOptions, countStats, cancellationToken).Count();
        }

        var limit = strategy.RespectsPreviewLimit ? options.MaxRecords : null;

        IEnumerable<Record> Records()
        {
            var sequence = 0;
            foreach (var row in session.ReadRows(profile.SheetName, profile.ReadOptions, stats, cancellationToken))
            {
                if (limit is { } max && sequence >= max)
                    yield break;
                sequence++;
                result.RecordCount = sequence;
                if ((sequence & 0x1FF) == 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    options.Progress?.Report(new ExportProgress(sequence, result.TotalRows));
                }
                yield return plan.BuildRecord(row, sequence, issues);
            }
        }

        try
        {
            if (wrapper.Enabled)
            {
                writer.WriteStartObject();
                var rootFields = BuildRootFields(plan, wrapper, result.TotalRows, issues);
                if (wrapper.IncludeMeta)
                {
                    var meta = new JsonObject
                    {
                        ["source"] = context.SourceFileName,
                        ["sheet"] = context.SheetName,
                        ["exportedAt"] = context.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", System.Globalization.CultureInfo.InvariantCulture),
                        ["count"] = result.TotalRows
                    };
                    Merge(meta, rootFields);
                    writer.WritePropertyName(string.IsNullOrWhiteSpace(wrapper.MetaPropertyName) ? "meta" : wrapper.MetaPropertyName);
                    meta.WriteTo(writer);
                }
                else
                {
                    foreach (var (name, value) in rootFields.ToList())
                    {
                        rootFields.Remove(name);
                        writer.WritePropertyName(name);
                        ShapeHelpers.WriteNode(writer, value);
                    }
                }

                writer.WritePropertyName(string.IsNullOrWhiteSpace(wrapper.DataPropertyName) ? "data" : wrapper.DataPropertyName);
                strategy.Write(writer, Records(), shapeContext);
                writer.WriteEndObject();
            }
            else
            {
                strategy.Write(writer, Records(), shapeContext);
            }
        }
        catch (ShapeException ex)
        {
            result.FatalError = ex.Message;
        }

        options.Progress?.Report(new ExportProgress(result.RecordCount, Math.Max(result.TotalRows, result.RecordCount)));

        // Tam sayım yapıldıysa istatistikler tüm sayfayı yansıtır (önizlemede yalnızca ilk N kayıt okunur).
        var fullStats = countStats ?? stats;
        result.SkippedEmptyRows = fullStats.SkippedEmptyRows;
        result.SkippedHiddenRows = fullStats.SkippedHiddenRows;
        result.MissingFormulaValues = fullStats.MissingFormulaValues;
        stats = fullStats;
        result.Issues = issues.Issues;
        result.IssueCount = issues.TotalCount;
        result.Warnings.AddRange(shapeContext.Warnings);
        if (stats.MissingFormulaValues > 0)
            result.Warnings.Add($"{stats.MissingFormulaValues} formül hücresinin hesaplanmış değeri yok. Dosyayı Excel'de açıp kaydederek değerleri güncelleyin.");
        result.Aborted = profile.ErrorPolicy == ConversionErrorPolicy.Stop && issues.TotalCount > 0;
        return result;
    }

    /// <summary>Önizleme: ilk N kaydı girintili JSON metni olarak üretir.</summary>
    public (string Json, ExportRunResult Result) Preview(IWorkbookSession session, MappingProfile profile, int maxRecords,
        CancellationToken cancellationToken = default)
    {
        using var ms = new MemoryStream();
        var previewProfile = profile.Clone();
        previewProfile.Format.WriteBom = false;
        var result = Run(ms, session, previewProfile, new ExportRunOptions { MaxRecords = maxRecords, ForceIndented = true }, cancellationToken);
        return (Encoding.UTF8.GetString(ms.GetBuffer(), 0, (int)ms.Length), result);
    }

    private static JsonObject BuildRootFields(ConversionPlan plan, WrapperOptions wrapper, int count, IssueCollector issues)
    {
        var obj = new JsonObject();
        foreach (var field in wrapper.RootFields.Where(f => f.Include && !string.IsNullOrWhiteSpace(f.Alias)))
        {
            var value = plan.EvaluateRootField(field, count, issues);
            var path = field.Alias.Split('.');
            var target = obj;
            for (var i = 0; i < path.Length - 1; i++)
            {
                if (target[path[i]] is not JsonObject child)
                    target[path[i]] = child = new JsonObject();
                target = child;
            }
            target[path[^1]] = value;
        }
        return obj;
    }

    private static void Merge(JsonObject target, JsonObject source)
    {
        foreach (var (name, value) in source.ToList())
        {
            source.Remove(name);
            target[name] = value;
        }
    }
}
