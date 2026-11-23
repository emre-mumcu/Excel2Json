using ExcelToJson.Core.Export;
using ExcelToJson.Core.Json;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Tests;

public class MultiSheetExportTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "ExcelToJsonTests", Guid.NewGuid().ToString("N"));

    private readonly TempWorkbook _wb = TempWorkbook.Create(w =>
    {
        var a = w.AddWorksheet("Ürünler");
        a.Cell("A1").Value = "Kod"; a.Cell("B1").Value = "Fiyat";
        a.Cell("A2").Value = "P1"; a.Cell("B2").Value = 10;
        a.Cell("A3").Value = "P2"; a.Cell("B3").Value = 20;
        var s = w.AddWorksheet("Ayarlar");
        s.Cell("A1").Value = "Anahtar"; s.Cell("B1").Value = "Değer";
        s.Cell("A2").Value = "timeout"; s.Cell("B2").Value = 30;
        s.Cell("A3").Value = "theme"; s.Cell("B3").Value = "dark";
    });

    public void Dispose()
    {
        _wb.Dispose();
        if (Directory.Exists(_dir))
            Directory.Delete(_dir, true);
    }

    private static MultiSheetExporter CreateExporter()
    {
        var engine = JsonExportEngine.CreateDefault();
        return new MultiSheetExporter(engine, new JsonFileExporter(engine));
    }

    private static MappingProfile Defaults() => new() { Format = new ExportFormatOptions { Indented = false } };

    /// <summary>Ayarlar sayfası için anahtar–değer profili (başka bir dosyadan kaydedilmiş gibi).</summary>
    private MappingProfile KeyValueProfile(Excel.IWorkbookSession session)
    {
        var p = session.CreateProfile("Ayarlar");
        p.Output.Shape = OutputShapes.KeyValue;
        p.Output.KeyFieldId = p.Field<ColumnField>("anahtar").Id;
        p.Output.ValueFieldId = p.Field<ColumnField>("deger").Id;
        return p;
    }

    [Fact]
    public void Combined_file_uses_each_sheets_own_profile()
    {
        using var session = _wb.Open();
        var jobs = new[]
        {
            MultiSheetExporter.PrepareJob(session, "Ürünler", null, Defaults()),
            MultiSheetExporter.PrepareJob(session, "Ayarlar", KeyValueProfile(session), Defaults())
        };
        var target = Path.Combine(_dir, "hepsi.json");
        var progress = new List<MultiSheetProgress>();

        var result = CreateExporter().ExportCombined(session, jobs, target, new ExportFormatOptions { Indented = false },
            new SyncProgress<MultiSheetProgress>(progress.Add));

        Assert.True(result.Succeeded);
        Assert.Equal(target, result.CombinedFilePath);
        Assert.Equal(4, result.TotalRecords);
        Assert.Equal("""{"Ürünler":[{"kod":"P1","fiyat":10},{"kod":"P2","fiyat":20}],"Ayarlar":{"timeout":30,"theme":"dark"}}""",
            File.ReadAllText(target));
        Assert.Contains(progress, p => p.SheetIndex == 1 && p.SheetName == "Ayarlar");
        Assert.Equal(1.0, progress[^1].Fraction, 3);
    }

    [Fact]
    public void Combined_file_is_not_written_when_a_sheet_fails()
    {
        using var session = _wb.Open();
        var failing = session.CreateProfile("Ürünler");
        failing.Field<ColumnField>("kod").Type = DataType.Integer; // "P1" → hata, politika: durdur
        var jobs = new[]
        {
            MultiSheetExporter.PrepareJob(session, "Ayarlar", null, Defaults()),
            MultiSheetExporter.PrepareJob(session, "Ürünler", failing, Defaults())
        };
        var target = Path.Combine(_dir, "hepsi.json");

        var result = CreateExporter().ExportCombined(session, jobs, target, new ExportFormatOptions());

        Assert.False(result.Succeeded);
        Assert.Null(result.CombinedFilePath);
        Assert.True(result.Sheets[1].Run.Aborted);
        Assert.False(File.Exists(target));
        Assert.Empty(Directory.GetFiles(_dir));
    }

    [Fact]
    public void Separate_files_continue_after_a_failing_sheet()
    {
        using var session = _wb.Open();
        var failing = session.CreateProfile("Ürünler");
        failing.Field<ColumnField>("kod").Type = DataType.Integer;
        failing.Format.FileNameTemplate = "{sheet}.json";
        var defaults = Defaults();
        defaults.Format.FileNameTemplate = "{sheet}.json";
        var jobs = new[]
        {
            MultiSheetExporter.PrepareJob(session, "Ürünler", failing, defaults),
            MultiSheetExporter.PrepareJob(session, "Ayarlar", KeyValueProfile(session), defaults)
        };
        jobs[1].Profile.Format.FileNameTemplate = "{sheet}.json";
        var paths = MultiSheetExporter.ResolveSeparatePaths(_dir, "kitap.xlsx", jobs, DateTime.Now);

        var result = CreateExporter().ExportSeparate(session, jobs, paths);

        Assert.False(result.Succeeded);
        Assert.Equal(1, result.SucceededCount);
        Assert.Null(result.Sheets[0].FilePath);
        Assert.Equal(Path.Combine(_dir, "Ayarlar.json"), result.Sheets[1].FilePath);
        Assert.Equal("""{"timeout":30,"theme":"dark"}""", File.ReadAllText(result.Sheets[1].FilePath!).ReplaceLineEndings("").Replace(" ", ""));
        Assert.Single(Directory.GetFiles(_dir));
    }

    [Fact]
    public void Separate_paths_are_unique_within_a_batch()
    {
        using var session = _wb.Open();
        var defaults = Defaults();
        defaults.Format.FileNameTemplate = "{file}.json"; // iki sayfa da aynı ada çözülür
        var jobs = new[]
        {
            MultiSheetExporter.PrepareJob(session, "Ürünler", null, defaults),
            MultiSheetExporter.PrepareJob(session, "Ayarlar", null, defaults)
        };
        var paths = MultiSheetExporter.ResolveSeparatePaths("C:\\out", "kitap.xlsx", jobs, DateTime.Now);
        Assert.Equal(["C:\\out\\kitap.json", "C:\\out\\kitap (2).json"], paths);
    }

    [Fact]
    public void Prepared_profile_is_rematched_to_the_target_sheet_by_header()
    {
        using var session = _wb.Open();
        // Başka sayfa için hazırlanmış profil: "Kod" başlığı Ayarlar sayfasında yok → uyarı
        var profile = session.CreateProfile("Ürünler");
        var job = MultiSheetExporter.PrepareJob(session, "Ayarlar", profile, Defaults());

        Assert.Equal("Ayarlar", job.Profile.SheetName);
        Assert.Contains(job.Warnings, w => w.Contains("Kod"));
    }

    private sealed class SyncProgress<T>(Action<T> report) : IProgress<T>
    {
        public void Report(T value) => report(value);
    }
}
