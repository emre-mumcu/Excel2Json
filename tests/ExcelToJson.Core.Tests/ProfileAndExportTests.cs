using System.Diagnostics;
using ExcelToJson.Core.Export;
using ExcelToJson.Core.Json;
using ExcelToJson.Core.Models;
using ExcelToJson.Core.Profiles;

namespace ExcelToJson.Core.Tests;

public class ProfileAndExportTests
{
    [Fact]
    public async Task Profile_round_trip_and_rematch_by_header_name()
    {
        using var original = TempWorkbook.FromRows("S", ["Kod", "Ad", "Fiyat"], ["P1", "Laptop", 10]);
        using var session = original.Open();
        var profile = session.CreateProfile();
        profile.Name = "Ürün profili";
        profile.SourceFileName = "urunler.xlsx";
        profile.Field<ColumnField>("fiyat").Type = DataType.Decimal;
        profile.Field<ColumnField>("ad").Alias = "product.name";
        profile.Fields.Add(new CustomField { Alias = "source", Value = "ERP" });
        profile.Output.Shape = OutputShapes.KeyedObject;
        profile.Output.KeyFieldId = profile.Field<ColumnField>("kod").Id;
        profile.Output.Wrapper.Enabled = true;
        profile.Output.Wrapper.RootFields.Add(new CustomField { Alias = "v", Value = "1" });

        var dir = Path.Combine(Path.GetTempPath(), "ExcelToJsonTests", Guid.NewGuid().ToString("N"));
        var store = new FileProfileStore(dir);
        await store.SaveAsync(profile);

        var list = await store.ListAsync();
        Assert.Single(list);
        Assert.Equal("Ürün profili", list[0].Name);
        Assert.NotNull(await store.FindForSourceAsync("URUNLER.xlsx", "s"));

        var loaded = await store.LoadAsync("Ürün profili");
        Assert.Equal(MappingProfile.CurrentSchemaVersion, loaded.SchemaVersion);
        Assert.Equal(profile.Fields.Count, loaded.Fields.Count);
        Assert.IsType<CustomField>(loaded.Fields[^1]);
        Assert.Equal(DataType.Decimal, loaded.Field<ColumnField>("fiyat").Type);
        Assert.Equal(profile.Output.KeyFieldId, loaded.Output.KeyFieldId);
        Assert.Single(loaded.Output.Wrapper.RootFields);

        // Sütunların yeri değişmiş, biri silinmiş, biri eklenmiş yeni dosya
        using var changed = TempWorkbook.FromRows("S", ["Yeni", "Fiyat", "Kod"], ["x", 20, "P9"]);
        using var session2 = changed.Open();
        var schema = session2.GetSchema("S", loaded.ReadOptions);
        var reconciled = ProfileReconciler.Reconcile(loaded, schema);

        Assert.Equal(3, reconciled.Profile.Field<ColumnField>("kod").ColumnIndex);
        Assert.Equal("C", reconciled.Profile.Field<ColumnField>("kod").ColumnLetter);
        Assert.Equal(2, reconciled.Profile.Field<ColumnField>("fiyat").ColumnIndex);
        Assert.Contains(reconciled.Warnings, w => w.Contains("Ad"));
        var added = reconciled.Profile.Fields.OfType<ColumnField>().Single(f => f.Header == "Yeni");
        Assert.False(added.Include);

        var (json, result) = session2.RunJson(reconciled.Profile);
        Assert.True(result.Succeeded);
        Assert.Contains("\"P9\":{\"kod\":\"P9\",\"fiyat\":20,\"source\":\"ERP\"}", json);

        Directory.Delete(dir, true);
    }

    [Fact]
    public void Profile_from_newer_schema_is_rejected()
    {
        Assert.Throws<ProfileFormatException>(() => ProfileSerializer.Deserialize("""{"schemaVersion": 99}"""));
        Assert.Throws<ProfileFormatException>(() => ProfileSerializer.Deserialize("not json"));
    }

    [Fact]
    public void File_name_template_resolution_and_sanitizing()
    {
        var now = new DateTime(2026, 11, 23, 9, 5, 7);
        Assert.Equal("Ürünler_20261123_090507.json", FileNameTemplate.Resolve(null, "urunler.xlsx", "Ürünler", now));
        Assert.Equal("urunler-A_B-2026-11-23.json", FileNameTemplate.Resolve("{file}-{sheet}-{date}", "C:\\x\\urunler.xlsx", "A/B", now));
        Assert.Equal("x_unknown_.json", FileNameTemplate.Resolve("x_{unknown}_", "f.xlsx", "s", now));
        Assert.Equal("20261123.json", FileNameTemplate.Resolve("{yyyyMMdd}.json", "f.xlsx", "s", now));
    }

    [Fact]
    public async Task Atomic_export_writes_file_and_keeps_target_on_failure()
    {
        using var wb = TempWorkbook.FromRows("S", ["Kod", "Miktar"], ["P1", "abc"]);
        using var session = wb.Open();
        var profile = session.CreateProfile();
        var dir = Path.Combine(Path.GetTempPath(), "ExcelToJsonTests", Guid.NewGuid().ToString("N"));
        var target = Path.Combine(dir, "out.json");
        var exporter = new JsonFileExporter(JsonExportEngine.CreateDefault());

        var ok = await exporter.ExportAsync(session, profile, target);
        Assert.True(ok.Succeeded);
        Assert.Equal("""[{"kod":"P1","miktar":"abc"}]""", await File.ReadAllTextAsync(target));

        // Stop politikasında hata → hedef dosya değişmez, geçici dosya kalmaz
        profile.Field<ColumnField>("miktar").Type = DataType.Integer;
        var failed = await exporter.ExportAsync(session, profile, target);
        Assert.False(failed.Succeeded);
        Assert.True(failed.Run.Aborted);
        Assert.Equal("""[{"kod":"P1","miktar":"abc"}]""", await File.ReadAllTextAsync(target));
        Assert.Single(Directory.GetFiles(dir));

        Assert.Equal(Path.Combine(dir, "out (2).json"), FileNameTemplate.GetAvailablePath(target));
        Directory.Delete(dir, true);
    }

    [Fact]
    public async Task Large_sheet_export_is_streamed_and_cancellable()
    {
        const int rows = 100_000;
        using var wb = TempWorkbook.Create(w =>
        {
            var ws = w.AddWorksheet("Big");
            ws.Cell(1, 1).Value = "Kod"; ws.Cell(1, 2).Value = "Ad"; ws.Cell(1, 3).Value = "Fiyat"; ws.Cell(1, 4).Value = "Tarih";
            var data = Enumerable.Range(1, rows).Select(i => new object[] { $"P{i:D6}", $"Ürün {i}", i * 1.5, new DateTime(2026, 1, 1).AddDays(i % 365) });
            ws.Cell(2, 1).InsertData(data);
        });
        using var session = wb.Open();
        var profile = session.CreateProfile("Big");
        var dir = Path.Combine(Path.GetTempPath(), "ExcelToJsonTests", Guid.NewGuid().ToString("N"));
        var exporter = new JsonFileExporter(JsonExportEngine.CreateDefault());
        var progress = new List<ExportProgress>();

        var sw = Stopwatch.StartNew();
        var result = exporter.Export(session, profile, Path.Combine(dir, "big.json"), new SyncProgress(progress.Add));
        sw.Stop();

        Assert.True(result.Succeeded);
        Assert.Equal(rows, result.Run.RecordCount);
        Assert.True(progress.Count > 10);
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(30), $"Export çok yavaş: {sw.Elapsed}");

        using var cts = new CancellationTokenSource();
        cts.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            exporter.ExportAsync(session, profile, Path.Combine(dir, "cancel.json"), null, cts.Token));
        Assert.False(File.Exists(Path.Combine(dir, "cancel.json")));

        Directory.Delete(dir, true);
    }

    private sealed class SyncProgress(Action<ExportProgress> report) : IProgress<ExportProgress>
    {
        public void Report(ExportProgress value) => report(value);
    }
}
