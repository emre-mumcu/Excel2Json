using System.Text;
using ClosedXML.Excel;
using ExcelToJson.Core.Excel;
using ExcelToJson.Core.Json;
using ExcelToJson.Core.Mapping;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Tests;

/// <summary>Testler için geçici Excel dosyaları üretir (repo'da binary dosya tutulmaz).</summary>
public sealed class TempWorkbook : IDisposable
{
    public string Path { get; }

    private TempWorkbook(string path) => Path = path;

    public static TempWorkbook Create(Action<XLWorkbook> build, string extension = ".xlsx")
    {
        var dir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ExcelToJsonTests");
        Directory.CreateDirectory(dir);
        var path = System.IO.Path.Combine(dir, Guid.NewGuid().ToString("N") + extension);
        using var wb = new XLWorkbook();
        build(wb);
        wb.SaveAs(path);
        return new TempWorkbook(path);
    }

    /// <summary>Tek sayfalı tablo: ilk satır başlık, sonrakiler veri.</summary>
    public static TempWorkbook FromRows(string sheetName, params object?[][] rows) => Create(wb =>
    {
        var ws = wb.AddWorksheet(sheetName);
        for (var r = 0; r < rows.Length; r++)
            for (var c = 0; c < rows[r].Length; c++)
                SetValue(ws.Cell(r + 1, c + 1), rows[r][c]);
    });

    public static void SetValue(IXLCell cell, object? value)
    {
        switch (value)
        {
            case null: break;
            case string s: cell.Value = s; break;
            case int i: cell.Value = i; break;
            case long l: cell.Value = l; break;
            case double d: cell.Value = d; break;
            case decimal m: cell.Value = m; break;
            case bool b: cell.Value = b; break;
            case DateTime dt: cell.Value = dt; break;
            default: cell.Value = value.ToString(); break;
        }
    }

    public void Dispose()
    {
        try { File.Delete(Path); } catch { /* yok say */ }
    }
}

internal static class TestExtensions
{
    public static IWorkbookSession Open(this TempWorkbook wb) =>
        new ClosedXmlExcelReader().OpenAsync(wb.Path).GetAwaiter().GetResult();

    /// <summary>Şemadan varsayılan profil oluşturur.</summary>
    public static MappingProfile CreateProfile(this IWorkbookSession session, string? sheet = null, SheetReadOptions? options = null)
    {
        sheet ??= session.Info.Sheets[0].Name;
        options ??= new SheetReadOptions();
        var schema = session.GetSchema(sheet, options);
        return new MappingProfile
        {
            SheetName = sheet,
            ReadOptions = options,
            Fields = FieldFactory.CreateColumnFields(schema, NamingStyle.CamelCase, true),
            Format = new ExportFormatOptions { Indented = false }
        };
    }

    public static (string Json, ExportRunResult Result) RunJson(this IWorkbookSession session, MappingProfile profile, int? max = null)
    {
        using var ms = new MemoryStream();
        var result = JsonExportEngine.CreateDefault().Run(ms, session, profile,
            new ExportRunOptions { MaxRecords = max, Now = new DateTime(2026, 11, 23, 10, 0, 0, DateTimeKind.Utc) });
        return (Encoding.UTF8.GetString(ms.ToArray()), result);
    }

    public static T Field<T>(this MappingProfile profile, string alias) where T : FieldDefinition =>
        profile.Fields.OfType<T>().Single(f => f.Alias == alias);
}
