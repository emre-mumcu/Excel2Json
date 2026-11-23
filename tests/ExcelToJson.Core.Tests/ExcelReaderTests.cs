using ExcelToJson.Core.Excel;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Tests;

public class ExcelReaderTests
{
    [Fact]
    public void Lists_sheets_with_used_range_and_hidden_flag()
    {
        using var wb = TempWorkbook.Create(w =>
        {
            var a = w.AddWorksheet("Ürünler");
            a.Cell("B2").Value = "Kod";
            a.Cell("D10").Value = "x";
            var h = w.AddWorksheet("Gizli");
            h.Cell("A1").Value = 1;
            h.Visibility = ClosedXML.Excel.XLWorksheetVisibility.Hidden;
            w.AddWorksheet("Boş");
        });
        using var session = wb.Open();

        Assert.Equal(3, session.Info.Sheets.Count);
        var first = session.Info.Sheets[0];
        Assert.Equal("Ürünler", first.Name);
        Assert.Equal("B2:D10", first.UsedRange);
        Assert.Equal(9, first.RowCount);
        Assert.Equal(3, first.ColumnCount);
        Assert.True(session.Info.Sheets[1].IsHidden);
        Assert.True(session.Info.Sheets[2].IsEmpty);
    }

    [Fact]
    public void Uses_selected_header_row_and_data_range()
    {
        using var wb = TempWorkbook.FromRows("S",
            ["Rapor başlığı"],
            [],
            ["Kod", "Ad"],
            ["P1", "Laptop"],
            ["P2", "Mouse"],
            ["P3", "Klavye"]);
        using var session = wb.Open();
        var options = new SheetReadOptions { HeaderRow = 3, DataEndRow = 5 };

        var schema = session.GetSchema("S", options);
        Assert.Equal(["Kod", "Ad"], schema.Columns.Select(c => c.Header));
        Assert.Equal(4, schema.DataStartRow);

        var rows = session.ReadRows("S", options).ToList();
        Assert.Equal([4, 5], rows.Select(r => r.RowNumber));
        Assert.Equal("P2", rows[1][1].Text);
    }

    [Fact]
    public void No_header_uses_column_letters_and_empty_headers_are_named()
    {
        using var wb = TempWorkbook.FromRows("S", ["a", null, "c"], ["1", "2", "3"]);
        using var session = wb.Open();

        var noHeader = session.GetSchema("S", new SheetReadOptions { NoHeader = true });
        Assert.Equal(["A", "B", "C"], noHeader.Columns.Select(c => c.Header));
        Assert.Equal(2, session.ReadRows("S", new SheetReadOptions { NoHeader = true }).Count());

        var withHeader = session.GetSchema("S", new SheetReadOptions());
        Assert.Equal("Column B", withHeader.Columns[1].Header);
    }

    [Fact]
    public void Skips_empty_rows_and_counts_them()
    {
        using var wb = TempWorkbook.FromRows("S", ["Kod"], ["P1"], [], ["  "], ["P2"]);
        using var session = wb.Open();
        var stats = new RowReadStats();

        var rows = session.ReadRows("S", new SheetReadOptions(), stats).ToList();
        Assert.Equal(2, rows.Count);
        Assert.Equal(2, stats.SkippedEmptyRows);

        var all = session.ReadRows("S", new SheetReadOptions { SkipEmptyRows = false }).ToList();
        Assert.Equal(4, all.Count);
    }

    [Fact]
    public void Hidden_rows_and_columns_can_be_excluded()
    {
        using var wb = TempWorkbook.Create(w =>
        {
            var ws = w.AddWorksheet("S");
            ws.Cell("A1").Value = "Kod"; ws.Cell("B1").Value = "Gizli"; ws.Cell("C1").Value = "Ad";
            ws.Cell("A2").Value = "P1"; ws.Cell("C2").Value = "Laptop";
            ws.Cell("A3").Value = "P2"; ws.Cell("C3").Value = "Mouse";
            ws.Row(3).Hide();
            ws.Column(2).Hide();
        });
        using var session = wb.Open();
        var options = new SheetReadOptions { IncludeHiddenRows = false, IncludeHiddenColumns = false };
        var stats = new RowReadStats();

        Assert.Equal(["Kod", "Ad"], session.GetSchema("S", options).Columns.Select(c => c.Header));
        Assert.Single(session.ReadRows("S", options, stats));
        Assert.Equal(1, stats.SkippedHiddenRows);

        Assert.Equal(2, session.ReadRows("S", new SheetReadOptions()).Count());
        Assert.True(session.GetSchema("S", new SheetReadOptions()).Columns[1].IsHidden);
    }

    [Fact]
    public async Task Sheet_with_whole_column_formatting_loads_quickly()
    {
        // Gerçek dosyada görülen durum: A1:B1048576 koşullu biçimlendirme + tüm sütunlara stil.
        // ws.RowsUsed(All) bu durumda ~1 milyon satırı dolaşıp takılıyordu.
        using var wb = TempWorkbook.Create(w =>
        {
            var ws = w.AddWorksheet("S");
            ws.Columns(4, 16384).Style.Font.Bold = true;
            foreach (var article in new[] { "der", "die", "das" })
                ws.Range("A1:B1048576").AddConditionalFormat().WhenIsTrue($"LEFT($A1,4)=\"{article} \"").Font.SetBold();
            ws.Cell("A1").Value = "Kod"; ws.Cell("B1").Value = "Ad";
            for (var r = 2; r <= 50; r++) { ws.Cell(r, 1).Value = $"P{r}"; ws.Cell(r, 2).Value = $"Ürün {r}"; }
            ws.Row(10).Hide();
        });
        using var session = wb.Open();

        var load = Task.Run(() => session.GetSchema("S", new SheetReadOptions()));
        var schema = await load.WaitAsync(TimeSpan.FromSeconds(3),TestContext.Current.CancellationToken);

        Assert.Equal(["Kod", "Ad"], schema.Columns.Select(c => c.Header));
        Assert.Equal(48, session.ReadRows("S", new SheetReadOptions { IncludeHiddenRows = false }).Count());
    }

    [Fact]
    public void Merged_cells_can_be_spread()
    {
        using var wb = TempWorkbook.Create(w =>
        {
            var ws = w.AddWorksheet("S");
            ws.Cell("A1").Value = "Grup"; ws.Cell("B1").Value = "Ürün";
            ws.Cell("A2").Value = "Elektronik";
            ws.Range("A2:A4").Merge();
            ws.Cell("B2").Value = "Laptop"; ws.Cell("B3").Value = "Mouse"; ws.Cell("B4").Value = "Klavye";
        });
        using var session = wb.Open();

        var spread = session.ReadRows("S", new SheetReadOptions { SpreadMergedCells = true }).ToList();
        Assert.All(spread, r => Assert.Equal("Elektronik", r[1].Text));

        var plain = session.ReadRows("S", new SheetReadOptions()).ToList();
        Assert.True(plain[1][1].IsBlank);
    }

    [Fact]
    public async Task Reads_file_locked_by_another_process()
    {
        using var wb = TempWorkbook.FromRows("S", ["Kod"], ["P1"]);
        // Excel dosyayı açtığında yazma kilidi alır ama okumaya izin verir.
        await using var lockStream = new FileStream(wb.Path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);

        using var session = await new ClosedXmlExcelReader().OpenAsync(wb.Path);
        Assert.Single(session.ReadRows("S", new SheetReadOptions()));
    }

    [Fact]
    public async Task Xls_files_give_a_friendly_error()
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xls");
        await File.WriteAllBytesAsync(path, [1, 2, 3]);
        try
        {
            var ex = await Assert.ThrowsAsync<ExcelReadException>(() => new ClosedXmlExcelReader().OpenAsync(path));
            Assert.Contains(".xlsx", ex.Message);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Schema_contains_samples_and_inferred_types()
    {
        using var wb = TempWorkbook.FromRows("S",
            ["Kod", "Fiyat", "Adet", "Aktif", "Tarih", "Barkod"],
            ["001", 10.5, 3, true, new DateTime(2026, 1, 2), 12345678901234567d],
            ["002", 20, 4, false, new DateTime(2026, 1, 3), 22345678901234567d]);
        using var session = wb.Open();

        var cols = session.GetSchema("S", new SheetReadOptions()).Columns;
        Assert.Equal(["001", "002"], cols[0].SampleValues);
        Assert.Equal(
            [DataType.String, DataType.Decimal, DataType.Integer, DataType.Boolean, DataType.Date, DataType.String],
            cols.Select(c => c.SuggestedType));
    }

    [Fact]
    public void Column_letters_round_trip()
    {
        Assert.Equal("A", ColumnLetters.FromNumber(1));
        Assert.Equal("Z", ColumnLetters.FromNumber(26));
        Assert.Equal("AA", ColumnLetters.FromNumber(27));
        Assert.Equal("XFD", ColumnLetters.FromNumber(16384));
        Assert.Equal(16384, ColumnLetters.ToNumber("XFD"));
    }
}
