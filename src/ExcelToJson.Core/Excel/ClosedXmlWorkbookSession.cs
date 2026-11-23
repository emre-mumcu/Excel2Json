using System.Collections.Concurrent;
using System.Globalization;
using ClosedXML.Excel;
using ExcelToJson.Core.Conversion;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Excel;

/// <summary>
/// Çalışma kitabını bellekte tutar. Her sayfa ilk erişimde bir kez ClosedXML'den
/// değişmez (immutable) bir <see cref="SheetCache"/>'e aktarılır; sonraki tüm okumalar bu cache'ten yapılır.
/// </summary>
internal sealed class ClosedXmlWorkbookSession : IWorkbookSession
{
    private readonly XLWorkbook _workbook;
    private readonly Stream _buffer;
    private readonly object _workbookLock = new();
    private readonly ConcurrentDictionary<string, Lazy<SheetCache>> _sheets = new(StringComparer.OrdinalIgnoreCase);
    private bool _disposed;

    public ClosedXmlWorkbookSession(string path, long fileSize, XLWorkbook workbook, Stream buffer)
    {
        _workbook = workbook;
        _buffer = buffer;

        var sheets = new List<SheetInfo>();
        foreach (var ws in workbook.Worksheets)
        {
            var range = ws.RangeUsed();
            if (range is null)
            {
                sheets.Add(new SheetInfo(ws.Name, ws.Position, ws.Visibility != XLWorksheetVisibility.Visible, null, 0, 0, 0, 0));
                continue;
            }

            var first = range.FirstCell().Address;
            var last = range.LastCell().Address;
            sheets.Add(new SheetInfo(
                ws.Name, ws.Position, ws.Visibility != XLWorksheetVisibility.Visible,
                $"{first.ToStringRelative()}:{last.ToStringRelative()}",
                first.RowNumber, last.RowNumber, first.ColumnNumber, last.ColumnNumber));
        }

        Info = new WorkbookInfo(Path.GetFullPath(path), Path.GetFileName(path), fileSize, DateTime.Now, sheets);
    }

    public WorkbookInfo Info { get; }

    public SheetInfo GetSheet(string sheetName) =>
        Info.Sheets.FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
        ?? throw new ExcelReadException($"\"{sheetName}\" adlı sayfa bulunamadı.");

    private SheetCache GetCache(string sheetName)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var info = GetSheet(sheetName);
        return _sheets.GetOrAdd(info.Name, _ => new Lazy<SheetCache>(() =>
        {
            lock (_workbookLock)
                return SheetCache.Build(_workbook.Worksheet(info.Name), info);
        })).Value;
    }

    public RawPreview GetRawPreview(string sheetName, int maxRows = 20)
    {
        var cache = GetCache(sheetName);
        if (cache.IsEmpty)
            return new RawPreview([], []);

        var letters = Enumerable.Range(cache.FirstColumn, cache.ColumnCount).Select(ColumnLetters.FromNumber).ToList();
        var rows = new List<RawPreviewRow>();
        for (var r = cache.FirstRow; r <= cache.LastRow && rows.Count < maxRows; r++)
        {
            var values = new string[cache.ColumnCount];
            for (var c = 0; c < values.Length; c++)
                values[c] = cache.Get(r, cache.FirstColumn + c).Formatted ?? string.Empty;
            rows.Add(new RawPreviewRow(r, cache.HiddenRows.Contains(r), values));
        }

        return new RawPreview(letters, rows);
    }

    public SheetSchema GetSchema(string sheetName, SheetReadOptions options, int sampleRowLimit = 200)
    {
        var cache = GetCache(sheetName);
        if (cache.IsEmpty)
            return new SheetSchema(cache.Name, null, 1, 0, []);

        var (headerRow, dataStart, dataEnd) = ResolveRows(cache, options);
        var columns = new List<ColumnInfo>();

        foreach (var col in VisibleColumns(cache, options))
        {
            var letter = ColumnLetters.FromNumber(col);
            string header;
            if (headerRow is null)
            {
                header = letter;
            }
            else
            {
                header = cache.Get(headerRow.Value, col, options.SpreadMergedCells).Formatted?.Trim() ?? string.Empty;
                if (header.Length == 0)
                    header = $"Column {letter}";
            }

            // Örnek değerler ve tip tahmini için ilk N veri satırı taranır.
            var samples = new List<string>(3);
            var typedSamples = new List<CellValue>();
            var scanned = 0;
            for (var r = dataStart; r <= dataEnd && scanned < sampleRowLimit; r++)
            {
                if (!options.IncludeHiddenRows && cache.HiddenRows.Contains(r))
                    continue;
                scanned++;
                var cell = cache.Get(r, col, options.SpreadMergedCells);
                if (cell.IsBlank)
                    continue;
                typedSamples.Add(cell);
                if (samples.Count < 3)
                    samples.Add(cell.Formatted ?? string.Empty);
            }

            columns.Add(new ColumnInfo(col, letter, header, cache.HiddenColumns.Contains(col), samples,
                TypeInference.Infer(typedSamples)));
        }

        return new SheetSchema(cache.Name, headerRow, dataStart, dataEnd, columns);
    }

    public IEnumerable<SheetRow> ReadRows(string sheetName, SheetReadOptions options, RowReadStats? stats = null,
        CancellationToken cancellationToken = default)
    {
        var cache = GetCache(sheetName);
        if (cache.IsEmpty)
            yield break;

        var (_, dataStart, dataEnd) = ResolveRows(cache, options);
        var columns = VisibleColumns(cache, options).ToArray();

        for (var r = dataStart; r <= dataEnd; r++)
        {
            if ((r & 0x3FF) == 0)
                cancellationToken.ThrowIfCancellationRequested();

            if (!options.IncludeHiddenRows && cache.HiddenRows.Contains(r))
            {
                if (stats is not null) stats.SkippedHiddenRows++;
                continue;
            }

            var cells = cache.GetRow(r, options.SpreadMergedCells);

            if (options.SkipEmptyRows && IsEmptyRow(cells, columns, cache.FirstColumn))
            {
                if (stats is not null) stats.SkippedEmptyRows++;
                continue;
            }

            if (stats is not null)
            {
                foreach (var c in columns)
                {
                    var i = c - cache.FirstColumn;
                    if (i < cells.Length && cells[i].MissingFormulaValue)
                        stats.MissingFormulaValues++;
                }
            }

            yield return new SheetRow(r, cache.FirstColumn, cells);
        }
    }

    private static bool IsEmptyRow(CellValue[] cells, int[] columns, int firstColumn)
    {
        if (cells.Length == 0)
            return true;
        foreach (var c in columns)
        {
            var i = c - firstColumn;
            if (i < cells.Length && !cells[i].IsBlank && !string.IsNullOrWhiteSpace(cells[i].Formatted))
                return false;
        }
        return true;
    }

    private static IEnumerable<int> VisibleColumns(SheetCache cache, SheetReadOptions options)
    {
        for (var c = cache.FirstColumn; c <= cache.LastColumn; c++)
        {
            if (!options.IncludeHiddenColumns && cache.HiddenColumns.Contains(c))
                continue;
            yield return c;
        }
    }

    private static (int? HeaderRow, int DataStart, int DataEnd) ResolveRows(SheetCache cache, SheetReadOptions options)
    {
        int? headerRow = options.NoHeader ? null : Math.Max(1, options.HeaderRow ?? cache.FirstRow);
        var dataStart = options.DataStartRow ?? (headerRow is null ? cache.FirstRow : headerRow.Value + 1);
        dataStart = Math.Max(dataStart, 1);
        var dataEnd = options.DataEndRow is { } end ? Math.Min(end, cache.LastRow) : cache.LastRow;
        return (headerRow, dataStart, dataEnd);
    }

    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        lock (_workbookLock)
        {
            _workbook.Dispose();
            _buffer.Dispose();
        }
    }

    /// <summary>Bir sayfanın değişmez hücre cache'i.</summary>
    private sealed class SheetCache
    {
        private static readonly CellValue[] EmptyRow = [];
        private CellValue[]?[] _rows = [];
        private (int R1, int C1, int R2, int C2)[] _merged = [];

        public string Name { get; private init; } = string.Empty;
        public bool IsEmpty { get; private init; }
        public int FirstRow { get; private init; }
        public int LastRow { get; private init; }
        public int FirstColumn { get; private init; }
        public int LastColumn { get; private init; }
        public int ColumnCount => LastColumn - FirstColumn + 1;
        public HashSet<int> HiddenRows { get; } = [];
        public HashSet<int> HiddenColumns { get; } = [];

        public CellValue Get(int row, int column, bool spreadMerged = false)
        {
            if (spreadMerged && TryGetMergeAnchor(row, column, out var ar, out var ac))
                (row, column) = (ar, ac);
            return GetRaw(row, column);
        }

        private CellValue GetRaw(int row, int column)
        {
            if (row < FirstRow || row > LastRow || column < FirstColumn || column > LastColumn)
                return CellValue.Blank;
            var cells = _rows[row - FirstRow];
            return cells is null ? CellValue.Blank : cells[column - FirstColumn];
        }

        public CellValue[] GetRow(int row, bool spreadMerged)
        {
            var cells = row < FirstRow || row > LastRow ? null : _rows[row - FirstRow];
            if (!spreadMerged || _merged.Length == 0)
                return cells ?? EmptyRow;

            CellValue[]? copy = null;
            foreach (var m in _merged)
            {
                if (row < m.R1 || row > m.R2)
                    continue;
                var anchor = GetRaw(m.R1, m.C1);
                for (var c = Math.Max(m.C1, FirstColumn); c <= Math.Min(m.C2, LastColumn); c++)
                {
                    if (row == m.R1 && c == m.C1)
                        continue;
                    copy ??= cells is null ? NewBlankRow() : (CellValue[])cells.Clone();
                    copy[c - FirstColumn] = anchor;
                }
            }
            return copy ?? cells ?? EmptyRow;
        }

        private CellValue[] NewBlankRow()
        {
            var arr = new CellValue[ColumnCount];
            Array.Fill(arr, CellValue.Blank);
            return arr;
        }

        private bool TryGetMergeAnchor(int row, int column, out int anchorRow, out int anchorColumn)
        {
            foreach (var m in _merged)
            {
                if (row >= m.R1 && row <= m.R2 && column >= m.C1 && column <= m.C2)
                {
                    anchorRow = m.R1;
                    anchorColumn = m.C1;
                    return true;
                }
            }
            anchorRow = anchorColumn = 0;
            return false;
        }

        public static SheetCache Build(IXLWorksheet ws, SheetInfo info)
        {
            if (info.IsEmpty)
                return new SheetCache { Name = info.Name, IsEmpty = true };

            var cache = new SheetCache
            {
                Name = info.Name,
                FirstRow = info.FirstRow,
                LastRow = info.LastRow,
                FirstColumn = info.FirstColumn,
                LastColumn = info.LastColumn
            };

            var rows = new CellValue[]?[info.RowCount];
            var columnCount = info.ColumnCount;

            foreach (var cell in ws.CellsUsed(XLCellsUsedOptions.Contents))
            {
                var r = cell.Address.RowNumber - info.FirstRow;
                var c = cell.Address.ColumnNumber - info.FirstColumn;
                if (r < 0 || r >= rows.Length || c < 0 || c >= columnCount)
                    continue;
                if (rows[r] is null)
                {
                    rows[r] = new CellValue[columnCount];
                    Array.Fill(rows[r]!, CellValue.Blank);
                }
                rows[r]![c] = ReadCell(cell);
            }
            cache._rows = rows;

            // Not: ws.RowsUsed(XLCellsUsedOptions.All) kullanılmaz; tüm satırlara biçim uygulanmış sayfalarda
            // ~1 milyon satırı dolaşıp takılabiliyor. Yalnızca kullanılan aralıktaki satırlar kontrol edilir.
            for (var r = info.FirstRow; r <= info.LastRow; r++)
            {
                if (ws.Row(r).IsHidden)
                    cache.HiddenRows.Add(r);
            }
            for (var c = info.FirstColumn; c <= info.LastColumn; c++)
            {
                if (ws.Column(c).IsHidden)
                    cache.HiddenColumns.Add(c);
            }

            cache._merged = ws.MergedRanges
                .Select(m => (m.FirstCell().Address.RowNumber, m.FirstCell().Address.ColumnNumber,
                              m.LastCell().Address.RowNumber, m.LastCell().Address.ColumnNumber))
                .ToArray();

            return cache;
        }

        private static CellValue ReadCell(IXLCell cell)
        {
            var hasFormula = cell.HasFormula;
            // ClosedXML formülleri yeniden hesaplamaz; dosyadaki cache'lenmiş değer kullanılır.
            var value = hasFormula ? cell.CachedValue : cell.Value;

            string formatted;
            try
            {
                formatted = cell.GetFormattedString(CultureInfo.CurrentCulture);
            }
            catch
            {
                formatted = value.ToString(CultureInfo.CurrentCulture);
            }

            return value.Type switch
            {
                XLDataType.Blank => CellValue.Blank with { MissingFormulaValue = hasFormula, Formatted = string.Empty },
                XLDataType.Boolean => new CellValue { Kind = CellKind.Boolean, Boolean = value.GetBoolean(), Formatted = formatted },
                XLDataType.Number => new CellValue { Kind = CellKind.Number, Number = value.GetNumber(), Formatted = formatted },
                XLDataType.Text => new CellValue { Kind = CellKind.Text, Text = value.GetText(), Formatted = formatted },
                XLDataType.DateTime => new CellValue { Kind = CellKind.DateTime, DateTime = value.GetDateTime(), Formatted = formatted },
                XLDataType.TimeSpan => new CellValue { Kind = CellKind.TimeSpan, TimeSpan = value.GetTimeSpan(), Formatted = formatted },
                XLDataType.Error => new CellValue { Kind = CellKind.Error, Text = formatted, Formatted = formatted },
                _ => CellValue.Blank
            };
        }
    }
}
