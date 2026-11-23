namespace ExcelToJson.Core.Models;

/// <summary>Sayfanın, okuma ayarlarına göre çözümlenmiş sütun yapısı.</summary>
public sealed record SheetSchema(
    string SheetName,
    int? HeaderRow,
    int DataStartRow,
    int DataEndRow,
    IReadOnlyList<ColumnInfo> Columns)
{
    public ColumnInfo? FindByIndex(int columnIndex) => Columns.FirstOrDefault(c => c.Index == columnIndex);
}

/// <summary>Excel sütun bilgisi.</summary>
public sealed record ColumnInfo(
    int Index,
    string Letter,
    string Header,
    bool IsHidden,
    IReadOnlyList<string> SampleValues,
    DataType SuggestedType);

/// <summary>Ham veri önizlemesi (ilk N satır).</summary>
public sealed record RawPreview(
    IReadOnlyList<string> ColumnLetters,
    IReadOnlyList<RawPreviewRow> Rows);

public sealed record RawPreviewRow(int RowNumber, bool IsHidden, IReadOnlyList<string> Values);

/// <summary>Okunan bir veri satırı. Cells, şemadaki sütun sırasına göre değil sayfa sütun numarasına göre erişilir.</summary>
public sealed class SheetRow
{
    private readonly CellValue[] _cells;
    private readonly int _firstColumn;

    public SheetRow(int rowNumber, int firstColumn, CellValue[] cells)
    {
        RowNumber = rowNumber;
        _firstColumn = firstColumn;
        _cells = cells;
    }

    public int RowNumber { get; }

    /// <summary>Sütun numarasına (1 tabanlı) göre hücre değeri.</summary>
    public CellValue this[int columnIndex]
    {
        get
        {
            var i = columnIndex - _firstColumn;
            return i >= 0 && i < _cells.Length ? _cells[i] : CellValue.Blank;
        }
    }
}

/// <summary>Satır okuma sırasında toplanan istatistikler.</summary>
public sealed class RowReadStats
{
    public int SkippedEmptyRows { get; set; }
    public int SkippedHiddenRows { get; set; }
    public int MissingFormulaValues { get; set; }
}
