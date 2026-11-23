namespace ExcelToJson.Core.Models;

/// <summary>Açılmış çalışma kitabının özet bilgisi.</summary>
public sealed record WorkbookInfo(
    string FilePath,
    string FileName,
    long FileSize,
    DateTime LoadedAt,
    IReadOnlyList<SheetInfo> Sheets);

/// <summary>Bir çalışma sayfasının özet bilgisi.</summary>
public sealed record SheetInfo(
    string Name,
    int Position,
    bool IsHidden,
    string? UsedRange,
    int FirstRow,
    int LastRow,
    int FirstColumn,
    int LastColumn)
{
    public bool IsEmpty => UsedRange is null;
    public int RowCount => IsEmpty ? 0 : LastRow - FirstRow + 1;
    public int ColumnCount => IsEmpty ? 0 : LastColumn - FirstColumn + 1;
}
