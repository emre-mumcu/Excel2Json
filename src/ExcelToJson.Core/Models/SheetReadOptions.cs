namespace ExcelToJson.Core.Models;

/// <summary>Sayfanın nasıl okunacağını belirleyen başlık/satır ayarları.</summary>
public sealed class SheetReadOptions
{
    /// <summary>Başlık satırı (1 tabanlı). Null ise kullanılan alanın ilk satırı.</summary>
    public int? HeaderRow { get; set; }

    /// <summary>True ise başlık yoktur; sütun adları A, B, C… olur.</summary>
    public bool NoHeader { get; set; }

    /// <summary>Veri başlangıç satırı. Null ise başlık + 1 (başlık yoksa ilk satır).</summary>
    public int? DataStartRow { get; set; }

    /// <summary>Opsiyonel veri bitiş satırı.</summary>
    public int? DataEndRow { get; set; }

    public bool SkipEmptyRows { get; set; } = true;
    public bool IncludeHiddenRows { get; set; } = true;
    public bool IncludeHiddenColumns { get; set; } = true;
    public bool SpreadMergedCells { get; set; }

    public SheetReadOptions Clone() => (SheetReadOptions)MemberwiseClone();
}
