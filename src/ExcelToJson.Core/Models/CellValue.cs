namespace ExcelToJson.Core.Models;

public enum CellKind : byte
{
    Blank,
    Number,
    Text,
    Boolean,
    DateTime,
    TimeSpan,
    Error
}

/// <summary>
/// Excel hücresinin ham (tipli) değeri ve Excel'deki biçimlendirilmiş metni.
/// ClosedXML'den bağımsızdır; böylece dönüşüm katmanı kolayca test edilebilir.
/// </summary>
public readonly record struct CellValue
{
    public CellKind Kind { get; init; }
    public double Number { get; init; }
    public bool Boolean { get; init; }
    public DateTime DateTime { get; init; }
    public TimeSpan TimeSpan { get; init; }
    /// <summary>Metin veya hata (#N/A) değeri.</summary>
    public string? Text { get; init; }
    /// <summary>Excel'de görünen biçimlendirilmiş metin.</summary>
    public string Formatted { get; init; }
    /// <summary>Formül var ama hesaplanmış (cache) değer yok.</summary>
    public bool MissingFormulaValue { get; init; }

    public static readonly CellValue Blank = new() { Kind = CellKind.Blank, Formatted = string.Empty };

    public bool IsBlank => Kind == CellKind.Blank || (Kind == CellKind.Text && string.IsNullOrEmpty(Text));

    public static CellValue FromNumber(double value, string? formatted = null) =>
        new() { Kind = CellKind.Number, Number = value, Formatted = formatted ?? value.ToString(System.Globalization.CultureInfo.InvariantCulture) };

    public static CellValue FromText(string value) =>
        new() { Kind = CellKind.Text, Text = value, Formatted = value };

    public static CellValue FromBoolean(bool value) =>
        new() { Kind = CellKind.Boolean, Boolean = value, Formatted = value ? "TRUE" : "FALSE" };

    public static CellValue FromDateTime(DateTime value, string? formatted = null) =>
        new() { Kind = CellKind.DateTime, DateTime = value, Formatted = formatted ?? value.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture) };

    public override string ToString() => Formatted;
}
