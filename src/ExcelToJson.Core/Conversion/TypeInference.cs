using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Conversion;

/// <summary>Örnek hücre değerlerinden veri tipi tahmini ("Auto" modunun önerisi).</summary>
public static class TypeInference
{
    /// <summary>15 haneden uzun sayılar double hassasiyetini aşar; string önerilir (TCKN, barkod vb.).</summary>
    public const double MaxSafeInteger = 999_999_999_999_999d;

    public static DataType Infer(IReadOnlyCollection<CellValue> samples)
    {
        if (samples.Count == 0)
            return DataType.String;

        bool allNumber = true, allWhole = true, allBool = true, allDate = true, anyTime = false;
        foreach (var s in samples)
        {
            switch (s.Kind)
            {
                case CellKind.Number:
                    allBool = allDate = false;
                    if (Math.Abs(s.Number) > MaxSafeInteger)
                        return DataType.String;
                    if (s.Number != Math.Floor(s.Number))
                        allWhole = false;
                    break;
                case CellKind.Boolean:
                    allNumber = allDate = false;
                    break;
                case CellKind.DateTime:
                    allNumber = allBool = false;
                    if (s.DateTime.TimeOfDay != TimeSpan.Zero)
                        anyTime = true;
                    break;
                default:
                    // Metin (baştaki sıfırlar, kodlar vb.) her zaman string kalır.
                    return DataType.String;
            }
        }

        if (allNumber)
            return allWhole ? DataType.Integer : DataType.Decimal;
        if (allBool)
            return DataType.Boolean;
        if (allDate)
            return anyTime ? DataType.DateTime : DataType.Date;
        return DataType.String;
    }

    /// <summary>Sabit değer metninden tip tahmini: true/false, tam sayı, ondalık (invariant), aksi halde string.</summary>
    public static DataType InferConstant(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return DataType.String;
        var t = text.Trim();
        if (t is "true" or "false")
            return DataType.Boolean;
        if (t.Length > 1 && t[0] == '0' && char.IsDigit(t[1]))
            return DataType.String; // baştaki sıfır korunur
        if (long.TryParse(t, System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out _))
            return DataType.Integer;
        if (decimal.TryParse(t, System.Globalization.NumberStyles.AllowLeadingSign | System.Globalization.NumberStyles.AllowDecimalPoint,
                System.Globalization.CultureInfo.InvariantCulture, out _))
            return DataType.Decimal;
        return DataType.String;
    }
}
