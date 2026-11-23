using System.Text.Json.Nodes;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Conversion;

public enum ConversionStatus
{
    Ok,
    Omit,
    Failed
}

/// <summary>Bir hücre/metin dönüşümünün sonucu. Value null ise JSON null yazılır.</summary>
public readonly record struct ConversionResult(ConversionStatus Status, JsonNode? Value, string? Error)
{
    public static ConversionResult Ok(JsonNode? value) => new(ConversionStatus.Ok, value, null);
    public static readonly ConversionResult Omitted = new(ConversionStatus.Omit, null, null);
    public static ConversionResult Fail(string error) => new(ConversionStatus.Failed, null, error);

    public bool IsOk => Status == ConversionStatus.Ok;
}

/// <summary>Excel hücre değerini hedef JSON tipine dönüştürür.</summary>
public interface IValueConverter
{
    /// <summary>Sütun alanı için hücre dönüşümü (trim, boş değer politikası, tip ayarları dahil).</summary>
    ConversionResult Convert(in CellValue cell, ColumnField field);

    /// <summary>Metin değeri hedef tipe dönüştürür (varsayılan değerler, sabitler, şablon sonuçları, dizi elemanları).</summary>
    ConversionResult ConvertText(string text, DataType type, TypeSettings settings);
}
