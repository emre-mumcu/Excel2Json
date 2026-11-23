using ExcelToJson.Core.Models;

namespace ExcelToJson.App.Components;

/// <summary>Enum değerlerinin arayüzde gösterilen Türkçe karşılıkları.</summary>
public static class Labels
{
    public static readonly IReadOnlyDictionary<DataType, string> DataType = new Dictionary<DataType, string>
    {
        [Core.Models.DataType.Auto] = "Auto",
        [Core.Models.DataType.String] = "Metin (String)",
        [Core.Models.DataType.Integer] = "Tam sayı (Integer)",
        [Core.Models.DataType.Decimal] = "Ondalık (Decimal)",
        [Core.Models.DataType.Boolean] = "Mantıksal (Boolean)",
        [Core.Models.DataType.Date] = "Tarih (Date)",
        [Core.Models.DataType.DateTime] = "Tarih-saat (DateTime)",
        [Core.Models.DataType.Array] = "Dizi (ayraçlı)"
    };

    public static readonly IReadOnlyDictionary<NamingStyle, string> NamingStyle = new Dictionary<NamingStyle, string>
    {
        [Core.Models.NamingStyle.CamelCase] = "camelCase",
        [Core.Models.NamingStyle.PascalCase] = "PascalCase",
        [Core.Models.NamingStyle.SnakeCase] = "snake_case",
        [Core.Models.NamingStyle.Original] = "Orijinal"
    };

    public static readonly IReadOnlyDictionary<EmptyValuePolicy, string> EmptyPolicy = new Dictionary<EmptyValuePolicy, string>
    {
        [EmptyValuePolicy.Null] = "null yaz",
        [EmptyValuePolicy.EmptyString] = "\"\" yaz",
        [EmptyValuePolicy.Omit] = "Alanı yazma",
        [EmptyValuePolicy.DefaultValue] = "Varsayılan değer"
    };

    public static readonly IReadOnlyDictionary<ConversionErrorPolicy, string> ErrorPolicy = new Dictionary<ConversionErrorPolicy, string>
    {
        [ConversionErrorPolicy.Stop] = "Export'u durdur",
        [ConversionErrorPolicy.SetNull] = "Hatalı değeri null yap",
        [ConversionErrorPolicy.KeepRaw] = "Ham metin olarak bırak"
    };

    public static readonly IReadOnlyDictionary<CustomFieldSource, string> Source = new Dictionary<CustomFieldSource, string>
    {
        [CustomFieldSource.Constant] = "Sabit değer",
        [CustomFieldSource.Template] = "Şablon",
        [CustomFieldSource.RowNumber] = "Satır numarası",
        [CustomFieldSource.AutoIncrement] = "Otomatik artan",
        [CustomFieldSource.Guid] = "GUID",
        [CustomFieldSource.Timestamp] = "Tarih/Saat (export anı)",
        [CustomFieldSource.Json] = "Sabit JSON"
    };

    public static readonly IReadOnlyDictionary<DuplicateKeyPolicy, string> DuplicateKey = new Dictionary<DuplicateKeyPolicy, string>
    {
        [DuplicateKeyPolicy.Error] = "Hata ver",
        [DuplicateKeyPolicy.LastWins] = "Sonuncu kazansın",
        [DuplicateKeyPolicy.FirstWins] = "İlki kazansın",
        [DuplicateKeyPolicy.ToArray] = "Değeri diziye çevir"
    };

    public static readonly IReadOnlyDictionary<ExistingFilePolicy, string> ExistingFile = new Dictionary<ExistingFilePolicy, string>
    {
        [ExistingFilePolicy.Overwrite] = "Üzerine yaz",
        [ExistingFilePolicy.AppendNumber] = "Numara ekle – ad (2).json",
        [ExistingFilePolicy.Ask] = "Bana sor"
    };

    public static readonly IReadOnlyDictionary<RowNumberMode, string> RowNumber = new Dictionary<RowNumberMode, string>
    {
        [RowNumberMode.Sequence] = "1'den başlayan sıra no",
        [RowNumberMode.ExcelRow] = "Excel satır numarası"
    };

    public static readonly IReadOnlyDictionary<TimestampMode, string> Timestamp = new Dictionary<TimestampMode, string>
    {
        [TimestampMode.Utc] = "UTC (DateTime.UtcNow)",
        [TimestampMode.Local] = "Yerel saat (DateTime.Now)"
    };

    public static readonly IReadOnlyDictionary<GroupedVariant, string> Grouped = new Dictionary<GroupedVariant, string>
    {
        [GroupedVariant.Dictionary] = "{ \"grup\": [ … ] }",
        [GroupedVariant.ArrayOfGroups] = "[ { \"grup\": …, \"items\": [ … ] } ]"
    };

    /// <summary>Dizi elemanları için anlamlı tipler.</summary>
    public static readonly DataType[] ArrayElementTypes =
        [Core.Models.DataType.String, Core.Models.DataType.Integer, Core.Models.DataType.Decimal, Core.Models.DataType.Boolean, Core.Models.DataType.Date];

    public static readonly (string Value, string Label)[] Cultures = [("tr-TR", "Türkçe (1.234,56)"), ("", "Invariant (1,234.56)")];
}
