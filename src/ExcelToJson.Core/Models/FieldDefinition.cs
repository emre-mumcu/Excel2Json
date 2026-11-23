using System.Text.Json.Serialization;

namespace ExcelToJson.Core.Models;

/// <summary>
/// JSON kaydındaki bir alanın tanımı. Excel sütunları ve ek alanlar aynı sıralı listede tutulur;
/// JSON alan sırası bu listenin sırasını takip eder.
/// </summary>
[JsonPolymorphic(TypeDiscriminatorPropertyName = "kind")]
[JsonDerivedType(typeof(ColumnField), "column")]
[JsonDerivedType(typeof(CustomField), "custom")]
public abstract class FieldDefinition
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    public bool Include { get; set; } = true;

    /// <summary>JSON alan adı. Nokta notasyonu iç içe nesne üretir (address.city).</summary>
    public string Alias { get; set; } = string.Empty;

    public DataType Type { get; set; } = DataType.Auto;

    public TypeSettings TypeSettings { get; set; } = new();

    [JsonIgnore]
    public abstract string DisplaySource { get; }

    public abstract FieldDefinition Clone();
}

/// <summary>Excel sütununa bağlı alan.</summary>
public sealed class ColumnField : FieldDefinition
{
    /// <summary>Sütun numarası (1 tabanlı). Profil yüklenirken başlığa göre yeniden çözülür.</summary>
    public int ColumnIndex { get; set; }
    public string ColumnLetter { get; set; } = string.Empty;
    public string Header { get; set; } = string.Empty;

    public EmptyValuePolicy EmptyPolicy { get; set; } = EmptyValuePolicy.Null;
    public string? DefaultValue { get; set; }
    public bool Trim { get; set; } = true;

    public override string DisplaySource => $"{ColumnLetter} · {Header}";

    public override FieldDefinition Clone()
    {
        var c = (ColumnField)MemberwiseClone();
        c.TypeSettings = TypeSettings.Clone();
        return c;
    }
}

/// <summary>Excel'de olmayan, JSON'a eklenen alan.</summary>
public sealed class CustomField : FieldDefinition
{
    public CustomFieldSource Source { get; set; } = CustomFieldSource.Constant;

    /// <summary>Sabit değer, şablon metni veya sabit JSON parçası.</summary>
    public string? Value { get; set; }

    public RowNumberMode RowNumberMode { get; set; } = RowNumberMode.Sequence;

    public long Start { get; set; } = 1;
    public long Step { get; set; } = 1;

    public TimestampMode TimestampMode { get; set; } = TimestampMode.Utc;

    /// <summary>Tarih/saat çıktı formatı. Boşsa ISO 8601.</summary>
    public string? Format { get; set; }

    public override string DisplaySource => Source switch
    {
        CustomFieldSource.Constant => "Sabit değer",
        CustomFieldSource.Template => "Şablon",
        CustomFieldSource.RowNumber => "Satır numarası",
        CustomFieldSource.AutoIncrement => "Otomatik artan",
        CustomFieldSource.Guid => "GUID",
        CustomFieldSource.Timestamp => "Tarih/Saat",
        CustomFieldSource.Json => "Sabit JSON",
        _ => Source.ToString()
    };

    public override FieldDefinition Clone()
    {
        var c = (CustomField)MemberwiseClone();
        c.TypeSettings = TypeSettings.Clone();
        return c;
    }
}

/// <summary>Tipe özgü dönüşüm ayarları.</summary>
public sealed class TypeSettings
{
    public const string DefaultTrueValues = "true, evet, e, yes, y, 1, x, var, doğru";
    public const string DefaultFalseValues = "false, hayır, h, no, n, 0, yok, yanlış";

    /// <summary>Tarih/saat çıktı formatı. Boşsa ISO 8601 (Date: yyyy-MM-dd, DateTime: yyyy-MM-ddTHH:mm:ss).</summary>
    public string? DateFormat { get; set; }

    /// <summary>Metin olarak girilmiş tarihler için opsiyonel giriş formatı (ör. dd.MM.yyyy).</summary>
    public string? DateInputFormat { get; set; }

    /// <summary>Metin sayı/tarih parse kültürü: "tr-TR" veya "" (Invariant).</summary>
    public string Culture { get; set; } = "tr-TR";

    public string TrueValues { get; set; } = DefaultTrueValues;
    public string FalseValues { get; set; } = DefaultFalseValues;

    /// <summary>Boolean: boş hücre false sayılır ("x / boş" kullanımı).</summary>
    public bool EmptyIsFalse { get; set; }

    public string ArraySeparator { get; set; } = ",";
    public DataType ArrayElementType { get; set; } = DataType.String;

    public TypeSettings Clone() => (TypeSettings)MemberwiseClone();
}
