namespace ExcelToJson.Core.Models;

/// <summary>Çıktı şekli kimlikleri. Yeni bir şekil eklemek için yeni bir IOutputShapeStrategy yazmak yeterlidir.</summary>
public static class OutputShapes
{
    public const string ArrayOfObjects = "arrayOfObjects";
    public const string KeyedObject = "keyedObject";
    public const string SingleObject = "singleObject";
    public const string KeyValue = "keyValue";
    public const string ArrayOfValues = "arrayOfValues";
    public const string ArrayOfArrays = "arrayOfArrays";
    public const string Grouped = "grouped";

    /// <summary>Alan yerine Excel satır numarasını anahtar olarak kullanmak için özel kimlik.</summary>
    public const string RowNumberKey = "__row";
}

/// <summary>Çıktı yapısı ve şekle özgü ayarlar.</summary>
public sealed class OutputOptions
{
    public string Shape { get; set; } = OutputShapes.ArrayOfObjects;

    /// <summary>Anahtar alanının Id'si veya <see cref="OutputShapes.RowNumberKey"/>.</summary>
    public string? KeyFieldId { get; set; }
    public DuplicateKeyPolicy DuplicateKeyPolicy { get; set; } = DuplicateKeyPolicy.Error;
    public bool KeepKeyInObject { get; set; } = true;

    /// <summary>Single object: Excel satır numarası. Null ise ilk kayıt.</summary>
    public int? SingleRowNumber { get; set; }

    /// <summary>Key–value ve Array of values için değer alanı Id'si.</summary>
    public string? ValueFieldId { get; set; }

    /// <summary>Array of arrays: ilk eleman başlık dizisi olsun.</summary>
    public bool IncludeHeaderRow { get; set; }

    public string? GroupFieldId { get; set; }
    public GroupedVariant GroupedVariant { get; set; } = GroupedVariant.Dictionary;
    public string ChildrenName { get; set; } = "items";
    public bool KeepGroupFieldInItems { get; set; }

    public WrapperOptions Wrapper { get; set; } = new();

    public OutputOptions Clone()
    {
        var c = (OutputOptions)MemberwiseClone();
        c.Wrapper = Wrapper.Clone();
        return c;
    }
}

/// <summary>Kök sarmalayıcı: { "meta": {..}, "data": [..] }</summary>
public sealed class WrapperOptions
{
    public bool Enabled { get; set; }
    public string DataPropertyName { get; set; } = "data";
    public bool IncludeMeta { get; set; } = true;
    public string MetaPropertyName { get; set; } = "meta";

    /// <summary>Kök seviye ek alanlar. Meta açıksa meta bloğuna, değilse köke yazılır.</summary>
    public List<CustomField> RootFields { get; set; } = [];

    public WrapperOptions Clone()
    {
        var c = (WrapperOptions)MemberwiseClone();
        c.RootFields = RootFields.Select(f => (CustomField)f.Clone()).ToList();
        return c;
    }
}

/// <summary>JSON yazım ve dosya ayarları.</summary>
public sealed class ExportFormatOptions
{
    public bool Indented { get; set; } = true;
    public int IndentSize { get; set; } = 2;
    public bool UnescapedUnicode { get; set; } = true;
    public bool WriteBom { get; set; }
    public string FileNameTemplate { get; set; } = "{sheet}_{yyyyMMdd_HHmmss}.json";
    public ExistingFilePolicy ExistingFilePolicy { get; set; } = ExistingFilePolicy.AppendNumber;

    public ExportFormatOptions Clone() => (ExportFormatOptions)MemberwiseClone();
}
