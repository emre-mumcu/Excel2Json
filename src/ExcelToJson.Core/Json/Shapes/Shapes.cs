using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelToJson.Core.Models;
using ExcelToJson.Core.Validation;

namespace ExcelToJson.Core.Json.Shapes;

/// <summary>[ {..}, {..} ] — stream edilir.</summary>
public sealed class ArrayOfObjectsShape : IOutputShapeStrategy
{
    public string Id => OutputShapes.ArrayOfObjects;
    public string DisplayName => "Nesne dizisi";
    public string Description => "Her satır bir nesne. En yaygın biçim.";
    public string Example => """[ { "code": "P001", "name": "Laptop" }, { ... } ]""";
    public int Order => 1;

    public IEnumerable<ValidationMessage> Validate(MappingProfile profile) => [];

    public void Write(Utf8JsonWriter writer, IEnumerable<Record> records, ShapeContext context)
    {
        writer.WriteStartArray();
        foreach (var r in records)
            r.ToJsonObject().WriteTo(writer);
        writer.WriteEndArray();
    }
}

/// <summary>{ "anahtar": {..} } — anahtar bir alandan veya Excel satır numarasından gelir.</summary>
public sealed class KeyedObjectShape : IOutputShapeStrategy
{
    public string Id => OutputShapes.KeyedObject;
    public string DisplayName => "Anahtarlı nesne";
    public string Description => "Seçilen alan (veya satır no) anahtar olur; her kayıt bir nesne.";
    public string Example => """{ "P001": { "name": "Laptop" }, "P002": { ... } }""";
    public int Order => 2;

    public IEnumerable<ValidationMessage> Validate(MappingProfile profile) =>
        ShapeHelpers.RequireField(profile, profile.Output.KeyFieldId, "Anahtar alanı", allowRowNumber: true);

    public void Write(Utf8JsonWriter writer, IEnumerable<Record> records, ShapeContext context)
    {
        var o = context.Options;
        var exclude = o.KeepKeyInObject ? null : o.KeyFieldId;
        ShapeHelpers.WriteKeyed(writer,
            records.Select(r => (ShapeHelpers.GetKey(r, o.KeyFieldId), r.RowNumber, (Func<JsonNode?>)(() => r.ToJsonObject(exclude)))),
            o.DuplicateKeyPolicy, context);
    }
}

/// <summary>Yalnızca tek kayıt: { .. }</summary>
public sealed class SingleObjectShape : IOutputShapeStrategy
{
    public string Id => OutputShapes.SingleObject;
    public string DisplayName => "Tek nesne";
    public string Description => "Yalnızca ilk kayıt (veya seçilen satır) tek bir nesne olarak.";
    public string Example => """{ "code": "P001", "name": "Laptop" }""";
    public int Order => 3;
    public bool RespectsPreviewLimit => false;

    public IEnumerable<ValidationMessage> Validate(MappingProfile profile)
    {
        if (profile.Output.SingleRowNumber is < 1)
            yield return ValidationMessage.Error(ValidationArea.Output, "Satır numarası 1 veya daha büyük olmalıdır.");
    }

    public void Write(Utf8JsonWriter writer, IEnumerable<Record> records, ShapeContext context)
    {
        var row = context.Options.SingleRowNumber;
        var record = row is null ? records.FirstOrDefault() : records.FirstOrDefault(r => r.RowNumber == row);
        if (record is null)
        {
            context.Warnings.Add(row is null ? "Hiç kayıt yok; null yazıldı." : $"{row}. satırda kayıt bulunamadı; null yazıldı.");
            writer.WriteNullValue();
            return;
        }
        record.ToJsonObject().WriteTo(writer);
    }
}

/// <summary>İki sütundan sözlük: { "anahtar": değer } — ayar/konfigürasyon sayfaları için.</summary>
public sealed class KeyValueShape : IOutputShapeStrategy
{
    public string Id => OutputShapes.KeyValue;
    public string DisplayName => "Anahtar–değer";
    public string Description => "İki sütundan sözlük. Ayar/konfigürasyon sayfaları için.";
    public string Example => """{ "timeout": 30, "theme": "dark" }""";
    public int Order => 4;

    public IEnumerable<ValidationMessage> Validate(MappingProfile profile)
    {
        foreach (var m in ShapeHelpers.RequireField(profile, profile.Output.KeyFieldId, "Anahtar alanı", allowRowNumber: true))
            yield return m;
        foreach (var m in ShapeHelpers.RequireField(profile, profile.Output.ValueFieldId, "Değer alanı"))
            yield return m;
    }

    public void Write(Utf8JsonWriter writer, IEnumerable<Record> records, ShapeContext context)
    {
        var o = context.Options;
        ShapeHelpers.WriteKeyed(writer,
            records.Select(r => (ShapeHelpers.GetKey(r, o.KeyFieldId), r.RowNumber, (Func<JsonNode?>)(() => ShapeHelpers.GetValue(r, o.ValueFieldId)))),
            o.DuplicateKeyPolicy, context);
    }
}

/// <summary>Tek sütun: [ "a", "b", "c" ] — stream edilir.</summary>
public sealed class ArrayOfValuesShape : IOutputShapeStrategy
{
    public string Id => OutputShapes.ArrayOfValues;
    public string DisplayName => "Değer dizisi";
    public string Description => "Tek bir alanın değerleri düz dizi olarak.";
    public string Example => """[ "Laptop", "Mouse", "Klavye" ]""";
    public int Order => 5;

    public IEnumerable<ValidationMessage> Validate(MappingProfile profile) =>
        ShapeHelpers.RequireField(profile, profile.Output.ValueFieldId, "Değer alanı");

    public void Write(Utf8JsonWriter writer, IEnumerable<Record> records, ShapeContext context)
    {
        var id = context.Options.ValueFieldId!;
        writer.WriteStartArray();
        foreach (var r in records)
        {
            if (r.TryGetValue(id, out var v) && !v.Omitted)
                ShapeHelpers.WriteNode(writer, v.Value);
        }
        writer.WriteEndArray();
    }
}

/// <summary>Satırlar değer dizisi: [ ["P001","Laptop"], .. ] — stream edilir.</summary>
public sealed class ArrayOfArraysShape : IOutputShapeStrategy
{
    public string Id => OutputShapes.ArrayOfArrays;
    public string DisplayName => "Dizi dizisi";
    public string Description => "Her satır, alan sırasıyla bir değer dizisi. İsteğe bağlı başlık satırı.";
    public string Example => """[ ["code","name"], ["P001","Laptop"], ... ]""";
    public int Order => 6;

    public IEnumerable<ValidationMessage> Validate(MappingProfile profile) => [];

    public void Write(Utf8JsonWriter writer, IEnumerable<Record> records, ShapeContext context)
    {
        writer.WriteStartArray();
        if (context.Options.IncludeHeaderRow)
        {
            writer.WriteStartArray();
            foreach (var f in context.Plan.Fields)
                writer.WriteStringValue(f.Alias);
            writer.WriteEndArray();
        }
        foreach (var r in records)
            r.ToJsonArray().WriteTo(writer);
        writer.WriteEndArray();
    }
}

/// <summary>
/// Seçilen alana göre gruplama. İki varyant:
/// { "grup": [ {..} ] } veya [ { "grupAlanı": "..", "items": [ .. ] } ]
/// </summary>
public sealed class GroupedShape : IOutputShapeStrategy
{
    public string Id => OutputShapes.Grouped;
    public string DisplayName => "Gruplanmış";
    public string Description => "Seçilen alana göre gruplanmış kayıtlar.";
    public string Example => """{ "Elektronik": [ { "code": "P001" }, ... ] }""";
    public int Order => 7;

    public IEnumerable<ValidationMessage> Validate(MappingProfile profile)
    {
        foreach (var m in ShapeHelpers.RequireField(profile, profile.Output.GroupFieldId, "Gruplama alanı"))
            yield return m;
        if (profile.Output.GroupedVariant == GroupedVariant.ArrayOfGroups && string.IsNullOrWhiteSpace(profile.Output.ChildrenName))
            yield return ValidationMessage.Error(ValidationArea.Output, "Alt kayıt dizisinin adı boş olamaz.");
    }

    public void Write(Utf8JsonWriter writer, IEnumerable<Record> records, ShapeContext context)
    {
        var o = context.Options;
        var groupId = o.GroupFieldId!;
        var exclude = o.KeepGroupFieldInItems ? null : groupId;
        var groups = new OrderedDictionary<string, (JsonNode? KeyValue, List<JsonObject> Items)>(StringComparer.Ordinal);

        foreach (var r in records)
        {
            var keyNode = r.TryGetValue(groupId, out var v) && !v.Omitted ? v.Value : null;
            var key = Record.ToKeyString(keyNode) ?? string.Empty;
            if (!groups.TryGetValue(key, out var group))
                groups[key] = group = (Record.Detach(keyNode)?.DeepClone(), []);
            group.Items.Add(r.ToJsonObject(exclude));
        }

        if (o.GroupedVariant == GroupedVariant.Dictionary)
        {
            writer.WriteStartObject();
            foreach (var (key, group) in groups)
            {
                writer.WritePropertyName(key);
                WriteItems(writer, group.Items);
            }
            writer.WriteEndObject();
        }
        else
        {
            var groupField = context.Plan.Find(groupId);
            var groupName = groupField?.Path[^1] ?? "group";
            var childrenName = string.IsNullOrWhiteSpace(o.ChildrenName) ? "items" : o.ChildrenName;
            writer.WriteStartArray();
            foreach (var (_, group) in groups)
            {
                writer.WriteStartObject();
                writer.WritePropertyName(groupName);
                ShapeHelpers.WriteNode(writer, group.KeyValue);
                writer.WritePropertyName(childrenName);
                WriteItems(writer, group.Items);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }
    }

    private static void WriteItems(Utf8JsonWriter writer, List<JsonObject> items)
    {
        writer.WriteStartArray();
        foreach (var item in items)
            item.WriteTo(writer);
        writer.WriteEndArray();
    }
}
