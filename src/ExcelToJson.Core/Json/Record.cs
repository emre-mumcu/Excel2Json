using System.Text.Json.Nodes;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Json;

/// <summary>Bir alanın dönüştürülmüş değeri. Omitted ise alan JSON'a hiç yazılmaz.</summary>
public readonly record struct FieldValue(PlannedField Field, JsonNode? Value, bool Omitted);

/// <summary>Bir Excel satırından üretilmiş, alan sırasını koruyan kayıt.</summary>
public sealed class Record(int rowNumber, int sequence, FieldValue[] values)
{
    public int RowNumber { get; } = rowNumber;
    public int Sequence { get; } = sequence;
    public IReadOnlyList<FieldValue> Values { get; } = values;

    public bool TryGetValue(string fieldId, out FieldValue value)
    {
        foreach (var v in Values)
        {
            if (v.Field.Id == fieldId)
            {
                value = v;
                return true;
            }
        }
        value = default;
        return false;
    }

    /// <summary>Kaydı JSON nesnesine çevirir. Nokta notasyonlu alias'lar iç içe nesne üretir.</summary>
    public JsonObject ToJsonObject(string? excludeFieldId = null)
    {
        var root = new JsonObject();
        foreach (var v in Values)
        {
            if (v.Omitted || v.Field.Id == excludeFieldId)
                continue;

            var target = root;
            var path = v.Field.Path;
            for (var i = 0; i < path.Length - 1; i++)
            {
                if (target[path[i]] is JsonObject child)
                {
                    target = child;
                }
                else
                {
                    child = new JsonObject();
                    target[path[i]] = child;
                    target = child;
                }
            }
            target[path[^1]] = Detach(v.Value);
        }
        return root;
    }

    /// <summary>Kayıt değerlerini alan sırasıyla düz dizi olarak verir (Array of arrays).</summary>
    public JsonArray ToJsonArray()
    {
        var array = new JsonArray();
        foreach (var v in Values)
            array.Add(v.Omitted ? null : Detach(v.Value));
        return array;
    }

    /// <summary>JsonNode yalnızca tek bir parent'a bağlanabilir; gerekirse kopyalanır.</summary>
    internal static JsonNode? Detach(JsonNode? node) => node is null || node.Parent is null ? node : node.DeepClone();

    /// <summary>JSON değerini sözlük anahtarı olarak kullanılabilecek metne çevirir.</summary>
    public static string? ToKeyString(JsonNode? node)
    {
        if (node is null)
            return null;
        if (node is JsonValue value && value.TryGetValue<string>(out var s))
            return s;
        return node.ToJsonString();
    }
}

/// <summary>Satır, sütun, ham değer ve hedef tip ile bir dönüşüm hatası.</summary>
public sealed record ConversionIssue(
    int RowNumber,
    string FieldId,
    string FieldAlias,
    string? ColumnLetter,
    string? RawValue,
    DataType TargetType,
    string Message);

/// <summary>Dönüşüm hatalarını toplar; bellek için saklanan hata sayısı sınırlıdır.</summary>
public sealed class IssueCollector(int maxStored = 500)
{
    private readonly List<ConversionIssue> _issues = [];

    public IReadOnlyList<ConversionIssue> Issues => _issues;
    public int TotalCount { get; private set; }

    public void Add(ConversionIssue issue)
    {
        TotalCount++;
        if (_issues.Count < maxStored)
            _issues.Add(issue);
    }
}

/// <summary>Çıktı şekli yazılırken oluşan, kullanıcıya gösterilecek hata (ör. tekrarlanan anahtar).</summary>
public sealed class ShapeException(string message) : Exception(message);
