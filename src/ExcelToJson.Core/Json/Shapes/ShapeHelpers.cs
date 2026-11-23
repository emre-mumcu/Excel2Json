using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelToJson.Core.Models;
using ExcelToJson.Core.Validation;

namespace ExcelToJson.Core.Json.Shapes;

internal static class ShapeHelpers
{
    /// <summary>Şekil için seçilmesi zorunlu bir alanı doğrular.</summary>
    public static IEnumerable<ValidationMessage> RequireField(MappingProfile profile, string? fieldId, string label, bool allowRowNumber = false)
    {
        if (allowRowNumber && fieldId == OutputShapes.RowNumberKey)
            yield break;
        if (string.IsNullOrEmpty(fieldId))
        {
            yield return ValidationMessage.Error(ValidationArea.Output, $"{label} seçilmelidir.");
            yield break;
        }
        var field = profile.FindField(fieldId);
        if (field is null)
            yield return ValidationMessage.Error(ValidationArea.Output, $"Seçilen {label.ToLowerInvariant()} artık mevcut değil.");
        else if (!field.Include)
            yield return ValidationMessage.Error(ValidationArea.Output, $"{label} (\"{field.Alias}\") dahil edilmiş alanlardan biri olmalıdır.", field.Id);
    }

    public static string? GetKey(Record record, string? keyFieldId)
    {
        if (keyFieldId == OutputShapes.RowNumberKey)
            return record.RowNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return keyFieldId is not null && record.TryGetValue(keyFieldId, out var v) && !v.Omitted
            ? Record.ToKeyString(v.Value)
            : null;
    }

    public static JsonNode? GetValue(Record record, string? fieldId) =>
        fieldId is not null && record.TryGetValue(fieldId, out var v) && !v.Omitted ? Record.Detach(v.Value) : null;

    public static void WriteNode(Utf8JsonWriter writer, JsonNode? node)
    {
        if (node is null)
            writer.WriteNullValue();
        else
            node.WriteTo(writer);
    }

    /// <summary>
    /// Anahtarlı sözlük yazar ve tekrarlanan anahtar politikasını uygular.
    /// Error ve FirstWins stream edilir; LastWins ve ToArray tüm değerleri bellekte toplar.
    /// </summary>
    public static void WriteKeyed(Utf8JsonWriter writer, IEnumerable<(string? Key, int Row, Func<JsonNode?> Value)> items,
        DuplicateKeyPolicy policy, ShapeContext context)
    {
        var skipped = 0;
        writer.WriteStartObject();

        switch (policy)
        {
            case DuplicateKeyPolicy.Error:
            case DuplicateKeyPolicy.FirstWins:
            {
                var seen = new Dictionary<string, int>(StringComparer.Ordinal);
                foreach (var (key, row, value) in items)
                {
                    if (string.IsNullOrEmpty(key)) { skipped++; continue; }
                    if (seen.TryGetValue(key, out var firstRow))
                    {
                        if (policy == DuplicateKeyPolicy.Error)
                            throw new ShapeException($"Tekrarlanan anahtar: \"{key}\" (satır {firstRow} ve {row}). Tekrarlanan anahtar politikasını değiştirebilirsiniz.");
                        continue;
                    }
                    seen[key] = row;
                    writer.WritePropertyName(key);
                    WriteNode(writer, value());
                }
                break;
            }
            case DuplicateKeyPolicy.LastWins:
            {
                var map = new OrderedDictionary<string, JsonNode?>(StringComparer.Ordinal);
                foreach (var (key, _, value) in items)
                {
                    if (string.IsNullOrEmpty(key)) { skipped++; continue; }
                    map[key] = value();
                }
                foreach (var (key, value) in map)
                {
                    writer.WritePropertyName(key);
                    WriteNode(writer, value);
                }
                break;
            }
            case DuplicateKeyPolicy.ToArray:
            {
                var map = new OrderedDictionary<string, List<JsonNode?>>(StringComparer.Ordinal);
                foreach (var (key, _, value) in items)
                {
                    if (string.IsNullOrEmpty(key)) { skipped++; continue; }
                    if (!map.TryGetValue(key, out var list))
                        map[key] = list = [];
                    list.Add(value());
                }
                foreach (var (key, list) in map)
                {
                    writer.WritePropertyName(key);
                    if (list.Count == 1)
                    {
                        WriteNode(writer, list[0]);
                    }
                    else
                    {
                        writer.WriteStartArray();
                        foreach (var n in list)
                            WriteNode(writer, n);
                        writer.WriteEndArray();
                    }
                }
                break;
            }
        }

        writer.WriteEndObject();
        if (skipped > 0)
            context.Warnings.Add($"Anahtarı boş olan {skipped} kayıt atlandı.");
    }
}
