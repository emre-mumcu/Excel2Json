using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelToJson.Core.Conversion;
using ExcelToJson.Core.Expressions;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Json;

/// <summary>Export anına ait bağlam bilgisi.</summary>
public sealed record ExportContext(DateTime Now, string SourceFileName, string SheetName)
{
    public DateTime UtcNow => Now.ToUniversalTime();
}

/// <summary>Plana alınmış (derlenmiş) bir alan.</summary>
public sealed class PlannedField
{
    internal PlannedField(FieldDefinition definition)
    {
        Definition = definition;
        Path = definition.Alias.Split('.');
    }

    public FieldDefinition Definition { get; }
    public string Id => Definition.Id;
    public string Alias => Definition.Alias;
    public string[] Path { get; }

    internal TemplateExpression? Template { get; set; }
    internal JsonNode? Constant { get; set; }
    internal string? ConstantError { get; set; }
}

/// <summary>
/// Profili bir kez "derler" (şablonları parse eder, sabitleri dönüştürür) ve her satır için
/// <see cref="Record"/> üretir.
/// </summary>
public sealed class ConversionPlan
{
    private readonly IValueConverter _converter;
    private readonly ExportContext _context;
    private readonly Dictionary<string, ColumnField> _templateLookup;
    private readonly ConversionErrorPolicy _errorPolicy;

    private ConversionPlan(MappingProfile profile, IValueConverter converter, ExportContext context)
    {
        _converter = converter;
        _context = context;
        _errorPolicy = profile.ErrorPolicy;
        Fields = profile.IncludedFields.Select(f => new PlannedField(f)).ToList();
        _templateLookup = BuildTemplateLookup(profile.Fields);

        foreach (var f in Fields)
        {
            if (f.Definition is CustomField custom)
                Prepare(f, custom);
        }
    }

    public IReadOnlyList<PlannedField> Fields { get; }

    public PlannedField? Find(string? id) => id is null ? null : Fields.FirstOrDefault(f => f.Id == id);

    public static ConversionPlan Create(MappingProfile profile, IValueConverter converter, ExportContext context) =>
        new(profile, converter, context);

    /// <summary>Şablon placeholder'larının çözülebileceği adlar: sütun alias'ları ve Excel başlıkları.</summary>
    public static Dictionary<string, ColumnField> BuildTemplateLookup(IEnumerable<FieldDefinition> fields)
    {
        var lookup = new Dictionary<string, ColumnField>(StringComparer.OrdinalIgnoreCase);
        var columns = fields.OfType<ColumnField>().ToList();
        foreach (var c in columns.Where(c => !string.IsNullOrWhiteSpace(c.Alias)))
            lookup.TryAdd(c.Alias, c);
        foreach (var c in columns.Where(c => !string.IsNullOrWhiteSpace(c.Header)))
            lookup.TryAdd(c.Header, c);
        return lookup;
    }

    private void Prepare(PlannedField planned, CustomField custom)
    {
        switch (custom.Source)
        {
            case CustomFieldSource.Template:
                if (TemplateExpression.TryParse(custom.Value, out var t, out var error))
                    planned.Template = t;
                else
                    planned.ConstantError = error;
                break;
            case CustomFieldSource.Constant:
                var r = _converter.ConvertText(custom.Value ?? string.Empty, custom.Type, custom.TypeSettings);
                if (r.Status == ConversionStatus.Failed)
                    planned.ConstantError = r.Error;
                else
                    planned.Constant = r.Value;
                break;
            case CustomFieldSource.Json:
                try
                {
                    planned.Constant = JsonNode.Parse(custom.Value ?? "null");
                }
                catch (JsonException ex)
                {
                    planned.ConstantError = "Geçersiz JSON: " + ex.Message;
                }
                break;
        }
    }

    /// <summary>Bir Excel satırından kayıt üretir. <paramref name="sequence"/> 1'den başlayan kayıt sırasıdır.</summary>
    public Record BuildRecord(SheetRow row, int sequence, IssueCollector issues)
    {
        var values = new FieldValue[Fields.Count];
        for (var i = 0; i < Fields.Count; i++)
        {
            var f = Fields[i];
            values[i] = f.Definition switch
            {
                ColumnField column => ConvertColumn(f, column, row, issues),
                CustomField custom => EvaluateCustom(f, custom, row, sequence, issues),
                _ => new FieldValue(f, null, true)
            };
        }
        return new Record(row.RowNumber, sequence, values);
    }

    private FieldValue ConvertColumn(PlannedField f, ColumnField column, SheetRow row, IssueCollector issues)
    {
        var cell = row[column.ColumnIndex];
        var result = _converter.Convert(cell, column);
        return result.Status switch
        {
            ConversionStatus.Ok => new FieldValue(f, result.Value, false),
            ConversionStatus.Omit => new FieldValue(f, null, true),
            _ => HandleFailure(f, row.RowNumber, column.ColumnLetter, cell.Formatted, column.Type, result.Error, issues)
        };
    }

    private FieldValue HandleFailure(PlannedField f, int rowNumber, string? columnLetter, string? raw, DataType type,
        string? error, IssueCollector issues)
    {
        issues.Add(new ConversionIssue(rowNumber, f.Id, f.Alias, columnLetter, raw, type, error ?? "Dönüşüm hatası"));
        return _errorPolicy switch
        {
            ConversionErrorPolicy.KeepRaw => new FieldValue(f, raw is null ? null : JsonValue.Create(raw), false),
            _ => new FieldValue(f, null, false)
        };
    }

    private FieldValue EvaluateCustom(PlannedField f, CustomField custom, SheetRow row, int sequence, IssueCollector issues)
    {
        if (f.ConstantError is not null)
            return HandleFailure(f, row.RowNumber, null, custom.Value, custom.Type, f.ConstantError, issues);

        switch (custom.Source)
        {
            case CustomFieldSource.Constant:
            case CustomFieldSource.Json:
                return new FieldValue(f, f.Constant?.DeepClone(), false);

            case CustomFieldSource.Template:
                var text = f.Template!.Evaluate(name =>
                    _templateLookup.TryGetValue(name, out var col)
                        ? ValueConverter.ToDisplayString(row[col.ColumnIndex], col.Trim)
                        : string.Empty);
                if (custom.Type is DataType.Auto or DataType.String)
                    return new FieldValue(f, JsonValue.Create(text), false);
                var r = _converter.ConvertText(text, custom.Type, custom.TypeSettings);
                return r.Status == ConversionStatus.Failed
                    ? HandleFailure(f, row.RowNumber, null, text, custom.Type, r.Error, issues)
                    : new FieldValue(f, r.Value, false);

            case CustomFieldSource.RowNumber:
                return new FieldValue(f, JsonValue.Create(custom.RowNumberMode == RowNumberMode.ExcelRow ? row.RowNumber : sequence), false);

            case CustomFieldSource.AutoIncrement:
                return new FieldValue(f, JsonValue.Create(custom.Start + (sequence - 1L) * custom.Step), false);

            case CustomFieldSource.Guid:
                return new FieldValue(f, JsonValue.Create(Guid.NewGuid().ToString()), false);

            case CustomFieldSource.Timestamp:
                return new FieldValue(f, JsonValue.Create(FormatTimestamp(custom, _context)), false);

            default:
                return new FieldValue(f, null, true);
        }
    }

    public static string FormatTimestamp(CustomField custom, ExportContext context)
    {
        var value = custom.TimestampMode == TimestampMode.Utc ? context.UtcNow : context.Now;
        if (!string.IsNullOrWhiteSpace(custom.Format))
            return value.ToString(custom.Format, CultureInfo.InvariantCulture);
        return custom.TimestampMode == TimestampMode.Utc
            ? value.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture)
            : new DateTimeOffset(value).ToString("yyyy-MM-dd'T'HH:mm:sszzz", CultureInfo.InvariantCulture);
    }

    /// <summary>Kök seviye ek alanın değeri (Template için {file}, {sheet}, {count} kullanılabilir).</summary>
    public JsonNode? EvaluateRootField(CustomField field, int count, IssueCollector issues)
    {
        switch (field.Source)
        {
            case CustomFieldSource.Constant:
            {
                var r = _converter.ConvertText(field.Value ?? string.Empty, field.Type, field.TypeSettings);
                if (r.Status == ConversionStatus.Failed)
                    issues.Add(new ConversionIssue(0, field.Id, field.Alias, null, field.Value, field.Type, r.Error ?? "Dönüşüm hatası"));
                return r.Value;
            }
            case CustomFieldSource.Json:
                try { return JsonNode.Parse(field.Value ?? "null"); }
                catch (JsonException ex)
                {
                    issues.Add(new ConversionIssue(0, field.Id, field.Alias, null, field.Value, DataType.String, "Geçersiz JSON: " + ex.Message));
                    return null;
                }
            case CustomFieldSource.Guid:
                return JsonValue.Create(Guid.NewGuid().ToString());
            case CustomFieldSource.Timestamp:
                return JsonValue.Create(FormatTimestamp(field, _context));
            case CustomFieldSource.Template:
                if (!TemplateExpression.TryParse(field.Value, out var t, out _))
                    return null;
                return JsonValue.Create(t.Evaluate(name => RootPlaceholder(name, count)));
            default:
                return null;
        }
    }

    public static readonly IReadOnlyList<string> RootPlaceholders = ["file", "sheet", "count"];

    private string? RootPlaceholder(string name, int count) => name.ToLowerInvariant() switch
    {
        "file" => _context.SourceFileName,
        "sheet" => _context.SheetName,
        "count" => count.ToString(CultureInfo.InvariantCulture),
        _ => string.Empty
    };
}
