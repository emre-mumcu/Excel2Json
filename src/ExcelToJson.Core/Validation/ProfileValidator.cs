using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelToJson.Core.Conversion;
using ExcelToJson.Core.Expressions;
using ExcelToJson.Core.Json;
using ExcelToJson.Core.Json.Shapes;
using ExcelToJson.Core.Mapping;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Validation;

/// <summary>Profilin (mapping + çıktı + export ayarları) export öncesi doğrulaması.</summary>
public sealed class ProfileValidator(OutputShapeRegistry shapes, IValueConverter converter)
{
    public static ProfileValidator CreateDefault() => new(OutputShapeRegistry.CreateDefault(), new ValueConverter());

    public ValidationResult Validate(MappingProfile profile, SheetSchema? schema = null)
    {
        var messages = new List<ValidationMessage>();
        ValidateSheet(profile, messages);
        ValidateFields(profile, schema, messages);
        ValidateOutput(profile, messages);
        ValidateFormat(profile, messages);
        return new ValidationResult(messages);
    }

    private static void ValidateSheet(MappingProfile profile, List<ValidationMessage> messages)
    {
        if (string.IsNullOrWhiteSpace(profile.SheetName))
        {
            messages.Add(ValidationMessage.Error(ValidationArea.Sheet, "Bir sayfa seçilmelidir."));
            return;
        }
        var o = profile.ReadOptions;
        if (o.HeaderRow is < 1)
            messages.Add(ValidationMessage.Error(ValidationArea.Sheet, "Başlık satırı 1 veya daha büyük olmalıdır."));
        if (o.DataStartRow is < 1)
            messages.Add(ValidationMessage.Error(ValidationArea.Sheet, "Veri başlangıç satırı 1 veya daha büyük olmalıdır."));
        if (o.DataStartRow is { } start && o.DataEndRow is { } end && end < start)
            messages.Add(ValidationMessage.Error(ValidationArea.Sheet, "Veri bitiş satırı, başlangıç satırından küçük olamaz."));
        if (!o.NoHeader && o.HeaderRow is { } h && o.DataStartRow is { } s && s <= h)
            messages.Add(ValidationMessage.Warning(ValidationArea.Sheet, "Veri başlangıç satırı başlık satırından sonra olmalıdır."));
    }

    private void ValidateFields(MappingProfile profile, SheetSchema? schema, List<ValidationMessage> messages)
    {
        var included = profile.IncludedFields.ToList();
        if (included.Count == 0)
        {
            messages.Add(ValidationMessage.Error(ValidationArea.Fields, "En az bir alan seçilmelidir."));
            return;
        }

        var validAliases = new List<FieldDefinition>();
        foreach (var f in included)
        {
            if (!AliasGenerator.IsValid(f.Alias, out var error))
                messages.Add(ValidationMessage.Error(ValidationArea.Fields, $"{Describe(f)}: {error}", f.Id));
            else
                validAliases.Add(f);
        }

        ValidatePaths(validAliases.Select(f => (f.Alias, (string?)f.Id)), ValidationArea.Fields, messages);

        var templateLookup = ConversionPlan.BuildTemplateLookup(profile.Fields);
        foreach (var f in included)
        {
            switch (f)
            {
                case ColumnField c:
                    if (schema is not null && schema.FindByIndex(c.ColumnIndex) is null)
                        messages.Add(ValidationMessage.Warning(ValidationArea.Fields, $"\"{c.Header}\" sütunu ({c.ColumnLetter}) sayfada bulunamadı; değerler boş gelecek.", c.Id));
                    if (c.EmptyPolicy == EmptyValuePolicy.DefaultValue && c.Type != DataType.Auto)
                    {
                        var r = converter.ConvertText(c.DefaultValue ?? string.Empty, c.Type, c.TypeSettings);
                        if (r.Status == ConversionStatus.Failed)
                            messages.Add(ValidationMessage.Error(ValidationArea.Fields, $"\"{c.Alias}\" varsayılan değeri geçersiz: {r.Error}", c.Id));
                    }
                    ValidateTypeSettings(f, messages);
                    break;
                case CustomField custom:
                    ValidateCustomField(custom, templateLookup, messages, isRoot: false);
                    break;
            }
        }
    }

    private static void ValidateTypeSettings(FieldDefinition f, List<ValidationMessage> messages)
    {
        if (f.Type == DataType.Array && string.IsNullOrEmpty(f.TypeSettings.ArraySeparator))
            messages.Add(ValidationMessage.Warning(ValidationArea.Fields, $"\"{f.Alias}\": dizi ayracı boş; ',' kullanılacak.", f.Id));
        if (f.Type is DataType.Date or DataType.DateTime && !string.IsNullOrWhiteSpace(f.TypeSettings.DateFormat))
        {
            try
            {
                _ = DateTime.Now.ToString(f.TypeSettings.DateFormat, System.Globalization.CultureInfo.InvariantCulture);
            }
            catch (FormatException)
            {
                messages.Add(ValidationMessage.Error(ValidationArea.Fields, $"\"{f.Alias}\": geçersiz tarih formatı.", f.Id));
            }
        }
    }

    private void ValidateCustomField(CustomField f, Dictionary<string, ColumnField> templateLookup,
        List<ValidationMessage> messages, bool isRoot)
    {
        var area = isRoot ? ValidationArea.Output : ValidationArea.Fields;
        switch (f.Source)
        {
            case CustomFieldSource.Template:
                if (string.IsNullOrEmpty(f.Value))
                {
                    messages.Add(ValidationMessage.Warning(area, $"\"{f.Alias}\": şablon boş.", f.Id));
                    break;
                }
                if (!TemplateExpression.TryParse(f.Value, out var t, out var error))
                {
                    messages.Add(ValidationMessage.Error(area, $"\"{f.Alias}\": şablon hatası – {error}", f.Id));
                    break;
                }
                foreach (var p in t.Placeholders)
                {
                    var known = isRoot
                        ? ConversionPlan.RootPlaceholders.Contains(p, StringComparer.OrdinalIgnoreCase)
                        : templateLookup.ContainsKey(p);
                    if (!known)
                        messages.Add(ValidationMessage.Error(area,
                            isRoot
                                ? $"\"{f.Alias}\": kök alan şablonlarında yalnızca {{file}}, {{sheet}}, {{count}} kullanılabilir ({{{p}}})."
                                : $"\"{f.Alias}\": şablondaki {{{p}}} hiçbir alan adına veya Excel başlığına karşılık gelmiyor.",
                            f.Id));
                }
                break;

            case CustomFieldSource.Json:
                if (!IsValidJson(f.Value, out var jsonError))
                    messages.Add(ValidationMessage.Error(area, $"\"{f.Alias}\": geçersiz JSON – {jsonError}", f.Id));
                break;

            case CustomFieldSource.Constant:
                var r = converter.ConvertText(f.Value ?? string.Empty, f.Type, f.TypeSettings);
                if (r.Status == ConversionStatus.Failed)
                    messages.Add(ValidationMessage.Error(area, $"\"{f.Alias}\": sabit değer {f.Type} tipine dönüştürülemedi – {r.Error}", f.Id));
                break;

            case CustomFieldSource.Timestamp when !string.IsNullOrWhiteSpace(f.Format):
                try
                {
                    _ = DateTime.Now.ToString(f.Format, System.Globalization.CultureInfo.InvariantCulture);
                }
                catch (FormatException)
                {
                    messages.Add(ValidationMessage.Error(area, $"\"{f.Alias}\": geçersiz tarih formatı.", f.Id));
                }
                break;

            case CustomFieldSource.RowNumber or CustomFieldSource.AutoIncrement when isRoot:
                messages.Add(ValidationMessage.Error(area, $"\"{f.Alias}\": satır numarası/otomatik artan kök alanlarda kullanılamaz.", f.Id));
                break;

            case CustomFieldSource.AutoIncrement when f.Step == 0:
                messages.Add(ValidationMessage.Warning(area, $"\"{f.Alias}\": artış değeri 0; tüm kayıtlar aynı değeri alır.", f.Id));
                break;
        }
    }

    public static bool IsValidJson(string? text, out string? error)
    {
        error = null;
        if (string.IsNullOrWhiteSpace(text))
        {
            error = "Değer boş.";
            return false;
        }
        try
        {
            JsonNode.Parse(text);
            return true;
        }
        catch (JsonException ex)
        {
            error = ex.Message;
            return false;
        }
    }

    /// <summary>
    /// Aynı seviyede ad tekrarı (duyarlı: hata, duyarsız: uyarı) ve nokta notasyonu çakışmalarını
    /// (hem "address" hem "address.city") denetler.
    /// </summary>
    public static void ValidatePaths(IEnumerable<(string Alias, string? FieldId)> fields, ValidationArea area, List<ValidationMessage> messages)
    {
        var list = fields.ToList();
        var exact = new Dictionary<string, string?>(StringComparer.Ordinal);
        var insensitive = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var (alias, id) in list)
        {
            if (exact.ContainsKey(alias))
            {
                messages.Add(ValidationMessage.Error(area, $"\"{alias}\" alan adı birden fazla kez kullanılmış.", id));
                continue;
            }
            exact[alias] = id;

            if (insensitive.TryGetValue(alias, out var other))
                messages.Add(ValidationMessage.Warning(area, $"\"{alias}\" ve \"{other}\" yalnızca büyük/küçük harf farkıyla ayrılıyor.", id));
            else
                insensitive[alias] = alias;
        }

        foreach (var (alias, id) in list)
        {
            var segments = alias.Split('.');
            for (var i = 1; i < segments.Length; i++)
            {
                var prefix = string.Join('.', segments[..i]);
                if (exact.ContainsKey(prefix))
                {
                    messages.Add(ValidationMessage.Error(area,
                        $"Nokta notasyonu çakışması: \"{prefix}\" hem değer hem de \"{alias}\" için nesne olarak kullanılıyor.", id));
                    break;
                }
            }
        }
    }

    private void ValidateOutput(MappingProfile profile, List<ValidationMessage> messages)
    {
        if (!shapes.TryGet(profile.Output.Shape, out var strategy))
        {
            messages.Add(ValidationMessage.Error(ValidationArea.Output, $"Bilinmeyen çıktı şekli: {profile.Output.Shape}"));
            return;
        }
        messages.AddRange(strategy.Validate(profile));

        var wrapper = profile.Output.Wrapper;
        if (!wrapper.Enabled)
            return;

        if (string.IsNullOrWhiteSpace(wrapper.DataPropertyName))
            messages.Add(ValidationMessage.Error(ValidationArea.Output, "Veri özelliğinin adı (data) boş olamaz."));
        if (wrapper.IncludeMeta && string.IsNullOrWhiteSpace(wrapper.MetaPropertyName))
            messages.Add(ValidationMessage.Error(ValidationArea.Output, "Meta özelliğinin adı boş olamaz."));

        var roots = wrapper.RootFields.Where(f => f.Include).ToList();
        foreach (var f in roots)
        {
            if (!AliasGenerator.IsValid(f.Alias, out var error))
                messages.Add(ValidationMessage.Error(ValidationArea.Output, $"Kök alan: {error}", f.Id));
            else
                ValidateCustomField(f, [], messages, isRoot: true);
        }

        var names = roots.Where(f => AliasGenerator.IsValid(f.Alias, out _)).Select(f => (f.Alias, (string?)f.Id)).ToList();
        if (wrapper.IncludeMeta)
        {
            foreach (var builtIn in new[] { "source", "sheet", "exportedAt", "count" })
            {
                if (names.Any(n => n.Alias == builtIn))
                    messages.Add(ValidationMessage.Warning(ValidationArea.Output, $"\"{builtIn}\" kök alanı yerleşik meta değerinin üzerine yazacak."));
            }
        }
        else if (names.Any(n => n.Alias == wrapper.DataPropertyName))
        {
            messages.Add(ValidationMessage.Error(ValidationArea.Output, $"Kök alan adı \"{wrapper.DataPropertyName}\" veri özelliğiyle çakışıyor."));
        }
        ValidatePaths(names, ValidationArea.Output, messages);
    }

    private static void ValidateFormat(MappingProfile profile, List<ValidationMessage> messages)
    {
        if (string.IsNullOrWhiteSpace(profile.Format.FileNameTemplate))
            messages.Add(ValidationMessage.Warning(ValidationArea.Export, "Dosya adı şablonu boş; varsayılan şablon kullanılacak."));
        if (profile.Format.IndentSize is < 1 or > 8)
            messages.Add(ValidationMessage.Warning(ValidationArea.Export, "Girinti boyutu 1–8 arasında olmalı; en yakın değer kullanılacak."));
    }

    private static string Describe(FieldDefinition f) => f switch
    {
        ColumnField c => $"{c.ColumnLetter} sütunu ({c.Header})",
        CustomField c => $"Ek alan \"{c.Alias}\"",
        _ => f.Alias
    };
}
