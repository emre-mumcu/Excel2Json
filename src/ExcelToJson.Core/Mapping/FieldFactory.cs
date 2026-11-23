using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Mapping;

/// <summary>Şemadan alan listesi üretir ve alias'ları yeniden hesaplar.</summary>
public static class FieldFactory
{
    public static List<FieldDefinition> CreateColumnFields(SheetSchema schema, NamingStyle style, bool transliterate)
    {
        var aliases = AliasGenerator.MakeUnique(schema.Columns.Select(c =>
            AliasGenerator.Generate(c.Header, style, transliterate, "column" + c.Letter)));

        return schema.Columns.Select((c, i) => (FieldDefinition)CreateColumnField(c, aliases[i])).ToList();
    }

    public static ColumnField CreateColumnField(ColumnInfo column, string alias) => new()
    {
        ColumnIndex = column.Index,
        ColumnLetter = column.Letter,
        Header = column.Header,
        Alias = alias,
        Include = true,
        Type = DataType.Auto
    };

    /// <summary>Tüm sütun alanlarının alias'larını seçilen stille yeniden üretir (ek alanlara dokunmaz).</summary>
    public static void RegenerateAliases(IList<FieldDefinition> fields, NamingStyle style, bool transliterate)
    {
        var columns = fields.OfType<ColumnField>().ToList();
        var reserved = fields.OfType<CustomField>().Select(f => f.Alias);
        var generated = AliasGenerator.MakeUnique(reserved.Concat(columns.Select(c =>
            AliasGenerator.Generate(c.Header, style, transliterate, "column" + c.ColumnLetter)))).Skip(fields.Count - columns.Count).ToList();
        for (var i = 0; i < columns.Count; i++)
            columns[i].Alias = generated[i];
    }

    /// <summary>
    /// Başlık/okuma ayarları değiştiğinde alanları yeni şemayla senkronlar (sütun numarasına göre).
    /// Kullanıcının elle değiştirmediği (otomatik üretilmiş) alias'lar yeni başlıktan yeniden üretilir;
    /// artık olmayan sütunlar çıkarılır, yeni sütunlar eklenir. Ek alanlar ve sıralama korunur.
    /// </summary>
    public static void SyncWithSchema(MappingProfile profile, SheetSchema schema)
    {
        var byIndex = schema.Columns.ToDictionary(c => c.Index);
        var seen = new HashSet<int>();
        var result = new List<FieldDefinition>();

        foreach (var field in profile.Fields)
        {
            if (field is not ColumnField column)
            {
                result.Add(field);
                continue;
            }
            if (!byIndex.TryGetValue(column.ColumnIndex, out var info) || !seen.Add(info.Index))
                continue;

            var autoAlias = AliasGenerator.Generate(column.Header, profile.NamingStyle, profile.TransliterateTurkish, "column" + column.ColumnLetter);
            var wasAuto = column.Alias == autoAlias || column.Alias.StartsWith(autoAlias + "_", StringComparison.Ordinal);
            column.Header = info.Header;
            column.ColumnLetter = info.Letter;
            if (wasAuto)
                column.Alias = AliasGenerator.Generate(info.Header, profile.NamingStyle, profile.TransliterateTurkish, "column" + info.Letter);
            result.Add(column);
        }

        foreach (var info in schema.Columns.Where(c => !seen.Contains(c.Index)))
            result.Add(CreateColumnField(info, AliasGenerator.Generate(info.Header, profile.NamingStyle, profile.TransliterateTurkish, "column" + info.Letter)));

        // Otomatik alias'lar çakışırsa benzersizleştir (elle verilenlere dokunmadan sondan ekleme yapılır).
        var unique = AliasGenerator.MakeUnique(result.Select(f => f.Alias));
        for (var i = 0; i < result.Count; i++)
            result[i].Alias = unique[i];

        profile.Fields = result;
    }

    public static CustomField CreateCustomField(IEnumerable<FieldDefinition> existing, CustomFieldSource source = CustomFieldSource.Constant)
    {
        var names = existing.Select(f => f.Alias).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var baseName = source switch
        {
            CustomFieldSource.Guid => "id",
            CustomFieldSource.RowNumber => "rowNumber",
            CustomFieldSource.AutoIncrement => "sequence",
            CustomFieldSource.Timestamp => "exportedAt",
            CustomFieldSource.Template => "text",
            CustomFieldSource.Json => "extra",
            _ => "newField"
        };
        var name = baseName;
        var n = 2;
        while (names.Contains(name))
            name = baseName + n++;
        return new CustomField { Alias = name, Source = source };
    }
}
