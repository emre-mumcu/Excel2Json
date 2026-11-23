using ExcelToJson.Core.Mapping;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Profiles;

public sealed record ReconcileResult(MappingProfile Profile, IReadOnlyList<string> Warnings);

/// <summary>
/// Yüklenen profili mevcut sayfa şemasına uyarlar. Sütunlar önce <b>başlık adına</b>, bulunamazsa
/// sütun harfine göre eşleştirilir. Bulunamayan sütunlar uyarı olarak raporlanır; profilde olmayan
/// yeni sütunlar listenin sonuna "dahil değil" olarak eklenir.
/// </summary>
public static class ProfileReconciler
{
    public static ReconcileResult Reconcile(MappingProfile source, SheetSchema schema)
    {
        var profile = source.Clone();
        var warnings = new List<string>();
        var used = new HashSet<int>();
        var fields = new List<FieldDefinition>();

        // Başlık eşleşmeleri önce yapılır ki harf yedeklemesi başka bir sütunun başlığını "çalmasın".
        var byHeader = new Dictionary<ColumnField, ColumnInfo>();
        foreach (var c in profile.Fields.OfType<ColumnField>())
        {
            var match = schema.Columns.FirstOrDefault(col => !used.Contains(col.Index) &&
                string.Equals(col.Header.Trim(), c.Header.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match is not null)
            {
                byHeader[c] = match;
                used.Add(match.Index);
            }
        }

        foreach (var field in profile.Fields)
        {
            if (field is not ColumnField column)
            {
                fields.Add(field);
                continue;
            }

            if (!byHeader.TryGetValue(column, out var match))
            {
                match = schema.Columns.FirstOrDefault(col => !used.Contains(col.Index) &&
                    string.Equals(col.Letter, column.ColumnLetter, StringComparison.OrdinalIgnoreCase));
                if (match is not null)
                {
                    used.Add(match.Index);
                    warnings.Add($"\"{column.Header}\" başlığı bulunamadı; {column.ColumnLetter} sütunu (\"{match.Header}\") ile eşleştirildi.");
                }
            }

            if (match is null)
            {
                warnings.Add($"\"{column.Header}\" ({column.ColumnLetter}) sütunu Excel'de artık bulunmuyor; profilden çıkarıldı.");
                continue;
            }

            column.ColumnIndex = match.Index;
            column.ColumnLetter = match.Letter;
            column.Header = match.Header;
            fields.Add(column);
        }

        var newColumns = schema.Columns.Where(c => !used.Contains(c.Index)).ToList();
        if (newColumns.Count > 0)
        {
            var existing = fields.Select(f => f.Alias).ToList();
            var aliases = AliasGenerator.MakeUnique(existing.Concat(newColumns.Select(c =>
                AliasGenerator.Generate(c.Header, profile.NamingStyle, profile.TransliterateTurkish, "column" + c.Letter))))
                .Skip(existing.Count).ToList();
            for (var i = 0; i < newColumns.Count; i++)
            {
                var f = FieldFactory.CreateColumnField(newColumns[i], aliases[i]);
                f.Include = false;
                fields.Add(f);
            }
            warnings.Add($"Profilde olmayan {newColumns.Count} yeni sütun listeye eklendi (dahil değil).");
        }

        profile.Fields = fields;
        profile.SheetName = schema.SheetName;
        return new ReconcileResult(profile, warnings);
    }
}
