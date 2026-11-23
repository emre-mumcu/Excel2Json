using System.Text;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Mapping;

/// <summary>Excel başlığından JSON alan adı (alias) üretir.</summary>
public static class AliasGenerator
{
    /// <summary>Türkçe karakterleri ASCII karşılıklarına çevirir (ı→i, ş→s, ğ→g, ü→u, ö→o, ç→c).</summary>
    public static string Transliterate(string text)
    {
        var sb = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            sb.Append(ch switch
            {
                'ı' => 'i', 'İ' => 'I',
                'ş' => 's', 'Ş' => 'S',
                'ğ' => 'g', 'Ğ' => 'G',
                'ü' => 'u', 'Ü' => 'U',
                'ö' => 'o', 'Ö' => 'O',
                'ç' => 'c', 'Ç' => 'C',
                _ => ch
            });
        }
        return sb.ToString();
    }

    public static string Generate(string header, NamingStyle style, bool transliterateTurkish, string fallback = "field")
    {
        header = (header ?? string.Empty).Trim();
        if (transliterateTurkish)
            header = Transliterate(header);

        if (style == NamingStyle.Original)
            return header.Length == 0 ? fallback : header;

        var words = SplitWords(header);
        if (words.Count == 0)
            return fallback;

        var sb = new StringBuilder();
        for (var i = 0; i < words.Count; i++)
        {
            var w = words[i];
            switch (style)
            {
                case NamingStyle.CamelCase:
                    sb.Append(i == 0 ? ToLower(w) : Capitalize(w));
                    break;
                case NamingStyle.PascalCase:
                    sb.Append(Capitalize(w));
                    break;
                case NamingStyle.SnakeCase:
                    if (i > 0) sb.Append('_');
                    sb.Append(ToLower(w));
                    break;
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Kelimelere böler: harf/rakam dışı karakterler ayraçtır; ayrıca küçük→büyük harf geçişleri
    /// (productCode → product, Code) ve harf↔rakam olmayan ardışık büyük harf grupları (HTTPServer → HTTP, Server) ayrılır.
    /// </summary>
    internal static List<string> SplitWords(string text)
    {
        var words = new List<string>();
        var current = new StringBuilder();

        void Flush()
        {
            if (current.Length > 0)
            {
                words.Add(current.ToString());
                current.Clear();
            }
        }

        for (var i = 0; i < text.Length; i++)
        {
            var ch = text[i];
            if (!char.IsLetterOrDigit(ch))
            {
                Flush();
                continue;
            }

            if (current.Length > 0)
            {
                var prev = text[i - 1];
                var next = i + 1 < text.Length ? text[i + 1] : '\0';
                var lowerToUpper = char.IsLower(prev) && char.IsUpper(ch);
                var acronymEnd = char.IsUpper(prev) && char.IsUpper(ch) && char.IsLower(next);
                if (lowerToUpper || acronymEnd)
                    Flush();
            }
            current.Append(ch);
        }
        Flush();
        return words;
    }

    private static string ToLower(string w)
    {
        var sb = new StringBuilder(w.Length);
        foreach (var ch in w)
            sb.Append(ch switch { 'İ' => 'i', 'I' => 'i', _ => char.ToLowerInvariant(ch) });
        return sb.ToString();
    }

    private static string Capitalize(string w)
    {
        if (w.Length == 0)
            return w;
        var lower = ToLower(w);
        var first = lower[0] switch { 'i' => 'I', 'ı' => 'I', _ => char.ToUpperInvariant(lower[0]) };
        // Türkçe 'i' büyük harfi 'İ' olmalı; ama orijinal başlıkta 'İ' varsa onu koruyalım.
        if (w[0] == 'İ') first = 'İ';
        return first + lower[1..];
    }

    /// <summary>Tekrarlanan adları benzersizleştirir: name, name_2, name_3…</summary>
    public static List<string> MakeUnique(IEnumerable<string> names)
    {
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<string>();
        foreach (var name in names)
        {
            var candidate = name;
            var n = 2;
            while (!used.Add(candidate))
                candidate = $"{name}_{n++}";
            result.Add(candidate);
        }
        return result;
    }

    /// <summary>Alias geçerli mi? Boş olamaz, nokta ile başlayıp bitemez, boş segment içeremez.</summary>
    public static bool IsValid(string? alias, out string? error)
    {
        error = null;
        if (string.IsNullOrWhiteSpace(alias))
        {
            error = "Alan adı boş olamaz.";
            return false;
        }
        if (alias != alias.Trim())
        {
            error = "Alan adı boşlukla başlayıp bitemez.";
            return false;
        }
        if (alias.Split('.').Any(s => s.Length == 0))
        {
            error = "Nokta notasyonunda boş segment olamaz (ör. \"a..b\" veya \".a\").";
            return false;
        }
        if (alias.Any(char.IsControl))
        {
            error = "Alan adı kontrol karakteri içeremez.";
            return false;
        }
        return true;
    }
}
