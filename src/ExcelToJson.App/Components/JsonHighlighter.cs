using System.Net;
using System.Text;
using Microsoft.AspNetCore.Components;

namespace ExcelToJson.App.Components;

/// <summary>
/// JSON metnini satır numaralı, sözdizimi renklendirilmiş HTML'e çevirir (harici kütüphane gerektirmez).
/// Girdi Utf8JsonWriter çıktısı olduğundan her zaman geçerli JSON'dur; basit bir tarayıcı yeterlidir.
/// </summary>
public static class JsonHighlighter
{
    public const int MaxChars = 400_000;

    public static MarkupString ToHtml(string? json)
    {
        if (string.IsNullOrEmpty(json))
            return new MarkupString(string.Empty);

        var truncated = json.Length > MaxChars;
        if (truncated)
            json = json[..MaxChars];

        var sb = new StringBuilder(json.Length * 3);
        var lineNo = 1;
        sb.Append("<div class=\"ln\"><span class=\"n\">1</span><span class=\"t\">");

        var i = 0;
        while (i < json.Length)
        {
            var c = json[i];
            switch (c)
            {
                case '\r':
                    i++;
                    break;
                case '\n':
                    lineNo++;
                    sb.Append("</span></div><div class=\"ln\"><span class=\"n\">").Append(lineNo).Append("</span><span class=\"t\">");
                    i++;
                    break;
                case '"':
                {
                    var end = i + 1;
                    while (end < json.Length && json[end] != '"')
                        end += json[end] == '\\' ? 2 : 1;
                    end = Math.Min(end + 1, json.Length);
                    var token = json[i..end];
                    // Ardından ':' geliyorsa anahtar
                    var j = end;
                    while (j < json.Length && json[j] == ' ') j++;
                    var cls = j < json.Length && json[j] == ':' ? "j-key" : "j-str";
                    Span(sb, cls, token);
                    i = end;
                    break;
                }
                case '{' or '}' or '[' or ']' or ',' or ':':
                    Span(sb, "j-punc", c.ToString());
                    i++;
                    break;
                case 't' when Matches(json, i, "true"):
                    Span(sb, "j-bool", "true");
                    i += 4;
                    break;
                case 'f' when Matches(json, i, "false"):
                    Span(sb, "j-bool", "false");
                    i += 5;
                    break;
                case 'n' when Matches(json, i, "null"):
                    Span(sb, "j-null", "null");
                    i += 4;
                    break;
                default:
                    if (c == '-' || char.IsDigit(c))
                    {
                        var end = i + 1;
                        while (end < json.Length && (char.IsDigit(json[end]) || json[end] is '.' or 'e' or 'E' or '+' or '-'))
                            end++;
                        Span(sb, "j-num", json[i..end]);
                        i = end;
                    }
                    else
                    {
                        sb.Append(WebUtility.HtmlEncode(c.ToString()));
                        i++;
                    }
                    break;
            }
        }

        sb.Append("</span></div>");
        if (truncated)
            sb.Append("<div class=\"ln\"><span class=\"n\"></span><span class=\"t j-null\">… (önizleme kısaltıldı)</span></div>");
        return new MarkupString(sb.ToString());
    }

    private static bool Matches(string s, int i, string word) =>
        string.CompareOrdinal(s, i, word, 0, word.Length) == 0;

    private static void Span(StringBuilder sb, string cls, string text) =>
        sb.Append("<span class=\"").Append(cls).Append("\">").Append(WebUtility.HtmlEncode(text)).Append("</span>");
}
