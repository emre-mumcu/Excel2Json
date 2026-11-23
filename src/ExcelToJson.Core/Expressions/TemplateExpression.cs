using System.Text;

namespace ExcelToJson.Core.Expressions;

/// <summary>
/// "{firstName} {lastName}" veya "PRD-{productCode}" gibi şablonlar.
/// Placeholder adı bir alias'a veya Excel başlığına referans verir. "{{" ve "}}" kaçış dizileridir.
/// </summary>
public sealed class TemplateExpression
{
    private readonly List<Segment> _segments;

    private TemplateExpression(List<Segment> segments) => _segments = segments;

    private readonly record struct Segment(bool IsPlaceholder, string Text);

    public IEnumerable<string> Placeholders => _segments.Where(s => s.IsPlaceholder).Select(s => s.Text);

    public static bool TryParse(string? template, out TemplateExpression expression, out string? error)
    {
        var segments = new List<Segment>();
        var text = new StringBuilder();
        template ??= string.Empty;
        error = null;

        for (var i = 0; i < template.Length; i++)
        {
            var ch = template[i];
            if (ch == '{')
            {
                if (i + 1 < template.Length && template[i + 1] == '{')
                {
                    text.Append('{');
                    i++;
                    continue;
                }
                var end = template.IndexOf('}', i + 1);
                if (end < 0)
                {
                    expression = new TemplateExpression([]);
                    error = $"Kapanmamış süslü parantez (konum {i + 1}).";
                    return false;
                }
                var name = template[(i + 1)..end].Trim();
                if (name.Length == 0)
                {
                    expression = new TemplateExpression([]);
                    error = "Boş placeholder: {}";
                    return false;
                }
                if (text.Length > 0)
                {
                    segments.Add(new Segment(false, text.ToString()));
                    text.Clear();
                }
                segments.Add(new Segment(true, name));
                i = end;
            }
            else if (ch == '}')
            {
                if (i + 1 < template.Length && template[i + 1] == '}')
                    i++;
                text.Append('}');
            }
            else
            {
                text.Append(ch);
            }
        }
        if (text.Length > 0)
            segments.Add(new Segment(false, text.ToString()));

        expression = new TemplateExpression(segments);
        return true;
    }

    public static TemplateExpression Parse(string? template) =>
        TryParse(template, out var e, out var error) ? e : throw new FormatException(error);

    /// <summary>Placeholder'ları <paramref name="resolver"/> ile çözerek metni üretir. Çözülemeyen placeholder boş yazılır.</summary>
    public string Evaluate(Func<string, string?> resolver)
    {
        if (_segments.Count == 1 && !_segments[0].IsPlaceholder)
            return _segments[0].Text;
        var sb = new StringBuilder();
        foreach (var s in _segments)
            sb.Append(s.IsPlaceholder ? resolver(s.Text) : s.Text);
        return sb.ToString();
    }
}
