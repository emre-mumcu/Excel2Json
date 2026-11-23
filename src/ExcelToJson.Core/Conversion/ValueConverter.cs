using System.Collections.Concurrent;
using System.Globalization;
using System.Text.Json.Nodes;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Conversion;

/// <summary>
/// Varsayılan değer dönüştürücü. Girdi, kullanıcının seçtiği kültürle (varsayılan tr-TR) parse edilir;
/// JSON çıktısı her zaman kültürden bağımsızdır (sayılar '.', tarihler ISO 8601).
/// </summary>
public sealed class ValueConverter : IValueConverter
{
    public const string IsoDateFormat = "yyyy-MM-dd";
    public const string IsoDateTimeFormat = "yyyy-MM-ddTHH:mm:ss";

    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;
    private static readonly CultureInfo Turkish = CultureInfo.GetCultureInfo("tr-TR");
    private static readonly ConcurrentDictionary<string, HashSet<string>> WordSetCache = new();
    private static readonly StringComparer TurkishIgnoreCase = StringComparer.Create(Turkish, ignoreCase: true);

    public ConversionResult Convert(in CellValue cell, ColumnField field)
    {
        var settings = field.TypeSettings;
        var type = field.Type;

        // Boş değer kontrolü (trim sonrası boş metin de boş sayılır)
        var isBlank = cell.IsBlank || (cell.Kind == CellKind.Text && field.Trim && string.IsNullOrWhiteSpace(cell.Text));
        if (isBlank)
        {
            if (type == DataType.Boolean && settings.EmptyIsFalse)
                return ConversionResult.Ok(JsonValue.Create(false));
            return field.EmptyPolicy switch
            {
                EmptyValuePolicy.Null => ConversionResult.Ok(null),
                EmptyValuePolicy.EmptyString => ConversionResult.Ok(JsonValue.Create(string.Empty)),
                EmptyValuePolicy.Omit => ConversionResult.Omitted,
                EmptyValuePolicy.DefaultValue => ConvertDefault(field),
                _ => ConversionResult.Ok(null)
            };
        }

        if (cell.Kind == CellKind.Error && type != DataType.String)
            return ConversionResult.Fail($"Hücrede Excel hatası var ({cell.Text}).");

        return type switch
        {
            DataType.Auto => ConvertAuto(cell, field.Trim),
            DataType.String => ConversionResult.Ok(JsonValue.Create(ToDisplayString(cell, field.Trim))),
            DataType.Integer => ToInteger(cell, settings, field.Trim),
            DataType.Decimal => ToDecimal(cell, settings, field.Trim),
            DataType.Boolean => ToBoolean(cell, settings, field.Trim),
            DataType.Date or DataType.DateTime => ToDate(cell, type, settings, field.Trim),
            DataType.Array => ToArray(ToDisplayString(cell, field.Trim), settings),
            _ => ConversionResult.Fail($"Bilinmeyen tip: {type}")
        };
    }

    public ConversionResult ConvertText(string text, DataType type, TypeSettings settings)
    {
        text ??= string.Empty;
        switch (type)
        {
            case DataType.Auto:
                return ConvertText(text, TypeInference.InferConstant(text), settings.WithInvariantCulture());
            case DataType.String:
                return ConversionResult.Ok(JsonValue.Create(text));
            case DataType.Array:
                return ToArray(text, settings);
            default:
                if (string.IsNullOrWhiteSpace(text))
                    return type == DataType.Boolean && settings.EmptyIsFalse
                        ? ConversionResult.Ok(JsonValue.Create(false))
                        : ConversionResult.Ok(null);
                var cell = CellValue.FromText(text.Trim());
                return type switch
                {
                    DataType.Integer => ToInteger(cell, settings, true),
                    DataType.Decimal => ToDecimal(cell, settings, true),
                    DataType.Boolean => ToBoolean(cell, settings, true),
                    DataType.Date or DataType.DateTime => ToDate(cell, type, settings, true),
                    _ => ConversionResult.Ok(JsonValue.Create(text))
                };
        }
    }

    private ConversionResult ConvertDefault(ColumnField field)
    {
        var text = field.DefaultValue ?? string.Empty;
        // Auto tipte varsayılan değerin tipi metinden tahmin edilir (invariant).
        return field.Type == DataType.Auto
            ? ConvertText(text, DataType.Auto, field.TypeSettings)
            : ConvertText(text, field.Type, field.TypeSettings);
    }

    /// <summary>Auto: hücrenin doğal tipine göre dönüşüm.</summary>
    private static ConversionResult ConvertAuto(in CellValue cell, bool trim) => cell.Kind switch
    {
        CellKind.Number when Math.Abs(cell.Number) > TypeInference.MaxSafeInteger
            => ConversionResult.Ok(JsonValue.Create(NumberToPlainString(cell.Number))),
        CellKind.Number when cell.Number == Math.Floor(cell.Number)
            => ConversionResult.Ok(JsonValue.Create((long)cell.Number)),
        CellKind.Number => ConversionResult.Ok(JsonValue.Create(Normalize((decimal)cell.Number))),
        CellKind.Boolean => ConversionResult.Ok(JsonValue.Create(cell.Boolean)),
        CellKind.DateTime => ConversionResult.Ok(JsonValue.Create(cell.DateTime.TimeOfDay == TimeSpan.Zero
            ? cell.DateTime.ToString(IsoDateFormat, Invariant)
            : cell.DateTime.ToString(IsoDateTimeFormat, Invariant))),
        CellKind.TimeSpan => ConversionResult.Ok(JsonValue.Create(cell.TimeSpan.ToString("c", Invariant))),
        _ => ConversionResult.Ok(JsonValue.Create(ToDisplayString(cell, trim)))
    };

    /// <summary>String tipi için hücrenin görünen metni. Bilimsel gösterime düşen büyük sayılar tam yazılır.</summary>
    public static string ToDisplayString(in CellValue cell, bool trim)
    {
        string s;
        if (cell.Kind == CellKind.Text)
            s = cell.Text ?? string.Empty;
        else if (cell.Kind == CellKind.Number && (string.IsNullOrEmpty(cell.Formatted) || cell.Formatted.Contains("E+", StringComparison.OrdinalIgnoreCase)))
            s = NumberToPlainString(cell.Number);
        else
            s = cell.Formatted ?? string.Empty;
        return trim ? s.Trim() : s;
    }

    private static string NumberToPlainString(double number)
    {
        // Tam sayılar double'ın tuttuğu değerle birebir (bilimsel gösterim ve 15 hane yuvarlaması olmadan) yazılır.
        if (number == Math.Floor(number) && !double.IsInfinity(number))
            return new System.Numerics.BigInteger(number).ToString(Invariant);
        if (Math.Abs(number) < 7.9e28)
            return Normalize((decimal)number).ToString(Invariant);
        return number.ToString("R", Invariant);
    }

    private static ConversionResult ToInteger(in CellValue cell, TypeSettings settings, bool trim)
    {
        switch (cell.Kind)
        {
            case CellKind.Number:
                if (cell.Number != Math.Floor(cell.Number))
                    return ConversionResult.Fail("Ondalıklı değer tam sayıya dönüştürülemez.");
                if (cell.Number is > long.MaxValue or < long.MinValue)
                    return ConversionResult.Fail("Değer tam sayı aralığının dışında.");
                return ConversionResult.Ok(JsonValue.Create((long)cell.Number));
            case CellKind.Boolean:
                return ConversionResult.Fail("Mantıksal değer tam sayıya dönüştürülemez.");
            case CellKind.Text:
                var text = trim ? cell.Text!.Trim() : cell.Text!;
                if (long.TryParse(text, NumberStyles.Integer | NumberStyles.AllowThousands, GetCulture(settings), out var l))
                    return ConversionResult.Ok(JsonValue.Create(l));
                if (decimal.TryParse(text, NumberStyles.Number, GetCulture(settings), out var d) && d == decimal.Truncate(d)
                    && d is <= long.MaxValue and >= long.MinValue)
                    return ConversionResult.Ok(JsonValue.Create((long)d));
                return ConversionResult.Fail($"\"{text}\" tam sayı değil.");
            default:
                return ConversionResult.Fail("Değer tam sayıya dönüştürülemez.");
        }
    }

    private static ConversionResult ToDecimal(in CellValue cell, TypeSettings settings, bool trim)
    {
        switch (cell.Kind)
        {
            case CellKind.Number:
                try
                {
                    return ConversionResult.Ok(JsonValue.Create(Normalize((decimal)cell.Number)));
                }
                catch (OverflowException)
                {
                    return ConversionResult.Ok(JsonValue.Create(cell.Number));
                }
            case CellKind.Text:
                var text = trim ? cell.Text!.Trim() : cell.Text!;
                if (decimal.TryParse(text, NumberStyles.Number | NumberStyles.AllowExponent, GetCulture(settings), out var d))
                    return ConversionResult.Ok(JsonValue.Create(Normalize(d)));
                return ConversionResult.Fail($"\"{text}\" ondalık sayı değil ({CultureLabel(settings)}).");
            default:
                return ConversionResult.Fail("Değer ondalık sayıya dönüştürülemez.");
        }
    }

    private static ConversionResult ToBoolean(in CellValue cell, TypeSettings settings, bool trim)
    {
        switch (cell.Kind)
        {
            case CellKind.Boolean:
                return ConversionResult.Ok(JsonValue.Create(cell.Boolean));
            case CellKind.Number when cell.Number is 0 or 1:
                return ConversionResult.Ok(JsonValue.Create(cell.Number == 1));
            case CellKind.Number:
            case CellKind.Text:
                var text = cell.Kind == CellKind.Text ? cell.Text! : cell.Formatted;
                if (trim) text = text.Trim();
                if (WordSet(settings.TrueValues).Contains(text))
                    return ConversionResult.Ok(JsonValue.Create(true));
                if (WordSet(settings.FalseValues).Contains(text))
                    return ConversionResult.Ok(JsonValue.Create(false));
                return ConversionResult.Fail($"\"{text}\" true/false değerlerinden biri değil.");
            default:
                return ConversionResult.Fail("Değer mantıksal (true/false) değere dönüştürülemez.");
        }
    }

    private static ConversionResult ToDate(in CellValue cell, DataType type, TypeSettings settings, bool trim)
    {
        DateTime value;
        switch (cell.Kind)
        {
            case CellKind.DateTime:
                value = cell.DateTime;
                break;
            case CellKind.Number:
                // Excel tarihleri seri numarası olarak saklar (1900 tarih sistemi).
                if (cell.Number is < -657435 or > 2958465)
                    return ConversionResult.Fail("Sayı geçerli bir Excel tarih seri numarası değil.");
                value = DateTime.FromOADate(cell.Number);
                break;
            case CellKind.Text:
                var text = trim ? cell.Text!.Trim() : cell.Text!;
                if (!TryParseDate(text, settings, out value))
                    return ConversionResult.Fail($"\"{text}\" tarih olarak okunamadı.");
                break;
            default:
                return ConversionResult.Fail("Değer tarihe dönüştürülemez.");
        }

        var format = string.IsNullOrWhiteSpace(settings.DateFormat)
            ? (type == DataType.Date ? IsoDateFormat : IsoDateTimeFormat)
            : settings.DateFormat;
        try
        {
            return ConversionResult.Ok(JsonValue.Create(value.ToString(format, Invariant)));
        }
        catch (FormatException)
        {
            return ConversionResult.Fail($"Geçersiz tarih formatı: {format}");
        }
    }

    private static bool TryParseDate(string text, TypeSettings settings, out DateTime value)
    {
        if (!string.IsNullOrWhiteSpace(settings.DateInputFormat) &&
            DateTime.TryParseExact(text, settings.DateInputFormat, GetCulture(settings), DateTimeStyles.AllowWhiteSpaces, out value))
            return true;
        if (DateTime.TryParse(text, GetCulture(settings), DateTimeStyles.AllowWhiteSpaces, out value))
            return true;
        return DateTime.TryParse(text, Invariant, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.RoundtripKind, out value);
    }

    private ConversionResult ToArray(string text, TypeSettings settings)
    {
        var separator = string.IsNullOrEmpty(settings.ArraySeparator) ? "," : settings.ArraySeparator;
        var array = new JsonArray();
        if (string.IsNullOrWhiteSpace(text))
            return ConversionResult.Ok(array);

        var elementType = settings.ArrayElementType is DataType.Array or DataType.Auto ? DataType.String : settings.ArrayElementType;
        foreach (var part in text.Split(separator, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))
        {
            var r = ConvertText(part, elementType, settings);
            if (r.Status == ConversionStatus.Failed)
                return ConversionResult.Fail($"Dizi elemanı dönüştürülemedi: {r.Error}");
            array.Add(r.Value);
        }
        return ConversionResult.Ok(array);
    }

    private static HashSet<string> WordSet(string list) =>
        WordSetCache.GetOrAdd(list ?? string.Empty, l => new HashSet<string>(
            l.Split([',', ';', '|'], StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries),
            TurkishIgnoreCase));

    private static CultureInfo GetCulture(TypeSettings settings) =>
        string.IsNullOrWhiteSpace(settings.Culture) ? Invariant :
        settings.Culture.Equals("tr-TR", StringComparison.OrdinalIgnoreCase) ? Turkish :
        CultureInfo.GetCultureInfo(settings.Culture);

    private static string CultureLabel(TypeSettings settings) =>
        string.IsNullOrWhiteSpace(settings.Culture) ? "Invariant" : settings.Culture;

    /// <summary>Gereksiz sondaki sıfırları atar (12.50 → 12.5).</summary>
    private static decimal Normalize(decimal value) => value / 1.000000000000000000000000000000000m;
}

internal static class TypeSettingsExtensions
{
    /// <summary>Auto sabit değerleri invariant kültürle yorumlanır.</summary>
    public static TypeSettings WithInvariantCulture(this TypeSettings settings)
    {
        var c = settings.Clone();
        c.Culture = string.Empty;
        return c;
    }
}
