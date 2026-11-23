using ExcelToJson.Core.Conversion;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Tests;

public class ValueConverterTests
{
    private readonly ValueConverter _converter = new();

    private string? Convert(CellValue cell, DataType type, Action<ColumnField>? configure = null)
    {
        var field = new ColumnField { Alias = "x", Type = type };
        configure?.Invoke(field);
        var r = _converter.Convert(cell, field);
        return r.Status switch
        {
            ConversionStatus.Omit => "<omit>",
            ConversionStatus.Failed => "<fail>",
            _ => r.Value?.ToJsonString() ?? "null"
        };
    }

    [Theory]
    [InlineData("1.234,56", "tr-TR", "1234.56")]
    [InlineData("12,5", "tr-TR", "12.5")]
    [InlineData("1,234.56", "", "1234.56")]
    [InlineData("-0,75", "tr-TR", "-0.75")]
    public void Decimal_text_is_parsed_with_selected_culture(string text, string culture, string expected) =>
        Assert.Equal(expected, Convert(CellValue.FromText(text), DataType.Decimal, f => f.TypeSettings.Culture = culture));

    [Fact]
    public void Numbers_are_written_invariant_without_trailing_zeros()
    {
        Assert.Equal("10.5", Convert(CellValue.FromNumber(10.50), DataType.Decimal));
        Assert.Equal("42", Convert(CellValue.FromNumber(42), DataType.Integer));
        Assert.Equal("<fail>", Convert(CellValue.FromNumber(4.2), DataType.Integer));
        Assert.Equal("1234", Convert(CellValue.FromText("1.234"), DataType.Integer));
        Assert.Equal("<fail>", Convert(CellValue.FromText("abc"), DataType.Integer));
    }

    [Fact]
    public void Dates_from_serial_numbers_datetime_cells_and_text()
    {
        Assert.Equal("\"2026-11-23\"", Convert(CellValue.FromNumber(new DateTime(2026, 11, 23).ToOADate()), DataType.Date));
        Assert.Equal("\"2026-11-23T14:30:00\"", Convert(CellValue.FromDateTime(new DateTime(2026, 11, 23, 14, 30, 0)), DataType.DateTime));
        Assert.Equal("\"2026-11-23\"", Convert(CellValue.FromText("23.11.2026"), DataType.Date));
        Assert.Equal("\"23/11/2026\"", Convert(CellValue.FromText("23.11.2026"), DataType.Date, f => f.TypeSettings.DateFormat = "dd'/'MM'/'yyyy"));
        Assert.Equal("\"2026-11-23\"", Convert(CellValue.FromText("11-23-2026"), DataType.Date, f => f.TypeSettings.DateInputFormat = "MM-dd-yyyy"));
        Assert.Equal("<fail>", Convert(CellValue.FromText("yarın"), DataType.Date));
    }

    [Theory]
    [InlineData("evet", "true")]
    [InlineData("HAYIR", "false")]
    [InlineData("Evet", "true")]
    [InlineData("x", "true")]
    [InlineData("1", "true")]
    [InlineData("yok", "false")]
    [InlineData("belki", "<fail>")]
    public void Boolean_text_variants(string text, string expected) =>
        Assert.Equal(expected, Convert(CellValue.FromText(text), DataType.Boolean));

    [Fact]
    public void Boolean_from_cells_and_empty_as_false()
    {
        Assert.Equal("true", Convert(CellValue.FromBoolean(true), DataType.Boolean));
        Assert.Equal("false", Convert(CellValue.FromNumber(0), DataType.Boolean));
        Assert.Equal("false", Convert(CellValue.Blank, DataType.Boolean, f => f.TypeSettings.EmptyIsFalse = true));
        Assert.Equal("null", Convert(CellValue.Blank, DataType.Boolean));
    }

    [Fact]
    public void Delimited_arrays_with_element_types()
    {
        Assert.Equal("[\"a\",\"b\",\"c\"]", Convert(CellValue.FromText("a, b ,c"), DataType.Array));
        Assert.Equal("[1,2,3]", Convert(CellValue.FromText("1;2;3"), DataType.Array, f =>
        {
            f.TypeSettings.ArraySeparator = ";";
            f.TypeSettings.ArrayElementType = DataType.Integer;
        }));
        Assert.Equal("[]", Convert(CellValue.FromText(" "), DataType.Array, f => f.EmptyPolicy = EmptyValuePolicy.DefaultValue));
        Assert.Equal("<fail>", Convert(CellValue.FromText("1|x"), DataType.Array, f =>
        {
            f.TypeSettings.ArraySeparator = "|";
            f.TypeSettings.ArrayElementType = DataType.Integer;
        }));
    }

    [Fact]
    public void Large_numbers_and_leading_zeros_stay_strings()
    {
        // 15 haneden uzun sayılar Auto modda string (double'ın tuttuğu değer bilimsel gösterimsiz yazılır)
        Assert.Equal("\"12345678901234568\"", Convert(CellValue.FromNumber(12345678901234567d), DataType.Auto));
        Assert.Equal("\"1234567890123456\"", Convert(CellValue.FromNumber(1234567890123456d), DataType.Auto));
        Assert.Equal("12345678901", Convert(CellValue.FromNumber(12345678901d), DataType.Auto));
        // Bilimsel gösterimle formatlanmış sayı String tipinde tam yazılır
        Assert.Equal("\"8690000000012\"", Convert(CellValue.FromNumber(8690000000012d, "8,69E+12"), DataType.String));
        // Metin olarak girilmiş baştaki sıfırlar korunur
        Assert.Equal("\"00123\"", Convert(CellValue.FromText("00123"), DataType.Auto));
    }

    [Fact]
    public void Empty_value_policies()
    {
        Assert.Equal("null", Convert(CellValue.Blank, DataType.String));
        Assert.Equal("\"\"", Convert(CellValue.Blank, DataType.String, f => f.EmptyPolicy = EmptyValuePolicy.EmptyString));
        Assert.Equal("<omit>", Convert(CellValue.Blank, DataType.String, f => f.EmptyPolicy = EmptyValuePolicy.Omit));
        Assert.Equal("5", Convert(CellValue.Blank, DataType.Integer, f => { f.EmptyPolicy = EmptyValuePolicy.DefaultValue; f.DefaultValue = "5"; }));
        Assert.Equal("true", Convert(CellValue.Blank, DataType.Auto, f => { f.EmptyPolicy = EmptyValuePolicy.DefaultValue; f.DefaultValue = "true"; }));
        // Trim açıkken yalnızca boşluk içeren metin boş sayılır
        Assert.Equal("null", Convert(CellValue.FromText("   "), DataType.String));
        Assert.Equal("\"   \"", Convert(CellValue.FromText("   "), DataType.String, f => f.Trim = false));
    }

    [Fact]
    public void Trim_and_auto_types()
    {
        Assert.Equal("\"abc\"", Convert(CellValue.FromText("  abc "), DataType.String));
        Assert.Equal("3", Convert(CellValue.FromNumber(3), DataType.Auto));
        Assert.Equal("3.25", Convert(CellValue.FromNumber(3.25), DataType.Auto));
        Assert.Equal("true", Convert(CellValue.FromBoolean(true), DataType.Auto));
        Assert.Equal("\"2026-01-02\"", Convert(CellValue.FromDateTime(new DateTime(2026, 1, 2)), DataType.Auto));
    }

    [Fact]
    public void Error_cells_fail_unless_string()
    {
        var err = new CellValue { Kind = CellKind.Error, Text = "#N/A", Formatted = "#N/A" };
        Assert.Equal("<fail>", Convert(err, DataType.Integer));
        Assert.Equal("\"#N/A\"", Convert(err, DataType.String));
    }
}
