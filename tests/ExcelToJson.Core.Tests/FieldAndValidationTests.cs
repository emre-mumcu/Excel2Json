using ExcelToJson.Core.Expressions;
using ExcelToJson.Core.Models;
using ExcelToJson.Core.Validation;

namespace ExcelToJson.Core.Tests;

public class FieldAndValidationTests : IDisposable
{
    private readonly TempWorkbook _wb = TempWorkbook.FromRows("S",
        ["Ad", "Soyad", "Şehir", "Posta Kodu", "Miktar"],
        ["Ali", "Yılmaz", "İzmir", "35000", "3"],
        ["Ayşe", "Demir", "Ankara", "06000", "abc"]);

    public void Dispose() => _wb.Dispose();

    [Fact]
    public void Dot_notation_creates_nested_objects()
    {
        using var session = _wb.Open();
        var p = session.CreateProfile();
        p.Field<ColumnField>("sehir").Alias = "address.city";
        p.Field<ColumnField>("postaKodu").Alias = "address.zip";
        p.Fields.RemoveAll(f => f.Alias is "soyad" or "miktar");

        var (json, _) = session.RunJson(p, 1);
        Assert.Equal("""[{"ad":"Ali","address":{"city":"İzmir","zip":"35000"}}]""", json);
    }

    [Fact]
    public void Nesting_conflicts_and_duplicates_are_reported()
    {
        var messages = new List<ValidationMessage>();
        ProfileValidator.ValidatePaths([("address", "1"), ("address.city", "2"), ("name", "3"), ("name", "4"), ("Name", "5")],
            ValidationArea.Fields, messages);

        Assert.Contains(messages, m => m.Severity == ValidationSeverity.Error && m.Message.Contains("çakışması") && m.FieldId == "2");
        Assert.Contains(messages, m => m.Severity == ValidationSeverity.Error && m.FieldId == "4");
        Assert.Contains(messages, m => m.Severity == ValidationSeverity.Warning && m.FieldId == "5");
    }

    [Fact]
    public void Validator_checks_selection_aliases_templates_keys_and_json()
    {
        using var session = _wb.Open();
        var p = session.CreateProfile();
        var validator = ProfileValidator.CreateDefault();
        Assert.True(validator.Validate(p).IsValid);

        p.Field<ColumnField>("ad").Alias = "a..b";
        p.Fields.Add(new CustomField { Alias = "full", Source = CustomFieldSource.Template, Value = "{Ad} {yok}" });
        p.Fields.Add(new CustomField { Alias = "extra", Source = CustomFieldSource.Json, Value = "{ bozuk" });
        p.Output.Shape = OutputShapes.KeyedObject;

        var result = validator.Validate(p);
        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, m => m.Message.Contains("boş segment"));
        Assert.Contains(result.Errors, m => m.Message.Contains("{yok}"));
        Assert.DoesNotContain(result.Errors, m => m.Message.Contains("{Ad}"));
        Assert.Contains(result.Errors, m => m.Message.Contains("geçersiz JSON"));
        Assert.Contains(result.Errors, m => m.Area == ValidationArea.Output && m.Message.Contains("Anahtar alanı"));

        foreach (var f in p.Fields) f.Include = false;
        Assert.Contains(validator.Validate(p).Errors, m => m.Message.Contains("En az bir alan"));
    }

    [Fact]
    public void Custom_fields_constant_template_sequence_rownumber_guid_timestamp_json()
    {
        using var session = _wb.Open();
        var p = session.CreateProfile();
        // Sütunlar hariç tutulsa bile şablonlarda kullanılabilir.
        foreach (var f in p.Fields) f.Include = false;
        p.Fields.Add(new CustomField { Alias = "source", Value = "ERP" });
        p.Fields.Add(new CustomField { Alias = "isActive", Value = "true" });
        p.Fields.Add(new CustomField { Alias = "version", Value = "2" });
        p.Fields.Add(new CustomField { Alias = "fullName", Source = CustomFieldSource.Template, Value = "{ad} {Soyad}" });
        p.Fields.Add(new CustomField { Alias = "code", Source = CustomFieldSource.Template, Value = "PRD-{Posta Kodu}" });
        p.Fields.Add(new CustomField { Alias = "seq", Source = CustomFieldSource.AutoIncrement, Start = 100, Step = 10 });
        p.Fields.Add(new CustomField { Alias = "row", Source = CustomFieldSource.RowNumber, RowNumberMode = RowNumberMode.ExcelRow });
        p.Fields.Add(new CustomField { Alias = "no", Source = CustomFieldSource.RowNumber });
        p.Fields.Add(new CustomField { Alias = "at", Source = CustomFieldSource.Timestamp });
        p.Fields.Add(new CustomField { Alias = "tags", Source = CustomFieldSource.Json, Value = """["a",{"b":1}]""" });
        p.Fields.Add(new CustomField { Alias = "id", Source = CustomFieldSource.Guid });

        var (json, result) = session.RunJson(p);
        Assert.True(result.Succeeded, string.Join(", ", result.Issues.Select(i => i.Message)));
        var arr = System.Text.Json.Nodes.JsonNode.Parse(json)!.AsArray();
        var second = arr[1]!.AsObject();

        Assert.Equal("ERP", (string?)second["source"]);
        Assert.True((bool)second["isActive"]!);
        Assert.Equal(2, (int)second["version"]!);
        Assert.Equal("Ayşe Demir", (string?)second["fullName"]);
        Assert.Equal("PRD-06000", (string?)second["code"]);
        Assert.Equal(110, (int)second["seq"]!);
        Assert.Equal(3, (int)second["row"]!);
        Assert.Equal(2, (int)second["no"]!);
        Assert.Equal("2026-11-23T10:00:00Z", (string?)second["at"]);
        Assert.Equal("""["a",{"b":1}]""", second["tags"]!.ToJsonString());
        Assert.True(Guid.TryParse((string?)second["id"], out _));
        Assert.NotEqual((string?)arr[0]!["id"], (string?)second["id"]);
        // Sıralama korunur
        Assert.Equal(["source", "isActive", "version", "fullName", "code", "seq", "row", "no", "at", "tags", "id"],
            second.Select(kv => kv.Key));
    }

    [Fact]
    public void Template_parser_handles_escapes_and_errors()
    {
        var t = TemplateExpression.Parse("{{literal}} {a}-{ b }");
        Assert.Equal(["a", "b"], t.Placeholders);
        Assert.Equal("{literal} 1-2", t.Evaluate(n => n == "a" ? "1" : "2"));
        Assert.False(TemplateExpression.TryParse("{açık", out _, out _));
    }

    [Theory]
    [InlineData(ConversionErrorPolicy.SetNull, """[{"miktar":3},{"miktar":null}]""", true)]
    [InlineData(ConversionErrorPolicy.KeepRaw, """[{"miktar":3},{"miktar":"abc"}]""", true)]
    [InlineData(ConversionErrorPolicy.Stop, null, false)]
    public void Conversion_error_policies(ConversionErrorPolicy policy, string? expected, bool succeeded)
    {
        using var session = _wb.Open();
        var p = session.CreateProfile();
        p.Fields.RemoveAll(f => f.Alias != "miktar");
        p.Field<ColumnField>("miktar").Type = DataType.Integer;
        p.ErrorPolicy = policy;

        var (json, result) = session.RunJson(p);
        Assert.Equal(succeeded, result.Succeeded);
        Assert.Equal(1, result.IssueCount);
        var issue = result.Issues[0];
        Assert.Equal(3, issue.RowNumber);
        Assert.Equal("E", issue.ColumnLetter);
        Assert.Equal("abc", issue.RawValue);
        Assert.Equal(DataType.Integer, issue.TargetType);
        if (expected is not null)
            Assert.Equal(expected, json);
    }

    [Fact]
    public void Constant_json_validation()
    {
        Assert.True(ProfileValidator.IsValidJson("""{"a":[1,2]}""", out _));
        Assert.False(ProfileValidator.IsValidJson("{a:1}", out var error));
        Assert.NotNull(error);
    }
}
