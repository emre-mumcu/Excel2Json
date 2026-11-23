using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Tests;

public class OutputShapeTests : IDisposable
{
    private readonly TempWorkbook _wb = TempWorkbook.FromRows("Ürünler",
        ["Kod", "Ad", "Kategori", "Fiyat"],
        ["P1", "Laptop", "Elektronik", 1500.5],
        ["P2", "Mouse", "Elektronik", 25],
        ["P3", "Kalem", "Kırtasiye", 2]);

    public void Dispose() => _wb.Dispose();

    private (string Json, Json.ExportRunResult Result) Run(Action<MappingProfile> configure)
    {
        using var session = _wb.Open();
        var profile = session.CreateProfile();
        configure(profile);
        return session.RunJson(profile);
    }

    private static string Id(MappingProfile p, string alias) => p.Fields.Single(f => f.Alias == alias).Id;

    [Fact]
    public void Array_of_objects_with_typed_values_and_unescaped_turkish()
    {
        var (json, _) = Run(_ => { });
        Assert.Equal("""[{"kod":"P1","ad":"Laptop","kategori":"Elektronik","fiyat":1500.5},{"kod":"P2","ad":"Mouse","kategori":"Elektronik","fiyat":25},{"kod":"P3","ad":"Kalem","kategori":"Kırtasiye","fiyat":2}]""", json);
    }

    [Fact]
    public void Turkish_characters_can_be_escaped()
    {
        var (json, _) = Run(p => p.Format.UnescapedUnicode = false);
        Assert.Contains("K\\u0131rtasiye", json);
    }

    [Fact]
    public void Keyed_object_by_field_without_key_in_object()
    {
        var (json, _) = Run(p =>
        {
            p.Output.Shape = OutputShapes.KeyedObject;
            p.Output.KeyFieldId = Id(p, "kod");
            p.Output.KeepKeyInObject = false;
            p.Fields.RemoveAll(f => f.Alias is "kategori" or "fiyat");
        });
        Assert.Equal("""{"P1":{"ad":"Laptop"},"P2":{"ad":"Mouse"},"P3":{"ad":"Kalem"}}""", json);
    }

    [Fact]
    public void Keyed_object_by_row_number()
    {
        var (json, _) = Run(p =>
        {
            p.Output.Shape = OutputShapes.KeyedObject;
            p.Output.KeyFieldId = OutputShapes.RowNumberKey;
            p.Fields.RemoveAll(f => f.Alias != "kod");
        });
        Assert.Equal("""{"2":{"kod":"P1"},"3":{"kod":"P2"},"4":{"kod":"P3"}}""", json);
    }

    [Theory]
    [InlineData(DuplicateKeyPolicy.FirstWins, """{"Elektronik":"Laptop","Kırtasiye":"Kalem"}""")]
    [InlineData(DuplicateKeyPolicy.LastWins, """{"Elektronik":"Mouse","Kırtasiye":"Kalem"}""")]
    [InlineData(DuplicateKeyPolicy.ToArray, """{"Elektronik":["Laptop","Mouse"],"Kırtasiye":"Kalem"}""")]
    public void Key_value_with_duplicate_policies(DuplicateKeyPolicy policy, string expected)
    {
        var (json, result) = Run(p =>
        {
            p.Output.Shape = OutputShapes.KeyValue;
            p.Output.KeyFieldId = Id(p, "kategori");
            p.Output.ValueFieldId = Id(p, "ad");
            p.Output.DuplicateKeyPolicy = policy;
        });
        Assert.True(result.Succeeded);
        Assert.Equal(expected, json);
    }

    [Fact]
    public void Duplicate_key_error_policy_fails()
    {
        var (_, result) = Run(p =>
        {
            p.Output.Shape = OutputShapes.KeyedObject;
            p.Output.KeyFieldId = Id(p, "kategori");
            p.Output.DuplicateKeyPolicy = DuplicateKeyPolicy.Error;
        });
        Assert.False(result.Succeeded);
        Assert.Contains("Elektronik", result.FatalError);
    }

    [Fact]
    public void Single_object_first_or_selected_row()
    {
        var (first, _) = Run(p => { p.Output.Shape = OutputShapes.SingleObject; p.Fields.RemoveAll(f => f.Alias != "kod"); });
        Assert.Equal("""{"kod":"P1"}""", first);

        var (row4, _) = Run(p => { p.Output.Shape = OutputShapes.SingleObject; p.Output.SingleRowNumber = 4; p.Fields.RemoveAll(f => f.Alias != "kod"); });
        Assert.Equal("""{"kod":"P3"}""", row4);
    }

    [Fact]
    public void Array_of_values()
    {
        var (json, _) = Run(p => { p.Output.Shape = OutputShapes.ArrayOfValues; p.Output.ValueFieldId = Id(p, "fiyat"); });
        Assert.Equal("[1500.5,25,2]", json);
    }

    [Fact]
    public void Array_of_arrays_with_header()
    {
        var (json, _) = Run(p =>
        {
            p.Output.Shape = OutputShapes.ArrayOfArrays;
            p.Output.IncludeHeaderRow = true;
            p.Fields.RemoveAll(f => f.Alias is "kategori" or "fiyat");
        });
        Assert.Equal("""[["kod","ad"],["P1","Laptop"],["P2","Mouse"],["P3","Kalem"]]""", json);
    }

    [Fact]
    public void Grouped_dictionary_variant()
    {
        var (json, _) = Run(p =>
        {
            p.Output.Shape = OutputShapes.Grouped;
            p.Output.GroupFieldId = Id(p, "kategori");
            p.Fields.RemoveAll(f => f.Alias is "ad" or "fiyat");
        });
        Assert.Equal("""{"Elektronik":[{"kod":"P1"},{"kod":"P2"}],"Kırtasiye":[{"kod":"P3"}]}""", json);
    }

    [Fact]
    public void Grouped_array_variant_with_custom_children_name()
    {
        var (json, _) = Run(p =>
        {
            p.Output.Shape = OutputShapes.Grouped;
            p.Output.GroupFieldId = Id(p, "kategori");
            p.Output.GroupedVariant = GroupedVariant.ArrayOfGroups;
            p.Output.ChildrenName = "urunler";
            p.Fields.RemoveAll(f => f.Alias is "ad" or "fiyat");
        });
        Assert.Equal("""[{"kategori":"Elektronik","urunler":[{"kod":"P1"},{"kod":"P2"}]},{"kategori":"Kırtasiye","urunler":[{"kod":"P3"}]}]""", json);
    }

    [Fact]
    public void Wrapper_with_meta_and_root_fields()
    {
        var (json, _) = Run(p =>
        {
            p.Fields.RemoveAll(f => f.Alias != "kod");
            p.Output.Wrapper.Enabled = true;
            p.Output.Wrapper.DataPropertyName = "items";
            p.Output.Wrapper.RootFields.Add(new CustomField { Alias = "version", Value = "2" });
            p.Output.Wrapper.RootFields.Add(new CustomField { Alias = "title", Source = CustomFieldSource.Template, Value = "{sheet} ({count})" });
        });
        Assert.Equal("""{"meta":{"source":""" + "\"" + Path.GetFileName(_wb.Path) + "\"" +
                     ""","sheet":"Ürünler","exportedAt":"2026-11-23T10:00:00Z","count":3,"version":2,"title":"Ürünler (3)"},"items":[{"kod":"P1"},{"kod":"P2"},{"kod":"P3"}]}""", json);
    }

    [Fact]
    public void Wrapper_without_meta_puts_root_fields_at_root()
    {
        var (json, _) = Run(p =>
        {
            p.Fields.RemoveAll(f => f.Alias != "kod");
            p.Output.Shape = OutputShapes.ArrayOfValues;
            p.Output.ValueFieldId = Id(p, "kod");
            p.Output.Wrapper.Enabled = true;
            p.Output.Wrapper.IncludeMeta = false;
            p.Output.Wrapper.RootFields.Add(new CustomField { Alias = "source", Value = "ERP" });
        });
        Assert.Equal("""{"source":"ERP","data":["P1","P2","P3"]}""", json);
    }

    [Fact]
    public void Indented_output_with_custom_indent_and_bom()
    {
        using var session = _wb.Open();
        var profile = session.CreateProfile();
        profile.Fields.RemoveAll(f => f.Alias != "kod");
        profile.Format.Indented = true;
        profile.Format.IndentSize = 4;
        profile.Format.WriteBom = true;
        using var ms = new MemoryStream();
        Json.JsonExportEngine.CreateDefault().Run(ms, session, profile, new Json.ExportRunOptions { MaxRecords = 1 });
        var bytes = ms.ToArray();
        Assert.Equal([0xEF, 0xBB, 0xBF], bytes[..3]);
        var text = System.Text.Encoding.UTF8.GetString(bytes[3..]).ReplaceLineEndings("\n");
        Assert.Equal("[\n    {\n        \"kod\": \"P1\"\n    }\n]", text);
    }

    [Fact]
    public void Preview_is_limited_to_requested_records()
    {
        using var session = _wb.Open();
        var profile = session.CreateProfile();
        var (json, result) = Json.JsonExportEngine.CreateDefault().Preview(session, profile, 2);
        Assert.Equal(2, result.RecordCount);
        Assert.Equal(2, System.Text.Json.Nodes.JsonNode.Parse(json)!.AsArray().Count);
    }
}
