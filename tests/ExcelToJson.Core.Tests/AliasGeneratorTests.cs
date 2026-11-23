using ExcelToJson.Core.Mapping;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Tests;

public class AliasGeneratorTests
{
    [Theory]
    [InlineData("Ürün Kodu", NamingStyle.CamelCase, true, "urunKodu")]
    [InlineData("Ürün Kodu", NamingStyle.PascalCase, true, "UrunKodu")]
    [InlineData("Ürün Kodu", NamingStyle.SnakeCase, true, "urun_kodu")]
    [InlineData("Ürün Kodu", NamingStyle.CamelCase, false, "ürünKodu")]
    [InlineData("Ürün Kodu", NamingStyle.Original, false, "Ürün Kodu")]
    [InlineData("İL ADI", NamingStyle.CamelCase, true, "ilAdi")]
    [InlineData("Şehir/İlçe", NamingStyle.SnakeCase, true, "sehir_ilce")]
    [InlineData("productCode", NamingStyle.SnakeCase, true, "product_code")]
    [InlineData("HTTPStatus", NamingStyle.CamelCase, true, "httpStatus")]
    [InlineData("ID", NamingStyle.CamelCase, true, "id")]
    [InlineData("  Fiyat (TL)  ", NamingStyle.CamelCase, true, "fiyatTl")]
    [InlineData("Çğış Öü", NamingStyle.CamelCase, true, "cgisOu")]
    [InlineData("---", NamingStyle.CamelCase, true, "field")]
    public void Generates_aliases(string header, NamingStyle style, bool transliterate, string expected) =>
        Assert.Equal(expected, AliasGenerator.Generate(header, style, transliterate));

    [Fact]
    public void Duplicate_headers_are_made_unique()
    {
        Assert.Equal(["name", "name_2", "code", "name_3"], AliasGenerator.MakeUnique(["name", "name", "code", "name"]));
    }

    [Theory]
    [InlineData("a", true)]
    [InlineData("address.city", true)]
    [InlineData("", false)]
    [InlineData(" a", false)]
    [InlineData("a..b", false)]
    [InlineData(".a", false)]
    [InlineData("a.", false)]
    public void Validates_aliases(string alias, bool valid) =>
        Assert.Equal(valid, AliasGenerator.IsValid(alias, out _));

    [Fact]
    public void Field_factory_creates_unique_aliases_from_duplicate_headers()
    {
        var schema = new SheetSchema("S", 1, 2, 3,
        [
            new ColumnInfo(1, "A", "Ad", false, [], DataType.String),
            new ColumnInfo(2, "B", "Ad", false, [], DataType.String),
            new ColumnInfo(3, "C", "Column C", false, [], DataType.String)
        ]);
        var fields = FieldFactory.CreateColumnFields(schema, NamingStyle.CamelCase, true);
        Assert.Equal(["ad", "ad_2", "columnC"], fields.Select(f => f.Alias));
    }
}
