namespace ExcelToJson.Core.Models;

/// <summary>
/// Bir dönüşümün tüm ayarları: başlık ayarları, alanlar, çıktı şekli ve export ayarları.
/// Uygulamanın çalışma durumu doğrudan bu nesnedir; profil olarak JSON'a kaydedilir.
/// </summary>
public sealed class MappingProfile
{
    public const int CurrentSchemaVersion = 1;

    public int SchemaVersion { get; set; } = CurrentSchemaVersion;
    public string Name { get; set; } = string.Empty;
    public string? SourceFileName { get; set; }
    public string? SheetName { get; set; }
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;

    public SheetReadOptions ReadOptions { get; set; } = new();

    public NamingStyle NamingStyle { get; set; } = NamingStyle.CamelCase;
    public bool TransliterateTurkish { get; set; } = true;

    public List<FieldDefinition> Fields { get; set; } = [];

    public ConversionErrorPolicy ErrorPolicy { get; set; } = ConversionErrorPolicy.Stop;

    public OutputOptions Output { get; set; } = new();

    public ExportFormatOptions Format { get; set; } = new();

    public IEnumerable<FieldDefinition> IncludedFields => Fields.Where(f => f.Include);

    public FieldDefinition? FindField(string? id) => id is null ? null : Fields.FirstOrDefault(f => f.Id == id);

    public MappingProfile Clone()
    {
        var c = (MappingProfile)MemberwiseClone();
        c.ReadOptions = ReadOptions.Clone();
        c.Fields = Fields.Select(f => f.Clone()).ToList();
        c.Output = Output.Clone();
        c.Format = Format.Clone();
        return c;
    }
}
