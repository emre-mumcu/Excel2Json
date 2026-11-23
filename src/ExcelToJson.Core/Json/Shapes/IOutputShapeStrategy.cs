using System.Text.Json;
using ExcelToJson.Core.Models;
using ExcelToJson.Core.Validation;

namespace ExcelToJson.Core.Json.Shapes;

/// <summary>Bir çıktı şeklinin yazım bağlamı.</summary>
public sealed class ShapeContext(MappingProfile profile, ConversionPlan plan, IssueCollector issues, CancellationToken cancellationToken)
{
    public MappingProfile Profile { get; } = profile;
    public OutputOptions Options => Profile.Output;
    public ConversionPlan Plan { get; } = plan;
    public IssueCollector Issues { get; } = issues;
    public CancellationToken CancellationToken { get; } = cancellationToken;
    public List<string> Warnings { get; } = [];
}

/// <summary>
/// Çıktı şekli stratejisi. Yeni bir şekil eklemek için bu arayüzü uygulayan bir sınıf yazıp
/// <see cref="OutputShapeRegistry"/>'ye kaydetmek yeterlidir; mevcut kod değişmez.
/// </summary>
public interface IOutputShapeStrategy
{
    string Id { get; }
    string DisplayName { get; }
    string Description { get; }
    /// <summary>UI'da gösterilecek küçük örnek çıktı.</summary>
    string Example { get; }
    int Order { get; }

    /// <summary>Önizlemede kayıt sınırı uygulanabilir mi? (Tek nesne gibi şekiller tüm kayıtlara bakmalıdır.)</summary>
    bool RespectsPreviewLimit => true;

    IEnumerable<ValidationMessage> Validate(MappingProfile profile);

    /// <summary>Veri kısmını (tek bir JSON değeri) yazar. Kayıtlar tembel üretilir; mümkünse stream edilmelidir.</summary>
    void Write(Utf8JsonWriter writer, IEnumerable<Record> records, ShapeContext context);
}

/// <summary>Kayıtlı çıktı şekilleri.</summary>
public sealed class OutputShapeRegistry
{
    private readonly Dictionary<string, IOutputShapeStrategy> _strategies;

    public OutputShapeRegistry(IEnumerable<IOutputShapeStrategy> strategies) =>
        _strategies = strategies.ToDictionary(s => s.Id, StringComparer.OrdinalIgnoreCase);

    public IEnumerable<IOutputShapeStrategy> All => _strategies.Values.OrderBy(s => s.Order);

    public IOutputShapeStrategy Get(string id) =>
        _strategies.TryGetValue(id, out var s) ? s : throw new ShapeException($"Bilinmeyen çıktı şekli: {id}");

    public bool TryGet(string id, out IOutputShapeStrategy strategy) => _strategies.TryGetValue(id, out strategy!);

    public static OutputShapeRegistry CreateDefault() => new(
    [
        new ArrayOfObjectsShape(),
        new KeyedObjectShape(),
        new SingleObjectShape(),
        new KeyValueShape(),
        new ArrayOfValuesShape(),
        new ArrayOfArraysShape(),
        new GroupedShape()
    ]);
}
