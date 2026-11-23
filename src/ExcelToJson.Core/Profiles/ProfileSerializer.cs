using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Profiles;

public sealed class ProfileFormatException(string message, Exception? inner = null) : Exception(message, inner);

/// <summary>Profil JSON serileştirmesi ve şema sürümü migration'ı.</summary>
public static class ProfileSerializer
{
    public static JsonSerializerOptions Options { get; } = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) },
        AllowTrailingCommas = true,
        ReadCommentHandling = JsonCommentHandling.Skip
    };

    public static string Serialize(MappingProfile profile) => JsonSerializer.Serialize(profile, Options);

    public static MappingProfile Deserialize(string json)
    {
        JsonNode? node;
        try
        {
            node = JsonNode.Parse(json, documentOptions: new JsonDocumentOptions { AllowTrailingCommas = true, CommentHandling = JsonCommentHandling.Skip });
        }
        catch (JsonException ex)
        {
            throw new ProfileFormatException("Profil dosyası geçerli bir JSON değil.", ex);
        }

        if (node is not JsonObject obj)
            throw new ProfileFormatException("Profil dosyası bir JSON nesnesi olmalıdır.");

        var version = obj["schemaVersion"]?.GetValue<int>() ?? 0;
        if (version > MappingProfile.CurrentSchemaVersion)
            throw new ProfileFormatException($"Profil daha yeni bir uygulama sürümüyle oluşturulmuş (şema sürümü {version}). Lütfen uygulamayı güncelleyin.");

        Migrate(obj, version);

        try
        {
            return obj.Deserialize<MappingProfile>(Options) ?? throw new ProfileFormatException("Profil okunamadı.");
        }
        catch (JsonException ex)
        {
            throw new ProfileFormatException("Profil dosyası okunamadı: " + ex.Message, ex);
        }
    }

    /// <summary>Eski şema sürümlerini güncel sürüme yükseltir. Yeni sürümlerde adımlar buraya eklenir.</summary>
    private static void Migrate(JsonObject obj, int fromVersion)
    {
        if (fromVersion < 1)
        {
            // Sürüm 0 (sürüm alanı olmayan) profiller aynı yapıya sahip kabul edilir.
        }
        obj["schemaVersion"] = MappingProfile.CurrentSchemaVersion;
    }
}
