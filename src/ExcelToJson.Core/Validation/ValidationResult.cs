namespace ExcelToJson.Core.Validation;

public enum ValidationSeverity
{
    Warning,
    Error
}

/// <summary>Mesajın ait olduğu adım (UI navigasyonunda rozet göstermek için).</summary>
public enum ValidationArea
{
    File,
    Sheet,
    Fields,
    Output,
    Export
}

public sealed record ValidationMessage(ValidationSeverity Severity, ValidationArea Area, string Message, string? FieldId = null)
{
    public static ValidationMessage Error(ValidationArea area, string message, string? fieldId = null) =>
        new(ValidationSeverity.Error, area, message, fieldId);

    public static ValidationMessage Warning(ValidationArea area, string message, string? fieldId = null) =>
        new(ValidationSeverity.Warning, area, message, fieldId);
}

public sealed class ValidationResult
{
    public ValidationResult(IEnumerable<ValidationMessage> messages) => Messages = messages.ToList();

    public static ValidationResult Empty { get; } = new([]);

    public IReadOnlyList<ValidationMessage> Messages { get; }
    public bool IsValid => Messages.All(m => m.Severity != ValidationSeverity.Error);
    public IEnumerable<ValidationMessage> Errors => Messages.Where(m => m.Severity == ValidationSeverity.Error);
    public IEnumerable<ValidationMessage> Warnings => Messages.Where(m => m.Severity == ValidationSeverity.Warning);

    public bool HasErrors(ValidationArea area) => Errors.Any(m => m.Area == area);
    public bool HasWarnings(ValidationArea area) => Warnings.Any(m => m.Area == area);
    public IEnumerable<ValidationMessage> ForField(string fieldId) => Messages.Where(m => m.FieldId == fieldId);
}
