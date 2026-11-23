namespace ExcelToJson.Core.Models;

/// <summary>JSON alanının hedef veri tipi.</summary>
public enum DataType
{
    Auto,
    String,
    Integer,
    Decimal,
    Boolean,
    Date,
    DateTime,
    Array
}

/// <summary>Başlıktan alias üretirken kullanılan adlandırma stili.</summary>
public enum NamingStyle
{
    CamelCase,
    PascalCase,
    SnakeCase,
    Original
}

/// <summary>Hücre boş olduğunda ne yazılacağı.</summary>
public enum EmptyValuePolicy
{
    Null,
    EmptyString,
    Omit,
    DefaultValue
}

/// <summary>Tip dönüşümü başarısız olduğunda uygulanacak politika.</summary>
public enum ConversionErrorPolicy
{
    Stop,
    SetNull,
    KeepRaw
}

/// <summary>Ek alanın değer kaynağı.</summary>
public enum CustomFieldSource
{
    Constant,
    Template,
    RowNumber,
    AutoIncrement,
    Guid,
    Timestamp,
    Json
}

public enum RowNumberMode
{
    /// <summary>Excel'deki gerçek satır numarası.</summary>
    ExcelRow,
    /// <summary>1'den başlayan kayıt sıra numarası.</summary>
    Sequence
}

public enum TimestampMode
{
    Local,
    Utc
}

public enum DuplicateKeyPolicy
{
    Error,
    LastWins,
    FirstWins,
    ToArray
}

public enum GroupedVariant
{
    /// <summary>{ "grup": [ ... ] }</summary>
    Dictionary,
    /// <summary>[ { "grupAlanı": "..", "items": [ ... ] } ]</summary>
    ArrayOfGroups
}

public enum ExistingFilePolicy
{
    Overwrite,
    AppendNumber,
    Ask
}

public enum ThemeMode
{
    System,
    Light,
    Dark
}
