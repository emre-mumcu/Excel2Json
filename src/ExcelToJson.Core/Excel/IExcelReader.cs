using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Excel;

/// <summary>Excel dosyasını açıp bir oturum (cache'lenmiş çalışma kitabı) döndürür.</summary>
public interface IExcelReader
{
    /// <summary>Desteklenen uzantılar (.xlsx, .xlsm).</summary>
    IReadOnlyList<string> SupportedExtensions { get; }

    Task<IWorkbookSession> OpenAsync(string path, CancellationToken cancellationToken = default);
}

/// <summary>
/// Bir kez okunmuş ve bellekte tutulan çalışma kitabı. Sayfa değiştirmek dosyayı yeniden açmaz.
/// Tüm üyeler thread-safe'tir; arka plan iş parçacıklarından çağrılabilir.
/// </summary>
public interface IWorkbookSession : IDisposable
{
    WorkbookInfo Info { get; }

    SheetInfo GetSheet(string sheetName);

    /// <summary>İlk <paramref name="maxRows"/> satırın biçimlendirilmiş ham görünümü.</summary>
    RawPreview GetRawPreview(string sheetName, int maxRows = 20);

    /// <summary>Okuma ayarlarına göre sütunları, başlıkları, örnek değerleri ve tahmini tipleri çıkarır.</summary>
    SheetSchema GetSchema(string sheetName, SheetReadOptions options, int sampleRowLimit = 200);

    /// <summary>Veri satırlarını tembel (lazy) olarak döndürür.</summary>
    IEnumerable<SheetRow> ReadRows(string sheetName, SheetReadOptions options, RowReadStats? stats = null,
        CancellationToken cancellationToken = default);
}

/// <summary>Kullanıcıya gösterilebilecek (Türkçe mesajlı) Excel okuma hatası.</summary>
public sealed class ExcelReadException(string message, Exception? inner = null) : Exception(message, inner);
