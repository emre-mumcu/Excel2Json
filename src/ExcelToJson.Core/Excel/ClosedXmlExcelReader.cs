using ClosedXML.Excel;
using ExcelToJson.Core.Models;

namespace ExcelToJson.Core.Excel;

/// <summary>ClosedXML tabanlı Excel okuyucu.</summary>
public sealed class ClosedXmlExcelReader : IExcelReader
{
    public IReadOnlyList<string> SupportedExtensions { get; } = [".xlsx", ".xlsm"];

    public Task<IWorkbookSession> OpenAsync(string path, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        return Task.Run<IWorkbookSession>(() => Open(path, cancellationToken), cancellationToken);
    }

    private IWorkbookSession Open(string path, CancellationToken ct)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext == ".xls")
            throw new ExcelReadException("Eski .xls biçimi desteklenmiyor. Lütfen dosyayı Excel'de açıp \".xlsx\" olarak kaydedin.");
        if (!SupportedExtensions.Contains(ext))
            throw new ExcelReadException($"Desteklenmeyen dosya türü: \"{ext}\". Yalnızca .xlsx ve .xlsm dosyaları açılabilir.");
        if (!File.Exists(path))
            throw new ExcelReadException($"Dosya bulunamadı: {path}");

        // Dosya Excel'de açık (kilitli) olsa bile okunabilsin diye FileShare.ReadWrite ile açılır
        // ve belleğe kopyalanır; böylece dosya oturum boyunca kilitli tutulmaz.
        var buffer = new MemoryStream();
        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            fs.CopyTo(buffer);
        }
        catch (IOException ex)
        {
            throw new ExcelReadException("Dosya okunamadı. Başka bir uygulama dosyayı özel olarak kilitlemiş olabilir.", ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new ExcelReadException("Dosyaya erişim izni yok.", ex);
        }

        ct.ThrowIfCancellationRequested();
        buffer.Position = 0;

        XLWorkbook workbook;
        try
        {
            workbook = new XLWorkbook(buffer);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            buffer.Dispose();
            throw new ExcelReadException("Dosya geçerli bir Excel çalışma kitabı olarak açılamadı. Dosya bozuk veya parola korumalı olabilir.", ex);
        }

        return new ClosedXmlWorkbookSession(path, buffer.Length, workbook, buffer);
    }
}
