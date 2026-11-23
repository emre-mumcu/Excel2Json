using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Extensions.Logging;

namespace ExcelToJson.App.Services;

/// <summary>Windows kabuğu işlemleri: klasör/dosya açma, panoya kopyalama.</summary>
public sealed class ShellService(ILogger<ShellService> logger)
{
    private Form? _owner;

    public void Attach(Form owner) => _owner = owner;

    public void OpenFolder(string path)
    {
        try
        {
            if (File.Exists(path))
                Process.Start("explorer.exe", $"/select,\"{path}\"");
            else if (Directory.Exists(path))
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Klasör açılamadı: {Path}", path);
        }
    }

    public void OpenFile(string path)
    {
        try
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Dosya açılamadı: {Path}", path);
            OpenFolder(path);
        }
    }

    /// <summary>Pano STA iş parçacığı gerektirir; işlem UI iş parçacığında yapılır.</summary>
    public Task CopyToClipboardAsync(string text)
    {
        if (_owner is null)
            return Task.CompletedTask;
        return _owner.InvokeAsync(() => Clipboard.SetText(string.IsNullOrEmpty(text) ? " " : text));
    }
}
