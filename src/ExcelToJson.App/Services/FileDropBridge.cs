namespace ExcelToJson.App.Services;

/// <summary>
/// Host (WinForms/WebView2) tarafında yakalanan olayları Blazor bileşenlerine iletir:
/// sürükle-bırakılan dosyaların tam yolları ve klavye kısayolları.
/// </summary>
public sealed class FileDropBridge
{
    public event Action<IReadOnlyList<string>>? FilesDropped;
    public event Action<string>? ShortcutPressed;

    public void RaiseFilesDropped(IReadOnlyList<string> paths)
    {
        if (paths.Count > 0)
            FilesDropped?.Invoke(paths);
    }

    public void RaiseShortcut(string shortcut) => ShortcutPressed?.Invoke(shortcut);
}
