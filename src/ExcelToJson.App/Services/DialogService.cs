using System.Windows.Forms;

namespace ExcelToJson.App.Services;

public enum OverwriteChoice
{
    Overwrite,
    AppendNumber,
    Cancel
}

/// <summary>Native Windows diyalogları. Razor tarafından await edilebilir.</summary>
public interface INativeDialogService
{
    Task<string?> PickExcelFileAsync();
    Task<string?> PickFolderAsync(string? initialFolder);
    Task<string?> PickProfileToImportAsync();
    Task<string?> PickProfileExportPathAsync(string suggestedName);
    Task<OverwriteChoice> AskOverwriteAsync(string path);
    Task<OverwriteChoice> AskOverwriteAsync(IReadOnlyList<string> paths);
    Task<bool> ConfirmAsync(string title, string message);
}

/// <summary>
/// Diyaloglar her zaman WinForms UI iş parçacığında, ana formun sahipliğinde açılır.
/// Blazor tarafındaki çağrı, diyalog kapanana kadar await edilir.
/// </summary>
public sealed class DialogService : INativeDialogService
{
    private Form? _owner;

    public void Attach(Form owner) => _owner = owner;

    private Task<T> OnUiThread<T>(Func<Form, T> action)
    {
        var owner = _owner ?? throw new InvalidOperationException("DialogService bir forma bağlanmadı.");
        // Control.InvokeAsync çağrıyı UI iş parçacığının mesaj kuyruğuna koyar ve sonucu Task olarak döndürür.
        return owner.InvokeAsync(() => action(owner));
    }

    public Task<string?> PickExcelFileAsync() => OnUiThread(owner =>
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Excel dosyası seçin",
            Filter = "Excel çalışma kitapları (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Tüm dosyalar (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false
        };
        return dialog.ShowDialog(owner) == DialogResult.OK ? dialog.FileName : null;
    });

    public Task<string?> PickFolderAsync(string? initialFolder) => OnUiThread(owner =>
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = "JSON dosyasının kaydedileceği klasörü seçin",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true
        };
        if (!string.IsNullOrEmpty(initialFolder) && Directory.Exists(initialFolder))
            dialog.InitialDirectory = initialFolder;
        return dialog.ShowDialog(owner) == DialogResult.OK ? dialog.SelectedPath : null;
    });

    public Task<string?> PickProfileToImportAsync() => OnUiThread(owner =>
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Profil içe aktar",
            Filter = "Profil dosyaları (*.json)|*.json|Tüm dosyalar (*.*)|*.*",
            CheckFileExists = true
        };
        return dialog.ShowDialog(owner) == DialogResult.OK ? dialog.FileName : null;
    });

    public Task<string?> PickProfileExportPathAsync(string suggestedName) => OnUiThread(owner =>
    {
        using var dialog = new SaveFileDialog
        {
            Title = "Profili dışa aktar",
            Filter = "Profil dosyaları (*.profile.json)|*.profile.json|JSON (*.json)|*.json",
            FileName = suggestedName + ".profile.json",
            OverwritePrompt = true
        };
        return dialog.ShowDialog(owner) == DialogResult.OK ? dialog.FileName : null;
    });

    public Task<OverwriteChoice> AskOverwriteAsync(string path) => OnUiThread(owner =>
    {
        var result = MessageBox.Show(owner,
            $"\"{Path.GetFileName(path)}\" dosyası zaten var.\n\nÜzerine yazılsın mı?\n\nEvet: üzerine yaz\nHayır: numara ekleyerek yeni dosya oluştur\nİptal: export'u iptal et",
            "Dosya zaten var", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        return result switch
        {
            DialogResult.Yes => OverwriteChoice.Overwrite,
            DialogResult.No => OverwriteChoice.AppendNumber,
            _ => OverwriteChoice.Cancel
        };
    });

    public Task<OverwriteChoice> AskOverwriteAsync(IReadOnlyList<string> paths)
    {
        if (paths.Count == 1)
            return AskOverwriteAsync(paths[0]);
        var list = string.Join("\n", paths.Take(5).Select(p => "• " + Path.GetFileName(p)));
        if (paths.Count > 5)
            list += $"\n• … ve {paths.Count - 5} dosya daha";
        return OnUiThread(owner =>
        {
            var result = MessageBox.Show(owner,
                $"{paths.Count} dosya zaten var:\n{list}\n\nHepsinin üzerine yazılsın mı?\n\nEvet: üzerine yaz\nHayır: numara ekleyerek yeni dosyalar oluştur\nİptal: export'u iptal et",
                "Dosyalar zaten var", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            return result switch
            {
                DialogResult.Yes => OverwriteChoice.Overwrite,
                DialogResult.No => OverwriteChoice.AppendNumber,
                _ => OverwriteChoice.Cancel
            };
        });
    }

    public Task<bool> ConfirmAsync(string title, string message) => OnUiThread(owner =>
        MessageBox.Show(owner, message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes);
}
