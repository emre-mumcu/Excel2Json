using ExcelToJson.Core.Excel;
using ExcelToJson.Core.Export;
using ExcelToJson.Core.Json;
using ExcelToJson.Core.Mapping;
using ExcelToJson.Core.Models;
using ExcelToJson.Core.Profiles;
using ExcelToJson.Core.Validation;
using Microsoft.Extensions.Logging;

namespace ExcelToJson.App.Services;

public enum WizardStep
{
    File,
    Sheet,
    Fields,
    Output,
    Export
}

public sealed record BusyInfo(string Message, bool Cancellable)
{
    public double? Progress { get; set; }
    public string? Detail { get; set; }
}

public sealed record ErrorInfo(string Title, string Message, string? Details);

/// <summary>
/// Uygulamanın tek merkezi durumu. Bileşenler <see cref="Changed"/> olayına abone olur.
/// Uzun işlemler arka planda (Task.Run) yürütülür; UI hiçbir zaman bloklanmaz.
/// </summary>
public sealed class AppState : IDisposable
{
    private readonly IExcelReader _reader;
    private readonly JsonExportEngine _engine;
    private readonly JsonFileExporter _exporter;
    private readonly MultiSheetExporter _multiExporter;
    private readonly ProfileValidator _validator;
    private readonly IProfileStore _profiles;
    private readonly SettingsService _settings;
    private readonly NotificationService _notify;
    private readonly INativeDialogService _dialogs;
    private readonly ShellService _shell;
    private readonly ILogger<AppState> _logger;

    private CancellationTokenSource? _busyCts;
    private CancellationTokenSource? _previewCts;

    public AppState(IExcelReader reader, JsonExportEngine engine, JsonFileExporter exporter, MultiSheetExporter multiExporter, ProfileValidator validator,
        IProfileStore profiles, SettingsService settings, NotificationService notify, INativeDialogService dialogs,
        ShellService shell, ILogger<AppState> logger)
    {
        _reader = reader;
        _engine = engine;
        _exporter = exporter;
        _multiExporter = multiExporter;
        _validator = validator;
        _profiles = profiles;
        _settings = settings;
        _notify = notify;
        _dialogs = dialogs;
        _shell = shell;
        _logger = logger;
        Profile = CreateDefaultProfile();
        OutputFolder = settings.Current.LastOutputFolder;
        Revalidate();
    }

    public event Action? Changed;

    /// <summary>Ctrl+S ile profil adı sorulması gerektiğinde tetiklenir.</summary>
    public event Action? SaveProfileRequested;

    private void Notify() => Changed?.Invoke();

    // ---------- Durum ----------
    public WizardStep CurrentStep { get; private set; } = WizardStep.File;
    public IWorkbookSession? Session { get; private set; }
    public WorkbookInfo? Workbook => Session?.Info;
    public SheetInfo? Sheet { get; private set; }
    public SheetSchema? Schema { get; private set; }
    public RawPreview? RawPreview { get; private set; }
    public MappingProfile Profile { get; private set; }
    public string? ActiveProfileName { get; private set; }
    public ProfileSummary? SuggestedProfile { get; private set; }
    public IReadOnlyList<string> ProfileWarnings { get; private set; } = [];

    public ValidationResult Validation { get; private set; } = ValidationResult.Empty;
    public string? PreviewJson { get; private set; }
    public ExportRunResult? PreviewResult { get; private set; }
    public string? PreviewError { get; private set; }
    public bool PreviewLoading { get; private set; }

    /// <summary>Son export denemesinde oluşan dönüşüm hataları (export durdurulduysa).</summary>
    public ExportRunResult? LastExportRun { get; private set; }
    public FileExportResult? LastExport { get; private set; }

    public string? HighlightFieldId { get; private set; }
    public BusyInfo? Busy { get; private set; }
    public ErrorInfo? Error { get; private set; }
    public string? OutputFolder { get; private set; }

    /// <summary>Çoklu sayfa export seçimleri (dosya açıldığında sıfırlanır).</summary>
    public MultiSheetSelection Multi { get; private set; } = new();

    public AppSettings Settings => _settings.Current;
    public bool HasSheet => Session is not null && Schema is not null && Profile.SheetName is not null;

    public string ResolvedFileName => FileNameTemplate.Resolve(Profile.Format.FileNameTemplate,
        Workbook?.FileName ?? "dosya.xlsx", Profile.SheetName ?? "Sayfa", DateTime.Now, ActiveProfileName);

    // ---------- Navigasyon ----------
    public void GoTo(WizardStep step)
    {
        CurrentStep = step;
        Notify();
    }

    public void FocusField(string? fieldId)
    {
        HighlightFieldId = fieldId;
        if (fieldId is not null)
        {
            if (Profile.FindField(fieldId) is not null)
                CurrentStep = WizardStep.Fields;
            else if (Profile.Output.Wrapper.RootFields.Any(f => f.Id == fieldId))
                CurrentStep = WizardStep.Output;
        }
        Notify();
    }

    public void ClearHighlight() => HighlightFieldId = null;

    public void DismissError()
    {
        Error = null;
        Notify();
    }

    // ---------- Dosya ----------
    public async Task OpenFileAsync(string path)
    {
        var opened = await RunBusyAsync("Dosya okunuyor…", ct => _reader.OpenAsync(path, ct), "Dosya açılamadı");
        if (opened is null)
            return;

        Session?.Dispose();
        Session = opened;
        Sheet = null;
        Schema = null;
        RawPreview = null;
        ActiveProfileName = null;
        SuggestedProfile = null;
        ProfileWarnings = [];
        LastExport = null;
        LastExportRun = null;
        Profile = CreateDefaultProfile();
        _settings.AddRecentFile(opened.Info.FilePath);
        Multi = new MultiSheetSelection();
        _logger.LogInformation("Dosya açıldı: {Path} ({Sheets} sayfa)", path, opened.Info.Sheets.Count);

        var first = opened.Info.Sheets.FirstOrDefault(s => !s.IsHidden && !s.IsEmpty) ?? opened.Info.Sheets.FirstOrDefault();
        CurrentStep = WizardStep.Sheet;
        if (first is not null)
            await SelectSheetAsync(first.Name);
        else
            Revalidate();
        Notify();
    }

    public async Task PickAndOpenFileAsync()
    {
        var path = await _dialogs.PickExcelFileAsync();
        if (path is not null)
            await OpenFileAsync(path);
    }

    /// <summary>F5: dosyayı diskten yeniden okur; mevcut eşleştirme başlık adına göre korunur.</summary>
    public async Task RefreshAsync()
    {
        if (Workbook is null)
            return;
        var path = Workbook.FilePath;
        var opened = await RunBusyAsync("Dosya yeniden okunuyor…", ct => _reader.OpenAsync(path, ct), "Dosya yenilenemedi");
        if (opened is null)
            return;

        Session?.Dispose();
        Session = opened;
        var sheetName = Profile.SheetName;
        if (sheetName is not null && opened.Info.Sheets.Any(s => s.Name == sheetName))
        {
            var profile = Profile.Clone();
            var loaded = await RunBusyAsync("Sayfa okunuyor…", _ => Task.Run(() =>
            {
                var raw = opened.GetRawPreview(sheetName);
                var schema = opened.GetSchema(sheetName, profile.ReadOptions);
                return (raw, schema, ProfileReconciler.Reconcile(profile, schema));
            }), "Sayfa okunamadı");
            if (loaded is var (raw, schema, reconciled) && schema is not null)
            {
                Sheet = opened.GetSheet(sheetName);
                RawPreview = raw;
                Schema = schema;
                Profile = reconciled.Profile;
                ProfileWarnings = reconciled.Warnings;
            }
        }
        else
        {
            Sheet = null;
            Schema = null;
            RawPreview = null;
            CurrentStep = WizardStep.Sheet;
        }
        _notify.Info("Dosya yenilendi", Path.GetFileName(path));
        ProfileChanged();
    }

    public void CloseFile()
    {
        Session?.Dispose();
        Session = null;
        Sheet = null;
        Schema = null;
        RawPreview = null;
        Profile = CreateDefaultProfile();
        ActiveProfileName = null;
        SuggestedProfile = null;
        PreviewJson = null;
        PreviewResult = null;
        CurrentStep = WizardStep.File;
        Revalidate();
        Notify();
    }

    // ---------- Sayfa ----------
    public async Task SelectSheetAsync(string sheetName)
    {
        var session = Session;
        if (session is null)
            return;

        var profile = CreateDefaultProfile();
        profile.SheetName = sheetName;
        var result = await RunBusyAsync("Sayfa okunuyor…", _ => Task.Run(() =>
        {
            var raw = session.GetRawPreview(sheetName);
            var schema = session.GetSchema(sheetName, profile.ReadOptions);
            return (raw, schema);
        }), "Sayfa okunamadı");
        if (result.schema is null)
            return;

        Sheet = session.GetSheet(sheetName);
        RawPreview = result.raw;
        Schema = result.schema;
        profile.Fields = FieldFactory.CreateColumnFields(result.schema, profile.NamingStyle, profile.TransliterateTurkish);
        Profile = profile;
        ActiveProfileName = null;
        ProfileWarnings = [];
        ProfileChanged();

        // Aynı dosya + sayfa için kaydedilmiş profil varsa öner.
        try
        {
            SuggestedProfile = await _profiles.FindForSourceAsync(session.Info.FileName, sheetName);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Profil önerisi alınamadı");
            SuggestedProfile = null;
        }
        Notify();
    }

    /// <summary>Başlık/okuma ayarları değişti: şema yeniden çıkarılır, alanlar sütun numarasına göre senkronlanır.</summary>
    public async Task ApplyReadOptionsAsync()
    {
        var session = Session;
        var sheetName = Profile.SheetName;
        if (session is null || sheetName is null)
            return;
        var options = Profile.ReadOptions.Clone();
        var schema = await RunBusyAsync("Sayfa yeniden çözümleniyor…",
            _ => Task.Run(() => session.GetSchema(sheetName, options)), "Başlık ayarları uygulanamadı");
        if (schema is null)
            return;
        Schema = schema;
        FieldFactory.SyncWithSchema(Profile, schema);
        ProfileChanged();
    }


    // ---------- Profil ----------
    private MappingProfile CreateDefaultProfile() => new()
    {
        NamingStyle = _settings.Current.DefaultNamingStyle,
        TransliterateTurkish = _settings.Current.DefaultTransliterateTurkish,
        Output = new OutputOptions { Shape = _settings.Current.DefaultShape }
    };

    /// <summary>Profil üzerinde bir değişiklik yapıldı: doğrula, önizlemeyi (debounce ile) güncelle.</summary>
    public void ProfileChanged()
    {
        Revalidate();
        SchedulePreview();
        Notify();
    }

    public void RegenerateAliases()
    {
        FieldFactory.RegenerateAliases(Profile.Fields, Profile.NamingStyle, Profile.TransliterateTurkish);
        ProfileChanged();
    }

    public void AddCustomField(CustomFieldSource source = CustomFieldSource.Constant, int? insertAt = null)
    {
        var field = FieldFactory.CreateCustomField(Profile.Fields, source);
        if (insertAt is { } i && i >= 0 && i <= Profile.Fields.Count)
            Profile.Fields.Insert(i, field);
        else
            Profile.Fields.Add(field);
        HighlightFieldId = field.Id;
        ProfileChanged();
    }

    public void RemoveField(FieldDefinition field)
    {
        Profile.Fields.Remove(field);
        ProfileChanged();
    }

    public void MoveField(FieldDefinition field, int newIndex)
    {
        var old = Profile.Fields.IndexOf(field);
        if (old < 0)
            return;
        newIndex = Math.Clamp(newIndex, 0, Profile.Fields.Count - 1);
        if (old == newIndex)
            return;
        Profile.Fields.RemoveAt(old);
        Profile.Fields.Insert(newIndex, field);
        ProfileChanged();
    }

    public async Task<IReadOnlyList<ProfileSummary>> ListProfilesAsync()
    {
        try
        {
            return await _profiles.ListAsync();
        }
        catch (Exception ex)
        {
            ShowError("Profiller listelenemedi", ex);
            return [];
        }
    }

    public void RequestSaveProfile()
    {
        if (!HasSheet)
        {
            _notify.Warning("Kaydedilecek ayar yok", "Önce bir dosya ve sayfa seçin.");
            return;
        }
        if (ActiveProfileName is not null)
            _ = SaveProfileAsync(ActiveProfileName);
        else
            SaveProfileRequested?.Invoke();
    }

    public async Task<bool> SaveProfileAsync(string name)
    {
        try
        {
            var profile = Profile.Clone();
            profile.Name = name.Trim();
            profile.SourceFileName = Workbook?.FileName;
            await _profiles.SaveAsync(profile);
            Profile.Name = profile.Name;
            Profile.SourceFileName = profile.SourceFileName;
            ActiveProfileName = profile.Name;
            SuggestedProfile = null;
            _notify.Success("Profil kaydedildi", profile.Name);
            Notify();
            return true;
        }
        catch (Exception ex)
        {
            ShowError("Profil kaydedilemedi", ex);
            return false;
        }
    }

    public async Task LoadProfileAsync(string name)
    {
        try
        {
            var profile = await _profiles.LoadAsync(name);
            await ApplyProfileAsync(profile);
            ActiveProfileName = profile.Name;
            Notify();
        }
        catch (Exception ex)
        {
            ShowError("Profil yüklenemedi", ex);
        }
    }

    public async Task DeleteProfileAsync(string name)
    {
        try
        {
            await _profiles.DeleteAsync(name);
            if (ActiveProfileName == name)
                ActiveProfileName = null;
            Notify();
        }
        catch (Exception ex)
        {
            ShowError("Profil silinemedi", ex);
        }
    }

    public async Task ImportProfileAsync()
    {
        var path = await _dialogs.PickProfileToImportAsync();
        if (path is null)
            return;
        try
        {
            var profile = await _profiles.ImportFromFileAsync(path);
            if (string.IsNullOrWhiteSpace(profile.Name))
                profile.Name = Path.GetFileNameWithoutExtension(path).Replace(".profile", "");
            await ApplyProfileAsync(profile);
            ActiveProfileName = null;
            _notify.Success("Profil içe aktarıldı", "Kalıcı olarak saklamak için \"Profili kaydet\" kullanın.");
            Notify();
        }
        catch (Exception ex)
        {
            ShowError("Profil içe aktarılamadı", ex);
        }
    }

    public async Task ExportProfileAsync()
    {
        var name = ActiveProfileName ?? Profile.SheetName ?? "profil";
        var path = await _dialogs.PickProfileExportPathAsync(name);
        if (path is null)
            return;
        try
        {
            var profile = Profile.Clone();
            profile.Name = name;
            profile.SourceFileName = Workbook?.FileName;
            await _profiles.ExportToFileAsync(profile, path);
            _notify.Success("Profil dışa aktarıldı", Path.GetFileName(path),
                new ToastAction("Klasörü aç", () => _shell.OpenFolder(path)));
        }
        catch (Exception ex)
        {
            ShowError("Profil dışa aktarılamadı", ex);
        }
    }

    /// <summary>Profili mevcut çalışma kitabına uygular; sütunlar başlık adına göre yeniden eşleştirilir.</summary>
    public async Task ApplyProfileAsync(MappingProfile loaded)
    {
        var session = Session;
        if (session is null)
        {
            _notify.Warning("Önce bir Excel dosyası açın", "Profil, açık dosyanın sayfasına uygulanır.");
            return;
        }

        var sheetName = loaded.SheetName is not null && session.Info.Sheets.Any(s => string.Equals(s.Name, loaded.SheetName, StringComparison.OrdinalIgnoreCase))
            ? session.GetSheet(loaded.SheetName).Name
            : Profile.SheetName ?? session.Info.Sheets.FirstOrDefault()?.Name;
        if (sheetName is null)
            return;

        var result = await RunBusyAsync("Profil uygulanıyor…", _ => Task.Run(() =>
        {
            var raw = session.GetRawPreview(sheetName);
            var schema = session.GetSchema(sheetName, loaded.ReadOptions);
            return (raw, schema, reconciled: ProfileReconciler.Reconcile(loaded, schema));
        }), "Profil uygulanamadı");
        if (result.schema is null)
            return;

        Sheet = session.GetSheet(sheetName);
        RawPreview = result.raw;
        Schema = result.schema;
        Profile = result.reconciled.Profile;
        ProfileWarnings = result.reconciled.Warnings;
        SuggestedProfile = null;
        if (loaded.SheetName is not null && !string.Equals(loaded.SheetName, sheetName, StringComparison.OrdinalIgnoreCase))
            ProfileWarnings = [$"Profildeki \"{loaded.SheetName}\" sayfası bulunamadı; \"{sheetName}\" sayfasına uygulandı.", .. ProfileWarnings];
        if (ProfileWarnings.Count > 0)
            _notify.Warning("Profil uyarılarla uygulandı", string.Join("\n", ProfileWarnings.Take(3)));
        else
            _notify.Success("Profil uygulandı", loaded.Name);
        CurrentStep = WizardStep.Fields;
        ProfileChanged();
    }

    public async Task ApplySuggestedProfileAsync()
    {
        if (SuggestedProfile is { } s)
            await LoadProfileAsync(s.Name);
    }

    public void DismissSuggestedProfile()
    {
        SuggestedProfile = null;
        Notify();
    }

    // ---------- Doğrulama ve önizleme ----------
    private void Revalidate()
    {
        var messages = new List<ValidationMessage>();
        if (Session is null)
            messages.Add(ValidationMessage.Error(ValidationArea.File, "Bir Excel dosyası seçin."));
        else
            messages.AddRange(_validator.Validate(Profile, Schema).Messages);

        if (string.IsNullOrWhiteSpace(OutputFolder))
            messages.Add(ValidationMessage.Error(ValidationArea.Export, "Hedef klasör seçilmelidir."));
        else if (!Directory.Exists(OutputFolder))
            messages.Add(ValidationMessage.Error(ValidationArea.Export, "Hedef klasör bulunamadı."));

        Validation = new ValidationResult(messages);
    }

    public void SchedulePreview()
    {
        _previewCts?.Cancel();
        _previewCts?.Dispose();
        _previewCts = null;

        var session = Session;
        if (session is null || !HasSheet)
        {
            PreviewJson = null;
            PreviewResult = null;
            PreviewError = null;
            PreviewLoading = false;
            return;
        }

        var cts = new CancellationTokenSource();
        _previewCts = cts;
        PreviewLoading = true;
        _ = RunPreviewAsync(session, Profile.Clone(), Math.Clamp(_settings.Current.PreviewRecordCount, 1, 500), cts.Token);
    }

    private async Task RunPreviewAsync(IWorkbookSession session, MappingProfile snapshot, int count, CancellationToken ct)
    {
        try
        {
            await Task.Delay(300, ct); // debounce
            if (!snapshot.IncludedFields.Any())
            {
                PreviewJson = null;
                PreviewResult = null;
                PreviewError = "Önizleme için en az bir alan seçin.";
            }
            else
            {
                var (json, result) = await Task.Run(() => _engine.Preview(session, snapshot, count, ct), ct);
                ct.ThrowIfCancellationRequested();
                PreviewJson = json;
                PreviewResult = result;
                PreviewError = result.FatalError;
            }
            PreviewLoading = false;
            Notify();
        }
        catch (OperationCanceledException)
        {
            // Yeni bir önizleme isteği geldi.
        }
        catch (ObjectDisposedException)
        {
            // Dosya bu sırada kapatıldı.
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Önizleme üretilemedi");
            PreviewJson = null;
            PreviewError = "Önizleme üretilemedi: " + ex.Message;
            PreviewLoading = false;
            Notify();
        }
    }

    public void SetPreviewCount(int count)
    {
        _settings.Update(s => s.PreviewRecordCount = Math.Clamp(count, 1, 500));
        SchedulePreview();
        Notify();
    }

    // ---------- Export ----------
    public async Task PickOutputFolderAsync()
    {
        var folder = await _dialogs.PickFolderAsync(OutputFolder ?? (Workbook is null ? null : Path.GetDirectoryName(Workbook.FilePath)));
        if (folder is null)
            return;
        OutputFolder = folder;
        _settings.Update(s => s.LastOutputFolder = folder);
        ProfileChanged();
    }

    public async Task ExportAsync()
    {
        Revalidate();
        if (Busy is not null)
            return;
        if (!Validation.IsValid)
        {
            var first = Validation.Errors.First();
            if (first.Area == ValidationArea.Export && string.IsNullOrWhiteSpace(OutputFolder))
            {
                await PickOutputFolderAsync();
                Revalidate();
                if (!Validation.IsValid)
                {
                    if (!string.IsNullOrWhiteSpace(OutputFolder))
                        ShowValidationErrors();
                    return;
                }
            }
            else
            {
                ShowValidationErrors();
                return;
            }
        }

        var session = Session!;
        var target = Path.Combine(OutputFolder!, ResolvedFileName);
        if (File.Exists(target))
        {
            var policy = Profile.Format.ExistingFilePolicy;
            if (policy == ExistingFilePolicy.Ask)
            {
                var choice = await _dialogs.AskOverwriteAsync(target);
                if (choice == OverwriteChoice.Cancel)
                    return;
                policy = choice == OverwriteChoice.Overwrite ? ExistingFilePolicy.Overwrite : ExistingFilePolicy.AppendNumber;
            }
            if (policy == ExistingFilePolicy.AppendNumber)
                target = FileNameTemplate.GetAvailablePath(target);
        }

        var snapshot = Profile.Clone();
        _logger.LogInformation("Export başladı: {Sheet} → {Target}", snapshot.SheetName, target);

        var result = await RunBusyAsync("JSON dosyası oluşturuluyor…", async ct =>
        {
            var busy = Busy;
            var progress = new Progress<ExportProgress>(p =>
            {
                if (busy is null) return;
                busy.Progress = p.Fraction;
                busy.Detail = $"{p.Processed:N0} / {p.Total:N0} kayıt";
                Notify();
            });
            return await _exporter.ExportAsync(session, snapshot, target, progress, ct);
        }, "Export başarısız oldu", cancellable: true);

        if (result is null)
            return;

        LastExportRun = result.Run;
        if (result.Succeeded)
        {
            LastExport = result;
            _logger.LogInformation("Export tamamlandı: {Count} kayıt, {Size} bayt, {Duration}", result.Run.RecordCount, result.FileSize, result.Duration);
            var path = result.FilePath!;
            var msg = $"{result.Run.RecordCount:N0} kayıt · {FormatSize(result.FileSize)} · {result.Duration.TotalSeconds:0.0} sn";
            if (result.Run.IssueCount > 0)
                msg += $" · {result.Run.IssueCount} dönüşüm uyarısı";
            _notify.Success($"Export tamamlandı: {Path.GetFileName(path)}", msg,
                new ToastAction("Klasörü aç", () => _shell.OpenFolder(path)),
                new ToastAction("Dosyayı aç", () => _shell.OpenFile(path)));
        }
        else if (result.Run.Aborted)
        {
            _notify.Error("Export durduruldu",
                $"{result.Run.IssueCount} dönüşüm hatası bulundu. Hatalar önizleme panelinde listelendi; politika \"Export'u durdur\" olduğu için dosya yazılmadı.");
            PreviewResult = result.Run;
        }
        else
        {
            _notify.Error("Export başarısız", result.Run.FatalError);
        }
        Notify();
    }

    private void ShowValidationErrors()
    {
        var errors = Validation.Errors.ToList();
        var first = errors[0];
        _notify.Error($"Export öncesi {errors.Count} sorun var", string.Join("\n", errors.Take(3).Select(e => "• " + e.Message)));
        CurrentStep = first.Area switch
        {
            ValidationArea.File => WizardStep.File,
            ValidationArea.Sheet => WizardStep.Sheet,
            ValidationArea.Fields => WizardStep.Fields,
            ValidationArea.Output => WizardStep.Output,
            _ => WizardStep.Export
        };
        HighlightFieldId = first.FieldId;
        Notify();
    }

    // ---------- Çoklu sayfa export ----------

    /// <summary>
    /// Çoklu sayfa listesini günceller: yeni sayfalar eklenir; varsayılan seçim açık sayfa için "mevcut ayarlar",
    /// diğerleri için aynı dosya+sayfaya kayıtlı en son profil, yoksa "varsayılan (tüm sütunlar)".
    /// </summary>
    public async Task<IReadOnlyList<ProfileSummary>> RefreshMultiSheetAsync()
    {
        var profiles = await ListProfilesAsync();
        if (Workbook is not { } wb)
            return profiles;

        foreach (var sheet in wb.Sheets.Where(s => !s.IsEmpty))
        {
            if (!Multi.Items.TryGetValue(sheet.Name, out var item))
            {
                var saved = profiles.FirstOrDefault(p =>
                    string.Equals(p.SourceFileName, wb.FileName, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(p.SheetName, sheet.Name, StringComparison.OrdinalIgnoreCase));
                item = new MultiSheetItem
                {
                    Selected = sheet.Name == Profile.SheetName,
                    Choice = saved is not null ? MultiSheetItem.ProfilePrefix + saved.Name : MultiSheetItem.DefaultChoice
                };
                Multi.Items[sheet.Name] = item;
            }

            // "Mevcut ayarlar" yalnızca açık sayfa için geçerlidir.
            if (sheet.Name == Profile.SheetName && item.Choice == MultiSheetItem.DefaultChoice)
                item.Choice = MultiSheetItem.CurrentChoice;
            else if (sheet.Name != Profile.SheetName && item.Choice == MultiSheetItem.CurrentChoice)
                item.Choice = MultiSheetItem.DefaultChoice;
            if (item.ProfileName is { } name && !profiles.Any(p => p.Name == name))
                item.Choice = sheet.Name == Profile.SheetName ? MultiSheetItem.CurrentChoice : MultiSheetItem.DefaultChoice;
        }
        Notify();
        return profiles;
    }

    public async Task ExportMultiAsync()
    {
        if (Busy is not null || Session is not { } session)
            return;

        var selected = session.Info.Sheets
            .Where(s => Multi.Items.TryGetValue(s.Name, out var i) && i.Selected)
            .Select(s => (Sheet: s.Name, Item: Multi.Items[s.Name]))
            .ToList();
        if (selected.Count == 0)
        {
            _notify.Warning("Sayfa seçilmedi", "Export için en az bir sayfa seçin.");
            return;
        }

        if (string.IsNullOrWhiteSpace(OutputFolder) || !Directory.Exists(OutputFolder))
        {
            await PickOutputFolderAsync();
            if (string.IsNullOrWhiteSpace(OutputFolder) || !Directory.Exists(OutputFolder))
                return;
        }

        // 1) Her sayfa için profili hazırla (kayıtlı profiller başlık adına göre yeniden eşleştirilir) ve doğrula.
        var current = Profile.Clone();
        var defaults = CreateDefaultProfile();
        defaults.Format = Profile.Format.Clone();
        var jobs = await RunBusyAsync("Sayfalar hazırlanıyor…", async ct =>
        {
            var list = new List<SheetExportJob>();
            foreach (var (sheet, item) in selected)
            {
                ct.ThrowIfCancellationRequested();
                if (item.Choice == MultiSheetItem.CurrentChoice && sheet == current.SheetName)
                {
                    list.Add(new SheetExportJob(sheet, current.Clone(), []));
                    continue;
                }
                var saved = item.ProfileName is { } name ? await _profiles.LoadAsync(name, ct) : null;
                list.Add(await Task.Run(() => MultiSheetExporter.PrepareJob(session, sheet, saved, defaults), ct));
            }
            return list;
        }, "Sayfalar hazırlanamadı");
        if (jobs is null)
            return;

        var errors = jobs
            .SelectMany(j => _validator.Validate(j.Profile).Errors.Select(e => $"{j.SheetName}: {e.Message}"))
            .ToList();
        if (errors.Count > 0)
        {
            Error = new ErrorInfo("Çoklu export başlatılamadı",
                $"{errors.Count} doğrulama hatası var. Hatalı sayfanın profilini düzeltin veya o sayfa için \"Varsayılan\" seçin.",
                string.Join(Environment.NewLine, errors));
            Notify();
            return;
        }

        // 2) Hedef yolları çöz ve mevcut dosyaları politikaya göre ele al.
        var now = DateTime.Now;
        var folder = OutputFolder!;
        var targets = Multi.Mode == MultiSheetMode.SingleFile
            ? [Path.Combine(folder, FileNameTemplate.Resolve(Multi.CombinedFileNameTemplate, session.Info.FileName, "sayfalar", now, ActiveProfileName))]
            : MultiSheetExporter.ResolveSeparatePaths(folder, session.Info.FileName, jobs, now);

        var conflicts = targets.Where(File.Exists).ToList();
        if (conflicts.Count > 0)
        {
            var policy = Profile.Format.ExistingFilePolicy;
            if (policy == ExistingFilePolicy.Ask)
            {
                var choice = await _dialogs.AskOverwriteAsync(conflicts);
                if (choice == OverwriteChoice.Cancel)
                    return;
                policy = choice == OverwriteChoice.Overwrite ? ExistingFilePolicy.Overwrite : ExistingFilePolicy.AppendNumber;
            }
            if (policy == ExistingFilePolicy.AppendNumber)
                targets = targets.Select(FileNameTemplate.GetAvailablePath).ToList();
        }

        // 3) Export
        _logger.LogInformation("Çoklu export başladı: {Count} sayfa, mod {Mode}", jobs.Count, Multi.Mode);
        var mode = Multi.Mode;
        var format = Profile.Format.Clone();
        var result = await RunBusyAsync($"{jobs.Count} sayfa export ediliyor…", async ct =>
        {
            var busy = Busy;
            var progress = new Progress<MultiSheetProgress>(p =>
            {
                if (busy is null) return;
                busy.Progress = p.Fraction;
                busy.Detail = $"Sayfa {p.SheetIndex + 1}/{p.SheetCount} · {p.SheetName} · {p.Sheet.Processed:N0} / {p.Sheet.Total:N0} kayıt";
                Notify();
            });
            return mode == MultiSheetMode.SingleFile
                ? await _multiExporter.ExportCombinedAsync(session, jobs, targets[0], format, progress, ct)
                : await _multiExporter.ExportSeparateAsync(session, jobs, targets, progress, ct);
        }, "Çoklu export başarısız oldu");
        if (result is null)
            return;

        Multi.LastResult = result;
        var warnings = jobs.SelectMany(j => j.Warnings.Select(w => $"{j.SheetName}: {w}")).ToList();
        _logger.LogInformation("Çoklu export bitti: {Ok}/{Count} sayfa, {Records} kayıt", result.SucceededCount, result.Sheets.Count, result.TotalRecords);

        var files = result.WrittenFiles.ToList();
        var actions = new List<ToastAction>();
        if (files.Count > 0)
            actions.Add(new ToastAction("Klasörü aç", () => _shell.OpenFolder(files[0])));
        if (files.Count == 1)
            actions.Add(new ToastAction("Dosyayı aç", () => _shell.OpenFile(files[0])));

        if (result.Succeeded)
        {
            var msg = $"{result.TotalRecords:N0} kayıt · {files.Count} dosya · {result.Duration.TotalSeconds:0.0} sn";
            if (warnings.Count > 0)
                msg += $" · {warnings.Count} profil uyarısı";
            _notify.Success($"{result.Sheets.Count} sayfa export edildi", msg, [.. actions]);
        }
        else
        {
            var failed = result.Sheets.Where(s => !s.Succeeded)
                .Select(s => $"• {s.SheetName}: " + (s.Run.FatalError ?? $"{s.Run.IssueCount} dönüşüm hatası (export durduruldu)"));
            var title = mode == MultiSheetMode.SingleFile
                ? "Çoklu export durduruldu – dosya yazılmadı"
                : $"{result.SucceededCount}/{result.Sheets.Count} sayfa export edildi";
            _notify.Show(result.SucceededCount > 0 && mode == MultiSheetMode.SeparateFiles ? ToastKind.Warning : ToastKind.Error,
                title, string.Join("\n", failed.Take(4)), [.. actions]);
        }
        Notify();
    }

    public static string FormatSize(long bytes) => bytes switch
    {
        < 1024 => $"{bytes} B",
        < 1024 * 1024 => $"{bytes / 1024.0:0.#} KB",
        _ => $"{bytes / 1024.0 / 1024.0:0.#} MB"
    };

    // ---------- Meşguliyet ve hata yönetimi ----------
    public void CancelBusy() => _busyCts?.Cancel();

    private async Task<T?> RunBusyAsync<T>(string message, Func<CancellationToken, Task<T>> work, string errorTitle, bool cancellable = true)
    {
        _busyCts?.Dispose();
        _busyCts = new CancellationTokenSource();
        Busy = new BusyInfo(message, cancellable);
        Notify();
        var token = _busyCts.Token;
        Task<T>? task = null;
        try
        {
            task = work(token);
            // İş iptal belirtecini dinlemese bile (ör. ClosedXML çağrıları) arayüz iptalde hemen serbest kalır.
            return await task.WaitAsync(token);
        }
        catch (OperationCanceledException)
        {
            // Arka planda süren iş sonradan tamamlanırsa ürettiği kaynak (ör. çalışma kitabı oturumu) kapatılır.
            if (task is not null && !task.IsCompleted)
            {
                _ = task.ContinueWith(t =>
                {
                    if (t.IsCompletedSuccessfully && t.Result is IDisposable d)
                        d.Dispose();
                    else
                        _ = t.Exception; // gözlemlenmemiş hata bırakma
                }, TaskScheduler.Default);
            }
            _notify.Info("İşlem iptal edildi");
            return default;
        }
        catch (Exception ex)
        {
            ShowError(errorTitle, ex);
            return default;
        }
        finally
        {
            Busy = null;
            Notify();
        }
    }

    /// <summary>Hatayı loglar ve kullanıcıya teknik olmayan Türkçe bir mesaj + "Detaylar" gösterir.</summary>
    public void ShowError(string title, Exception ex)
    {
        var userMessage = ex switch
        {
            ExcelReadException or ProfileFormatException => ex.Message,
            FileNotFoundException => "Dosya bulunamadı.",
            UnauthorizedAccessException => "Bu konuma erişim izniniz yok. Farklı bir klasör seçmeyi deneyin.",
            IOException => "Dosya işlemi başarısız oldu. Dosya başka bir uygulama tarafından kullanılıyor olabilir.",
            _ => "Beklenmeyen bir hata oluştu. Ayrıntılar günlük dosyasına kaydedildi."
        };
        if (ex is ExcelReadException or ProfileFormatException)
            _logger.LogWarning(ex, "{Title}", title);
        else
            _logger.LogError(ex, "{Title}", title);
        Error = new ErrorInfo(title, userMessage, ex.ToString());
        Notify();
    }

    public void Dispose()
    {
        _previewCts?.Cancel();
        _previewCts?.Dispose();
        _busyCts?.Dispose();
        Session?.Dispose();
    }
}
