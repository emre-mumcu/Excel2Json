using System.Drawing;
using System.Windows.Forms;
using ExcelToJson.App.Components;
using ExcelToJson.App.Services;
using Microsoft.AspNetCore.Components.WebView;
using Microsoft.AspNetCore.Components.WebView.WindowsForms;
using Microsoft.Extensions.Logging;
using Microsoft.Web.WebView2.Core;

namespace ExcelToJson.App;

/// <summary>
/// Ana pencere: tüm alanı kaplayan bir BlazorWebView barındırır.
/// WebView2 içine bırakılan dosyaların tam yolu DOM'dan alınamadığı için, sayfadaki küçük bir betik
/// dosyaları <c>chrome.webview.postMessageWithAdditionalObjects</c> ile host'a gönderir; host burada
/// <see cref="CoreWebView2File.Path"/> üzerinden tam yolu alıp <see cref="FileDropBridge"/> ile Blazor'a iletir.
/// Klavye kısayolları da aynı köprüden gelir (WebView2 odaktayken form KeyPreview çalışmaz).
/// </summary>
public sealed class MainForm : Form
{
    private const string MessagePrefix = "x2j:";
    private readonly SettingsService _settings;
    private readonly FileDropBridge _bridge;
    private readonly ILogger<MainForm> _logger;
    private readonly BlazorWebView _webView;

    public MainForm(IServiceProvider services, SettingsService settings, DialogService dialogs, ShellService shell,
        FileDropBridge bridge, ThemeService theme, ILogger<MainForm> logger)
    {
        _settings = settings;
        _bridge = bridge;
        _logger = logger;

        Text = "Excel → JSON Dönüştürücü";
        AutoScaleMode = AutoScaleMode.Dpi;
        StartPosition = FormStartPosition.CenterScreen;

        // Boyutlar mantıksal (96 DPI) birimlerde tanımlanır ve monitör ölçeğine göre büyütülür.
        var scale = DeviceDpi / 96f;
        MinimumSize = new Size((int)(1024 * scale), (int)(700 * scale));
        var work = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1920, 1080);
        Size = new Size(Math.Min((int)(1440 * scale), (int)(work.Width * 0.92)), Math.Min((int)(900 * scale), (int)(work.Height * 0.92)));
        if (Size.Width < MinimumSize.Width || Size.Height < MinimumSize.Height)
            WindowState = FormWindowState.Maximized;

        Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        BackColor = theme.IsDark ? Color.FromArgb(31, 31, 31) : Color.FromArgb(250, 250, 250);

        // Form seviyesinde sürükle-bırak (WebView2 alanı dışında kalan durumlar için).
        AllowDrop = true;
        DragEnter += (_, e) => e.Effect = e.Data?.GetDataPresent(DataFormats.FileDrop) == true ? DragDropEffects.Copy : DragDropEffects.None;
        DragDrop += (_, e) =>
        {
            if (e.Data?.GetData(DataFormats.FileDrop) is string[] files)
                _bridge.RaiseFilesDropped(files);
        };

        dialogs.Attach(this);
        shell.Attach(this);

        _webView = new BlazorWebView
        {
            Dock = DockStyle.Fill,
            HostPage = Path.Combine("wwwroot", "index.html"),
            Services = services
        };
        _webView.RootComponents.Add<AppRoot>("#app");
        _webView.BlazorWebViewInitialized += OnWebViewInitialized;
        Controls.Add(_webView);

        RestoreWindowPlacement();
    }

    private void OnWebViewInitialized(object? sender, BlazorWebViewInitializedEventArgs e)
    {
        var core = e.WebView.CoreWebView2;
        // F5 (sayfa yenileme), Ctrl+P vb. tarayıcı kısayollarını kapat; tuşlar yine de DOM'a ulaşır
        // ve uygulamanın kendi kısayolları (Ctrl+O/S/E, F5) olarak işlenir.
        core.Settings.AreBrowserAcceleratorKeysEnabled = false;
#if !DEBUG
        core.Settings.AreDevToolsEnabled = false;
#endif
        core.Settings.IsStatusBarEnabled = false;
        core.Settings.IsZoomControlEnabled = false;
        core.WebMessageReceived += OnWebMessageReceived;
        _webView.WebView.DefaultBackgroundColor = BackColor;
    }

    private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        string? message;
        try
        {
            message = e.TryGetWebMessageAsString();
        }
        catch (ArgumentException)
        {
            return;
        }

        if (message is null || !message.StartsWith(MessagePrefix, StringComparison.Ordinal))
            return; // Blazor'un kendi IPC mesajları

        var body = message[MessagePrefix.Length..];
        if (body == "drop")
        {
            var paths = new List<string>();
            if (e.AdditionalObjects is { } objects)
            {
                foreach (var obj in objects)
                {
                    if (obj is CoreWebView2File file && !string.IsNullOrEmpty(file.Path))
                        paths.Add(file.Path);
                }
            }
            _logger.LogInformation("Sürükle-bırak: {Count} dosya", paths.Count);
            _bridge.RaiseFilesDropped(paths);
        }
        else if (body.StartsWith("key:", StringComparison.Ordinal))
        {
            _bridge.RaiseShortcut(body[4..]);
        }
    }

    private void RestoreWindowPlacement()
    {
        var p = _settings.Current.Window;
        if (p is null || p.Width < MinimumSize.Width || p.Height < MinimumSize.Height)
            return;

        var bounds = new Rectangle(p.X, p.Y, p.Width, p.Height);
        // Kayıtlı konum artık hiçbir ekranda görünmüyorsa (ör. monitör çıkarıldı) varsayılanı kullan.
        if (!Screen.AllScreens.Any(s => s.WorkingArea.IntersectsWith(bounds)))
            return;

        StartPosition = FormStartPosition.Manual;
        Bounds = bounds;
        if (p.Maximized)
            WindowState = FormWindowState.Maximized;
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        var bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
        _settings.Update(s => s.Window = new WindowPlacement
        {
            X = bounds.X,
            Y = bounds.Y,
            Width = bounds.Width,
            Height = bounds.Height,
            Maximized = WindowState == FormWindowState.Maximized
        });
        base.OnFormClosing(e);
    }
}
