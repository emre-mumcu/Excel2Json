using System.Diagnostics;
using System.Windows.Forms;
using ExcelToJson.App.Services;
using ExcelToJson.Core.Excel;
using ExcelToJson.Core.Export;
using ExcelToJson.Core.Json;
using ExcelToJson.Core.Json.Shapes;
using ExcelToJson.Core.Profiles;
using ExcelToJson.Core.Validation;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.FluentUI.AspNetCore.Components;
using Microsoft.Web.WebView2.Core;
using DialogResult = System.Windows.Forms.DialogResult;
using DialogService = ExcelToJson.App.Services.DialogService;
using MessageBox = System.Windows.Forms.MessageBox;
using Serilog;

namespace ExcelToJson.App;

internal static class Program
{
    private const string WebView2DownloadUrl = "https://developer.microsoft.com/microsoft-edge/webview2/";

    [STAThread]
    private static void Main()
    {
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.File(Path.Combine(AppPaths.LogDirectory, "log-.txt"),
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 14,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {SourceContext}: {Message:lj}{NewLine}{Exception}")
            .CreateLogger();

        try
        {
            ApplicationConfiguration.Initialize();
            RegisterGlobalExceptionHandlers();

            if (!IsWebView2Available())
                return;

            using var services = ConfigureServices();
            var form = services.GetRequiredService<MainForm>();
            Log.Information("Uygulama başlatıldı. Sürüm {Version}", Application.ProductVersion);
            Application.Run(form);
        }
        catch (Exception ex)
        {
            Log.Fatal(ex, "Uygulama beklenmeyen bir hatayla sonlandı");
            MessageBox.Show(
                "Uygulama beklenmeyen bir hatayla karşılaştı ve kapanacak.\n\nAyrıntılar günlük dosyasına yazıldı:\n" + AppPaths.LogDirectory,
                "Excel → JSON", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            Log.CloseAndFlush();
        }
    }

    private static ServiceProvider ConfigureServices()
    {
        var services = new ServiceCollection();

        services.AddWindowsFormsBlazorWebView();
#if DEBUG
        services.AddBlazorWebViewDeveloperTools();
#endif
        services.AddFluentUIComponents();
        services.AddLogging(b => b.ClearProviders().AddSerilog(dispose: false));

        // Core
        services.AddSingleton<IExcelReader, ClosedXmlExcelReader>();
        services.AddSingleton<Core.Conversion.IValueConverter, Core.Conversion.ValueConverter>();
        services.AddSingleton(_ => OutputShapeRegistry.CreateDefault());
        services.AddSingleton<JsonExportEngine>();
        services.AddSingleton<JsonFileExporter>();
        services.AddSingleton<MultiSheetExporter>();
        services.AddSingleton<ProfileValidator>();
        services.AddSingleton<IProfileStore>(_ => new FileProfileStore(AppPaths.ProfileDirectory));

        // App
        services.AddSingleton<SettingsService>();
        services.AddSingleton<ThemeService>();
        services.AddSingleton<DialogService>();
        services.AddSingleton<INativeDialogService>(sp => sp.GetRequiredService<DialogService>());
        services.AddSingleton<ShellService>();
        services.AddSingleton<FileDropBridge>();
        services.AddSingleton<NotificationService>();
        services.AddSingleton<AppState>();
        services.AddSingleton<MainForm>();

        return services.BuildServiceProvider();
    }

    private static void RegisterGlobalExceptionHandlers()
    {
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (_, e) =>
        {
            Log.Error(e.Exception, "UI iş parçacığında işlenmeyen hata");
            MessageBox.Show("Beklenmeyen bir hata oluştu. Ayrıntılar günlük dosyasına yazıldı.\n\n" + e.Exception.Message,
                "Excel → JSON", MessageBoxButtons.OK, MessageBoxIcon.Error);
        };
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
            Log.Fatal(e.ExceptionObject as Exception, "İşlenmeyen hata (AppDomain)");
        TaskScheduler.UnobservedTaskException += (_, e) =>
        {
            Log.Error(e.Exception, "Gözlenmeyen görev hatası");
            e.SetObserved();
        };
    }

    /// <summary>WebView2 Runtime kurulu değilse kullanıcıya anlaşılır bir mesaj ve indirme bağlantısı gösterir.</summary>
    private static bool IsWebView2Available()
    {
        try
        {
            var version = CoreWebView2Environment.GetAvailableBrowserVersionString();
            Log.Information("WebView2 Runtime sürümü: {Version}", version);
            return true;
        }
        catch (WebView2RuntimeNotFoundException ex)
        {
            Log.Warning(ex, "WebView2 Runtime bulunamadı");
            var answer = MessageBox.Show(
                "Bu uygulamanın çalışması için Microsoft Edge WebView2 Runtime gereklidir, ancak bilgisayarınızda kurulu değil.\n\n" +
                "İndirme sayfasını şimdi açmak ister misiniz?\n" + WebView2DownloadUrl,
                "WebView2 Runtime gerekli", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (answer == DialogResult.Yes)
                Process.Start(new ProcessStartInfo(WebView2DownloadUrl) { UseShellExecute = true });
            return false;
        }
    }
}
