# Excel → JSON Dönüştürücü

Excel çalışma kitaplarını (.xlsx / .xlsm) yapılandırılabilir JSON dosyalarına dönüştüren Windows masaüstü uygulaması.
WinForms + Blazor Hybrid (BlazorWebView) + Microsoft Fluent UI Blazor ile yazılmıştır.

## Gereksinimler

- Windows 10/11
- .NET 11 SDK (geliştirme) / .NET 11 Desktop Runtime (çalıştırma)
- Microsoft Edge **WebView2 Runtime** (Windows 11'de hazır gelir; yoksa uygulama açılışta indirme bağlantısı gösterir)

## Derleme, test ve çalıştırma

```powershell
dotnet build ExcelToJson.sln
dotnet test                      # Microsoft.Testing.Platform (global.json) + xUnit v3
dotnet run --project src/ExcelToJson.App
```

## Çözüm yapısı

```
ExcelToJson.sln
├─ src/ExcelToJson.Core        (net10.0, UI bağımsız)
│   ├─ Models/        WorkbookInfo, SheetInfo, CellValue, FieldDefinition (ColumnField / CustomField), MappingProfile…
│   ├─ Excel/         IExcelReader, ClosedXmlExcelReader, IWorkbookSession (sayfa cache'i)
│   ├─ Conversion/    IValueConverter, ValueConverter, TypeInference
│   ├─ Mapping/       AliasGenerator, FieldFactory
│   ├─ Expressions/   TemplateExpression ("{firstName} {lastName}")
│   ├─ Json/          ConversionPlan, Record, JsonExportEngine, Shapes/ (IOutputShapeStrategy + 7 strateji)
│   ├─ Validation/    ProfileValidator, ValidationResult
│   ├─ Profiles/      IProfileStore, FileProfileStore, ProfileSerializer (schemaVersion), ProfileReconciler
│   └─ Export/        JsonFileExporter (atomik yazım), MultiSheetExporter, FileNameTemplate
├─ src/ExcelToJson.App         (net11.0-windows, WinForms + BlazorWebView)
│   ├─ MainForm.cs    BlazorWebView host, sürükle-bırak ve kısayol köprüsü
│   ├─ Services/      AppState, DialogService, SettingsService, ThemeService, FileDropBridge, ShellService, NotificationService
│   ├─ Components/    Kabuk, adım navigasyonu, önizleme paneli, diyaloglar, editörler
│   ├─ Steps/         FileStep, SheetStep, FieldsStep, OutputStep, ExportStep
│   └─ wwwroot/       index.html, css/app.css
└─ tests/ExcelToJson.Core.Tests (xUnit v3; test Excel dosyaları ClosedXML ile kodda üretilir)
```

## Kullanım

Sol taraftaki adımlar arasında serbestçe geçilebilir; eksik veya hatalı adımlar kırmızı/turuncu rozetle işaretlenir.

1. **Dosya**: "Dosya Seç" (Ctrl+O), pencereye sürükle-bırak veya son açılan 10 dosyadan biri. Dosya Excel'de açık olsa bile okunur.
2. **Sayfa**: Sayfa kartları (kullanılan alan, satır/sütun sayısı, gizli işareti), ilk 20 satırın ham önizlemesi.
   Başlık satırı, veri başlangıç/bitiş satırı, "başlık yok", boş satır atlama, gizli satır/sütun, birleştirilmiş hücre seçenekleri.
   Aynı dosya + sayfa için kayıtlı bir profil varsa uygulanması önerilir.
3. **Alanlar**: Sütun seçimi, alias (nokta notasyonu = iç içe nesne), veri tipi ve tipe özgü ayarlar (🔧), boş değer politikası, trim.
   Arama, tümünü seç / hiçbiri / ters çevir, adlandırma stili (camelCase, PascalCase, snake_case, orijinal) ve Türkçe karakter dönüşümü.
   Sıralama sürükle-bırak veya oklarla yapılır. **Alan Ekle** ile ek alan: sabit, şablon, satır no, otomatik artan, GUID, tarih/saat, sabit JSON.
4. **Çıktı**: Yedi çıktı şekli (örnek çıktılarıyla) ve şekle özgü ayarlar; isteğe bağlı kök sarmalayıcı (`meta` + `data`) ve kök alanlar.
5. **Export**: Hedef klasör, dosya adı şablonu (`{file}`, `{sheet}`, `{profile}`, `{date}`, `{time}`, `{yyyyMMdd_HHmmss}`…),
   dosya varsa davranışı, girinti, Türkçe karakterleri kaçışsız yazma, BOM. Son kontrol listesi ve **Export** (Ctrl+E).

### Çoklu sayfa export

Çalışma kitabında birden fazla dolu sayfa varsa Export adımında **Çoklu sayfa export** kartı görünür:

- Export edilecek sayfaları seçin; her sayfa için profil seçin:
  - **Mevcut ayarlar**: yalnızca açık sayfa için; Alanlar/Çıktı adımlarındaki güncel ayarlar,
  - **Varsayılan**: tüm sütunlar, Auto tip, nesne dizisi,
  - **Kayıtlı profil**: ★ işaretliler bu dosya + sayfa için kaydedilmiştir (otomatik seçilir); başka sayfanın profili de
    seçilebilir, sütunlar başlık adına göre yeniden eşleştirilir.
- **Ayrı dosyalar**: her sayfa, kendi profilinin dosya adı şablonu ve format ayarlarıyla ayrı dosyaya yazılır
  (aynı ada çözülen dosyalara numara eklenir). Bir sayfanın başarısız olması diğerlerini durdurmaz.
- **Tek dosya**: `{ "Ürünler": [...], "Ayarlar": {...} }`. Her sayfanın değeri kendi profilinin çıktı şekliyle üretilir;
  girinti/kodlama Export adımındaki format ayarlarından gelir. Herhangi bir sayfa başarısız olursa dosya yazılmaz.
- Export öncesi tüm sayfaların profilleri doğrulanır. Hedef klasör ve "dosya zaten varsa" ayarı ortaktır
  ("Bana sor" seçiliyse tüm çakışmalar için tek soru sorulur).

Alanlar, Çıktı ve Export adımlarında sağda (dar pencerede altta) daraltılabilir **canlı JSON önizlemesi** bulunur:
sözdizimi renklendirme, satır numaraları, panoya kopyalama, toplam kayıt / atlanan boş satır / dönüşüm uyarısı özeti.
Bir dönüşüm hatasına tıklamak ilgili alana gider ve onu vurgular.

### Klavye kısayolları

| Kısayol | İşlem |
|---|---|
| Ctrl+O | Dosya aç |
| Ctrl+S | Profili kaydet |
| Ctrl+E | Export |
| F5 | Dosyayı diskten yeniden oku (eşleştirme korunur) |

### Veri tipleri

| Tip | Davranış |
|---|---|
| Auto | Hücrenin Excel tipine göre; 15 haneden uzun sayılar string |
| String | Excel'de görünen biçimlendirilmiş metin (baştaki sıfırlar korunur) |
| Integer / Decimal | Metin sayılar seçilen kültürle (tr-TR / Invariant) okunur; çıktı her zaman `1234.56` |
| Boolean | Excel TRUE/FALSE, 1/0 ve yapılandırılabilir değer listeleri (evet/hayır, x/boş…) |
| Date / DateTime | Excel tarihi, seri numarası veya metin; çıktı ISO 8601 ya da özel format |
| Array | Ayraçlı metin (`,` `;` `\|` veya özel) → tipli eleman dizisi |

Dönüşüm hatası politikası: export'u durdur / null yap / ham metni bırak.

### Profiller ve ayarlar

| Ne | Nerede |
|---|---|
| Profiller | `%AppData%\ExcelToJson\profiles\*.profile.json` |
| Uygulama ayarları | `%AppData%\ExcelToJson\settings.json` (tema, varsayılanlar, son dosyalar, pencere konumu) |
| Günlükler | `%LocalAppData%\ExcelToJson\logs\log-YYYYMMDD.txt` (14 gün) |

Profil yüklenirken sütunlar önce **başlık adına**, bulunamazsa sütun harfine göre eşleştirilir; bulunamayan sütunlar uyarı olarak gösterilir.
Profil dosyaları `schemaVersion` içerir; dışa/içe aktarma ile paylaşılabilir.

## Teknik notlar

- **Cache**: Çalışma kitabı `FileShare.ReadWrite` ile açılıp belleğe kopyalanır (dosya kilitlenmez). Her sayfa ilk erişimde bir kez
  değişmez bir hücre cache'ine aktarılır; sayfa değiştirmek dosyayı yeniden açmaz.
- **Formüller**: ClosedXML formülleri yeniden hesaplamaz; dosyadaki cache'lenmiş değer kullanılır. Değer yoksa önizlemede uyarı gösterilir.
- **Büyük dosyalar**: JSON `Utf8JsonWriter` ile doğrudan dosyaya stream edilir; önizleme yalnızca ilk N kaydı üretir.
  100.000 satırlık sayfa testlerde birkaç saniyede export edilir; işlem ilerleme çubuğu ile arka planda çalışır ve iptal edilebilir.
  Not: Gruplanmış şekil ile "sonuncu kazansın" / "diziye çevir" politikaları doğaları gereği kayıtları bellekte toplar.
- **Atomik yazım**: Önce aynı klasörde geçici dosyaya yazılır, başarılı olunca hedefe taşınır.
- **Yeni çıktı şekli eklemek**: `IOutputShapeStrategy` uygulayın ve `OutputShapeRegistry`'ye kaydedin; mevcut kod değişmez.
- **Sürükle-bırak**: BlazorWebView pencerenin tamamını kapladığı için form seviyesindeki `DragDrop` olayları WebView2'ye ulaşmaz ve
  DOM bırakılan dosyanın tam yolunu vermez. `index.html` içindeki küçük bir köprü betiği dosyaları
  `chrome.webview.postMessageWithAdditionalObjects` ile host'a gönderir; `MainForm` tam yolu `CoreWebView2File.Path` ile alıp
  `FileDropBridge` üzerinden Blazor'a iletir. Klavye kısayolları da aynı köprüden geçer (tarayıcı kısayolları kapatılmıştır).
