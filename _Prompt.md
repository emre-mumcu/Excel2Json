# Görev: Modern "Excel → JSON Dönüştürücü" Masaüstü Uygulaması (.NET + WinForms + Blazor Hybrid)

Bir kıdemli .NET geliştiricisi gibi davran. Aşağıda tarif edilen Windows masaüstü uygulamasını sıfırdan, üretim kalitesinde tasarla ve geliştir. Kod, mimari ve kullanıcı deneyimi açısından sürdürülebilir, test edilebilir ve genişletilebilir olmalı. Arayüz metinleri **Türkçe** olacak; kod içi isimlendirme (sınıf, metot, değişken) **İngilizce**, açıklayıcı yorumlar Türkçe olabilir.

---

## 1. Amaç

Kullanıcı bir Excel dosyası seçer, içinden bir sayfa (sheet) seçer, o sayfadaki istediği sütunları seçip her birine isterse farklı bir JSON alan adı (alias) ve veri tipi verir, Excel'de olmayan ek alanlar tanımlar, JSON çıktısının yapısını (object, array, array of objects vb.) belirler ve sonucu seçtiği bir klasöre `.json` dosyası olarak kaydeder.

## 2. Referans: Mevcut Basit Uygulama (bunun yerine geçecek)

`a:\AIStore\ExcelToJson` klasöründe bu işin basit bir sürümü var (.NET WinForms + ClosedXML). Gerekirse inceleyebilirsin ama doğrudan kopyalama; yeni mimariye göre yeniden yaz. Mevcut sürümün özeti:

- `.xlsx` / `.xlsm` okur, sayfaları listeler, kullanılan alanın ilk satırını başlık kabul eder.
- DataGridView'de sütun seçimi + JSON alan adı düzenleme var; alan adı başlıktan otomatik üretilir (harf/rakam dışı → `_`, ilk harf küçük).
- Tüm değerler `GetFormattedString()` ile **string** aktarılır (tip desteği yok).
- Çıktı: Array of objects veya satır numarası anahtarlı Object.
- Doğrulama: en az bir sütun seçilmeli, alan adı boş olamaz, alan adları benzersiz olmalı.
- Boş satırları atlar; JSON `UnsafeRelaxedJsonEscaping` ile Türkçe karakterleri kaçışsız yazar, UTF-8 (BOM'suz) kaydeder.

**Mevcut sürümün eksikleri (yeni sürümde çözülmeli):** veri tipi yok, ek alan yok, iç içe yapı yok, başlık satırı seçilemiyor, ayarlar/profil kaydedilmiyor, her sayfa değişiminde dosya yeniden açılıyor, Excel'de açık (kilitli) dosya okunamıyor, büyük dosyalarda UI donuyor, test yok, UI eski görünümlü.

---

## 3. Teknoloji ve UI Kararı

### Seçilen yaklaşım: WinForms host + Blazor Hybrid (BlazorWebView) + Fluent UI Blazor

- Uygulama bir **.NET WinForms** uygulamasıdır (`net11.0-windows`). Kullandığın paket sürümlerinin hedef framework ile uyumlu olduğunu doğrula.
- Ana form, `Microsoft.AspNetCore.Components.WebView.WindowsForms` paketindeki **`BlazorWebView`** kontrolünü tam ekran barındırır. Arayüz Razor bileşenleriyle C# olarak yazılır (JavaScript gerekmez). Görüntü katmanı WebView2'dir.
- Bileşen kütüphanesi: **Microsoft Fluent UI Blazor** (`Microsoft.FluentUI.AspNetCore.Components`). Windows 11 Fluent tasarım diliyle uyumlu, açık/koyu tema destekli ve ücretsizdir.
- Excel okuma: **ClosedXML** (mevcut uygulamayla aynı).
- JSON: **System.Text.Json** (`JsonNode` / `Utf8JsonWriter`).
- DI: `Microsoft.Extensions.DependencyInjection` (BlazorWebView'in `Services` koleksiyonu üzerinden).
- Loglama: `Microsoft.Extensions.Logging` + dosyaya log (ör. Serilog file sink) → `%LocalAppData%\ExcelToJson\logs`.

**Neden bu seçim:** Saf WinForms kontrolleri modern görünmüyor. Blazor Hybrid hem WinForms gereksinimini karşılıyor hem de web tabanlı modern bir arayüz (Fluent bileşenleri, tema, animasyon, responsive layout) sağlıyor. İş mantığının tamamı yine C#/.NET'te kalıyor.

### Hybrid'e özgü teknik notlar (dikkat et)
- **Dosya/klasör seçimi:** Native diyaloglar (`OpenFileDialog`, `FolderBrowserDialog`) bir `IDialogService` arkasında kullanılmalı ve WinForms UI thread'inde çalıştırılmalı (`Form.Invoke`/`SynchronizationContext`). Bu servisin Razor tarafından `await` edilebilir olması gerekiyor.
- **Sürükle-bırak:** WebView2 içine bırakılan dosyanın tam yolu alınamaz. Dosya sürükle-bırak desteği WinForms form seviyesinde (`AllowDrop`, `DragDrop`) yakalanmalı ve bir event/servis ile Blazor tarafına iletilmeli.
- **WebView2 Runtime:** Kurulu değilse uygulama açılışta kullanıcıya anlaşılır bir mesaj göstermeli ve indirme linki vermeli.
- **Uzun işlemler:** Excel okuma ve JSON üretimi `Task.Run` ile arka planda yapılmalı. UI'da ilerleme göstergesi ve iptal (`CancellationToken`) olmalı.

### Alternatif (sadece Blazor Hybrid'de çözülemeyen bir engel çıkarsa)
Saf WinForms + `Application.SetColorMode(SystemColorMode.System)` (koyu tema) + **Krypton Toolkit** (ücretsiz, modern görünümlü WinForms kontrolleri). Bu yola geçersen önce nedenini açıkla.

---

## 4. Çözüm Mimarisi

```
ExcelToJson.sln
├─ src/
│  ├─ ExcelToJson.Core/          (UI bağımsız class library – net10.0)
│  │   ├─ Models/                (WorkbookInfo, SheetInfo, SheetData, ColumnInfo, FieldMapping,
│  │   │                          CustomField, ExportOptions, OutputShape, DataType, MappingProfile…)
│  │   ├─ Excel/                 (IExcelReader, ClosedXmlExcelReader)
│  │   ├─ Conversion/            (IValueConverter: string→int/decimal/bool/date…)
│  │   ├─ Json/                  (IJsonBuilder, şekil stratejileri – her OutputShape için ayrı strateji)
│  │   ├─ Expressions/           (Ek alanlar için şablon/ifade çözümleyici)
│  │   ├─ Validation/            (Mapping ve export doğrulayıcıları, ValidationResult)
│  │   └─ Profiles/              (IProfileStore – JSON dosyası olarak kaydet/yükle)
│  └─ ExcelToJson.App/           (WinForms + BlazorWebView – net10.0-windows)
│      ├─ MainForm.cs            (BlazorWebView host, drag-drop köprüsü)
│      ├─ Services/              (DialogService, SettingsService, ThemeService, FileDropBridge)
│      ├─ Components/            (Razor bileşenleri)
│      ├─ Pages/ veya Steps/
│      └─ wwwroot/               (index.html, app.css)
└─ tests/
   └─ ExcelToJson.Core.Tests/    (xUnit; test Excel dosyaları kodla ClosedXML ile üretilsin)
```

İlkeler:
- Core katmanı UI'dan tamamen bağımsız olacak. İleride CLI veya başka bir UI eklenebilmeli.
- Çıktı şekilleri **Strategy pattern** ile uygulanacak (`IOutputShapeStrategy`). Yeni bir şekil eklemek mevcut kodu değiştirmeyi gerektirmemeli.
- Workbook dosyası bir kez okunup oturum boyunca cache'lenecek. Sayfa değiştirince dosya yeniden açılmayacak; "Yenile" butonu ile yeniden okunabilecek.
- Dosya, Excel'de açık olsa bile okunabilmeli: `FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)` ile açılmalı.

---

## 5. Fonksiyonel Gereksinimler

### 5.1 Dosya Seçimi
- "Dosya Seç" butonu (`.xlsx`, `.xlsm`) ve forma sürükle-bırak. `.xls` için anlaşılır bir hata mesajı gösterilmeli ("dosyayı .xlsx olarak kaydedin").
- Son açılan 10 dosya listesi (Recent files). Dosya artık yoksa listeden düşürülmeli.
- Dosya adı, boyutu ve sayfa sayısı gösterilmeli.

### 5.2 Sayfa (Sheet) Seçimi
- Sayfalar kart veya liste halinde gösterilmeli: ad, kullanılan alan (ör. `A1:K250`), tahmini satır/sütun sayısı, gizli sayfa ise işaret.
- Sayfa seçilince ilk N satırlık **ham veri önizlemesi** (ör. 20 satır) bir tabloda gösterilmeli.
- **Başlık ayarları:**
  - Başlık satırı numarası (varsayılan: kullanılan alanın ilk satırı)
  - Veri başlangıç satırı (varsayılan: başlık + 1)
  - Opsiyonel veri bitiş satırı
  - "Başlık yok" seçeneği (bu durumda sütun adları `A`, `B`, `C`… olur)
- Seçenekler: tamamen boş satırları atla (varsayılan açık), gizli satırları/sütunları dahil et/etme, birleştirilmiş hücrelerde (merged cells) değeri tüm hücrelere yay.

### 5.3 Sütun Eşleştirme (Mapping) Tablosu
Her Excel sütunu için bir satır:

| Alan | Açıklama |
|---|---|
| Dahil et | Checkbox. Ayrıca "tümünü seç / hiçbirini seçme / ters çevir" butonları |
| Sütun | Excel harfi (A, B…) – salt okunur |
| Excel başlığı | Salt okunur |
| Örnek değerler | İlk 3 dolu değer (tooltip veya küçük metin) |
| JSON alan adı (alias) | Düzenlenebilir. Başlıktan otomatik üretilir. Adlandırma stili seçilebilir: camelCase (varsayılan), PascalCase, snake_case, orijinal. Türkçe karakter dönüşümü seçeneği (ı→i, ş→s, ğ→g, ü→u, ö→o, ç→c) |
| Veri tipi | `Auto`, `String`, `Integer`, `Decimal`, `Boolean`, `Date`, `DateTime`, `Array (ayraçlı)`. Auto, örnek verilerden tip tahmin eder ama kullanıcı değiştirebilir |
| Tip ayarı | Tipe göre değişir: tarih için çıktı formatı (ISO 8601 varsayılan, özel format girilebilir), decimal için ondalık ayracı/kültür (tr-TR / Invariant), boolean için true/false kabul edilen değerler (`evet/hayır`, `1/0`, `x/boş`…), array için ayraç (`,` `;` `|`) ve eleman tipi |
| Boş değer | `null` yaz / `""` yaz / alanı hiç yazma / varsayılan değer kullan |
| Trim | Baştaki/sondaki boşlukları temizle (varsayılan açık) |

- Sütunlar sürüklenerek veya yukarı/aşağı butonlarıyla **sıralanabilmeli**. JSON alan sırası bu sırayı takip etmeli.
- Arama/filtre kutusu (çok sütunlu sayfalar için).
- **İç içe yapı (nesting):** Alias içinde nokta notasyonu iç içe nesne üretmeli. Örnek: `address.city` ve `address.zip` → `{ "address": { "city": "...", "zip": "..." } }`. Çakışmalar doğrulanmalı (ör. hem `address` hem `address.city` tanımlıysa hata).

### 5.4 Ek Alanlar (Excel'de olmayan, JSON'a eklenecek alanlar)
"Alan Ekle" ile yeni alan tanımlanır. Her alanın adı (nokta notasyonu destekli), tipi ve değer kaynağı olur:

| Kaynak türü | Örnek |
|---|---|
| Sabit değer | `"source": "ERP"`, `"isActive": true`, `"version": 2` |
| Şablon (diğer alanlardan) | `"{firstName} {lastName}"`, `"PRD-{productCode}"`. Placeholder'lar alias'a veya Excel başlığına referans verebilir |
| Satır numarası | Excel satır no veya 1'den başlayan sıra no |
| Otomatik artan | Başlangıç ve artış değeri ile |
| GUID | Her kayıt için yeni GUID |
| Tarih/Saat | Export anı (`DateTime.Now` / `UtcNow`, format seçilebilir) |
| Sabit JSON | Kullanıcının yazdığı geçerli bir JSON parçası (object/array). Kaydetmeden önce doğrulanmalı |

- Ek alanlar, sütunlarla aynı sıralama listesinde yer alabilmeli. Yani kullanıcı ek alanı istediği konuma taşıyabilmeli.
- Ayrıca **kök seviye ek alanlar** (her kayda değil, çıktının köküne bir kez eklenen alanlar) tanımlanabilmeli; bkz. 5.5 Wrapper.

### 5.5 Çıktı Yapısı (Output Shape)
Kullanıcı seçer. Her seçeneğin yanında küçük bir örnek çıktı gösterilmeli:

1. **Array of objects** (varsayılan): `[ {..}, {..} ]`
2. **Object (anahtarlı sözlük):** `{ "<anahtar>": {..} }`. Anahtar seçilen bir sütundan (ör. `productCode`) veya Excel satır numarasından gelir. Tekrarlanan anahtar politikası seçilebilmeli: hata ver / sonuncu kazansın / ilk kazansın / değeri diziye çevir. Seçenek: anahtar alanı nesnenin içinde de kalsın mı?
3. **Single object:** Sadece tek kayıt (ilk satır veya seçilen satır no) → `{ .. }`
4. **Key–value object:** İki sütundan sözlük → `{ "<keyColumn>": <valueColumn> }`. Ayar/konfigürasyon sayfaları için.
5. **Array of values:** Tek sütun → `[ "a", "b", "c" ]`
6. **Array of arrays:** Satırlar değer dizisi olarak → `[ ["P001","Laptop"], .. ]`. Opsiyonel ilk eleman başlık dizisi.
7. **Gruplanmış:** Seçilen sütuna göre gruplama → `{ "<grup>": [ {..}, {..} ] }` veya `[ { "<grupAlanı>": "..", "items": [ .. ] } ]` (iki varyant). Kullanıcı children dizisinin adını belirleyebilmeli.

**Wrapper (kök sarmalayıcı) – opsiyonel ve tüm şekillerle birleştirilebilir:**
```json
{
  "meta": { "source": "urunler.xlsx", "sheet": "Ürünler", "exportedAt": "2026-11-23T10:00:00Z", "count": 250 },
  "data": [ ... ]
}
```
- Kök özellik adı (`data`) değiştirilebilir. `meta` bloğu açılıp kapatılabilir ve kullanıcı buraya kendi kök alanlarını ekleyebilir.

### 5.6 Önizleme
- Canlı JSON önizlemesi: varsayılan olarak ilk 10 kayıt (ayarlanabilir). Ayarlar değiştikçe debounce ile güncellenmeli (~300 ms).
- Syntax highlighting (ör. basit bir Razor bileşeni veya hafif bir highlight kütüphanesi), satır numaraları, "Panoya kopyala" butonu.
- Önizleme altında özet: toplam kayıt, atlanan boş satır, dönüşüm uyarısı sayısı.

### 5.7 Doğrulama ve Hata Yönetimi
- Export öncesi doğrulamalar:
  - En az bir alan seçilmeli.
  - Alias boş olmamalı ve geçerli olmalı.
  - Aynı seviyede alan adları benzersiz olmalı (büyük/küçük harf duyarsız uyarı, duyarlı hata).
  - Nokta notasyonu çakışmaları olmamalı.
  - Şablon placeholder'ları var olan alanlara referans vermeli.
  - Anahtar sütunu seçilmiş olmalı (anahtarlı şekillerde).
  - Sabit JSON geçerli olmalı.
- **Tip dönüşüm hataları:** Her hata satır no, sütun, ham değer ve hedef tip ile listelenmeli. Politika seçilebilir: export'u durdur / hatalı değeri `null` yap / ham string olarak bırak. Hata listesi bir panelde gösterilmeli ve tıklanınca ilgili alan vurgulanmalı.
- Tüm beklenmeyen hatalar loglanmalı. Kullanıcıya teknik olmayan Türkçe bir mesaj ve "Detaylar" bölümü gösterilmeli.

### 5.8 Export ve Kaydetme
- **Hedef klasör seçimi** (`FolderBrowserDialog`). Son kullanılan klasör hatırlanmalı.
- **Dosya adı şablonu:** varsayılan `{sheet}_{yyyyMMdd_HHmmss}.json`. Desteklenen placeholder'lar: `{file}`, `{sheet}`, `{date}`, `{yyyyMMdd}` vb. Geçersiz dosya adı karakterleri temizlenmeli.
- Dosya varsa: üzerine yaz / numara ekle / sor.
- Format ayarları:
  - Girintili (indented) veya minified
  - Girinti boyutu
  - Türkçe karakterleri kaçışsız yaz (varsayılan açık)
  - Encoding: UTF-8 BOM'suz (varsayılan) / BOM'lu
- Yazma önce geçici dosyaya yapılıp sonra taşınmalı (atomik yazım).
- Export sonrası toast bildirimi gösterilmeli: "Klasörü aç" ve "Dosyayı aç" aksiyonlarıyla.
- **(Opsiyonel – Faz 3)** Çoklu sayfa export: seçilen birden fazla sayfayı, her biri kendi profiliyle, ayrı dosyalara veya tek dosyada `{ "<sheet>": [...] }` olarak export etme.

### 5.9 Profiller (Mapping Şablonları)
- Tüm ayarlar (başlık ayarları, mapping, ek alanlar, çıktı şekli, export ayarları) **profil** olarak JSON dosyasına kaydedilip yüklenebilmeli.
- Profil `%AppData%\ExcelToJson\profiles` altında tutulmalı. Ayrıca dışa/içe aktarılabilmeli (paylaşım için).
- Aynı dosya + sayfa tekrar açıldığında son kullanılan profil otomatik önerilmeli.
- Profil yüklenirken Excel'de artık bulunmayan sütunlar uyarı olarak gösterilmeli. Eşleştirme sütun harfine değil **başlık adına** göre yapılmalı; başlık bulunamazsa harfe göre yapılmalı.
- Profil formatı bir `schemaVersion` alanı içermeli (ileride migration için).

### 5.10 Uygulama Ayarları
- Tema: Sistem / Açık / Koyu (Fluent `DesignTheme`)
- Varsayılan adlandırma stili, varsayılan çıktı şekli, önizleme kayıt sayısı
- Ayarlar `%AppData%\ExcelToJson\settings.json` dosyasında tutulmalı.
- Pencere boyutu ve konumu hatırlanmalı.

---

## 6. UI / UX Tasarımı

- Fluent UI tabanlı, sade ve ferah bir tasarım. Sol tarafta dikey bir **adım navigasyonu (wizard/stepper)**, sağda içerik alanı:
  1. **Dosya** – seçim kartı, sürükle-bırak alanı, son dosyalar
  2. **Sayfa** – sayfa kartları + ham veri önizleme + başlık ayarları
  3. **Alanlar** – mapping tablosu + ek alanlar (sekme veya bölüm)
  4. **Çıktı** – şekil seçimi (görsel kartlar ve örnek JSON), wrapper ayarları
  5. **Export** – klasör, dosya adı, format ayarları, son doğrulama özeti, "Export" butonu
- Adımlar arasında serbest geçiş olmalı (sıkı sihirbaz değil). Eksik veya hatalı adımlar navigasyonda uyarı rozetiyle gösterilmeli.
- **Alanlar, Çıktı ve Export** adımlarında sağda daraltılabilir bir **canlı JSON önizleme paneli** olmalı (geniş ekranda yan yana, dar ekranda alt alta).
- Üst çubukta: uygulama adı, aktif dosya/sayfa bilgisi, profil kaydet/yükle menüsü, tema seçici.
- Uzun işlemlerde progress ring/bar ve iptal butonu gösterilmeli. UI hiçbir zaman donmamalı.
- Klavye kısayolları:
  - `Ctrl+O` – dosya aç
  - `Ctrl+S` – profili kaydet
  - `Ctrl+E` – export
  - `F5` – dosyayı yenile
- Erişilebilirlik: yeterli kontrast, focus göstergeleri, tooltip'ler.
- Minimum pencere boyutu 1024×700; yüksek DPI desteği (`ApplicationHighDpiMode.PerMonitorV2`).

---

## 7. Teknik Detaylar ve Edge Case'ler

- **Hücre okuma:** Tipli dönüşüm için ham değer (`XLCellValue`: number / DateTime / bool / text) kullanılmalı. `String` tipi için `GetFormattedString()` kullanılmalı. Excel'de sayı olarak saklanan tarihler `Date` tipinde doğru dönüştürülmeli.
- **Formüller:** ClosedXML formülleri yeniden hesaplamaz; cache'lenmiş değer kullanılmalı. Hesaplanmış değer yoksa uyarı verilmeli.
- **Kültür:** Girdi parse işlemi kullanıcının seçtiği kültürle (varsayılan tr-TR), JSON çıktısı her zaman invariant ile yapılmalı (sayılar `.` ile, tarihler ISO 8601).
- **Büyük sayılar:** 15 haneden uzun sayılar (TCKN, barkod vb.) `Auto` modda string önerilmeli. Baştaki sıfırlar korunmalı.
- **Büyük dosyalar:** 100k+ satırda bellek ve süre makul olmalı. Önizleme yalnızca ilk N kaydı üretmeli. Export, `Utf8JsonWriter` ile stream olarak dosyaya yazmalı (tüm JSON'u string olarak bellekte tutmamalı).
- Boş başlıklar `Column X` olarak adlandırılmalı. Aynı başlıktan birden fazla varsa otomatik alias'lar `name`, `name_2` şeklinde benzersizleştirilmeli.
- Async/await doğru kullanılmalı: `async void` yalnızca event handler'larda, UI thread'e dönüş güvenli olmalı.

---

## 8. Test

`ExcelToJson.Core.Tests` (xUnit, FluentAssertions veya Shouldly isteğe bağlı) en az şunları kapsamalı:
- Excel okuma: başlık satırı seçimi, boş satır atlama, merged cell, gizli satır, kilitli dosya.
- Tip dönüşümleri: tr-TR ondalık, tarih (serial number + metin), boolean varyantları, ayraçlı array, büyük sayılar, boş değer politikaları.
- Alias üretimi: camelCase, snake_case, Türkçe karakter dönüşümü, tekrarlanan başlıklar.
- Her çıktı şekli için beklenen JSON (snapshot/string karşılaştırması). Buna anahtar çakışma politikaları ve gruplama dahil.
- Nokta notasyonuyla iç içe yapı ve çakışma doğrulaması.
- Ek alanlar: sabit, şablon, otomatik artan, GUID formatı, sabit JSON doğrulaması.
- Profil kaydet/yükle round-trip ve başlık adına göre yeniden eşleştirme.

Test Excel dosyaları testler içinde ClosedXML ile programatik olarak üretilmeli (repo'ya binary dosya koyma).

---

## 9. Çalışma Planı (fazlar halinde ilerle)

Her fazın sonunda derlenebilir ve çalışır bir durum olsun. Faz sonunda ne yaptığını, nasıl test edildiğini ve açık kalan konuları kısaca raporla.

- **Faz 1 – İskelet & Çekirdek:** Solution yapısı, BlazorWebView + Fluent UI kurulumu, tema, DialogService, Excel okuma (cache + FileShare), sayfa seçimi, başlık ayarları, basit mapping (seç + alias), Array of objects export, klasöre kaydetme, temel testler.
- **Faz 2 – Tipler & Şekiller:** Veri tipleri ve dönüşüm politikaları, boş değer politikaları, nokta notasyonu, tüm çıktı şekilleri + wrapper, canlı önizleme, doğrulama paneli.
- **Faz 3 – Ek Alanlar & Profiller:** Ek alan türleri, kök ek alanlar, profil kaydet/yükle/otomatik öner, uygulama ayarları, son dosyalar, sürükle-bırak, dosya adı şablonu.
- **Faz 4 – Cilalama:** Büyük dosya performansı (stream export, iptal), klavye kısayolları, toast bildirimleri, çoklu sayfa export (opsiyonel), README (kurulum, kullanım, ekran açıklamaları), testlerin tamamlanması.

## 10. Kabul Kriterleri

- `dotnet build` uyarısız (veya gerekçeli uyarılarla) derlenmeli ve `dotnet test` yeşil olmalı.
- Kullanıcı dosya seç → sayfa seç → sütun seç/alias/tip ver → ek alan ekle → çıktı şekli seç → klasör seç → export akışını fare ile ve takılmadan tamamlayabilmeli.
- Üretilen JSON her şekil için geçerli olmalı ve seçilen tiplere uygun olmalı (sayılar tırnaksız, boolean `true/false`, tarihler ISO).
- Excel'de açık bir dosya okunabilmeli. 100.000 satırlık bir sayfa UI'ı dondurmadan export edilebilmeli.
- Profil kaydedip uygulamayı yeniden başlattıktan sonra aynı çıktı tek tıkla yeniden üretilebilmeli.

## 11. Çalışma Şekli

- Belirsiz veya birden fazla makul seçeneği olan bir kararda (ör. paket sürümü uyumsuzluğu, UI teknolojisinde engel) önce kısa bir öneriyle bana sor. Küçük kararlarda makul varsayılanı seçip devam et ve raporda belirt.
- Kullanılan NuGet paketlerinin güncel ve kararlı sürümlerini kullan. Ücretli/lisanslı bileşen kullanma.
- Gereksiz soyutlama yapma; okunabilirlik ve basitlik öncelikli olsun. Ancak Core/UI ayrımı ve strateji yapısı korunmalı.
