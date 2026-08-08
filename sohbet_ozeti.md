# Stok Panel — Durum Özeti (07.08.2026 itibarıyla)

Bu dosya, yeni bir sohbete geçerken Claude'a bağlam vermek için hazırlanmıştır.
**Yeni sohbette kullanıcının söylemesi gereken tek şey:** "GitHub'daki stok-panel
reposunu çek, analiz et" — Claude tüm dosyaları (bu dosya dahil) doğrudan
`raw.githubusercontent.com/fincanlarnecip/stok-panel/main/...` üzerinden okur,
zip/Drive/manuel yükleme gerekmez (repo public).

## Proje nedir
Google Apps Script backend (`Kod.js`) + tek dosya HTML/JS frontend (`index.html`)
ile çalışan bir seramik/parke mağazası stok takip paneli + artık **PWA (telefona
kurulabilir uygulama)**. GitHub Pages üzerinde barındırılıyor, veriler Google
Sheets'te tutuluyor. Repo: https://github.com/fincanlarnecip/stok-panel
(sahibi: `fincanlarnecip`, çalışma hesabı: `fincanlaryapi@gmail.com`)

## ⚙️ DEPLOY ARTIK TAMAMEN OTOMATİK — manuel adım YOK
- `index.html` değiştir → commit et → GitHub Pages birkaç dakikada günceller.
- `Kod.js` (backend) değiştir → commit et → GitHub Actions (`.github/workflows/deploy.yml`)
  otomatik olarak `clasp push` + `clasp deploy -i <deploymentId>` çalıştırır, birkaç
  saniyede Apps Script'e yansır. **Apps Script editörüne hiç girmeye gerek yok.**
- ⚠️ **KRİTİK — clasp deploy MUTLAKA `-i <deploymentId>` ile çağrılmalı.** Deployment
  ID olmadan her `clasp deploy` YENİ bir deployment (farklı URL) oluşturur ve panelin
  gerçekten kullandığı sabit adres GÜNCELLENMEZ (bu yüzden bir ara "pin silme" gibi
  backend değişiklikleri sessizce hiç işlemiyordu — kök neden buydu, düzeltildi).
  Doğru deployment ID: `AKfycbxHEjTuSVDVCL1EcoaPWdxHGxjPTOfUrf0mrIRZfjkoiMl7VC-Bcr6gIMqbTzH0vmr4Yw`
  (bu, `index.html` içindeki `const API = ...` ile aynı olmalı — değişirse ikisini
  birden güncelle).
- `.claspignore` dosyası var ve **SADECE `Kod.js` + `appsscript.json`** push ediliyor.
  Bunu SAKIN silme — silinirse repo kökündeki `index.html`, `lib/*.js`, ve eski
  `fincanlarstok_v3_fixed.gs` gibi dosyalar da yanlışlıkla Apps Script projesine
  push edilir ve `SHEET_ID` gibi değişkenler iki kez tanımlanıp
  "Identifier already declared" syntax hatası verir (bir kez yaşandı, düzeltildi —
  Apps Script editöründen elle 3 fazlalık dosya silindi).
- Script ID: `17AUz3HZF0RBGkKQkQcCcSLDtc44zLCmzS4GsQwvOJZT5BgOV-1ZYcxDl`

## 🔑 GitHub yazma erişimi (Claude'un doğrudan commit atabilmesi için)
Kullanıcı Claude'a **fine-grained personal access token** veriyor (sadece
`stok-panel` reposu, **Contents: Read and write** + **Workflows: Read and write**
izinli — Workflows izni `.github/workflows/*` dosyalarını değiştirebilmek için
şart, sıradan Contents izni yetmiyor). Token her yeni sohbette kullanıcı tarafından
**tekrar verilmesi gerekir** (Claude'un sohbetler arası hafızası yok). Token
`github.com/settings/personal-access-tokens` adresinden istenildiğinde iptal
edilebilir. Token ile Claude GitHub API'sinin `contents` endpoint'ini kullanarak
(bash + curl, base64 encode) doğrudan commit atıyor — kullanıcının dosya
indirip/yükleme yapmasına gerek kalmıyor.

## 📱 PWA (telefon uygulaması) kuruldu — 07.08.2026
- `manifest.json`, `sw.js` (service worker), `icon-192.png`, `icon-512.png` eklendi.
- Chrome'da "Ana ekrana ekle" / "Uygulamayı yükle" ile telefon uygulaması gibi kurulabiliyor.
- ⚠️ **Service worker "network-first" stratejisiyle çalışıyor** (v2, CACHE_ADI="stok-panel-v2") —
  yani sayfa isteklerinde HER ZAMAN önce ağdan taze veri çekmeye çalışır, sadece
  offline'da önbelleğe düşer. Bu, ilk PWA denemesinde "uygulama güncellenmiyor"
  sorununu çözmek için sonradan eklendi (ilk versiyon cache-first'ti, hataliydi).
  **Yeni bir önbellekleme sorunu çıkarsa** CACHE_ADI'yi bir üst versiyona
  (v3, v4...) çıkarmak eski önbelleği otomatik temizler (activate handler'da
  eski cache key'leri siliniyor).
- Kullanıcı bir ara "hala eski görünüyor" sorunu yaşadı ama kök neden kodda değil,
  cihazın WebAPK önbelleğindeydi — çözüm: **normal tarayıcıda (gizli sekmede)
  doğrula, sonra uygulamayı kaldır+yeniden kur**. Bu sıra takip edilirse sorun
  çıkmıyor.

## 📱 Mobil için yapılan UI değişiklikleri (07.08.2026)
Hepsi `@media (max-width: 640px)` içinde, masaüstü görünümü etkilenmiyor:
- **Stok listesi tablo → kart görünümü**: her ürün ayrı kutu, alt alta. STOK_KODU,
  MARKA, ETIKET, NOT, NUMUNE, GORSEL sütunları mobilde gizli. STOK_ADI kartın
  başlığı, altında Merkez Depo/Ayrılmış/Satılabilir/Fiyat etiket+değer satırları.
- **Arama kutusu** artık kendi satırında, boydan boya (16px font — iOS zoom'unu
  önlemek için), diğer kontroller (Tam eşleşme, Filtre, Excel, 0 ve eski) altına
  sarıyor.
- **Tarih rozeti** (`#dbadge`) mobilde JS ile (`mobilTarihRozetiKonumla()`, resize
  listener'lı) topbar'dan alınıp arama barındaki `.pills` satırının sonuna
  taşınıyor; masaüstünde eski yerinde (sağ üst) kalıyor.
- **"📦 Listedeki Depo Toplamı"** satırı eklendi — tarih rozetinin hemen üstünde,
  o an filtrelenmiş/listelenen ürünlerin Merkez Depo toplamını gösterir
  (`updStats()` içinde hesaplanıyor, `#s-depo-toplam` id'li elemente yazılıyor).
- **Arama/etiket sıralama mantığı değişti:** Sonuçlar artık **Merkez Depo
  miktarına göre çoktan aza** sıralanıyor:
  - Arama kutusuna yazınca (`applyF` içindeki `rows.sort`): tam eşleşenler yine
    üstte, aynı grup içinde depo miktarına göre azalan.
  - Etiket filtresi (sol menüden etikete tıklama, arama kutusu boşken): açılış
    önceliği / marka içi sıralama **atlanıyor**, doğrudan depo miktarına göre
    azalan sıralanıyor (kullanıcı özellikle böyle istedi: "etiket seçilirse
    etiketi dikkate al").
  - Bu iki sıralama bloğu ayrı kod parçaları, ikisi de güncellendi.

## 🗺️ Mağaza Bölümleri — mobil "nokta → kart" pin sistemi (07.08.2026)
Hem **Ayarlar → Mağaza Bölümleri** (düzenlenebilir, `mbPinlerCiz()`) hem de
**üst menü → Kontrol** (salt-okunur, `kontrolPinlerCiz()`) sayfalarında aynı mantık:
- Mobilde (`window.innerWidth<=640`) pinler varsayılan olarak **küçük renkli
  noktalar** (16px daire, stok durumuna göre kırmızı/sarı/yeşil) — isim yazısı yok.
- Bir noktaya dokununca **kart olarak genişliyor**: ürün adı/ölçüsü üstte, miktar
  altta **sarı renkte, "m²" birimiyle** (`mbGenisletilenPinId` / `kontrolGenisletilenPinId`
  global değişkenleriyle takip ediliyor — tek seferde sadece bir pin açık olabilir).
- Aynı karta tekrar dokununca **kapanıyor**; başka bir noktaya dokununca eskisi
  kapanıp yenisi açılıyor; fotoğrafın boş alanına dokununca da açık kart kapanıyor
  (Ayarlar'da `mbFotoTiklandi()`, Kontrol'de `kontrol-img`'nin onclick'i ile).
- Ayarlar'daki genişletilmiş kartta ayrıca **⋮ simgesi** var — ona dokununca
  "Yön Değiştir / Kaldır" menüsü açılıyor (`mbPinMenuAc`). Kontrol'de bu yok
  (salt-okunur, düzenleme imkanı yok).
- **Masaüstünde davranış değişmedi** — Ayarlar'da eski iki satırlı etiket (isim
  üstte, miktar altta), Kontrol'de küçük/kısaltılmış etiket + hover'da büyüme
  (CSS `.kontrol-pin-label:hover`).
- Bölüm açılıp kapandığında (`mbBolumAc`/`mbEditorKapat`,
  `kontrolBolumAc`/`kontrolEditorKapat`) genişletilen pin id'si sıfırlanıyor.

## 🐛 Düzeltilen hatalar (07.08.2026)
- **Pin silme çalışmıyordu** ("Kaldır"a basınca hiçbir şey olmuyordu, ne hata ne
  değişiklik): Kök neden yukarıdaki "clasp deploy -i olmadan yeni deployment
  oluşturma" sorunuydu — panel hiçbir zaman güncel backend kodunu çalıştırmıyordu.
  `-i <deploymentId>` eklenince ve `.claspignore` ile dosya kirliliği (duplicate
  SHEET_ID) temizlenince düzeldi. Ayrıca backend (`magazaPinSil`) artık gerçekten
  silinmediyse `{ok:false, hata:"..."}` dönüyor (önceden sessizce `{ok:true}`
  dönüyordu), frontend (`mbPinSil`) da hata varsa `alert()` ile gösteriyor —
  ileride benzer bir sorun çıkarsa artık sessiz kalmaz, hemen görünür olur.
- **Kontrol sayfası eklenirken DOM yapısı bozulmuştu** (üstte büyük boş alan
  çıkıyordu) — `page-kontrol` div'i yanlışlıkla ana sayfa sarmalayıcısının
  dışına eklenmişti, doğru katmana taşınarak düzeltildi.

## ⏳ AÇIK KONU — "Sistem aktif" gecikmesi (henüz çözülmedi)
Panel açılışında "Bağlanıyor..." → "Sistem aktif" durumuna geçmesi uzun sürüyor
şikayeti var. Kök neden büyük ihtimalle backend'deki `getAllData()` fonksiyonunun
birden fazla ayrı sheet/spreadsheet okuması (Stoklar, ayrı Fiyat spreadsheet'i,
GPD, Parke, Notlar, Etiketler, UrunEt, Markalar, Ayarlar, Numuneler+Gruplar,
Teklifler, UrunGorseller, MagazaBolumleri — hepsi sıralı, ayrı ayrı okunuyor) +
Apps Script'in soğuk başlangıç gecikmesi. Kodun içinde zaten adım adım süre
ölçen `⏱ [N]` logları var. **Sonraki adım:** kullanıcı bilgisayardan Apps Script
editörü → Executions (Yürütmeler) panelinden gerçek süre dağılımını (hangi adım
ne kadar sürüyor) paylaşacak, ona göre hedefli optimizasyon yapılacak (örn.
CacheService — ama tek değer 100KB sınırı olduğu için tüm getAllData çıktısını
doğrudan cachelemek mümkün olmayabilir, parçalı/kısmi cache gerekebilir).

## Dosyalar (repo kökü)
- `index.html` — tüm frontend (~760KB+)
- `Kod.js` — tüm backend (Apps Script), clasp ile senkron
- `appsscript.json`, `.clasp.json` — Apps Script proje bağlantı dosyaları
- `.claspignore` — **SADECE Kod.js + appsscript.json push edilsin diye** (silme!)
- `.github/workflows/deploy.yml` — otomatik deploy (clasp push + deploy -i ile)
- `manifest.json`, `sw.js`, `icon-192.png`, `icon-512.png` — PWA dosyaları
- `lib/xlsx.full.min.js`, `lib/pdf.min.mjs`, `lib/pdf.worker.min.mjs` — yerel kütüphaneler
- `gpd_fiyat_listesi.csv` — artık kullanılmıyor (Liste Güncelleme özelliği yerini aldı)
- `sohbet_ozeti.md` — bu dosya

## Bilinen/onaylanmamış eski konular (hâlâ geçerli olabilir, doğrulanmadı)
- Backend'deki `sheetToObj` fonksiyonundaki `===\"1\"`/`!==\"0\"` gibi katı string
  karşılaştırmalarının frontend'de hâlâ gözden geçirilmemiş olma ihtimali var —
  "tik kayboluyor" tarzı bir şikayet gelirse buradan başla.
- GPD_MARKA_KODU tarihte 89'dan 20'ye taşındı — çok eski/deploy edilmemiş bir
  önbellek varsa hâlâ 89 görünebilir.

## Genel ilerleyiş notu
Kullanıcı teknik detaylara çok hakim değil (adım adım, ekran görüntülü yönlendirme
gerekiyor), ama artık GitHub üzerinden tam otomatik bir sistemi var ve bunu
sorunsuz kullanabiliyor. Değişiklik isteklerinde çoğu zaman "mobil için" diye
belirtiyor — masaüstü davranışını bozmadan sadece mobil CSS/JS dallarını
hedeflemeye özellikle dikkat et.
