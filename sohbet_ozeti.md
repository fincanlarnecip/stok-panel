# Stok Panel — Durum Özeti (05.08.2026 itibarıyla)

Bu dosya, yeni bir sohbete geçerken Claude'a bağlam vermek için hazırlanmıştır.

## Proje nedir
Google Apps Script backend (`fincanlarstok_v3_fixed.gs`) + tek dosya HTML/JS frontend
(`index.html`) ile çalışan bir seramik/parke mağazası stok takip paneli. GitHub Pages
üzerinde barındırılıyor, veriler Google Sheets'te tutuluyor.

## ÖNEMLİ — Deploy adımları (her `.gs` değişikliğinde ZORUNLU)
1. `index.html` → GitHub reposunda dosyayı değiştirip commit et.
2. `fincanlarstok_v3_fixed.gs` → Google Apps Script editöründe içeriği değiştir, kaydet,
   sonra **Deploy → Manage deployments → (kalem ikonu) → Version: New version → Deploy**
   yapılmadan değişiklikler ASLA devreye girmez.
3. `lib/` klasöründeki dosyalar (xlsx.full.min.js, pdf.min.mjs, pdf.worker.min.mjs) —
   GitHub'da `lib` klasörünün İÇİNE yüklenmeli (kök dizine değil).

## Bu sohbette tamamlanan başlıca özellikler
- **Excel/PDF kütüphaneleri yerelleştirildi** (CDN erişim sorunu için) — `lib/` klasöründe.
- **Ayarlar sayfasına eklenen yeni sekmeler:**
  - 🚫 Liste Dışı — marka/ürün koduna göre dışlama (arama+doğrudan kod ile ekleme)
  - ⭐ Öncelik Sıralamaları (mevcuttu, stok 0 olsa bile göster istisnası eklendi)
  - 🔀 Marka İçi Sıralama — ara-ekle mantığıyla markaya özel ürün sırası
  - 🏷️ Takma Ad — özel ürün ismi (orijinal isim korunur, altında küçük gösterilir)
  - 📋 Liste Güncelleme — GPD ve Parke için Excel/PDF ile toplu fiyat listesi güncelleme
  - 🔗 Eşleşen Stoklar — kutu/palet gibi farklı stok kartlarını gruplama (numune durumu paylaşımı)
  - ❓ Yardım — SSS + Alt+F1 kısayolu, tıklanabilir/yanıp sönen bağlantılar
  - 🔘 Varsayılan Tikler — dağınık tüm aç/kapa tiklerin toplu görünümü
  - 🗺️ Mağaza Bölümleri — EN SON eklenen büyük özellik: stant/bölüm fotoğrafı yükleyip
    üzerine tıklayarak ürün "pin"leme, canlı stok miktarı + isim/ölçü rozetleri,
    sürükle-taşı, yatay/dikey yön değiştirme, tam eşleşmeli arama.
- **Ana ekran (Liste) değişiklikleri:**
  - ALIŞ FİYATI sütunu artık net maliyet fiyatını (Fiyat Geçmişi'ndeki gerçek en yeni
    kayıttan, veri yüklenirken bir kez hesaplanıp `S.fiyatlar[kod].alisFiyatiGercek`'e
    yazılır) gösteriyor.
  - Arama: "Tam eşleşme" tiki, ölçü eşdeğeri renklendirmesi (yeşil=birebir, sarı=eşdeğer
    ör. 60x60↔60,5x60,5), ölçü başa alma ayarı (94/PVC-PPRC grubu hariç).
  - "0 ve eski stokları göster" artık varsayılan işaretli.
  - Ctrl+Space / Alt+T (KeyT ile Caps Lock bağımsız) kısayolları.
- **Hata düzeltmeleri:**
  - Dosyaya gömülü "11:17" kalıntı statik metin/duplicate sync-bar temizlendi.
  - Ayarlar kaydedilince localStorage önbelleği de güncelleniyor (yenileme titremesi azaldı).
  - Bağlantı koptuğunda tekrar denemede "Sunucuya bağlanılıyor" mesajının asılı kalması
    düzeltildi (retry başarılı olunca artık temizleniyor, gerçek hata metni gösteriliyor).
  - **KÖK NEDENİ BULUNAN AMA HENÜZ DÜZELTİLMEYEN HATA:** Backend'deki `sheetToObj`
    fonksiyonu artık `String(row[1])` ile değer sütununu metne çeviriyor (bu KISMEN
    düzeltildi — bkz. commit). Frontend'deki bazı `===\"1\"` / `!==\"0\"` katı karşılaştırmaları
    hâlâ gözden geçirilmemiş olabilir, ileride benzer "tik kayboluyor" şikayeti gelirse
    buradan başla.

## Bilinen/onaylanmamış konular
- Kullanıcı "GPD ürünleri silinmiş" dedi, sonra "silinmemiş, kontrol ettim" diye düzeltti —
  yanlış alarmmış, GPD_MARKA_KODU="20" olarak doğru duruyor, ekstra aksiyon gerekmedi.
- GPD_MARKA_KODU tarihte 89'dan 20'ye taşındı (yorum satırında not var) — deploy edilmemiş
  eski bir Apps Script sürümü varsa hâlâ 89 görünebilir, kullanıcı "89" dediğinde bunu göz
  önünde bulundur.

## Dosyalar
- `index.html` — tüm frontend (tek dosya, ~750KB)
- `Kod.js` — tüm backend (Apps Script kodu; eski adı `fincanlarstok_v3_fixed.gs`, artık clasp
  ile senkron olduğu için `Kod.js` adıyla repoda duruyor)
- `appsscript.json`, `.clasp.json` — clasp/Apps Script proje bağlantı dosyaları
- `.github/workflows/deploy.yml` — GitHub Actions otomasyonu (bkz. aşağıda)
- `lib/xlsx.full.min.js`, `lib/pdf.min.mjs`, `lib/pdf.worker.min.mjs` — yerel kütüphaneler
- `gpd_fiyat_listesi.csv` — orijinal referans veri (artık kullanılmıyor, Liste Güncelleme
  özelliği bunun yerini aldı)
- `sohbet_ozeti.md` — bu dosya, her oturum sonunda güncellenir

## 06.08.2026 — GitHub üzerinden tek noktadan otomasyon kuruldu
**ARTIK MANUEL "Deploy → New version" ADIMI YOK.** Repo: https://github.com/fincanlarnecip/stok-panel
(sahibi: `fincanlarnecip`, çalışma hesabı: `fincanlaryapi@gmail.com`)

- clasp (Google'ın resmi Apps Script CLI'ı) kuruldu, proje şu Script ID'ye bağlandı:
  `17AUz3HZF0RBGkKQkQcCcSLDtc44zLCmzS4GsQwvOJZT5BgOV-1ZYcxDl`
- `.clasprc.json` (clasp kimlik bilgisi) GitHub'da **`CLASPRC_JSON`** adlı repository secret
  olarak saklanıyor (Settings → Secrets and variables → Actions).
- `.github/workflows/deploy.yml` şunu yapıyor: `Kod.js`, `appsscript.json` veya `.clasp.json`
  dosyalarından biri `main` branch'e push edildiğinde otomatik olarak:
  1. `clasp push --force` (kodu Apps Script projesine gönderir)
  2. `clasp deploy` (yeni bir deployment/versiyon oluşturur, otomatik canlıya alır)
- **Yeni iş akışı:** `.gs`/backend kodunda değişiklik → GitHub'da `Kod.js`'i düzenle/yükle →
  commit et → Actions sekmesinde (https://github.com/fincanlarnecip/stok-panel/actions)
  birkaç saniye içinde yeşil tik görülür → değişiklik otomatik canlıda.
- ⚠️ **Dikkat:** GitHub'ın web düzenleyicisinde büyük dosyada (Kod.js ~70KB) elle satır
  silme/ekleme risklidir — bir kez yanlışlıkla parantez silinip "Unexpected end of input"
  syntax hatası çıktı, önceki çalışan commit'ten (raw URL üzerinden) kopyalayıp düzeltildi.
  Bundan sonra büyük değişiklikleri Claude'a yaptırıp tam dosyayı yapıştırmak, veya yerel
  editörde (VS Code vb.) düzenleyip öyle yüklemek daha güvenli.
- `index.html` değişikliği zaten eskisi gibi normal commit ile GitHub Pages'e yansıyor,
  ayrı bir adım gerekmiyor.

## Yeni sohbet başlatma yöntemi (GÜNCEL — Drive/zip artık kullanılmıyor)
Kullanıcı artık zip veya Drive yüklemek zorunda değil. Yeni sohbette şu yeterli:
> "GitHub'daki stok-panel reposunu çek, analiz et"
Claude, `raw.githubusercontent.com/fincanlarnecip/stok-panel/main/...` üzerinden tüm
dosyaları (index.html, Kod.js, sohbet_ozeti.md vb.) doğrudan okuyabiliyor (repo public).

Sohbet sonunda Claude bu dosyayı (sohbet_ozeti.md) günceller. Kullanıcı, Claude'a
**yazma izinli bir GitHub fine-grained personal access token** verdi (sadece bu repoya,
sadece "Contents: Read and write" izniyle) — böylece Claude artık dosyayı GitHub'a
**doğrudan kendisi commit edebiliyor**, kullanıcının manuel yükleme yapmasına gerek yok.
Token her yeni sohbette kullanıcı tarafından tekrar verilmesi gerekir (Claude'un
sohbetler arası hafızası yok) — token GitHub'da `github.com/settings/personal-access-tokens`
adresinden istenildiğinde iptal edilebilir (Revoke).
