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
- `fincanlarstok_v3_fixed.gs` — tüm backend
- `lib/xlsx.full.min.js`, `lib/pdf.min.mjs`, `lib/pdf.worker.min.mjs` — yerel kütüphaneler
- `gpd_fiyat_listesi.csv` — orijinal referans veri (artık kullanılmıyor, Liste Güncelleme
  özelliği bunun yerini aldı)
