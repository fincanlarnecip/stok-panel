// ════════════════════════════════════════════════
// STOK YÖNETİM SİSTEMİ — Google Apps Script v3 (DÜZELTİLMİŞ)
// fincanlaryapi@gmail.com
// Düzeltmeler:
//   - getFiyatlar(): tüm geçmiş kayıtları döndürür (fiyat geçmişi için)
//   - getAllData(): fiyatlar hem son fiyat hem geçmiş olarak gelir
// ════════════════════════════════════════════════

const SHEET_ID = "17-eyhwLd-3vIkH4HArhPnc3Ty7gkrZVYMERYJynXG4Q";

// ── Fatura fiyat verilerinin bulunduğu Sheets (fincanlaryapi hesabı)
const FIYAT_SHEET_ID  = "19t4MsvudC8X7knZ_dymBm5fghcbZcpAMwOmUXZxDPPQ";
const FIYAT_SHEET_ADI = "FATURAFIYAT";

const SHEETS = {
  notlar:    "Notlar",
  etiketler: "Etiketler",
  urunEt:    "UrunEtiketleri",
  markalar:  "Markalar",
  ayarlar:   "Ayarlar",
  stoklar:   "Stoklar",
  log:       "Log",
  numuneler:      "Numuneler",
  numuneGruplari: "NumuneGruplari",
  teklifler:      "Teklifler",
  urunGorselleri: "UrunGorselleri",
  magazaBolumleri: "MagazaBolumleri",
  magazaPinler:    "MagazaPinler",
  gpdListesi:     "GPD_Fiyat_Listesi",
  parkeListesi:   "PARKE_Fiyat_Listesi",
};

// ── GPD 2026 Fiyat Listesi ──
// Mail ile gelen stok güncelleme dosyasından TAMAMEN BAĞIMSIZ, ayrı bir sayfada (GPD_Fiyat_Listesi)
// tutulur — Stoklar sayfası her mail güncellemesinde silinip yeniden yazıldığı için GPD ürünleri
// oraya karışmaz. Panelde marka olarak "GPD" görünmesi için, mevcut marka sistemi stok kodunun
// İLK 2 RAKAMINA göre çalıştığından (bkz. index.html getBrand()), GPD ürünlerine bu 2 haneli kod
// önek olarak eklenir (ör. "20-FLB09"). Markalar sayfasında bu kod yoksa otomatik eklenir.
const GPD_MARKA_KODU = "20";
const GPD_MARKA_ADI  = "GPD";
const GPD_MARKA_RENK = "#7c3aed";

// ── Parke Fiyat Listesi ──
// GPD ile birebir aynı mantık/yapı (bkz. yukarıdaki not) — Ayarlar > Liste Güncelleme sekmesinden
// Excel/PDF yükleyerek doldurulur. NOT: Bu, Hesaplama Modülündeki (Seramik+Parke m² hesaplayıcı)
// "parke" ile tamamen ayrı ve ilgisiz bir özelliktir — sadece fiyat listesi verisidir.
const PARKE_MARKA_KODU = "21";
const PARKE_MARKA_ADI  = "Parke";
const PARKE_MARKA_RENK = "#a16207";

// ════════════════════════════════════════════════
// YARDIMCI FONKSİYONLAR
// ════════════════════════════════════════════════

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e8edf5");
  }
  return sheet;
}

function sheetToObj(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const result = {};
  // NOT: Değer sütunu (row[1]) burada mutlaka String()'e çevrilmeli. Google Sheets, hücreye
  // "1" gibi rakamsal görünen bir metin yazıldığında bunu otomatik olarak SAYI olarak saklıyor.
  // Bu değer frontend'de "===\"1\"" gibi katı (tip duyarlı) karşılaştırmalarla kontrol edildiği
  // için, sayı olarak gelirse hiçbir zaman eşleşmiyor ve ayar her zaman "kapalı" görünüyordu.
  data.slice(1).forEach(row => { if (row[0]) result[String(row[0])] = String(row[1]); });
  return result;
}

function sheetToArr(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1)
    .map(row => ({ t: row[0], c: row[1] || "#d97706" }))
    .filter(e => e.t);
}

function sheetToRows(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues().slice(1)
    .map(row => ({ k: String(row[0]), a: row[1], c: row[2] || "#2563eb" }))
    .filter(m => m.k);
  // Sheette (eski hatalardan kalma) aynı KOD'a sahip tekrar eden satırlar olsa bile,
  // istemci HİÇBİR ZAMAN duplikasyon görmesin diye burada okuma anında tekilleştiriyoruz.
  // Sheetin kendisi hâlâ şişebilir, o yüzden ayrıca zamanlı bir temizlik tetikleyicisi kullanılıyor.
  const gorulen = new Set();
  return rows.filter(m => {
    if (gorulen.has(m.k)) return false;
    gorulen.add(m.k);
    return true;
  });
}

function logEntry(tip, detay, tarih) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEETS.log, ["TARİH", "TİP", "DETAY"]);
  sheet.appendRow([tarih || new Date(), tip, detay]);
}

function logError(err) {
  try { logEntry("HATA", err.message, new Date()); } catch(e) {}
}

function parseDateFromFilename(name) {
  let m = name.match(/(\d{2})[._\-](\d{2})[._\-](\d{4})/);
  if (m) return m[3] + "-" + m[2].padStart(2,"0") + "-" + m[1].padStart(2,"0");
  m = name.match(/(\d{4})[._\-](\d{2})[._\-](\d{2})/);
  if (m) return m[1] + "-" + m[2] + "-" + m[3];
  m = name.match(/(\d{2})[._\-](\d{2})/);
  if (m) return new Date().getFullYear() + "-" + m[2].padStart(2,"0") + "-" + m[1].padStart(2,"0");
  return new Date().toISOString().split("T")[0];
}

function findColIndex(headers, candidates) {
  for (const c of candidates) {
    const idx = headers.indexOf(c);
    if (idx >= 0) return idx;
  }
  return -1;
}

// ════════════════════════════════════════════════
// FİYAT VERİLERİNİ OKU — DÜZELTİLDİ
// Son fiyat map'i + tüm geçmiş kayıtları döndürür
// ════════════════════════════════════════════════

// ════════════════════════════════════════════════
// GPD 2026 FİYAT LİSTESİ — ayrı sayfa, mail güncellemesinden bağımsız
// Sayfa düzeni: A1 = Alış İskontomuz (%), 2. satır başlıklar, 3. satırdan itibaren veri.
// Başlıklar: STOK_KODU | STOK_ADI | KOLI_ADEDI | YUZEY | BRUT_FIYAT
// Alış Fiyatımız = BRUT_FIYAT × (1 − A1/100) — her okumada canlı hesaplanır, sayfaya yazılmaz.
// ════════════════════════════════════════════════

function gpdMarkayiKaydet(ss) {
  const sheet = getOrCreateSheet(ss, SHEETS.markalar, ["KOD","AD","RENK"]);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === GPD_MARKA_KODU) return; // zaten güncel kodla kayıtlı
    // Eski koddan (ör. 89) yeni koda (20) geçiş: aynı isimde (GPD) ama eski/farklı kodlu bir satır
    // bulunursa, yeni bir satır eklemek yerine SADECE o satırın kodunu günceller — iki ayrı "GPD"
    // markası oluşmasın diye.
    if (String(data[i][1]) === GPD_MARKA_ADI) {
      sheet.getRange(i + 1, 1).setValue(GPD_MARKA_KODU);
      return;
    }
  }
  sheet.appendRow([GPD_MARKA_KODU, GPD_MARKA_ADI, GPD_MARKA_RENK]);
}

function getGpdUrunleri(ss) {
  const sheet = ss.getSheetByName(SHEETS.gpdListesi);
  if (!sheet) return { iskonto: 0, urunler: [], fiyatlar: {} };

  const data = sheet.getDataRange().getValues();
  const iskonto = parseFloat(data[0] && data[0][0]) || 0;
  if (data.length < 3) return { iskonto, urunler: [], fiyatlar: {} };

  // 2. satır (index 1) başlık, 3. satırdan (index 2) itibaren veri
  const headers = (data[1] || []).map(h => String(h).toUpperCase().trim().replace(/ /g, "_"));
  const colKod  = findColIndex(headers, ["STOK_KODU","ÜRÜN_KODU","URUN_KODU","KOD"]);
  const colAd   = findColIndex(headers, ["STOK_ADI","ÜRÜN_ADI","URUN_ADI","AD"]);
  const colKoli = findColIndex(headers, ["KOLI_ADEDI","KOLİ_ADEDİ"]);
  const colYuz  = findColIndex(headers, ["YUZEY","YÜZEY"]);
  const colBrut = findColIndex(headers, ["BRUT_FIYAT","BRÜT_FIYAT","BRUT_FİYAT"]);

  const urunler = [];
  const fiyatlar = {};

  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    const hamKod = colKod >= 0 ? String(row[colKod] || "").trim() : "";
    if (!hamKod) continue;
    const brutFiyat = colBrut >= 0 ? (parseFloat(row[colBrut]) || 0) : 0;
    const alisFiyati = Math.round(brutFiyat * (1 - iskonto / 100) * 100) / 100;
    // Panelin marka sistemi stok kodunun ilk 2 rakamına göre çalıştığı için GPD kodu önek olarak eklenir.
    const stokKodu = GPD_MARKA_KODU + "-" + hamKod;

    urunler.push({
      STOK_KODU:   stokKodu,
      STOK_ADI:    (colAd >= 0 ? String(row[colAd] || "").trim() : "") +
                   (colYuz >= 0 && row[colYuz] ? " (" + String(row[colYuz]).trim() + ")" : ""),
      MERKEZ_DEPO: "", AYRILMIS: "", SATILABILIR: "",
      MARKA: GPD_MARKA_ADI,
      KOLI_ADEDI:  colKoli >= 0 ? (parseFloat(row[colKoli]) || 0) : 0,
      YUZEY:       colYuz  >= 0 ? String(row[colYuz] || "") : "",
      BRUT_FIYAT:  brutFiyat,
      ALIS_FIYATI: alisFiyati,
    });

    // Mevcut FİYAT sütununun (Malzeme Listesi vb. her yerde) çalışması için fiyatlar map'ine
    // de ekleniyor — birimFiyat = Liste (Brüt) Fiyatı, maliyetFiyat/netFiyat = bizim Alış Fiyatımız.
    fiyatlar[stokKodu] = {
      birimFiyat:   brutFiyat,
      iskonto:      iskonto,
      netFiyat:     alisFiyati,
      nakliyePayi:  0,
      maliyetFiyat: alisFiyati,
      kdvOrani:     20,
      faturaNo:     "",
      fatTarih:     "",
      tedarikci:    "GPD",
      faturaLink:   "",
      edmLink:      "",
    };
  }

  return { iskonto, urunler, fiyatlar };
}

function parkeMarkayiKaydet(ss) {
  const sheet = getOrCreateSheet(ss, SHEETS.markalar, ["KOD","AD","RENK"]);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === PARKE_MARKA_KODU) return;
    if (String(data[i][1]) === PARKE_MARKA_ADI) {
      sheet.getRange(i + 1, 1).setValue(PARKE_MARKA_KODU);
      return;
    }
  }
  sheet.appendRow([PARKE_MARKA_KODU, PARKE_MARKA_ADI, PARKE_MARKA_RENK]);
}

function getParkeUrunleri(ss) {
  const sheet = ss.getSheetByName(SHEETS.parkeListesi);
  if (!sheet) return { iskonto: 0, urunler: [], fiyatlar: {} };

  const data = sheet.getDataRange().getValues();
  const iskonto = parseFloat(data[0] && data[0][0]) || 0;
  if (data.length < 3) return { iskonto, urunler: [], fiyatlar: {} };

  const headers = (data[1] || []).map(h => String(h).toUpperCase().trim().replace(/ /g, "_"));
  const colKod  = findColIndex(headers, ["STOK_KODU","ÜRÜN_KODU","URUN_KODU","KOD"]);
  const colAd   = findColIndex(headers, ["STOK_ADI","ÜRÜN_ADI","URUN_ADI","AD"]);
  const colKoli = findColIndex(headers, ["KOLI_ADEDI","KOLİ_ADEDİ"]);
  const colYuz  = findColIndex(headers, ["YUZEY","YÜZEY"]);
  const colBrut = findColIndex(headers, ["BRUT_FIYAT","BRÜT_FIYAT","BRUT_FİYAT"]);

  const urunler = [];
  const fiyatlar = {};

  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    const hamKod = colKod >= 0 ? String(row[colKod] || "").trim() : "";
    if (!hamKod) continue;
    const brutFiyat = colBrut >= 0 ? (parseFloat(row[colBrut]) || 0) : 0;
    const alisFiyati = Math.round(brutFiyat * (1 - iskonto / 100) * 100) / 100;
    const stokKodu = PARKE_MARKA_KODU + "-" + hamKod;

    urunler.push({
      STOK_KODU:   stokKodu,
      STOK_ADI:    (colAd >= 0 ? String(row[colAd] || "").trim() : "") +
                   (colYuz >= 0 && row[colYuz] ? " (" + String(row[colYuz]).trim() + ")" : ""),
      MERKEZ_DEPO: "", AYRILMIS: "", SATILABILIR: "",
      MARKA: PARKE_MARKA_ADI,
      KOLI_ADEDI:  colKoli >= 0 ? (parseFloat(row[colKoli]) || 0) : 0,
      YUZEY:       colYuz  >= 0 ? String(row[colYuz] || "") : "",
      BRUT_FIYAT:  brutFiyat,
      ALIS_FIYATI: alisFiyati,
    });

    fiyatlar[stokKodu] = {
      birimFiyat:   brutFiyat,
      iskonto:      iskonto,
      netFiyat:     alisFiyati,
      nakliyePayi:  0,
      maliyetFiyat: alisFiyati,
      kdvOrani:     20,
      faturaNo:     "",
      fatTarih:     "",
      tedarikci:    "Parke",
      faturaLink:   "",
      edmLink:      "",
    };
  }

  return { iskonto, urunler, fiyatlar };
}

// ── Liste Güncelleme (GPD / Parke ortak) ──
// Ayarlar > Liste Güncelleme sekmesinden Excel/PDF yükleyip önizlemede onaylanan satırları
// buraya gönderir. Sayfa tamamen temizlenip (2 başlık satırı korunarak) yeniden yazılır.
function listeGuncelle(body) {
  const hangi = String(body.liste || "").toLowerCase(); // "gpd" | "parke"
  if (hangi !== "gpd" && hangi !== "parke") return { error: "Geçersiz liste türü" };

  const sheetAdi = hangi === "gpd" ? SHEETS.gpdListesi : SHEETS.parkeListesi;
  const iskonto  = parseFloat(body.iskonto) || 0;
  let rows = [];
  try { rows = JSON.parse(body.rows || "[]"); } catch (e) { rows = []; }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, sheetAdi, [""]);
  sheet.clear();

  sheet.getRange(1, 1).setValue(iskonto);
  sheet.getRange(2, 1, 1, 5).setValues([["STOK_KODU","STOK_ADI","YUZEY","KOLI_ADEDI","BRUT_FIYAT"]]);

  if (rows.length > 0) {
    const out = rows.map(r => [
      String(r.kod || ""),
      String(r.ad || ""),
      String(r.yuzey || ""),
      parseFloat(r.koliAdedi) || 0,
      parseFloat(r.brutFiyat) || 0,
    ]);
    sheet.getRange(3, 1, out.length, 5).setValues(out);
  }

  if (hangi === "gpd") gpdMarkayiKaydet(ss);
  else parkeMarkayiKaydet(ss);

  return { ok: true, satirSayisi: rows.length };
}

function getFiyatlar() {
  const fiyatMap    = {};  // stok_kodu → en son fiyat objesi
  const gecmisMap   = {};  // stok_kodu → [tüm fiyat kayıtları]

  try {
    const ss = SpreadsheetApp.openById(FIYAT_SHEET_ID);
    const sh = ss.getSheetByName(FIYAT_SHEET_ADI);
    if (!sh) return { fiyatMap, gecmisMap };

    const data = sh.getDataRange().getValues();
    if (data.length < 2) return { fiyatMap, gecmisMap };

    const h = data[0];
    const colKod  = h.indexOf("STOK_KODU");
    const colFiy  = h.indexOf("BIRIM_FIYAT");
    const colIsk  = h.indexOf("ISKONTO");
    const colNet  = h.indexOf("NET_FIYAT");
    const colNav  = h.indexOf("NAKLIYE_PAYI");
    const colMal  = h.indexOf("MALIYET_FIYAT");
    const colKdv  = h.indexOf("KDV_ORANI");
    const colFNo  = h.indexOf("FATURA_NO");
    const colFTar = h.indexOf("FATURA_TARIHI");
    const colTed  = h.indexOf("TEDARIKCI");
    const colLnk  = h.indexOf("FATURA_LINK");
    const colEdm  = h.indexOf("EDM_LINK");

    // Satırlar en yeniden eskiye sıralı (insertRowAfter(1) sayesinde)
    data.slice(1).forEach(row => {
      const kod = String(row[colKod] || "").trim();
      if (!kod) return;

      const kayit = {
        birimFiyat:   parseFloat(row[colFiy]) || 0,
        iskonto:      parseFloat(row[colIsk]) || 0,
        netFiyat:     colNet >= 0 ? parseFloat(row[colNet]) || 0 : 0,
        nakliyePayi:  colNav >= 0 ? parseFloat(row[colNav]) || 0 : 0,
        maliyetFiyat: colMal >= 0 ? parseFloat(row[colMal]) || 0 : 0,
        kdvOrani:     parseFloat(row[colKdv]) || 0,
        faturaNo:     String(row[colFNo]  || ""),
        fatTarih:     String(row[colFTar] || ""),
        tedarikci:    String(row[colTed]  || ""),
        faturaLink:   colLnk >= 0 ? String(row[colLnk] || "") : "",
        edmLink:      colEdm >= 0 ? String(row[colEdm] || "") : "",
      };

      // En son fiyat: ilk karşılaşılan kayıt (en üstte = en yeni)
      if (!fiyatMap[kod]) {
        fiyatMap[kod] = kayit;
      }

      // Geçmiş: tüm kayıtlar
      if (!gecmisMap[kod]) gecmisMap[kod] = [];
      gecmisMap[kod].push(kayit);
    });

    Logger.log("Fiyat verisi okundu: " + Object.keys(fiyatMap).length + " ürün, " +
      Object.values(gecmisMap).reduce((t, a) => t + a.length, 0) + " toplam kayıt");

  } catch(e) {
    Logger.log("Fiyat okuma hatası: " + e.message);
  }

  return { fiyatMap, gecmisMap };
}

// ════════════════════════════════════════════════
// EXCEL → STOK VERİSİ
// ════════════════════════════════════════════════

function parseExcelFile(excelFile) {
  const token    = ScriptApp.getOAuthToken();
  const boundary = "stok_boundary_xyz";

  const metaJson = JSON.stringify({
    name: "temp_stok_" + Date.now(),
    mimeType: "application/vnd.google-apps.spreadsheet"
  });

  const part1 = "--" + boundary + "\r\n" +
    "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
    metaJson + "\r\n" +
    "--" + boundary + "\r\n" +
    "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n";

  const part3 = "\r\n--" + boundary + "--";

  const allBytes = Utilities.newBlob(part1).getBytes()
    .concat(excelFile.getBlob().getBytes())
    .concat(Utilities.newBlob(part3).getBytes());

  const response = UrlFetchApp.fetch(
    "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
    {
      method: "POST",
      headers: {
        "Authorization": "Bearer " + token,
        "Content-Type": "multipart/related; boundary=" + boundary
      },
      payload: allBytes,
      muteHttpExceptions: true
    }
  );

  const result = JSON.parse(response.getContentText());
  if (!result.id) {
    throw new Error("Excel→Sheets dönüşüm hatası: " + response.getContentText().substring(0, 200));
  }

  const tempSS = SpreadsheetApp.openById(result.id);
  const sheet  = tempSS.getSheets()[0];
  const data   = sheet.getDataRange().getValues();

  try { DriveApp.getFileById(result.id).setTrashed(true); } catch(e) {}

  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h).toUpperCase().trim().replace(/ /g, "_"));

  const colMap = {
    STOK_KODU:   findColIndex(headers, ["STOK_KODU","STOK KODU","KOD","STOKKODU"]),
    STOK_ADI:    findColIndex(headers, ["STOK_ADI","STOK ADI","AD","STOKADI","AÇIKLAMA"]),
    MERKEZ_DEPO: findColIndex(headers, ["MERKEZ_DEPO","MERKEZ DEPO","DEPO","MERKEZ"]),
    AYRILMIS:    findColIndex(headers, ["AYRILMIS","AYRILAN","AYRILMIŞ","REZERVE"]),
    SATILABILIR: findColIndex(headers, ["SATILABILIR","SATILABİLİR","NET","KULLANILABILIR"]),
  };

  return data.slice(1).map(row => {
    const kod = colMap.STOK_KODU >= 0 ? String(row[colMap.STOK_KODU] || "").trim() : "";
    if (!kod) return null;
    return {
      STOK_KODU:   kod,
      STOK_ADI:    colMap.STOK_ADI    >= 0 ? String(row[colMap.STOK_ADI]    || "").trim() : "",
      MERKEZ_DEPO: colMap.MERKEZ_DEPO >= 0 ? parseFloat(row[colMap.MERKEZ_DEPO] || 0) || 0 : 0,
      AYRILMIS:    colMap.AYRILMIS    >= 0 ? parseFloat(row[colMap.AYRILMIS]    || 0) || 0 : 0,
      SATILABILIR: colMap.SATILABILIR >= 0 ? parseFloat(row[colMap.SATILABILIR] || 0) || 0 : 0,
    };
  }).filter(r => r && r.STOK_KODU);
}

// ════════════════════════════════════════════════
// STOK KAYDET
// ════════════════════════════════════════════════

// ── HAFİF "SON GÜNCELLEME" ZAMAN DAMGASI ──
// Panelin arka planda sürekli sorgulayabileceği, TÜM veriyi değil sadece tek bir zaman
// bilgisini döndüren hafif bir uç nokta. Böylece "yeni stok geldi mi" kontrolü, tüm
// Stoklar/Markalar/Numuneler sayfalarını okumak zorunda kalmadan çok hızlı yapılabilir.
function sonGuncellemeZamaniniYaz() {
  PropertiesService.getScriptProperties().setProperty("sonStokGuncelleme", new Date().toISOString());
}
function sonGuncellemeZamaniniGetir() {
  const v = PropertiesService.getScriptProperties().getProperty("sonStokGuncelleme");
  return { sonGuncelleme: v || null };
}

function saveStoklar(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEETS.stoklar,
    ["STOK_KODU","STOK_ADI","MERKEZ_DEPO","AYRILMIS","SATILABILIR","TARIH"]);

  const MAX_YUKLEME = 7; // Son 7 yükleme saklanır (20'den düşürüldü — sheet küçük kaldıkça getAllData çok daha hızlı okunur)
  const tarih    = body.tarih || new Date().toISOString().split("T")[0];
  const yeniList = body.stoklar || [];

  // Mevcut verideki benzersiz tarihleri bul
  const mevcutData = sheet.getDataRange().getValues();
  const tarihSet = new Set();
  const tarihIdx = 5; // TARIH kolonu
  if (mevcutData.length > 1) {
    mevcutData.slice(1).forEach(row => {
      const t = String(row[tarihIdx] || "").split("T")[0].trim();
      if (t && !t.includes("pasif")) tarihSet.add(t);
    });
  }

  // Yeni tarih zaten varsa üzerine yaz (aynı gün tekrar yükleme)
  // Yoksa ekle ve limit kontrolü yap
  const tumTarihler = [...tarihSet].sort();
  if (!tarihSet.has(tarih)) {
    tumTarihler.push(tarih);
    // 20'den fazla yükleme varsa en eskiyi sil
    while (tumTarihler.length > MAX_YUKLEME) {
      tumTarihler.shift();
    }
  }

  // Yeni tarihin satırlarını oluştur
  const yeniSatirlar = yeniList.map(r => [
    r.STOK_KODU, r.STOK_ADI, r.MERKEZ_DEPO, r.AYRILMIS, r.SATILABILIR, tarih
  ]);

  // Mevcut veriden bu tarihe ait satırları çıkar, kalan tarihleri koru
  const korunacak = mevcutData.length > 1
    ? mevcutData.slice(1).filter(row => {
        const t = String(row[tarihIdx] || "").split("T")[0].trim();
        return tumTarihler.includes(t) && t !== tarih;
      })
    : [];

  // Sheet'i yeniden yaz
  sheet.clearContents();
  sheet.appendRow(["STOK_KODU","STOK_ADI","MERKEZ_DEPO","AYRILMIS","SATILABILIR","TARIH"]);
  const tumSatirlar = [...korunacak, ...yeniSatirlar];
  if (tumSatirlar.length > 0) {
    sheet.getRange(2, 1, tumSatirlar.length, 6).setValues(tumSatirlar);
  }

  Logger.log("Stok kaydedildi: " + tarih + " → " + yeniList.length + " ürün, toplam " + tumSatirlar.length + " satır, " + tumTarihler.length + " yükleme");
  sonGuncellemeZamaniniYaz();
  return { ok: true, count: yeniList.length, tarih: tarih, yuklemeSayisi: tumTarihler.length };
}

// ════════════════════════════════════════════════
// MAİL OTOMASYONU
// ════════════════════════════════════════════════

function checkStokMail() {
  // ÖNEMLİ: Eskiden "işlendi" durumu Gmail THREAD'i (konuşma) üzerinden takip ediliyordu.
  // Aynı konu başlığıyla her gün gelen bir mail Gmail'de TEK bir thread'de birleşiyorsa,
  // ilk gün işlenip etiketlenince tüm thread "işlendi" sayılıyor ve sonraki günlerin
  // yeni mesajları sisteme HİÇ BAKILMADAN atlanıyordu (log'da hiçbir iz bile kalmıyordu).
  // Artık her mesajın kendi benzersiz kimliği (getId()) ayrı bir tabloda tutuluyor;
  // aynı thread'de olsalar bile her yeni mesaj kendi başına değerlendiriliyor.
  const query   = "to:fincanlaryapi@gmail.com has:attachment newer_than:7d";
  const threads = GmailApp.search(query);

  let islenenMesajSayisi = 0, toplamUrunSayisi = 0;

  if (threads.length === 0) {
    Logger.log("Yeni stok maili yok.");
    return { islenen: 0, urunSayisi: 0 };
  }

  let label = GmailApp.getUserLabelByName("stok-islendi");
  if (!label) label = GmailApp.createLabel("stok-islendi");

  const islenenler = islenenMesajIdleriniGetir();
  const stokFolder = getOrCreateFolder("STOK");

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const msgId = msg.getId();
      if (islenenler.has(msgId)) return; // bu MESAJ (thread değil) daha önce işlendi

      let basarili = false, hataVarMi = false;
      msg.getAttachments().forEach(att => {
        const name = att.getName();
        if (!name.match(/\.(xlsx?|csv)$/i)) return;

        const existing = stokFolder.getFilesByName(name);
        while (existing.hasNext()) existing.next().setTrashed(true);
        const savedFile = stokFolder.createFile(att);

        try {
          const stoklar = parseExcelFile(savedFile);
          if (stoklar.length > 0) {
            const tarih = parseDateFromFilename(name);
            saveStoklar({ stoklar: stoklar, tarih: tarih });
            logEntry("otomatik", name + " → " + stoklar.length + " ürün", new Date());
            basarili = true;
            toplamUrunSayisi += stoklar.length;
          } else {
            hataVarMi = true;
            logEntry("hata", name + ": dosyadan hiç ürün okunamadı (0 satır — başlık/kolon eşleşmesi kontrol edilmeli)", new Date());
          }
        } catch(err) {
          hataVarMi = true;
          logEntry("hata", name + ": " + err.message, new Date());
        }
      });
      // Sadece gerçekten başarılıysa (ya da işlenecek uygun ek hiç yoksa) bu MESAJI işlendi say;
      // hata varsa işaretlenmez, bir sonraki kontrolde tekrar denenir.
      if (basarili || !hataVarMi) {
        mesajIslenmisIsaretle(msgId);
        thread.addLabel(label); // Gmail'de görsel referans için (artık mantığın parçası değil)
        if (basarili) islenenMesajSayisi++;
      }
    });
  });

  return { islenen: islenenMesajSayisi, urunSayisi: toplamUrunSayisi };
}

// Manuel "Maili Şimdi Kontrol Et" butonu tarafından çağrılır. checkStokMail() ile aynı işi yapar,
// tek farkı doPost/doGet üzerinden senkron olarak çağrılıp sonucu doğrudan istemciye döndürmesi.
function checkMailNow() {
  try {
    return checkStokMail();
  } catch (err) {
    return { error: err.message };
  }
}

// Hangi Gmail mesajlarının (thread değil, tekil mesaj) daha önce başarıyla işlendiğini
// ayrı bir sayfada (İşlenenMailler) takip eder. Sadece MESAJ_ID ve TARIH tutulur.
function islenenMesajIdleriniGetir() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, "İşlenenMailler", ["MESAJ_ID","TARIH"]);
  const data = sheet.getDataRange().getValues();
  const set = new Set();
  for (let i = 1; i < data.length; i++) set.add(String(data[i][0]));
  return set;
}

function mesajIslenmisIsaretle(msgId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, "İşlenenMailler", ["MESAJ_ID","TARIH"]);
  sheet.appendRow([msgId, new Date()]);
  // Tablo sınırsız büyümesin diye 60 günden eski kayıtları temizle (arama zaten newer_than:7d ile sınırlı).
  const data = sheet.getDataRange().getValues();
  if (data.length > 500) {
    const sinirTarihi = new Date(Date.now() - 60*24*60*60*1000);
    const kalanlar = [data[0]];
    for (let i = 1; i < data.length; i++) {
      const tarih = data[i][1] instanceof Date ? data[i][1] : new Date(data[i][1]);
      if (tarih > sinirTarihi) kalanlar.push(data[i]);
    }
    if (kalanlar.length < data.length) {
      sheet.clearContents();
      sheet.getRange(1, 1, kalanlar.length, 2).setValues(kalanlar);
    }
  }
}

// ════════════════════════════════════════════════
// TETİKLEYİCİ
// ════════════════════════════════════════════════

function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("checkStokMail").timeBased().everyMinutes(5).create();
  Logger.log("✅ Tetikleyici kuruldu (her 5 dakikada bir).");
}

// ════════════════════════════════════════════════
// WEB API
// ════════════════════════════════════════════════

function doGet(e) {
  if (e.parameter && e.parameter.payload) {
    try {
      const parsed = JSON.parse(decodeURIComponent(e.parameter.payload));
      e = Object.assign({}, e, { postData: { contents: JSON.stringify(parsed) } });
    } catch(err) {}
  }
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    const body   = e.postData ? JSON.parse(e.postData.contents) : e.parameter;
    const action = body.action;
    let result;
    switch (action) {
      case "getAll":      result = getAllData(body.tarih);  break;
      case "sonGuncelleme": result = sonGuncellemeZamaniniGetir(); break;
      case "checkMailNow": result = checkMailNow(); break;
      case "saveNot":     result = saveNot(body);       break;
      case "saveEtiket":  result = saveEtiket(body);    break;
      case "saveUrunEt":  result = saveUrunEt(body);    break;
      case "urunEtGuncelleTek": result = urunEtGuncelleTek(body); break;
      case "etiketTamSil": result = etiketTamSil(body); break;
      case "saveMarka":   result = saveMarka(body);     break;
      case "markaEkleTek":      result = markaEkleTek(body);      break;
      case "markaSilTek":       result = markaSilTek(body);       break;
      case "markaGuncelleTek":  result = markaGuncelleTek(body);  break;
      case "markaSiraDegistir": result = markaSiraDegistir(body); break;
      case "saveAyar":       result = saveAyar(body);          break;
      case "listeGuncelle":  result = listeGuncelle(body);     break;
      case "saveMagazaBolum":  result = saveMagazaBolum(body);   break;
      case "silMagazaBolum":   result = silMagazaBolum(body);    break;
      case "magazaPinKaydet":  result = magazaPinKaydet(body);   break;
      case "magazaPinSil":     result = magazaPinSil(body);      break;
      case "saveNumuneler":  result = saveNumuneler(body);     break;
      case "saveNumuneGruplari": result = saveNumuneGruplari(body); break;
      case "numuneEkleTek":      result = numuneEkleTek(body);      break;
      case "numuneSilTek":       result = numuneSilTek(body);       break;
      case "numuneGuncelleTek":  result = numuneGuncelleTek(body);  break;
      case "numuneSiraDegistir": result = numuneSiraDegistir(body); break;
      case "numuneGrupSilTek":   result = numuneGrupSilTek(body);   break;
      case "numuneGrupEkleTek":  result = numuneGrupEkleTek(body);  break;
      case "saveTeklifler": result = saveTeklifler(body); break;
      case "saveStoklar": result = saveStoklar(body);   break;
      case "saveUrunGorsel": result = saveUrunGorsel(body); break;
      case "silUrunGorsel":  result = silUrunGorsel(body);  break;
      case "guncelleGorselEtiket": result = guncelleGorselEtiket(body); break;
      default:            result = { error: "Bilinmeyen: " + action };
    }
    return jsonResponse(result);
  } catch(err) {
    logError(err);
    return jsonResponse({ error: err.message });
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════
// VERİ FONKSİYONLARI
// ════════════════════════════════════════════════

function tarihStr(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Europe/Istanbul", "yyyy-MM-dd");
  }
  const s = String(val || "");
  const m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[0];
  const d = new Date(s);
  if (!isNaN(d)) return Utilities.formatDate(d, "Europe/Istanbul", "yyyy-MM-dd");
  return "";
}

function getAllData(seciliTarih) {
  const _t0 = Date.now();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  Logger.log("⏱ [1] Spreadsheet açıldı: " + (Date.now()-_t0) + "ms");

  const stokSheet = ss.getSheetByName(SHEETS.stoklar);
  let stoklar     = [];
  let stokTarih   = "";
  let stokTarihler = [];

  if (stokSheet) {
    const _t1 = Date.now();
    const data = stokSheet.getDataRange().getValues();
    Logger.log("⏱ [2] Stoklar okundu (" + (data.length-1) + " satır): " + (Date.now()-_t1) + "ms");
    if (data.length > 1) {
      const headers = data[0].map(h => String(h).toUpperCase().trim());
      const tarihIdx = headers.indexOf("TARIH");

      // TEK GEÇİŞTE: hem benzersiz tarihleri topla hem satırları tarihe göre grupla
      // (öncesinde veri 2 kez taranıyordu — tarih toplama + filtreleme — artık tek seferde)
      const tarihSet = new Set();
      const gruplar = new Map(); // tarihStr -> o tarihe ait ham satırlar
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const t = tarihStr(row[tarihIdx]);
        if (!t || t.includes("pasif")) continue;
        tarihSet.add(t);
        if (!gruplar.has(t)) gruplar.set(t, []);
        gruplar.get(t).push(row);
      }
      stokTarihler = [...tarihSet].sort().reverse();

      // En son tarih veya seçili tarih
      stokTarih = seciliTarih && tarihSet.has(seciliTarih)
        ? seciliTarih
        : (stokTarihler[0] || "");

      // Sadece hedef tarihin satırlarını (zaten gruplanmış) nesneye çevir
      stoklar = (gruplar.get(stokTarih) || [])
        .map(row => {
          const obj = {};
          headers.forEach((h, i) => obj[h] = row[i]);
          return obj;
        })
        .filter(r => r.STOK_KODU);
      Logger.log("⏱ [3] Stoklar filtrelendi (" + stoklar.length + " ürün, " + stokTarihler.length + " tarih): " + (Date.now()-_t1) + "ms (toplam)");
    }
  }

  // ── Fiyat verilerini oku — DÜZELTİLDİ: hem son fiyat hem geçmiş ──
  const _t2 = Date.now();
  let fiyatlar  = {};
  let fiyatGecmis = {};
  try {
    const sonuc = getFiyatlar();
    fiyatlar    = sonuc.fiyatMap;
    fiyatGecmis = sonuc.gecmisMap;
  } catch(e) {
    Logger.log("Fiyat okuma atlandı: " + e.message);
  }
  Logger.log("⏱ [4] Fiyatlar okundu (ayrı spreadsheet): " + (Date.now()-_t2) + "ms");

  // ── GPD 2026 Fiyat Listesi — ayrı sayfadan okunur, Stoklar'a hiç yazılmaz (mail güncellemesi
  // bunu asla silmez/değiştirmez). Ürünler listeye, fiyatları fiyatlar map'ine eklenir.
  const _t2b = Date.now();
  try {
    gpdMarkayiKaydet(ss);
    const gpd = getGpdUrunleri(ss);
    if (gpd.urunler.length) {
      stoklar = stoklar.concat(gpd.urunler);
      Object.assign(fiyatlar, gpd.fiyatlar);
    }
    Logger.log("⏱ [4b] GPD listesi okundu (" + gpd.urunler.length + " ürün): " + (Date.now()-_t2b) + "ms");
  } catch(e) {
    Logger.log("GPD listesi okuma atlandı: " + e.message);
  }

  // ── Parke Fiyat Listesi — GPD ile aynı mantık, ayrı sayfa/marka. ──
  const _t2c = Date.now();
  try {
    parkeMarkayiKaydet(ss);
    const parke = getParkeUrunleri(ss);
    if (parke.urunler.length) {
      stoklar = stoklar.concat(parke.urunler);
      Object.assign(fiyatlar, parke.fiyatlar);
    }
    Logger.log("⏱ [4c] Parke listesi okundu (" + parke.urunler.length + " ürün): " + (Date.now()-_t2c) + "ms");
  } catch(e) {
    Logger.log("Parke listesi okuma atlandı: " + e.message);
  }

  const _t3 = Date.now();
  const notlar = sheetToObj(ss, SHEETS.notlar);
  Logger.log("⏱ [5] Notlar okundu: " + (Date.now()-_t3) + "ms");

  const _t4 = Date.now();
  const etiketler = sheetToArr(ss, SHEETS.etiketler);
  Logger.log("⏱ [6] Etiketler okundu: " + (Date.now()-_t4) + "ms");

  const _t5 = Date.now();
  const urunEt = sheetToObj(ss, SHEETS.urunEt);
  Logger.log("⏱ [7] UrunEt okundu: " + (Date.now()-_t5) + "ms");

  const _t6 = Date.now();
  const markalar = sheetToRows(ss, SHEETS.markalar);
  Logger.log("⏱ [8] Markalar okundu: " + (Date.now()-_t6) + "ms");

  const _t7 = Date.now();
  const ayarlar = sheetToObj(ss, SHEETS.ayarlar);
  Logger.log("⏱ [9] Ayarlar okundu: " + (Date.now()-_t7) + "ms");

  const _t8 = Date.now();
  const numuneGruplariListesi = getNumuneGruplari(ss);
  const numuneler = getNuumuneler(ss, numuneGruplariListesi);
  Logger.log("⏱ [10] Numuneler+Gruplar okundu: " + (Date.now()-_t8) + "ms");

  const _t9 = Date.now();
  const teklifler = getTeklifler(ss);
  Logger.log("⏱ [11] Teklifler okundu: " + (Date.now()-_t9) + "ms");

  const _t10 = Date.now();
  const urunGorseller = getUrunGorselMap(ss);
  Logger.log("⏱ [12] Ürün görselleri okundu: " + (Date.now()-_t10) + "ms");

  const _t11 = Date.now();
  let magazaBolumleri = [];
  try { magazaBolumleri = getMagazaBolumleri(ss); } catch(e) { Logger.log("Mağaza bölümleri okuma atlandı: " + e.message); }
  Logger.log("⏱ [13] Mağaza bölümleri okundu: " + (Date.now()-_t11) + "ms");

  Logger.log("⏱ [TOPLAM] getAllData: " + (Date.now()-_t0) + "ms");

  return {
    stoklar:        stoklar,
    stokTarih:      stokTarih,
    stokTarihler:   stokTarihler,
    notlar:         notlar,
    etiketler:      etiketler,
    urunEt:         urunEt,
    markalar:       markalar,
    ayarlar:        ayarlar,
    fiyatlar:       fiyatlar,
    fiyatGecmis:    fiyatGecmis,
    numuneler:      numuneler,
    numuneGruplari: numuneGruplariListesi,
    teklifler:      teklifler,
    urunGorseller:  urunGorseller,
    magazaBolumleri: magazaBolumleri,
  };
}

function saveNot(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEETS.notlar, ["STOK_KODU","NOT","TARIH"]);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(body.kod)) {
      if (body.not) sheet.getRange(i+1,2,1,2).setValues([[body.not, new Date()]]);
      else sheet.deleteRow(i+1);
      return { ok: true };
    }
  }
  if (body.not) sheet.appendRow([body.kod, body.not, new Date()]);
  return { ok: true };
}

// Aynı sayfaya aynı anda birden fazla kayıt isteği gelirse (örn. hızlı hızlı
// tıklama) birbirinin üzerine yazmasın diye kilit kullanan ve tüm satırları
// TEK seferde (appendRow yerine setValues ile) yazan yardımcı fonksiyon.
function kilitliTopluYaz(sheetName, headers, rows) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000); // en fazla 15sn bekle
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = getOrCreateSheet(ss, sheetName, headers);
    sheet.clearContents();
    const all = [headers].concat(rows);
    sheet.getRange(1, 1, all.length, headers.length).setValues(all);
  } finally {
    lock.releaseLock();
  }
}

function saveEtiket(body) {
  const rows = (body.etiketler || []).map(e => [e.t||e, e.c||"#d97706"]);
  kilitliTopluYaz(SHEETS.etiketler, ["ETIKET","RENK"], rows);
  return { ok: true };
}

function saveUrunEt(body) {
  const rows = [];
  Object.entries(body.urunEt || {}).forEach(([kod, tags]) => {
    if (tags && tags.length > 0) rows.push([kod, tags.join(",")]);
  });
  kilitliTopluYaz(SHEETS.urunEt, ["STOK_KODU","ETIKETLER"], rows);
  return { ok: true };
}

// ── Ürün-etiket ataması için TEK SATIR güncelleme ──
// ÖNEMLİ: saveUrunEt() tüm STOK_KODU->ETİKET eşleşmesini istemcinin hafızasındaki kopyayla baştan
// yazar. Eskiden bu, sadece "Tamam" butonuna (saveTE) basılınca çağrılıyordu; ama tüm pencereler artık
// boşluğa tıklayınca da kapandığından (İptal ile aynı davranış), kullanıcı bir etiketi işaretleyip
// boşluğa tıklayınca değişiklik hiç kaydedilmeden kayboluyordu. Artık her işaretleme/ekleme anında,
// SADECE o ürünün satırını güncelleyen bu uç nokta çağrılıyor.
function urunEtGuncelleTek(body) {
  // body: {kod, etiketler: [...]}
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = getOrCreateSheet(ss, SHEETS.urunEt, ["STOK_KODU","ETIKETLER"]);
    const data = sheet.getDataRange().getValues();
    let row = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.kod)) { row = i + 1; break; }
    }
    const etiketStr = (body.etiketler || []).join(",");
    if (row === -1) {
      if (etiketStr) sheet.appendRow([body.kod, etiketStr]);
    } else if (!etiketStr) {
      sheet.deleteRow(row); // ürünün artık hiç etiketi kalmadıysa satırı temizle
    } else {
      sheet.getRange(row, 2).setValue(etiketStr);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function etiketTamSil(body) {
  // body: {etiket} — etiket tanımını siler VE bu etiketi taşıyan tüm ürünlerden kaldırır.
  // Az sık yapılan bir yönetici işlemi olduğu için burada birden fazla satır etkilenebilir,
  // ama işlem sheetteki GÜNCEL veri üzerinde (kilit altında okunup yazılarak) yapılıyor —
  // istemcinin hafızasındaki eski bir kopya asla kullanılmıyor.
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const eSheet = ss.getSheetByName(SHEETS.etiketler);
    if (eSheet) {
      const eData = eSheet.getDataRange().getValues();
      for (let i = eData.length - 1; i >= 1; i--) {
        if (String(eData[i][0]) === String(body.etiket)) eSheet.deleteRow(i + 1);
      }
    }
    const uSheet = ss.getSheetByName(SHEETS.urunEt);
    if (uSheet) {
      const uData = uSheet.getDataRange().getValues();
      for (let i = uData.length - 1; i >= 1; i--) {
        const mevcut = String(uData[i][1] || "").split(",").map(s => s.trim()).filter(Boolean);
        const yeni = mevcut.filter(x => x !== body.etiket);
        if (yeni.length !== mevcut.length) {
          if (yeni.length === 0) uSheet.deleteRow(i + 1);
          else uSheet.getRange(i + 1, 2).setValue(yeni.join(","));
        }
      }
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function saveMarka(body) {
  const rows = (body.markalar || []).map(m => [m.k, m.a, m.c||"#2563eb"]);
  kilitliTopluYaz(SHEETS.markalar, ["KOD","AD","RENK"], rows);
  return { ok: true };
}

// ── Marka için TEK SATIR işlemleri ──
// ÖNEMLİ: saveMarka() tüm listeyi (istemcinin hafızasındaki kopyayı) baştan yazar.
// delBrand() gibi fonksiyonlar hiç sync çağırmadığı için silme hiç kaydedilmiyordu (marka geri geliyordu).
// Ayrıca birden fazla sekme/cihaz açıkken tam liste yazımı diğer oturumun eklediği/sildiği markaları da
// ezebilir. Bu yüzden ekle/sil/güncelle/sırala artık tek satır (ya da sırada sadece KOD listesi) ile,
// kilit altında ve sheetteki güncel veriden çalışıyor.

function markaSatirBul(sheet, kod) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(kod)) return i + 1; // 1-indexli sheet satırı
  }
  return -1;
}

function markaEkleTek(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = getOrCreateSheet(ss, SHEETS.markalar, ["KOD","AD","RENK"]);
    if (markaSatirBul(sheet, body.k) !== -1) return { ok: true, already: true };
    sheet.appendRow([body.k, body.a, body.c || "#2563eb"]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function markaSilTek(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.markalar);
    if (!sheet) return { ok: true };
    // Geçmişteki kilitsiz/tam-üzerine-yazma döneminden kalma olası tekrar eden satırlar da
    // silinsin diye SADECE ilk eşleşen değil, o koda ait TÜM satırlar siliniyor.
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(body.k)) sheet.deleteRow(i + 1);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function markaGuncelleTek(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.markalar);
    if (!sheet) return { ok: true };
    const row = markaSatirBul(sheet, body.k);
    if (row === -1) return { ok: true };
    if (body.a !== undefined) sheet.getRange(row, 2).setValue(body.a);
    if (body.c !== undefined) sheet.getRange(row, 3).setValue(body.c);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function markaSiraDegistir(body) {
  // body: {kodSirasi: [kod1, kod2, ...]} — istenen yeni sıralama (sadece kod listesi).
  // Ad/renk sheetten TAZE okunur, böylece sıralama işlemi eşzamanlı bir ekleme/silme/yeniden
  // adlandırmayı asla ezmez. Client listesinde olmayan ama sheette bulunan kodlar kaybolmasın
  // diye sona eklenir.
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = getOrCreateSheet(ss, SHEETS.markalar, ["KOD","AD","RENK"]);
    const data = sheet.getDataRange().getValues();
    const mevcut = {}; // kod -> [ad, renk]
    for (let i = 1; i < data.length; i++) {
      mevcut[String(data[i][0])] = [data[i][1], data[i][2]];
    }
    const siraliKodlar = (body.kodSirasi || []).map(String).filter(k => mevcut[k]);
    Object.keys(mevcut).forEach(k => { if (siraliKodlar.indexOf(k) === -1) siraliKodlar.push(k); });
    const yeniSatirlar = siraliKodlar.map(k => [k, mevcut[k][0], mevcut[k][1]]);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 3).setValues([["KOD","AD","RENK"]]);
    if (yeniSatirlar.length > 0) sheet.getRange(2, 1, yeniSatirlar.length, 3).setValues(yeniSatirlar);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}


// ════════════════════════════════════════════════
// NUMUNE FONKSİYONLARI
// ════════════════════════════════════════════════

function getNuumuneler(ss, numuneGruplariListesi) {
  const sheet = ss.getSheetByName(SHEETS.numuneler);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const gruplar = numuneGruplariListesi || getNumuneGruplari(ss);
  const gecerliGrupIdler = new Set(gruplar.map(g => g.id));
  const gorulen = new Set(); // "grupId|stokKodu"
  return data.slice(1).map(row => ({
    grupId:    String(row[0] || ""),
    stokKodu:  String(row[1] || ""),
    magazaVar: row[2] === true || row[2] === "TRUE" || row[2] === 1,
    konum:     String(row[3] || ""),
    notlar:    String(row[4] || ""),
  })).filter(r => {
    if (!r.grupId || !r.stokKodu) return false;
    if (!gecerliGrupIdler.has(r.grupId)) return false; // artık var olmayan gruba ait yetim satır
    const anahtar = r.grupId + "|" + r.stokKodu;
    if (gorulen.has(anahtar)) return false; // tekrar eden satır — istemci hiç görmesin
    gorulen.add(anahtar);
    return true;
  });
}

function getNumuneGruplari(ss) {
  const sheet = ss.getSheetByName(SHEETS.numuneGruplari);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1).map(row => ({
    id:   String(row[0] || ""),
    ad:   String(row[1] || ""),
    sira: parseInt(row[2]) || 0,
  })).filter(r => r.id && r.ad);
}

function saveNumuneler(body) {
  const rows = (body.numuneler || []).map(n =>
    [n.grupId, n.stokKodu, n.magazaVar ? "TRUE" : "FALSE", n.konum || "", n.notlar || ""]);
  kilitliTopluYaz(SHEETS.numuneler, ["GRUP_ID","STOK_KODU","MAGAZA_VAR","KONUM","NOTLAR"], rows);
  return { ok: true, count: rows.length };
}

// ── Numune için TEK SATIR işlemleri ──
// ÖNEMLİ: Yukarıdaki saveNumuneler() tüm listeyi (istemcinin hafızasındaki kopyayı) baştan yazar.
// Birden fazla sekme/cihaz açıkken, biri eski veriyle kaydederse diğerinin eklediği ürünler silinmiş
// gibi görünür ("kendiliğinden siliniyor" hatası). Bu yüzden ekle/sil/güncelle/sırala işlemleri artık
// tek satır üzerinde, kilit altında çalışıyor — asla tüm tabloyu istemcinin kopyasıyla değiştirmiyor.

function numuneSatirBul(sheet, grupId, stokKodu) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(grupId) && String(data[i][1]) === String(stokKodu)) {
      return i + 1; // 1-indexli sheet satırı
    }
  }
  return -1;
}

function numuneEkleTek(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = getOrCreateSheet(ss, SHEETS.numuneler, ["GRUP_ID","STOK_KODU","MAGAZA_VAR","KONUM","NOTLAR"]);
    if (numuneSatirBul(sheet, body.grupId, body.stokKodu) !== -1) {
      return { ok: true, already: true };
    }
    sheet.appendRow([body.grupId, body.stokKodu, "FALSE", "", ""]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function numuneSilTek(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.numuneler);
    if (!sheet) return { ok: true };
    const row = numuneSatirBul(sheet, body.grupId, body.stokKodu);
    if (row !== -1) sheet.deleteRow(row);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function numuneGuncelleTek(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.numuneler);
    if (!sheet) return { ok: true };
    const row = numuneSatirBul(sheet, body.grupId, body.stokKodu);
    if (row === -1) return { ok: true };
    if (body.magazaVar !== undefined) sheet.getRange(row, 3).setValue(body.magazaVar ? "TRUE" : "FALSE");
    if (body.konum !== undefined) sheet.getRange(row, 4).setValue(body.konum);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function numuneSiraDegistir(body) {
  // body: {grupId, stokKodu, yon}  yon: -1 yukarı, 1 aşağı
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.numuneler);
    if (!sheet) return { ok: true };
    const data = sheet.getDataRange().getValues();
    const grupSatirlari = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.grupId)) grupSatirlari.push(i);
    }
    const idx = grupSatirlari.findIndex(i => String(data[i][1]) === String(body.stokKodu));
    if (idx === -1) return { ok: true };
    const yeniIdx = idx + body.yon;
    if (yeniIdx < 0 || yeniIdx >= grupSatirlari.length) return { ok: true };
    const rowA = grupSatirlari[idx] + 1, rowB = grupSatirlari[yeniIdx] + 1;
    const a = sheet.getRange(rowA, 1, 1, 5).getValues()[0];
    const b = sheet.getRange(rowB, 1, 1, 5).getValues()[0];
    sheet.getRange(rowA, 1, 1, 5).setValues([b]);
    sheet.getRange(rowB, 1, 1, 5).setValues([a]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function numuneGrupSilTek(body) {
  // body: {grupId} — grubu ve içindeki tüm numune satırlarını siler
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const gSheet = ss.getSheetByName(SHEETS.numuneGruplari);
    if (gSheet) {
      const gData = gSheet.getDataRange().getValues();
      for (let i = gData.length - 1; i >= 1; i--) {
        if (String(gData[i][0]) === String(body.grupId)) gSheet.deleteRow(i + 1);
      }
    }
    const nSheet = ss.getSheetByName(SHEETS.numuneler);
    if (nSheet) {
      const nData = nSheet.getDataRange().getValues();
      for (let i = nData.length - 1; i >= 1; i--) {
        if (String(nData[i][0]) === String(body.grupId)) nSheet.deleteRow(i + 1);
      }
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function numuneGrupEkleTek(body) {
  // body: {id, ad, sira} — varsa günceller (isim değişikliği), yoksa ekler
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = getOrCreateSheet(ss, SHEETS.numuneGruplari, ["ID","AD","SIRA"]);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) {
        sheet.getRange(i + 1, 2).setValue(body.ad);
        return { ok: true };
      }
    }
    sheet.appendRow([body.id, body.ad, body.sira != null ? body.sira : (data.length - 1)]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ── TEK SEFERLİK TEMİZLİK ──
// Bu fonksiyonu Apps Script editöründe üstteki fonksiyon seçim kutusundan "temizleNumuneler"i
// seçip ▶ Çalıştır (Run) butonuna basarak ELLE çalıştırın. Web'den çağrılmaz, güvenlidir.
// Yapmadığı: hiçbir grubu veya ürünü silmez. Yaptığı sadece:
//   1) Aynı GRUP_ID + STOK_KODU'na sahip birebir tekrar eden satırları teke indirir (ilkini tutar).
//   2) NumuneGruplari sayfasında artık hiç karşılığı olmayan (silinmiş grup) "yetim" satırları siler.
// Çalıştırdıktan sonra Executions/Yürütmeler sekmesinden loglara bakıp kaç satır etkilendiğini görebilirsiniz.
function temizleNumuneler() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const nSheet = ss.getSheetByName(SHEETS.numuneler);
    const gSheet = ss.getSheetByName(SHEETS.numuneGruplari);
    if (!nSheet) { Logger.log("Numuneler sayfası bulunamadı."); return; }

    const gecerliGrupIdler = new Set(
      gSheet ? gSheet.getDataRange().getValues().slice(1).map(r => String(r[0])) : []
    );

    const data = nSheet.getDataRange().getValues();
    const headers = data[0];
    const gorulen = new Set(); // "grupId|stokKodu"
    const tutulacak = [headers];
    let dupSayisi = 0, yetimSayisi = 0;

    for (let i = 1; i < data.length; i++) {
      const grupId = String(data[i][0] || "");
      const stokKodu = String(data[i][1] || "");
      if (!grupId || !stokKodu) continue;
      if (!gecerliGrupIdler.has(grupId)) { yetimSayisi++; continue; } // artık var olmayan grup
      const anahtar = grupId + "|" + stokKodu;
      if (gorulen.has(anahtar)) { dupSayisi++; continue; } // tekrar eden satır
      gorulen.add(anahtar);
      tutulacak.push(data[i]);
    }

    nSheet.clearContents();
    nSheet.getRange(1, 1, tutulacak.length, headers.length).setValues(tutulacak);

    Logger.log("Temizlik tamamlandı. Önce: " + (data.length - 1) + " satır, sonra: " + (tutulacak.length - 1) + " satır. " +
               "Silinen tekrar: " + dupSayisi + ", silinen yetim: " + yetimSayisi + ".");
  } finally {
    lock.releaseLock();
  }
}

// Aynı mantık Markalar sayfası için: aynı KOD'a sahip tekrar eden satırları teke indirir (ilkini tutar).
// Apps Script editöründe fonksiyon kutusundan "temizleMarkalar"ı seçip ▶ Çalıştır ile elle çalıştırabilir,
// ya da aşağıdaki gunlukTemizlik() ile otomatik tetikleyiciye bağlayabilirsiniz.
function temizleMarkalar() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.markalar);
    if (!sheet) { Logger.log("Markalar sayfası bulunamadı."); return; }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const gorulen = new Set();
    const tutulacak = [headers];
    let dupSayisi = 0;

    for (let i = 1; i < data.length; i++) {
      const kod = String(data[i][0] || "");
      if (!kod) continue;
      if (gorulen.has(kod)) { dupSayisi++; continue; }
      gorulen.add(kod);
      tutulacak.push(data[i]);
    }

    sheet.clearContents();
    sheet.getRange(1, 1, tutulacak.length, headers.length).setValues(tutulacak);

    Logger.log("Marka temizliği tamamlandı. Önce: " + (data.length - 1) + " satır, sonra: " + (tutulacak.length - 1) + " satır. " +
               "Silinen tekrar: " + dupSayisi + ".");
  } finally {
    lock.releaseLock();
  }
}

// ── OTOMATİK GÜNLÜK TEMİZLİK ──
// Bu fonksiyonu bir kez zamanlı tetikleyiciye (trigger) bağlarsanız, sheetler kendiliğinden
// düzenli aralıklarla küçülür — elle çalıştırmaya gerek kalmaz.
// Kurulum (tek seferlik, ~1 dakika):
//   1) Apps Script editöründe sol menüden saat ikonlu "Tetikleyiciler / Triggers" sekmesine girin.
//   2) Sağ alttaki "+ Tetikleyici Ekle" butonuna basın.
//   3) Fonksiyon: gunlukTemizlik, Etkinlik kaynağı: Zamana dayalı, Zamanlayıcı tipi: Günlük zamanlayıcı,
//      istediğiniz saat aralığını seçin (örn. gece 03:00-04:00), Kaydet.
// Bu kurulduktan sonra Markalar ve Numuneler sayfaları her gün otomatik olarak tekilleştirilir.
function gunlukTemizlik() {
  temizleMarkalar();
  temizleNumuneler();
}

function saveNumuneGruplari(body) {
  const rows = (body.gruplar || []).map((g, i) => [g.id, g.ad, g.sira || i]);
  kilitliTopluYaz(SHEETS.numuneGruplari, ["ID","AD","SIRA"], rows);
  return { ok: true };
}

function saveAyar(body) {
  const rows = Object.entries(body.ayarlar || {}).map(([k,v]) => [k, v]);
  kilitliTopluYaz(SHEETS.ayarlar, ["ANAHTAR","DEGER"], rows);
  return { ok: true };
}

// ════════════════════════════════════════════════
// TEKLİF FONKSİYONLARI
// ════════════════════════════════════════════════

function getTeklifler(ss) {
  const sheet = ss.getSheetByName(SHEETS.teklifler);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1).map(row => {
    let items = [];
    try { items = JSON.parse(row[4] || "[]"); } catch(e) {}
    return {
      id:        String(row[0] || ""),
      musteri:   String(row[1] || ""),
      tarih:     String(row[2] || ""),
      olusturma: String(row[3] || ""),
      items:     items,
      not:       String(row[5] || ""),
      kdvAcik:   row[6] === true || row[6] === "TRUE" || row[6] === 1,
    };
  }).filter(t => t.id);
}

function saveTeklifler(body) {
  const rows = (body.teklifler || []).map(t => [
    t.id, t.musteri || "", t.tarih || "", t.olusturma || "",
    JSON.stringify(t.items || []), t.not || "", t.kdvAcik === false ? "FALSE" : "TRUE"
  ]);
  kilitliTopluYaz(SHEETS.teklifler, ["ID","MUSTERI","TARIH","OLUSTURMA","ITEMS_JSON","NOT","KDV_ACIK"], rows);
  return { ok: true, count: rows.length };
}

// ════════════════════════════════════════════════
// ÜRÜN GÖRSELLERİ (Drive tabanlı, çoklu görsel/ürün)
// ════════════════════════════════════════════════

// ── MAĞAZA BÖLÜMLERİ (fotoğraf üzerine ürün pinleme) ──
// Her bölüm (Sıra Stant, Metal Stant, Canlı Mekan 1 vb.) bir fotoğraf + üzerine iğnelenmiş
// ürün konumlarından (pin) oluşur. Fotoğraflar Google Drive'da saklanır (Sheets hücresi
// fotoğraf boyutunu kaldıramaz), Sheets'te sadece dosya ID'si/URL'si + pin koordinatları tutulur.
function getMagazaBolumleri(ss) {
  const bSheet = ss.getSheetByName(SHEETS.magazaBolumleri);
  const pSheet = ss.getSheetByName(SHEETS.magazaPinler);
  const bolumler = [];
  if (!bSheet) return bolumler;
  const bData = bSheet.getDataRange().getValues();
  const pData = pSheet ? pSheet.getDataRange().getValues() : [];

  bData.slice(1).forEach(row => {
    const id = String(row[0] || "");
    if (!id) return;
    bolumler.push({
      id: id,
      ad: String(row[1] || ""),
      fileId: String(row[2] || ""),
      url: String(row[3] || ""),
      pinler: [],
    });
  });

  pData.slice(1).forEach(row => {
    const bolumId = String(row[0] || "");
    const b = bolumler.find(x => x.id === bolumId);
    if (!b) return;
    b.pinler.push({
      id: String(row[1] || ""),
      stokKodu: String(row[2] || ""),
      x: Number(row[3]) || 0,
      y: Number(row[4]) || 0,
      yon: String(row[5] || "yatay"),
    });
  });

  return bolumler;
}

// body: { bolumId, ad, base64, mimeType } — base64 varsa yeni/değişen fotoğraf, yoksa sadece ad güncellenir.
function saveMagazaBolum(body) {
  const ad = String(body.ad || "").trim();
  if (!ad) return { ok: false, hata: "Bölüm adı gerekli" };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEETS.magazaBolumleri, ["ID","AD","FILE_ID","URL","TARIH"]);
  const data = sheet.getDataRange().getValues();

  let bolumId = String(body.bolumId || "").trim();
  let satirIdx = -1;
  if (bolumId) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === bolumId) { satirIdx = i + 1; break; }
    }
  }
  if (!bolumId) bolumId = "bl_" + Date.now();

  let fileId = satirIdx > 0 ? String(data[satirIdx - 1][2] || "") : "";
  let url    = satirIdx > 0 ? String(data[satirIdx - 1][3] || "") : "";

  if (body.base64) {
    const klasor = getOrCreateFolder("MAGAZA_BOLUM_FOTO");
    const mime = body.mimeType || "image/jpeg";
    const ext  = mime.indexOf("png") > -1 ? "png" : "jpg";
    const bytes = Utilities.base64Decode(body.base64);
    const blob  = Utilities.newBlob(bytes, mime, bolumId + "." + ext);
    if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {} }
    const file = klasor.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fileId = file.getId();
    url = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1600";
  }

  const satir = [bolumId, ad, fileId, url, Utilities.formatDate(new Date(), "Europe/Istanbul", "dd/MM/yyyy HH:mm")];
  if (satirIdx > 0) sheet.getRange(satirIdx, 1, 1, satir.length).setValues([satir]);
  else sheet.appendRow(satir);

  return { ok: true, bolumId: bolumId, ad: ad, fileId: fileId, url: url };
}

// body: { bolumId }
function silMagazaBolum(body) {
  const bolumId = String(body.bolumId || "").trim();
  if (!bolumId) return { ok: false, hata: "Eksik parametre (bolumId)" };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const bSheet = ss.getSheetByName(SHEETS.magazaBolumleri);
  if (bSheet) {
    const data = bSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === bolumId) {
        const fileId = String(data[i][2] || "");
        if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {} }
        bSheet.deleteRow(i + 1);
      }
    }
  }
  const pSheet = ss.getSheetByName(SHEETS.magazaPinler);
  if (pSheet) {
    const pData = pSheet.getDataRange().getValues();
    for (let i = pData.length - 1; i >= 1; i--) {
      if (String(pData[i][0]) === bolumId) pSheet.deleteRow(i + 1);
    }
  }
  return { ok: true };
}

// body: { bolumId, pinId, stokKodu, x, y } — pinId varsa günceller (konum taşıma), yoksa yeni pin ekler.
function magazaPinKaydet(body) {
  const bolumId = String(body.bolumId || "").trim();
  const stokKodu = String(body.stokKodu || "").trim();
  if (!bolumId || !stokKodu) return { ok: false, hata: "Eksik parametre (bolumId/stokKodu)" };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEETS.magazaPinler, ["BOLUM_ID","PIN_ID","STOK_KODU","X","Y","YON"]);
  const data = sheet.getDataRange().getValues();

  let pinId = String(body.pinId || "").trim();
  let satirIdx = -1;
  let eskiYon = "yatay";
  if (pinId) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === pinId) { satirIdx = i + 1; eskiYon = String(data[i][5] || "yatay"); break; }
    }
  }
  if (!pinId) pinId = "pin_" + Date.now() + "_" + Math.floor(Math.random()*1000);
  const yon = body.yon != null ? String(body.yon) : eskiYon;

  const satir = [bolumId, pinId, stokKodu, Number(body.x) || 0, Number(body.y) || 0, yon];
  if (satirIdx > 0) sheet.getRange(satirIdx, 1, 1, satir.length).setValues([satir]);
  else sheet.appendRow(satir);

  return { ok: true, pinId: pinId };
}

// body: { pinId }
function magazaPinSil(body) {
  const pinId = String(body.pinId || "").trim();
  if (!pinId) return { ok: false, hata: "Eksik parametre (pinId)" };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.magazaPinler);
  if (!sheet) return { ok: false, hata: "MagazaPinler sayfası bulunamadı" };
  const data = sheet.getDataRange().getValues();
  let silindi = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]).trim() === pinId) { sheet.deleteRow(i + 1); silindi++; }
  }
  if (silindi === 0) return { ok: false, hata: "Eşleşen pin bulunamadı (pinId: " + pinId + ")" };
  return { ok: true, silinen: silindi };
}

function getUrunGorselMap(ss) {
  const sheet = ss.getSheetByName(SHEETS.urunGorselleri);
  const map = {};
  if (!sheet) return map;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return map;
  data.slice(1).forEach(row => {
    const kod = String(row[0] || "").trim();
    if (!kod) return;
    if (!map[kod]) map[kod] = [];
    map[kod].push({
      sira:   Number(row[1]) || 0,
      fileId: String(row[2] || ""),
      url:    String(row[3] || ""),
      etiket: String(row[5] || ""),
    });
  });
  Object.keys(map).forEach(kod => map[kod].sort((a,b) => a.sira - b.sira));
  return map;
}

// body: { kod, sira, base64, mimeType, etiket }
// base64: data URL öneki olmadan sadece base64 gövdesi (data:...;base64, kısmı çıkarılmış olmalı)
// etiket: görsele isteğe bağlı kısa açıklama ("Ön Yüz", "Detay" gibi) — opsiyonel, yeni fotoğraf
//         yüklerken de gönderilebilir, yoksa boş bırakılır (sonradan guncelleGorselEtiket ile de eklenebilir).
function saveUrunGorsel(body) {
  const kod  = String(body.kod || "").trim();
  const sira = Number(body.sira) || 1;
  if (!kod || !body.base64) return { ok: false, hata: "Eksik parametre (kod/base64)" };

  const ss     = SpreadsheetApp.openById(SHEET_ID);
  const sheet  = getOrCreateSheet(ss, SHEETS.urunGorselleri, ["STOK_KODU","SIRA","FILE_ID","URL","TARIH","ETIKET"]);
  const klasor = getOrCreateFolder("URUN_GORSELLERI");

  const mime = body.mimeType || "image/jpeg";
  const ext  = mime.indexOf("png") > -1 ? "png" : "jpg";
  const dosyaAdi = kod + "_" + sira + "." + ext;

  // Aynı isimde eskisi varsa çöpe at (üzerine yazma etkisi)
  const mevcutlar = klasor.getFilesByName(dosyaAdi);
  while (mevcutlar.hasNext()) mevcutlar.next().setTrashed(true);

  const bytes = Utilities.base64Decode(body.base64);
  const blob  = Utilities.newBlob(bytes, mime, dosyaAdi);
  const file  = klasor.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const url = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1000";

  // Sheet'te aynı kod+sira satırı varsa güncelle, yoksa ekle
  const data = sheet.getDataRange().getValues();
  let bulunduSatir = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === kod && Number(data[i][1]) === sira) { bulunduSatir = i + 1; break; }
  }
  // Etiket: body'de gönderildiyse onu kullan; güncellemede gönderilmediyse eski etiketi koru.
  let etiket = body.etiket != null ? String(body.etiket) : "";
  if (bulunduSatir > 0 && body.etiket == null) {
    etiket = String(data[bulunduSatir - 1][5] || "");
  }
  const satir = [kod, sira, fileId, url, Utilities.formatDate(new Date(), "Europe/Istanbul", "dd/MM/yyyy HH:mm"), etiket];
  if (bulunduSatir > 0) {
    // Eski dosyayı da çöpe at (kod+sira aynıysa eski görsel artık gereksiz)
    try {
      const eskiFileId = String(data[bulunduSatir - 1][2] || "");
      if (eskiFileId && eskiFileId !== fileId) DriveApp.getFileById(eskiFileId).setTrashed(true);
    } catch(e) {}
    sheet.getRange(bulunduSatir, 1, 1, satir.length).setValues([satir]);
  } else {
    sheet.appendRow(satir);
  }

  return { ok: true, kod: kod, sira: sira, fileId: fileId, url: url, etiket: etiket };
}

// body: { kod, sira }
function silUrunGorsel(body) {
  const kod  = String(body.kod || "").trim();
  const sira = Number(body.sira);
  if (!kod) return { ok: false, hata: "Eksik parametre (kod)" };

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.urunGorselleri);
  if (!sheet) return { ok: true };

  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === kod && Number(data[i][1]) === sira) {
      const fileId = String(data[i][2] || "");
      try { if (fileId) DriveApp.getFileById(fileId).setTrashed(true); } catch(e) {}
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: true };
}

// body: { kod, sira, etiket } — SADECE görsel etiketini (ör. "Ön Yüz", "Detay") günceller,
// fotoğrafı yeniden yüklemeye gerek yoktur. Görsel galerisindeki her küçük resmin altındaki
// metin kutusundan çağrılır.
function guncelleGorselEtiket(body) {
  const kod  = String(body.kod || "").trim();
  const sira = Number(body.sira);
  if (!kod) return { ok: false, hata: "Eksik parametre (kod)" };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.urunGorselleri);
    if (!sheet) return { ok: true };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === kod && Number(data[i][1]) === sira) {
        sheet.getRange(i + 1, 6).setValue(String(body.etiket || "")); // 6. kolon = ETIKET
        return { ok: true };
      }
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }


  // test: github otomasyonu 05.08.2026
}
