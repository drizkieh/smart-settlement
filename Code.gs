/** ============================================================
 *  CA SETTLEMENT — Backend (Google Apps Script)
 *  Sheets  : CA_Header, Struk_Detail
 *  Drive   : folder induk -> subfolder per Nomor_CA
 *
 *  SETUP:
 *  1. Ganti SHEET_ID & PARENT_FOLDER_ID di bawah
 *  2. File > Project properties > Script properties -> tambah key:
 *       PIN_ADMIN  = pin rahasia admin
 *       PIN_PM     = pin rahasia PM (read-only)
 *     PIN TIDAK ditulis di kode / di HTML. Validasi PIN sepenuhnya
 *     terjadi di server (lihat resolveRole_ di bawah), jadi siapa pun
 *     yang lihat source HTML/kode ini TIDAK bisa melihat PIN asli.
 *  3. FOTO STRUK TIDAK LAGI PAKAI DRIVE LINK SHARING (ANYONE/DOMAIN).
 *     File struk disimpan PRIVATE (cuma bisa dibaca akun pemilik script).
 *     Frontend minta gambarnya lewat action 'getReceiptImage' yang digate
 *     PIN — backend baca file (sebagai pemiliknya, selalu boleh) lalu
 *     kirim balik base64. Ini karena banyak akun Google Workspace kantor
 *     mematikan fitur "link sharing" (ANYONE_WITH_LINK maupun
 *     DOMAIN_WITH_LINK) di level kebijakan admin, yang bikin
 *     file.setSharing() selalu gagal dengan "Access denied: DriveApp"
 *     walau scope OAuth & Advanced Service sudah benar. Constant
 *     RECEIPT_SHARING_MODE di bawah sudah TIDAK DIPAKAI lagi, dibiarkan
 *     saja kalau suatu saat mau dipakai lagi di lingkungan yang izinnya
 *     lebih longgar.
 *  4. WAJIB aktifkan Advanced Service "Drive API" (dipakai buat OCR gratis):
 *     di editor Apps Script -> klik "Layanan"/"Services" (ikon +) di sidebar
 *     kiri -> cari "Drive API" -> Add. Tanpa ini, OCR akan error "Drive is
 *     not defined".
 *     Setelah itu WAJIB jalankan salah satu fungsi (misal setupSheets) SEKALI
 *     manual dari editor (pilih di dropdown -> klik Jalankan), lalu setujui
 *     popup izin baru yang muncul (Review permissions -> Advanced -> Go to
 *     project -> Allow). Ini cuma perlu sekali per project. Habis itu WAJIB
 *     Deploy > Manage deployments > Edit (pensil) > New version, supaya web
 *     app yang sedang jalan ikut pakai izin barunya.
 *  5. Jalankan setupSheets() sekali dari editor untuk bikin header kolom
 *  6. Deploy > New deployment > Web app
 *     Execute as: Me
 *     Who has access: Anyone
 *     (URL web app ini boleh publik — keamanan sekarang bergantung ke PIN
 *     yang divalidasi di server, bukan ke URL-nya dirahasiakan, dan foto
 *     struk sendiri private, cuma bisa diambil lewat action yang digate PIN)
 *  7. Setiap kali edit Code.gs, WAJIB Deploy > Manage deployments > Edit (pensil)
 *     > New version, kalau nggak URL lama tetap jalanin kode versi lama
 *
 *  CATATAN OCR: Hybrid — coba Google AI (Gemini) dulu kalau ada
 *  GEMINI_API_KEY di Script Properties, fallback ke Google Drive OCR gratis
 *  kalau tidak ada key / limit habis / error. Jadi tanpa key pun tetap jalan
 *  (Drive OCR), tapi dengan key akurasinya jauh lebih tinggi (confidence high).
 *  Dapatkan key gratis di https://aistudio.google.com → Get API key → Create.
 * ============================================================ */
const SHEET_ID = '1HPC2u3WzActS4IefMfjVmmq7_zwnKkRS1kxiQHrMdyA';
const PARENT_FOLDER_ID = '1GtZUUd7jyHHZmgleqlmIhbsIn-i_JTOx';

const SHEET_HEADER = 'CA_Header';
const SHEET_DETAIL = 'Struk_Detail';

// SUDAH TIDAK DIPAKAI — foto struk sekarang selalu private & diambil lewat
// action 'getReceiptImage' (proxy lewat backend). Dibiarkan di sini kalau
// suatu saat mau dipakai lagi di akun/lingkungan yang izin sharing-nya
// tidak dibatasi admin Workspace.
const RECEIPT_SHARING_MODE = 'ANYONE';

// Batas pemanggilan OCR per PIN per hari — jaga-jaga kalau endpoint
// disalahgunakan / dipanggil berulang di luar app.
const OCR_DAILY_LIMIT_PER_PIN = 150; // Gemini free tier ~1500/hari, Drive OCR juga ada kuota; 150 aman

// Model Gemini untuk OCR AI. 2.0-flash paling murah & cepat untuk struk.
const GEMINI_MODEL = 'gemini-2.0-flash';

// Bahasa OCR fallback Drive. 'id' = Indonesia.
const OCR_LANGUAGE = 'id';

/* ================= AUTH ================= */
// Role mana saja yang boleh memanggil tiap action. Ini dicek di server,
// jadi PIN & role BUKAN sekadar gerbang tampilan di HTML.
const ACTION_ROLES = {
  whoami:          ['admin', 'pm'],
  saveCaInfo:      ['admin'],
  uploadStruk:     ['admin'],
  updateStruk:     ['admin'],
  deleteStruk:     ['admin'],
  ocrStruk:        ['admin'],
  getCaData:       ['admin', 'pm'],
  getAllCa:        ['admin', 'pm'],
  getReceiptImage: ['admin', 'pm']
};

function getRolePins_() {
  const props = PropertiesService.getScriptProperties();
  // FIX: sebelumnya kode ini salah baca key 'ism2026' / 'pmism2026' (bekas PIN
  // lama), padahal Script Property yang benar namanya PIN_ADMIN / PIN_PM.
  // Itu sebabnya login selalu gagal walau PIN_ADMIN/PIN_PM sudah diisi.
  const admin = props.getProperty('PIN_ADMIN');
  const pm = props.getProperty('PIN_PM');
  if (!admin || !pm) {
    throw new Error('PIN_ADMIN / PIN_PM belum diset di Script Properties. Lihat komentar SETUP di atas.');
  }
  return { admin: admin, pm: pm };
}

function resolveRole_(pin) {
  if (!pin) return null;
  const pins = getRolePins_();
  if (pin === pins.admin) return 'admin';
  if (pin === pins.pm) return 'pm';
  return null;
}

/* ================= ENTRY POINT ================= */

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const data = body.data || {};

    const allowedRoles = ACTION_ROLES[action];
    if (!allowedRoles) throw new Error('Action tidak dikenal: ' + action);

    const role = resolveRole_(data.pin);
    if (!role) throw new Error('PIN salah atau tidak dikirim.');
    if (allowedRoles.indexOf(role) === -1) {
      throw new Error('Role ' + role + ' tidak diizinkan melakukan aksi ini.');
    }

    let result;
    switch (action) {
      case 'whoami':          result = { role: role };                              break;
      case 'saveCaInfo':      result = saveCaInfo(data);                             break;
      case 'uploadStruk':     result = uploadStruk(data);                            break;
      case 'updateStruk':     result = updateStruk(data);                            break;
      case 'deleteStruk':     result = deleteStruk(data);                            break;
      case 'getCaData':       result = getCaData(data);                              break;
      case 'getAllCa':        result = getAllCa();                                   break;
      case 'ocrStruk':        result = ocrStruk_(data.base64, data.mimeType, data.pin); break;
      case 'getReceiptImage': result = getReceiptImage(data);                        break;
      default: throw new Error('Action tidak dikenal: ' + action);
    }

    return jsonResponse({ success: true, result: result });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ================= SETUP (jalankan sekali manual dari editor) ================= */

function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  let header = ss.getSheetByName(SHEET_HEADER);
  if (!header) header = ss.insertSheet(SHEET_HEADER);
  header.clear();
  header.appendRow([
    'Nomor_CA', 'Nama_Pemohon', 'Departemen', 'Jumlah_CA_Diterima',
    'Tanggal_Pengajuan', 'Status', 'Drive_Folder_ID', 'Drive_Folder_URL',
    'Created_At', 'Updated_At', 'Diinput_Oleh'
  ]);

  let detail = ss.getSheetByName(SHEET_DETAIL);
  if (!detail) detail = ss.insertSheet(SHEET_DETAIL);
  detail.clear();
  detail.appendRow([
    'ID_Struk', 'Nomor_CA', 'Tanggal_Struk', 'Merchant', 'Kategori',
    'Keterangan', 'Total', 'Confidence_OCR', 'Nama_File',
    'Drive_File_ID', 'Drive_File_URL', 'Created_At', 'Diinput_Oleh'
  ]);
}

/* ================= HELPER: LOCK / SANITASI / NORMALISASI ================= */

// Bungkus operasi tulis-ke-sheet dengan lock supaya dua request yang datang
// nyaris bersamaan (misal dua admin submit di detik yang sama) nggak saling
// tabrakan / bikin data dobel.
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// Cegah formula-injection: kalau teks user diawali =, +, -, atau @, Sheets bisa
// menganggapnya formula saat dibuka manual. Kasih prefix apostrof biar dianggap teks.
function sanitizeForSheet_(val) {
  if (typeof val !== 'string') return val;
  return /^[=+\-@]/.test(val) ? "'" + val : val;
}

function normalizeNomor_(s) {
  return String(s || '').trim();
}
// dipakai buat PEMBANDINGAN saja (case-insensitive), bukan buat disimpan —
// nilai yang disimpan tetap format asli yang diketik user (setelah di-trim).
function nomorKey_(s) {
  return normalizeNomor_(s).toUpperCase();
}

/* ================= OCR (HYBRID: Gemini AI → fallback Drive OCR) ================= */

function checkOcrQuota_(pin) {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd');
  const key = 'OCR_COUNT_' + today + '_' + Utilities.base64EncodeWebSafe(pin || '').slice(0, 16);
  const count = Number(props.getProperty(key) || 0);
  if (count >= OCR_DAILY_LIMIT_PER_PIN) {
    throw new Error('Batas OCR harian tercapai (' + OCR_DAILY_LIMIT_PER_PIN + 'x/hari). Isi manual dulu, atau coba lagi besok.');
  }
  props.setProperty(key, String(count + 1));
}

// Entry point OCR yang dipanggil doPost. Mengembalikan bentuk JSON yang SAMA
// persis seperti versi Anthropic sebelumnya, jadi index.html tidak perlu diubah:
// { total, merchant, tanggal, kategori_suggest, confidence }
// Google Doc hasil OCR nggak bisa di-convert ke text/plain lewat getAs() atau
// dibuka lewat DocumentApp (butuh scope 'documents' terpisah & kadang error
// "conversion not supported"). Cara paling reliable: panggil langsung Drive
// REST export endpoint pakai token OAuth punya script ini sendiri — cukup
// modal scope "drive" yang memang sudah wajib buat OCR.
function exportGoogleDocAsText_(fileId) {
  const url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) + '/export?mimeType=text%2Fplain';
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Gagal ambil teks hasil OCR (HTTP ' + res.getResponseCode() + '): ' + res.getContentText().slice(0, 200));
  }
  return res.getContentText('UTF-8');
}

function ocrStruk_(base64, mimeType, pin) {
  checkOcrQuota_(pin);

  // 1) Coba Gemini AI dulu kalau ada key — akurasi tinggi, ~1500/hari gratis
  const geminiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (geminiKey && String(geminiKey).trim()) {
    try {
      return geminiOcr_(base64, mimeType, String(geminiKey).trim());
    } catch (e) {
      Logger.log('Gemini OCR gagal, fallback Drive OCR: ' + e);
      // jangan throw — lanjut fallback ke Drive OCR di bawah
    }
  }

  // 2) Fallback Drive OCR gratis (perlu Advanced Service Drive API)
  if (typeof Drive === 'undefined') {
    throw new Error('GEMINI_API_KEY belum diisi dan Drive API belum diaktifkan. Isi GEMINI_API_KEY di Script Properties atau aktifkan Drive API (Services > Drive API).');
  }

  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType || 'image/jpeg', 'struk_ocr_temp');

  let tempDocId = null;
  try {
    const tempDoc = Drive.Files.create(
      { name: 'TEMP_OCR_' + new Date().getTime(), mimeType: MimeType.GOOGLE_DOCS },
      blob,
      { ocr: true, ocrLanguage: OCR_LANGUAGE }
    );
    tempDocId = tempDoc.id;

    const text = exportGoogleDocAsText_(tempDocId);
    return parseReceiptText_(text);
  } finally {
    if (tempDocId) {
      try { DriveApp.getFileById(tempDocId).setTrashed(true); } catch (e) { /* abaikan */ }
    }
  }
}

function geminiOcr_(base64, mimeType, apiKey) {
  const prompt = 'Kamu adalah OCR struk Indonesia. Dari gambar struk ini ekstrak JSON saja tanpa markdown: '
    + '{"total": number|null, "merchant": string|null, "tanggal": "YYYY-MM-DD"|null, "kategori_suggest": string} '
    + 'Aturan: total = angka grand total yang dibayar (tanpa Rp/titik/koma, contoh 125000). merchant = nama toko. tanggal = format YYYY-MM-DD jika terbaca. kategori_suggest pilih salah satu: Transportasi, Makan & Minum, Akomodasi, Komunikasi, Perlengkapan / ATK, Lainnya. Jika tidak yakin isi null jangan mengarang.';
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(GEMINI_MODEL) + ':generateContent?key=' + encodeURIComponent(apiKey);
  const payload = {
    contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: mimeType || 'image/jpeg', data: base64 } }] }],
    generationConfig: { temperature: 0.1, maxOutputTokens: 512, responseMimeType: 'application/json' }
  };
  const resp = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code !== 200) throw new Error('Gemini OCR gagal (' + code + '): ' + body.slice(0, 400));
  const j = JSON.parse(body);
  const text = j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts
    ? j.candidates[0].content.parts.map(function(p){return p.text||'';}).join('') : '';
  let parsed;
  try { parsed = JSON.parse(text); } catch (e2) {
    const m = text.match(/\{[\s\S]*\}/);
    parsed = m ? JSON.parse(m[0]) : null;
  }
  if (!parsed) throw new Error('Gemini tidak mengembalikan JSON valid: ' + text.slice(0,200));
  const total = parsed.total != null && String(parsed.total).trim() !== '' ? Number(String(parsed.total).replace(/[^0-9]/g,'')) : null;
  if (total != null && (isNaN(total) || total <= 0)) throw new Error('Gemini total tidak valid');
  return {
    total: total || null,
    merchant: parsed.merchant ? String(parsed.merchant).trim().slice(0,60) : null,
    tanggal: parsed.tanggal ? String(parsed.tanggal).trim() : null,
    kategori_suggest: parsed.kategori_suggest || 'Lainnya',
    confidence: (total != null && total > 0) ? 'high' : 'low'
  };
}

// Baca teks hasil OCR mentah dan cari total, merchant, tanggal, kategori.
// Bukan ilmu pasti (regex-based) — makanya hasilnya tetap harus dicek user
// di layar review sebelum disimpan.
function parseReceiptText_(text) {
  const lines = String(text || '').split('\n').map(l => l.trim()).filter(Boolean);

  const total = extractTotal_(lines);
  const merchant = lines.length ? lines[0].slice(0, 60) : null;
  const tanggal = extractTanggal_(text);
  const kategori = guessKategori_(text);
  const confidence = total ? 'medium' : 'low';

  return {
    total: total,
    merchant: merchant,
    tanggal: tanggal,
    kategori_suggest: kategori,
    confidence: confidence
  };
}

function extractNumberFromLine_(line) {
  const matches = line.match(/(\d{1,3}(?:[.,]\d{3})+|\d{4,})/g);
  if (!matches) return null;
  // Ambil angka terakhir di baris itu (biasanya nominal muncul setelah labelnya)
  const raw = matches[matches.length - 1];
  const digits = raw.replace(/[.,]/g, '');
  const num = parseInt(digits, 10);
  return isNaN(num) || num <= 0 ? null : num;
}

function extractTotal_(lines) {
  // Urutan prioritas label, dari yang paling spesifik/paling sering benar dulu.
  const priorityPatterns = [
    /grand\s*total/i,
    /total\s*bayar|total\s*tagihan|total\s*akhir|total\s*belanja/i,
    /\bjumlah\s*bayar\b/i,
    /\btotal\b/i,
    /\bjumlah\b/i
  ];

  for (let p = 0; p < priorityPatterns.length; p++) {
    const pattern = priorityPatterns[p];
    for (let i = 0; i < lines.length; i++) {
      if (!pattern.test(lines[i])) continue;
      if (/sub\s*-?\s*total/i.test(lines[i])) continue; // skip subtotal
      let num = extractNumberFromLine_(lines[i]);
      if (num == null && lines[i + 1]) num = extractNumberFromLine_(lines[i + 1]);
      if (num != null) return num;
    }
  }
  return null;
}

function extractTanggal_(text) {
  // Format dd/mm/yyyy atau dd-mm-yyyy
  let m = text.match(/\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\b/);
  if (m) {
    let [, d, mo, y] = m;
    if (y.length === 2) y = '20' + y;
    d = d.padStart(2, '0'); mo = mo.padStart(2, '0');
    if (Number(mo) >= 1 && Number(mo) <= 12 && Number(d) >= 1 && Number(d) <= 31) {
      return y + '-' + mo + '-' + d;
    }
  }
  // Format yyyy-mm-dd
  m = text.match(/\b(20\d{2})-(\d{1,2})-(\d{1,2})\b/);
  if (m) {
    let [, y, mo, d] = m;
    return y + '-' + mo.padStart(2, '0') + '-' + d.padStart(2, '0');
  }
  return null;
}

function guessKategori_(text) {
  const t = String(text || '').toLowerCase();
  if (/grab|gojek|taxi|taksi|parkir|\btol\b|bensin|pertamina|shell|spbu|bbm/.test(t)) return 'Transportasi';
  if (/restoran|resto\b|cafe|kafe|kopi|coffee|makan|nasi|warung|kfc|mcdonald|pizza|burger/.test(t)) return 'Makan & Minum';
  if (/hotel|inn\b|resort|penginapan|guesthouse|homestay/.test(t)) return 'Akomodasi';
  if (/pulsa|paket data|internet|wifi|telkomsel|indosat|xl axiata|smartfren|tri\b/.test(t)) return 'Komunikasi';
  if (/\batk\b|alat tulis|fotocopy|fotokopi|print\b|percetakan/.test(t)) return 'Perlengkapan / ATK';
  return 'Lainnya';
}

/* ================= CA HEADER ================= */

function saveCaInfo(data) {
  return withLock_(() => {
    const nomorDisplay = normalizeNomor_(data.nomor);
    if (!nomorDisplay) throw new Error('Nomor CA wajib diisi');
    const key = nomorKey_(nomorDisplay);

    const sheet = getSheet(SHEET_HEADER);
    const rows = sheet.getDataRange().getValues();
    const idx = rows.findIndex(r => nomorKey_(r[0]) === key);
    const now = new Date();

    const nama = sanitizeForSheet_(data.nama);
    const dept = sanitizeForSheet_(data.dept);
    const jumlah = Number(data.jumlah) || 0;

    if (idx > -1) {
      sheet.getRange(idx + 1, 2, 1, 4).setValues([[nama, dept, jumlah, data.status || rows[idx][5]]]);
      sheet.getRange(idx + 1, 10).setValue(now); // Updated_At
      return { nomor: rows[idx][0], mode: 'updated', folderUrl: rows[idx][7] };
    }

    // CA baru -> buat folder Drive khusus CA ini
    const folder = DriveApp.getFolderById(PARENT_FOLDER_ID).createFolder(nomorDisplay);
    sheet.appendRow([
      nomorDisplay, nama, dept, jumlah, now,
      data.status || 'Draft', folder.getId(), folder.getUrl(), now, now,
      sanitizeForSheet_(data.diinputOleh || '')
    ]);
    return { nomor: nomorDisplay, mode: 'created', folderUrl: folder.getUrl() };
  });
}

/* ================= STRUK DETAIL ================= */

function uploadStruk(data) {
  // data: { nomorCa, idStruk, tanggal, merchant, kategori, keterangan, total, confidence, fileName, mimeType, base64, diinputOleh }
  const folderId = getCaFolderId(data.nomorCa);
  const folder = DriveApp.getFolderById(folderId);

  // Upload ke Drive di LUAR lock (lambat) — cuma penulisan baris sheet yang dikunci di bawah,
  // supaya beberapa admin bisa upload foto barengan tanpa saling nunggu.
  const blob = Utilities.newBlob(Utilities.base64Decode(data.base64), data.mimeType, data.fileName);
  const file = folder.createFile(blob);
  // File SENGAJA dibiarkan private (default) — tidak dipanggil file.setSharing().
  // Banyak akun Google Workspace kantor mematikan link sharing (ANYONE/DOMAIN)
  // lewat kebijakan admin, yang bikin setSharing() selalu gagal dengan
  // "Access denied: DriveApp" walau scope OAuth sudah benar. Karena itu foto
  // diambil lewat action 'getReceiptImage' (proxy, dibaca sebagai pemilik file),
  // bukan lewat link Drive langsung.

  return withLock_(() => {
    const sheet = getSheet(SHEET_DETAIL);
    const now = new Date();
    sheet.appendRow([
      data.idStruk, normalizeNomor_(data.nomorCa), data.tanggal, sanitizeForSheet_(data.merchant), data.kategori,
      sanitizeForSheet_(data.keterangan), Number(data.total) || 0, data.confidence, sanitizeForSheet_(data.fileName),
      file.getId(), file.getUrl(), now, sanitizeForSheet_(data.diinputOleh || '')
    ]);
    return { idStruk: data.idStruk, driveFileId: file.getId(), driveFileUrl: file.getUrl() };
  });
}

function updateStruk(data) {
  // data: { idStruk, tanggal, merchant, kategori, keterangan, total }
  return withLock_(() => {
    const sheet = getSheet(SHEET_DETAIL);
    const rows = sheet.getDataRange().getValues();
    const idx = rows.findIndex(r => r[0] === data.idStruk);
    if (idx === -1) throw new Error('Struk tidak ditemukan: ' + data.idStruk);

    sheet.getRange(idx + 1, 3, 1, 5).setValues([[
      data.tanggal, sanitizeForSheet_(data.merchant), data.kategori, sanitizeForSheet_(data.keterangan), Number(data.total) || 0
    ]]);
    return { idStruk: data.idStruk, mode: 'updated' };
  });
}

function deleteStruk(data) {
  // data: { idStruk }
  return withLock_(() => {
    const sheet = getSheet(SHEET_DETAIL);
    const rows = sheet.getDataRange().getValues();
    const idx = rows.findIndex(r => r[0] === data.idStruk);
    if (idx === -1) throw new Error('Struk tidak ditemukan: ' + data.idStruk);

    const driveFileId = rows[idx][9];
    try { DriveApp.getFileById(driveFileId).setTrashed(true); } catch (e) {}
    sheet.deleteRow(idx + 1);
    return { idStruk: data.idStruk, mode: 'deleted' };
  });
}

/* ================= FOTO STRUK (PROXY — BUKAN LINK DRIVE LANGSUNG) ================= */

// Dipanggil frontend buat nampilin foto struk. File Drive-nya private (nggak
// pernah di-setSharing), jadi cara satu-satunya buat lihat isinya adalah lewat
// backend ini — yang jalan sebagai pemilik file (selalu boleh baca file
// miliknya sendiri, nggak butuh izin sharing apa pun). Akses ke action ini
// sendiri tetap digate PIN lewat ACTION_ROLES seperti action lain.
function getReceiptImage(data) {
  // data: { fileId }
  if (!data.fileId) throw new Error('fileId wajib diisi');
  const file = DriveApp.getFileById(data.fileId);
  const blob = file.getBlob();
  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    mimeType: blob.getContentType()
  };
}

/* ================= READ ================= */

function getCaData(data) {
  // data: { nomorCa }
  const key = nomorKey_(data.nomorCa);
  const headerSheet = getSheet(SHEET_HEADER);
  const headerRows = headerSheet.getDataRange().getValues();
  const headerRow = headerRows.find(r => nomorKey_(r[0]) === key);
  if (!headerRow) throw new Error('Nomor CA tidak ditemukan');

  const detailSheet = getSheet(SHEET_DETAIL);
  const detailRows = detailSheet.getDataRange().getValues().slice(1)
    .filter(r => nomorKey_(r[1]) === key)
    .map(r => ({
      idStruk: r[0], nomorCa: r[1], tanggal: fmtDate(r[2]), merchant: r[3],
      kategori: r[4], keterangan: r[5], total: r[6], confidence: r[7],
      fileName: r[8], driveFileId: r[9], driveFileUrl: r[10], createdAt: r[11]
    }));

  return {
    header: {
      nomor: headerRow[0], nama: headerRow[1], dept: headerRow[2],
      jumlah: headerRow[3], tanggalPengajuan: headerRow[4], status: headerRow[5],
      folderUrl: headerRow[7]
    },
    details: detailRows
  };
}

function getAllCa() {
  // dipakai dashboard PM: rekap semua CA + total terpakai per CA
  const headerRows = getSheet(SHEET_HEADER).getDataRange().getValues().slice(1);
  const detailRows = getSheet(SHEET_DETAIL).getDataRange().getValues().slice(1);

  const totalsByCA = {};
  detailRows.forEach(r => {
    const key = nomorKey_(r[1]);
    totalsByCA[key] = (totalsByCA[key] || 0) + (Number(r[6]) || 0);
  });

  return headerRows.map(r => {
    const nomor = r[0];
    const jumlah = Number(r[3]) || 0;
    const terpakai = totalsByCA[nomorKey_(nomor)] || 0;
    return {
      nomor: nomor, nama: r[1], dept: r[2], jumlah: jumlah,
      tanggalPengajuan: fmtDate(r[4]), status: r[5], folderUrl: r[7],
      totalTerpakai: terpakai, sisa: jumlah - terpakai
    };
  }).reverse(); // terbaru di atas
}

/* ================= HELPER ================= */

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet tidak ditemukan: ' + name);
  return sheet;
}

function getCaFolderId(nomorCa) {
  const sheet = getSheet(SHEET_HEADER);
  const rows = sheet.getDataRange().getValues();
  const key = nomorKey_(nomorCa);
  const row = rows.find(r => nomorKey_(r[0]) === key);
  if (!row) throw new Error('Nomor CA belum tersimpan — simpan info CA dulu sebelum upload struk');
  return row[6];
}

function fmtDate(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd');
  }
  return String(v);
}
