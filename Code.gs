/**
 * ============================================================
 *  CA Settle — Backend (Google Apps Script)
 *  Web App untuk index.html di folder ini
 *  Setup: set Script Properties PIN_ADMIN, PIN_PM, (opsional) GEMINI_API_KEY, DRIVE_FOLDER_ID
 * ============================================================
 */

// ---------- KONFIG ----------
const CFG = {
  SHEET_CA: 'CA',
  SHEET_STRUK: 'Struk',
  // Script Properties keys
  PROP_PIN_ADMIN: 'PIN_ADMIN',
  PROP_PIN_PM: 'PIN_PM',
  PROP_GEMINI_KEY: 'GEMINI_API_KEY',
  PROP_DRIVE_FOLDER: 'DRIVE_FOLDER_ID',
  // OCR limit harian sederhana (opsional)
  OCR_LIMIT_PER_DAY: 200
};

// ---------- HTTP HANDLERS ----------
function doGet() {
  return ContentService.createTextOutput('CA Settle backend aktif').setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  // Lock untuk hindari race saat tulis sheet/drive
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const data = body.data || {};

    // whoami tidak butuh validasi pin lama — pin ada di data.pin
    // semua action lain validasi pin juga
    let auth = null;
    if (action !== 'whoami') {
      auth = authByPin_(data.pin);
      if (!auth.ok) return jsonFail_(auth.error);
    }

    let result;
    switch (action) {
      case 'whoami':       result = handleWhoami_(data); break;
      case 'saveCaInfo':   result = handleSaveCaInfo_(data, auth); break;
      case 'getCaData':    result = handleGetCaData_(data, auth); break;
      case 'getAllCa':     result = handleGetAllCa_(data, auth); break;
      case 'ocrStruk':     result = handleOcrStruk_(data, auth); break;
      case 'uploadStruk':  result = handleUploadStruk_(data, auth); break;
      case 'updateStruk':  result = handleUpdateStruk_(data, auth); break;
      case 'deleteStruk':  result = handleDeleteStruk_(data, auth); break;
      default: return jsonFail_('Action tidak dikenal: ' + action);
    }
    return jsonOk_(result);
  } catch (err) {
    return jsonFail_(err && err.message ? err.message : String(err));
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// ---------- AUTH ----------
function handleWhoami_(data) {
  const pin = String(data.pin || '').trim();
  const auth = authByPin_(pin);
  if (!auth.ok) throw new Error(auth.error);
  return { role: auth.role };
}

function authByPin_(pin) {
  const props = PropertiesService.getScriptProperties();
  const pinAdmin = String(props.getProperty(CFG.PROP_PIN_ADMIN) || '').trim();
  const pinPm    = String(props.getProperty(CFG.PROP_PIN_PM) || '').trim();
  if (!pinAdmin && !pinPm) {
    // fallback jika belum di-setup: izinkan pin apapun sebagai admin (mode dev)
    // GANTI segera di Script Properties!
    return { ok: true, role: 'admin' };
  }
  if (pin && pin === pinAdmin) return { ok: true, role: 'admin' };
  if (pin && pin === pinPm)    return { ok: true, role: 'pm' };
  return { ok: false, error: 'PIN salah, coba lagi.' };
}

function requireAdmin_(auth) {
  if (!auth || auth.role !== 'admin') throw new Error('Aksi ini hanya untuk Admin — PIN PM tidak diizinkan.');
}

// ---------- SHEET HELPERS ----------
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function ensureSheets_() {
  const ss = ss_();
  let ca = ss.getSheetByName(CFG.SHEET_CA);
  if (!ca) {
    ca = ss.insertSheet(CFG.SHEET_CA);
    ca.appendRow(['nomor','nama','dept','jumlah','status','diinputOleh','createdAt','updatedAt']);
    ca.getRange(1,1,1,8).setFontWeight('bold').setBackground('#E4E9F1');
    ca.setFrozenRows(1);
  }
  let st = ss.getSheetByName(CFG.SHEET_STRUK);
  if (!st) {
    st = ss.insertSheet(CFG.SHEET_STRUK);
    st.appendRow(['idStruk','nomorCa','tanggal','merchant','kategori','keterangan','total','confidence','driveFileId','driveFileUrl','diinputOleh','createdAt']);
    st.getRange(1,1,1,12).setFontWeight('bold').setBackground('#E4E9F1');
    st.setFrozenRows(1);
  }
  return { ca: ca, st: st };
}

function findRowByCol_(sheet, colIdx1Based, value) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const vals = sheet.getRange(2, colIdx1Based, last - 1, 1).getValues();
  const target = String(value).trim().toLowerCase();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toLowerCase() === target) return i + 2;
  }
  return -1;
}

function normalizeNomor_(s) {
  return String(s || '').trim().replace(/\s+/g, ' ');
}

// ---------- CA HANDLERS ----------
function handleSaveCaInfo_(data, auth) {
  requireAdmin_(auth);
  const nomorRaw = normalizeNomor_(data.nomor);
  if (!nomorRaw) throw new Error('Nomor CA wajib diisi.');
  ensureSheets_();
  const sh = ss_().getSheetByName(CFG.SHEET_CA);
  const row = findRowByCol_(sh, 1, nomorRaw);
  const now = new Date();
  const jumlah = Number(data.jumlah) || 0;
  const status = data.status || 'Draft';
  if (!['Draft','Diajukan','Settled'].includes(status)) throw new Error('Status tidak valid.');
  // Validasi transisi status sederhana: Settled hanya boleh jika sudah ada struk
  if (status === 'Settled' && row === -1) {
    // CA baru langsung Settled — izinkan tapi warn di log
    Logger.log('CA baru langsung Settled: ' + nomorRaw);
  }

  if (row === -1) {
    sh.appendRow([nomorRaw, String(data.nama||''), String(data.dept||''), jumlah, status, String(data.diinputOleh||''), now, now]);
  } else {
    // cek transisi: tidak boleh mundur dari Settled ke Draft tanpa konfirmasi — di backend tetap izinkan, frontend yang konfirmasi
    sh.getRange(row, 2).setValue(String(data.nama||''));
    sh.getRange(row, 3).setValue(String(data.dept||''));
    sh.getRange(row, 4).setValue(jumlah);
    sh.getRange(row, 5).setValue(status);
    sh.getRange(row, 6).setValue(String(data.diinputOleh||''));
    sh.getRange(row, 8).setValue(now);
  }
  return { nomor: nomorRaw };
}

function handleGetCaData_(data, auth) {
  ensureSheets_();
  const nomor = normalizeNomor_(data.nomorCa);
  if (!nomor) throw new Error('Nomor CA wajib diisi.');
  const caSh = ss_().getSheetByName(CFG.SHEET_CA);
  const row = findRowByCol_(caSh, 1, nomor);
  if (row === -1) throw new Error('CA ' + nomor + ' tidak ditemukan.');
  const vals = caSh.getRange(row, 1, 1, 8).getValues()[0];
  const header = { nomor: vals[0], nama: vals[1], dept: vals[2], jumlah: Number(vals[3])||0, status: vals[4] };

  const stSh = ss_().getSheetByName(CFG.SHEET_STRUK);
  const last = stSh.getLastRow();
  const details = [];
  if (last >= 2) {
    const all = stSh.getRange(2, 1, last - 1, 12).getValues();
    for (let i = 0; i < all.length; i++) {
      const r = all[i];
      if (String(r[1]).trim().toLowerCase() === nomor.toLowerCase()) {
        details.push({
          idStruk: String(r[0]), driveFileId: String(r[8]||''), driveFileUrl: String(r[9]||''),
          tanggal: formatDateISO_(r[2]), merchant: String(r[3]||''), kategori: String(r[4]||''),
          keterangan: String(r[5]||''), total: Number(r[6])||0, confidence: String(r[7]||'')
        });
      }
    }
  }
  return { header: header, details: details };
}

function handleGetAllCa_(data, auth) {
  ensureSheets_();
  const caSh = ss_().getSheetByName(CFG.SHEET_CA);
  const stSh = ss_().getSheetByName(CFG.SHEET_STRUK);
  const lastCa = caSh.getLastRow();
  if (lastCa < 2) return [];
  const caRows = caSh.getRange(2, 1, lastCa - 1, 8).getValues();

  // hitung totalTerpakai per CA dari sheet Struk
  const mapTotal = {};
  const lastSt = stSh.getLastRow();
  if (lastSt >= 2) {
    const stRows = stSh.getRange(2, 1, lastSt - 1, 12).getValues();
    for (let i = 0; i < stRows.length; i++) {
      const nomorCa = normalizeNomor_(stRows[i][1]);
      const tot = Number(stRows[i][6])||0;
      const key = nomorCa.toLowerCase();
      mapTotal[key] = (mapTotal[key]||0) + tot;
    }
  }

  const out = [];
  for (let i = 0; i < caRows.length; i++) {
    const r = caRows[i];
    const nomor = String(r[0]||'').trim();
    if (!nomor) continue;
    const jumlah = Number(r[3])||0;
    const totalTerpakai = mapTotal[nomor.toLowerCase()]||0;
    out.push({
      nomor: nomor, nama: String(r[1]||''), dept: String(r[2]||''),
      jumlah: jumlah, totalTerpakai: totalTerpakai, sisa: jumlah - totalTerpakai, status: String(r[4]||'Draft')
    });
  }
  // terbaru di atas
  out.reverse();
  return out;
}

// ---------- STRUK HANDLERS ----------
function handleUploadStruk_(data, auth) {
  requireAdmin_(auth);
  const nomorCa = normalizeNomor_(data.nomorCa);
  if (!nomorCa) throw new Error('Nomor CA wajib diisi.');
  if (!data.idStruk) throw new Error('idStruk wajib diisi.');
  const total = Number(data.total);
  if (!total || total <= 0) throw new Error('Total harus > 0.');
  ensureSheets_();
  const caSh = ss_().getSheetByName(CFG.SHEET_CA);
  if (findRowByCol_(caSh, 1, nomorCa) === -1) throw new Error('CA ' + nomorCa + ' belum disimpan. Simpan Info CA dulu.');

  // cek duplikat idStruk
  const stSh = ss_().getSheetByName(CFG.SHEET_STRUK);
  if (findRowByCol_(stSh, 1, data.idStruk) !== -1) throw new Error('Struk ' + data.idStruk + ' sudah ada.');

  // validasi sisa budget — warning saja, tetap izinkan simpan (frontend akan confirm)
  // simpan file ke Drive
  let driveFileId = '';
  let driveFileUrl = '';
  if (data.base64) {
    const up = saveToDrive_(data.base64, data.mimeType || 'image/jpeg', data.fileName || (data.idStruk + '.jpg'), nomorCa);
    driveFileId = up.fileId;
    driveFileUrl = up.fileUrl;
  }

  stSh.appendRow([
    String(data.idStruk), nomorCa, data.tanggal || '', String(data.merchant||''), String(data.kategori||'Lainnya'),
    String(data.keterangan||''), total, String(data.confidence||''), driveFileId, driveFileUrl,
    String(data.diinputOleh||''), new Date()
  ]);

  return { driveFileId: driveFileId, driveFileUrl: driveFileUrl };
}

function handleUpdateStruk_(data, auth) {
  requireAdmin_(auth);
  if (!data.idStruk) throw new Error('idStruk wajib diisi.');
  ensureSheets_();
  const sh = ss_().getSheetByName(CFG.SHEET_STRUK);
  const row = findRowByCol_(sh, 1, data.idStruk);
  if (row === -1) throw new Error('Struk ' + data.idStruk + ' tidak ditemukan.');
  const total = Number(data.total);
  if (!total || total <= 0) throw new Error('Total harus > 0.');
  sh.getRange(row, 3).setValue(data.tanggal || '');
  sh.getRange(row, 4).setValue(String(data.merchant||''));
  sh.getRange(row, 5).setValue(String(data.kategori||'Lainnya'));
  sh.getRange(row, 6).setValue(String(data.keterangan||''));
  sh.getRange(row, 7).setValue(total);
  return { ok: true };
}

function handleDeleteStruk_(data, auth) {
  requireAdmin_(auth);
  if (!data.idStruk) throw new Error('idStruk wajib diisi.');
  ensureSheets_();
  const sh = ss_().getSheetByName(CFG.SHEET_STRUK);
  const row = findRowByCol_(sh, 1, data.idStruk);
  if (row === -1) throw new Error('Struk ' + data.idStruk + ' tidak ditemukan.');
  // coba hapus file Drive juga (best-effort)
  try {
    const fileId = String(sh.getRange(row, 9).getValue()||'').trim();
    if (fileId) DriveApp.getFileById(fileId).setTrashed(true);
  } catch (e) { Logger.log('Gagal trash file: ' + e); }
  sh.deleteRow(row);
  return { ok: true };
}

// ---------- OCR ----------
function handleOcrStruk_(data, auth) {
  requireAdmin_(auth);
  if (!data.base64) throw new Error('Gambar struk wajib diisi.');
  // rate limit sederhana via CacheService
  const cache = CacheService.getScriptCache();
  const todayKey = 'ocr_count_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let cnt = Number(cache.get(todayKey) || '0');
  if (cnt >= CFG.OCR_LIMIT_PER_DAY) throw new Error('Batas OCR harian tercapai (' + CFG.OCR_LIMIT_PER_DAY + '), coba lagi besok.');
  cache.put(todayKey, String(cnt + 1), 21600);

  const props = PropertiesService.getScriptProperties();
  const geminiKey = String(props.getProperty(CFG.PROP_GEMINI_KEY) || '').trim();

  if (!geminiKey) {
    // tanpa API key: kembalikan manual agar frontend tetap bisa lanjut isi manual
    return { total: null, merchant: null, tanggal: null, kategori_suggest: 'Lainnya', confidence: 'manual' };
  }

  try {
    const mimeType = data.mimeType || 'image/jpeg';
    const prompt = 'Kamu adalah OCR struk Indonesia. Dari gambar struk ini, ekstrak JSON saja tanpa markdown:\n'
      + '{"total": number|null, "merchant": string|null, "tanggal": "YYYY-MM-DD"|null, "kategori_suggest": string}\n'
      + 'Aturan: total = angka grand total yang dibayar (tanpa Rp/titik/koma, contoh 125000). merchant = nama toko. tanggal = format YYYY-MM-DD jika terbaca. kategori_suggest pilih salah satu: Transportasi, Makan & Minum, Akomodasi, Komunikasi, Perlengkapan / ATK, Lainnya. Jika tidak yakin, isi null dan jangan mengarang.';

    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + encodeURIComponent(geminiKey);
    const payload = {
      contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: mimeType, data: data.base64 } }] }],
      generationConfig: { temperature: 0.1, maxOutputTokens: 512, responseMimeType: 'application/json' }
    };
    const resp = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    const body = resp.getContentText();
    if (code !== 200) throw new Error('OCR gagal (' + code + '): ' + body.slice(0, 300));
    const j = JSON.parse(body);
    const text = j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts
      ? j.candidates[0].content.parts.map(function(p){return p.text||'';}).join('') : '';
    let parsed;
    try { parsed = JSON.parse(text); } catch (e2) {
      // coba ekstrak JSON dari text
      const m = text.match(/\{[\s\S]*\}/);
      parsed = m ? JSON.parse(m[0]) : null;
    }
    if (!parsed) throw new Error('OCR tidak mengembalikan JSON valid');
    const total = parsed.total != null && String(parsed.total).trim() !== '' ? Number(String(parsed.total).replace(/[^0-9]/g,'')) : null;
    const merchant = parsed.merchant ? String(parsed.merchant).trim() : null;
    const tanggal = parsed.tanggal ? String(parsed.tanggal).trim() : null;
    const kategori = parsed.kategori_suggest || 'Lainnya';
    const confidence = (total != null && total > 0) ? 'medium' : 'manual';
    return { total: total || null, merchant: merchant, tanggal: tanggal, kategori_suggest: kategori, confidence: confidence };
  } catch (err) {
    Logger.log('OCR error: ' + err);
    // fallback manual — jangan throw agar UX tetap lanjut
    return { total: null, merchant: null, tanggal: null, kategori_suggest: 'Lainnya', confidence: 'manual', _warn: String(err).slice(0,200) };
  }
}

// ---------- DRIVE ----------
function saveToDrive_(base64, mimeType, fileName, nomorCa) {
  const props = PropertiesService.getScriptProperties();
  const folderId = String(props.getProperty(CFG.PROP_DRIVE_FOLDER) || '').trim();
  let folder;
  if (folderId) {
    try { folder = DriveApp.getFolderById(folderId); } catch (e) { folder = DriveApp.getRootFolder(); }
    // subfolder per CA
    const safeCa = nomorCa.replace(/[^a-zA-Z0-9-_]/g, '_');
    const it = folder.getFoldersByName(safeCa);
    folder = it.hasNext() ? it.next() : folder.createFolder(safeCa);
  } else {
    folder = DriveApp.getRootFolder();
  }
  const pure = base64.indexOf(',') !== -1 ? base64.split(',')[1] : base64;
  const bytes = Utilities.base64Decode(pure);
  const blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', fileName);
  const file = folder.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return { fileId: file.getId(), fileUrl: file.getUrl() };
}

// ---------- UTIL ----------
function formatDateISO_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(v).trim();
  // coba parse dd/MM/yyyy atau yyyy-MM-dd
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0,10);
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m) return m[3] + '-' + ('0'+m[2]).slice(-2) + '-' + ('0'+m[1]).slice(-2);
  return s;
}

function jsonOk_(result) {
  return ContentService.createTextOutput(JSON.stringify({ success: true, result: result })).setMimeType(ContentService.MimeType.JSON);
}
function jsonFail_(msg) {
  return ContentService.createTextOutput(JSON.stringify({ success: false, error: String(msg) })).setMimeType(ContentService.MimeType.JSON);
}

// ---------- SETUP HELPER (jalankan sekali dari Editor) ----------
function setupCasettle() {
  ensureSheets_();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(CFG.PROP_PIN_ADMIN)) props.setProperty(CFG.PROP_PIN_ADMIN, '123456');
  if (!props.getProperty(CFG.PROP_PIN_PM)) props.setProperty(CFG.PROP_PIN_PM, '654321');
  Logger.log('Setup selesai. Sheets CA & Struk siap. PIN default admin=123456 pm=654321 — ganti di Script Properties!');
}
