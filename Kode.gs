/**
 * =========================================================================
 *  DASHBOARD JADWAL & ABSEN TOKO  —  v2 (cepat)
 * =========================================================================
 *  Perubahan utama dibanding v1:
 *  - 1 panggilan server per tanggal (getDayBundle) berisi SEMUA toko:
 *    pindah toko jadi instan karena data sudah ada di browser.
 *  - Cache dipecah per potongan (batas CacheService 100 KB/kunci tidak
 *    lagi membuat cache gagal diam-diam lalu baca spreadsheet terus).
 *  - Log foto di-cache dan hanya dibaca sekali, bukan tiap ganti toko.
 *  - Upload: folder Drive dicari sekali (ID disimpan), sharing diatur di
 *    folder (bukan per file), tanpa Session.getActiveUser().
 * =========================================================================
 */

const CONFIG = {
  SPREADSHEET_ID: '1XLWPJ2me7jQfBd7SekOnYV_JVAAPfyo_1LaqEsUIPgA',
  SHEET_NAME: 'Sheet1',

  DATA_START_ROW: 4,
  COL_NIK: 1,
  COL_NAMA: 2,
  COL_JABATAN: 3,
  COL_LOKASI: 4,
  COL_TANGGAL_START: 6,
  JUMLAH_TANGGAL: 30,

  DRIVE_ROOT_FOLDER: 'Foto Absensi Toko',
  LOG_SHEET_NAME: 'LOG_FOTO',

  CACHE_SECONDS: 900,      // jadwal (tombol Refresh menghapusnya)
  LOG_CACHE_SECONDS: 300,  // log foto (otomatis dihapus saat upload/hapus)

  // Jika foto tidak tampil di dashboard setelah upload, ubah ke true
  // (artinya sharing diatur per file; lebih lambat ~1 detik per foto).
  SHARE_PER_FILE: false,
};

/* ------------------------------------------------------------------ */
/*  WEB APP                                                            */
/* ------------------------------------------------------------------ */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Dashboard Jadwal & Absen Toko')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Jalankan SEKALI dari editor Apps Script (untuk otorisasi + siapkan folder & log). */
function setupSekali() {
  const folder = getRootFolder_();
  getLogSheet_();
  Logger.log('Folder foto siap: ' + folder.getUrl());
}

/* ------------------------------------------------------------------ */
/*  CACHE BERPOTONGAN                                                  */
/* ------------------------------------------------------------------ */

const CHUNK_ = 30000; // karakter per potongan (aman < 100 KB)

function cachePutBig_(key, obj, seconds) {
  try {
    const s = JSON.stringify(obj);
    const n = Math.ceil(s.length / CHUNK_) || 1;
    const m = {};
    m[key + '#n'] = String(n);
    for (let i = 0; i < n; i++) m[key + '#' + i] = s.substr(i * CHUNK_, CHUNK_);
    CacheService.getScriptCache().putAll(m, seconds);
  } catch (err) {}
}

function cacheGetBig_(key) {
  try {
    const cache = CacheService.getScriptCache();
    const n = parseInt(cache.get(key + '#n'), 10);
    if (!n) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(key + '#' + i);
    const got = cache.getAll(keys);
    let s = '';
    for (let i = 0; i < n; i++) {
      const part = got[key + '#' + i];
      if (part == null) return null;
      s += part;
    }
    return JSON.parse(s);
  } catch (err) {
    return null;
  }
}

function cacheRemoveBig_(key) {
  try {
    const cache = CacheService.getScriptCache();
    const n = parseInt(cache.get(key + '#n'), 10) || 0;
    const keys = [key + '#n'];
    for (let i = 0; i < n; i++) keys.push(key + '#' + i);
    cache.removeAll(keys);
  } catch (err) {}
}

/* ------------------------------------------------------------------ */
/*  DATA JADWAL                                                        */
/* ------------------------------------------------------------------ */

function getScheduleData_() {
  const cached = cacheGetBig_('SCHEDULE_DATA');
  if (cached) return cached;

  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const needCols = CONFIG.COL_TANGGAL_START + CONFIG.JUMLAH_TANGGAL - 1;
  const lastCol = Math.max(sheet.getLastColumn(), needCols);
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  const dayNameRow = values[1];
  const dateNumRow = values[2];

  const dateCols = [];
  for (let i = 0; i < CONFIG.JUMLAH_TANGGAL; i++) {
    const col = CONFIG.COL_TANGGAL_START + i;
    dateCols.push({
      colIndex: i,
      sheetCol: col,
      tanggal: dateNumRow[col - 1],
      hari: dayNameRow[col - 1],
    });
  }

  const employees = [];
  const storeSet = {};
  for (let r = CONFIG.DATA_START_ROW - 1; r < lastRow; r++) {
    const row = values[r];
    const nik = row[CONFIG.COL_NIK - 1];
    const nama = row[CONFIG.COL_NAMA - 1];
    if (!nik || !nama) continue;

    const lokasi = String(row[CONFIG.COL_LOKASI - 1] || '').trim();
    const jabatan = row[CONFIG.COL_JABATAN - 1];
    const shifts = [];
    for (let i = 0; i < CONFIG.JUMLAH_TANGGAL; i++) {
      const raw = row[CONFIG.COL_TANGGAL_START + i - 1];
      shifts.push(raw ? String(raw).trim() : '');
    }
    employees.push({ nik: String(nik), nama: String(nama).trim(), jabatan: jabatan, lokasi: lokasi, shifts: shifts });
    if (lokasi) storeSet[lokasi] = true;
  }

  const data = {
    dateCols: dateCols,
    employees: employees,
    stores: Object.keys(storeSet).sort(),
    generatedAt: new Date().toISOString(),
  };
  cachePutBig_('SCHEDULE_DATA', data, CONFIG.CACHE_SECONDS);
  return data;
}

function classifyShift_(code) {
  if (!code) return { group: 'kosong', label: '-', jam: '' };
  const c = String(code).trim().toUpperCase();

  if (c === 'O') return { group: 'libur', label: 'Libur', jam: '' };
  if (c === 'C') return { group: 'cuti', label: 'Cuti / Izin', jam: '' };
  if (c === 'AO') return { group: 'libur_tambahan', label: 'Libur Tambahan', jam: '' };

  // P7 / S15 / M23  (angka jam mulai opsional)
  const m = c.match(/^([PSM])(\d{0,2})$/);
  if (m) {
    const jam = m[2] ? m[2].padStart(2, '0') + ':00' : '';
    const group = m[1] === 'P' ? 'pagi' : (m[1] === 'S' ? 'siang' : 'malam');
    const label = group === 'pagi' ? 'Pagi' : (group === 'siang' ? 'Siang' : 'Malam');
    return { group: group, label: label, jam: jam };
  }

  // Cadangan: kode angka shift 1 / 2 / 3
  if (c === '1') return { group: 'pagi', label: 'Pagi', jam: '' };
  if (c === '2') return { group: 'siang', label: 'Siang', jam: '' };
  if (c === '3') return { group: 'malam', label: 'Malam', jam: '' };

  return { group: 'lainnya', label: c, jam: '' };
}

function namaHariPanjang_(kodeHari) {
  const map = { Sen: 'Senin', Sel: 'Selasa', Rab: 'Rabu', Kam: 'Kamis', Jum: "Jum'at", Sab: 'Sabtu', Min: 'Minggu' };
  return map[kodeHari] || kodeHari;
}

/* ------------------------------------------------------------------ */
/*  PERIODE                                                            */
/* ------------------------------------------------------------------ */

function getPeriodSettings() {
  const startDate = PropertiesService.getScriptProperties().getProperty('PERIOD_START_DATE');
  return { startDate: startDate || null };
}

function savePeriodSettings(startDateStr) {
  const props = PropertiesService.getScriptProperties();
  if (!startDateStr) props.deleteProperty('PERIOD_START_DATE');
  else props.setProperty('PERIOD_START_DATE', startDateStr);
  return getPeriodSettings();
}

/**
 * Menentukan kolom "hari ini" TANPA bergantung pada pengaturan manual:
 * 1) kolom yang angka tanggalnya sama dengan tanggal hari ini (prioritas yang
 *    nama harinya juga cocok), 2) cadangan: tanggal awal di Pengaturan.
 * Tanggal kalender kolom lain dihitung mundur/maju dari hari ini.
 */
function resolveToday_(data, period) {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const todayMid = new Date(Utilities.formatDate(now, tz, 'yyyy-MM-dd') + 'T00:00:00');
  const day = parseInt(Utilities.formatDate(now, tz, 'd'), 10);
  const kodeHari = ['', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab', 'Min'][parseInt(Utilities.formatDate(now, tz, 'u'), 10)];

  const cands = data.dateCols.filter(function (d) { return parseInt(d.tanggal, 10) === day; });
  if (cands.length) {
    const exact = cands.filter(function (d) { return String(d.hari).trim() === kodeHari; })[0];
    const idx = (exact || cands[0]).colIndex;
    const base = new Date(todayMid.getTime());
    base.setDate(base.getDate() - idx);
    return { idx: idx, base: base };
  }
  if (period && period.startDate) {
    const base = new Date(period.startDate + 'T00:00:00');
    const diff = Math.round((todayMid - base) / 86400000);
    return { idx: (diff >= 0 && diff < CONFIG.JUMLAH_TANGGAL) ? diff : 0, base: base };
  }
  return { idx: 0, base: null };
}

function realIso_(base, ci) {
  if (!base) return '';
  const d = new Date(base.getTime());
  d.setDate(d.getDate() + ci);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/* ------------------------------------------------------------------ */
/*  LOG FOTO                                                           */
/* ------------------------------------------------------------------ */

function getLogSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    sheet.appendRow([
      'Timestamp', 'Toko', 'Tanggal Periode', 'Hari', 'Tanggal Real',
      'Shift/Keterangan', 'Catatan', 'Uploader', 'File ID', 'URL Foto', 'Device Time'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Baris log ringkas (dengan cache). Urutan: [ts, toko, tanggal, ket, catatan, fileId, deviceTime] */
function getLogRows_() {
  const cached = cacheGetBig_('LOG_ROWS');
  if (cached) return cached;

  const sheet = getLogSheet_();
  const last = sheet.getLastRow();
  const rows = [];
  if (last >= 2) {
    const v = sheet.getRange(2, 1, last - 1, 11).getValues();
    for (let i = 0; i < v.length; i++) {
      const r = v[i];
      if (!r[8]) continue;
      rows.push([
        r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
        String(r[1] || '').trim(),
        String(r[2]),
        String(r[5] || ''),
        String(r[6] || ''),
        String(r[8]),
        String(r[10] || ''),
      ]);
    }
  }
  cachePutBig_('LOG_ROWS', rows, CONFIG.LOG_CACHE_SECONDS);
  return rows;
}

function getRootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('ROOT_FOLDER_ID');
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (err) {}
  }
  const it = DriveApp.getFoldersByName(CONFIG.DRIVE_ROOT_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.DRIVE_ROOT_FOLDER);
  try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
  props.setProperty('ROOT_FOLDER_ID', folder.getId());
  return folder;
}

/* ------------------------------------------------------------------ */
/*  BUNDLE PER HARI (semua toko sekaligus)                             */
/* ------------------------------------------------------------------ */

function buildBundle_(colIndex, data, base) {
  const ci = Math.max(0, Math.min(CONFIG.JUMLAH_TANGGAL - 1, Number(colIndex) || 0));
  const dateCol = data.dateCols[ci];

  const byStore = {};
  const stores = data.stores.map(function (name) {
    const s = { toko: name, pagi: [], siang: [], malam: [], lain: [], photos: [] };
    byStore[name] = s;
    return s;
  });

  data.employees.forEach(function (e) {
    const st = byStore[e.lokasi];
    if (!st) return;
    const code = e.shifts[ci];
    const info = classifyShift_(code);
    if (info.group === 'kosong') return;
    const item = { n: e.nama, j: e.jabatan || '', k: code, t: info.jam };
    if (info.group === 'pagi' || info.group === 'siang' || info.group === 'malam') st[info.group].push(item);
    else st.lain.push(item);
  });

  const tgl = String(dateCol.tanggal);
  const rows = getLogRows_();
  for (let i = rows.length - 1; i >= 0; i--) { // terbaru dulu
    const r = rows[i];
    if (r[2] !== tgl) continue;
    const st = byStore[r[1]];
    if (!st) continue;
    st.photos.push({ id: r[5], ket: r[3], cat: r[4], dt: r[6], ts: r[0] });
  }

  return {
    ci: ci,
    tanggal: dateCol.tanggal,
    hari: namaHariPanjang_(dateCol.hari),
    tanggalReal: realIso_(base, ci),
    stores: stores,
  };
}

function getDayBundle(colIndex) {
  const data = getScheduleData_();
  const today = resolveToday_(data, getPeriodSettings());
  return buildBundle_(colIndex, data, today.base);
}

function getInitialAppData() {
  const data = getScheduleData_();
  const period = getPeriodSettings();
  const today = resolveToday_(data, period);

  const realMap = {};
  data.dateCols.forEach(function (d) { realMap[d.colIndex] = realIso_(today.base, d.colIndex); });

  return {
    stores: data.stores,
    dateCols: data.dateCols,
    employees: data.employees,
    periode: period,
    todayColIndex: today.idx,
    tanggalRealMap: realMap,
    bundle: buildBundle_(today.idx, data, today.base),
  };
}

function refreshAll() {
  cacheRemoveBig_('SCHEDULE_DATA');
  cacheRemoveBig_('LOG_ROWS');
  return getInitialAppData();
}

/* ------------------------------------------------------------------ */
/*  UPLOAD & HAPUS FOTO                                                */
/* ------------------------------------------------------------------ */

function uploadAttendancePhoto(payload) {
  // Cegah foto ganda jika HP mengirim ulang (koneksi putus-nyambung)
  const cache = CacheService.getScriptCache();
  const dupKey = payload.clientId ? 'UP_' + payload.clientId : '';
  if (dupKey) {
    const hit = cache.get(dupKey);
    if (hit) return { fileId: hit };
  }

  const root = getRootFolder_();
  const blob = Utilities.newBlob(
    Utilities.base64Decode(payload.base64),
    payload.mimeType || 'image/jpeg',
    payload.storeName + '_Tgl' + payload.tanggal + '_' + Date.now() + '.jpg'
  );

  const file = root.createFile(blob);
  const fileId = file.getId();
  if (CONFIG.SHARE_PER_FILE) {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }

  getLogSheet_().appendRow([
    new Date(),
    payload.storeName,
    payload.tanggal,
    payload.hari,
    payload.tanggalReal || '',
    payload.keterangan || '',
    payload.catatan || '',
    '',
    fileId,
    'https://drive.google.com/file/d/' + fileId + '/view',
    payload.deviceTime || ''
  ]);

  cacheRemoveBig_('LOG_ROWS');
  if (dupKey) { try { cache.put(dupKey, fileId, 1800); } catch (err) {} }
  return { fileId: fileId };
}

function deletePhoto(fileId) {
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (err) {}

  const sheet = getLogSheet_();
  const last = sheet.getLastRow();
  if (last >= 2) {
    const hit = sheet.getRange(2, 9, last - 1, 1).createTextFinder(fileId).matchEntireCell(true).findNext();
    if (hit) sheet.deleteRow(hit.getRow());
  }
  cacheRemoveBig_('LOG_ROWS');
  return true;
}
