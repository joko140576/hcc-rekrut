/**
 * BACKEND DASHBOARD PEMENUHAN PENGGANTI  (v2 - lebih stabil & cepat)
 * Google Apps Script Web App (anti-CORS)
 *
 * Cara deploy:
 * 1. Buka spreadsheet > Extensions > Apps Script, tempel kode ini.
 * 2. Deploy > New deployment > Type: Web app
 *    Execute as: Me | Who has access: Anyone
 * 3. Salin URL /exec ke konstanta GAS_URL di index.html.
 * Setiap kali kode diubah: Deploy > Manage deployments > Edit > Version: New version > Deploy
 * (URL /exec tetap sama).
 */

const SPREADSHEET_ID = '1uOyBYbSVFhprHvn7-eUu8INgrwR4nXyHr3WN7nRCZyA';
const SHEET_NAME = 'pegawai-pengganti'; // dibuat otomatis bila belum ada
const PROJECT_SHEET = 'project';
const PEMENUHAN_SHEET = 'pegawai-project'; // nama tab sesuai spreadsheet
const HEADERS = ['Nama', 'NPP', 'Tanggal Keluar', 'Status Keluar', 'Kantor',
  'No Surat Nitro', 'Tanggal Surat', 'Status Pengganti', 'Nama Pengganti', 'Tanggal Masuk'];

const CACHE_TTL = 300;       // detik; cache dihapus otomatis saat ada tambah/hapus data
const CACHE_PREFIX = 'list_v2';
const CHUNK = 30000;         // batas CacheService 100KB/nilai -> dipecah per 30.000 karakter

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function fmt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v).trim();
}

/* ---------- Cache server (CacheService) ---------- */
function cacheGet_() {
  try {
    const c = CacheService.getScriptCache();
    const n = Number(c.get(CACHE_PREFIX + '_n') || 0);
    if (!n) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(CACHE_PREFIX + '_' + i);
    const parts = c.getAll(keys);
    let s = '';
    for (let i = 0; i < n; i++) {
      const p = parts[keys[i]];
      if (p === null || p === undefined) return null;
      s += p;
    }
    return s;
  } catch (e) { return null; }
}

function cachePut_(str) {
  try {
    const c = CacheService.getScriptCache();
    const n = Math.ceil(str.length / CHUNK);
    const o = {};
    for (let i = 0; i < n; i++) o[CACHE_PREFIX + '_' + i] = str.substr(i * CHUNK, CHUNK);
    o[CACHE_PREFIX + '_n'] = String(n);
    c.putAll(o, CACHE_TTL);
  } catch (e) { /* cache gagal tidak boleh menggagalkan respons */ }
}

function cacheClear_() {
  try {
    const c = CacheService.getScriptCache();
    const n = Number(c.get(CACHE_PREFIX + '_n') || 0);
    const keys = [CACHE_PREFIX + '_n'];
    for (let i = 0; i < n; i++) keys.push(CACHE_PREFIX + '_' + i);
    c.removeAll(keys);
  } catch (e) {}
}

/* ---------- PEMENUHAN PROJECT (sheet "project" & "pemenuhan project") ---------- */
// Struktur kolom dibaca dari baris 1 (header) masing-masing sheet, jadi tidak ditulis ulang di sini.
function normName_(s) { return String(s).toLowerCase().replace(/[^a-z0-9]/g, ''); }

function findSheet_(name) {
  const all = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets();
  const sh = all.filter(function (s) { return normName_(s.getName()) === normName_(name); })[0];
  if (!sh) throw new Error('Sheet "' + name + '" tidak ditemukan. Sheet yang ada: ' + all.map(function (s) { return s.getName(); }).join(', '));
  return sh;
}

function tabOf_(kind) { return findSheet_(kind === 'prj' ? PROJECT_SHEET : PEMENUHAN_SHEET); }

function readTab_(sh) {
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (!lr || !lc) return { headers: [], rows: [] };
  const vals = sh.getRange(1, 1, lr, lc).getValues();
  const rows = [];
  for (let i = 1; i < vals.length; i++) {
    const v = vals[i].map(fmt_);
    if (v.some(function (x) { return x !== ''; })) rows.push({ row: i + 1, v: v });
  }
  return { headers: vals[0].map(fmt_), rows: rows };
}

function readProjects_() {
  return { ok: true, project: readTab_(tabOf_('prj')), pemenuhan: readTab_(tabOf_('pmp')), updated: new Date().getTime() };
}

function addTab_(sh, values) {
  const lc = sh.getLastColumn();
  const hs = sh.getRange(1, 1, 1, lc).getValues()[0].map(fmt_);
  const row = hs.map(function (h) {
    const v = values ? values[h] : '';
    return v === undefined || v === null ? '' : String(v).trim().replace(/\s+/g, ' ');
  });
  if (!row.some(function (x) { return x !== ''; })) throw new Error('Data kosong');
  if (sh.getName() === findSheet_(PROJECT_SHEET).getName() && sh.getLastRow() > 1) {
    // Kombinasi project + kantor tidak boleh duplikat
    let pk = hs.findIndex(function (h) { return /^(nama\s*)?project$/i.test(h); });
    if (pk < 0) pk = hs.findIndex(function (h) { return /project/i.test(h); });
    if (pk < 0) pk = 0;
    const kk = hs.findIndex(function (h, i) { return i !== pk && /kantor|cabang|lokasi|wilayah|area/i.test(h); });
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, lc).getValues();
    const keyOf = function (r) { return normName_(r[pk]) + '|' + (kk >= 0 ? normName_(r[kk]) : ''); };
    const mine = keyOf(row);
    if (row[pk] && data.some(function (r) { return keyOf(r) === mine; })) throw new Error('Project untuk kantor ini sudah ada, tidak boleh duplikat');
  }
  const next = sh.getLastRow() + 1;
  hs.forEach(function (h, i) { if (/npp|nip|nik/i.test(h)) sh.getRange(next, i + 1).setNumberFormat('@'); });
  sh.getRange(next, 1, 1, lc).setValues([row.map(function (x, i) {
    return /kebutuhan|kuota|target|jumlah|qty/i.test(hs[i]) && x !== '' && !isNaN(Number(x)) ? Number(x) : x;
  })]);
}

function deleteTab_(sh, rowNo, check) {
  if (!(rowNo > 1) || rowNo > sh.getLastRow()) throw new Error('Baris tidak valid, muat ulang data');
  const cur = sh.getRange(rowNo, 1, 1, sh.getLastColumn()).getValues()[0].map(fmt_).join('|');
  if (check !== undefined && check !== null && String(check) !== cur) throw new Error('Data sudah berubah, muat ulang lalu coba lagi');
  sh.deleteRow(rowNo);
}

/* ---------- GET ?action=list[&fresh=1] ---------- */
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'projects') return json_(readProjects_());
    const fresh = e && e.parameter && e.parameter.fresh === '1';
    if (!fresh) {
      const hit = cacheGet_();
      if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
    }

    const sh = getSheet_();
    const last = sh.getLastRow();
    let rows = [];
    if (last > 1) {
      rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues().map(function (r, i) {
        return {
          row: i + 2,
          nama: fmt_(r[0]), npp: fmt_(r[1]), tanggalKeluar: fmt_(r[2]),
          statusKeluar: fmt_(r[3]), kantor: fmt_(r[4]), noSurat: fmt_(r[5]),
          tanggalSurat: fmt_(r[6]), statusPengganti: fmt_(r[7]),
          namaPengganti: fmt_(r[8]), tanggalMasuk: fmt_(r[9])
        };
      }).filter(function (r) { return r.nama || r.npp; });
    }
    const str = JSON.stringify({ ok: true, data: rows, updated: new Date().getTime() });
    cachePut_(str);
    return ContentService.createTextOutput(str).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ---------- POST (text/plain) body JSON: {action:'add'|'delete', ...} ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    lock.waitLock(20000);
    locked = true;
    const p = JSON.parse(e.postData.contents);
    const sh = getSheet_();

    if (p.action === 'add') {
      const row = [p.nama, p.npp, p.tanggalKeluar, p.statusKeluar, p.kantor,
        p.noSurat, p.tanggalSurat, p.statusPengganti, p.namaPengganti, p.tanggalMasuk]
        .map(function (v) { return v === undefined || v === null ? '' : String(v).trim().replace(/\s+/g, ' '); });
      if (!row[0] || !row[1]) return json_({ ok: false, error: 'Nama dan NPP wajib diisi' });
      const next = sh.getLastRow() + 1;
      const rng = sh.getRange(next, 1, 1, HEADERS.length);
      rng.setNumberFormat('@'); // simpan sebagai teks (NPP tidak kehilangan angka 0 di depan)
      rng.setValues([row]);
      SpreadsheetApp.flush();
      cacheClear_();
      return json_({ ok: true, message: 'Data tersimpan' });
    }

    if (p.action === 'delete') {
      const rowNo = Number(p.row);
      if (!(rowNo > 1) || rowNo > sh.getLastRow()) return json_({ ok: false, error: 'Baris tidak valid, muat ulang data' });
      // Pengaman: pastikan baris yang dihapus masih milik data yang sama
      // (nomor baris bisa bergeser jika ada yang menghapus lebih dulu).
      if (p.npp !== undefined && p.npp !== null) {
        const cur = sh.getRange(rowNo, 1, 1, 2).getValues()[0];
        if (String(cur[1]).trim() !== String(p.npp).trim() || String(cur[0]).trim() !== String(p.nama || '').trim()) {
          cacheClear_();
          return json_({ ok: false, error: 'Data sudah berubah, muat ulang lalu coba lagi' });
        }
      }
      sh.deleteRow(rowNo);
      SpreadsheetApp.flush();
      cacheClear_();
      return json_({ ok: true, message: 'Data dihapus' });
    }

    if (p.action === 'padd') {
      addTab_(tabOf_(p.sheet), p.values);
      SpreadsheetApp.flush();
      return json_({ ok: true, message: 'Data tersimpan' });
    }

    if (p.action === 'pdel') {
      deleteTab_(tabOf_(p.sheet), Number(p.row), p.check);
      SpreadsheetApp.flush();
      return json_({ ok: true, message: 'Data dihapus' });
    }

    return json_({ ok: false, error: 'Action tidak dikenal' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    if (locked) lock.releaseLock();
  }
}

// Opsional: jalankan sekali dari editor ini bila ingin menghapus cache secara manual
function clearCache() { cacheClear_(); }
