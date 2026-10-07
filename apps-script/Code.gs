/**
 * Code.gs — gabungan semua file .gs (Config, Util, Setup, Transaksi, Alur, Web).
 * Tempel utuh ke satu file Code.gs di Apps Script. Urutan bagian jangan diubah.
 */

/* ======================= Config ======================= */

/**
 * Config.gs — satu-satunya tempat definisi struktur data.
 * Mengubah nama tab, kolom, atau status cukup di sini.
 */
var CFG = {
  TZ: 'Asia/Jakarta',

  SHEET: { TRX: 'TRANSAKSI', LOG: 'LOG', USER: 'USER', CONFIG: 'CONFIG' },

  STATUS: {
    MASUK: 'MASUK',
    SPM: 'SPM_DIBAGI',
    LOADING: 'LOADING',
    SELESAI: 'SELESAI_MUAT',
    RECEIVED: 'RECEIVED',
    DONE: 'COMPLETED'
  },

  ROLE: {
    SECURITY: 'SECURITY',
    JURU: 'JURU_MUAT',
    CHECKER: 'CHECKER'
  },

  COLS: {
    TRX: [
      'ID', 'Tanggal', 'No_Antrean', 'No_Polisi', 'Sopir', 'Tujuan',
      'Jam_Masuk', 'Dicatat_Security',
      'No_SPM', 'Jam_Dibagi', 'PIC_Muat',
      'Jam_Start', 'Jam_Finish', 'Juru_Muat',
      'Jam_Receive', 'Checker', 'Jam_Completed',
      'Status', 'Exception', 'Alasan_Exception', 'Input_Susulan',
      'Waktu_Ke_SPM', 'Waiting', 'Loading', 'Lead_Time', 'Over_Target', 'FIFO_OK',
      'Urut_Bagi_SPM', 'Urut_Muat', 'Selisih_SPM', 'Selisih_Muat', 'Tanda'
    ],
    LOG: ['Timestamp', 'ID', 'Status_Lama', 'Status_Baru', 'Pelaku', 'Peran', 'Catatan'],
    USER: ['Nama', 'Peran', 'Aktif'],
    CONFIG: ['Kunci', 'Nilai', 'Keterangan']
  },

  TEXT_COLS: ['ID', 'Tanggal', 'No_Polisi', 'No_SPM'],
  TIME_COLS: ['Jam_Masuk', 'Jam_Dibagi', 'Jam_Start', 'Jam_Finish', 'Jam_Receive', 'Jam_Completed'],

  DEFAULT_CONFIG: [
    ['TARGET_WAITING_MENIT', 30, 'Batas waktu tunggu (Start Loading - Jam Masuk) sebelum dihitung Over Target'],
    ['TARGET_LOADING_MENIT', 60, 'Batas durasi loading (Receive - Start)'],
    ['TARGET_LEAD_MENIT', 120, 'Batas lead time total (Receive - Masuk)'],
    ['ADMIN_EMAILS', '', 'Email admin, pisahkan dengan koma']
  ],

  SEED_USER: [
    ['Contoh Security', 'SECURITY', 'Ya'],
    ['Contoh Juru Muat', 'JURU_MUAT', 'Ya'],
    ['Contoh Checker', 'CHECKER', 'Ya']
  ]
};

/** Alur status: setiap aksi hanya sah dari status tertentu oleh peran tertentu. */
CFG.FLOW = {
  start:   { from: 'MASUK',        to: 'LOADING',      role: 'JURU_MUAT', time: 'Jam_Start',     actor: 'Juru_Muat' },
  receive: { from: 'LOADING',      to: 'RECEIVED',     role: 'CHECKER',   time: 'Jam_Receive',   actor: 'Checker' },
  done:    { from: 'RECEIVED',     to: 'COMPLETED',    role: 'CHECKER',   time: 'Jam_Completed' }
};

/** Kinerja: lama simpan hasil baca USER/CONFIG (detik) dan jumlah baris TRANSAKSI terakhir yang dibaca. */
CFG.CACHE_SEC = 60;
CFG.TAIL_ROWS = 300;

/** Tab REKAP: tabel per tanggal berisi rumus (diisi dari Setup.gs). */
CFG.REKAP = { SHEET: 'REKAP', FIRST_ROW: 13, ROWS: 80 };

/* ======================= Util ======================= */

/**
 * Util.gs — fungsi bantu umum. Tidak berisi aturan bisnis.
 * memo_ hidup hanya selama satu eksekusi (satu klik), jadi tidak pernah basi antar klik.
 * Pembacaan USER dan CONFIG juga disimpan di CacheService selama CFG.CACHE_SEC detik.
 */

var memo_ = { ss: null, sheets: {}, cfg: null, users: null };

/** Tunggu kunci; bila sibuk, beri pesan yang mudah dipahami (bukan error teknis). */
function waitLock_(lock) {
  try { lock.waitLock(10000); }
  catch (e) { throw new Error('Sistem sedang sibuk. Tunggu sebentar lalu tekan lagi.'); }
}

function resetMemo_() {
  memo_ = { ss: null, sheets: {}, cfg: null, users: null };
}

function sheet_(name) {
  if (!memo_.sheets[name]) {
    memo_.ss = memo_.ss || SpreadsheetApp.getActive();
    var s = memo_.ss.getSheetByName(name);
    if (!s) throw new Error('Tab ' + name + ' belum ada. Jalankan Control Tower > Setup tab otomatis.');
    memo_.sheets[name] = s;
  }
  return memo_.sheets[name];
}

/**
 * Urutan kolom sebuah tab sistem (TRANSAKSI, LOG) diambil dari Config, tanpa membaca sheet.
 * Header di baris 1 tidak boleh diubah; menu Control Tower > Cek struktur memeriksanya.
 */
function head_(name) {
  var cols = null;
  Object.keys(CFG.SHEET).forEach(function (k) { if (CFG.SHEET[k] === name) cols = CFG.COLS[k]; });
  if (!cols) throw new Error('Tab ' + name + ' tidak dikenal.');
  return cols;
}

function fmt_(d, pattern) {
  return Utilities.formatDate(d, CFG.TZ, pattern);
}

function today_() {
  return fmt_(new Date(), 'yyyy-MM-dd');
}

function dateStr_(v) {
  return v instanceof Date ? fmt_(v, 'yyyy-MM-dd') : String(v || '');
}

function hm_(v) {
  return v instanceof Date ? fmt_(v, 'HH:mm') : '';
}

function pad_(n) {
  return n < 10 ? '0' + n : String(n);
}

function toObjects_(head, values, firstRow) {
  return values.map(function (r, i) {
    var o = { _row: firstRow + i };
    head.forEach(function (k, j) { o[k] = r[j]; });
    return o;
  });
}

/** Baca seluruh baris sebuah tab kecil (USER, CONFIG) sebagai array objek. */
function readRows_(name) {
  var v = sheet_(name).getDataRange().getValues();
  var head = v.shift();
  return toObjects_(head, v, 2);
}

/**
 * Baca hanya n baris terakhir sebuah tab besar (TRANSAKSI).
 * Data selalu ditambah di bawah, jadi data hari ini ada di ujung.
 * Waktu baca tetap sama berapa pun lamanya sistem dipakai.
 */
function readTail_(name, n) {
  var s = sheet_(name);
  var last = s.getLastRow();
  if (last < 2) return [];
  var head = head_(name);
  var start = Math.max(2, last - n + 1);
  return toObjects_(head, s.getRange(start, 1, last - start + 1, head.length).getValues(), start);
}

/** Seperti readRows_, tetapi memakai CacheService untuk tab kecil yang jarang berubah. */
function cachedRows_(name) {
  var cache = CacheService.getScriptCache();
  var key = 'rows:' + name;
  var hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  var rows = readRows_(name);
  cache.put(key, JSON.stringify(rows), CFG.CACHE_SEC);
  return rows;
}

/** Tambah satu baris berdasarkan objek; kolom yang tidak diisi dibiarkan kosong. */
function appendObj_(name, obj) {
  sheet_(name).appendRow(head_(name).map(function (k) { return obj[k] === undefined ? '' : obj[k]; }));
}

/**
 * Ubah satu baris dengan satu kali tulis.
 * base (opsional) = objek baris lengkap yang sudah ada di memori, supaya tidak perlu dibaca ulang.
 */
function updateObj_(name, row, obj, base) {
  var s = sheet_(name), head = head_(name);
  var cur = base;
  if (!cur) {
    cur = {};
    var v = s.getRange(row, 1, 1, head.length).getValues()[0];
    head.forEach(function (k, j) { cur[k] = v[j]; });
  }
  var vals = head.map(function (k) {
    var x = obj[k] !== undefined ? obj[k] : cur[k];
    return x === undefined ? '' : x;
  });
  s.getRange(row, 1, 1, head.length).setValues([vals]);
}

/** Ambil nilai CONFIG berdasarkan kunci. */
function cfg_(key, fallback) {
  if (!memo_.cfg) memo_.cfg = cachedRows_(CFG.SHEET.CONFIG);
  for (var i = 0; i < memo_.cfg.length; i++) {
    if (memo_.cfg[i].Kunci === key && memo_.cfg[i].Nilai !== '') return memo_.cfg[i].Nilai;
  }
  return fallback;
}

/** Nomor kolom (1, 2, 3 ...) menjadi huruf (A, B, C ...). */
function colLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/* ======================= Setup ======================= */

/**
 * Setup.gs — membuat 4 tab beserta header, format, dan data awal.
 * Aman dijalankan ulang: tab dan data yang sudah ada tidak ditimpa.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Control Tower')
    .addItem('Setup tab otomatis', 'setup')
    .addItem('Perbarui REKAP', 'buatRekap_')
    .addItem('Cek struktur', 'cekStruktur')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActive();

  Object.keys(CFG.SHEET).forEach(function (key) {
    var name = CFG.SHEET[key];
    var s = ss.getSheetByName(name) || ss.insertSheet(name);
    if (s.getLastRow() === 0) {
      var head = CFG.COLS[key];
      s.getRange(1, 1, 1, head.length).setValues([head])
        .setFontWeight('bold').setBackground('#e8eef0');
      s.setFrozenRows(1);
    }
  });

  tambahKolom_();
  seed_(CFG.SHEET.USER, CFG.SEED_USER);
  seed_(CFG.SHEET.CONFIG, CFG.DEFAULT_CONFIG);
  formatTrx_();
  buatRekap_();
  CacheService.getScriptCache().removeAll(['rows:' + CFG.SHEET.USER, 'rows:' + CFG.SHEET.CONFIG]);

  [CFG.SHEET.TRX, CFG.SHEET.LOG].forEach(function (name) {
    var s = sheet_(name);
    if (!s.getProtections(SpreadsheetApp.ProtectionType.SHEET).length) {
      s.protect().setDescription('Diubah oleh sistem').setWarningOnly(true);
    }
  });

  SpreadsheetApp.getUi().alert(
    'Setup selesai.\n\nLangkah berikutnya:\n' +
    '1. Ganti nama contoh di tab USER dengan nama asli.\n' +
    '2. Isi ADMIN_EMAILS di tab CONFIG.\n' +
    '3. Buka tab REKAP, pilih tanggal di sel B3.\n' +
    '4. Deploy sebagai Web app.'
  );
}

function seed_(name, rows) {
  var s = sheet_(name);
  if (s.getLastRow() > 1) return;
  s.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function formatTrx_() {
  var s = sheet_(CFG.SHEET.TRX);
  var head = CFG.COLS.TRX;
  // Format per kolom penuh, supaya baris baru yang ditambah sistem ikut berformat teks atau jam.
  CFG.TEXT_COLS.forEach(function (c) {
    var L = colLetter_(head.indexOf(c) + 1);
    s.getRange(L + ':' + L).setNumberFormat('@');
  });
  CFG.TIME_COLS.forEach(function (c) {
    var L = colLetter_(head.indexOf(c) + 1);
    s.getRange(L + ':' + L).setNumberFormat('HH:mm:ss');
  });
}

/**
 * Tambah kolom baru di ujung kanan pada tab lama (tanpa menghapus atau menggeser data).
 * Hanya dikerjakan bila header lama persis sama dengan awal daftar kolom di Config.
 */
function tambahKolom_() {
  Object.keys(CFG.SHEET).forEach(function (k) {
    var s = sheet_(CFG.SHEET[k]), want = CFG.COLS[k];
    var have = s.getRange(1, 1, 1, Math.max(s.getLastColumn(), 1)).getValues()[0];
    while (have.length && have[have.length - 1] === '') have.pop();
    var prefix = have.length < want.length && have.every(function (c, i) { return c === want[i]; });
    if (!prefix) return;
    var add = want.slice(have.length);
    s.getRange(1, have.length + 1, 1, add.length).setValues([add]).setFontWeight('bold').setBackground('#e8eef0');
  });
}

/** Isi sel REKAP berupa rumus murni (tanpa script), supaya selalu mengikuti TRANSAKSI. */
function rekapFormulas_() {
  var T = CFG.SHEET.TRX, cols = CFG.COLS.TRX;
  var L = function (n) { return colLetter_(cols.indexOf(n) + 1); };
  var whole = function (n) { return T + '!$' + L(n) + ':$' + L(n); };
  var part = function (n) { return T + '!$' + L(n) + '$2:$' + L(n) + '$20000'; };
  var cnt = function (tanda) {
    return '=SUMPRODUCT(--(' + part('Tanggal') + '=$D$3)' + (tanda ? ',--(' + part('Tanda') + '="' + tanda + '")' : '') + ')';
  };
  var pick = ['No_Antrean', 'No_Polisi', 'Jam_Masuk', 'Jam_Start', 'Urut_Muat', 'Selisih_Muat', 'Tanda', 'Status'];
  var R0 = CFG.REKAP.FIRST_ROW, rows = [];
  for (var k = 0; k < CFG.REKAP.ROWS; k++) {
    rows.push(pick.map(function (n, j) {
      var at = 'INDEX(' + whole(n) + ',$B$4+' + k + ')';
      if (j === 0) return '=IF($B$4=0,"",IF(INDEX(' + whole('Tanggal') + ',$B$4+' + k + ')=$D$3,' + at + ',""))';
      return '=IF($A' + (R0 + k) + '="","",IF(' + at + '="","",' + at + '))';
    }));
  }
  return {
    B3: '=TODAY()',
    D3: '=TEXT(B3,"yyyy-mm-dd")',
    B4: '=IFERROR(MATCH($D$3,' + whole('Tanggal') + ',0),0)',
    summary: [
      ['Total mobil', cnt('')], ['Sesuai urutan', cnt('SESUAI')], ['Lompat antrean', cnt('LOMPAT')],
      ['Tertunda', cnt('TERTUNDA')], ['Beda urutan', cnt('BEDA_URUT')]
    ],
    header: ['Datang ke-', 'No. polisi', 'Jam masuk', 'Jam mulai muat', 'Muat ke-', 'Selisih', 'Tanda', 'Status'],
    rows: rows
  };
}

/** Buat atau perbarui tab REKAP. Aman dijalankan ulang. */
function buatRekap_() {
  var ss = SpreadsheetApp.getActive();
  var s = ss.getSheetByName(CFG.REKAP.SHEET) || ss.insertSheet(CFG.REKAP.SHEET);
  var f = rekapFormulas_(), R0 = CFG.REKAP.FIRST_ROW;

  s.getRange('A1').setValue('Rekap harian: urutan datang dan urutan muat').setFontWeight('bold').setFontSize(14);
  s.getRange('A3').setValue('Pilih tanggal');
  s.getRange('B3').setFormula(f.B3).setNumberFormat('yyyy-mm-dd').setBackground('#fff9c4');
  s.getRange('C3').setValue('teks tanggal');
  s.getRange('D3').setFormula(f.D3);
  s.getRange('A4').setValue('Baris awal di TRANSAKSI');
  s.getRange('B4').setFormula(f.B4);
  f.summary.forEach(function (x, i) {
    s.getRange(6 + i, 1).setValue(x[0]);
    s.getRange(6 + i, 2).setFormula(x[1]);
  });
  s.getRange(R0 - 1, 1, 1, f.header.length).setValues([f.header])
    .setFontWeight('bold').setBackground('#e8eef0');
  s.getRange(R0 - 1, 1, f.rows.length + 1, 12).clearFormat();
  s.getRange(R0 - 1, f.header.length + 1, f.rows.length + 1, 12 - f.header.length).clearContent();
  s.getRange(R0 - 1, 1, 1, f.header.length).setFontWeight('bold').setBackground('#e8eef0');
  s.getRange(R0, 1, f.rows.length, f.header.length).setFormulas(f.rows);
  ['C', 'D'].forEach(function (L) { s.getRange(L + R0 + ':' + L + (R0 + f.rows.length - 1)).setNumberFormat('HH:mm'); });

  var rng = s.getRange('G' + R0 + ':G' + (R0 + f.rows.length - 1));
  var rule = function (txt, bg, fg) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(txt)
      .setBackground(bg).setFontColor(fg).setRanges([rng]).build();
  };
  s.setConditionalFormatRules([
    rule('SESUAI', '#e6f4ea', '#137333'), rule('LOMPAT', '#fef3e0', '#9a5b00'),
    rule('TERTUNDA', '#fce8e6', '#c5221f'), rule('BEDA_URUT', '#f3e8fd', '#7627bb')
  ]);
}

/** Pastikan header 4 tab sama dengan Config. Jalankan bila ada yang menyisipkan atau memindah kolom. */
function cekStruktur() {
  var bad = [];
  var ss = SpreadsheetApp.getActive();
  Object.keys(CFG.SHEET).forEach(function (k) {
    var s = ss.getSheetByName(CFG.SHEET[k]);
    if (!s) { bad.push(CFG.SHEET[k] + ': tab tidak ada'); return; }
    var want = CFG.COLS[k];
    var have = s.getRange(1, 1, 1, want.length).getValues()[0];
    want.forEach(function (c, i) {
      if (have[i] !== c) bad.push(CFG.SHEET[k] + ' kolom ' + (i + 1) + ': seharusnya ' + c + ', ditemukan ' + have[i]);
    });
  });
  SpreadsheetApp.getUi().alert(bad.length ? 'Struktur tidak sesuai:\n' + bad.join('\n') : 'Struktur 4 tab sesuai.');
}

/* ======================= Transaksi ======================= */

/**
 * Transaksi.gs — aturan bisnis. Semua perubahan data lewat sini.
 * Tahap 1: pencatatan kendaraan masuk (Security).
 */

function userNames_(role) {
  if (!memo_.users) memo_.users = cachedRows_(CFG.SHEET.USER);
  return memo_.users
    .filter(function (u) { return u.Peran === role && u.Aktif === 'Ya'; })
    .map(function (u) { return u.Nama; });
}

function assertUser_(name, role) {
  if (userNames_(role).indexOf(name) === -1) {
    throw new Error('Nama tidak terdaftar atau tidak aktif untuk peran ' + role + '.');
  }
}

function log_(id, oldStatus, newStatus, actor, role, note) {
  appendObj_(CFG.SHEET.LOG, {
    Timestamp: new Date(), ID: id, Status_Lama: oldStatus, Status_Baru: newStatus,
    Pelaku: actor, Peran: role, Catatan: note || ''
  });
}

function byQueue_(a, b) {
  return a.No_Antrean - b.No_Antrean;
}

/** Baris pada tanggal tertentu dari kumpulan baris yang sudah dibaca, urut nomor antrean. */
function rowsOfDate_(all, date) {
  return all.filter(function (r) { return dateStr_(r.Tanggal) === date; }).sort(byQueue_);
}

function listByDate_(date) {
  return rowsOfDate_(readTail_(CFG.SHEET.TRX, CFG.TAIL_ROWS), date);
}

function nextQueueNo_(rows) {
  return rows.reduce(function (m, r) { return Math.max(m, Number(r.No_Antrean) || 0); }, 0) + 1;
}

function securityViewFrom_(date, rows) {
  return { tanggal: date, next: pad_(nextQueueNo_(rows)), list: rows.map(card_).reverse() };
}

/** Tampilan hari ini untuk halaman Security. */
function securityView_() {
  var date = today_();
  return securityViewFrom_(date, listByDate_(date));
}

/**
 * Catat kendaraan masuk.
 * p = { pelaku, nopol, sopir, tujuan, paksa }
 * Mengembalikan { dup: true } bila No. Polisi yang sama masih aktif hari ini,
 * kecuali p.paksa bernilai true.
 */
function recordEntry_(p) {
  var nopol = String(p.nopol || '').toUpperCase().replace(/\s+/g, ' ').trim();
  if (!nopol) throw new Error('No. polisi wajib diisi.');
  assertUser_(p.pelaku, CFG.ROLE.SECURITY);

  var lock = LockService.getScriptLock();
  waitLock_(lock);
  try {
    var date = today_();
    var rows = listByDate_(date);

    var active = rows.some(function (r) {
      return r.No_Polisi === nopol && r.Status !== CFG.STATUS.DONE;
    });
    if (active && !p.paksa) return { dup: true };

    var no = nextQueueNo_(rows);
    var id = date.replace(/-/g, '') + '-' + pad_(no);

    var rec = {
      ID: id, Tanggal: date, No_Antrean: no, No_Polisi: nopol,
      Sopir: String(p.sopir || '').trim(), Tujuan: String(p.tujuan || '').trim(),
      Jam_Masuk: new Date(), Dicatat_Security: p.pelaku,
      Status: CFG.STATUS.MASUK, Exception: 'Tidak', Input_Susulan: 'Tidak'
    };
    appendObj_(CFG.SHEET.TRX, rec);
    log_(id, '', CFG.STATUS.MASUK, p.pelaku, CFG.ROLE.SECURITY, '');

    rows.push(rec);
    return { ok: true, no: pad_(no), nopol: nopol, view: securityViewFrom_(date, rows) };
  } finally {
    lock.releaseLock();
  }
}

/* ======================= Alur ======================= */

/**
 * Alur.gs — tahap 2: perpindahan status, FIFO, hitungan waktu, tampilan per peran, KPI.
 * Semua perpindahan status lewat act_(), diatur oleh CFG.FLOW.
 */

function mins_(a, b) {
  return (a instanceof Date && b instanceof Date) ? Math.round((b - a) / 60000) : '';
}

/** Angka terbesar pada satu kolom (0 bila belum ada). */
function maxNum_(rows, key) {
  return rows.reduce(function (m, x) { return typeof x[key] === 'number' ? Math.max(m, x[key]) : m; }, 0);
}

/**
 * Bandingkan urutan aktual dengan No. Antrean (urutan datang).
 * Selisih negatif = dilayani lebih awal dari gilirannya, positif = tertunda.
 * Tanda: SESUAI, LOMPAT, TERTUNDA, atau BEDA_URUT (campuran). Kosong bila belum ada urutan aktual.
 */
function urut_(m) {
  var n = Number(m.No_Antrean);
  var o = {
    Selisih_SPM: typeof m.Urut_Bagi_SPM === 'number' ? m.Urut_Bagi_SPM - n : '',
    Selisih_Muat: typeof m.Urut_Muat === 'number' ? m.Urut_Muat - n : ''
  };
  var d = [o.Selisih_SPM, o.Selisih_Muat].filter(function (x) { return x !== ''; });
  var early = d.some(function (x) { return x < 0; }), late = d.some(function (x) { return x > 0; });
  o.Tanda = !d.length ? '' : (early && late) ? 'BEDA_URUT' : early ? 'LOMPAT' : late ? 'TERTUNDA' : 'SESUAI';
  return o;
}

/** Hitung ulang kolom waktu dan Over_Target dari baris (sudah digabung dengan perubahan). */
function calc_(m) {
  var o = {
    Waiting: mins_(m.Jam_Masuk, m.Jam_Start),
    Loading: mins_(m.Jam_Start, m.Jam_Receive),
    Lead_Time: mins_(m.Jam_Masuk, m.Jam_Receive)
  };
  var tw = Number(cfg_('TARGET_WAITING_MENIT', 30));
  var tl = Number(cfg_('TARGET_LOADING_MENIT', 60));
  var tt = Number(cfg_('TARGET_LEAD_MENIT', 120));
  var over = (o.Waiting !== '' && o.Waiting > tw) ||
             (o.Loading !== '' && o.Loading > tl) ||
             (o.Lead_Time !== '' && o.Lead_Time > tt);
  o.Over_Target = over ? 'Ya' : 'Tidak';
  var u = urut_(m);
  Object.keys(u).forEach(function (k) { o[k] = u[k]; });
  return o;
}

/** Nomor antrean terkecil di antara mobil yang masih menunggu loading. */
function nextFifo_(rows) {
  var waiting = rows.filter(function (r) { return r.Status === CFG.STATUS.MASUK; });
  if (!waiting.length) return '';
  return pad_(Math.min.apply(null, waiting.map(function (r) { return Number(r.No_Antrean); })));
}

/**
 * Satu pintu untuk semua perubahan status.
 * p = { action, id, pelaku, page, alasan? }
 */
function act_(p) {
  var f = CFG.FLOW[p.action];
  if (!f) throw new Error('Aksi tidak dikenal.');
  assertUser_(p.pelaku, f.role);

  var lock = LockService.getScriptLock();
  waitLock_(lock);
  try {
    var all = readTail_(CFG.SHEET.TRX, CFG.TAIL_ROWS);
    var r = all.filter(function (x) { return x.ID === p.id; })[0];
    if (!r) throw new Error('Data tidak ditemukan.');
    if (r.Status !== f.from) throw new Error('Status sudah berubah. Muat ulang halaman.');

    var set = { Status: f.to };
    set[f.time] = new Date();
    if (f.actor) set[f.actor] = p.pelaku;
    var note = '';
    var day = rowsOfDate_(all, dateStr_(r.Tanggal));

    if (p.action === 'start') {
      var first = nextFifo_(rowsOfDate_(all, dateStr_(r.Tanggal)));
      var skipped = first !== '' && Number(r.No_Antrean) > Number(first);
      if (skipped) {
        var why = String(p.alasan || '').trim();
        set.Exception = 'Ya';
        set.Alasan_Exception = why;
        note = 'Exception' + (why ? ': ' + why : ' (tanpa alasan)');
      }
      set.FIFO_OK = skipped ? 'Tidak' : 'Ya';
      set.Urut_Muat = maxNum_(day, 'Urut_Muat') + 1;
      note = (note ? note + '; ' : '') + 'Urut muat ke-' + set.Urut_Muat;
    }

    var merged = {};
    Object.keys(r).forEach(function (k) { merged[k] = r[k]; });
    Object.keys(set).forEach(function (k) { merged[k] = set[k]; });
    var calc = calc_(merged);
    Object.keys(calc).forEach(function (k) { set[k] = calc[k]; });

    updateObj_(CFG.SHEET.TRX, r._row, set, merged);
    log_(r.ID, r.Status, f.to, p.pelaku, f.role, note);

    // perbarui salinan di memori agar tampilan tidak perlu membaca sheet lagi
    Object.keys(set).forEach(function (k) { r[k] = set[k]; });
    return { ok: true, view: view_(p.page, rowsOfDate_(all, today_())) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Tampilan ---------- */

function card_(r) {
  var S = CFG.STATUS;
  var base = { MASUK: r.Jam_Masuk, SPM_DIBAGI: r.Jam_Dibagi, LOADING: r.Jam_Start,
               SELESAI_MUAT: r.Jam_Finish, RECEIVED: r.Jam_Receive }[r.Status];
  var num = function (x) { return typeof x === 'number' ? x : null; };
  return {
    id: r.ID, no: pad_(Number(r.No_Antrean)), nopol: r.No_Polisi,
    status: r.Status, sopir: r.Sopir, tujuan: r.Tujuan,
    masuk: hm_(r.Jam_Masuk), start: hm_(r.Jam_Start),
    receive: hm_(r.Jam_Receive), completed: hm_(r.Jam_Completed),
    security: r.Dicatat_Security, juru: r.Juru_Muat, checker: r.Checker,
    waiting: num(r.Waiting), loading: num(r.Loading), lead: num(r.Lead_Time),
    mnt: base instanceof Date ? Math.round((new Date() - base) / 60000) : null,
    exc: r.Exception === 'Ya', alasan: r.Alasan_Exception,
    urutMuat: num(r.Urut_Muat), selMuat: num(r.Selisih_Muat), tanda: r.Tanda || ''
  };
}

function avg_(rows, key) {
  var v = rows.map(function (r) { return r[key]; })
              .filter(function (x) { return typeof x === 'number'; });
  return v.length ? Math.round(v.reduce(function (a, b) { return a + b; }, 0) / v.length) : null;
}

function kpi_(rows) {
  var S = CFG.STATUS;
  var started = rows.filter(function (r) { return r.FIFO_OK === 'Ya' || r.FIFO_OK === 'Tidak'; });
  var fifoOk = started.filter(function (r) { return r.FIFO_OK === 'Ya'; }).length;
  return {
    total: rows.length,
    waiting: rows.filter(function (r) { return r.Status === S.MASUK || r.Status === S.SPM; }).length,
    loading: rows.filter(function (r) { return r.Status === S.LOADING; }).length,
    completed: rows.filter(function (r) { return r.Status === S.DONE; }).length,
    avgWaiting: avg_(rows, 'Waiting'),
    avgLoading: avg_(rows, 'Loading'),
    fifo: started.length ? Math.round(fifoOk * 100 / started.length) : null,
    over: rows.filter(function (r) { return r.Over_Target === 'Ya'; }).length
  };
}

/** Data tampilan per halaman, selalu untuk tanggal hari ini. */
function view_(page, rowsIn) {
  var date = today_();
  var rows = rowsIn || listByDate_(date);
  var S = CFG.STATUS, R = CFG.ROLE;
  var v = { tanggal: date };

  if (page === 'juru') {
    v.users = userNames_(R.JURU);
    v.list = rows.filter(function (r) { return r.Status === S.MASUK || r.Status === S.LOADING; }).map(card_);
    v.next = nextFifo_(rows);
  } else if (page === 'checker') {
    v.users = userNames_(R.CHECKER);
    v.list = rows.filter(function (r) { return r.Status === S.LOADING || r.Status === S.RECEIVED; }).map(card_);
  } else if (page === 'dashboard') {
    v.kpi = kpi_(rows);
    v.list = rows.map(card_);
  } else {
    throw new Error('Halaman tidak dikenal.');
  }
  return v;
}

/* ======================= Web ======================= */

/**
 * Web.gs — pintu masuk web app dan API untuk halaman.
 * Satu link, satu halaman per peran: ...?page=security | pic | juru | checker | dashboard
 */

var PAGES = {
  security: 'Security',
  juru: 'Juru',
  checker: 'Checker',
  dashboard: 'Dashboard'
};

/* ---- API JSON (dipakai halaman di GitHub Pages) ---- */
var API_VIEW_ = { apiView: 1, apiSecurityInit: 1, apiSecurityList: 1 };   // boleh lewat GET (hanya membaca)
var API_ACT_  = { apiAct: 1, apiSecurityEntry: 1 };                       // hanya lewat POST (mengubah data)

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function apiRun_(fn, arg, table) {
  try {
    if (!table[fn]) throw new Error('Perintah tidak dikenal.');
    var f = { apiView: apiView, apiSecurityInit: apiSecurityInit, apiSecurityList: apiSecurityList,
              apiAct: apiAct, apiSecurityEntry: apiSecurityEntry }[fn];
    return json_({ ok: true, data: f(arg) });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err).replace(/^Error:\s*/, '') });
  }
}

function doPost(e) {
  var b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: 'Data tidak valid.' }); }
  return apiRun_(String(b.fn || ''), b.arg, API_ACT_);
}

function doGet(e) {
  if (e && e.parameter && e.parameter.fn) {
    var arg = null;
    try { arg = JSON.parse(e.parameter.arg || 'null'); } catch (x) {}
    return apiRun_(String(e.parameter.fn), arg, API_VIEW_);
  }
  var key = String((e && e.parameter && e.parameter.page) || 'security').toLowerCase();
  if (!PAGES[key]) return HtmlService.createHtmlOutput('Halaman tidak ditemukan.');
  var t = HtmlService.createTemplateFromFile(PAGES[key]);
  t.page = key;
  return t.evaluate()
    .setTitle('Loading Control Tower')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

var NAV = [
  ['security', 'Security',  'M12 1L3 5v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V5l-9-4z'],
  ['juru',     'Juru muat', 'M20 8h-3V4H3c-1.1 0-2 .9-2 2v11h2c0 1.66 1.34 3 3 3s3-1.34 3-3h6c0 1.66 1.34 3 3 3s3-1.34 3-3h2v-5l-3-4zM6 18.5c-.83 0-1.5-.67-1.5-1.5s.67-1.5 1.5-1.5 1.5.67 1.5 1.5-.67 1.5-1.5 1.5zm13.5-9l1.96 2.5H17V9.5h2.5zm-1.5 9c-.83 0-1.5-.67-1.5-1.5s.67-1.5 1.5-1.5 1.5.67 1.5 1.5-.67 1.5-1.5 1.5z'],
  ['checker',  'Checker',   'M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z'],
  ['dashboard','Dashboard', 'M3 13h8V3H3v10zm0 8h8v-6H3v6zm10 0h8V11h-8v10zm0-18v6h8V3h-8z']
];

/** Menu bawah (HP) atau sidebar (desktop), dibuat di server agar link memakai URL web app. */
function navHtml(page) {
  var base = ScriptApp.getService().getUrl();
  var items = NAV.map(function (n) {
    return '<a href="' + base + '?page=' + n[0] + '" target="_top"' + (n[0] === page ? ' class="on" aria-current="page"' : '') + '>' +
           '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="' + n[2] + '"/></svg><span>' + n[1] + '</span></a>';
  }).join('');
  return '<nav class="nav" aria-label="Menu">' + items + '</nav>';
}

function include(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

/* ---- Security ---- */
function apiSecurityInit() {
  var v = securityView_();
  v.users = userNames_(CFG.ROLE.SECURITY);
  return v;
}
function apiSecurityList() { return securityView_(); }
function apiSecurityEntry(p) { return recordEntry_(p); }

/* ---- PIC, Juru muat, Checker, Dashboard ---- */
function apiView(page) { return view_(page); }
function apiAct(p) { return act_(p); }
