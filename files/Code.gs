/**
 * DASHBOARD STATUS RELOKASI — Backend API (Google Apps Script)
 * ------------------------------------------------------------
 * Frontend: static site di Cloudflare Pages (index.html).
 * Backend : script ini, di-deploy sebagai Web App (JSON API).
 *
 * Endpoint (semua lewat POST ke URL /exec, body JSON, Content-Type text/plain):
 *   {action:'login',  username, password} -> {success, token, kategori, nama, namaList, username}
 *   {action:'data',   token}              -> array data dashboard (sesuai hak akses akun)
 *   {action:'logout', token}              -> {success:true}
 *
 * Cara pasang:
 * 1) Buka Google Sheet sumber data > Extensions > Apps Script.
 * 2) Ganti isi Code.gs dengan file ini. HAPUS file HTML "Dashboard" (sudah pindah ke Cloudflare).
 * 3) Jalankan setupLoginSheet() sekali (kalau sheet "Login Dashboard" belum ada).
 * 4) Isi CONFIG.FRONTEND_URL dengan alamat Cloudflare Pages (opsional, untuk menu di Sheet).
 * 5) Deploy > New deployment > Web app:
 *      Execute as : Me
 *      Who has access : Anyone   (WAJIB "Anyone", bukan "Anyone with Google account",
 *                                 supaya bisa dipanggil dari domain Cloudflare)
 * 6) Salin URL /exec ke variabel API_URL di index.html.
 * 7) Update kode berikutnya: Deploy > Manage deployments > pensil > Version: New version
 *    (URL tetap sama).
 */

var CONFIG = {
  SHEET_NAME: 'Data Status Relokasi', // ganti jika nama tab sheet berbeda
  LOGIN_SHEET_NAME: 'Login Dashboard', // sheet "database" akun login (dibuat via setupLoginSheet())
  FRONTEND_URL: 'https://GANTI-NAMA-PROJECT.pages.dev', // alamat dashboard di Cloudflare Pages
  SESSION_TTL_SEC: 6 * 60 * 60, // masa berlaku sesi login (maks 21600 = 6 jam, batas CacheService)
  MAX_FAILED_LOGIN: 5,          // gagal login berturut-turut sebelum akun dikunci sementara
  LOCK_SEC: 10 * 60,            // lama kunci sementara (detik)
  STEP_COLUMNS: [
    'BA Survey', 'Pembuatan Alur', 'Permintaan Kabel', 'Pengiriman Kabel',
    'Penarikan Kabel', 'Cut Over', 'Opname/QC',
    'BOQ', 'PR-PO', 'BAST', 'Invoice', 'Payment'
  ],
  META_COLUMNS: [
    'No', 'ID Project APJATEL', 'Kategori Program', 'Nama Program',
    'Lokasi Relokasi', 'Nama Instansi', 'Nama Vendor', 'Waktu Relokasi',
    'Wilayah', 'Status Relokasi', 'Nama Operator', '% Progress',
    'Statuss All Progress', 'Status Mitra'
  ],
  // Kolom baru: PIC per tahap (bisa kosong di banyak baris, itu wajar)
  PIC_COLUMNS: [
    'PIC Mitra - Teknis', 'PIC Mitra - Administrasi',
    'PIC Operator - Teknis', 'PIC Operator - Administrasi'
  ],
  // Kolom baru: durasi (hari) per tahap. BA Survey tidak punya kolom durasi di sheet
  // sumber, jadi dihitung terpisah (lihat baSurveyDuration di bawah).
  // Nilai di luar 0-730 hari dianggap tidak valid (data sumbernya kadang error/kosong)
  // dan diabaikan saat dihitung rata-rata.
  DURATION_COLUMN_MAP: {
    'BOQ': 'Durasi BOQ - Mitra',
    'PR-PO': 'Durasi PO - Operator',
    'Pembuatan Alur': 'Durasi Pembuatan Alur -Mitra',
    'Permintaan Kabel': 'Durasi Permintaan Kabel',
    'Pengiriman Kabel': 'Durasi Pengiriman Kabel - Operator',
    'Penarikan Kabel': 'Durasi Penarikan Kabel - Mitra',
    'Cut Over': 'Durasi Cut Over - Operator',
    'Opname/QC': 'Durasi Opname - Operator',
    'BAST': 'Durasi BAST - Mitra',
    'Invoice': 'Durasi Invoice - Mitra',
    'Payment': 'Durasi Payment - Operator'
  },
  DURATION_MAX_VALID_DAYS: 365, // diperketat dari 730 — banyak kolom tanggal sumber belum terisi lengkap
  // Kolom tanggal per tahap (dipakai untuk cari tanggal input TERBARU di seluruh tahapan,
  // ditampilkan sebagai info "Update data terakhir" di dashboard). Boleh diisi array kalau
  // satu tahap punya lebih dari satu kolom tanggal (mis. Penarikan Kabel: Mulai & Selesai).
  TANGGAL_COLUMN_MAP: {
    'BA Survey': 'Tanggal BA Survey',
    'Pembuatan Alur': 'Tanggal Pembuatan Alur',
    'Permintaan Kabel': 'Tanggal Permintaan Kabel',
    'Pengiriman Kabel': 'Tanggal Pengiriman Kabel',
    'Penarikan Kabel': ['Tanggal Mulai Penarikan Kabel', 'Tanggal Selesai Penarikan Kabel'],
    'Cut Over': 'Tanggal Cut Over',
    'Opname/QC': 'Tanggal Opname/QC',
    'BOQ': 'Tanggal BOQ',
    'PR-PO': 'Tanggal PR-PO',
    'BAST': 'Tanggal BAST',
    'Invoice': 'Tanggal Invoice',
    'Payment': 'Tanggal Payment'
  }
};

/** Menambahkan menu custom saat spreadsheet dibuka */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Dashboard Relokasi')
    .addItem('Buka Dashboard', 'showDashboard')
    .addToUi();
}

/**
 * SETUP DATABASE LOGIN — jalankan SEKALI dari editor Apps Script
 * (pilih fungsi ini di dropdown atas, klik Run).
 * Membuat sheet "Login Dashboard" kalau belum ada, dengan kolom:
 *   Username | Password | Kategori | Nama | Aktif
 * - Kategori diisi salah satu: "Vendor", "Operator", atau "Admin"
 * - Nama HARUS SAMA PERSIS dengan nilai di kolom "Nama Vendor" atau
 *   "Nama Operator" pada sheet data (case-sensitive tidak masalah, tapi
 *   ejaan harus sama). Kosongkan untuk Admin.
 *   BISA DIISI LEBIH DARI SATU NAMA, dipisah koma — misal satu login operator
 *   yang menaungi beberapa nama badan usaha: "PT A, PT B, PT C". Login itu akan
 *   melihat gabungan data dari semua nama yang disebut.
 * - Aktif diisi Y untuk aktif, N untuk nonaktifkan akun tanpa menghapus baris.
 * Tambah/ubah/hapus baris akun langsung di sheet ini kapan saja, tidak perlu
 * ubah kode. Password disimpan apa adanya (plain text) di sheet — akses ke
 * sheet ini sendiri tetap dibatasi lewat permission Google Sheet seperti biasa.
 */
function setupLoginSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.LOGIN_SHEET_NAME);
  if (sheet) {
    SpreadsheetApp.getUi().alert('Sheet "' + CONFIG.LOGIN_SHEET_NAME + '" sudah ada, tidak dibuat ulang.');
    return;
  }
  sheet = ss.insertSheet(CONFIG.LOGIN_SHEET_NAME);
  var headers = ['Username', 'Password', 'Kategori', 'Nama', 'Aktif'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.getRange(2, 1, 4, headers.length).setValues([
    ['admin', 'ubah-password-ini', 'Admin', '', 'Y'],
    ['vendor1', 'ubah-password-ini', 'Vendor', 'Nama Vendor Contoh (samakan persis dgn sheet data)', 'Y'],
    ['operator1', 'ubah-password-ini', 'Operator', 'Nama Operator Contoh (samakan persis dgn sheet data)', 'Y'],
    ['operator2', 'ubah-password-ini', 'Operator', 'Nama Operator A, Nama Operator B (contoh multi-operator, pisah koma)', 'Y']
  ]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  SpreadsheetApp.getUi().alert(
    'Sheet "' + CONFIG.LOGIN_SHEET_NAME + '" berhasil dibuat.\n\n' +
    'Silakan edit baris contohnya: ganti password, dan isi kolom "Nama" persis ' +
    'sama dengan "Nama Vendor" / "Nama Operator" yang ada di sheet data. ' +
    'Bisa diisi lebih dari satu nama dipisah koma kalau satu login perlu melihat ' +
    'gabungan beberapa vendor/operator. Tambah baris baru untuk akun lain kapan saja.'
  );
}

/**
 * Verifikasi kredensial (internal, tidak diekspos langsung ke client).
 * Return: { success:true, kategori, nama, namaList, username } atau { success:false, message }
 * namaList = array nama (hasil pecah kolom "Nama" berdasarkan koma), dipakai untuk
 * menyaring data kalau satu login menaungi lebih dari satu vendor/operator.
 */
function checkCredentials_(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.LOGIN_SHEET_NAME);
  if (!sheet) {
    return { success: false, message: 'Sheet "' + CONFIG.LOGIN_SHEET_NAME + '" belum dibuat. Jalankan fungsi setupLoginSheet() dulu dari editor Apps Script.' };
  }
  if (!username || !password) {
    return { success: false, message: 'Username dan password wajib diisi.' };
  }
  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  var uInput = String(username).trim().toLowerCase();
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var uRow = String(row[idx['Username']] || '').trim().toLowerCase();
    if (uRow !== uInput) continue;

    var pRow = String(row[idx['Password']] || '');
    var aktif = String(row[idx['Aktif']] || 'Y').trim().toUpperCase();
    if (aktif === 'N') {
      return { success: false, message: 'Akun ini sudah dinonaktifkan.' };
    }
    if (pRow !== String(password)) {
      return { success: false, message: 'Username atau password salah.' };
    }

    var kategori = String(row[idx['Kategori']] || '').trim();
    var namaRaw = String(row[idx['Nama']] || '').trim();
    var namaList = namaRaw.split(',').map(function (s) { return s.trim(); }).filter(Boolean);
    if (kategori !== 'Admin' && kategori !== 'Vendor' && kategori !== 'Operator') {
      return { success: false, message: 'Kategori akun tidak valid: "' + kategori + '". Harus Admin/Vendor/Operator.' };
    }
    if (kategori !== 'Admin' && namaList.length === 0) {
      return { success: false, message: 'Akun kategori ' + kategori + ' belum diisi kolom "Nama".' };
    }
    return { success: true, kategori: kategori, nama: namaRaw, namaList: namaList, username: row[idx['Username']] };
  }
  return { success: false, message: 'Username atau password salah.' };
}

/** Menu di Sheet: membuka dashboard di Cloudflare Pages (tab baru) */
function showDashboard() {
  var html = HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif">Dashboard sekarang berjalan di:</p>' +
    '<p><a href="' + CONFIG.FRONTEND_URL + '" target="_blank" style="font-family:sans-serif">' +
    CONFIG.FRONTEND_URL + '</a></p>'
  ).setWidth(420).setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard Status Relokasi');
}

/** Respons JSON standar */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Health check: buka URL /exec di browser untuk memastikan deployment hidup */
function doGet() {
  return json_({ ok: true, service: 'Dashboard Status Relokasi API' });
}

/** Router API */
function doPost(e) {
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    switch (req.action) {
      case 'login':
        return json_({ ok: true, data: apiLogin_(req.username, req.password) });
      case 'data':
        return json_({ ok: true, data: getDashboardData(requireSession_(req.token)) });
      case 'logout':
        if (req.token) CacheService.getScriptCache().remove('sess_' + req.token);
        return json_({ ok: true, data: { success: true } });
      default:
        return json_({ ok: false, error: 'Aksi tidak dikenali.' });
    }
  } catch (err) {
    return json_({ ok: false, code: err.code || 'ERROR', error: err.message || String(err) });
  }
}

/** Login: cek kredensial, batasi percobaan gagal, terbitkan token sesi */
function apiLogin_(username, password) {
  var cache = CacheService.getScriptCache();
  var uKey = 'fail_' + String(username || '').trim().toLowerCase();
  var fails = parseInt(cache.get(uKey) || '0', 10);
  if (fails >= CONFIG.MAX_FAILED_LOGIN) {
    return { success: false, message: 'Terlalu banyak percobaan gagal. Coba lagi dalam beberapa menit.' };
  }
  var res = checkCredentials_(username, password);
  if (!res.success) {
    if (String(username || '').trim()) cache.put(uKey, String(fails + 1), CONFIG.LOCK_SEC);
    return res;
  }
  cache.remove(uKey);
  var token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  cache.put('sess_' + token, JSON.stringify({
    kategori: res.kategori, nama: res.nama, namaList: res.namaList, username: res.username
  }), CONFIG.SESSION_TTL_SEC);
  res.token = token;
  return res;
}

/** Ambil identitas dari token sesi di server (client TIDAK bisa menentukan hak aksesnya sendiri) */
function requireSession_(token) {
  var raw = token ? CacheService.getScriptCache().get('sess_' + token) : null;
  if (!raw) {
    var err = new Error('Sesi berakhir, silakan login kembali.');
    err.code = 'AUTH';
    throw err;
  }
  return JSON.parse(raw);
}

/**
 * Dipanggil internal oleh doPost() (action 'data') dengan user hasil requireSession_()
 * Membaca, membersihkan, dan mengembalikan data siap pakai untuk dashboard.
 * @param {Object} user - { kategori:'Admin'|'Vendor'|'Operator', nama:string }, hasil dari login().
 *   Wajib diisi (dashboard tidak mengizinkan akses tanpa login) kecuali kategori Admin.
 */
function getDashboardData(user) {
  if (!user || !user.kategori) {
    throw new Error('Akses ditolak. Silakan login terlebih dahulu.');
  }
  if (user.kategori !== 'Admin' && user.kategori !== 'Vendor' && user.kategori !== 'Operator') {
    throw new Error('Kategori login tidak dikenali.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    // fallback: cari nama sheet yang cocok tanpa peduli spasi/huruf besar-kecil
    var target = CONFIG.SHEET_NAME.trim().toLowerCase();
    sheet = ss.getSheets().filter(function (s) {
      return s.getName().trim().toLowerCase() === target;
    })[0];
  }

  if (!sheet) {
    var available = ss.getSheets().map(function (s) { return s.getName(); }).join(', ');
    throw new Error('Sheet "' + CONFIG.SHEET_NAME + '" tidak ditemukan. Sheet yang ada: ' + available);
  }

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    throw new Error('Sheet "' + sheet.getName() + '" ditemukan tapi tidak berisi data (hanya ' + values.length + ' baris).');
  }

  // Cari baris header secara otomatis (bisa jadi bukan baris 1 kalau ada judul/merge di atasnya)
  var headerRowIdx = -1;
  for (var i = 0; i < Math.min(15, values.length); i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (String(values[i][j]).trim() === 'ID Project APJATEL') { headerRowIdx = i; break; }
    }
    if (headerRowIdx !== -1) break;
  }

  if (headerRowIdx === -1) {
    var preview = values.slice(0, 5).map(function (r) {
      return r.slice(0, 6).map(function (c) { return String(c).trim(); }).join(' | ');
    }).join('  ///  ');
    throw new Error('Baris header dengan kolom "ID Project APJATEL" tidak ditemukan di 15 baris pertama sheet "' +
      sheet.getName() + '". Isi 5 baris pertama: ' + preview);
  }

  var headers = values[headerRowIdx];
  var dataRows = values.slice(headerRowIdx + 1);
  var colIndex = {};
  headers.forEach(function (h, i) {
    colIndex[String(h).trim()] = i;
  });

  if (colIndex['ID Project APJATEL'] === undefined) {
    throw new Error('Kolom "ID Project APJATEL" tidak ditemukan di sheet "' + sheet.getName() +
      '". Header yang terbaca: ' + headers.map(function(h){return String(h).trim();}).join(' | '));
  }

  function get(row, name) {
    var idx = colIndex[name];
    if (idx === undefined) return '';
    return row[idx];
  }

  var idPattern = /^(APJ|APJTL)/i;

  function normStep(v) {
    if (v === '' || v === null || v === undefined) return 'Belum';
    var s = String(v).trim().toLowerCase();
    if (s === 'done') return 'Done';
    if (s === 'process' || s === 'proses') return 'Proses';
    return 'Belum'; // N/Y, kosong, atau nilai lain dianggap belum jalan
  }

  function toYear(v) {
    if (v === '' || v === null || v === undefined) return null;
    if (v instanceof Date) return v.getFullYear();
    var n = parseInt(v, 10);
    return isNaN(n) ? null : n;
  }

  function toDate(v) {
    if (v instanceof Date) return v;
    return null;
  }

  var MS_PER_DAY = 24 * 60 * 60 * 1000;

  // BA Survey tidak punya kolom "Durasi" tersendiri di sheet sumber.
  // Dihitung sendiri per RUAS LOKASI RELOKASI: selisih hari antara tanggal Pembuatan Alur
  // TERAWAL dan tanggal BA Survey TERAWAL, di antara semua baris (operator) untuk lokasi
  // yang sama — merepresentasikan berapa lama fase survei berlangsung sampai fase
  // Pembuatan Alur mulai dikerjakan.
  var baSurveyRange = {}; // lokasi -> { minSurvey: Date, minAlur: Date }
  for (var pr = 0; pr < dataRows.length; pr++) {
    var prow = dataRows[pr];
    var lok = get(prow, 'Lokasi Relokasi');
    if (lok === '' || lok === null || lok === undefined) continue;
    lok = String(lok).trim();
    var dSurvey = toDate(get(prow, 'Tanggal BA Survey'));
    var dAlur = toDate(get(prow, 'Tanggal Pembuatan Alur'));
    if (!baSurveyRange[lok]) baSurveyRange[lok] = { minSurvey: null, minAlur: null };
    if (dSurvey && (!baSurveyRange[lok].minSurvey || dSurvey < baSurveyRange[lok].minSurvey)) {
      baSurveyRange[lok].minSurvey = dSurvey;
    }
    if (dAlur && (!baSurveyRange[lok].minAlur || dAlur < baSurveyRange[lok].minAlur)) {
      baSurveyRange[lok].minAlur = dAlur;
    }
  }
  function baSurveyDuration(lok) {
    var rng = baSurveyRange[lok];
    if (!rng || !rng.minSurvey || !rng.minAlur) return null;
    var days = Math.round((rng.minAlur - rng.minSurvey) / MS_PER_DAY);
    return (days >= 0 && days <= CONFIG.DURATION_MAX_VALID_DAYS) ? days : null;
  }

  var out = [];
  for (var r = 0; r < dataRows.length; r++) {
    var row = dataRows[r];
    var idProject = get(row, 'ID Project APJATEL');
    if (idProject === '' || idProject === null || idProject === undefined) continue;

    var kategori = get(row, 'Kategori Program');
    idProject = String(idProject).trim();
    kategori = String(kategori).trim();

    // Perbaiki baris yang kolom ID Project & Kategori Program-nya tertukar
    // (masalah data yang ditemukan di sheet sumber)
    if (!idPattern.test(idProject) && idPattern.test(kategori)) {
      var tmp = idProject;
      idProject = kategori;
      kategori = tmp;
    }
    if (kategori === 'Pembanguan MRT') kategori = 'Pembangunan MRT';

    var statusRelokasi = String(get(row, 'Status Relokasi') || 'Unknown').trim();
    if (statusRelokasi === 'On Progres') statusRelokasi = 'On Progress';
    if (statusRelokasi === '') statusRelokasi = 'Unknown';

    var record = {
      'No': get(row, 'No'),
      'ID Project APJATEL': idProject,
      'Kategori Program': kategori || 'Unknown',
      'Nama Program': get(row, 'Nama Program') || '',
      'Lokasi Relokasi': get(row, 'Lokasi Relokasi') || '',
      'Nama Instansi': get(row, 'Nama Instansi') || '',
      'Nama Vendor': String(get(row, 'Nama Vendor') || 'Belum Ada Vendor').trim() || 'Belum Ada Vendor',
      'Waktu Relokasi': toYear(get(row, 'Waktu Relokasi')),
      'Wilayah': String(get(row, 'Wilayah') || 'Unknown').trim() || 'Unknown',
      'Status Relokasi': statusRelokasi,
      'Nama Operator': String(get(row, 'Nama Operator') || 'Belum Ada Operator').trim() || 'Belum Ada Operator',
      'Status Mitra': String(get(row, 'Status Mitra') || 'Unknown').trim() || 'Unknown'
    };

    // % Progress & Statuss All Progress TIDAK diambil dari kolom sheet (seringkali
    // tidak sinkron / ketinggalan update dibanding status 12 tahapan yang sebenarnya).
    // Dihitung sendiri dari status 12 tahapan supaya selalu konsisten dengan section
    // "Progres 12 Tahap Relokasi" di dashboard.
    var doneCount = 0;
    CONFIG.STEP_COLUMNS.forEach(function (step) {
      var v = normStep(get(row, step));
      record[step] = v;
      if (v === 'Done') doneCount++;
    });
    record['% Progress'] = Math.round((doneCount / CONFIG.STEP_COLUMNS.length) * 100);
    record['Statuss All Progress'] = (doneCount === CONFIG.STEP_COLUMNS.length) ? 'Complete' : 'On Going';

    CONFIG.PIC_COLUMNS.forEach(function (col) {
      var v = get(row, col);
      record[col] = (v === null || v === undefined || v === '') ? '' : String(v).trim();
    });

    Object.keys(CONFIG.DURATION_COLUMN_MAP).forEach(function (step) {
      var col = CONFIG.DURATION_COLUMN_MAP[step];
      var v = get(row, col);
      var n = Number(v);
      var valid = (v !== '' && v !== null && !isNaN(n) && n >= 0 && n <= CONFIG.DURATION_MAX_VALID_DAYS);
      record['Durasi_' + step] = valid ? n : null;
    });
    record['Durasi_BA Survey'] = baSurveyDuration(String(get(row, 'Lokasi Relokasi') || '').trim());

    // Tanggal input terbaru di antara SEMUA kolom tanggal tahapan untuk baris ini
    // (dipakai client untuk hitung "Update data terakhir" dari seluruh/filtered data).
    // Tanggal di masa depan (typo tahun di sheet, mis. 2029) diabaikan supaya tidak
    // menampilkan info yang tidak masuk akal.
    var now = new Date();
    var latestDate = null;
    Object.keys(CONFIG.TANGGAL_COLUMN_MAP).forEach(function (step) {
      var cols = CONFIG.TANGGAL_COLUMN_MAP[step];
      if (!Array.isArray(cols)) cols = [cols];
      cols.forEach(function (col) {
        var d = toDate(get(row, col));
        if (d && d <= now && (!latestDate || d > latestDate)) latestDate = d;
      });
    });
    record['_TanggalTerbaru'] = latestDate ? latestDate.getTime() : null;

    // Tanggal terbaru PER TAHAP (timestamp ms, atau null) — dipakai popup detail
    // saat bar "Progres 12 tahap relokasi" diklik. Tanggal masa depan diabaikan.
    Object.keys(CONFIG.TANGGAL_COLUMN_MAP).forEach(function (step) {
      var cols = CONFIG.TANGGAL_COLUMN_MAP[step];
      if (!Array.isArray(cols)) cols = [cols];
      var best = null;
      cols.forEach(function (col) {
        var d = toDate(get(row, col));
        if (d && d <= now && (!best || d > best)) best = d;
      });
      record['Tgl_' + step] = best ? best.getTime() : null;
    });

    out.push(record);
  }

  // Cuma tampilkan data dengan Status Mitra "Aktif" (mitra yang statusnya Non Aktif
  // atau kosong disembunyikan dari dashboard, berlaku untuk semua kategori login).
  out = out.filter(function (rec) {
    return rec['Status Mitra'] === 'Aktif';
  });

  // Batasi data sesuai kategori & nama login (kecuali Admin, yang melihat semua data)
  if (user.kategori === 'Vendor') {
    var vendorSet = (user.namaList || [user.nama]).map(function (n) { return n.trim().toLowerCase(); });
    out = out.filter(function (rec) {
      return vendorSet.indexOf(rec['Nama Vendor'].trim().toLowerCase()) !== -1;
    });
  } else if (user.kategori === 'Operator') {
    var operatorSet = (user.namaList || [user.nama]).map(function (n) { return n.trim().toLowerCase(); });
    out = out.filter(function (rec) {
      return operatorSet.indexOf(rec['Nama Operator'].trim().toLowerCase()) !== -1;
    });
  }

  return out;
}