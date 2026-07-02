/**
 * SISTEM ABSENSI SEKOLAH - GOOGLE APPS SCRIPT
 * Fitur: Dashboard, Manajemen Guru/Staff, Absensi, Rekap, & Pengaturan
 */

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sistem Absensi Sekolah')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Inisialisasi Spreadsheet dan Sheet Otomatis (hanya buat jika belum ada)
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    'Pengaturan': [['Key', 'Value'], ['nama_sekolah', 'Sekolah Contoh'], ['kepala_sekolah', '-'], ['nip_kepsek', '-'], ['logo_url', '']],
    'Guru':       [['NIP', 'Nama', 'Jabatan']],
    'Staff':      [['NIP', 'Nama', 'Jabatan']],
    'Absensi':    [['Timestamp', 'ID', 'Nama', 'Kategori', 'Tipe', 'Status']]
  };

  for (let name in sheets) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, sheets[name].length, sheets[name][0].length).setValues(sheets[name]);
      sheet.getRange(1, 1, 1, sheets[name][0].length).setFontWeight('bold').setBackground('#f3f3f3');
    }
  }
  return "Database berhasil disiapkan!";
}

// Reset penuh: hapus semua data di setiap sheet, tulis ulang header & default settings
function resetDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Definisi tiap sheet: [header_row, ...default_data_rows]
  const definitions = {
    'Pengaturan': [
      ['Key', 'Value'],
      ['nama_sekolah', 'Sekolah Contoh'],
      ['kepala_sekolah', '-'],
      ['nip_kepsek', '-'],
      ['logo_url', '']
    ],
    'Guru':    [['NIP', 'Nama', 'Jabatan']],
    'Staff':   [['NIP', 'Nama', 'Jabatan']],
    'Absensi': [['Timestamp', 'ID', 'Nama', 'Kategori', 'Tipe', 'Status']]
  };

  for (let name in definitions) {
    let sheet = ss.getSheetByName(name);

    if (!sheet) {
      // Buat baru kalau belum ada
      sheet = ss.insertSheet(name);
    } else {
      // Hapus seluruh isi sheet
      sheet.clearContents();
    }

    const data = definitions[name];
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    sheet.getRange(1, 1, 1, data[0].length).setFontWeight('bold').setBackground('#f3f3f3');
  }

  return "✅ Database berhasil direset! Semua data Guru, Staff, dan Absensi telah dihapus.";
}

// Fungsi CRUD & Data Handling
function getData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h.toLowerCase()] = row[i]);
    return obj;
  });
}

function saveData(sheetName, data) {
  // Validasi NIP unik di seluruh database (Guru + Staff)
  const nipBaru = String(data[0]).trim();
  const duplikat = _isNipExists(nipBaru, null);
  if (duplikat) {
    return { ok: false, pesan: 'NIP ' + nipBaru + ' sudah terdaftar sebagai ' + duplikat + '. NIP harus unik!' };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { ok: false, pesan: 'Sheet tidak ditemukan!' };
  sheet.appendRow(data);
  return { ok: true, pesan: 'Data berhasil disimpan!' };
}

function deleteData(sheetName, nip) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return "Sheet tidak ditemukan!";
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(nip)) {
      sheet.deleteRow(i + 1);
      return "Data berhasil dihapus!";
    }
  }
  return "Data tidak ditemukan!";
}

function updateData(sheetName, nipLama, dataBaruArr) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { ok: false, pesan: 'Sheet tidak ditemukan!' };

  // Validasi NIP baru unik (kecualikan NIP lama milik sendiri)
  const nipBaru = String(dataBaruArr[0]).trim();
  if (nipBaru !== String(nipLama).trim()) {
    const duplikat = _isNipExists(nipBaru, String(nipLama).trim());
    if (duplikat) {
      return { ok: false, pesan: 'NIP ' + nipBaru + ' sudah digunakan oleh ' + duplikat + '. NIP harus unik!' };
    }
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(nipLama)) {
      sheet.getRange(i + 1, 1, 1, dataBaruArr.length).setValues([dataBaruArr]);
      return { ok: true, pesan: 'Data berhasil diperbarui!' };
    }
  }
  return { ok: false, pesan: 'Data tidak ditemukan!' };
}

// Helper: cek apakah NIP sudah ada di Guru atau Staff
// excludeNip: NIP yang dikecualikan (untuk kasus edit - NIP lama sendiri)
// Return: nama sheet jika duplikat, null jika tidak
function _isNipExists(nip, excludeNip) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ['Guru', 'Staff'];
  for (let s = 0; s < sheetNames.length; s++) {
    const sheet = ss.getSheetByName(sheetNames[s]);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      const existingNip = String(values[i][0]).trim();
      if (existingNip === nip && existingNip !== excludeNip) {
        return sheetNames[s]; // kembalikan nama sheet tempat NIP ditemukan
      }
    }
  }
  return null;
}

function updateSettings(settingsObj) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pengaturan');
  sheet.clear();
  sheet.appendRow(['Key', 'Value']);
  for (let key in settingsObj) {
    sheet.appendRow([key, settingsObj[key]]);
  }
  return true;
}

function getSettings() {
  const data = getData('Pengaturan');
  let settings = {};
  data.forEach(item => settings[item.key] = item.value);
  return settings;
}

function submitAbsen(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const absenSheet = ss.getSheetByName('Absensi');
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // Cek status absen hari ini untuk NIP ini
  const status = _getStatusHariIni(absenSheet, payload.id, today, tz);

  if (payload.tipe === 'Datang') {
    if (status.sudahDatang) {
      return { ok: false, pesan: "⚠️ Anda sudah tercatat DATANG hari ini!" };
    }
  } else if (payload.tipe === 'Pulang') {
    if (!status.sudahDatang) {
      return { ok: false, pesan: "⚠️ Anda belum absen DATANG hari ini!" };
    }
    if (status.sudahPulang) {
      return { ok: false, pesan: "⚠️ Anda sudah tercatat PULANG hari ini!" };
    }
  }

  absenSheet.appendRow([
    new Date(),
    payload.id,
    payload.nama,
    payload.kategori,
    payload.tipe,
    'Hadir'
  ]);
  return { ok: true, pesan: "✅ Absen " + payload.tipe + " berhasil dicatat!" };
}

// Cek status absen hari ini untuk satu NIP (dipanggil dari frontend saat pilih nama)
function getStatusAbsenHariIni(nip) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const absenSheet = ss.getSheetByName('Absensi');
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  return _getStatusHariIni(absenSheet, nip, today, tz);
}

// Helper internal: kembalikan { sudahDatang, sudahPulang }
function _getStatusHariIni(sheet, nip, today, tz) {
  let sudahDatang = false, sudahPulang = false;
  if (!sheet || sheet.getLastRow() <= 1) return { sudahDatang, sudahPulang };
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  values.forEach(row => {
    const ts = row[0];
    if (!ts) return;
    const tsDate = ts instanceof Date ? ts : new Date(ts);
    const rowDate = Utilities.formatDate(tsDate, tz, 'yyyy-MM-dd');
    if (rowDate !== today) return;
    if (String(row[1]) !== String(nip)) return;
    const tipe = String(row[4]).trim();
    if (tipe === 'Datang') sudahDatang = true;
    if (tipe === 'Pulang') sudahPulang = true;
  });
  return { sudahDatang, sudahPulang };
}

// Ambil rekap absensi berdasarkan bulan (format: "YYYY-MM") - difilter di server
function getRekapBulan(bulan) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Absensi');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  // Baca langsung berdasarkan posisi kolom (bukan nama header)
  // Urutan kolom: [0]Timestamp [1]ID/NIP [2]Nama [3]Kategori [4]Tipe [5]Status
  const tz = Session.getScriptTimeZone();
  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const result = [];

  values.forEach(row => {
    const ts = row[0];
    if (!ts) return;
    const tsDate = ts instanceof Date ? ts : new Date(ts);
    const rowBulan = Utilities.formatDate(tsDate, tz, 'yyyy-MM');
    if (rowBulan === bulan) {
      result.push({
        timestamp: Utilities.formatDate(tsDate, tz, 'dd/MM/yyyy HH:mm'),
        nip:       String(row[1]),   // Kolom B: ID/NIP
        nama:      String(row[2]),   // Kolom C: Nama
        kategori:  String(row[3]),   // Kolom D: Kategori (Guru/Staff)
        tipe:      String(row[4]),   // Kolom E: Tipe (Datang/Pulang)
        status:    String(row[5])    // Kolom F: Status
      });
    }
  });

  return result;
}

// Ambil rekap absensi berdasarkan tanggal spesifik (format: "YYYY-MM-DD") - difilter di server
function getRekapTanggal(tanggal) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Absensi');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const tz = Session.getScriptTimeZone();
  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const result = [];

  values.forEach(row => {
    const ts = row[0];
    if (!ts) return;
    const tsDate = ts instanceof Date ? ts : new Date(ts);
    const rowTanggal = Utilities.formatDate(tsDate, tz, 'yyyy-MM-dd');
    if (rowTanggal === tanggal) {
      result.push({
        timestamp: Utilities.formatDate(tsDate, tz, 'dd/MM/yyyy HH:mm'),
        nip:       String(row[1]),
        nama:      String(row[2]),
        kategori:  String(row[3]),
        tipe:      String(row[4]),
        status:    String(row[5])
      });
    }
  });

  return result;
}

function getStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // Baca sheet Absensi secara langsung (nilai raw) untuk perbandingan tanggal yang akurat
  const absenSheet = ss.getSheetByName('Absensi');
  let datang = 0, pulang = 0, totalHariIni = 0;

  if (absenSheet && absenSheet.getLastRow() > 1) {
    const values = absenSheet.getRange(2, 1, absenSheet.getLastRow() - 1, 6).getValues();
    values.forEach(row => {
      const ts = row[0];
      if (!ts) return;
      const tsDate = ts instanceof Date ? ts : new Date(ts);
      const rowDate = Utilities.formatDate(tsDate, tz, 'yyyy-MM-dd');
      if (rowDate === today) {
        totalHariIni++;
        const tipe = String(row[4]).trim();
        if (tipe === 'Datang') datang++;
        else if (tipe === 'Pulang') pulang++;
      }
    });
  }

  return {
    total_guru: getData('Guru').length,
    total_staff: getData('Staff').length,
    absen_hari_ini: totalHariIni,
    datang: datang,
    pulang: pulang
  };
}