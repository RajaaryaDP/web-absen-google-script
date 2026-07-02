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

// Inisialisasi Spreadsheet dan Sheet Otomatis
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    'Pengaturan': [['Key', 'Value'], ['nama_sekolah', 'Sekolah Contoh'], ['kepala_sekolah', '-'], ['nip_kepsek', '-'], ['logo_url', '']],
    'Guru': [['NIP', 'Nama', 'Jabatan']],
    'Staff': [['NIP', 'Nama', 'Jabatan']],
    'Absensi': [['Timestamp', 'ID', 'Nama', 'Kategori', 'Tipe', 'Status']] // Kategori: Guru/Staff, Tipe: Datang/Pulang
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  sheet.appendRow(data);
  return true;
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
  if (!sheet) return "Sheet tidak ditemukan!";
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(nipLama)) {
      sheet.getRange(i + 1, 1, 1, dataBaruArr.length).setValues([dataBaruArr]);
      return "Data berhasil diperbarui!";
    }
  }
  return "Data tidak ditemukan!";
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