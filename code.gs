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
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Absensi');
  const now = new Date();
  ss.appendRow([
    now, 
    payload.id, 
    payload.nama, 
    payload.kategori, 
    payload.tipe, 
    'Hadir'
  ]);
  return "Absen " + payload.tipe + " berhasil dikirim!";
}

// Ambil rekap absensi berdasarkan bulan (format: "YYYY-MM") - difilter di server
function getRekapBulan(bulan) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Absensi');
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0].map(h => h.toString().toLowerCase());
  const result = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const ts = row[0];
    if (!ts) continue;
    const tsDate = ts instanceof Date ? ts : new Date(ts);
    const rowBulan = Utilities.formatDate(tsDate, Session.getScriptTimeZone(), 'yyyy-MM');
    if (rowBulan === bulan) {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = row[idx] instanceof Date
        ? Utilities.formatDate(row[idx], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
        : row[idx]);
      result.push(obj);
    }
  }
  return result;
}

function getStats() {
  const absensi = getData('Absensi');
  const today = new Date().toLocaleDateString();
  const todayAbsen = absensi.filter(a => new Date(a.timestamp).toLocaleDateString() === today);
  
  return {
    total_guru: getData('Guru').length,
    total_staff: getData('Staff').length,
    absen_hari_ini: todayAbsen.length,
    datang: todayAbsen.filter(a => a.tipe === 'Datang').length,
    pulang: todayAbsen.filter(a => a.tipe === 'Pulang').length
  };
}