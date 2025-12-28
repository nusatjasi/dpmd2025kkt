const SPREADSHEET_ID = '1kq1ZyFlwop2IWshiothH0IFwdp6Tl_wv-MHH541Of-g';
const UPLOAD_FOLDER_ID = '1hXA7MuOg2TbhfRtVfZeuIZj07Yd_883g';

function doGet(e) {
  const page = (e && e.parameter && e.parameter.p) || 'index';
  initializeSystem();

  try {
    return HtmlService.createTemplateFromFile(page)
      .evaluate()
      .setTitle('DPMD Digital - Kab. Kepulauan Tanimbar')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return HtmlService.createHtmlOutput('<h3>Halaman tidak ditemukan: ' + page + '</h3>');
  }
}

/**
 * Helper untuk menyisipkan file HTML lain (seperti CSS atau JS)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Mendapatkan URL dasar aplikasi untuk navigasi
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Membuat sheet yang diperlukan secara otomatis
 * dan menambahkan user Super Admin jika kosong.
 */
function initializeSystem() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetsToInit = {
    'Users': ['id', 'fullname', 'username', 'password', 'role'],
    'Settings': ['Key', 'Value'],
    'DataPegawai': [
      'id', 'nama', 'tanggal_lahir', 'nip', 'jenis_kelamin', 'agama', 'status_pegawai',
      'pangkat', 'gol_ruang', 'tmt_pangkat', 'jabatan', 'eselon', 'tmt_jabatan',
      'masa_kerja_tahun', 'masa_kerja_bulan', 'pendidikan_terakhir', 'no_karpeg',
      'tmt_cpns', 'tmt_pns', 'naik_pangkat_berikutnya', 'naik_gaji_berikutnya', 'deviceId'
    ],
    'SuratMasuk': ['id', 'noAgenda', 'tglTerima', 'sifatSurat', 'noSurat', 'tglSurat', 'pengirim', 'perihal', 'linkSurat', 'tujuanSurat', 'nomorWa', 'disposisi'],
    'DataAbsensi': ['Timestamp', 'Tanggal', 'Waktu', 'Nama', 'NIP', 'Sesi', 'Status', 'Keterangan', 'Latitude', 'Longitude', 'DeviceId']
  };

  for (let name in sheetsToInit) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(sheetsToInit[name]);
      sheet.getRange(1, 1, 1, sheetsToInit[name].length).setFontWeight('bold').setBackground('#f3f3f3');
    }

    // Jika sheet Users baru saja dibuat atau kosong (hanya header), tambahkan admin default
    if (name === 'Users' && sheet.getLastRow() <= 1) {
      // Cek apakah kolom pertama (header) sudah ada, jika belum tulis header dulu
      if (sheet.getLastRow() === 0) sheet.appendRow(sheetsToInit[name]);

      sheet.appendRow([
        Utilities.getUuid(),
        'Administrator',
        'admin',
        'admin123',
        'Super Admin'
      ]);
    }
  }
}

/* --- AUTHENTICATION --- */
function login(username, password) {
  const users = getSheetData('users');
  const user = users.find(u => u.username === username && u.password === password);

  if (user) {
    return { success: true, user: user };
  }
  return { success: false, message: 'Username atau Password salah!' };
}

/* --- SETTINGS --- */


function saveAppSetting(key, value) {
  const sheet = getSheet('Settings');
  const data = sheet.getDataRange().getValues();
  let found = false;

  if (data.length === 0) sheet.appendRow(['Key', 'Value']);

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      found = true;
      break;
    }
  }

  if (!found) {
    sheet.appendRow([key, value]);
  }

  return { success: true, message: 'Pengaturan disimpan' };
}

/* --- GENERIC DATA HANDLERS --- */
function getSheetData(sheetName) {
  const realSheetName = getRealSheetName(sheetName);
  return getDataFromSheet(realSheetName);
}

function saveData(sheetName, payload) {
  const realSheetName = getRealSheetName(sheetName);
  const sheet = getSheet(realSheetName);
  const data = sheet.getDataRange().getValues();

  let headers = [];
  if (data.length > 0) {
    headers = data[0];
  } else {
    headers = Object.keys(payload).filter(k => k !== 'id');
    headers.unshift('id');
    sheet.appendRow(headers);
  }

  if (!payload.id) payload.id = Utilities.getUuid();

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(payload.id)) {
      rowIndex = i + 1;
      break;
    }
  }

  const rowData = headers.map(h => payload[h] || '');

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  return { success: true, message: 'Data berhasil disimpan!', id: payload.id };
}

function deleteData(sheetName, id) {
  const realSheetName = getRealSheetName(sheetName);
  const sheet = getSheet(realSheetName);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Data berhasil dihapus' };
    }
  }
  return { success: false, message: 'ID tidak ditemukan' };
}

/* --- FILE UPLOAD --- */
function uploadPhoto(data, name) {
  try {
    const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
    const blob = Utilities.newBlob(Utilities.base64Decode(data.split(',')[1]), MimeType.JPEG, name);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      success: true,
      url: file.getDownloadUrl(),
      id: file.getId()
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/* --- HELPER --- */
function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    // Jika perlu header spesifik bisa ditambahkan di sini atau lewat initializeSystem
  }
  return sheet;
}

// --- Dashboard ---
function getDashboardStats() {
  const pegawai = getDataFromSheet('DataPegawai');

  const stats = {
    totalPegawai: pegawai.length,
    pns: pegawai.filter(p => p.status_pegawai === 'PNS').length,
    pppk: pegawai.filter(p => p.status_pegawai === 'PPPK').length,
    eselon: {},
    golongan: {},
    agama: {},
    birthdays: []
  };

  pegawai.forEach(p => {
    // Agama
    if (p.agama) stats.agama[p.agama] = (stats.agama[p.agama] || 0) + 1;

    // Eselon
    let e = p.eselon || 'Non Eselon';
    stats.eselon[e] = (stats.eselon[e] || 0) + 1;

    // Golongan (Simplify for Chart)
    if (p.gol_ruang) {
      let g = p.gol_ruang.split('/')[0];
      stats.golongan['Gol ' + g] = (stats.golongan['Gol ' + g] || 0) + 1;
    }

    // Birthdays this month
    if (p.tanggal_lahir) {
      const bday = new Date(p.tanggal_lahir);
      const now = new Date();
      if (bday.getMonth() === now.getMonth()) {
        stats.birthdays.push({ nama: p.nama, tgl: bday.getDate() });
      }
    }
  });

  return jsonResponse({ status: 'success', data: stats });
}

function getDataFromSheet(name) {
  const sheet = getSheet(name);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function getRealSheetName(key) {
  const map = {
    'bezeting': 'DataPegawai',
    'suratmasuk': 'SuratMasuk',
    'master': 'MasterData',
    'users': 'Users',
    'profil': 'Users',
    'dashboard': 'DataPegawai'
  };
  return map[key] || key;
}

/* --- MASTER DATA --- */
function getMasterData() {
  const sheet = getSheet('MasterData');
  const data = sheet.getDataRange().getValues();
  let result = {};

  if (data.length > 1) {
    const headers = data[0];
    headers.forEach(h => result[h] = []);
    for (let i = 1; i < data.length; i++) {
      headers.forEach((key, colIndex) => {
        const val = data[i][colIndex];
        if (val) result[key].push(val);
      });
    }
  } else {
    result = {
      agama: ["Katolik", "Protestan", "Islam", "Hindu", "Buddha", "Konghucu"],
      pangkat: ["Pembina Utama", "Pembina Utama Madya", "Pembina Utama Muda", "Pembina Tingkat I", "Pembina", "Penata Tingkat I", "Penata", "Penata Muda Tingkat I", "Penata Muda", "N/A (untuk PPPK)"],
      golongan: ["IV/e", "IV/d", "IV/c", "IV/b", "IV/a", "III/d", "III/c", "III/b", "III/a", "II/d", "II/c", "II/b", "II/a", "I/d", "I/c", "I/b", "I/a", "N/A"],
      eselon: ["II", "III", "IV", "Non Eselon"],
      jabatan: ["Kepala Dinas", "Sekretaris", "Kepala Bidang"],
      pendidikan: ["SD", "SMP", "SMA/SMK", "D1", "D2", "D3", "D4", "S1", "S2", "S3"]
    };
    saveMasterData(result);
  }
  return { status: 'success', data: result };
}

function saveMasterData(payload) {
  const sheet = getSheet('MasterData');
  sheet.clear();
  const keys = Object.keys(payload);
  if (keys.length === 0) return { status: 'success' };

  sheet.appendRow(keys);
  let maxLen = 0;
  keys.forEach(k => {
    if (payload[k].length > maxLen) maxLen = payload[k].length;
  });

  const rows = [];
  for (let i = 0; i < maxLen; i++) {
    const row = keys.map(k => payload[k][i] || '');
    rows.push(row);
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, keys.length).setValues(rows);
  }
  return { status: 'success', message: 'Master data saved' };
}

/* --- ATTENDANCE --- */
function saveAbsensi(payload) {
  try {
    const sheet = getSheet('DataAbsensi');
    const timestamp = new Date();
    const newRow = [
      timestamp.toISOString(),
      payload.tanggal || timestamp.toISOString().split('T')[0],
      payload.waktu || Utilities.formatDate(timestamp, "GMT+7", "HH:mm:ss"),
      payload.nama,
      payload.nip,
      payload.session || 'Pagi',
      payload.status || 'Hadir',
      payload.keterangan || '-',
      payload.latitude || 0,
      payload.longitude || 0,
      payload.deviceId || '-'
    ];
    sheet.appendRow(newRow);

    // Register deviceId if not already registered
    if (payload.pegawaiId) {
      const pSheet = getSheet('DataPegawai');
      const pData = pSheet.getDataRange().getValues();
      const pHeaders = pData[0];
      const deviceIdIdx = pHeaders.indexOf('deviceId');

      if (deviceIdIdx !== -1) {
        for (let i = 1; i < pData.length; i++) {
          if (String(pData[i][0]) === String(payload.pegawaiId)) {
            if (!pData[i][deviceIdIdx]) {
              pSheet.getRange(i + 1, deviceIdIdx + 1).setValue(payload.deviceId);
            }
            break;
          }
        }
      }
    }

    return { status: 'success', message: 'Absensi berhasil' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function getAbsensi(date) {
  try {
    const data = getDataFromSheet('DataAbsensi');
    const results = data.filter(row => row.Tanggal === date);
    return { status: 'success', data: results };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function updateAbsensiStatus(payload) {
  try {
    const sheet = getSheet('DataAbsensi');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const dateIdx = headers.indexOf('Tanggal');
    const nipIdx = headers.indexOf('NIP');
    const statusIdx = headers.indexOf('Status');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][dateIdx] === payload.date && String(data[i][nipIdx]) === String(payload.nip)) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, statusIdx + 1).setValue(payload.status);
    } else {
      // If not found, create a manual entry
      const timestamp = new Date();
      const newRow = headers.map(h => {
        if (h === 'Timestamp') return timestamp.toISOString();
        if (h === 'Tanggal') return payload.date;
        if (h === 'Waktu') return '-';
        if (h === 'Nama') return payload.nama;
        if (h === 'NIP') return payload.nip;
        if (h === 'Status') return payload.status;
        if (h === 'Sesi') return 'Manual';
        return '-';
      });
      sheet.appendRow(newRow);
    }
    return { status: 'success', message: 'Status diperbarui' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/* --- APP SETTINGS --- */
function saveAppSettings(payload) {
  try {
    for (let key in payload) {
      saveAppSetting(key, typeof payload[key] === 'object' ? JSON.stringify(payload[key]) : payload[key]);
    }
    return { status: 'success', message: 'Settings saved' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function getAppSettings() {
  try {
    const data = getDataFromSheet('Settings');
    const settings = {};
    data.forEach(row => {
      let val = row.Value;
      if (typeof val === 'string' && (val.startsWith('{') || val.startsWith('['))) {
        try { val = JSON.parse(val); } catch (e) { }
      }
      settings[row.Key] = val;
    });
    return { status: 'success', data: settings };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

