const SPREADSHEET_ID = '1kq1ZyFlwop2IWshiothH0IFwdp6Tl_wv-MHH541Of-g';
const UPLOAD_FOLDER_ID = '1hXA7MuOg2TbhfRtVfZeuIZj07Yd_883g';

function doGet(e) {
  const action = e.parameter.action;

  if (action === 'getPegawai') {
    return getPegawai();
  } else if (action === 'getSuratMasuk') {
    return getSuratMasuk();
  } else if (action === 'getMasterData') {
    return getMasterData();
  } else if (action === 'getDashboardStats') {
    return getDashboardStats();
  } else if (action === 'getAppSettings') {
    return getAppSettings();
  } else if (action === 'getAbsensi') {
    return getAbsensi(e.parameter.date);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Invalid action' })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  if (!e.postData) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'No post data' })).setMimeType(ContentService.MimeType.JSON);
  }

  const data = JSON.parse(e.postData.contents);
  const action = data.action;

  if (action === 'savePegawai') {
    return savePegawai(data.payload);
  } else if (action === 'deletePegawai') {
    return deletePegawai(data.id);
  } else if (action === 'saveSuratMasuk') {
    return saveSuratMasuk(data.payload);
  } else if (action === 'deleteSuratMasuk') {
    return deleteSuratMasuk(data.id);
  } else if (action === 'saveMasterData') {
    return saveMasterData(data.payload);
  } else if (action === 'saveAbsensi') {
    return saveAbsensi(data.payload);
  } else if (action === 'updateAbsensiStatus') {
    return updateAbsensiStatus(data.payload);
  } else if (action === 'uploadFile') {
    return uploadFile(data.fileData, data.fileName, data.mimeType);
  } else if (action === 'saveAppSettings') {
    return saveAppSettings(data.payload);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Invalid action' })).setMimeType(ContentService.MimeType.JSON);
}

// --- Helper Functions ---
function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function getDataFromSheet(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Only header or empty

  const headers = data[0];
  const rows = data.slice(1);

  return rows.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// --- Pegawai ---
function getPegawai() {
  const data = getDataFromSheet('DataPegawai');
  return jsonResponse({ status: 'success', data: data });
}

function savePegawai(payload) {
  const sheet = getSheet('DataPegawai');
  const data = sheet.getDataRange().getValues();
  const headers = data.length > 0 ? data[0] : ['id', 'nama', 'nip', 'jenis_kelamin', 'pangkat', 'gol_ruang', 'jabatan', 'agama', 'status_pegawai', 'tanggal_lahir'];

  if (data.length === 0) sheet.appendRow(headers);

  let rowIndex = -1;
  if (!payload.id) payload.id = new Date().getTime().toString();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(payload.id)) {
      rowIndex = i + 1;
      break;
    }
  }

  const rowData = headers.map(header => payload[header] || '');

  if (rowIndex > 0) {
    // Update
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    // Insert
    sheet.appendRow(rowData);
  }

  return jsonResponse({ status: 'success', message: 'Data saved', id: payload.id });
}

function deletePegawai(id) {
  const sheet = getSheet('DataPegawai');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ status: 'success', message: 'Data deleted' });
    }
  }
  return jsonResponse({ status: 'error', message: 'ID not found' });
}

// --- Surat Masuk ---
function getSuratMasuk() {
  const data = getDataFromSheet('SuratMasuk');
  // Parse JSON fields if necessary (like disposisi)
  const parsedData = data.map(item => {
    if (item.disposisi && typeof item.disposisi === 'string') {
      try { item.disposisi = JSON.parse(item.disposisi); } catch (e) { }
    }
    return item;
  });
  return jsonResponse({ status: 'success', data: parsedData });
}

function saveSuratMasuk(payload) {
  const sheet = getSheet('SuratMasuk');
  const data = sheet.getDataRange().getValues();
  const headers = data.length > 0 ? data[0] : ['id', 'noAgenda', 'tglTerima', 'sifatSurat', 'noSurat', 'tglSurat', 'pengirim', 'perihal', 'linkSurat', 'tujuanSurat', 'nomorWa', 'disposisi'];

  if (data.length === 0) sheet.appendRow(headers);

  let rowIndex = -1;
  let existingData = {};

  if (!payload.id) payload.id = 'surat_' + new Date().getTime();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(payload.id)) {
      rowIndex = i + 1;
      // Map existing row to object
      headers.forEach((h, k) => {
        existingData[h] = data[i][k];
      });
      break;
    }
  }

  // Merge payload into existing data for updates
  const finalData = rowIndex > 0 ? { ...existingData, ...payload } : payload;

  // Stringify complex objects
  if (typeof finalData.disposisi === 'object') finalData.disposisi = JSON.stringify(finalData.disposisi);

  // Map back to row array
  const rowData = headers.map(header => finalData[header] || '');

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  return jsonResponse({ status: 'success', message: 'Surat saved', id: payload.id });
}

function deleteSuratMasuk(id) {
  const sheet = getSheet('SuratMasuk');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ status: 'success', message: 'Surat deleted' });
    }
  }
  return jsonResponse({ status: 'error', message: 'ID not found' });
}

// --- Data Master ---
function getMasterData() {
  const sheet = getSheet('MasterData');
  const data = sheet.getDataRange().getValues();
  let result = {};

  if (data.length > 0) {
    const headers = data[0]; // e.g., 'agama', 'pangkat', etc.
    // Initialize arrays
    headers.forEach(h => result[h] = []);

    // Fill arrays, skipping empty cells
    for (let i = 1; i < data.length; i++) {
      headers.forEach((key, colIndex) => {
        const val = data[i][colIndex];
        if (val) result[key].push(val);
      });
    }
  } else {
    // Default Initial Data if empty
    result = {
      agama: ["Katolik", "Protestan", "Islam", "Hindu", "Buddha", "Konghucu"],
      pangkat: ["Pembina Utama", "Pembina Utama Madya", "Pembina Utama Muda", "Pembina Tingkat I", "Pembina", "Penata Tingkat I", "Penata", "Penata Muda Tingkat I", "Penata Muda", "N/A (untuk PPPK)"],
      golongan: ["IV/e", "IV/d", "IV/c", "IV/b", "IV/a", "III/d", "III/c", "III/b", "III/a", "II/d", "II/c", "II/b", "II/a", "I/d", "I/c", "I/b", "I/a", "N/A"],
      eselon: ["II", "III", "IV", "Non Eselon"],
      jabatan: ["Kepala Dinas", "Sekretaris", "Kepala Bidang Kelembagaan", "Kepala Bidang Pemberdayaan", "Kepala Seksi A", "Kepala Seksi B", "Kepala Seksi C", "Analis Data"],
      pendidikan: ["SD", "SMP", "SMA/SMK", "D1", "D2", "D3", "D4", "S1", "S2", "S3"]
    };
    saveMasterData(result); // Persist defaults
  }

  return jsonResponse({ status: 'success', data: result });
}

function saveMasterData(payload) {
  const sheet = getSheet('MasterData');
  sheet.clear();

  const keys = Object.keys(payload);
  if (keys.length === 0) return jsonResponse({ status: 'success' });

  sheet.appendRow(keys); // Headers

  // Find max length
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

  return jsonResponse({ status: 'success', message: 'Master data saved' });
}

// --- Dashboard ---
function getDashboardStats() {
  // Only fetching critical info to avoid heavy payloads
  const pegawai = getDataFromSheet('DataPegawai');
  const surat = getDataFromSheet('SuratMasuk');

  return jsonResponse({
    status: 'success',
    data: {
      totalPegawai: pegawai.length,
      pns: pegawai.filter(p => p.status_pegawai === 'PNS').length,
      pppk: pegawai.filter(p => p.status_pegawai === 'PPPK').length,
      eselon: pegawai.reduce((acc, p) => { let e = p.eselon || 'Non Eselon'; acc[e] = (acc[e] || 0) + 1; return acc; }, {}),
      golongan: pegawai.reduce((acc, p) => { if (p.status_pegawai === 'PNS' && p.gol_ruang) { let g = p.gol_ruang.split('/')[0]; acc['Gol ' + g] = (acc['Gol ' + g] || 0) + 1; } return acc; }, {}),
      agama: pegawai.reduce((acc, p) => { acc[p.agama] = (acc[p.agama] || 0) + 1; return acc; }, {}),
      upcomingBirthdays: pegawai.map(p => ({ nama: p.nama, tanggal: p.tanggal_lahir })).filter(p => !!p.tanggal), // Simplified
      upcomingRetirements: pegawai.map(p => ({ nama: p.nama, tanggal_lahir: p.tanggal_lahir })).filter(p => !!p.tanggal_lahir)
    }
  });
}


// --- Settings & Attendance ---
function getAppSettings() {
  const sheet = getSheet('Settings');
  const data = sheet.getDataRange().getValues();
  let settings = {};

  if (data.length <= 1) {
    // Default settings
    settings = {
      officeLat: -7.9812985, // Example default
      officeLon: 131.2948271,
      jarakMaksimal: 100,
      morning_active: true,
      evening_active: true
    };
    saveAppSettings(settings);
  } else {
    // Assuming Key-Value pairs in columns A and B
    for (let i = 1; i < data.length; i++) {
      let key = data[i][0];
      let val = data[i][1];
      // Convert booleans/numbers
      if (val === 'true') val = true;
      if (val === 'false') val = false;
      if (!isNaN(parseFloat(val)) && isFinite(val) && typeof val !== 'boolean') val = parseFloat(val);
      settings[key] = val;
    }
  }
  return jsonResponse({ status: 'success', data: settings });
}

function saveAppSettings(payload) {
  const sheet = getSheet('Settings');
  sheet.clear();
  sheet.appendRow(['Key', 'Value']);

  const rows = Object.keys(payload).map(key => [key, payload[key]]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
  return jsonResponse({ status: 'success', message: 'Settings saved' });
}

function saveAbsensi(payload) {
  const sheet = getSheet('DataAbsensi');
  const pegawaiSheet = getSheet('DataPegawai');

  // Headers check
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) {
    sheet.appendRow(['Timestamp', 'Tanggal', 'Waktu', 'Nama', 'NIP', 'Sesi', 'Status', 'Keterangan', 'Latitude', 'Longitude', 'DeviceId']);
  }

  // Check double processing
  // Although client checks, backend should also check uniqueness context (Date + NIP + Sesi)
  // Implementing simplified check based on client trust for now to match Firebase speed, 
  // but ideally we iterate to check duplicates.

  // Check Device ID Registration
  const pegawaiData = pegawaiSheet.getDataRange().getValues();
  let headers = pegawaiData[0];
  let idCol = headers.indexOf('id');
  let deviceIdCol = headers.indexOf('registeredDeviceId');
  let nipCol = headers.indexOf('nip');

  // Add registeredDeviceId column if missing
  if (deviceIdCol === -1) {
    deviceIdCol = headers.length;
    pegawaiSheet.getRange(1, deviceIdCol + 1).setValue('registeredDeviceId');
    // Refresh data
  }

  let pegawaiRowIndex = -1;
  for (let i = 1; i < pegawaiData.length; i++) {
    if (String(pegawaiData[i][idCol]) === String(payload.pegawaiId)) {
      pegawaiRowIndex = i + 1;
      const registeredId = pegawaiData[i][deviceIdCol];
      if (registeredId && String(registeredId) !== String(payload.deviceId)) {
        return jsonResponse({ status: 'error', message: 'Device ID mismatch' });
      }
      if (!registeredId) {
        // Register device
        pegawaiSheet.getRange(pegawaiRowIndex, deviceIdCol + 1).setValue(payload.deviceId);
      }
      break;
    }
  }

  // Validation: Check if already present today
  const today = payload.tanggal || new Date().toISOString().split('T')[0]; // Use payload date if allowed/provided
  const rows = sheet.getDataRange().getValues();
  const hasAbsen = rows.some((row, i) => {
    if (i === 0) return false;
    // row[1] is Tanggal, row[4] is NIP, row[5] is Sesi
    // Note: Date comparison needs care with formats. Assuming string match YYYY-MM-DD
    return row[1] === today && String(row[4]) === String(payload.nip) && row[5] === payload.session;
  });

  if (hasAbsen) {
    return jsonResponse({ status: 'error', message: 'Sudah melakukan absensi untuk sesi ini hari ini.' });
  }

  const timestamp = new Date();
  const newRow = [
    timestamp.toISOString(),
    today, // Tanggal
    payload.waktu, // Waktu
    payload.nama,
    payload.nip,
    payload.session,
    payload.status,
    payload.keterangan,
    payload.lokasi.latitude,
    payload.lokasi.longitude,
    payload.deviceId
  ];

  sheet.appendRow(newRow);
  return jsonResponse({ status: 'success', message: 'Absensi berhasil' });
}

function getAbsensi(date) {
  const sheet = getSheet('DataAbsensi');
  const data = sheet.getDataRange().getValues();
  const results = [];

  if (data.length > 1) {
    // Columns: 0:Timestamp, 1:Tanggal, 2:Waktu, 3:Nama, 4:NIP, 5:Sesi, 6:Status, 7:Keterangan, 8:Lat, 9:Lon, 10:DeviceId
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === date) {
        results.push({
          timestamp: data[i][0],
          tanggal: data[i][1],
          waktu: data[i][2],
          nama: data[i][3],
          nip: data[i][4],
          session: data[i][5],
          status: data[i][6],
          keterangan: data[i][7],
          pegawaiId: getPegawaiIdByNip(data[i][4]) // Helper to get ID if needed, or just rely on NIP
        });
      }
    }
  }
  return jsonResponse({ status: 'success', data: results });
}

function updateAbsensiStatus(payload) {
  // Payload: { date, NIP (or nama), status }
  // Logic: Find all records for this user & date, update status. If none, create new "Manual" record.
  const sheet = getSheet('DataAbsensi');
  const data = sheet.getDataRange().getValues();
  let updated = false;

  // Columns: 1:Tanggal, 4:NIP
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === payload.date && String(data[i][4]) === String(payload.nip)) {
      sheet.getRange(i + 1, 7).setValue(payload.status); // Column 6 (0-indexed) is index 7 (1-indexed) -> Status
      updated = true;
    }
  }

  if (!updated) {
    // Create new record
    const newRow = [
      new Date().toISOString(),
      payload.date,
      '-', // Waktu
      payload.nama,
      payload.nip,
      'Manual', // Session
      payload.status,
      'Manual Update',
      0, 0, ''
    ];
    sheet.appendRow(newRow);
  }

  return jsonResponse({ status: 'success', message: 'Status updated' });
}

function getPegawaiIdByNip(nip) {
  // Helper to find ID from DataPegawai. Optional but good for consistency.
  const sheet = getSheet('DataPegawai');
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return null;
  const headers = data[0];
  const idCol = headers.indexOf('id');
  const nipCol = headers.indexOf('nip');

  if (idCol === -1 || nipCol === -1) {
    console.error("Headers 'id' or 'nip' not found in DataPegawai sheet.");
    return null;
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][nipCol]) === String(nip)) return data[i][idCol];
  }
  return null;
}


// --- File Upload ---
function uploadFile(data, name, mimeType) {
  try {
    const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
    const blob = Utilities.newBlob(Utilities.base64Decode(data), mimeType, name);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return jsonResponse({
      status: 'success',
      url: file.getUrl(),
      webViewLink: file.getWebViewLink(),
      id: file.getId()
    });
  } catch (e) {
    return jsonResponse({ status: 'error', message: e.toString() });
  }
}
