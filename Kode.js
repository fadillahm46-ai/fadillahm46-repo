// =========================================================================
// BACKEND GOOGLE APPS SCRIPT - SE DASHBOARD MONITORING & MANPOWER PORTAL
// =========================================================================

/**
 * Konfigurasi Spreadsheet ID (Opsional):
 * Jika script dibuat dari menu: Ekstensi > Apps Script di Spreadsheet, biarkan KOSONG ("").
 * Jika script dibuat terpisah di script.google.com, isi dengan ID spreadsheet Anda.
 */
var SPREADSHEET_ID = "";

/**
 * Helper mendapatkan spreadsheet aktif secara aman
 */
function getSpreadsheet() {
  var ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = SpreadsheetApp.getActive();
  } catch (e) {}

  if (!ss && typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID.trim() !== "") {
    try {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID.trim());
    } catch (err) {}
  }
  return ss;
}

/**
 * Helper mencari Sheet Karyawan secara akurat dan tahan error
 */
function getSheetKaryawan(ss) {
  if (!ss) ss = getSpreadsheet();
  if (!ss) {
    throw new Error("Spreadsheet tidak ditemukan! Pastikan membuka Apps Script dari Spreadsheet (Ekstensi > Apps Script).");
  }

  var sheet = ss.getSheetByName('Data Karyawan');
  if (sheet) return sheet;

  // Cari lembar mana saja yang memiliki kolom 'NRP'
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    if (sheets[s].getName() === 'Data Unit') continue;
    var row1 = sheets[s].getRange(1, 1, 1, Math.min(sheets[s].getMaxColumns(), 30)).getDisplayValues()[0];
    for (var c = 0; c < row1.length; c++) {
      if (String(row1[c]).trim().toUpperCase() === 'NRP') {
        sheets[s].setName('Data Karyawan');
        return sheets[s];
      }
    }
  }
  return sheets[0];
}

/**
 * FUNGSI OTOMATIS MEMBUAT SEMUA KOLOM BARU
 * Pilih fungsi ini di menu dropdown atas Apps Script, lalu klik 'Jalankan'.
 */
function tambahSemuaKolomBaru() {
  var ss = getSpreadsheet();
  if (!ss) {
    Logger.log("Gagal: Spreadsheet tidak ditemukan.");
    return;
  }

  var sheet = getSheetKaryawan(ss);

  var standardHeaders = [
    "NoHP", "Password", "StatusAkun",
    "Foto_BIB", "Foto_TIA", "Foto_SIM", "Foto_MCU", "Foto_DTAG", "Foto_KET", "Foto_SIO"
  ];

  ensureHeadersExist(sheet, standardHeaders);

  var pesan = "Berhasil! Kolom NoHP, Password, StatusAkun, dan seluruh kolom Foto Dokumen sudah aktif di sheet [" + sheet.getName() + "].";
  Logger.log(pesan);
  try {
    SpreadsheetApp.getUi().alert("Hasil Setup", pesan, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {}
}

/**
 * Memastikan kolom baru otomatis terpasang tanpa error batas kolom
 */
function ensureHeadersExist(sheet, newHeadersList) {
  if (!sheet) sheet = getSheetKaryawan();
  var maxCols = sheet.getMaxColumns();
  var r1 = sheet.getRange(1, 1, 1, maxCols).getValues()[0];
  
  var existingHeaders = [];
  var lastColWithHeader = 0;
  for (var c = 0; c < r1.length; c++) {
    var hName = String(r1[c]).trim();
    if (hName !== "") {
      existingHeaders.push(hName);
      lastColWithHeader = c + 1;
    }
  }

  var toAdd = [];
  newHeadersList.forEach(function(key) {
    var exists = existingHeaders.some(function(h) {
      return h.toLowerCase() === key.toLowerCase();
    });
    if (!exists) {
      toAdd.push(key);
    }
  });

  if (toAdd.length > 0) {
    var neededCols = lastColWithHeader + toAdd.length;
    if (neededCols > sheet.getMaxColumns()) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), neededCols - sheet.getMaxColumns() + 3);
    }
    for (var i = 0; i < toAdd.length; i++) {
      var colIdx = lastColWithHeader + i + 1;
      sheet.getRange(1, colIdx)
        .setValue(toAdd[i])
        .setFontWeight("bold")
        .setBackground("#f3f4f6")
        .setHorizontalAlignment("center");
    }
    SpreadsheetApp.flush();
  }
}

/**
 * Inisialisasi awal saat request pertama kali masuk
 */
function setupSheets() {
  var ss = getSpreadsheet();
  if (!ss) return;

  var sheetKaryawan = getSheetKaryawan(ss);
  var standardHeaders = [
    "NoHP", "Password", "StatusAkun",
    "Foto_BIB", "Foto_TIA", "Foto_SIM", "Foto_MCU", "Foto_DTAG", "Foto_KET", "Foto_SIO"
  ];
  ensureHeadersExist(sheetKaryawan, standardHeaders);

  var headersUnit = ["ID", "NoUnit", "Model", "ExpKom_BIB", "ExpKom_TIA", "ExpKom_TMA"];
  var sheetUnit = ss.getSheetByName('Data Unit');
  if (!sheetUnit) {
    sheetUnit = ss.insertSheet('Data Unit');
    sheetUnit.appendRow(headersUnit);
    sheetUnit.getRange(1, 1, 1, headersUnit.length).setFontWeight("bold").setBackground("#f3f4f6");
  }
}

// ROUTING GET
function doGet(e) {
  setupSheets();
  var action = e && e.parameter ? e.parameter.action : '';
  
  if (action === 'getKaryawan') {
    return responseJSON(getData('Data Karyawan'));
  } else if (action === 'getUnit') {
    return responseJSON(getData('Data Unit'));
  }
  
  return responseJSON({ error: 'Action GET tidak valid' });
}

// ROUTING POST
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;

    setupSheets();

    if (action === 'saveKaryawan') {
      return responseJSON(saveData('Data Karyawan', payload.data, 'NRP'));
    } 
    else if (action === 'deleteKaryawan') {
      return responseJSON(deleteData('Data Karyawan', payload.id));
    } 
    else if (action === 'saveUnit') {
      return responseJSON(saveData('Data Unit', payload.data, 'NoUnit'));
    } 
    else if (action === 'deleteUnit') {
      return responseJSON(deleteData('Data Unit', payload.id));
    } 
    else if (action === 'importKaryawanBulk') {
      return responseJSON(importBulk('Data Karyawan', payload.data, 'NRP'));
    }

    return responseJSON({ error: 'Action POST tidak valid' });

  } catch (err) {
    return responseJSON({ error: err.toString() });
  }
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function generateId() {
  return 'id_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
}

// MENGAMBIL DATA DARI SHEET (GET)
function getData(sheetName) {
  var ss = getSpreadsheet();
  if (!ss) return [];
  var sheet = (sheetName === 'Data Karyawan') ? getSheetKaryawan(ss) : (ss.getSheetByName(sheetName) || ss.getSheets()[0]);
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var headers = data[0];
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      
      // Standarisasi tanggal
      if (Object.prototype.toString.call(val) === '[object Date]') {
        var ms = val.getTime() - (val.getTimezoneOffset() * 60000);
        val = new Date(ms).toISOString().split('T')[0];
      }
      
      // Standarisasi NoHP jika ada awalan kutip
      if (headers[j] === 'NoHP' && typeof val === 'string') {
        val = val.replace(/^'/, '');
      }

      obj[headers[j]] = val;
    }
    result.push(obj);
  }
  
  return result.reverse();
}

// MENYIMPAN / EDIT DATA DENGAN PROTEKSI BATAS SEL 50.000 KARAKTER
function saveData(sheetName, dataObj, primaryKey) {
  var ss = getSpreadsheet();
  if (!ss) return { success: false, error: 'Spreadsheet tidak ditemukan' };
  var sheet = (sheetName === 'Data Karyawan') ? getSheetKaryawan(ss) : (ss.getSheetByName(sheetName) || ss.getSheets()[0]);
  
  // Pastikan kolom dataObj ada di sheet
  var incomingKeys = [];
  for (var k in dataObj) {
    if (k) incomingKeys.push(k);
  }
  ensureHeadersExist(sheet, incomingKeys);

  var data = sheet.getDataRange().getValues();
  var displayData = sheet.getDataRange().getDisplayValues();
  var headers = data[0];

  if (!dataObj.ID || String(dataObj.ID).indexOf("temp_") === 0) {
    dataObj.ID = generateId();
  }

  if (sheetName === 'Data Karyawan' || sheet.getName() === 'Data Karyawan') {
    if (!dataObj.Password) dataObj.Password = String(dataObj.NRP || '').trim();
    if (!dataObj.StatusAkun) dataObj.StatusAkun = 'Aktif';
  }

  var rowIndex = -1;
  var idIndex = headers.indexOf('ID');
  var pkIndex = headers.indexOf(primaryKey);

  for (var i = 1; i < data.length; i++) {
    var rowId = String(data[i][idIndex] || '');
    var rowPk = String(displayData[i][pkIndex] || '').trim();
    var targetPk = String(dataObj[primaryKey] || '').trim();

    if ((dataObj.ID && rowId === String(dataObj.ID)) || 
        (targetPk && rowPk.toUpperCase() === targetPk.toUpperCase())) {
      rowIndex = i + 1;
      if (rowId) dataObj.ID = rowId;
      
      for (var h = 0; h < headers.length; h++) {
        var key = headers[h];
        if (dataObj[key] === undefined && data[i][h] !== undefined) {
          dataObj[key] = data[i][h];
        }
      }
      break;
    }
  }

  var rowData = [];
  for (var h = 0; h < headers.length; h++) {
    var val = dataObj[headers[h]];
    if (val === undefined || val === null) {
      val = "";
    } else if (typeof val === 'string' && val.length > 49000) {
      val = val.substring(0, 49000); // Proteksi batas karakter Google Sheets
    }
    
    // Simpan NoHP sebagai teks agar angka '0' depan tidak hilang
    if (headers[h] === 'NoHP' && val !== "") {
      val = "'" + String(val).replace(/^'/, '');
    }
    rowData.push(val);
  }

  if (headers.length > sheet.getMaxColumns()) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns() + 2);
  }

  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  return { success: true, id: dataObj.ID };
}

// MENGHAPUS DATA
function deleteData(sheetName, id) {
  var ss = getSpreadsheet();
  if (!ss) return { success: false, error: 'Spreadsheet tidak ditemukan' };
  var sheet = (sheetName === 'Data Karyawan') ? getSheetKaryawan(ss) : (ss.getSheetByName(sheetName) || ss.getSheets()[0]);
  if (!sheet) return { success: false, error: 'Sheet tidak ditemukan' };

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idIndex = headers.indexOf('ID');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'ID data tidak ditemukan' };
}

// IMPORT MASSAL EXCEL
function importBulk(sheetName, dataArray, primaryKey) {
  if (!dataArray || dataArray.length === 0) return { success: true };

  var ss = getSpreadsheet();
  if (!ss) return { success: false, error: 'Spreadsheet tidak ditemukan' };
  var sheet = (sheetName === 'Data Karyawan') ? getSheetKaryawan(ss) : (ss.getSheetByName(sheetName) || ss.getSheets()[0]);
  
  var standardHeaders = [
    "NoHP", "Password", "StatusAkun",
    "Foto_BIB", "Foto_TIA", "Foto_SIM", "Foto_MCU", "Foto_DTAG", "Foto_KET", "Foto_SIO"
  ];
  ensureHeadersExist(sheet, standardHeaders);

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idIndex = headers.indexOf('ID');
  var pkIndex = headers.indexOf(primaryKey);

  var existingRows = {}; 
  for (var i = 1; i < data.length; i++) {
    var pkVal = String(data[i][pkIndex]).trim().toUpperCase();
    if (pkVal) existingRows[pkVal] = i + 1;
  }

  dataArray.forEach(function(obj) {
    var pkVal = String(obj[primaryKey]).trim().toUpperCase();
    if (!pkVal) return;

    var rowIndex = existingRows[pkVal];

    if (!obj.ID || String(obj.ID).indexOf("temp_") === 0) {
      obj.ID = generateId();
    }
    if (!obj.Password) obj.Password = pkVal;
    if (!obj.StatusAkun) obj.StatusAkun = 'Aktif';

    var rowData = headers.map(function(h) {
      var val = obj[h] !== undefined ? obj[h] : "";
      if (h === 'NoHP' && val !== "") val = "'" + String(val).replace(/^'/, '');
      return val;
    });

    if (rowIndex) {
      rowData[idIndex] = sheet.getRange(rowIndex, idIndex + 1).getValue();
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
      existingRows[pkVal] = sheet.getLastRow();
    }
  });

  return { success: true, count: dataArray.length };
}
