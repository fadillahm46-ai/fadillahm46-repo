// =========================================================================
// BACKEND GOOGLE APPS SCRIPT - SE DASHBOARD MONITORING & MANPOWER PORTAL
// =========================================================================

/**
 * Konfigurasi Spreadsheet ID (Opsional jika script terikat ke spreadsheet langsung)
 */
var SPREADSHEET_ID = "";

/**
 * FOLDER GOOGLE DRIVE PENYIMPANAN FOTO DOKUMEN & SERTIFIKAT
 */
var DRIVE_FOLDER_ID = "1pA4e24uLetVNBhj0GmrlV0sS-PvPREC9";

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

  sheets[0].setName('Data Karyawan');
  return sheets[0];
}

/**
 * DAFTAR LENGKAP 27 HEADER KOLOM DATA KARYAWAN (A - AA)
 */
var ALL_KARYAWAN_HEADERS = [
  "ID", "NRP", "Nama", "Perusahaan", "Jabatan", "TglLahir",
  "ExpSIMPER_BIB", "ExpSIMPER_TIA", "ExpSIM_B2", "TglMCU", "ExpMCU",
  "TglLOTOTO", "ExpDangerTag", "TglAwareness", "ExpKetinggian", "TglSIO", "ExpSIO",
  "NoHP", "Password", "StatusAkun",
  "Foto_BIB", "Foto_TIA", "Foto_SIM", "Foto_MCU", "Foto_DTAG", "Foto_KET", "Foto_SIO"
];

/**
 * DAFTAR LENGKAP HEADER KOLOM DATA UNIT
 */
var ALL_UNIT_HEADERS = [
  "ID", "NoUnit", "Model", "ExpKom_BIB", "ExpKom_TIA", "ExpKom_TMA"
];

/**
 * FUNGSI UTAMA: SETUP DATABASE LENGKAP SATU KALI KLIK
 */
function setupDatabaseLengkap() {
  var ss = getSpreadsheet();
  if (!ss) {
    Logger.log("Gagal: Spreadsheet tidak ditemukan.");
    return;
  }

  var sheetKaryawan = getSheetKaryawan(ss);
  var maxCols = Math.max(sheetKaryawan.getMaxColumns(), 1);
  var r1 = sheetKaryawan.getRange(1, 1, 1, maxCols).getDisplayValues()[0];
  var hasNRP = r1.some(function(h) { return String(h).trim().toUpperCase() === "NRP"; });
  var hasData = sheetKaryawan.getLastRow() > 1;

  if (!hasNRP && !hasData) {
    sheetKaryawan.clear();
    if (sheetKaryawan.getMaxColumns() < ALL_KARYAWAN_HEADERS.length) {
      sheetKaryawan.insertColumnsAfter(sheetKaryawan.getMaxColumns(), ALL_KARYAWAN_HEADERS.length - sheetKaryawan.getMaxColumns() + 2);
    }
    sheetKaryawan.getRange(1, 1, 1, ALL_KARYAWAN_HEADERS.length)
      .setValues([ALL_KARYAWAN_HEADERS])
      .setFontWeight("bold")
      .setBackground("#f3f4f6")
      .setHorizontalAlignment("center");
    sheetKaryawan.setFrozenRows(1);
  } else {
    ensureHeadersExist(sheetKaryawan, ALL_KARYAWAN_HEADERS);
  }

  var sheetUnit = ss.getSheetByName('Data Unit');
  if (!sheetUnit) {
    sheetUnit = ss.insertSheet('Data Unit');
    sheetUnit.getRange(1, 1, 1, ALL_UNIT_HEADERS.length)
      .setValues([ALL_UNIT_HEADERS])
      .setFontWeight("bold")
      .setBackground("#f3f4f6")
      .setHorizontalAlignment("center");
    sheetUnit.setFrozenRows(1);
  } else {
    ensureHeadersExist(sheetUnit, ALL_UNIT_HEADERS);
  }

  // Cek akses folder Drive
  try {
    var testFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    Logger.log("Folder Drive Terhubung: " + testFolder.getName());
  } catch(errF) {
    Logger.log("Peringatan Folder Drive: " + errF.toString());
  }

  SpreadsheetApp.flush();
  Logger.log("Database siap & seluruh kolom tersusun rapi.");
}

function tambahSemuaKolomBaru() {
  setupDatabaseLengkap();
}

function setupSheets() {
  setupDatabaseLengkap();
}

function ensureHeadersExist(sheet, headersList) {
  if (!sheet) return;
  var maxCols = Math.max(sheet.getMaxColumns(), 1);
  var r1 = sheet.getRange(1, 1, 1, maxCols).getDisplayValues()[0];
  
  var existingHeaders = [];
  var lastColWithHeader = 0;
  for (var c = 0; c < r1.length; c++) {
    var hName = String(r1[c]).trim();
    if (hName !== "") {
      existingHeaders.push(hName.toLowerCase());
      lastColWithHeader = c + 1;
    }
  }

  var toAdd = [];
  headersList.forEach(function(key) {
    if (existingHeaders.indexOf(String(key).trim().toLowerCase()) === -1) {
      toAdd.push(key);
    }
  });

  if (toAdd.length > 0) {
    var neededCols = lastColWithHeader + toAdd.length;
    if (neededCols > sheet.getMaxColumns()) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), neededCols - sheet.getMaxColumns() + 2);
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

// ROUTING GET
function doGet(e) {
  setupDatabaseLengkap();
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

    setupDatabaseLengkap();

    if (action === 'uploadDocPhoto') {
      return responseJSON(handlePhotoUploadToDrive(payload));
    }
    else if (action === 'saveKaryawan') {
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

/**
 * UPLOAD FOTO LANGSUNG KE GOOGLE DRIVE & SIMPAN TAUTAN KE SHEET
 */
function handlePhotoUploadToDrive(payload) {
  try {
    var ss = getSpreadsheet();
    if (!ss) return { success: false, error: 'Spreadsheet tidak ditemukan.' };
    var sheet = getSheetKaryawan(ss);
    if (!sheet) return { success: false, error: 'Sheet Data Karyawan tidak ditemukan.' };

    var rowId = payload.id;
    var nrp = payload.nrp;
    var docKey = payload.docKey;
    var base64 = payload.base64;
    var colHeader = "Foto_" + docKey;

    if (!base64) {
      return { success: false, error: 'Data foto tidak diterima.' };
    }

    var folder;
    try {
      folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    } catch(errFolder) {
      return { success: false, error: 'Gagal membuka folder Google Drive. Pastikan ID folder benar: ' + errFolder.toString() };
    }

    var cleanBase64 = base64;
    var mimeType = "image/jpeg";
    if (cleanBase64.indexOf(";base64,") > -1) {
      var parts = cleanBase64.split(";base64,");
      mimeType = parts[0].replace("data:", "").trim();
      cleanBase64 = parts[1];
    } else if (cleanBase64.indexOf("base64,") > -1) {
      cleanBase64 = cleanBase64.split("base64,")[1];
    }

    var fileName = (nrp ? nrp : "MP") + "_" + docKey + "_" + Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd_HHmmss") + ".jpg";
    var decoded = Utilities.base64Decode(cleanBase64);
    var blob = Utilities.newBlob(decoded, mimeType, fileName);
    var file = folder.createFile(blob);
    
    // Set hak akses agar gambar dapat tampil langsung
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(errShare) {}

    var fileId = file.getId();
    var driveUrl = "https://lh3.googleusercontent.com/d/" + fileId;

    ensureHeadersExist(sheet, [colHeader]);

    var data = sheet.getDataRange().getValues();
    var displayData = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    var colIndex = headers.indexOf(colHeader);
    var idIndex = headers.indexOf('ID');
    var nrpIndex = headers.indexOf('NRP');

    if (colIndex === -1) {
      return { success: false, error: 'Kolom ' + colHeader + ' tidak ditemukan di sheet.' };
    }

    var targetRow = -1;
    for (var i = 1; i < data.length; i++) {
      var rId = String(data[i][idIndex] || '');
      var rNrp = String(displayData[i][nrpIndex] || '').trim();
      if ((rowId && rId === String(rowId)) || (nrp && rNrp.toUpperCase() === String(nrp).trim().toUpperCase())) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow > -1) {
      sheet.getRange(targetRow, colIndex + 1).setValue(driveUrl);
      SpreadsheetApp.flush();
      return { success: true, url: driveUrl, docKey: docKey, id: rowId, nrp: nrp, fileId: fileId };
    } else {
      return { success: true, url: driveUrl, docKey: docKey, warning: 'Foto tersimpan di Drive, baris sheet belum diperbarui.' };
    }
  } catch(e) {
    return { success: false, error: e.toString() };
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
      
      if (Object.prototype.toString.call(val) === '[object Date]') {
        var ms = val.getTime() - (val.getTimezoneOffset() * 60000);
        val = new Date(ms).toISOString().split('T')[0];
      }
      
      if (headers[j] === 'NoHP' && typeof val === 'string') {
        val = val.replace(/^'/, '');
      }

      obj[headers[j]] = val;
    }
    result.push(obj);
  }
  
  return result.reverse();
}

// MENYIMPAN / EDIT DATA DENGAN INTERSEPSI FOTO KE DRIVE OTOMATIS
function saveData(sheetName, dataObj, primaryKey) {
  var ss = getSpreadsheet();
  if (!ss) return { success: false, error: 'Spreadsheet tidak ditemukan' };
  var sheet = (sheetName === 'Data Karyawan') ? getSheetKaryawan(ss) : (ss.getSheetByName(sheetName) || ss.getSheets()[0]);
  
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
    var key = headers[h];
    var val = dataObj[key];

    if (val === undefined || val === null) {
      val = "";
    } else if (typeof val === 'string' && val.indexOf('data:image/') === 0) {
      // Jika data base64 terkirim, otomatis upload ke Drive dan ganti dengan URL
      try {
        var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        var cleanB64 = val.split(";base64,")[1] || val.split("base64,")[1] || val;
        var mime = val.substring(5, val.indexOf(";")) || "image/jpeg";
        var decoded = Utilities.base64Decode(cleanB64);
        var blob = Utilities.newBlob(decoded, mime, (dataObj.NRP || "MP") + "_" + key + "_" + Date.now() + ".jpg");
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        val = "https://lh3.googleusercontent.com/d/" + file.getId();
        dataObj[key] = val;
      } catch(errB64) {
        val = "";
      }
    } else if (typeof val === 'string' && val.length > 49000) {
      val = val.substring(0, 49000);
    }
    
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
  
  ensureHeadersExist(sheet, ALL_KARYAWAN_HEADERS);

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
