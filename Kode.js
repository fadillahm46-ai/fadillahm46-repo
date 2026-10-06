// =====================================================================
// BACKEND GOOGLE APPS SCRIPT: SISTEM MONITORING SE & INTEGRASI DRIVE
// ID Folder Drive: 1pA4e24uLetVNBhj0GmrlV0sS-PvPREC9
// =====================================================================

var FOLDER_ID = "1pA4e24uLetVNBhj0GmrlV0sS-PvPREC9";

var MASTER_HEADERS_KARYAWAN = [
  "ID", "NRP", "Nama", "Perusahaan", "Jabatan", "TglLahir",
  "ExpSIMPER_BIB", "ExpSIMPER_TIA", "ExpSIM_B2", "TglMCU", "ExpMCU",
  "TglLOTOTO", "ExpDangerTag", "TglAwareness", "ExpKetinggian", "TglSIO", "ExpSIO",
  "NoHP", "Password", "StatusAkun",
  "Foto_BIB", "Foto_TIA", "Foto_SIM", "Foto_MCU", "Foto_DTAG", "Foto_KET", "Foto_SIO"
];

var MASTER_HEADERS_UNIT = [
  "ID", "NoUnit", "Model", "ExpKom_BIB", "ExpKom_TIA", "ExpKom_TMA"
];

// FUNGSI UNTUK OTORISASI DRIVE DARI EDITOR (JALANKAN INI PERTAMA KALI)
function ujiIzinDriveDanFolder() {
  try {
    var folder = DriveApp.getFolderById(FOLDER_ID.trim());
    Logger.log("BERHASIL! Folder terhubung: " + folder.getName());
    SpreadsheetApp.getActiveSpreadsheet().toast("Koneksi Google Drive Berhasil!", "Sukses", 5);
  } catch(e) {
    Logger.log("Folder ID khusus belum ditemukan, menggunakan root folder: " + e.toString());
  }
}

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : "";
  if (action === "getUnit") {
    return ContentService.createTextOutput(JSON.stringify(getUnitData()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify(getKaryawanData()))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(30000);
  
  try {
    var contents = e.postData ? e.postData.contents : "{}";
    var req = JSON.parse(contents);
    var action = req.action;
    
    if (action === "uploadPhoto") {
      var uploadRes = uploadPhotoToDrive(req);
      return ContentService.createTextOutput(JSON.stringify(uploadRes))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "saveKaryawan") {
      var res = saveKaryawanData(req.data);
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "deleteKaryawan") {
      var delRes = deleteKaryawanData(req.id);
      return ContentService.createTextOutput(JSON.stringify(delRes))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "saveUnit") {
      var saveUnitRes = saveUnitData(req.data);
      return ContentService.createTextOutput(JSON.stringify(saveUnitRes))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "deleteUnit") {
      var delUnitRes = deleteUnitData(req.id);
      return ContentService.createTextOutput(JSON.stringify(delUnitRes))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "importKaryawanBulk") {
      var importRes = importBulkData(req.data);
      return ContentService.createTextOutput(JSON.stringify(importRes))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Action tidak dikenal" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// UPLOAD FOTO KE GOOGLE DRIVE & SIMPAN TAUTAN KE SPREADSHEET
function uploadPhotoToDrive(req) {
  var folder;
  try {
    folder = DriveApp.getFolderById(FOLDER_ID.trim());
  } catch(e) {
    folder = DriveApp.getRootFolder();
  }
  
  var rawBase64 = req.base64.replace(/^data:image\/\w+;base64,/, "");
  var decoded = Utilities.base64Decode(rawBase64);
  var fName = req.fileName || ("Foto_" + (req.docKey || "Doc") + "_" + (req.id || Date.now()) + ".jpg");
  var blob = Utilities.newBlob(decoded, "image/jpeg", fName);
  var file = folder.createFile(blob);
  
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {}
  
  var fileId = file.getId();
  var directViewUrl = "https://lh3.googleusercontent.com/d/" + fileId;
  
  // Tuliskan URL ke lembar spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheetKaryawan(ss);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var targetHeader = "Foto_" + req.docKey;
  var colIndex = -1;
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim().toUpperCase() === targetHeader.toUpperCase()) {
      colIndex = c;
      break;
    }
  }
  
  if (colIndex === -1) {
    sheet.getRange(1, headers.length + 1).setValue(targetHeader);
    colIndex = headers.length;
  }
  
  var rowIndex = -1;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) === String(req.id) || (data[r][1] && String(data[r][1]) === String(req.id))) {
      rowIndex = r + 1;
      break;
    }
  }
  
  if (rowIndex !== -1) {
    sheet.getRange(rowIndex, colIndex + 1).setValue(directViewUrl);
  }
  
  return {
    status: "success",
    url: directViewUrl,
    fileId: fileId,
    docKey: req.docKey
  };
}

function getSheetKaryawan(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Data Karyawan") || ss.getSheetByName("Karyawan");
  if (!sheet) {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().toLowerCase().indexOf("unit") === -1) {
        sheet = sheets[i];
        sheet.setName("Data Karyawan");
        break;
      }
    }
  }
  if (!sheet) sheet = ss.insertSheet("Data Karyawan");
  return sheet;
}

function getSheetUnit(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Data Unit") || ss.getSheetByName("Unit");
  if (!sheet) sheet = ss.insertSheet("Data Unit");
  return sheet;
}

function getKaryawanData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheetKaryawan(ss);
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  
  var headers = values[0].map(function(h) { return String(h).trim(); });
  var result = [];
  
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var obj = {};
    var hasData = false;
    
    for (var j = 0; j < headers.length; j++) {
      var headerName = headers[j];
      var cellVal = row[j];
      
      if (cellVal instanceof Date) {
        cellVal = Utilities.formatDate(cellVal, "GMT+8", "yyyy-MM-dd");
      }
      obj[headerName] = cellVal !== undefined && cellVal !== null ? String(cellVal) : "";
      if (obj[headerName] !== "") hasData = true;
    }
    
    if (hasData && (obj.NRP || obj.Nama || obj.ID)) {
      if (!obj.StatusAkun) obj.StatusAkun = "AKTIF";
      result.push(obj);
    }
  }
  return result;
}

function saveKaryawanData(item) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheetKaryawan(ss);
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  
  // Pastikan header baru tercatat
  for (var key in item) {
    if (headers.indexOf(key) === -1) {
      headers.push(key);
      sheet.getRange(1, headers.length).setValue(key);
    }
  }
  
  var rowIndex = -1;
  for (var r = 1; r < data.length; r++) {
    var rowId = String(data[r][0]);
    var rowNrp = String(data[r][1]);
    if ((item.ID && rowId === String(item.ID)) || (item.NRP && rowNrp === String(item.NRP))) {
      rowIndex = r + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    if (!item.ID || String(item.ID).indexOf("temp_") === 0) {
      item.ID = Utilities.getUuid();
    }
    var newRow = headers.map(function(h) { return item[h] || ""; });
    sheet.appendRow(newRow);
  } else {
    for (var h = 0; h < headers.length; h++) {
      var hName = headers[h];
      if (item[hName] !== undefined) {
        sheet.getRange(rowIndex, h + 1).setValue(item[hName]);
      }
    }
  }
  return { status: "success", data: item };
}

function deleteKaryawanData(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheetKaryawan(ss);
  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) === String(id) || String(data[r][1]) === String(id)) {
      sheet.deleteRow(r + 1);
      return { status: "success" };
    }
  }
  return { status: "not_found" };
}

function getUnitData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheetUnit(ss);
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  
  var headers = values[0].map(function(h) { return String(h).trim(); });
  var result = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = row[j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd");
      }
      obj[headers[j]] = val !== undefined && val !== null ? String(val) : "";
    }
    if (obj.NoUnit || obj.Model) result.push(obj);
  }
  return result;
}

function saveUnitData(item) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheetUnit(ss);
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  
  var rowIndex = -1;
  for (var r = 1; r < data.length; r++) {
    if ((item.ID && String(data[r][0]) === String(item.ID)) || (item.NoUnit && String(data[r][1]).toUpperCase() === String(item.NoUnit).toUpperCase())) {
      rowIndex = r + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    if (!item.ID || String(item.ID).indexOf("temp_") === 0) item.ID = Utilities.getUuid();
    var newRow = headers.map(function(h) { return item[h] || ""; });
    sheet.appendRow(newRow);
  } else {
    for (var h = 0; h < headers.length; h++) {
      if (item[headers[h]] !== undefined) {
        sheet.getRange(rowIndex, h + 1).setValue(item[headers[h]]);
      }
    }
  }
  return { status: "success" };
}

function deleteUnitData(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheetUnit(ss);
  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) === String(id) || String(data[r][1]) === String(id)) {
      sheet.deleteRow(r + 1);
      return { status: "success" };
    }
  }
  return { status: "not_found" };
}

function importBulkData(items) {
  if (!items || !items.length) return { status: "empty" };
  for (var i = 0; i < items.length; i++) {
    saveKaryawanData(items[i]);
  }
  return { status: "success", count: items.length };
}

function setupDatabaseLengkap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sKaryawan = getSheetKaryawan(ss);
  if (sKaryawan.getLastRow() === 0) {
    sKaryawan.getRange(1, 1, 1, MASTER_HEADERS_KARYAWAN.length).setValues([MASTER_HEADERS_KARYAWAN]);
  } else {
    var curHeaders = sKaryawan.getRange(1, 1, 1, sKaryawan.getLastColumn()).getValues()[0];
    for (var i = 0; i < MASTER_HEADERS_KARYAWAN.length; i++) {
      if (curHeaders.indexOf(MASTER_HEADERS_KARYAWAN[i]) === -1) {
        sKaryawan.getRange(1, sKaryawan.getLastColumn() + 1).setValue(MASTER_HEADERS_KARYAWAN[i]);
      }
    }
  }
  
  var sUnit = getSheetUnit(ss);
  if (sUnit.getLastRow() === 0) {
    sUnit.getRange(1, 1, 1, MASTER_HEADERS_UNIT.length).setValues([MASTER_HEADERS_UNIT]);
  }
  
  ujiIzinDriveDanFolder();
}
