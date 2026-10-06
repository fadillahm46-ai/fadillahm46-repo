// =========================================================================
// BACKEND GOOGLE APPS SCRIPT - SE DASHBOARD MONITORING & GOOGLE DRIVE UPLOAD
// =========================================================================

var DRIVE_FOLDER_ID = "1pA4e24uLetVNBhj0GmrlV0sS-PvPREC9";

var ALL_KARYAWAN_HEADERS = [
  "ID", "NRP", "Nama", "Perusahaan", "Jabatan", "TglLahir",
  "ExpSIMPER_BIB", "ExpSIMPER_TIA", "ExpSIM_B2", "TglMCU", "ExpMCU",
  "TglLOTOTO", "ExpDangerTag", "TglAwareness", "ExpKetinggian", "TglSIO", "ExpSIO",
  "NoHP", "Password", "StatusAkun",
  "Foto_BIB", "Foto_TIA", "Foto_SIM", "Foto_MCU", "Foto_DTAG", "Foto_KET", "Foto_SIO"
];

var ALL_UNIT_HEADERS = [
  "ID", "NoUnit", "Model", "ExpKom_BIB", "ExpKom_TIA", "ExpKom_TMA"
];

function getSpreadsheet() {
  var ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = SpreadsheetApp.getActive();
  } catch (e) {}
  return ss;
}

function getSheetKaryawan(ss) {
  if (!ss) ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Data Karyawan') || ss.getSheetByName('Karyawan');
  if (!sheet) {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().toLowerCase().indexOf('unit') === -1) {
        sheet = sheets[i];
        sheet.setName('Data Karyawan');
        break;
      }
    }
  }
  if (!sheet) sheet = ss.insertSheet('Data Karyawan');
  return sheet;
}

function getSheetUnit(ss) {
  if (!ss) ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Data Unit') || ss.getSheetByName('Unit');
  if (!sheet) sheet = ss.insertSheet('Data Unit');
  return sheet;
}

// FUNGSI PENGUJIAN IZIN GOOGLE DRIVE (JALANKAN DARI EDITOR APPS SCRIPT)
function ujiIzinDriveDanFolder() {
  var folder = dapatkanFolderDrive();
  Logger.log("BERHASIL! Folder aktif: " + folder.getName() + " (ID: " + folder.getId() + ")");
  SpreadsheetApp.getActiveSpreadsheet().toast("Koneksi Google Drive Berhasil Terhubung!", "Sukses", 5);
}

function dapatkanFolderDrive() {
  var folder = null;
  try {
    if (DRIVE_FOLDER_ID && DRIVE_FOLDER_ID.trim() !== "") {
      folder = DriveApp.getFolderById(DRIVE_FOLDER_ID.trim());
    }
  } catch (errId) {
    Logger.log("Peringatan ID Folder: " + errId.toString());
  }

  // Jika folder ID khusus tidak dapat dibuka, buat/gunakan folder cadangan otomatis
  if (!folder) {
    var namaFolder = "SE_Monitoring_Photos";
    var folders = DriveApp.getFoldersByName(namaFolder);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(namaFolder);
    }
  }
  return folder;
}

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : '';
  if (action === 'getUnit') {
    return responseJSON(getUnitData());
  }
  return responseJSON(getKaryawanData());
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(30000);

  try {
    var contents = e.postData ? e.postData.contents : "{}";
    var payload = JSON.parse(contents);
    var action = payload.action;

    // Mendukung baik 'uploadPhoto' maupun 'uploadDocPhoto'
    if (action === 'uploadPhoto' || action === 'uploadDocPhoto') {
      return responseJSON(handlePhotoUploadToDrive(payload));
    }
    else if (action === 'saveKaryawan') {
      return responseJSON(saveKaryawanData(payload.data));
    }
    else if (action === 'deleteKaryawan') {
      return responseJSON(deleteKaryawanData(payload.id));
    }
    else if (action === 'saveUnit') {
      return responseJSON(saveUnitData(payload.data));
    }
    else if (action === 'deleteUnit') {
      return responseJSON(deleteUnitData(payload.id));
    }
    else if (action === 'importKaryawanBulk') {
      return responseJSON(importBulkData(payload.data));
    }

    return responseJSON({ success: false, status: 'error', message: 'Action tidak dikenal' });

  } catch (err) {
    return responseJSON({ success: false, status: 'error', error: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function handlePhotoUploadToDrive(payload) {
  try {
    var folder = dapatkanFolderDrive();
    var base64 = payload.base64 || "";
    var docKey = payload.docKey || "DOC";
    var nrp = payload.nrp || "";
    var id = payload.id || "";

    if (!base64) {
      return { success: false, status: "error", error: "Data gambar tidak diterima" };
    }

    var cleanBase64 = base64;
    if (cleanBase64.indexOf(",") > -1) {
      cleanBase64 = cleanBase64.split(",")[1];
    }
    cleanBase64 = cleanBase64.replace(/\s/g, "");

    var decoded = Utilities.base64Decode(cleanBase64);
    var fileName = (nrp ? nrp : "MP") + "_" + docKey + "_" + Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd_HHmmss") + ".jpg";
    var blob = Utilities.newBlob(decoded, "image/jpeg", fileName);
    var file = folder.createFile(blob);

    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (eShare) {}

    var fileId = file.getId();
    var directUrl = "https://lh3.googleusercontent.com/d/" + fileId;
    var targetColHeader = "Foto_" + docKey;

    var ss = getSpreadsheet();
    var sheet = getSheetKaryawan(ss);
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });

    var colIndex = headers.indexOf(targetColHeader);
    if (colIndex === -1) {
      sheet.getRange(1, headers.length + 1).setValue(targetColHeader);
      colIndex = headers.length;
      headers.push(targetColHeader);
    }

    var rowIndex = -1;
    var idIndex = headers.indexOf("ID");
    var nrpIndex = headers.indexOf("NRP");

    for (var r = 1; r < data.length; r++) {
      var rId = idIndex > -1 ? String(data[r][idIndex] || "").trim() : "";
      var rNrp = nrpIndex > -1 ? String(data[r][nrpIndex] || "").trim() : "";
      if ((id && rId === String(id).trim()) || (nrp && rNrp.toUpperCase() === String(nrp).trim().toUpperCase())) {
        rowIndex = r + 1;
        break;
      }
    }

    if (rowIndex > -1) {
      sheet.getRange(rowIndex, colIndex + 1).setValue(directUrl);
      SpreadsheetApp.flush();
    }

    return {
      success: true,
      status: "success",
      url: directUrl,
      fileId: fileId,
      docKey: docKey,
      rowIndex: rowIndex
    };

  } catch (errUpload) {
    return {
      success: false,
      status: "error",
      error: errUpload.toString()
    };
  }
}

function getKaryawanData() {
  var ss = getSpreadsheet();
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
        var ms = cellVal.getTime() - (cellVal.getTimezoneOffset() * 60000);
        cellVal = new Date(ms).toISOString().split('T')[0];
      }
      if (headerName === 'NoHP' && typeof cellVal === 'string') {
        cellVal = cellVal.replace(/^'/, '');
      }

      obj[headerName] = cellVal !== undefined && cellVal !== null ? String(cellVal) : '';
      if (obj[headerName] !== '') hasData = true;
    }

    if (hasData && (obj.NRP || obj.Nama || obj.ID)) {
      if (!obj.StatusAkun) obj.StatusAkun = 'Aktif';
      result.push(obj);
    }
  }
  return result.reverse();
}

function saveKaryawanData(item) {
  var ss = getSpreadsheet();
  var sheet = getSheetKaryawan(ss);
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });

  for (var key in item) {
    if (headers.indexOf(key) === -1) {
      headers.push(key);
      sheet.getRange(1, headers.length).setValue(key);
    }
  }

  var rowIndex = -1;
  var idIndex = headers.indexOf('ID');
  var nrpIndex = headers.indexOf('NRP');

  for (var r = 1; r < data.length; r++) {
    var rId = idIndex > -1 ? String(data[r][idIndex] || '').trim() : '';
    var rNrp = nrpIndex > -1 ? String(data[r][nrpIndex] || '').trim() : '';
    if ((item.ID && rId === String(item.ID).trim()) || (item.NRP && rNrp.toUpperCase() === String(item.NRP).trim().toUpperCase())) {
      rowIndex = r + 1;
      if (rId) item.ID = rId;
      break;
    }
  }

  if (rowIndex === -1) {
    if (!item.ID || String(item.ID).indexOf('temp_') === 0) item.ID = Utilities.getUuid();
    if (!item.Password) item.Password = String(item.NRP || '').trim();
    if (!item.StatusAkun) item.StatusAkun = 'Aktif';

    var newRow = headers.map(function(h) {
      var val = item[h] || '';
      if (h === 'NoHP' && val !== '') val = "'" + String(val).replace(/^'/, '');
      return val;
    });
    sheet.appendRow(newRow);
  } else {
    for (var h = 0; h < headers.length; h++) {
      var hName = headers[h];
      if (item[hName] !== undefined) {
        var v = item[hName];
        if (hName === 'NoHP' && v !== '') v = "'" + String(v).replace(/^'/, '');
        sheet.getRange(rowIndex, h + 1).setValue(v);
      }
    }
  }

  return { success: true, status: 'success', data: item };
}

function deleteKaryawanData(id) {
  var ss = getSpreadsheet();
  var sheet = getSheetKaryawan(ss);
  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) === String(id) || String(data[r][1]) === String(id)) {
      sheet.deleteRow(r + 1);
      return { success: true, status: 'success' };
    }
  }
  return { success: false, status: 'not_found' };
}

function getUnitData() {
  var ss = getSpreadsheet();
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
        var ms = val.getTime() - (val.getTimezoneOffset() * 60000);
        val = new Date(ms).toISOString().split('T')[0];
      }
      obj[headers[j]] = val !== undefined && val !== null ? String(val) : '';
    }
    if (obj.NoUnit || obj.Model) result.push(obj);
  }
  return result;
}

function saveUnitData(item) {
  var ss = getSpreadsheet();
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
    if (!item.ID || String(item.ID).indexOf('temp_') === 0) item.ID = Utilities.getUuid();
    var newRow = headers.map(function(h) { return item[h] || ''; });
    sheet.appendRow(newRow);
  } else {
    for (var h = 0; h < headers.length; h++) {
      if (item[headers[h]] !== undefined) {
        sheet.getRange(rowIndex, h + 1).setValue(item[headers[h]]);
      }
    }
  }
  return { success: true, status: 'success' };
}

function deleteUnitData(id) {
  var ss = getSpreadsheet();
  var sheet = getSheetUnit(ss);
  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) === String(id) || String(data[r][1]) === String(id)) {
      sheet.deleteRow(r + 1);
      return { success: true, status: 'success' };
    }
  }
  return { success: false, status: 'not_found' };
}

function importBulkData(items) {
  if (!items || !items.length) return { success: true, status: 'empty' };
  for (var i = 0; i < items.length; i++) {
    saveKaryawanData(items[i]);
  }
  return { success: true, status: 'success', count: items.length };
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupDatabaseLengkap() {
  var ss = getSpreadsheet();
  var sKaryawan = getSheetKaryawan(ss);
  if (sKaryawan.getLastRow() === 0) {
    sKaryawan.getRange(1, 1, 1, ALL_KARYAWAN_HEADERS.length).setValues([ALL_KARYAWAN_HEADERS]);
  } else {
    var curHeaders = sKaryawan.getRange(1, 1, 1, sKaryawan.getLastColumn()).getValues()[0];
    for (var i = 0; i < ALL_KARYAWAN_HEADERS.length; i++) {
      if (curHeaders.indexOf(ALL_KARYAWAN_HEADERS[i]) === -1) {
        sKaryawan.getRange(1, sKaryawan.getLastColumn() + 1).setValue(ALL_KARYAWAN_HEADERS[i]);
      }
    }
  }

  var sUnit = getSheetUnit(ss);
  if (sUnit.getLastRow() === 0) {
    sUnit.getRange(1, 1, 1, ALL_UNIT_HEADERS.length).setValues([ALL_UNIT_HEADERS]);
  }

  ujiIzinDriveDanFolder();
}
