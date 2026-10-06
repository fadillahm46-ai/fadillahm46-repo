// =========================================================================
// BACKEND GOOGLE APPS SCRIPT - SE DASHBOARD MONITORING & MANPOWER PORTAL
// =========================================================================

/**
 * JALANKAN FUNGSI INI DARI APPS SCRIPT UNTUK MENAMBAHKAN KOLOM SECARA INSTAN!
 * Pilih 'tambahSemuaKolomBaru' di menu dropdown atas, lalu klik tombol 'Jalankan' (Run).
 */
function tambahSemuaKolomBaru() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Cari sheet 'Data Karyawan' atau gunakan sheet pertama jika namanya berbeda
  var sheet = ss.getSheetByName('Data Karyawan');
  if (!sheet) {
    sheet = ss.getSheets()[0];
  }

  // Daftar kolom baru yang akan ditambahkan jika belum ada
  var kolomBaru = [
    "NoHP",
    "Password",
    "StatusAkun",
    "Foto_BIB",
    "Foto_TIA",
    "Foto_SIM",
    "Foto_MCU",
    "Foto_DTAG",
    "Foto_KET",
    "Foto_SIO"
  ];

  var lastCol = sheet.getLastColumn();
  var existingHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  
  var kolomDitambahkan = [];

  kolomBaru.forEach(function(namaKolom) {
    // Cek apakah header kolom sudah ada (tidak sensitif huruf besar/kecil)
    var sudahAda = existingHeaders.some(function(header) {
      return String(header).trim().toLowerCase() === namaKolom.toLowerCase();
    });

    if (!sudahAda) {
      lastCol++;
      sheet.getRange(1, lastCol)
        .setValue(namaKolom)
        .setFontWeight("bold")
        .setBackground("#f3f4f6")
        .setHorizontalAlignment("center");
      existingHeaders.push(namaKolom);
      kolomDitambahkan.push(namaKolom);
    }
  });

  SpreadsheetApp.flush();
  Logger.log("Selesai! Kolom yang berhasil ditambahkan: " + kolomDitambahkan.join(", "));
}

// Inisialisasi otomatis jika dipanggil lewat request web
function setupSheets() {
  tambahSemuaKolomBaru();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var headersUnit = ["ID", "NoUnit", "Model", "ExpKom_BIB", "ExpKom_TIA", "ExpKom_TMA"];
  var sheetUnit = ss.getSheetByName('Data Unit');
  if (!sheetUnit) {
    sheetUnit = ss.insertSheet('Data Unit');
    sheetUnit.appendRow(headersUnit);
    sheetUnit.getRange(1, 1, 1, headersUnit.length).setFontWeight("bold").setBackground("#f3f4f6");
  }
}

// =========================================================================
// ROUTING GET (TARIK DATA DARI SHEET KE FRONTEND)
// =========================================================================
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

// =========================================================================
// ROUTING POST (SIMPAN, EDIT, HAPUS, IMPORT DATA)
// =========================================================================
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

// =========================================================================
// FUNGSI PENDUKUNG (HELPERS)
// =========================================================================

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function generateId() {
  return 'id_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
}

// MENGAMBIL DATA DARI SHEET (GET)
function getData(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName) || ss.getSheets()[0];
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
      
      obj[headers[j]] = val;
    }
    result.push(obj);
  }
  
  return result.reverse();
}

// MENYIMPAN ATAU MENGEDIT DATA (POST) -> ANTI DUPLIKAT
function saveData(sheetName, dataObj, primaryKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName) || ss.getSheets()[0];
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  if (!dataObj.ID || String(dataObj.ID).indexOf("temp_") === 0) {
    dataObj.ID = generateId();
  }

  // Set default akun & password jika belum disetel
  if (sheetName === 'Data Karyawan' || sheet.getName() === ss.getSheets()[0].getName()) {
    if (!dataObj.Password) dataObj.Password = String(dataObj.NRP || '').trim();
    if (!dataObj.StatusAkun) dataObj.StatusAkun = 'Aktif';
  }

  var rowIndex = -1;
  var idIndex = headers.indexOf('ID');
  var pkIndex = headers.indexOf(primaryKey);

  for (var i = 1; i < data.length; i++) {
    var rowId = String(data[i][idIndex]);
    var rowPk = String(data[i][pkIndex]);

    if ((dataObj.ID && rowId === String(dataObj.ID)) || 
        (dataObj[primaryKey] && rowPk.toUpperCase() === String(dataObj[primaryKey]).toUpperCase())) {
      rowIndex = i + 1;
      dataObj.ID = rowId;
      
      for (var h = 0; h < headers.length; h++) {
        var key = headers[h];
        if (dataObj[key] === undefined) {
          dataObj[key] = data[i][h];
        }
      }
      break;
    }
  }

  var rowData = [];
  for (var h = 0; h < headers.length; h++) {
    var val = dataObj[headers[h]];
    rowData.push(val !== undefined ? val : "");
  }

  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  return { success: true, id: dataObj.ID };
}

// MENGHAPUS DATA (POST)
function deleteData(sheetName, id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName) || ss.getSheets()[0];
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

// IMPORT MASSAL EXCEL (POST)
function importBulk(sheetName, dataArray, primaryKey) {
  if (!dataArray || dataArray.length === 0) return { success: true };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName) || ss.getSheets()[0];
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idIndex = headers.indexOf('ID');
  var pkIndex = headers.indexOf(primaryKey);

  var existingRows = {}; 
  for(var i = 1; i < data.length; i++) {
    var pkVal = String(data[i][pkIndex]).toUpperCase();
    if(pkVal) existingRows[pkVal] = i + 1;
  }

  dataArray.forEach(function(obj) {
    var pkVal = String(obj[primaryKey]).toUpperCase();
    if (!pkVal) return;

    var rowIndex = existingRows[pkVal];

    if (!obj.ID || String(obj.ID).indexOf("temp_") === 0) {
      obj.ID = generateId();
    }
    if (!obj.Password) obj.Password = pkVal;
    if (!obj.StatusAkun) obj.StatusAkun = 'Aktif';

    var rowData = headers.map(function(h) {
      return obj[h] !== undefined ? obj[h] : "";
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
