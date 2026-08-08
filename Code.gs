function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sistem Jurnal Digital - SMPN 3 Babelan')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 1. FUNGSI INISIALISASI DATABASE OTOMATIS
function initDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Tab Pengaturan
  var sheetSetup = ss.getSheetByName("SetUpDatabase");
  if (!sheetSetup) {
    sheetSetup = ss.insertSheet("SetUpDatabase");
    sheetSetup.appendRow(["Key", "Value"]);
    sheetSetup.getRange("A1:B1").setFontWeight("bold").setBackground("#d1fae5");
    var defaults = [
      ["kop1", "PEMERINTAH KABUPATEN BEKASI"],
      ["kop2", "DINAS PENDIDIKAN"],
      ["namaSekolah", "SMP NEGERI 3 BABELAN"],
      ["alamat", "Jl. KH. Noer Ali, Babelan, Kabupaten Bekasi, Jawa Barat"],
      ["namaKepsek", "Dra. Hj. ULFAH, M.M"],
      ["nipKepsek", "196911051995122002"],
      ["runningText", "Selamat Datang di Sistem Jurnal Harian Pegawai SMP Negeri 3 Babelan"],
      ["logoBase64", ""],
      ["ttdKepsekBase64", ""],
      ["lastUpdated", new Date().getTime().toString()]
    ];
    sheetSetup.getRange(2, 1, defaults.length, 2).setValues(defaults);
  }

  // Tab Pegawai
  var sheetPegawai = ss.getSheetByName("Pegawai");
  if (!sheetPegawai) {
    sheetPegawai = ss.insertSheet("Pegawai");
    sheetPegawai.appendRow(["nip", "nama", "jabatan", "jenis", "email", "username", "password", "ttdBase64"]);
    sheetPegawai.getRange("A1:H1").setFontWeight("bold").setBackground("#dbeafe");
  }

  // Tab Jurnal
  var sheetJurnal = ss.getSheetByName("Jurnal");
  if (!sheetJurnal) {
    sheetJurnal = ss.insertSheet("Jurnal");
    sheetJurnal.appendRow(["id", "jenisJurnal", "tanggal", "pegawai", "activities"]);
    sheetJurnal.getRange("A1:E1").setFontWeight("bold").setBackground("#fef9c3");
  }
}

// 2. SISTEM PEMBACA DATA PANJANG (ANTI LIMIT 50.000 KARAKTER)
function readChunkedRow(rowArray, startIndex) {
    var fullString = "";
    for (var i = startIndex; i < rowArray.length; i++) {
        if (rowArray[i]) {
            fullString += rowArray[i];
        }
    }
    return fullString;
}

// 3. SISTEM PENYIMPAN DATA PANJANG (AUTO-SPLIT KOLOM)
function writeChunkedData(sheet, rowNum, startCol, fullString) {
   var maxCol = sheet.getMaxColumns();
   if(maxCol >= startCol) {
     sheet.getRange(rowNum, startCol, 1, maxCol - startCol + 1).clearContent();
   }
   
   var strVal = fullString ? fullString.toString() : "";
   var chunks = [];
   for (var c = 0; c < strVal.length; c += 49000) {
     chunks.push(strVal.substring(c, c + 49000));
   }
   if (chunks.length === 0) chunks.push("");
   
   var neededCols = startCol + chunks.length - 1;
   if (neededCols > sheet.getMaxColumns()) {
     sheet.insertColumnsAfter(sheet.getMaxColumns(), neededCols - sheet.getMaxColumns());
   }
   
   sheet.getRange(rowNum, startCol, 1, chunks.length).setValues([chunks]);
}

// 4. MENGAMBIL SELURUH DATA (ANTI-NULL / ANTI ERROR DATE SPREADSHEET)
function getInitialData() {
  initDatabase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = { settings: {}, pegawai: [], jurnal: [] };

  // Tarik Data Setup (Memaksa jadi text menggunakan getDisplayValues)
  var sheetSetup = ss.getSheetByName("SetUpDatabase");
  if(sheetSetup && sheetSetup.getLastRow() > 0) {
      var sData = sheetSetup.getDataRange().getDisplayValues();
      for (var i = 1; i < sData.length; i++) {
        data.settings[sData[i][0]] = readChunkedRow(sData[i], 1);
      }
  }

  // Tarik Data Pegawai (Memaksa jadi text menggunakan getDisplayValues)
  var pSheet = ss.getSheetByName("Pegawai");
  if (pSheet && pSheet.getLastRow() > 1) {
    var pData = pSheet.getDataRange().getDisplayValues();
    for (var i = 1; i < pData.length; i++) {
      data.pegawai.push({
        nip: pData[i][0], nama: pData[i][1], jabatan: pData[i][2],
        jenis: pData[i][3], email: pData[i][4], username: pData[i][5],
        password: pData[i][6], 
        ttdBase64: readChunkedRow(pData[i], 7)
      });
    }
  }

  // Tarik Data Jurnal (Memaksa jadi text menggunakan getDisplayValues)
  var jSheet = ss.getSheetByName("Jurnal");
  if (jSheet && jSheet.getLastRow() > 1) {
    var jData = jSheet.getDataRange().getDisplayValues();
    for (var i = 1; i < jData.length; i++) {
      var actString = readChunkedRow(jData[i], 4);
      var activities = [];
      try { activities = JSON.parse(actString); } catch(e) {}
      
      data.jurnal.push({
        id: jData[i][0], jenisJurnal: jData[i][1], tanggal: jData[i][2],
        pegawai: jData[i][3], activities: activities
      });
    }
  }
  
  return data;
}

// 5. MENYIMPAN PENGATURAN SEKOLAH PERMANEN
function saveSettingsToDatabase(settingsObj) {
  initDatabase();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SetUpDatabase");
  var data = sheet.getDataRange().getValues();
  
  settingsObj.lastUpdated = new Date().getTime().toString();

  for (var key in settingsObj) {
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        writeChunkedData(sheet, i + 1, 2, settingsObj[key]); 
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key]);
      writeChunkedData(sheet, sheet.getLastRow(), 2, settingsObj[key]);
    }
  }
  return true;
}

// 6. MENYIMPAN AKUN PEGAWAI PERMANEN
function savePegawai(peg) {
  initDatabase();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pegawai");
  var data = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === peg.nip.toString()) {
      sheet.getRange(i + 1, 1, 1, 7).setValues([[peg.nip, peg.nama, peg.jabatan, peg.jenis, peg.email, peg.username, peg.password]]);
      writeChunkedData(sheet, i + 1, 8, peg.ttdBase64);
      found = true; break;
    }
  }
  if (!found) {
     sheet.appendRow([peg.nip, peg.nama, peg.jabatan, peg.jenis, peg.email, peg.username, peg.password]);
     writeChunkedData(sheet, sheet.getLastRow(), 8, peg.ttdBase64);
  }
  return true;
}

// 7. MENYIMPAN JURNAL PERMANEN
function saveJurnalMultiple(jurnals) {
  initDatabase();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Jurnal");
  var data = sheet.getDataRange().getValues();
  
  for(var k=0; k<jurnals.length; k++) {
    var jur = jurnals[k];
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === jur.id.toString()) {
        sheet.getRange(i + 1, 1, 1, 4).setValues([[jur.id, jur.jenisJurnal, jur.tanggal, jur.pegawai]]);
        writeChunkedData(sheet, i + 1, 5, JSON.stringify(jur.activities));
        found = true; break;
      }
    }
    if (!found) {
        sheet.appendRow([jur.id, jur.jenisJurnal, jur.tanggal, jur.pegawai]);
        writeChunkedData(sheet, sheet.getLastRow(), 5, JSON.stringify(jur.activities));
    }
  }
  return true;
}

// 8. HAPUS JURNAL
function deleteJurnalData(ids) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Jurnal");
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (ids.indexOf(data[i][0].toString()) > -1) {
      sheet.deleteRow(i + 1);
    }
  }
  return true;
}

// 9. GANTI PASSWORD
function changeUserPassword(nip, oldPass, newPass) {
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pegawai");
   var data = sheet.getDataRange().getValues();
   for(var i=1; i<data.length; i++) {
      if(data[i][0].toString() === nip.toString()) {
         if(data[i][6].toString() === oldPass.toString()) {
            sheet.getRange(i+1, 7).setValue(newPass);
            return {success: true};
         } else {
            return {success: false, message: "Password lama salah!"};
         }
      }
   }
   return {success: false, message: "User tidak ditemukan!"};
}
