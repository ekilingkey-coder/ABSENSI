/**
 * =========================================================================
 * GOOGLE APPS SCRIPT - SIPGN BGN (BACKEND SERVER & DATABASE RELAWAN MBG)
 * =========================================================================
 * Endpoint API Webhook:
 * 1. Menerima presensi absensi relawan & simpan foto ke Google Drive (per bulan)
 * 2. Sinkronisasi data relawan (Ambil / Simpan / Edit data relawan)
 * 3. Sinkronisasi data divisi
 * 4. Verifikasi login opsional sisi server (action: "verifyLogin")
 * =========================================================================
 */

var TIMEZONE = "Asia/Jakarta";
var DEFAULT_HONOR = 100000;

function doPost(e) {
  var lock = LockService.getScriptLock();
  var hasLock = lock.tryLock(15000);

  if (!hasLock) {
    return createJsonResponse({
      status: "error",
      message: "Server sedang sibuk memproses antrean lain. Silakan coba beberapa detik lagi."
    });
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({
        status: "error",
        message: "Payload request kosong atau tidak valid."
      });
    }

    var data = JSON.parse(e.postData.contents);

    // 1. ACTION: Sinkronisasi daftar relawan & divisi dari Admin
    if (data.action === "syncVolunteers") {
      if (!Array.isArray(data.volunteers)) {
        return createJsonResponse({
          status: "error",
          message: "Data 'volunteers' harus berupa array. Sinkronisasi dibatalkan demi keamanan data."
        });
      }

      var result = saveVolunteersToSheet(data.volunteers, data.divisions);
      return createJsonResponse({
        status: "success",
        message: "Daftar relawan berhasil disinkronkan ke Google Sheets",
        count: result.count
      });
    }

    // 2. ACTION OPSIONAL: Verifikasi login relawan di sisi server
    if (data.action === "verifyLogin") {
      return handleVerifyLogin(data.id, data.name, data.password);
    }

    // 3. DEFAULT ACTION: Simpan Presensi Relawan
    if (!data.name || String(data.name).trim() === "") {
      return createJsonResponse({
        status: "error",
        message: "Nama relawan wajib diisi."
      });
    }

    var sheet = getOrCreateAttendanceSheet();
    var attendanceId = String(data.id || ("AT-" + Date.now()));

    // Cegah duplikasi jika frontend melakukan retry dengan ID presensi yang sama
    if (data.id && isAttendanceIdExists(sheet, attendanceId)) {
      return createJsonResponse({
        status: "success",
        duplicate: true,
        message: "Presensi dengan ID " + attendanceId + " sudah tercatat sebelumnya."
      });
    }

    // Simpan foto ke Google Drive jika ada swafoto base64
    var photoUrl = "-";
    var photoWarning = null;
    if (data.photo && typeof data.photo === "string" && data.photo.indexOf("data:image") !== -1) {
      var photoResult = savePhotoToDrive(data.photo, data.name, data.type);
      photoUrl = photoResult.url;
      if (!photoResult.success) {
        photoWarning = photoResult.error;
      }
    }

    var honorValue = parseHonor(data.honor, DEFAULT_HONOR);
    var rowData = [
      new Date(),                                   // A: Waktu Server Masuk
      attendanceId,                                 // B: ID Presensi
      String(data.reportDate || "-"),               // C: Tanggal Laporan Operasional
      String(data.name || "").trim(),               // D: Nama Relawan
      String(data.role || "").trim(),               // E: Divisi
      String(data.type || "MASUK").trim(),          // F: Tipe Presensi (MASUK/PULANG)
      String(data.gps || ""),                       // G: Titik Koordinat GPS
      String(data.taskNote || ""),                  // H: Catatan Tugas / Menu
      photoUrl,                                     // I: Link Foto Bukti di Google Drive
      honorValue,                                   // J: Standar Honor
      String(data.status || "Valid (GPS & Selfie)") // K: Status SIPGN
    ];

    // Tulis dengan format kolom yang tepat agar tanggal/GPS/ID tidak terkorupsi & bebas formula injection
    var nextRow = sheet.getLastRow() + 1;
    var rowRange = sheet.getRange(nextRow, 1, 1, rowData.length);

    sheet.getRange(nextRow, 1, 1, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    sheet.getRange(nextRow, 2, 1, 8).setNumberFormat("@"); // Kolom B..I Plain Text
    sheet.getRange(nextRow, 10, 1, 1).setNumberFormat("#,##0"); // Kolom J Honor
    sheet.getRange(nextRow, 11, 1, 1).setNumberFormat("@"); // Kolom K Status

    rowRange.setValues([rowData]);

    var responsePayload = {
      status: "success",
      message: "Data presensi berhasil disimpan ke Google Sheets & Drive BGN",
      id: attendanceId,
      driveUrl: photoUrl
    };
    if (photoWarning) {
      responsePayload.photoWarning = photoWarning;
    }

    return createJsonResponse(responsePayload);

  } catch (error) {
    return createJsonResponse({
      status: "error",
      message: error.toString()
    });
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var action = params.action || "getVolunteers";

    // 1. ACTION: Ambil daftar relawan & divisi
    if (action === "getVolunteers") {
      var volunteerSheet = getOrCreateVolunteerSheet();
      var rows = volunteerSheet.getDataRange().getValues();
      var volunteers = [];

      for (var i = 1; i < rows.length; i++) {
        var row = rows[i];
        if (row[0] !== "" || row[1] !== "") {
          volunteers.push({
            id: String(row[0] || ("V" + i)),
            name: String(row[1] || ""),
            role: String(row[2] || "PERSIAPAN"),
            honor: parseHonor(row[3], DEFAULT_HONOR),
            password: String(row[4] !== "" && row[4] !== null ? row[4] : "1234")
          });
        }
      }

      var divisions = getDivisionsFromSheet();

      return createJsonResponse({
        status: "success",
        totalVolunteers: volunteers.length,
        volunteers: volunteers,
        divisions: divisions
      });
    }

    // 2. ACTION: Ambil riwayat presensi (mendukung filter opsional ?reportDate=...&limit=...)
    if (action === "getAttendances") {
      var sheet = getOrCreateAttendanceSheet();
      var aRows = sheet.getDataRange().getValues();
      if (aRows.length <= 1) {
        return createJsonResponse({ status: "success", total: 0, data: [] });
      }

      var headers = aRows[0];
      var filterDate = params.reportDate ? String(params.reportDate).trim() : null;
      var limit = params.limit ? parseInt(params.limit, 10) : 0;
      var list = [];

      for (var j = 1; j < aRows.length; j++) {
        var aRow = aRows[j];
        // Lewati baris yang benar-benar kosong
        if (aRow[0] === "" && aRow[1] === "" && aRow[3] === "") continue;

        var item = {};
        for (var k = 0; k < headers.length; k++) {
          var headerName = headers[k];
          if (!headerName) continue;

          var cellVal = aRow[k];
          if (cellVal instanceof Date) {
            // Format sesuai zona waktu WIB agar tidak terkonversi ke UTC
            cellVal = Utilities.formatDate(cellVal, TIMEZONE, "yyyy-MM-dd HH:mm:ss");
          }
          item[headerName] = cellVal;
        }

        if (filterDate && String(item["Tanggal Laporan Operasional"]) !== filterDate) {
          continue;
        }

        list.push(item);
      }

      // Jika ada parameter limit, ambil N data terbaru
      if (limit > 0 && list.length > limit) {
        list = list.slice(list.length - limit);
      }

      return createJsonResponse({
        status: "success",
        total: list.length,
        data: list
      });
    }

    // Default test response
    return createJsonResponse({
      status: "success",
      message: "SIPGN API Server Aktif & Terhubung!"
    });

  } catch (err) {
    return createJsonResponse({
      status: "error",
      message: err.toString()
    });
  }
}

/**
 * Verifikasi login di sisi server (opsional untuk peningkatan keamanan frontend)
 */
function handleVerifyLogin(id, name, password) {
  var sheet = getOrCreateVolunteerSheet();
  var rows = sheet.getDataRange().getValues();
  var inputPin = String(password || "");

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var rowId = String(row[0] || "");
    var rowName = String(row[1] || "");
    var rowPin = String(row[4] !== "" && row[4] !== null ? row[4] : "1234");

    var isMatch = (id && rowId === String(id)) || (name && rowName.toLowerCase() === String(name).toLowerCase());
    if (isMatch) {
      if (rowPin === inputPin) {
        return createJsonResponse({
          status: "success",
          volunteer: {
            id: rowId,
            name: rowName,
            role: String(row[2] || "PERSIAPAN"),
            honor: parseHonor(row[3], DEFAULT_HONOR)
          }
        });
      } else {
        return createJsonResponse({ status: "error", message: "PIN / Password salah." });
      }
    }
  }
  return createJsonResponse({ status: "error", message: "Data relawan tidak ditemukan." });
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseHonor(val, fallback) {
  if (val === undefined || val === null || val === "") return fallback;
  var num = Number(val);
  return isNaN(num) ? fallback : num;
}

function formatCodeId(index) {
  var numStr = String(index + 1);
  while (numStr.length < 3) numStr = "0" + numStr;
  return "V" + numStr;
}

function isAttendanceIdExists(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;
  var ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) return true;
  }
  return false;
}

function getActiveSpreadsheetSafe() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Spreadsheet aktif tidak ditemukan. Pastikan Apps Script dibuat dari menu Extensions -> Apps Script di dalam Google Sheets.");
  }
  return ss;
}

/**
 * Membuat atau mengambil sheet "Data_Relawan"
 */
function getOrCreateVolunteerSheet() {
  var ss = getActiveSpreadsheetSafe();
  var sheet = ss.getSheetByName("Data_Relawan");

  if (!sheet) {
    sheet = ss.insertSheet("Data_Relawan");
    var headers = ["ID Relawan", "Nama Lengkap", "Divisi", "Honor Harian (Rp)", "PIN / Password"];
    sheet.appendRow(headers);

    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#0D2F72");
    headerRange.setFontColor("#FFFFFF");
    headerRange.setFontWeight("bold");
    headerRange.setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Menyimpan seluruh daftar relawan & divisi ke Google Sheets
 */
function saveVolunteersToSheet(volunteers, divisions) {
  var sheet = getOrCreateVolunteerSheet();

  // Bersihkan isi baris lama (aman terhadap frozen row)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 5)).clearContent();
  }

  if (volunteers && volunteers.length > 0) {
    var rowsToInsert = [];
    for (var i = 0; i < volunteers.length; i++) {
      var v = volunteers[i] || {};
      rowsToInsert.push([
        String(v.id || formatCodeId(i)),
        String(v.name || "").trim(),
        String(v.role || "").trim(),
        parseHonor(v.honor, DEFAULT_HONOR),
        String(v.password !== undefined && v.password !== null && v.password !== "" ? v.password : "1234")
      ]);
    }

    // Format teks sebelum menulis agar PIN dengan nol di depan (misal "0123") tidak hilang
    sheet.getRange(2, 1, rowsToInsert.length, 3).setNumberFormat("@");
    sheet.getRange(2, 4, rowsToInsert.length, 1).setNumberFormat("#,##0");
    sheet.getRange(2, 5, rowsToInsert.length, 1).setNumberFormat("@");

    sheet.getRange(2, 1, rowsToInsert.length, 5).setValues(rowsToInsert);
  }

  if (Array.isArray(divisions)) {
    saveDivisionsToSheet(divisions);
  }

  SpreadsheetApp.flush();
  return { count: (volunteers ? volunteers.length : 0) };
}

/**
 * Sheet untuk daftar Divisi
 */
function getOrCreateDivisionSheet() {
  var ss = getActiveSpreadsheetSafe();
  var sheet = ss.getSheetByName("Daftar_Divisi");
  if (!sheet) {
    sheet = ss.insertSheet("Daftar_Divisi");
    sheet.appendRow(["Nama Divisi"]);
    var hr = sheet.getRange(1, 1, 1, 1);
    hr.setBackground("#F59E0B");
    hr.setFontColor("#000000");
    hr.setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function saveDivisionsToSheet(divisions) {
  var sheet = getOrCreateDivisionSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 1)).clearContent();
  }

  var rows = [];
  for (var i = 0; i < divisions.length; i++) {
    if (divisions[i]) {
      rows.push([String(divisions[i]).trim()]);
    }
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat("@");
    sheet.getRange(2, 1, rows.length, 1).setValues(rows);
  }
}

function getDivisionsFromSheet() {
  var ss = getActiveSpreadsheetSafe();
  var sheet = ss.getSheetByName("Daftar_Divisi");
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  var divs = [];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0]) divs.push(String(rows[i][0]));
  }
  return divs;
}

/**
 * Membuat atau mengambil sheet "Presensi_Relawan" lengkap dengan Header rapi
 */
function getOrCreateAttendanceSheet() {
  var ss = getActiveSpreadsheetSafe();
  var sheet = ss.getSheetByName("Presensi_Relawan");

  if (!sheet) {
    sheet = ss.insertSheet("Presensi_Relawan");
    var headers = [
      "Waktu Server",
      "ID Presensi",
      "Tanggal Laporan Operasional",
      "Nama Relawan",
      "Divisi",
      "Jenis Presensi",
      "Koordinat GPS",
      "Aktivitas / Catatan",
      "Link Swafoto Drive",
      "Honor (Rp)",
      "Status Verifikasi"
    ];
    sheet.appendRow(headers);

    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#1B365D");
    headerRange.setFontColor("#FFFFFF");
    headerRange.setFontWeight("bold");
    headerRange.setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Mengambil atau membuat folder di Google Drive
 */
function getOrCreateFolder(parentFolder, folderName, publicView) {
  var folders = parentFolder
    ? parentFolder.getFoldersByName(folderName)
    : DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  var folder = parentFolder
    ? parentFolder.createFolder(folderName)
    : DriveApp.createFolder(folderName);

  if (publicView) {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  return folder;
}

/**
 * Menyimpan Base64 Image ke Folder "Foto_Presensi_Relawan_MBG/YYYY-MM" di Google Drive
 */
function savePhotoToDrive(base64Data, name, type) {
  try {
    var match = base64Data.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,([\s\S]+)$/);
    if (!match) {
      return { success: false, url: "Format foto tidak valid", error: "Base64 header tidak valid" };
    }

    var contentType = match[1];
    var base64Content = match[2];
    var ext = "jpg";
    if (contentType.indexOf("png") !== -1) ext = "png";
    else if (contentType.indexOf("webp") !== -1) ext = "webp";

    var now = new Date();
    var monthFolderName = Utilities.formatDate(now, TIMEZONE, "yyyy-MM");
    var timestamp = Utilities.formatDate(now, TIMEZONE, "yyyyMMdd_HHmmss");

    var rootFolder = getOrCreateFolder(null, "Foto_Presensi_Relawan_MBG", true);
    var monthFolder = getOrCreateFolder(rootFolder, monthFolderName, true);

    var safeType = String(type || "PRESENSI").replace(/[^a-zA-Z0-9_-]/g, "_");
    var safeName = String(name || "Relawan").replace(/[^a-zA-Z0-9_-]/g, "_").replace(/_+/g, "_");
    var fileName = safeType + "_" + safeName + "_" + timestamp + "." + ext;

    var bytes = Utilities.base64Decode(base64Content);
    var blob = Utilities.newBlob(bytes, contentType, fileName);

    var file = monthFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { success: true, url: file.getUrl() };
  } catch (err) {
    return { success: false, url: "Gagal simpan foto: " + err.toString(), error: err.toString() };
  }
}