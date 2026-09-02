// =====================================================================
// APLIKASI PRESENSI QR CODE - Google Apps Script Backend
// File: Code.gs
// =====================================================================

const SHEET_NAME_USER = 'Data_User';
const SHEET_NAME_LOG  = 'Log_Presensi';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Aplikasi Presensi QR Code')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Endpoint untuk scanner eksternal. Responsnya tidak dibaca oleh browser
// eksternal (mode no-cors), namun validasi dan pencatatan tetap dilakukan server.
function doPost(e) {
  let request;
  try {
    request = JSON.parse(e.postData.contents);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Format request tidak valid.'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  const result = request.action === 'recordPresensi'
    ? recordPresensi(request.payload)
    : { success: false, message: 'Aksi tidak dikenali.' };
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getUserList() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME_USER);
    if (!sheet) throw new Error("Sheet '" + SHEET_NAME_USER + "' tidak ditemukan!");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const users = [];
    data.forEach(function(row) {
      const id = row[0]; const nama = row[1]; const jabatan = row[2]; const statusAktif = row[3];
      if (id && nama && String(statusAktif).trim().toLowerCase() === 'aktif') {
        users.push({ id: id, nama: nama, jabatan: jabatan });
      }
    });
    return users;
  } catch (error) {
    Logger.log('Error di getUserList: ' + error.toString());
    throw error;
  }
}

function getTodayPresensi() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_LOG);
    if (!sheet || sheet.getLastRow() < 2) return [];

    const timeZone = Session.getScriptTimeZone();
    const today = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues()
      .filter(function(row) {
        return Utilities.formatDate(new Date(row[0]), timeZone, 'yyyy-MM-dd') === today;
      })
      .reverse()
      .slice(0, 20)
      .map(function(row) {
        return {
          name: String(row[2] || '-'),
          ok: String(row[3]).toLowerCase() === 'hadir',
          desc: String(row[3] || 'Presensi'),
          time: Utilities.formatDate(new Date(row[0]), timeZone, 'HH:mm')
        };
      });
  } catch (error) {
    Logger.log('Error di getTodayPresensi: ' + error.toString());
    return [];
  }
}

function recordPresensi(qrPayload) {
  try {
    let data;
    try { data = JSON.parse(qrPayload); } catch (e) {
      return { success: false, message: 'QR Code tidak valid atau tidak dapat dibaca.' };
    }
    const { id, nama, jabatan, token, timestamp } = data;
    if (!id || !nama || !token || !timestamp) {
      return { success: false, message: 'Data QR Code tidak lengkap.' };
    }
    const now = new Date().getTime();
    const qrTime = new Date(timestamp).getTime();
    const validityMs = 5 * 60 * 1000;
    if (isNaN(qrTime)) return { success: false, message: 'Timestamp pada QR Code tidak valid.' };
    if ((now - qrTime) > validityMs) return { success: false, message: 'QR Code sudah kedaluwarsa. Minta user untuk generate ulang.' };
    const userList = getUserList();
    const userValid = userList.some(function(u) { return u.id === id; });
    if (!userValid) return { success: false, message: "ID User '" + id + "' tidak ditemukan atau tidak aktif." };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(SHEET_NAME_LOG);
    if (!logSheet) throw new Error("Sheet '" + SHEET_NAME_LOG + "' tidak ditemukan!");
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const logData = logSheet.getLastRow() > 1
      ? logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 2).getValues() : [];
    const sudahHadir = logData.some(function(row) {
      const logDate = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      return row[1] === id && logDate === today;
    });
    if (sudahHadir) return { success: false, message: nama + ' sudah tercatat hadir hari ini.' };
    const scanTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    logSheet.appendRow([scanTime, id, nama, 'Hadir', 'Scan QR Admin']);
    Logger.log('Presensi berhasil: ' + id + ' - ' + nama + ' pada ' + scanTime);
    return { success: true, message: 'Presensi berhasil dicatat!', data: { nama: nama, jabatan: jabatan || '-', waktu: scanTime, status: 'Hadir' } };
  } catch (error) {
    Logger.log('Error di recordPresensi: ' + error.toString());
    return { success: false, message: 'Terjadi kesalahan server: ' + error.toString() };
  }
}
