// =====================================================================
// APLIKASI PRESENSI QR CODE - Google Apps Script Backend
// File: Code.gs
// =====================================================================

const SHEET_NAME_USER = 'Data_User';
const SHEET_NAME_LOG  = 'Log_Presensi';
const SHEET_NAME_JOURNAL = 'Jurnal_Kegiatan';
const SHEET_NAME_STUDENT = 'Data_Siswa';
const SHEET_NAME_STUDENT_ATTENDANCE = 'Absensi_Siswa';
const JOURNAL_PHOTO_FOLDER_ID = '12VbH_aeHsBicV0anaqhXJJElHY6RZiuu';

function doGet(e) {
  if (e && e.parameter && e.parameter.api === 'dashboard') {
    const callback = String(e.parameter.callback || '');
    if (!/^[A-Za-z_$][0-9A-Za-z_$]{0,80}$/.test(callback)) {
      return ContentService.createTextOutput('Invalid callback.');
    }
    const date = String(e.parameter.date || '');
    const dashboard = getStudentAttendanceDashboard(date);
    return ContentService.createTextOutput(callback + '(' + JSON.stringify({
      total: dashboard.total,
      records: dashboard.records,
      stats: dashboard.stats,
      byTeacher: dashboard.byTeacher,
      byActivity: dashboard.byActivity
    }) + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Aplikasi Presensi QR Code')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getStudentAttendanceDashboard(date) {
  const targetDate = date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_STUDENT_ATTENDANCE);
  const empty = { total: 0, records: [], stats: { Hadir: 0, Izin: 0, Sakit: 0, Alpa: 0 }, byTeacher: [], byActivity: [] };
  if (!sheet || sheet.getLastRow() < 2) return empty;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues().filter(function(row) {
    return String(row[1]).slice(0, 10) === targetDate;
  });
  const stats = { Hadir: 0, Izin: 0, Sakit: 0, Alpa: 0 };
  const teachers = {}, activities = {};
  const records = rows.map(function(row) {
    const status = String(row[8] || '');
    if (stats[status] !== undefined) stats[status]++;
    const teacher = String(row[3] || '-');
    const activity = String(row[4] || '-');
    teachers[teacher] = (teachers[teacher] || 0) + (status === 'Hadir' ? 1 : 0);
    activities[activity] = (activities[activity] || 0) + (status === 'Hadir' ? 1 : 0);
    return { waktu: row[0], guru: teacher, kegiatan: activity, siswa: String(row[6] || '-'), kelas: String(row[7] || '-'), status: status };
  }).reverse();
  return {
    total: rows.length,
    records: records.slice(0, 100),
    stats: stats,
    byTeacher: Object.keys(teachers).map(function(name) { return { name: name, hadir: teachers[name] }; }),
    byActivity: Object.keys(activities).map(function(name) { return { name: name, hadir: activities[name] }; })
  };
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
    if (!users.length) {
      data.forEach(function(row) {
        if (row[0] && row[1]) users.push({ id: String(row[0]), nama: String(row[1]), jabatan: String(row[2] || 'Guru') });
      });
    }
    Logger.log('Daftar user/guru dikirim: ' + users.length);
    return users;
  } catch (error) {
    Logger.log('Error di getUserList: ' + error.toString());
    throw error;
  }
}

function getTeacherList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_USER);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .filter(function(row) { return row[0] && row[1]; })
    .map(function(row) { return { id: String(row[0]), nama: String(row[1]), jabatan: String(row[2] || 'Guru'), status: String(row[3] || '') }; });
}

function getStudentsByTeacher(teacherId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_STUDENT);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .filter(function(row) { return String(row[0]).trim() === String(teacherId).trim() && row[1] && row[2]; })
    .map(function(row) { return { guruId: String(row[0]), id: String(row[1]), nama: String(row[2]), kelas: String(row[3] || '-') }; });
}

function saveManualAttendance(data) {
  if (!data || !data.tanggal || !data.guruId || !data.kegiatan || !Array.isArray(data.records) || !data.records.length) {
    throw new Error('Guru, tanggal, kegiatan, dan daftar siswa wajib diisi.');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME_STUDENT_ATTENDANCE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_STUDENT_ATTENDANCE);
    sheet.appendRow(['Waktu Simpan', 'Tanggal', 'ID Guru', 'Nama Guru', 'Kegiatan', 'ID Siswa', 'Nama Siswa', 'Kelas', 'Status']);
  }
  const validStatuses = ['Hadir', 'Izin', 'Sakit', 'Alpa'];
  const rows = data.records.map(function(record) {
    if (!record.id || !record.nama || validStatuses.indexOf(record.status) === -1) throw new Error('Data status siswa tidak valid.');
    return [new Date(), data.tanggal, String(data.guruId), String(data.guruNama || ''), String(data.kegiatan).trim(), String(record.id), String(record.nama), String(record.kelas || '-'), record.status];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return { success: true, message: rows.length + ' data absensi berhasil disimpan.' };
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
    const userList = getTeacherList();
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

function saveJournal(journal) {
  if (!journal || !journal.tanggal || !journal.user || !journal.kegiatan || !journal.photoData) {
    throw new Error('Tanggal, nama, kegiatan, dan foto wajib diisi.');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME_JOURNAL);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_JOURNAL);
    sheet.appendRow(['Waktu Simpan', 'Tanggal Kegiatan', 'ID Peserta', 'Nama', 'Jabatan', 'Kegiatan', 'Keterangan', 'Foto']);
  }
  const imageData = journal.photoData.split(',')[1];
  const blob = Utilities.newBlob(Utilities.base64Decode(imageData), journal.photoType || 'image/jpeg', journal.photoName || 'foto.jpg');
  const parentFolder = DriveApp.getFolderById(JOURNAL_PHOTO_FOLDER_ID);
  const dateFolderName = String(journal.tanggal);
  const dateFolders = parentFolder.getFoldersByName(dateFolderName);
  const photoFolder = dateFolders.hasNext() ? dateFolders.next() : parentFolder.createFolder(dateFolderName);
  const photo = photoFolder.createFile(blob);
  sheet.appendRow([new Date(), journal.tanggal, journal.user.id || '', journal.user.nama || '', journal.user.jabatan || '', journal.kegiatan, journal.keterangan || '', photo.getUrl()]);
  return { success: true, message: 'Jurnal berhasil disimpan.' };
}
