/**
 * ==============================================================================
 * GOOGLE APPS SCRIPT: SYNC STATUS & PEWARNAAN OTOMATIS REMARK SE2026
 * ==============================================================================
 * 
 * Script ini digunakan untuk menghubungkan bot Python (tambah_remark_27a.py)
 * dengan Google Spreadsheet secara live. Ketika bot selesai memproses baris data:
 * 1. Otomatis menambahkan kolom "Status Pengerjaan" jika belum ada.
 * 2. Mengisi status "SELESAI" pada baris yang dikerjakan.
 * 3. Memberi warna hijau lembut (#d9ead3) pada seluruh baris yang selesai.
 * 4. Mendukung banyak bot/pengguna berjalan bersamaan (menggunakan ScriptLock).
 * 
 * ------------------------------------------------------------------------------
 * CARA MEMASANG (HANYA 1 MENIT):
 * ------------------------------------------------------------------------------
 * 1. Buka Google Spreadsheet target:
 *    https://docs.google.com/spreadsheets/d/12FWSyYak-scXS5skHFUSQbVkjy6Oh8ofejusqnOb5nA/edit?gid=0#gid=0
 * 2. Di menu atas, klik: 'Ekstensi' (Extensions) -> 'Apps Script'.
 * 3. Hapus seluruh isi kode default yang ada di layar editor,
 *    lalu COPY-PASTE SELURUH kode file ini ke sana.
 * 4. Klik ikon Disket (Simpan / Ctrl+S). Beri nama proyek misalnya: "Sync Remark SE2026".
 * 5. Klik tombol biru 'Terapkan' (Deploy) di pojok kanan atas -> pilih 'Deployment baru' (New deployment).
 * 6. Di samping tulisan "Pilih jenis", klik ikon Roda Gigi (Gear) -> pilih 'Aplikasi web' (Web app).
 * 7. Isi form pengaturan berikut:
 *    - Deskripsi          : Sync Remark SE2026
 *    - Jalankan sebagai   : 'Saya' (Me - email akun Google Anda)
 *    - Yang memiliki akses: 'Siapa saja' (Anyone)  <--- [PENTING!]
 * 8. Klik tombol 'Terapkan' (Deploy).
 * 9. Jika muncul pop-up izin otorisasi Google:
 *    - Klik 'Otorisasi akses' (Review permissions)
 *    - Pilih akun Google Anda
 *    - Jika muncul "Google hasn't verified this app", klik tulisan kecil 'Advanced' / 'Lanjutan' di kiri bawah
 *    - Klik 'Go to Sync Remark SE2026 (unsafe)'
 *    - Klik tombol 'Allow' / 'Izinkan'
 * 10. Salin "URL Aplikasi web" yang diberikan (berformat: https://script.google.com/macros/s/.../exec).
 * 11. Masukkan URL tersebut ke file 'sheet_sync_config.json' di folder bot ini, atau masukkan saat bot dijalankan.
 * ==============================================================================
 */

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var lock = LockService.getScriptLock();
  // Kunci script selama max 30 detik agar antrean dari banyak PC aman tanpa tabrakan data
  if (!lock.tryLock(30000)) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Server Google Sheets sedang sibuk, silakan coba lagi beberapa saat."
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // Cari sheet target: 'Data detail anomali', atau 'Dtata_NTB', atau sheet pertama aktif
    var sheet = ss.getSheetByName("Data detail anomali") || 
                ss.getSheetByName("Dtata_NTB") || 
                ss.getSheetByName("Sheet1") || 
                ss.getActiveSheet();

    var params = {};
    if (e && e.postData && e.postData.contents) {
      try {
        params = JSON.parse(e.postData.contents);
      } catch (err) {
        params = e.parameter || {};
      }
    } else if (e && e.parameter) {
      params = e.parameter;
    }

    // Ping / Test Endpoint
    if (params.action === "ping" || params.action === "test") {
      return ContentService.createTextOutput(JSON.stringify({
        status: "ok",
        message: "Google Apps Script Sync SE2026 aktif! Terhubung ke spreadsheet: " + ss.getName(),
        sheet_name: sheet.getName(),
        total_rows: sheet.getLastRow(),
        total_cols: sheet.getLastColumn()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var targetLink = (params.link || params.link_fasih_sm || params.link_fasih || "").trim();
    var targetNama = (params.nama_usaha || "").trim();
    var rowNum = parseInt(params.row_num || 0);
    var statusText = (params.status || "SELESAI").trim();
    var operator = (params.operator || "").trim();
    var timestamp = params.timestamp || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");

    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // 1. Cari atau buat kolom 'Status Pengerjaan'
    var statusColIdx = -1;
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i] || "").trim().toLowerCase();
      if (h === "status pengerjaan" || h === "status remark" || h === "status selesai") {
        statusColIdx = i + 1;
        break;
      }
    }

    if (statusColIdx === -1) {
      statusColIdx = lastCol + 1;
      sheet.getRange(1, statusColIdx).setValue("Status Pengerjaan");
      sheet.getRange(1, statusColIdx).setFontWeight("bold");
      sheet.getRange(1, statusColIdx).setBackground("#e0e0e0");
    }

    // 2. Cari baris target di spreadsheet
    var targetRow = -1;
    var data = sheet.getDataRange().getValues();

    // Deteksi index kolom link_fasih
    var linkColIdx = -1;
    for (var j = 0; j < headers.length; j++) {
      var hj = String(headers[j] || "").toLowerCase();
      if (hj.includes("link_fasih") || hj === "link") {
        linkColIdx = j;
        break;
      }
    }

    // A. Cek kecocokan berdasarkan row_num (cepat jika sheet belum disortir ulang)
    if (rowNum > 1 && rowNum <= data.length) {
      if (linkColIdx !== -1 && targetLink) {
        var cellLink = String(data[rowNum - 1][linkColIdx] || "").trim();
        if (cellLink && (cellLink.includes(targetLink) || targetLink.includes(cellLink))) {
          targetRow = rowNum;
        }
      } else {
        targetRow = rowNum;
      }
    }

    // B. Jika row_num tidak cocok, cari scan seluruh baris berdasarkan link Fasih-SM
    if (targetRow === -1 && linkColIdx !== -1 && targetLink) {
      // Potong link bersih tanpa /edit atau query parameter untuk pencocokan akurat
      var cleanTarget = targetLink.split("?")[0].replace(/\/edit\/?$/, "").trim();
      for (var r = 1; r < data.length; r++) {
        var cLink = String(data[r][linkColIdx] || "").split("?")[0].replace(/\/edit\/?$/, "").trim();
        if (cLink && (cLink === cleanTarget || cLink.includes(cleanTarget) || cleanTarget.includes(cLink))) {
          targetRow = r + 1;
          break;
        }
      }
    }

    // C. Jika belum ketemu juga, cari berdasarkan nama usaha
    if (targetRow === -1 && targetNama) {
      var namaColIdx = -1;
      for (var k = 0; k < headers.length; k++) {
        if (String(headers[k] || "").toLowerCase().includes("nama_usaha")) {
          namaColIdx = k;
          break;
        }
      }
      if (namaColIdx !== -1) {
        for (var rn = 1; rn < data.length; rn++) {
          if (String(data[rn][namaColIdx] || "").trim().toLowerCase() === targetNama.toLowerCase()) {
            targetRow = rn + 1;
            break;
          }
        }
      }
    }

    if (targetRow > 1) {
      // 1. Tulis status pengerjaan
      var finalStatus = statusText;
      if (operator) {
        finalStatus += " (" + operator + ")";
      }
      sheet.getRange(targetRow, statusColIdx).setValue(finalStatus);

      // 2. Beri warna latar hijau lembut (#d9ead3) pada seluruh baris data tersebut
      var totalCols = Math.max(sheet.getLastColumn(), statusColIdx);
      sheet.getRange(targetRow, 1, 1, totalCols).setBackground("#d9ead3");

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        row: targetRow,
        message: "Baris " + targetRow + " berhasil diperbarui menjadi '" + finalStatus + "' dan diwarnai hijau."
      })).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({
        status: "not_found",
        message: "Data target tidak ditemukan di spreadsheet."
      })).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
