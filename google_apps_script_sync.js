/**
 * ==============================================================================
 * GOOGLE APPS SCRIPT: SYNC STATUS & PEWARNAAN OTOMATIS REMARK SE2026
 * ==============================================================================
 * 
 * Script ini digunakan untuk menghubungkan bot Python (tambah_remark_27a.py)
 * dengan Google Spreadsheet secara live. Ketika bot selesai memproses baris data:
 * 1. Otomatis mencari tab data yang tepat (berdasarkan gid=0 / nama tab / kolom link_fasih).
 * 2. Otomatis menambahkan kolom "Status Pengerjaan" jika belum ada.
 * 3. Mengisi status "SELESAI (Nama Operator)" pada baris yang dikerjakan.
 * 4. Memberi warna hijau lembut (#d9ead3) pada seluruh baris yang selesai.
 * 5. Mendukung banyak bot/pengguna berjalan bersamaan (menggunakan ScriptLock).
 * 
 * ------------------------------------------------------------------------------
 * CARA MEMASANG (HANYA 1 MENIT):
 * ------------------------------------------------------------------------------
 * 1. Buka Google Spreadsheet target Anda:
 *    - Untuk Data detail anomali: https://docs.google.com/spreadsheets/d/12FWSyYak-scXS5skHFUSQbVkjy6Oh8ofejusqnOb5nA/edit?gid=0#gid=0
 *    - (Jika ada sheet kedua, lakukan langkah yang sama pada spreadsheet kedua).
 * 2. Di menu atas, klik: 'Ekstensi' (Extensions) -> 'Apps Script'.
 * 3. Hapus seluruh isi kode default yang ada di layar editor,
 *    lalu COPY-PASTE SELURUH kode file ini ke sana.
 * 4. Klik ikon Disket (Simpan / Ctrl+S). Beri nama proyek misalnya: "Sync Remark SE2026".
 * 5. Klik tombol biru 'Terapkan' (Deploy) di pojok kanan atas -> pilih 'Kelola deployment' jika mengupdate,
 *    atau 'Deployment baru' (New deployment) jika baru pertama kali.
 * 6. Di samping tulisan "Pilih jenis", klik ikon Roda Gigi (Gear) -> pilih 'Aplikasi web' (Web app).
 * 7. Isi form pengaturan berikut:
 *    - Deskripsi          : Sync Remark SE2026 v2
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
 * 11. Masukkan URL tersebut ke file 'sheet_sync_config.json' di folder bot.
 * ==============================================================================
 */

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var lock = null;
  var hasLock = false;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var allSheets = ss.getSheets();

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

    // 1. CARI TAB TARGET SECARA CERDAS & AKURAT
    var sheet = null;

    // A. Cari berdasarkan parameter gid (misal gid=0 atau gid=293925488)
    if (params.gid !== undefined && params.gid !== null && params.gid !== "") {
      var reqGid = parseInt(params.gid);
      for (var s = 0; s < allSheets.length; s++) {
        if (allSheets[s].getSheetId() === reqGid) {
          sheet = allSheets[s];
          break;
        }
      }
    }

    // B. Cari tab yang memiliki kolom 'link_fasih' atau 'link_fasih_sm'
    if (!sheet) {
      for (var s = 0; s < allSheets.length; s++) {
        var numCols = Math.min(allSheets[s].getLastColumn(), 35);
        if (numCols > 0) {
          var testHeaders = allSheets[s].getRange(1, 1, 1, numCols).getValues()[0];
          var hasLink = testHeaders.some(function(h) {
            var strH = String(h || "").toLowerCase();
            return strH.includes("link_fasih");
          });
          if (hasLink) {
            sheet = allSheets[s];
            break;
          }
        }
      }
    }

    // C. Cari berdasarkan kandidat nama sheet (case-insensitive)
    if (!sheet) {
      var candidates = ["data detail anomali", "dtata_ntb", "dtata ntb", "sheet1", "data anomali", "anomali"];
      for (var s = 0; s < allSheets.length; s++) {
        var sName = allSheets[s].getName().trim().toLowerCase();
        for (var c = 0; c < candidates.length; c++) {
          if (sName === candidates[c] || sName.includes(candidates[c])) {
            sheet = allSheets[s];
            break;
          }
        }
        if (sheet) break;
      }
    }

    if (!sheet) {
      sheet = ss.getActiveSheet();
    }

    // A. Ping / Test Endpoint (Read-only, tanpa lock)
    if (params.action === "ping" || params.action === "test") {
      return ContentService.createTextOutput(JSON.stringify({
        status: "ok",
        message: "Google Apps Script Sync SE2026 aktif! Terhubung ke spreadsheet: " + ss.getName(),
        sheet_name: sheet.getName(),
        sheet_gid: sheet.getSheetId(),
        total_rows: sheet.getLastRow(),
        total_cols: sheet.getLastColumn()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // B. Ambil Pemetaan Kolom Dinamis Berdasarkan Nama Header (Tahan Geser / Sisip Kolom)
    var colMap = getSheetColumnMap(sheet);
    var statusColIdx = colMap.statusColIdx;
    var linkColIdx = colMap.linkColIdx;
    var namaColIdx = colMap.namaColIdx;
    var desaColIdx = colMap.desaColIdx;

    // C. Endpoint Cek Status 1 Assignment (Read-only, Realtime & Super Cepat)
    if (params.action === "check_status") {
      var targetLink = (params.link || params.link_fasih_sm || params.link_fasih || "").trim();
      var targetNama = (params.nama_usaha || "").trim();
      var rowNum = parseInt(params.row_num || 0);

      if (statusColIdx === -1) {
        return ContentService.createTextOutput(JSON.stringify({
          status: "ok",
          is_done: false,
          sheet_status: ""
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var data = sheet.getDataRange().getValues();
      var targetRow = -1;

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

      if (targetRow === -1 && linkColIdx !== -1 && targetLink) {
        var cleanTarget = targetLink.split("?")[0].replace(/\/edit\/?$/, "").trim();
        for (var r = 1; r < data.length; r++) {
          var cLink = String(data[r][linkColIdx] || "").split("?")[0].replace(/\/edit\/?$/, "").trim();
          if (cLink && (cLink === cleanTarget || cLink.includes(cleanTarget) || cleanTarget.includes(cLink))) {
            targetRow = r + 1;
            break;
          }
        }
      }

      if (targetRow === -1 && targetNama && namaColIdx !== -1) {
        for (var rn = 1; rn < data.length; rn++) {
          if (String(data[rn][namaColIdx] || "").trim().toLowerCase() === targetNama.toLowerCase()) {
            targetRow = rn + 1;
            break;
          }
        }
      }

      if (targetRow > 1) {
        var currentStatus = String(data[targetRow - 1][statusColIdx - 1] || "").trim();
        var stUpper = currentStatus.toUpperCase();
        var isDone = (stUpper.includes("SELESAI") || stUpper.includes("DONE") || stUpper.includes("SUDAH") || stUpper.includes("BERHASIL") || stUpper.includes("DILEWATI"));
        return ContentService.createTextOutput(JSON.stringify({
          status: "ok",
          is_done: isDone,
          sheet_status: currentStatus,
          row: targetRow
        })).setMimeType(ContentService.MimeType.JSON);
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "ok",
        is_done: false,
        sheet_status: ""
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // D. Endpoint Ambil Seluruh Data yang Selesai (Read-only, Sinkronisasi Awal/Menu)
    if (params.action === "get_completed") {
      var completedList = [];
      if (statusColIdx !== -1) {
        var data = sheet.getDataRange().getValues();
        for (var r = 1; r < data.length; r++) {
          var st = String(data[r][statusColIdx - 1] || "").trim();
          var stUpper = st.toUpperCase();
          if (stUpper.includes("SELESAI") || stUpper.includes("DONE") || stUpper.includes("SUDAH") || stUpper.includes("BERHASIL") || stUpper.includes("DILEWATI")) {
            completedList.push({
              link: linkColIdx !== -1 ? String(data[r][linkColIdx] || "") : "",
              nama_usaha: namaColIdx !== -1 ? String(data[r][namaColIdx] || "") : "",
              iddesa: desaColIdx !== -1 ? String(data[r][desaColIdx] || "") : "",
              status: st,
              row: r + 1
            });
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        status: "ok",
        count: completedList.length,
        items: completedList
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ==========================================================================
    // OPERASI WRITE / UPDATE STATUS KE GOOGLE SHEETS (DILINDUNGI SCRIPT LOCK)
    // ==========================================================================
    lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "Server Google Sheets sedang sibuk melayani pembaruan lain, silakan coba lagi beberapa saat."
      })).setMimeType(ContentService.MimeType.JSON);
    }
    hasLock = true;

    var targetLink = (params.link || params.link_fasih_sm || params.link_fasih || "").trim();
    var targetNama = (params.nama_usaha || "").trim();
    var rowNum = parseInt(params.row_num || 0);
    var statusText = (params.status || "SELESAI").trim();
    var operator = (params.operator || "").trim();
    var timestamp = params.timestamp || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");

    // 2. Cari atau buat kolom 'Status Pengerjaan' secara dinamis (Tahan Geser / Sisip Kolom)
    var writeColMap = getSheetColumnMap(sheet);
    var statusColIdx = writeColMap.statusColIdx;
    var linkColIdx = writeColMap.linkColIdx;
    var namaColIdx = writeColMap.namaColIdx;
    var lastCol = sheet.getLastColumn();

    if (statusColIdx === -1) {
      statusColIdx = lastCol + 1;
      sheet.getRange(writeColMap.headerRowNum, statusColIdx).setValue("Status Pengerjaan");
      sheet.getRange(writeColMap.headerRowNum, statusColIdx).setFontWeight("bold");
      sheet.getRange(writeColMap.headerRowNum, statusColIdx).setBackground("#e0e0e0");
    }

    // 3. Cari baris target di spreadsheet
    var targetRow = -1;
    var data = sheet.getDataRange().getValues();

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
    if (targetRow === -1 && targetNama && namaColIdx !== -1) {
      for (var rn = 1; rn < data.length; rn++) {
        if (String(data[rn][namaColIdx] || "").trim().toLowerCase() === targetNama.toLowerCase()) {
          targetRow = rn + 1;
          break;
        }
      }
    }

    if (targetRow > 1) {
      // 1. Tulis status pengerjaan beserta nama operator jika ada
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
        sheet: sheet.getName(),
        row: targetRow,
        message: "Tab '" + sheet.getName() + "' baris " + targetRow + " berhasil diperbarui menjadi '" + finalStatus + "' dan diwarnai hijau."
      })).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({
        status: "not_found",
        sheet: sheet.getName(),
        message: "Data target tidak ditemukan pada tab '" + sheet.getName() + "'."
      })).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    if (hasLock && lock) {
      try {
        lock.releaseLock();
      } catch (e) {}
    }
  }
}

/**
 * Mendeteksi posisi kolom penting secara dinamis dari nama header sheet.
 * Menjamin kekebalan jika kolom digeser, disisipkan kolom baru, atau diubah namanya
 * (misal spasi, underscore, huruf besar/kecil) oleh pengguna lain.
 */
function getSheetColumnMap(sheet) {
  var lastCol = sheet.getLastColumn();
  var maxScanRows = Math.min(sheet.getLastRow(), 5);
  if (lastCol <= 0 || maxScanRows <= 0) {
    return { statusColIdx: -1, linkColIdx: -1, namaColIdx: -1, desaColIdx: -1, headerRowNum: 1 };
  }

  var headerRows = sheet.getRange(1, 1, maxScanRows, lastCol).getValues();
  var bestCols = { statusColIdx: -1, linkColIdx: -1, namaColIdx: -1, desaColIdx: -1, headerRowNum: 1 };

  for (var r = 0; r < headerRows.length; r++) {
    var rowVals = headerRows[r];
    var currentCols = { statusColIdx: -1, linkColIdx: -1, namaColIdx: -1, desaColIdx: -1, headerRowNum: r + 1 };

    for (var c = 0; c < rowVals.length; c++) {
      var raw = String(rowVals[c] || "").trim().toLowerCase();
      var clean = raw.replace(/[^a-z0-9]/g, "");
      if (!clean) continue;

      // Status Pengerjaan (status, statuspengerjaan, statusremark, statusselesai, statusbot, statuscleaning)
      if (clean.indexOf("status") !== -1) {
        currentCols.statusColIdx = c + 1; // 1-based index untuk getRange
      }
      // Link Fasih (linkfasih, fasihsm, linkfasihsm, fasih, link) tapi BUKAN maps / google
      if ((clean.indexOf("fasih") !== -1 || clean.indexOf("link") !== -1) &&
          clean.indexOf("map") === -1 && clean.indexOf("google") === -1) {
        currentCols.linkColIdx = c; // 0-based index untuk array data[][]
      }
      // Nama Usaha (namausaha, namaperusahaan, nama)
      if (clean.indexOf("nama") !== -1) {
        if (currentCols.namaColIdx === -1 || clean.indexOf("usaha") !== -1 || clean.indexOf("perusahaan") !== -1) {
          currentCols.namaColIdx = c; // 0-based
        }
      }
      // ID Desa (iddesa, kodedesa, kddesa, idkelurahan, desa)
      if (clean.indexOf("desa") !== -1) {
        currentCols.desaColIdx = c; // 0-based
      }
    }

    // Jika baris ini memiliki link atau nama usaha, ini adalah baris header
    if (currentCols.linkColIdx !== -1 || currentCols.namaColIdx !== -1) {
      bestCols = currentCols;
      break;
    }
  }

  return bestCols;
}
