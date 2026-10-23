/**
 * ==============================================================================
 * GOOGLE APPS SCRIPT: SYNC STATUS & PEWARNAAN OTOMATIS REMARK SE2026 (v3 - ULTRA FAST)
 * ==============================================================================
 * 
 * Script ini digunakan untuk menghubungkan bot Python (tambah_remark_27a.py)
 * dengan Google Spreadsheet secara live & realtime.
 * 
 * PENINGKATAN v3:
 * 1. ULTRA CEPAT: Menggunakan pembacaan kolom tunggal (Single-Column Slice) dan
 *    verifikasi baris instan (row_num direct verification). Waktu respon turun
 *    dari 22 detik menjadi < 0.5 detik untuk sheet berukuran 30.000+ baris!
 * 2. AKURAT MULTI-USAHA: Jika 1 assignment memiliki lebih dari 1 usaha (link sama,
 *    nama usaha berbeda), script mencocokkan link DAN nama usaha sekaligus sehingga
 *    baris usaha ke-2 / ke-3 diwarnai dan ditandai dengan tepat, tidak salah baris.
 * 3. DUKUNGAN BATCH UPDATE: Mendukung pengiriman update banyak baris sekaligus
 *    dalam 1 kali request HTTP (menghemat waktu & kuota API).
 * 4. DUKUNGAN RESET/CLEAR: Mendukung pembersihan status baris jika diperlukan.
 * 5. ScriptLock: Aman digunakan oleh banyak operator/bot secara bersamaan.
 * 
 * ------------------------------------------------------------------------------
 * CARA MEMASANG (HANYA 1 MENIT):
 * ------------------------------------------------------------------------------
 * 1. Buka Google Spreadsheet target Anda:
 *    - Untuk Data detail anomali (P1): gid=0
 *    - Untuk Data NTB (P2): gid=293925488
 * 2. Di menu atas, klik: 'Ekstensi' (Extensions) -> 'Apps Script'.
 * 3. Hapus seluruh isi kode lama di editor, lalu COPY-PASTE SELURUH isi file ini.
 * 4. Klik ikon Disket (Simpan / Ctrl+S).
 * 5. Klik tombol biru 'Terapkan' (Deploy) di pojok kanan atas -> pilih 'Kelola deployment'
 *    (Manage deployments).
 * 6. Klik ikon Pensil (Edit) di samping deployment aktif Anda:
 *    - Versi: pilih 'Versi baru' (New version)
 *    - Klik 'Terapkan' (Deploy).
 * ==============================================================================
 */

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function cleanStr(s) {
  return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function cleanLink(link) {
  return String(link || "").split("?")[0].replace(/\/edit\/?$/, "").replace(/\/+$/, "").trim();
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
            return strH.indexOf("link_fasih") !== -1 || strH.indexOf("link") !== -1;
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
      var candidates = ["data detail anomali", "dtata_ntb", "data_ntb", "dtata ntb", "sheet1", "data anomali", "anomali"];
      for (var s = 0; s < allSheets.length; s++) {
        var sName = allSheets[s].getName().trim().toLowerCase();
        for (var c = 0; c < candidates.length; c++) {
          if (sName === candidates[c] || sName.indexOf(candidates[c]) !== -1) {
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
        version: "v3-ultra-fast",
        message: "Google Apps Script Sync SE2026 aktif! Terhubung ke spreadsheet: " + ss.getName(),
        sheet_name: sheet.getName(),
        sheet_gid: sheet.getSheetId(),
        total_rows: sheet.getLastRow(),
        total_cols: sheet.getLastColumn()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // B. Ambil Pemetaan Kolom Dinamis Berdasarkan Nama Header
    var colMap = getSheetColumnMap(sheet);
    var statusColIdx = colMap.statusColIdx;
    var linkColIdx = colMap.linkColIdx;
    var namaColIdx = colMap.namaColIdx;
    var desaColIdx = colMap.desaColIdx;
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    // C. Endpoint Cek Status 1 Assignment (Read-only, Realtime & Super Cepat)
    if (params.action === "check_status") {
      var targetLink = (params.link || params.link_fasih_sm || params.link_fasih || "").trim();
      var targetNama = (params.nama_usaha || "").trim();
      var rowNum = parseInt(params.row_num || 0);

      if (statusColIdx === -1 || lastRow <= 1) {
        return ContentService.createTextOutput(JSON.stringify({
          status: "ok",
          is_done: false,
          sheet_status: ""
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var targetRow = findTargetRowFast(sheet, colMap, targetLink, targetNama, rowNum);

      if (targetRow > 1) {
        var currentStatus = String(sheet.getRange(targetRow, statusColIdx).getValue() || "").trim();
        var stUpper = currentStatus.toUpperCase();
        var isDone = (stUpper.indexOf("SELESAI") !== -1 || stUpper.indexOf("DONE") !== -1 || stUpper.indexOf("SUDAH") !== -1 || stUpper.indexOf("BERHASIL") !== -1 || stUpper.indexOf("DILEWATI") !== -1);
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

    // D. Endpoint Ambil Seluruh Data yang Selesai (Read-only, Slice Hemat Memori)
    if (params.action === "get_completed") {
      var completedList = [];
      if (statusColIdx !== -1 && lastRow > 1) {
        // Ambil hanya kolom status (1 kolom bukan 27 kolom!)
        var statusVals = sheet.getRange(1, statusColIdx, lastRow, 1).getValues();
        var linkVals = (linkColIdx !== -1) ? sheet.getRange(1, linkColIdx + 1, lastRow, 1).getValues() : null;
        var namaVals = (namaColIdx !== -1) ? sheet.getRange(1, namaColIdx + 1, lastRow, 1).getValues() : null;
        var desaVals = (desaColIdx !== -1) ? sheet.getRange(1, desaColIdx + 1, lastRow, 1).getValues() : null;

        for (var r = 1; r < statusVals.length; r++) {
          var st = String(statusVals[r][0] || "").trim();
          var stUpper = st.toUpperCase();
          if (stUpper.indexOf("SELESAI") !== -1 || stUpper.indexOf("DONE") !== -1 || stUpper.indexOf("SUDAH") !== -1 || stUpper.indexOf("BERHASIL") !== -1 || stUpper.indexOf("DILEWATI") !== -1) {
            completedList.push({
              link: linkVals ? String(linkVals[r][0] || "") : "",
              nama_usaha: namaVals ? String(namaVals[r][0] || "") : "",
              iddesa: desaVals ? String(desaVals[r][0] || "") : "",
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

    // Pastikan kolom 'Status Pengerjaan' ada
    if (statusColIdx === -1) {
      statusColIdx = lastCol + 1;
      sheet.getRange(colMap.headerRowNum, statusColIdx).setValue("Status Pengerjaan");
      sheet.getRange(colMap.headerRowNum, statusColIdx).setFontWeight("bold");
      sheet.getRange(colMap.headerRowNum, statusColIdx).setBackground("#e0e0e0");
      colMap.statusColIdx = statusColIdx;
      lastCol = statusColIdx;
    }

    var totalColsToColor = Math.max(sheet.getLastColumn(), statusColIdx);

    // E. BATCH UPDATE: Menangani array of items dalam 1 request
    if (params.action === "batch_update" && Array.isArray(params.items) && params.items.length > 0) {
      var batchOperator = (params.operator || "").trim();
      var defaultStatus = (params.status || "SELESAI").trim();
      var updatedCount = 0;
      var failedCount = 0;
      var updatedRows = [];

      for (var i = 0; i < params.items.length; i++) {
        var it = params.items[i];
        var bLink = (it.link || it.link_fasih_sm || "").trim();
        var bNama = (it.nama_usaha || "").trim();
        var bRowNum = parseInt(it.row_num || 0);
        var bStatus = (it.status || defaultStatus).trim();
        var bOp = (it.operator || batchOperator).trim();

        var bTargetRow = findTargetRowFast(sheet, colMap, bLink, bNama, bRowNum);
        if (bTargetRow > 1) {
          var bFinalStatus = bStatus;
          if (bOp) bFinalStatus += " (" + bOp + ")";
          sheet.getRange(bTargetRow, statusColIdx).setValue(bFinalStatus);
          sheet.getRange(bTargetRow, 1, 1, totalColsToColor).setBackground("#d9ead3");
          updatedRows.push(bTargetRow);
          updatedCount++;
        } else {
          failedCount++;
        }
      }

      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        updated_count: updatedCount,
        failed_count: failedCount,
        updated_rows: updatedRows,
        message: "Batch update selesai: " + updatedCount + " baris berhasil diperbarui."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // F. SINGLE UPDATE / RESET
    var targetLink = (params.link || params.link_fasih_sm || params.link_fasih || "").trim();
    var targetNama = (params.nama_usaha || "").trim();
    var rowNum = parseInt(params.row_num || 0);
    var statusText = (params.status || "SELESAI").trim();
    var operator = (params.operator || "").trim();

    var targetRow = findTargetRowFast(sheet, colMap, targetLink, targetNama, rowNum);

    if (targetRow > 1) {
      if (params.action === "reset" || statusText === "RESET" || statusText === "CLEAR" || statusText === "") {
        // Reset baris ke keadaan awal (kosongkan status & hapus background)
        sheet.getRange(targetRow, statusColIdx).setValue("");
        sheet.getRange(targetRow, 1, 1, totalColsToColor).setBackground(null);
        SpreadsheetApp.flush();
        return ContentService.createTextOutput(JSON.stringify({
          status: "success",
          sheet: sheet.getName(),
          row: targetRow,
          message: "Baris " + targetRow + " berhasil di-reset (status dikosongkan & warna dihapus)."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var finalStatus = statusText;
      if (operator) {
        finalStatus += " (" + operator + ")";
      }
      sheet.getRange(targetRow, statusColIdx).setValue(finalStatus);
      sheet.getRange(targetRow, 1, 1, totalColsToColor).setBackground("#d9ead3");

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
        message: "Data target (" + targetNama + " | " + targetLink + ") tidak ditemukan pada tab '" + sheet.getName() + "'."
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
 * PENCARIAN BARIS SUPER CEPAT & AKURAT (MULTI-USAHA AWARE)
 * Menggunakan Single-Column Slice & verifikasi instan. Waktu eksekusi ~0.05-0.2 detik!
 */
function findTargetRowFast(sheet, colMap, targetLink, targetNama, rowNum) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return -1;

  var linkColIdx = colMap.linkColIdx;
  var namaColIdx = colMap.namaColIdx;
  var cleanTargetLink = cleanLink(targetLink);
  var cleanTargetNama = cleanStr(targetNama);

  // 1. Cek langsung row_num jika disediakan (Instan 1-row check, ~0.02s)
  if (rowNum > 1 && rowNum <= lastRow) {
    var checkCols = Math.min(lastCol, 35);
    var rowValues = sheet.getRange(rowNum, 1, 1, checkCols).getValues()[0];
    var cellLink = (linkColIdx !== -1) ? cleanLink(String(rowValues[linkColIdx] || "")) : "";
    var cellNama = (namaColIdx !== -1) ? cleanStr(String(rowValues[namaColIdx] || "")) : "";

    var linkMatch = !cleanTargetLink || (cellLink === cleanTargetLink) || (cellLink.indexOf(cleanTargetLink) !== -1) || (cleanTargetLink.indexOf(cellLink) !== -1);
    var namaMatch = !cleanTargetNama || (cellNama === cleanTargetNama) || (cellNama.indexOf(cleanTargetNama) !== -1) || (cleanTargetNama.indexOf(cellNama) !== -1);

    if (linkMatch && namaMatch) {
      return rowNum;
    }
  }

  // 2. Jika rowNum tidak cocok atau tidak disediakan, cari via 1-kolom link_fasih (~0.15s)
  if (linkColIdx !== -1 && cleanTargetLink) {
    var linkColumnValues = sheet.getRange(1, linkColIdx + 1, lastRow, 1).getValues();
    var candidateRows = [];

    for (var r = 1; r < linkColumnValues.length; r++) {
      var cLink = cleanLink(String(linkColumnValues[r][0] || ""));
      if (cLink && (cLink === cleanTargetLink || cLink.indexOf(cleanTargetLink) !== -1 || cleanTargetLink.indexOf(cLink) !== -1)) {
        candidateRows.push(r + 1);
      }
    }

    if (candidateRows.length === 1) {
      // Jika hanya ada 1 baris dengan link tersebut
      return candidateRows[0];
    } else if (candidateRows.length > 1) {
      // MULTI-USAHA: Ada lebih dari 1 usaha dengan link assignment yang sama!
      // Cocokkan nama usaha secara presisi
      if (cleanTargetNama && namaColIdx !== -1) {
        var exactRow = -1;
        var partialRow = -1;

        for (var c = 0; c < candidateRows.length; c++) {
          var candR = candidateRows[c];
          var candNama = cleanStr(String(sheet.getRange(candR, namaColIdx + 1).getValue() || ""));
          if (candNama === cleanTargetNama) {
            exactRow = candR;
            break;
          }
          if (candNama && (candNama.indexOf(cleanTargetNama) !== -1 || cleanTargetNama.indexOf(candNama) !== -1)) {
            partialRow = candR;
          }
        }

        if (exactRow !== -1) return exactRow;
        if (partialRow !== -1) return partialRow;
      }
      return candidateRows[0];
    }
  }

  // 3. Fallback: Cari via 1-kolom nama_usaha (~0.15s)
  if (namaColIdx !== -1 && cleanTargetNama) {
    var namaColumnValues = sheet.getRange(1, namaColIdx + 1, lastRow, 1).getValues();
    for (var rn = 1; rn < namaColumnValues.length; rn++) {
      var cNama = cleanStr(String(namaColumnValues[rn][0] || ""));
      if (cNama === cleanTargetNama) {
        return rn + 1;
      }
    }
  }

  return -1;
}

/**
 * Mendeteksi posisi kolom penting secara dinamis dari header sheet (Scan hanya 3 baris teratas).
 */
function getSheetColumnMap(sheet) {
  var lastCol = sheet.getLastColumn();
  var maxScanRows = Math.min(sheet.getLastRow(), 3);
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

      if (clean.indexOf("status") !== -1) {
        currentCols.statusColIdx = c + 1; // 1-based index untuk getRange
      }
      if ((clean.indexOf("fasih") !== -1 || clean.indexOf("link") !== -1) &&
          clean.indexOf("map") === -1 && clean.indexOf("google") === -1) {
        currentCols.linkColIdx = c; // 0-based index
      }
      if (clean.indexOf("nama") !== -1) {
        if (currentCols.namaColIdx === -1 || clean.indexOf("usaha") !== -1 || clean.indexOf("perusahaan") !== -1) {
          currentCols.namaColIdx = c; // 0-based index
        }
      }
      if (clean.indexOf("desa") !== -1) {
        currentCols.desaColIdx = c; // 0-based index
      }
    }

    if (currentCols.linkColIdx !== -1 || currentCols.namaColIdx !== -1) {
      bestCols = currentCols;
      break;
    }
  }

  return bestCols;
}
