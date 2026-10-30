/**
 * Fungsi utama untuk backup harian:
 * 1. Menambah data baru yang belum ada.
 * 2. Memperbarui (update) data lama jika ada revisi pada tanggal + hub yang sama.
 * 3. Data lama di backup tidak akan hilang meskipun di sheet asal sudah dihapus.
 */
function autoBackupDataWithUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Daftar sheet sumber, sheet tujuan, serta posisi kolom (0-based index)
  const sheetsToBackup = [
    // Kolom A = 0, Kolom B = 1, Kolom C = 2, Kolom D = 3
    { sourceName: "FM Reguler Days [Full]", backupName: "FM Reguler Days [Backup]", dateCol: 0, hubCol: 1 },
    { sourceName: "FM High Days [Full]", backupName: "FM High Days [Backup]", dateCol: 0, hubCol: 1 },
    { sourceName: "LM [Full]", backupName: "LM [Backup]", dateCol: 2, hubCol: 3 } // report_date di Kolom C (2), hub_name di Kolom D (3)
  ];

  sheetsToBackup.forEach(item => {
    processIncrementalBackup(ss, item.sourceName, item.backupName, item.dateCol, item.hubCol);
  });
}

function processIncrementalBackup(ss, sourceName, backupName, dateColIdx = 0, hubColIdx = 1) {
  const sourceSheet = ss.getSheetByName(sourceName);
  if (!sourceSheet) {
    Logger.log("Sheet sumber tidak ditemukan: " + sourceName);
    return;
  }

  let backupSheet = ss.getSheetByName(backupName);
  if (!backupSheet) {
    backupSheet = ss.insertSheet(backupName);
  }

  const sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length <= 1) {
    Logger.log("Tidak ada data di sheet sumber: " + sourceName);
    return;
  }

  const sourceHeader = sourceData[0];
  const sourceRows = sourceData.slice(1);

  const backupLastRow = backupSheet.getLastRow();

  // 1. Jika sheet backup masih kosong sama sekali, salin seluruh isi beserta header
  if (backupLastRow === 0) {
    backupSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
    Logger.log("Backup perdana berhasil dibuat untuk: " + backupName);
    return;
  }

  // 2. Jika sheet backup sudah ada data
  const backupRange = backupSheet.getDataRange();
  const backupData = backupRange.getValues();
  const backupRows = backupData.slice(1);

  // Fungsi pembantu membuat Kunci Unik yang konsisten (menangani objek Date)
  function generateKey(row) {
    let dateVal = row[dateColIdx];
    let hubVal = row[hubColIdx];

    if (dateVal instanceof Date) {
      dateVal = Utilities.formatDate(dateVal, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    }
    return String(dateVal).trim() + "_" + String(hubVal).trim();
  }

  // Buat Map posisi baris di sheet backup berdasarkan Kunci Unik
  const backupRowIndexMap = new Map();
  backupRows.forEach((row, index) => {
    const key = generateKey(row);
    backupRowIndexMap.set(key, index + 2); // +2 karena offset 1-based index & melewati header
  });

  const newRowsToAppend = [];
  let updatedCount = 0;

  sourceRows.forEach(sourceRow => {
    const sourceKey = generateKey(sourceRow);

    if (backupRowIndexMap.has(sourceKey)) {
      // DATA SUDAH ADA: Update baris yang bersangkutan jika isinya mengalami perubahan
      const targetRowNumber = backupRowIndexMap.get(sourceKey);
      backupSheet.getRange(targetRowNumber, 1, 1, sourceRow.length).setValues([sourceRow]);
      updatedCount++;
    } else {
      // DATA BELUM ADA: Masukkan ke daftar baris baru untuk ditambahkan di paling bawah
      newRowsToAppend.push(sourceRow);
    }
  });

  // Tambahkan semua data baru di baris paling bawah
  if (newRowsToAppend.length > 0) {
    backupSheet
      .getRange(backupSheet.getLastRow() + 1, 1, newRowsToAppend.length, newRowsToAppend[0].length)
      .setValues(newRowsToAppend);
  }

  Logger.log(`[${backupName}] Selesai: ${newRowsToAppend.length} baris baru ditambahkan, ${updatedCount} baris lama diperbarui.`);
}