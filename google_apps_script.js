/**
 * STOCKBIT SCREENER SAVER - GOOGLE APPS SCRIPT
 * 
 * Script ini bertugas:
 * 1. Menerima data screener dari extension Stockbit
 * 2. Menyimpan data mentah ke sheet 'record_data' (append mode)
 * 3. Menghasilkan/memperbarui tabel di sheet 'dashboard' secara dinamis
 *    sesuai kategori yang ada di 'setting_kategori'.
 */

// Webhook Handler: Menerima data POST dari Extension
function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);
    const screenerName = contents.screenerName || "Stockbit Screener";
    const dateStr = contents.date || Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    const dayStr = contents.day || getIndonesianDay(new Date());
    const rows = contents.rows || []; // Array of { emiten, price, return1d, bandarVal, netForeign, streak }

    if (rows.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({ status: "empty", message: "Tidak ada baris data." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const recordSheet = ss.getSheetByName("record_data");
    const settingSheet = ss.getSheetByName("setting_kategori");

    // 1. Dapatkan mapping Kategori dari setting_kategori
    const categoryMap = getCategoryMapping(settingSheet);
    const category = categoryMap[screenerName] || "Lainnya";

    // 2. Masukkan data ke sheet 'record_data'
    let lastRow = recordSheet.getLastRow();
    let nextNo = lastRow > 1 ? (recordSheet.getRange(lastRow, 1).getValue() || 0) + 1 : 1;

    let appendRows = [];
    rows.forEach(r => {
      appendRows.push([
        nextNo++,
        dateStr,
        dayStr,
        screenerName,
        category,
        r.emiten,
        r.price,
        r.return1d,
        r.bandarVal,
        r.netForeign,
        r.streak
      ]);
    });

    if (appendRows.length > 0) {
      recordSheet.getRange(lastRow + 1, 1, appendRows.length, appendRows[0].length).setValues(appendRows);
    }

    // 3. Render ulang Dashboard secara dinamis
    refreshDashboard();

    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      message: "Berhasil menambahkan " + appendRows.length + " emiten ke Google Sheets & Dashboard terupdate!" 
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Stockbit Screener Saver Web App is active!").setMimeType(ContentService.MimeType.TEXT);
}

// Helper: Nama Hari Indonesia
function getIndonesianDay(d) {
  const days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  return days[d.getDay()];
}

// Helper: Ambil Map Kategori
function getCategoryMapping(settingSheet) {
  let map = {};
  if (!settingSheet) return map;
  const values = settingSheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const screener = String(values[i][0]).trim();
    const cat = String(values[i][1]).trim();
    if (screener) {
      map[screener] = cat || "Lainnya";
    }
  }
  return map;
}

/**
 * Fungsi Otomatis untuk Menggambar Ulang Sheet 'dashboard'
 * Menyesuaikan jumlah emiten (dinamis) dan kategori baru
 */
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dash = ss.getSheetByName("dashboard");
  if (!dash) {
    dash = ss.insertSheet("dashboard", 0);
  }
  
  const recordSheet = ss.getSheetByName("record_data");
  const settingSheet = ss.getSheetByName("setting_kategori");
  if (!recordSheet) return;

  // Bersihkan konten dashboard lama mulai baris 4 kebawah
  dash.getRange("A4:Z500").clear();

  // Header Judul Dashboard
  dash.getRange("A1").setValue("STOCKBIT SCREENER DASHBOARD").setFontSize(15).setFontWeight("bold");
  dash.getRange("A2").setValue("Rekap otomatis per kategori: Menyesuaikan jumlah emiten dinamis dari database")
    .setFontSize(9).setFontStyle("italic").setFontColor("#718096");

  // Ambil semua daftar kategori unik dari setting_kategori
  const categoryMap = getCategoryMapping(settingSheet);
  let categories = Array.from(new Set(Object.values(categoryMap)));
  if (categories.length === 0) categories = ["Buy on Weakness", "Momentum / Breakout", "Trend Following"];

  // Ambil data mentah record_data
  const recordValues = recordSheet.getDataRange().getValues();
  // baris 0 adalah header: [No, Tanggal, Hari, Nama Screener, Kategori, Emiten, Harga, 1D Ret, Bandar, NetForeign, Streak]
  
  // Warna-warna tema untuk kategori
  const palette = [
    { sec: "#1E3A8A", th: "#2563EB" }, // Biru
    { sec: "#065F46", th: "#059669" }, // Hijau Emerald
    { sec: "#92400E", th: "#D97706" }, // Amber
    { sec: "#4C1D95", th: "#7C3AED" }, // Ungu
    { sec: "#831843", th: "#DB2777" }, // Pink
  ];

  let currentRow = 4;

  categories.forEach((catName, idx) => {
    const color = palette[idx % palette.length];

    // Filter record yang kategorinya sesuai
    let emitenMap = {}; // emiten -> { totalCount, screeners: Set, lastDate, lastPrice }

    for (let i = 1; i < recordValues.length; i++) {
      const row = recordValues[i];
      const rDate = Utilities.formatDate(new Date(row[1]), "Asia/Jakarta", "yyyy-MM-dd");
      const rScreener = row[3];
      const rCat = row[4];
      const rEmiten = row[5];
      const rPrice = row[6];

      if (rCat === catName && rEmiten) {
        if (!emitenMap[rEmiten]) {
          emitenMap[rEmiten] = {
            totalCount: 0,
            screeners: new Set(),
            lastDate: rDate,
            lastPrice: rPrice
          };
        }
        emitenMap[rEmiten].totalCount++;
        emitenMap[rEmiten].screeners.add(rScreener);
        emitenMap[rEmiten].lastDate = rDate; // karena append berurut, baris terakhir = tanggal terakhir
        emitenMap[rEmiten].lastPrice = rPrice;
      }
    }

    const emitenList = Object.keys(emitenMap);
    // Jika tidak ada data pada kategori ini, tetap buat tabel kosong dengan 1 baris keterangan
    
    // 1. Judul Kategori: List Hasil Screener "Nama Kategori"
    dash.getRange(currentRow, 1, 1, 6).merge()
      .setValue(`List Hasil Screener "${catName}"`)
      .setBackground(color.sec)
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setFontSize(11)
      .setVerticalAlignment("middle");
    dash.setRowHeight(currentRow, 26);
    currentRow++;

    // 2. Table Header
    const headers = ["No.", "Emiten", "Total Muncul", "Irisan Screener", "Terakhir Muncul", "Harga Terakhir"];
    const thRange = dash.getRange(currentRow, 1, 1, 6);
    thRange.setValues([headers])
      .setBackground(color.th)
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setFontSize(10)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    dash.setRowHeight(currentRow, 22);
    currentRow++;

    // 3. Data Baris Emiten (Dinamis sesuai jumlah emiten)
    if (emitenList.length === 0) {
      dash.getRange(currentRow, 1, 1, 6).merge()
        .setValue("Belum ada data emiten untuk kategori ini.")
        .setFontStyle("italic")
        .setFontColor("#A0AEC0")
        .setHorizontalAlignment("center");
      dash.setRowHeight(currentRow, 20);
      currentRow += 2;
    } else {
      let rowsToInsert = [];
      emitenList.forEach((emiten, eIdx) => {
        const item = emitenMap[emiten];
        rowsToInsert.push([
          eIdx + 1,
          emiten,
          item.totalCount,
          Array.from(item.screeners).join(", "),
          item.lastDate,
          item.lastPrice
        ]);
      });

      const dataRange = dash.getRange(currentRow, 1, rowsToInsert.length, 6);
      dataRange.setValues(rowsToInsert)
        .setFontSize(10)
        .setVerticalAlignment("middle")
        .setBorder(true, true, true, true, true, true, "#E2E8F0", SpreadsheetApp.BorderStyle.SOLID);

      // Alignment spesifik
      dash.getRange(currentRow, 1, rowsToInsert.length, 1).setHorizontalAlignment("center"); // No
      dash.getRange(currentRow, 2, rowsToInsert.length, 1).setHorizontalAlignment("center").setFontWeight("bold"); // Emiten
      dash.getRange(currentRow, 3, rowsToInsert.length, 1).setHorizontalAlignment("right"); // Total Muncul
      dash.getRange(currentRow, 4, rowsToInsert.length, 1).setHorizontalAlignment("left"); // Irisan
      dash.getRange(currentRow, 5, rowsToInsert.length, 1).setHorizontalAlignment("center"); // Terakhir Muncul
      dash.getRange(currentRow, 6, rowsToInsert.length, 1).setHorizontalAlignment("right").setNumberFormat("#,##0"); // Harga

      for (let r = 0; r < rowsToInsert.length; r++) {
        dash.setRowHeight(currentRow + r, 20);
      }

      currentRow += rowsToInsert.length + 2; // Beri jarak 2 baris ke tabel kategori berikutnya
    }
  });
}
