/**
 * STOCKBIT SCREENER SAVER - GOOGLE APPS SCRIPT
 * 
 * Kolom Dashboard:
 * 1. No.
 * 2. Emiten
 * 3. Total Muncul (Frekuensi jumlah hari/tanggal emiten muncul)
 * 4. Total Irisan (Jumlah jenis screener yang ditembus)
 * 5. Irisan Screener (Daftar nama screener yang ditembus)
 * 6. Terakhir Muncul (Tanggal terakhir emiten muncul)
 * 7. Harga Terakhir (Harga penutupan pada kemunculan terakhir)
 */

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "No post data received" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const contents = JSON.parse(e.postData.contents);
    const screenerName = contents.screenerName || "Stockbit Screener";
    const dateStr = contents.date || Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    const dayStr = contents.day || getIndonesianDay(new Date());
    const rows = contents.rows || [];

    if (rows.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({ status: "empty", message: "Tidak ada baris data." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const recordSheet = ss.getSheetByName("record_data");
    const settingSheet = ss.getSheetByName("setting_kategori");

    if (!recordSheet) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Sheet record_data tidak ditemukan!" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 1. Dapatkan mapping Kategori dari setting_kategori
    const categoryMap = getCategoryMapping(settingSheet);
    const category = categoryMap[screenerName] || "Lainnya";

    // 2. Masukkan data ke sheet 'record_data'
    let lastRow = recordSheet.getLastRow();
    let nextNo = lastRow > 1 ? (Number(recordSheet.getRange(lastRow, 1).getValue()) || 0) + 1 : 1;

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

function getIndonesianDay(d) {
  const days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  return days[d.getDay()];
}

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
 * Render Ulang Sheet Dashboard:
 * Menghitung Total Muncul (jumlah tanggal), Total Irisan (jumlah screener), Irisan Screener (nama-namanya)
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

  // Ambil daftar kategori unik dari setting_kategori
  const categoryMap = getCategoryMapping(settingSheet);
  let categories = Array.from(new Set(Object.values(categoryMap)));
  if (categories.length === 0) categories = ["Buy on Weakness", "Momentum / Breakout", "Trend Following"];

  const recordValues = recordSheet.getDataRange().getValues();
  // baris 0 header: [No, Tanggal, Hari, Nama Screener, Kategori, Emiten, Harga, 1D Ret, Bandar, NetForeign, Streak]

  const palette = [
    { sec: "#1E3A8A", th: "#2563EB" }, // Biru
    { sec: "#065F46", th: "#059669" }, // Hijau Emerald
    { sec: "#92400E", th: "#D97706" }, // Amber
    { sec: "#4C1D95", th: "#7C3AED" }, // Ungu
    { sec: "#831843", th: "#DB2777" }, // Pink
  ];

  let currentRow = 4;
  const colCount = 7; // No, Emiten, Total Muncul, Total Irisan, Irisan Screener, Terakhir Muncul, Harga Terakhir

  categories.forEach((catName, idx) => {
    const color = palette[idx % palette.length];

    // Petakan data per emiten di kategori ini
    // emiten -> { dates: Set, screeners: Set, lastDate: String, lastPrice: Number }
    let emitenMap = {};

    for (let i = 1; i < recordValues.length; i++) {
      const row = recordValues[i];
      if (!row[5]) continue;

      let rDate = "";
      if (row[1] instanceof Date) {
        rDate = Utilities.formatDate(row[1], "Asia/Jakarta", "yyyy-MM-dd");
      } else {
        rDate = String(row[1]).slice(0, 10);
      }

      const rScreener = String(row[3]).trim();
      const rCat = String(row[4]).trim();
      const rEmiten = String(row[5]).trim();
      const rPrice = row[6];

      if (rCat === catName && rEmiten) {
        if (!emitenMap[rEmiten]) {
          emitenMap[rEmiten] = {
            dates: new Set(),
            screeners: new Set(),
            lastDate: rDate,
            lastPrice: rPrice
          };
        }
        emitenMap[rEmiten].dates.add(rDate);
        emitenMap[rEmiten].screeners.add(rScreener);
        emitenMap[rEmiten].lastDate = rDate;
        emitenMap[rEmiten].lastPrice = rPrice;
      }
    }

    const emitenList = Object.keys(emitenMap);

    // 1. Judul Kategori
    dash.getRange(currentRow, 1, 1, colCount).merge()
      .setValue(`List Hasil Screener "${catName}"`)
      .setBackground(color.sec)
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setFontSize(11)
      .setVerticalAlignment("middle");
    dash.setRowHeight(currentRow, 26);
    currentRow++;

    // 2. Table Headers (7 Kolom)
    const headers = ["No.", "Emiten", "Total Muncul", "Total Irisan", "Irisan Screener", "Terakhir Muncul", "Harga Terakhir"];
    const thRange = dash.getRange(currentRow, 1, 1, colCount);
    thRange.setValues([headers])
      .setBackground(color.th)
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setFontSize(10)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    dash.setRowHeight(currentRow, 22);
    currentRow++;

    // 3. Baris Data Emiten
    if (emitenList.length === 0) {
      dash.getRange(currentRow, 1, 1, colCount).merge()
        .setValue("(Belum ada data emiten untuk kategori ini)")
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
          item.dates.size,                          // Total Muncul (Jumlah hari/tanggal unik)
          item.screeners.size,                      // Total Irisan (Jumlah screener unik)
          Array.from(item.screeners).join(", "),    // Irisan Screener (Daftar nama screener)
          item.lastDate,                            // Terakhir Muncul
          item.lastPrice                            // Harga Terakhir
        ]);
      });

      const dataRange = dash.getRange(currentRow, 1, rowsToInsert.length, colCount);
      dataRange.setValues(rowsToInsert)
        .setFontSize(10)
        .setVerticalAlignment("middle")
        .setBorder(true, true, true, true, true, true, "#E2E8F0", SpreadsheetApp.BorderStyle.SOLID);

      dash.getRange(currentRow, 1, rowsToInsert.length, 1).setHorizontalAlignment("center"); // No
      dash.getRange(currentRow, 2, rowsToInsert.length, 1).setHorizontalAlignment("center").setFontWeight("bold"); // Emiten
      dash.getRange(currentRow, 3, rowsToInsert.length, 1).setHorizontalAlignment("center"); // Total Muncul
      dash.getRange(currentRow, 4, rowsToInsert.length, 1).setHorizontalAlignment("center"); // Total Irisan
      dash.getRange(currentRow, 5, rowsToInsert.length, 1).setHorizontalAlignment("left");   // Irisan Screener
      dash.getRange(currentRow, 6, rowsToInsert.length, 1).setHorizontalAlignment("center"); // Terakhir Muncul
      dash.getRange(currentRow, 7, rowsToInsert.length, 1).setHorizontalAlignment("right").setNumberFormat("#,##0"); // Harga

      for (let r = 0; r < rowsToInsert.length; r++) {
        dash.setRowHeight(currentRow + r, 20);
      }

      currentRow += rowsToInsert.length + 2;
    }
  });

  // Atur lebar kolom
  dash.setColumnWidth(1, 55);  // No
  dash.setColumnWidth(2, 90);  // Emiten
  dash.setColumnWidth(3, 110); // Total Muncul
  dash.setColumnWidth(4, 100); // Total Irisan
  dash.setColumnWidth(5, 260); // Irisan Screener
  dash.setColumnWidth(6, 120); // Terakhir Muncul
  dash.setColumnWidth(7, 110); // Harga Terakhir
}
