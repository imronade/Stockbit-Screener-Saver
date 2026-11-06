/**
 * STOCKBIT SCREENER SAVER - GOOGLE APPS SCRIPT (With Date Range Filter)
 * 
 * Fitur:
 * 1. Filter Rentang Tanggal (Start Date - End Date) di sheet Dashboard (Cell C3 dan E3)
 * 2. Total Muncul dihitung hanya dalam rentang tanggal yang dipilih
 * 3. Total Irisan & Irisan Screener dihitung dalam rentang tanggal
 * 4. Otomatis re-render saat user mengubah tanggal di C3/E3 via onEdit trigger
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

    const categoryMap = getCategoryMapping(settingSheet);
    const category = categoryMap[screenerName] || "Lainnya";

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

// Trigger Otomatis: Ketika user mengubah Start Date / End Date di Dashboard, tabel langsung update!
function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() === "dashboard") {
      const a1 = e.range.getA1Notation();
      // C3 = Start Date, E3 = End Date
      if (a1 === "C3" || a1 === "E3") {
        refreshDashboard();
      }
    }
  } catch (err) {}
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

function normalizeDateStr(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Jakarta", "yyyy-MM-dd");
  }
  let s = String(val).trim();
  if (s.length >= 10) return s.slice(0, 10);
  return s;
}

/**
 * Render Ulang Sheet Dashboard dengan Filter Tanggal
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

  // Baca Filter Tanggal di C3 (Start Date) dan E3 (End Date)
  let startDate = normalizeDateStr(dash.getRange("C3").getValue());
  let endDate = normalizeDateStr(dash.getRange("E3").getValue());

  // Bersihkan data tabel lama mulai baris 5 ke bawah
  dash.getRange("A5:Z500").clear();

  // Header Judul Dashboard
  dash.getRange("A1").setValue("STOCKBIT SCREENER DASHBOARD").setFontSize(15).setFontWeight("bold");

  // Filter Bar di Baris 3
  dash.getRange("B3").setValue("📅 Start Date:").setFontWeight("bold").setFontSize(10).setHorizontalAlignment("right");
  dash.getRange("D3").setValue("s/d  End Date:").setFontWeight("bold").setFontSize(10).setHorizontalAlignment("right");
  dash.getRange("F3").setValue("(Ubah tanggal di samping untuk filter)")
    .setFontSize(9).setFontStyle("italic").setFontColor("#718096").setVerticalAlignment("middle");

  dash.getRange("C3").setNumberFormat("yyyy-mm-dd");
  dash.getRange("E3").setNumberFormat("yyyy-mm-dd");

  const categoryMap = getCategoryMapping(settingSheet);
  let categories = Array.from(new Set(Object.values(categoryMap)));
  if (categories.length === 0) categories = ["Buy on Weakness", "Momentum / Breakout", "Trend Following"];

  const recordValues = recordSheet.getDataRange().getValues();
  // baris 0 header: [No, Tanggal, Hari, Nama Screener, Kategori, Emiten, Harga, 1D Ret, Bandar, NetForeign, Streak]

  const palette = [
    { sec: "#1E3A8A", th: "#2563EB" },
    { sec: "#065F46", th: "#059669" },
    { sec: "#92400E", th: "#D97706" },
    { sec: "#4C1D95", th: "#7C3AED" },
    { sec: "#831843", th: "#DB2777" },
  ];

  let currentRow = 6;
  const colCount = 7;

  categories.forEach((catName, idx) => {
    const color = palette[idx % palette.length];
    let emitenMap = {};

    for (let i = 1; i < recordValues.length; i++) {
      const row = recordValues[i];
      if (!row[5]) continue;

      const rDate = normalizeDateStr(row[1]);
      const rScreener = String(row[3]).trim();
      const rCat = String(row[4]).trim();
      const rEmiten = String(row[5]).trim();
      const rPrice = row[6];

      // Terapkan Filter Rentang Tanggal jika ada
      if (startDate && rDate < startDate) continue;
      if (endDate && rDate > endDate) continue;

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
        .setValue("(Tidak ada data emiten dalam rentang tanggal ini)")
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
          item.dates.size,                          // Total Muncul dalam rentang
          item.screeners.size,                      // Total Irisan dalam rentang
          Array.from(item.screeners).join(", "),    // Irisan Screener
          item.lastDate,                            // Terakhir Muncul
          item.lastPrice                            // Harga Terakhir
        ]);
      });

      const dataRange = dash.getRange(currentRow, 1, rowsToInsert.length, colCount);
      dataRange.setValues(rowsToInsert)
        .setFontSize(10)
        .setVerticalAlignment("middle")
        .setBorder(true, true, true, true, true, true, "#E2E8F0", SpreadsheetApp.BorderStyle.SOLID);

      dash.getRange(currentRow, 1, rowsToInsert.length, 1).setHorizontalAlignment("center");
      dash.getRange(currentRow, 2, rowsToInsert.length, 1).setHorizontalAlignment("center").setFontWeight("bold");
      dash.getRange(currentRow, 3, rowsToInsert.length, 1).setHorizontalAlignment("center");
      dash.getRange(currentRow, 4, rowsToInsert.length, 1).setHorizontalAlignment("center");
      dash.getRange(currentRow, 5, rowsToInsert.length, 1).setHorizontalAlignment("left");
      dash.getRange(currentRow, 6, rowsToInsert.length, 1).setHorizontalAlignment("center");
      dash.getRange(currentRow, 7, rowsToInsert.length, 1).setHorizontalAlignment("right").setNumberFormat("#,##0");

      for (let r = 0; r < rowsToInsert.length; r++) {
        dash.setRowHeight(currentRow + r, 20);
      }

      currentRow += rowsToInsert.length + 2;
    }
  });

  dash.setColumnWidth(1, 55);  // No
  dash.setColumnWidth(2, 90);  // Emiten
  dash.setColumnWidth(3, 110); // Total Muncul
  dash.setColumnWidth(4, 100); // Total Irisan
  dash.setColumnWidth(5, 260); // Irisan Screener
  dash.setColumnWidth(6, 120); // Terakhir Muncul
  dash.setColumnWidth(7, 110); // Harga Terakhir
}
