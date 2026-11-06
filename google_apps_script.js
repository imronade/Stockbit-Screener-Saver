/**
 * STOCKBIT SCREENER SAVER - GOOGLE APPS SCRIPT (Dynamic Columns Support)
 * 
 * Otomatis mendeteksi kolom baru dari berbagai macam screener (misal: Fundamental, Market Cap, DER, dsb).
 * Jika ada kolom baru yang belum ada di 'record_data', script akan otomatis menambahkannya ke samping kanan.
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

    // 1. Kategori mapping
    const categoryMap = getCategoryMapping(settingSheet);
    const category = categoryMap[screenerName] || "Lainnya";

    // 2. Baca Header yang sudah ada di sheet 'record_data'
    let currentHeaders = [];
    const lastCol = recordSheet.getLastColumn();
    if (lastCol > 0) {
      currentHeaders = recordSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    }

    // Header Dasar Wajib
    const baseHeaders = ["No", "Tanggal", "Hari", "Nama Screener", "Kategori", "Emiten", "Harga Penutupan"];
    if (currentHeaders.length === 0) {
      currentHeaders = [...baseHeaders];
    }

    // 3. Deteksi Kolom/Metrik Baru yang Dibawa oleh Screener Ini
    // Kumpulkan semua nama metrik unik dari seluruh baris
    let incomingMetricNames = new Set();
    rows.forEach(r => {
      if (r.metrics) {
        Object.keys(r.metrics).forEach(m => incomingMetricNames.add(m));
      }
    });

    // Cek apakah ada kolom baru yang belum ada di record_data
    let newColumnsToAdd = [];
    incomingMetricNames.forEach(colName => {
      // Kecualikan jika sudah merupakan salah satu header dasar
      if (!currentHeaders.includes(colName) && !["symbol", "emiten", "saham", "price", "harga"].includes(colName.toLowerCase())) {
        newColumnsToAdd.push(colName);
        currentHeaders.push(colName);
      }
    });

    // Jika ada kolom baru, tambahkan header baru di baris 1 record_data
    if (newColumnsToAdd.length > 0) {
      const startNewCol = currentHeaders.length - newColumnsToAdd.length + 1;
      const headerRange = recordSheet.getRange(1, startNewCol, 1, newColumnsToAdd.length);
      headerRange.setValues([newColumnsToAdd])
        .setBackground("#00875A")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold")
        .setFontSize(10)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
      recordSheet.setRowHeight(1, 26);
    }

    // 4. Susun Baris Data Sesuai Urutan Header yang Terkini
    let lastRow = recordSheet.getLastRow();
    let nextNo = lastRow > 1 ? (Number(recordSheet.getRange(lastRow, 1).getValue()) || 0) + 1 : 1;

    let appendRows = [];
    rows.forEach(r => {
      let rowData = new Array(currentHeaders.length).fill("");
      rowData[0] = nextNo++;
      rowData[1] = dateStr;
      rowData[2] = dayStr;
      rowData[3] = screenerName;
      rowData[4] = category;
      rowData[5] = r.emiten;
      rowData[6] = r.price;

      // Isi kolom metrik dinamis
      if (r.metrics) {
        Object.keys(r.metrics).forEach(key => {
          let colIdx = currentHeaders.indexOf(key);
          if (colIdx !== -1) {
            rowData[colIdx] = r.metrics[key];
          }
        });
      }
      appendRows.push(rowData);
    });

    if (appendRows.length > 0) {
      recordSheet.getRange(lastRow + 1, 1, appendRows.length, currentHeaders.length).setValues(appendRows);
    }

    // 5. Update Dashboard
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

function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() === "dashboard") {
      const a1 = e.range.getA1Notation();
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
 * Render Ulang Sheet Dashboard
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

  const ruleDate = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .setHelpText("Klik dua kali untuk memilih tanggal dari kalender")
    .build();

  dash.getRange("C3").setDataValidation(ruleDate);
  dash.getRange("E3").setDataValidation(ruleDate);

  let startDate = normalizeDateStr(dash.getRange("C3").getValue());
  let endDate = normalizeDateStr(dash.getRange("E3").getValue());

  dash.getRange("A5:Z500").clear();

  dash.getRange("A1").setValue("STOCKBIT SCREENER DASHBOARD").setFontSize(15).setFontWeight("bold");

  dash.getRange("B3").setValue("📅 Start Date:").setFontWeight("bold").setFontSize(10).setHorizontalAlignment("right");
  dash.getRange("D3").setValue("s/d  End Date:").setFontWeight("bold").setFontSize(10).setHorizontalAlignment("right");
  dash.getRange("F3").setValue("(Double-click kotak untuk pilih tanggal kalender | Kosongkan untuk semua)")
    .setFontSize(9).setFontStyle("italic").setFontColor("#718096").setVerticalAlignment("middle");

  dash.getRange("C3").setNumberFormat("yyyy-mm-dd");
  dash.getRange("E3").setNumberFormat("yyyy-mm-dd");

  const categoryMap = getCategoryMapping(settingSheet);
  let categories = Array.from(new Set(Object.values(categoryMap)));
  if (categories.length === 0) categories = ["Buy on Weakness", "Momentum / Breakout", "Trend Following"];

  const recordValues = recordSheet.getDataRange().getValues();

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

    dash.getRange(currentRow, 1, 1, colCount).merge()
      .setValue(`List Hasil Screener "${catName}"`)
      .setBackground(color.sec)
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setFontSize(11)
      .setVerticalAlignment("middle");
    dash.setRowHeight(currentRow, 26);
    currentRow++;

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
          item.dates.size,
          item.screeners.size,
          Array.from(item.screeners).join(", "),
          item.lastDate,
          item.lastPrice
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

  dash.setColumnWidth(1, 55);
  dash.setColumnWidth(2, 90);
  dash.setColumnWidth(3, 110);
  dash.setColumnWidth(4, 100);
  dash.setColumnWidth(5, 260);
  dash.setColumnWidth(6, 120);
  dash.setColumnWidth(7, 110);
}
