// Stockbit Screener Saver (Excel + Google Sheets Sync)
(function () {
  'use strict';

  function getScreenerTitle() {
    const headings = document.querySelectorAll('h1, h2, h3, h4');
    for (let el of headings) {
      let t = el.innerText.trim();
      if (t && !t.includes('Stockbit') && !t.includes('Screener Rules') && t.length < 100) {
        return t;
      }
    }

    const candidates = document.querySelectorAll('span, div');
    for (let el of candidates) {
      if (el.children.length === 0 && el.innerText) {
        const text = el.innerText.trim();
        const next = el.nextElementSibling;
        if (next && (next.innerText.includes('Symbols') || next.querySelector('svg'))) {
          return text;
        }
      }
    }

    return 'Stockbit Screener';
  }

  function cleanFileName(name) {
    return name.replace(/[\\/:*?"<>|]/g, '').trim();
  }

  function cleanSheetName(name) {
    let clean = name.replace(/[\\/*?[\]:]/g, '').trim();
    if (clean.length > 31) clean = clean.substring(0, 31).trim();
    return clean || 'Screener';
  }

  function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff;
    return buf;
  }

  function extractTableData() {
    const tables = document.querySelectorAll('table');
    if (!tables.length) return null;

    let rowsData = [];
    tables.forEach((table) => {
      const rows = table.querySelectorAll('tr');
      rows.forEach((row, rowIndex) => {
        const cells = row.querySelectorAll('th, td');
        if (!cells.length) return;

        let rowCells = [];
        cells.forEach((cell) => {
          let text = cell.innerText
            .replace(/\u00A0/g, ' ')
            .replace(/\n/g, ' ')
            .trim();
          rowCells.push(text);
        });

        if (!rowsData[rowIndex]) {
          rowsData[rowIndex] = rowCells;
        } else {
          rowsData[rowIndex] = rowsData[rowIndex].concat(rowCells);
        }
      });
    });

    return rowsData.filter((r) => r.some((c) => c !== ''));
  }

  // 1. Fitur Download File Excel Lokal (.xlsx)
  function downloadExcel() {
    if (typeof XLSX === 'undefined') {
      alert('Library Excel sedang dimuat, silakan coba 2 detik lagi.');
      return;
    }

    const cleanRows = extractTableData();
    if (!cleanRows || !cleanRows.length) {
      alert('Tabel Stockbit Screener tidak ditemukan atau belum dimuat!');
      return;
    }

    const rawTitle = getScreenerTitle();
    const today = new Date().toISOString().slice(0, 10);
    const fileName = `${cleanFileName(rawTitle)}_${today}.xlsx`;
    const sheetName = cleanSheetName(rawTitle);

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(cleanRows);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
    const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  // 2. Fitur Sinkronisasi Langsung ke Google Sheets
  async function syncToGoogleSheets(btnSync) {
    let webhookUrl = localStorage.getItem('stockbit_gsheets_webhook');
    if (!webhookUrl) {
      webhookUrl = prompt(
        'Masukkan URL Web App Google Apps Script Anda:\n(Contoh: https://script.google.com/macros/s/.../exec)'
      );
      if (!webhookUrl || !webhookUrl.startsWith('https://script.google.com')) {
        alert('URL Google Apps Script tidak valid. Sinkronisasi dibatalkan.');
        return;
      }
      localStorage.setItem('stockbit_gsheets_webhook', webhookUrl.trim());
    }

    const cleanRows = extractTableData();
    if (!cleanRows || cleanRows.length <= 1) {
      alert('Tabel Stockbit Screener tidak ditemukan atau belum ada data saham!');
      return;
    }

    const header = cleanRows[0].map((h) => h.toLowerCase());
    const findIndex = (keywords) => {
      return header.findIndex((h) => keywords.some((k) => h.includes(k)));
    };

    const emitenIdx = findIndex(['symbol', 'emiten', 'saham']);
    const priceIdx = findIndex(['price', 'harga', 'last']);
    const returnIdx = findIndex(['price returns', '1 day', 'change', '%']);
    const bandarValIdx = findIndex(['bandar value']);
    const netForeignIdx = findIndex(['net foreign buy / sell', 'foreign buy', 'foreign']);
    const streakIdx = findIndex(['streak']);

    let parsedRows = [];
    for (let i = 1; i < cleanRows.length; i++) {
      const row = cleanRows[i];
      let emiten = emitenIdx !== -1 ? row[emitenIdx] : row[0];
      if (emiten) emiten = emiten.split(' ')[0].trim(); // Ambil kode saham saja

      let price = priceIdx !== -1 ? parseFloat(row[priceIdx].replace(/,/g, '')) || row[priceIdx] : '';
      let return1d = returnIdx !== -1 ? row[returnIdx] : '';
      let bandarVal = bandarValIdx !== -1 ? row[bandarValIdx] : '';
      let netForeign = netForeignIdx !== -1 ? row[netForeignIdx] : '';
      let streak = streakIdx !== -1 ? row[streakIdx] : '';

      parsedRows.push({
        emiten: emiten || '',
        price: price || '',
        return1d: return1d || '',
        bandarVal: bandarVal || '',
        netForeign: netForeign || '',
        streak: streak || ''
      });
    }

    const payload = {
      screenerName: getScreenerTitle(),
      date: new Date().toISOString().slice(0, 10),
      rows: parsedRows
    };

    // Ubah status tombol jadi Loading
    const originalText = btnSync.innerHTML;
    btnSync.innerHTML = '<span>⏳ Mengirim...</span>';
    btnSync.style.opacity = '0.7';

    try {
      await fetch(webhookUrl, {
        method: 'POST',
        mode: 'no-cors', // Penting untuk Web App Google Script
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload)
      });

      alert(`✅ Berhasil! Data "${payload.screenerName}" berhasil dikirim ke Google Sheets.`);
    } catch (err) {
      alert('Gagal mengirim ke Google Sheets: ' + err.message);
    } finally {
      btnSync.innerHTML = originalText;
      btnSync.style.opacity = '1';
    }
  }

  // 3. Pasang Floating Action Bar di Halaman Screener
  function injectFloatingButtons() {
    if (!window.location.href.includes('screener')) {
      const bar = document.getElementById('stockbit-saver-action-bar');
      if (bar) bar.remove();
      return;
    }

    if (document.getElementById('stockbit-saver-action-bar')) return;

    const bar = document.createElement('div');
    bar.id = 'stockbit-saver-action-bar';
    Object.assign(bar.style, {
      position: 'fixed',
      bottom: '24px',
      right: '24px',
      zIndex: '999999',
      display: 'flex',
      gap: '10px',
      fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif'
    });

    // Tombol 1: Download Excel
    const btnExcel = document.createElement('button');
    btnExcel.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 6px;">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
        <polyline points="7 10 12 15 17 10"></polyline>
        <line x1="12" y1="15" x2="12" y2="3"></line>
      </svg>
      <span>Download Excel</span>
    `;
    Object.assign(btnExcel.style, {
      backgroundColor: '#00ab6b',
      color: '#ffffff',
      border: 'none',
      borderRadius: '24px',
      padding: '10px 16px',
      fontSize: '13px',
      fontWeight: '600',
      cursor: 'pointer',
      boxShadow: '0 4px 14px rgba(0, 171, 107, 0.4)',
      display: 'flex',
      alignItems: 'center',
      transition: 'all 0.2s ease-in-out'
    });
    btnExcel.onmouseover = () => {
      btnExcel.style.backgroundColor = '#00925b';
      btnExcel.style.transform = 'translateY(-2px)';
    };
    btnExcel.onmouseout = () => {
      btnExcel.style.backgroundColor = '#00ab6b';
      btnExcel.style.transform = 'translateY(0)';
    };
    btnExcel.addEventListener('click', downloadExcel);

    // Tombol 2: Sync to Google Sheets
    const btnSync = document.createElement('button');
    btnSync.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 6px;">
        <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/>
      </svg>
      <span style="margin-left: 6px;">Sync to Sheets</span>
    `;
    Object.assign(btnSync.style, {
      backgroundColor: '#1a73e8',
      color: '#ffffff',
      border: 'none',
      borderRadius: '24px',
      padding: '10px 16px',
      fontSize: '13px',
      fontWeight: '600',
      cursor: 'pointer',
      boxShadow: '0 4px 14px rgba(26, 115, 232, 0.4)',
      display: 'flex',
      alignItems: 'center',
      transition: 'all 0.2s ease-in-out'
    });
    btnSync.onmouseover = () => {
      btnSync.style.backgroundColor = '#1557b0';
      btnSync.style.transform = 'translateY(-2px)';
    };
    btnSync.onmouseout = () => {
      btnSync.style.backgroundColor = '#1a73e8';
      btnSync.style.transform = 'translateY(0)';
    };
    btnSync.addEventListener('click', () => syncToGoogleSheets(btnSync));

    bar.appendChild(btnExcel);
    bar.appendChild(btnSync);
    document.body.appendChild(bar);
  }

  injectFloatingButtons();

  const observer = new MutationObserver(() => {
    injectFloatingButtons();
  });

  observer.observe(document.body, {
    childList: true,
    subtree: true
  });
})();
