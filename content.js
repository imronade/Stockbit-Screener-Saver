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

  // Fungsi Setting / Edit Webhook URL
  function promptWebhookUrl() {
    const currentUrl = localStorage.getItem('stockbit_gsheets_webhook') || '';
    const newUrl = prompt(
      'Masukkan / Edit Web App URL Google Apps Script Anda:\n(Harus berakhiran /exec)',
      currentUrl
    );
    if (newUrl !== null) {
      if (newUrl.trim().startsWith('https://script.google.com') && newUrl.trim().endsWith('/exec')) {
        localStorage.setItem('stockbit_gsheets_webhook', newUrl.trim());
        alert('✅ Web App URL berhasil disimpan!');
        return newUrl.trim();
      } else if (newUrl.trim() === '') {
        localStorage.removeItem('stockbit_gsheets_webhook');
        alert('URL Web App telah dihapus.');
        return null;
      } else {
        alert('⚠️ URL tidak valid!\nPastikan URL berawalan https://script.google.com/macros/s/... dan berakhiran /exec');
        return null;
      }
    }
    return currentUrl || null;
  }

  // 2. Fitur Sinkronisasi Langsung ke Google Sheets
  async function syncToGoogleSheets(btnSync) {
    let webhookUrl = localStorage.getItem('stockbit_gsheets_webhook');
    if (!webhookUrl) {
      webhookUrl = promptWebhookUrl();
      if (!webhookUrl) return;
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
      if (emiten) emiten = emiten.split(' ')[0].trim();

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

    const originalText = btnSync.innerHTML;
    btnSync.innerHTML = '<span>⏳ Mengirim...</span>';
    btnSync.style.opacity = '0.7';

    try {
      await fetch(webhookUrl, {
        method: 'POST',
        mode: 'no-cors',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload)
      });

      alert(`✅ Selesai! Data "${payload.screenerName}" (${parsedRows.length} emiten) berhasil dikirim ke Google Sheets.`);
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
      alignItems: 'center',
      gap: '8px',
      fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif'
    });

    // Tombol 1: Download Excel
    const btnExcel = document.createElement('button');
    btnExcel.title = 'Download hasil screener ke file Excel (.xlsx)';
    btnExcel.innerHTML = `
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 5px;">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
        <polyline points="7 10 12 15 17 10"></polyline>
        <line x1="12" y1="15" x2="12" y2="3"></line>
      </svg>
      <span>Excel</span>
    `;
    Object.assign(btnExcel.style, {
      backgroundColor: '#00ab6b',
      color: '#ffffff',
      border: 'none',
      borderRadius: '20px',
      padding: '8px 14px',
      fontSize: '13px',
      fontWeight: '600',
      cursor: 'pointer',
      boxShadow: '0 4px 12px rgba(0, 171, 107, 0.35)',
      display: 'flex',
      alignItems: 'center',
      transition: 'all 0.2s'
    });
    btnExcel.addEventListener('click', downloadExcel);

    // Tombol 2: Sync to Google Sheets
    const btnSync = document.createElement('button');
    btnSync.title = 'Kirim data langsung ke Google Sheets Anda';
    btnSync.innerHTML = `
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 5px;">
        <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/>
      </svg>
      <span>Sync Sheets</span>
    `;
    Object.assign(btnSync.style, {
      backgroundColor: '#1a73e8',
      color: '#ffffff',
      border: 'none',
      borderRadius: '20px',
      padding: '8px 14px',
      fontSize: '13px',
      fontWeight: '600',
      cursor: 'pointer',
      boxShadow: '0 4px 12px rgba(26, 115, 232, 0.35)',
      display: 'flex',
      alignItems: 'center',
      transition: 'all 0.2s'
    });
    btnSync.addEventListener('click', () => syncToGoogleSheets(btnSync));

    // Tombol 3: Icon Gear Settings (Untuk ganti/edit URL Web App kapan saja)
    const btnSetting = document.createElement('button');
    btnSetting.title = 'Edit / Ganti Web App URL Google Sheets';
    btnSetting.innerHTML = `
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="12" cy="12" r="3"></circle>
        <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path>
      </svg>
    `;
    Object.assign(btnSetting.style, {
      backgroundColor: '#ffffff',
      color: '#4a5568',
      border: '1px solid #e2e8f0',
      borderRadius: '50%',
      width: '34px',
      height: '34px',
      padding: '0',
      fontSize: '13px',
      cursor: 'pointer',
      boxShadow: '0 2px 8px rgba(0, 0, 0, 0.1)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      transition: 'all 0.2s'
    });
    btnSetting.onmouseover = () => {
      btnSetting.style.backgroundColor = '#f7fafc';
      btnSetting.style.color = '#1a202c';
    };
    btnSetting.onmouseout = () => {
      btnSetting.style.backgroundColor = '#ffffff';
      btnSetting.style.color = '#4a5568';
    };
    btnSetting.addEventListener('click', promptWebhookUrl);

    bar.appendChild(btnExcel);
    bar.appendChild(btnSync);
    bar.appendChild(btnSetting);
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
