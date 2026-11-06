// Stockbit Screener Saver (Unified Floating Menu & Modal Settings)
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

  // 1. Download to Excel (.xlsx)
  function downloadExcel() {
    if (typeof XLSX === 'undefined') {
      alert('Library Excel sedang dimuat, silakan coba beberapa detik lagi.');
      return;
    }

    const cleanRows = extractTableData();
    if (!cleanRows || !cleanRows.length) {
      alert('Tabel Stockbit Screener tidak ditemukan atau belum selesai dimuat!');
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

  // 2. Sync to Google Sheets
  async function syncToGoogleSheets(statusElement) {
    let webhookUrl = localStorage.getItem('stockbit_gsheets_webhook');
    if (!webhookUrl) {
      openSettingsModal('Harap isi Web App URL terlebih dahulu.');
      return;
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

    if (statusElement) statusElement.innerText = '⏳ Mengirim data...';

    try {
      await fetch(webhookUrl, {
        method: 'POST',
        mode: 'no-cors',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload)
      });

      alert(`✅ Berhasil! Data "${payload.screenerName}" (${parsedRows.length} emiten) terkirim ke Google Sheets.`);
    } catch (err) {
      alert('Gagal mengirim ke Google Sheets: ' + err.message);
    } finally {
      if (statusElement) statusElement.innerText = '';
    }
  }

  // 3. Modal Settings Pop-up
  function openSettingsModal(notice) {
    const existingModal = document.getElementById('sb-settings-modal');
    if (existingModal) existingModal.remove();

    const currentUrl = localStorage.getItem('stockbit_gsheets_webhook') || '';

    const overlay = document.createElement('div');
    overlay.id = 'sb-settings-modal';
    Object.assign(overlay.style, {
      position: 'fixed',
      top: '0',
      left: '0',
      width: '100vw',
      height: '100vh',
      backgroundColor: 'rgba(0, 0, 0, 0.5)',
      zIndex: '1000000',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif'
    });

    const modal = document.createElement('div');
    Object.assign(modal.style, {
      backgroundColor: '#ffffff',
      borderRadius: '12px',
      padding: '24px',
      width: '90%',
      maxWidth: '500px',
      boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04)'
    });

    modal.innerHTML = `
      <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
        <h3 style="margin: 0; font-size: 17px; font-weight: 700; color: #1a202c; display: flex; align-items: center; gap: 8px;">
          ⚙️ Pengaturan Web App URL
        </h3>
        <button id="sb-close-modal" style="background: none; border: none; font-size: 20px; cursor: pointer; color: #718096; padding: 0 4px;">&times;</button>
      </div>

      ${notice ? `<p style="margin: 0 0 12px; font-size: 13px; color: #e53e3e; font-weight: 500;">⚠️ ${notice}</p>` : ''}

      <p style="margin: 0 0 12px; font-size: 13px; color: #4a5568; line-height: 1.5;">
        Masukkan URL Google Apps Script Web App yang Anda dapatkan setelah klik <b>Deploy &gt; Web app</b> di spreadsheet:
      </p>

      <div style="margin-bottom: 16px;">
        <label style="display: block; font-size: 12px; font-weight: 600; color: #2d3748; margin-bottom: 6px;">
          Google Apps Script Web App URL:
        </label>
        <input id="sb-input-url" type="text" value="${currentUrl}" placeholder="https://script.google.com/macros/s/.../exec"
          ${currentUrl ? 'disabled' : ''}
          style="width: 100%; box-sizing: border-box; padding: 10px 12px; font-size: 13px; border: 1px solid #cbd5e0; border-radius: 6px; background: ${currentUrl ? '#f7fafc' : '#ffffff'}; outline: none;" />
        <small style="display: block; margin-top: 6px; font-size: 11px; color: #718096;">
          Pastikan berawalan <code>https://script.google.com/macros/s/</code> dan berakhiran <code>/exec</code>
        </small>
      </div>

      <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 20px;">
        <button id="sb-btn-edit" style="display: ${currentUrl ? 'inline-block' : 'none'}; padding: 8px 16px; border-radius: 6px; border: 1px solid #cbd5e0; background: #ffffff; color: #2d3748; font-size: 13px; font-weight: 600; cursor: pointer;">
          ✏️ Edit URL
        </button>
        <button id="sb-btn-save" style="padding: 8px 18px; border-radius: 6px; border: none; background: #00ab6b; color: #ffffff; font-size: 13px; font-weight: 600; cursor: pointer;">
          💾 Simpan
        </button>
      </div>
    `;

    overlay.appendChild(modal);
    document.body.appendChild(overlay);

    const inputUrl = document.getElementById('sb-input-url');
    const btnEdit = document.getElementById('sb-btn-edit');
    const btnSave = document.getElementById('sb-btn-save');
    const btnClose = document.getElementById('sb-close-modal');

    btnClose.onclick = () => overlay.remove();
    overlay.onclick = (e) => {
      if (e.target === overlay) overlay.remove();
    };

    btnEdit.onclick = () => {
      inputUrl.disabled = false;
      inputUrl.style.backgroundColor = '#ffffff';
      inputUrl.focus();
      btnEdit.style.display = 'none';
    };

    btnSave.onclick = () => {
      const val = inputUrl.value.trim();
      if (!val) {
        localStorage.removeItem('stockbit_gsheets_webhook');
        alert('URL telah dikosongkan.');
        overlay.remove();
        return;
      }

      if (!val.startsWith('https://script.google.com') || !val.endsWith('/exec')) {
        alert('⚠️ Format URL tidak valid!\nPastikan URL berawalan https://script.google.com/macros/s/ dan berakhiran /exec');
        return;
      }

      localStorage.setItem('stockbit_gsheets_webhook', val);
      alert('✅ Web App URL berhasil disimpan!');
      overlay.remove();
    };
  }

  // 4. Single Unified Floating Icon & Dropdown Menu
  function injectUnifiedFloatingButton() {
    if (!window.location.href.includes('screener')) {
      const container = document.getElementById('sb-unified-container');
      if (container) container.remove();
      return;
    }

    if (document.getElementById('sb-unified-container')) return;

    const container = document.createElement('div');
    container.id = 'sb-unified-container';
    Object.assign(container.style, {
      position: 'fixed',
      bottom: '24px',
      right: '24px',
      zIndex: '999999',
      fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif'
    });

    // Menu Popover (Hidden by default)
    const menu = document.createElement('div');
    menu.id = 'sb-unified-menu';
    Object.assign(menu.style, {
      position: 'absolute',
      bottom: '62px',
      right: '0',
      width: '210px',
      backgroundColor: '#ffffff',
      borderRadius: '12px',
      boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.15), 0 8px 10px -6px rgba(0, 0, 0, 0.1)',
      border: '1px solid #e2e8f0',
      overflow: 'hidden',
      display: 'none',
      flexDirection: 'column',
      padding: '6px 0',
      transition: 'all 0.2s ease-in-out'
    });

    const menuItemStyle = {
      display: 'flex',
      alignItems: 'center',
      gap: '10px',
      padding: '10px 16px',
      fontSize: '13px',
      fontWeight: '600',
      color: '#2d3748',
      background: 'none',
      border: 'none',
      width: '100%',
      textAlign: 'left',
      cursor: 'pointer',
      transition: 'background 0.15s'
    };

    // Item 1: Download to Excel
    const itemExcel = document.createElement('button');
    itemExcel.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#00ab6b" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
        <polyline points="7 10 12 15 17 10"></polyline>
        <line x1="12" y1="15" x2="12" y2="3"></line>
      </svg>
      <span>Download to Excel</span>
    `;
    Object.assign(itemExcel.style, menuItemStyle);
    itemExcel.onmouseover = () => itemExcel.style.backgroundColor = '#f0fdf4';
    itemExcel.onmouseout = () => itemExcel.style.backgroundColor = 'transparent';
    itemExcel.onclick = () => {
      menu.style.display = 'none';
      downloadExcel();
    };

    // Item 2: Sync to Sheets
    const itemSync = document.createElement('button');
    itemSync.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#1a73e8" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/>
      </svg>
      <span>Sync to Sheets</span>
    `;
    Object.assign(itemSync.style, menuItemStyle);
    itemSync.onmouseover = () => itemSync.style.backgroundColor = '#eff6ff';
    itemSync.onmouseout = () => itemSync.style.backgroundColor = 'transparent';
    itemSync.onclick = () => {
      menu.style.display = 'none';
      syncToGoogleSheets();
    };

    // Divider
    const divider = document.createElement('div');
    Object.assign(divider.style, {
      height: '1px',
      backgroundColor: '#edf2f7',
      margin: '4px 0'
    });

    // Item 3: Setting
    const itemSetting = document.createElement('button');
    itemSetting.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#718096" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="12" cy="12" r="3"></circle>
        <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path>
      </svg>
      <span>Setting Web App</span>
    `;
    Object.assign(itemSetting.style, menuItemStyle);
    itemSetting.onmouseover = () => itemSetting.style.backgroundColor = '#f7fafc';
    itemSetting.onmouseout = () => itemSetting.style.backgroundColor = 'transparent';
    itemSetting.onclick = () => {
      menu.style.display = 'none';
      openSettingsModal();
    };

    menu.appendChild(itemExcel);
    menu.appendChild(itemSync);
    menu.appendChild(divider);
    menu.appendChild(itemSetting);

    // Single Floating Main Action Button
    const mainBtn = document.createElement('button');
    mainBtn.id = 'sb-main-btn';
    mainBtn.title = 'Stockbit Screener Saver';
    mainBtn.innerHTML = `
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
        <polyline points="7 10 12 15 17 10"></polyline>
        <line x1="12" y1="15" x2="12" y2="3"></line>
      </svg>
    `;
    Object.assign(mainBtn.style, {
      backgroundColor: '#00ab6b',
      color: '#ffffff',
      border: 'none',
      borderRadius: '50%',
      width: '50px',
      height: '50px',
      cursor: 'pointer',
      boxShadow: '0 4px 14px rgba(0, 171, 107, 0.45)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      transition: 'all 0.2s ease-in-out'
    });

    mainBtn.onmouseover = () => {
      mainBtn.style.backgroundColor = '#00925b';
      mainBtn.style.transform = 'scale(1.06)';
    };
    mainBtn.onmouseout = () => {
      mainBtn.style.backgroundColor = '#00ab6b';
      mainBtn.style.transform = 'scale(1)';
    };

    mainBtn.onclick = (e) => {
      e.stopPropagation();
      menu.style.display = menu.style.display === 'flex' ? 'none' : 'flex';
    };

    document.addEventListener('click', (e) => {
      if (!container.contains(e.target)) {
        menu.style.display = 'none';
      }
    });

    container.appendChild(menu);
    container.appendChild(mainBtn);
    document.body.appendChild(container);
  }

  injectUnifiedFloatingButton();

  const observer = new MutationObserver(() => {
    injectUnifiedFloatingButton();
  });

  observer.observe(document.body, {
    childList: true,
    subtree: true
  });
})();
