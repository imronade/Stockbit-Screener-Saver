// Stockbit Screener Saver (XLSX / Excel)
(function () {
  'use strict';

  function getScreenerTitle() {
    // 1. Coba cari judul h1 / h2 / h3 / h4
    const headings = document.querySelectorAll('h1, h2, h3, h4');
    for (let el of headings) {
      let t = el.innerText.trim();
      if (t && !t.includes('Stockbit') && !t.includes('Screener Rules') && t.length < 100) {
        return t;
      }
    }

    // 2. Fallback: Cari elemen span/div berukuran font besar di area atas screener
    const candidates = document.querySelectorAll('span, div');
    for (let el of candidates) {
      if (el.children.length === 0 && el.innerText) {
        const text = el.innerText.trim();
        // Cek jika teks diikuti icon copy/edit atau kata Symbols
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
    // Excel sheet name max 31 karakter dan tidak boleh ada karakter khusus
    let clean = name.replace(/[\\/*?[\]:]/g, '').trim();
    if (clean.length > 31) {
      clean = clean.substring(0, 31).trim();
    }
    return clean || 'Screener';
  }

  function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff;
    return buf;
  }

  function downloadExcel() {
    if (typeof XLSX === 'undefined') {
      alert('Library Excel sedang dimuat, silakan coba 2 detik lagi.');
      return;
    }

    const tables = document.querySelectorAll('table');
    if (!tables.length) {
      alert('Tabel Stockbit Screener tidak ditemukan!');
      return;
    }

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

    // Filter baris kosong
    const cleanRows = rowsData.filter((r) => r.some((c) => c !== ''));

    if (!cleanRows.length) {
      alert('Tidak ada data tabel yang dapat diunduh.');
      return;
    }

    // Ambil Judul Screener dan tanggal hari ini
    const rawTitle = getScreenerTitle();
    const today = new Date().toISOString().slice(0, 10); // format YYYY-MM-DD
    const fileName = `${cleanFileName(rawTitle)}_${today}.xlsx`;
    const sheetName = cleanSheetName(rawTitle);

    // Buat Workbook & Worksheet Excel dengan SheetJS
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(cleanRows);

    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    // Download file .xlsx
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
    const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  function injectFloatingButton() {
    if (!window.location.href.includes('screener')) {
      const existingBtn = document.getElementById('stockbit-csv-download-btn');
      if (existingBtn) existingBtn.remove();
      return;
    }

    if (document.getElementById('stockbit-csv-download-btn')) return;

    const btn = document.createElement('button');
    btn.id = 'stockbit-csv-download-btn';
    btn.innerHTML = `
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-right: 6px;">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
        <polyline points="7 10 12 15 17 10"></polyline>
        <line x1="12" y1="15" x2="12" y2="3"></line>
      </svg>
      <span>Download Excel</span>
    `;

    Object.assign(btn.style, {
      position: 'fixed',
      bottom: '24px',
      right: '24px',
      zIndex: '999999',
      backgroundColor: '#00ab6b',
      color: '#ffffff',
      border: 'none',
      borderRadius: '24px',
      padding: '10px 18px',
      fontSize: '13px',
      fontWeight: '600',
      fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
      cursor: 'pointer',
      boxShadow: '0 4px 14px rgba(0, 171, 107, 0.4)',
      display: 'flex',
      alignItems: 'center',
      transition: 'all 0.2s ease-in-out'
    });

    btn.onmouseover = () => {
      btn.style.backgroundColor = '#00925b';
      btn.style.transform = 'translateY(-2px)';
      btn.style.boxShadow = '0 6px 18px rgba(0, 171, 107, 0.5)';
    };

    btn.onmouseout = () => {
      btn.style.backgroundColor = '#00ab6b';
      btn.style.transform = 'translateY(0)';
      btn.style.boxShadow = '0 4px 14px rgba(0, 171, 107, 0.4)';
    };

    btn.addEventListener('click', downloadExcel);
    document.body.appendChild(btn);
  }

  injectFloatingButton();

  const observer = new MutationObserver(() => {
    injectFloatingButton();
  });

  observer.observe(document.body, {
    childList: true,
    subtree: true
  });
})();
