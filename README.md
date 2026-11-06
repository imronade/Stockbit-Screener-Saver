# Stockbit Screener Saver

Browser Extension (Brave, Google Chrome, Microsoft Edge) untuk mengekspor dan menyinkronkan hasil filter saham dari **Stockbit Screener** langsung ke **Excel (`.xlsx`)** atau **Google Sheets** otomatis dengan 1 klik!

![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)
![Version: 1.2.1](https://img.shields.io/badge/Version-1.2.1-orange.svg)
![Manifest](https://img.shields.io/badge/Manifest-V3-blue)
![Platform](https://img.shields.io/badge/Browser-Chrome%20%7C%20Brave%20%7C%20Edge-brightgreen)

---

## 📌 About

**Stockbit Screener Saver** mempermudah investor dan trader saham mendokumentasikan hasil riset screener Stockbit harian. Ekstensi ini menyediakan menu floating ringkas (1 tombol utama dengan pop-up menu):
1. **Download to Excel (.xlsx)**: File langsung tersimpan di komputer dengan nama file dan sheet otomatis sesuai judul screener + tanggal.
2. **Sync to Google Sheets**: Mengirim data langsung ke Google Spreadsheet di Google Drive, otomatis mencatat ke database `record_data` dengan **penambahan kolom variabel secara dinamis**, serta memperbarui `dashboard` rekap frekuensi & irisan screener lengkap dengan filter rentang tanggal (*Date Picker*).
3. **Setting Web App**: Pop-up dialog untuk mengelola dan mengedit URL Web App Google Apps Script kapan saja.

- **Versi:** `1.2.1`
- **Lisensi:** MIT License
- **Kategori:** Productivity / Financial Tools

---

## ✨ Fitur Utama (v1.2.1)

- **🟢 1-Click Floating Menu:** Ikon mengambang minimalis di pojok kanan bawah yang membuka menu pop-up: *Download to Excel*, *Sync to Sheets*, dan *Setting Web App*.
- **📊 Auto-Dynamic Columns Detection:** Otomatis mendeteksi variabel/metrik baru dari screener (misal: Technical, Bandarmologi, hingga Fundamental seperti Market Cap, DER, Operating Cash Flow) dan otomatis menambahkan kolom baru ke sheet `record_data` tanpa merusak kolom yang sudah ada.
- **📅 Interactive Date Range Filter (Date Picker):** Di sheet `dashboard`, cukup klik dua kali (*double-click*) pada kotak **Start Date** dan **End Date** untuk memfilter saham dalam periode tertentu via pop-up kalender.
- **🔢 Pemisahan Total Muncul & Total Irisan:** Membedakan antara **Total Muncul** (jumlah hari/tanggal emiten terdeteksi) dan **Total Irisan** (banyaknya jenis screener yang berhasil ditembus).
- **📂 Clean & Adaptive Template:** Template resmi hanya menyediakan kolom identitas dasar wajib; kolom variabel lainnya akan tumbuh mandiri (*self-growing database*) sesuai screener masing-masing pengguna.
- **🔒 100% Aman & Privasi Terjaga:** Berjalan di sisi klien (*client-side*), tidak membaca/menyimpan password, token, maupun kredensial akun Stockbit.

---

## 💻 Panduan Instalasi Extension

Extension ini menggunakan standar **Manifest V3** dan dapat diinstal di semua browser berbasis Chromium:

### 1. Download Repository
```bash
git clone https://github.com/imronade/Stockbit-Screener-Saver.git
```
*(atau klik tombol **Code** > **Download ZIP** lalu ekstrak filenya).*

### 2. Pasang di Browser

#### 🦁 Brave Browser
1. Buka browser Brave, ketik `brave://extensions` di URL bar lalu tekan **Enter**.
2. Aktifkan toggle **Developer mode** di pojok kanan atas.
3. Klik tombol **Load unpacked** di pojok kiri atas.
4. Pilih folder `Stockbit_Screener_Saver`.

#### 🌐 Google Chrome
1. Buka Google Chrome, ketik `chrome://extensions` di URL bar lalu tekan **Enter**.
2. Aktifkan toggle **Developer mode** di pojok kanan atas.
3. Klik tombol **Load unpacked** di pojok kiri atas.
4. Pilih folder `Stockbit_Screener_Saver`.

#### 🌊 Microsoft Edge
1. Buka Microsoft Edge, ketik `edge://extensions` di URL bar lalu tekan **Enter**.
2. Aktifkan toggle **Developer mode** di menu sebelah kiri bawah.
3. Klik tombol **Load unpacked** di bagian atas.
4. Pilih folder `Stockbit_Screener_Saver`.

---

## 📑 Setup Integrasi Google Spreadsheet (Opsional)

Jika Anda ingin menggunakan fitur **Sync to Sheets**:

1. **Download Template Resmi:**
   - Download file template resmi dari repository: [**`Hasil_Screener_Stockbit_Template.xlsx`**](Hasil_Screener_Stockbit_Template.xlsx).
   - Upload file tersebut ke Google Drive Anda, buka filenya, lalu simpan sebagai Google Spreadsheet (**File > Save as Google Sheets**).
2. **Pasang Script Otomatisasi:**
   - Di Google Spreadsheet Anda, buka menu **Extensions** > **Apps Script**.
   - Salin seluruh isi file [`google_apps_script.js`](google_apps_script.js) dan tempelkan ke editor.
   - Klik tombol **Save 💾**.
3. **Deploy Web App:**
   - Klik tombol **Deploy** > **New deployment**.
   - Klik ikon gerigi ⚙️ > pilih jenis **Web app**.
   - **Execute as:** `Me (email Anda)`.
   - **Who has access:** `Anyone (Siapa saja)`.
   - Klik **Deploy** dan berikan izin (*Authorize*).
   - Salin **Web App URL** yang berakhiran `/exec`.
4. **Hubungkan ke Extension:**
   - Di halaman Stockbit Screener, klik ikon hijau di pojok kanan bawah > pilih **Setting Web App**.
   - Tempelkan (*paste*) URL Web App tadi lalu klik **Simpan**.

---

## 📖 Cara Penggunaan

1. Buka situs [Stockbit Screener](https://stockbit.com/screener) dan pastikan Anda sudah login.
2. Buka salah satu Preset Screener Anda (misal: *Loser Bandar Masuk*, *Breakout Vol*, atau *Flow + Funda*).
3. Tunggu hingga tabel saham selesai dimuat di layar.
4. Klik **ikon lingkaran hijau** di pojok kanan bawah:
   - Pilih **Download to Excel**: Mengunduh file `.xlsx` lokal dengan nama file dan sheet otomatis.
   - Pilih **Sync to Sheets**: Mengirim data langsung ke cloud Google Sheets (otomatis auto-add kolom variabel baru dan memperbarui dashboard).

---

## 🤝 Kontribusi & Kolaborasi

Proyek ini bersifat open-source di bawah lisensi MIT. Kami menyambut kontribusi dan ide pengembangan dari komunitas trader dan developer:

1. Fork repository ini
2. Buat branch fitur baru (`git checkout -b fitur-baru`)
3. Commit perubahan (`git commit -m 'Menambahkan fitur baru'`)
4. Push ke branch Anda (`git push origin fitur-baru`)
5. Ajukan **Pull Request**

---

## 📄 Lisensi

Didistribusikan di bawah **[MIT License](LICENSE)**. Bebas digunakan, dimodifikasi, dan dikembangkan baik untuk keperluan personal maupun komersial.
