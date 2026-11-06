# Stockbit Screener Saver

Browser Extension (Brave, Google Chrome, Microsoft Edge) untuk mengekspor dan menyinkronkan hasil filter saham dari **Stockbit Screener** langsung ke **Excel (`.xlsx`)** atau **Google Sheets** otomatis dengan 1 klik!

![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)
![Version: 1.2.0](https://img.shields.io/badge/Version-1.2.0-orange.svg)
![Manifest](https://img.shields.io/badge/Manifest-V3-blue)
![Platform](https://img.shields.io/badge/Browser-Chrome%20%7C%20Brave%20%7C%20Edge-brightgreen)

---

## 📌 About

**Stockbit Screener Saver** mempermudah investor dan trader saham mendokumentasikan hasil riset screener Stockbit harian. Ekstensi ini menyediakan dua opsi ekspor:
1. **Download Excel (.xlsx)**: File langsung tersimpan di komputer dengan nama file dan sheet otomatis sesuai judul screener + tanggal.
2. **Sync to Google Sheets**: Mengirim data mentah langsung ke spreadsheet cloud di Google Drive, otomatis mengelompokkan ke database `record_data`, dan memperbarui `dashboard` rekap frekuensi & irisan screener secara dinamis.

- **Versi:** `1.2.0`
- **Lisensi:** MIT License
- **Kategori:** Productivity / Financial Tools

---

## ✨ Fitur Utama

- **🟢 Download Excel (.xlsx):** Sekali klik, menghasilkan file Excel asli lengkap dengan nama file dan sheet dinamis `[Nama Screener]_[YYYY-MM-DD].xlsx`.
- **🔵 Sync to Google Sheets:** Integrasi webhook Google Apps Script yang otomatis mengisi database mentah dan merekap dashboard emiten.
- **📊 Auto-Scaling Dashboard:** Tabel rekap pada Google Spreadsheet otomatis menyesuaikan jumlah emiten harian dan otomatis membuat blok tabel baru jika Anda menambah kategori baru di `setting_kategori`.
- **🎯 Full Table Parsing:** Menggabungkan tabel header, kolom simbol/harga, dan kolom finansial Stockbit secara utuh.
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

Jika Anda atau rekan Anda ingin menggunakan fitur **Sync to Sheets**:

1. **Gunakan Template Spreadsheet Resmi:**
   - Download file template dari repository ini: [**`Hasil_Screener_Stockbit_Template.xlsx`**](Hasil_Screener_Stockbit_Template.xlsx).
   - Upload file tersebut ke Google Drive Anda, lalu buka dan simpan sebagai Google Spreadsheet (**File > Save as Google Sheets**).
   - Template ini sudah dilengkapi dengan sheet `panduan_penggunaan`, `dashboard`, `record_data`, dan `setting_kategori`.
2. **Pasang Script:**
   - Di Google Spreadsheet, buka menu **Extensions** > **Apps Script**.
   - Salin seluruh kode dari file [`google_apps_script.js`](google_apps_script.js) dan tempelkan ke editor Apps Script.
3. **Deploy Web App:**
   - Klik tombol **Deploy** > **New deployment**.
   - Pilih jenis **Web app**.
   - **Execute as:** `Me (email Anda)`.
   - **Who has access:** `Anyone`.
   - Klik **Deploy** dan berikan izin (*Authorize*).
   - Salin URL Web App yang berakhiran `/exec`.
4. **Hubungkan ke Extension:**
   - Buka Stockbit Screener, klik tombol biru **Sync to Sheets**.
   - Masukkan URL Web App yang sudah disalin tadi saat diminta prompt. URL ini akan tersimpan otomatis untuk klik-klik berikutnya.

---

## 📖 Cara Penggunaan

1. Buka situs [Stockbit Screener](https://stockbit.com/screener) dan pastikan Anda sudah login.
2. Buka salah satu Preset Screener Anda (misal: *Loser Bandar dan Asing Masuk*).
3. Tunggu hingga tabel saham selesai dimuat.
4. Perhatikan di pojok kanan bawah layar akan muncul dua tombol:
   - 🟢 **Download Excel**: Untuk mengunduh file `.xlsx` lokal.
   - 🔵 **Sync to Sheets**: Untuk mengirim langsung ke Google Spreadsheet.
5. Klik tombol yang diinginkan. Selesai!

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
