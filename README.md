# Stockbit Screener Saver 📈📊

Browser Extension (Brave, Google Chrome, Microsoft Edge) untuk mengunduh hasil filter saham dari **Stockbit Screener** langsung ke dalam format **Excel (`.xlsx`)** hanya dengan 1 klik!

![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)
![Manifest](https://img.shields.io/badge/Manifest-V3-blue)
![Platform](https://img.shields.io/badge/Browser-Chrome%20%7C%20Brave%20%7C%20Edge-brightgreen)

---

## ✨ Fitur Utama

- **🚀 1-Click Export:** Tombol melayang (*floating button*) **"Download Excel"** otomatis muncul di pojok kanan bawah halaman Screener Stockbit.
- **📑 Otomatisasi Nama File & Sheet:** Nama file dan sheet Excel otomatis menggunakan judul screener Anda ditambah tanggal saat download.
  - *Format:* `[Nama Screener]_[YYYY-MM-DD].xlsx` (Contoh: `Loser Keluar BB Bandar Masuk_2026-11-06.xlsx`).
- **🎯 Full Table Support:** Mendukung pemisahan tabel modern Stockbit (menggabungkan kolom simbol/harga di sisi kiri dan kolom data finansial di sisi kanan secara sempurna).
- **🔒 100% Aman & Privasi Terjaga:** 
  - Tidak ada data yang dikirim ke server luar (*100% Client-Side/Offline*).
  - Tidak membaca username, password, token, atau informasi akun Anda.
- **🌐 Kompatibel Multi-Browser:** Berjalan lancar di semua browser berbasis Chromium (**Google Chrome**, **Brave**, **Microsoft Edge**, Opera, dll).

---

## 💻 Panduan Instalasi

Extension ini menggunakan standar **Manifest V3** dan dapat diinstal dengan mode pengembang (*Developer Mode*):

### 1. Download / Clone Repository
Pastikan Anda telah mendownload repository ini:
```bash
git clone https://github.com/imronade/Stockbit-Screener-Saver.git
```
*(atau klik tombol **Code** > **Download ZIP** lalu ekstrak ke komputer Anda).*

---

### 2. Cara Pasang di Browser

#### 🦁 Brave Browser
1. Buka browser Brave, ketik `brave://extensions` di kolom URL lalu tekan **Enter**.
2. Di pojok kanan atas, aktifkan toggle **Developer mode** (Mode pengembang).
3. Klik tombol **Load unpacked** di pojok kiri atas.
4. Pilih folder hasil download/clone tadi (`Stockbit_Screener_Saver`).
5. Selesai!

#### 🌐 Google Chrome
1. Buka Google Chrome, ketik `chrome://extensions` di kolom URL lalu tekan **Enter**.
2. Di pojok kanan atas, aktifkan toggle **Developer mode**.
3. Klik tombol **Load unpacked** di pojok kiri atas.
4. Pilih folder `Stockbit_Screener_Saver`.
5. Selesai!

#### 🌊 Microsoft Edge
1. Buka Microsoft Edge, ketik `edge://extensions` di kolom URL lalu tekan **Enter**.
2. Di menu sebelah kiri bawah, aktifkan toggle **Developer mode**.
3. Klik tombol **Load unpacked** (Muat ekstensi yang belum dibongkar) di bagian atas.
4. Pilih folder `Stockbit_Screener_Saver`.
5. Selesai!

---

## 📖 Cara Penggunaan

1. Buka situs [Stockbit Screener](https://stockbit.com/screener) dan pastikan Anda sudah login.
2. Buka salah satu Preset Screener Anda (misal: *Loser Bandar dan Asing Masuk*).
3. Tunggu hingga tabel saham selesai dimuat.
4. Perhatikan di **pojok kanan bawah layar**, akan muncul tombol hijau bertuliskan **"Download Excel"**.
5. Klik tombol tersebut, dan file `.xlsx` akan langsung terunduh rapi ke folder Downloads Anda.

---

## 🤝 Kontribusi & Kolaborasi

Proyek ini bersifat open-source! Sangat terbuka untuk ide, saran, maupun kolaborasi fitur baru (seperti export format lain, filter otomatis, dll).

1. Fork repository ini
2. Buat branch fitur baru Anda (`git checkout -b fitur-baru`)
3. Commit perubahan Anda (`git commit -m 'Menambahkan fitur baru'`)
4. Push ke branch Anda (`git push origin fitur-baru`)
5. Buat **Pull Request**

---

## 📄 Lisensi

Proyek ini dilisensikan di bawah **[MIT License](LICENSE)**. Anda bebas menggunakan, memodifikasi, mendistribusikan, dan mengembangkan lebih lanjut software ini baik untuk keperluan pribadi maupun komersial.
