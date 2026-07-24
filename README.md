# Bot Otomatisasi Reject Anomali Fasih-SM BPS

Bot ini berfungsi untuk mengotomatisasi proses klik *checkbox* dan tombol *Reject* pada halaman detail *assignment* anomali di sistem web Fasih-SM BPS. Script ini akan mensimulasikan klik layaknya manusia pada browser Chromium.

## 🚀 Cara Penggunaan (Untuk Pemula)

1. **Siapkan File Excel:**
   Pastikan Anda memiliki file Excel yang berisi link anomali (bot akan membaca kolom `R`). Anda tidak perlu memfilter kode kecamatan, bot akan memproses semua link yang ada di file Excel tersebut.
2. **Masukkan ke Folder Data:**
   Buka folder proyek ini, lalu masukkan file Excel tersebut ke dalam folder `data` yang sudah tersedia.
3. **Jalankan Bot:**
   Klik dua kali (Double-click) pada file **`run.bat`**.
   - Jika komputer Anda belum memiliki Python, bot akan mendownload dan menginstallnya secara otomatis (tunggu 1-2 menit lalu jalankan ulang).
   - Bot juga akan otomatis mendownload *library* dan browser yang dibutuhkan (hanya pada run pertama).
4. **Login SSO:**
   Setelah terminal dan browser terbuka, silakan login ke Fasih-SM melalui layar SSO BPS.
5. **Mulai Otomatisasi:**
   Jika sudah berhasil masuk ke *Dashboard*, kembali ke jendela terminal hitam, lalu tekan tombol **ENTER** untuk mulai mengotomatisasi pekerjaan Anda.

## 🛠️ Fitur Unggulan
- **Otomatisasi Instalasi:** Pengguna tidak perlu menginstall `python`, `pip`, atau *library* secara manual. Cukup klik `run.bat`.
- **Anti Bot-Detection (WAF):** Dilengkapi sistem *random delay* untuk mensimulasikan kecepatan manusia membaca data, guna mencegah IP terblokir oleh *firewall* BPS.
- **Smart UI Locator:** Memiliki pencarian tombol Reject dan Konfirmasi secara dinamis berbasis JavaScript yang anti-gagal, bahkan saat halamannya lambat dimuat.
- **Resume Capability:** Jika terjadi error atau internet putus, cukup jalankan ulang `run.bat`. Sistem memiliki *cache* (`processed_links.json`) untuk mengingat link mana saja yang sudah selesai, sehingga tidak akan memproses ulang dari awal.

## 📁 Struktur Folder
```text
📦 Anomali-SE2026
 ┣ 📂 data                     <-- TARUH FILE EXCEL (.xlsx) ANDA DI SINI
 ┣ 📂 chrome_profile_anomali   <-- Folder cache cookies agar tidak perlu login terus
 ┣ 📂 venv                     <-- (Terbuat otomatis) Folder sistem Python internal
 ┣ 📜 run.bat                  <-- KLIK INI UNTUK MENJALANKAN BOT
 ┣ 📜 reject_anomali.py        <-- Script utama mesin Python
 ┣ 📜 processed_links.json     <-- File log riwayat link yang sudah sukses
 ┗ 📜 README.md
```

## 📝 Changelog

### 2026-07-24
- Menyediakan folder `data/` secara langsung dalam repositori agar pengguna bisa langsung menaruh file Excel data.
- Memperbarui `.gitignore` agar seluruh file isian di dalam folder `data/` otomatis terabaikan (tidak ter-push), hanya menjaga struktur folder melalui `.gitkeep`.
- **Fix `run.bat` Exiting**: Menulis ulang struktur `run.bat` menggunakan label `goto` untuk menghilangkan error sintaks Windows Batch (`ke was unexpected`), menangani otomatis pemindahan file Excel di folder utama ke folder `data/`, serta mencegah jendela CMD tertutup otomatis.
- **Auto Bot/SSO Recovery**: Menambahkan penanganan otomatis saat terdeteksi bot/WAF/SSO timeout (tunggu 5 detik, auto-refresh, dan auto-click "Lanjutkan dengan SSO").
