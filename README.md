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
- **Fix Timeout Tombol 'Kirim'**: Mengubah klik tombol "Kirim" dari single-shot 30 detik menjadi retry 3x dengan timeout pendek (5 detik), ditambah fallback pencarian tombol via JavaScript. Jika tombol memang tidak ada (checkbox sudah tercentang sebelumnya), bot akan langsung lanjut ke fase Reject tanpa error.
- **Fix Navigasi ke Halaman Salah**: Menambahkan verifikasi URL setelah navigasi untuk memastikan browser benar-benar berada di halaman `/edit`.

### 2026-07-25
- **Fix Skip /edit Tanpa Retry**: Jika browser di-redirect keluar dari halaman `/edit` (bukan wilayah admin), langsung skip tanpa retry 3x. Sebelumnya bot sia-sia mencoba navigasi ulang 3 kali padahal hasilnya pasti sama.
- **Fix Cache Skip**: Link yang di-skip karena tidak bisa akses `/edit` sekarang langsung disimpan ke cache (`processed_links.json`), sehingga tidak akan diproses ulang di run berikutnya.

### 2026-08-06
- **Deteksi Dinamis Kolom Excel (Dukungan Format NIK / Missing Value)**: Memperbarui fungsi pembacaan file Excel agar secara otomatis mendeteksi baris header, letak kolom `Link Fasih` (baik di Kolom R, Kolom T, dll.), serta kolom `Tindak Lanjut`.
- **Dukungan File Tanpa Kolom Status**: Jika file Excel tidak memiliki kolom status `Tindak Lanjut` (seperti file Anomali NIK Missing Value), bot akan otomatis memproses semua data visible dan mengandalkan cache lokal `processed_links.json` untuk melacak progress pengerjaan.
- **Robustness Filter Excel**: Menambahkan validasi `ws.row_dimensions[row].height == 0` selain `hidden` untuk menjamin 100% data yang di-filter pada Excel oleh user akan di-skip saat menarik link awal.
- **Rekapan Rincian Per File Excel**: Menambahkan tampilan ringkasan sebelum browser dibuka yang menyajikan jumlah Total Link Valid, Sudah Diproses (Cache), dan Sisa Diproses secara terpisah untuk setiap file Excel di folder `data`.
- **Variabel Kontrol `PROSES_MISSING_VALUE_NIK` (Default: False)**: Menambahkan variabel sakelar `PROSES_MISSING_VALUE_NIK = False` di bagian paling atas `reject_anomali.py` agar pengguna bisa dengan mudah mengaktifkan/men-disable pemrosesan file anomali Missing Value NIK secara manual.

### 2026-07-27
- **Dukungan Filter Excel (Auto-Detect Hidden Rows)**: Mengubah mekanisme pembacaan file Excel menggunakan `openpyxl` agar mengecek status baris tersembunyi (`hidden`). Jika pengguna melakukan filter (AutoFilter) pada Excel, bot hanya akan memproses data yang **tampil/visible**, dan otomatis melewati (skip) baris yang ter-filter.
- **Filter Otomatis Status Kolom O**: Menambahkan validasi pada Kolom O (kolom 15). Bot hanya akan mengambil dan memproses link yang statusnya **"Belum Ditindaklanjuti"**, sedangkan status lain ("Sudah Ditindaklanjuti...") akan otomatis dilewati.
