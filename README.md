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
 ┣ 📂 data                     <-- FOLDER UTAMA DATA EXCEL
 ┃ ┣ 📂 Anomali                <-- Taruh file Excel anomali di sini (cth: anomaly 28 sept.xlsx)
 ┃ ┣ 📂 BKU Ditautkan          <-- Taruh file Excel BKU di sini (cth: 12. BKU Ditautkan tapi Status Ganda.xlsx)
 ┃ ┗ 📂 Kawasan                <-- Taruh file Excel anomali kawasan di sini (cth: sqllab...xlsx)
 ┣ 📂 chrome_profile_anomali   <-- Folder cache cookies agar tidak perlu login terus
 ┣ 📂 venv                     <-- (Terbuat otomatis) Folder sistem Python internal
 ┣ 📜 run.bat                  <-- KLIK INI UNTUK MENJALANKAN MENU BOT
 ┣ 📜 run_edit_admin.bat       <-- Launcher khusus Bot Edit Anomali by Admin
 ┣ 📜 run_status_ditemukan.bat <-- Launcher khusus Bot Ubah Status Ditemukan
 ┣ 📜 run_ubah_kawasan.bat     <-- Launcher khusus Bot Ubah Jenis Kawasan ke Diluar Kawasan
 ┣ 📜 run_tambah_remark.bat    <-- Launcher khusus Bot Tambah Remark 27.a Nilai Produksi
 ┣ 📜 edit_anomali_admin.py    <-- Script bot konfirmasi & perbaikan anomali by admin
 ┣ 📜 reject_anomali.py        <-- Script bot reject anomali
 ┣ 📜 ubah_status_ditemukan.py <-- Script bot ubah status '1. Ditemukan' & Submit Paksa
 ┣ 📜 ubah_jenis_kawasan.py    <-- Script bot ubah '8. d. Jenis kawasan' ke '10. Diluar kawasan'
 ┣ 📜 tambah_remark_27a.py     <-- Script bot tambah remark #0 pada pertanyaan 27.a nilai produksi
 ┣ 📜 processed_edit_admin.json <-- Log cache bot edit by admin
 ┣ 📜 processed_links.json     <-- Log cache bot reject
 ┣ 📜 processed_status_ditemukan.json <-- Log cache bot ubah status
 ┣ 📜 processed_jenis_kawasan.json    <-- Log cache bot ubah kawasan
 ┣ 📜 processed_tambah_remark.json    <-- Log cache bot tambah remark 27.a
 ┗ 📜 README.md
```

## 📝 Changelog

### 2026-10-22
- **Optimasi Kecepatan Eksekusi Bot 2x-3x Lebih Cepat (`tambah_remark_27a.py`)**:
  - Memangkas jeda waktu tidur buatan (*idle sleep*) pada seluruh alur pengerjaan assignment (buka menu Blok II, buka rincian usaha, pencarian scroll pertanyaan 27.a, input remark #0, simpan catatan, buka menu Catatan, toggle switch anomali, submit konfirmasi, hingga jeda antar-assignment).
  - Mengurangi total delay menganggur dari ~35+ detik menjadi hanya ~8–10 detik per assignment tanpa mengorbankan stabilitas interaksi Playwright.
  - Mempercepat interval polling kesiapan kuesioner dan feedback submit agar bot langsung merespons begitu server Fasih-SM selesai memproses data.
- **Fix False-Alarm Sesi Terputus & Eliminasi Jeda Login Ulang (`tambah_remark_27a.py`)**:
  - Memperbaiki bug false-positive "Error koneksi / halaman blank: halaman kosong (blank)" yang kerap muncul setiap beberapa assignment dan memaksa bot berhenti meminta tekan ENTER.
  - Masalah terjadi karena Fasih-SM merupakan Single Page Application (SPA), di mana teks `document.body` belum selesai di-render oleh JavaScript sesaat setelah event `DOMContentLoaded`. Pengecekan blank prematur ini kini telah dihapus dari fungsi validasi koneksi.
  - Menambahkan mekanisme *Self-Healing Auto-Reload*: jika kuesioner lambat dimuat (>12 detik), bot otomatis me-reload halaman 1 kali di latar belakang tanpa mengganggu jalannya otomatisasi.
  - Menambahkan *Self-Check Guard* pada `pause_for_relogin`: bot memverifikasi URL aktif di browser terlebih dahulu. Jika URL masih di dalam `/app/assignment/` dan bukan login SSO, bot langsung melanjutkan pengerjaan secara mulus (*unattended*) tanpa menghentikan proses.
  - Memperbaiki penanganan exception agar kesalahan timeout elemen atau JS DOM tidak salah diidentifikasi sebagai `SESSION_LOST`.
- **Deteksi Header Kolom Dinamis & Tahan Perubahan Pengguna (`tambah_remark_27a.py`, `google_apps_script_sync.js`)**:
  - Mengubah seluruh sistem pembacaan dan penulisan kolom agar 100% berbasis nama header (*header name dynamic mapping*), bukan berdasarkan nomor/huruf indeks kolom statis.
  - Kebal terhadap penggeseran urutan kolom, penyisipan kolom baru oleh teman/rekan lain, perubahan huruf besar/kecil, variasi spasi, maupun *underscore* (misal: "Link Fasih SM", "link_fasih", "Status", "Status Pengerjaan", "Nama Usaha", "Kode Desa").
  - Menambahkan fungsi `getSheetColumnMap` pada skrip Google Apps Script untuk mendeteksi header secara dinamis di baris 1-5, memastikan penulisan status dan pewarnaan hijau selalu tepat sasaran tanpa menduplikasi kolom.
  - Menambahkan cache TTL 60 detik pada `get_source_column_letters` di Python agar jika rekan kerja mengubah struktur spreadsheet saat bot sedang berjalan, bot otomatis menyesuaikan dengan struktur kolom terbaru.
- **Update Utilitas Push Histori & Paket Distribusi (`push_local_cache_to_sheets.py`, `Bot-Tambah-Remark-SE2026.zip`)**:
  - Menambahkan konfigurasi `sys.stdout.reconfigure(encoding='utf-8')` pada Windows console di `push_local_cache_to_sheets.py` agar aman dari error encoding karakter emoji pada nama tab Google Spreadsheet.
  - Mengemas ulang arsip `Bot-Tambah-Remark-SE2026.zip` dengan kode termutakhir dan template konfigurasi siap pakai (`operator_name` kosong) untuk dibagikan ke rekan operator.
- **Fix URL Routing Multi-Spreadsheet Live Sync (`tambah_remark_27a.py`)**:
  - Memperbaiki bug perutean URL Web App pada `sync_to_google_sheet`: sebelumnya penulisan status selesai untuk Prioritas 2 (Data NTB / Karangawen) terkirim ke spreadsheet yang salah karena pencocokan nama ("Dtata_NTB" vs "Prioritas 2: Data NTB") tidak cocok dan jatuh ke fallback Prioritas 1.
  - Mengimplementasikan fungsi `resolve_gas_url_for_item` yang cerdas untuk memetakan nama sumber, parameter gid, dan alias spreadsheet target ke URL Web App yang tepat secara akurat.
  - Memperbarui `sheet_sync_config.json` dengan alias lengkap dan mengeksekusi sinkronisasi histori lokal sehingga baris Karangawen (*BAWANG MERAH PAK KASWADI* baris 10015) serta baris Bonang lainnya kini sudah langsung bertuliskan `SELESAI (Muhshi)` dan berwarna hijau di Google Spreadsheet target.
- **Fix Deteksi Kolom Dinamis & Sinkronisasi Total Awal Multi-Operator (`tambah_remark_27a.py`)**:
  - Memperbaiki bug deteksi kolom di Google Sheets: Prioritas 1 dan Prioritas 2 memiliki indeks kolom yang berbeda (Prioritas 1: link di Z, status di AA; Prioritas 2: link di AA, status di AB). Bot kini menggunakan deteksi huruf kolom secara dinamis (`get_source_column_letters`), mencegah kegagalan query status pada Prioritas 2.
  - Menyempurnakan fungsi `is_done(it)` agar langsung memperhitungkan `it.get("is_sheet_done")`, serta memindai seluruh data live spreadsheet di awal dan menyinkronkan setiap baris selesai (dari operator manapun) ke dalam cache JSON lokal (`processed_tambah_remark.json`).
  - Memperbaiki filter auto-hide wilayah: Kecamatan yang sudah tuntas 100% (seperti Bonang dan Wonosalam pada Prioritas 1) kini otomatis disembunyikan dari daftar pilihan, dan kecamatan yang sebagian sudah dikerjakan rekan lain (seperti Mijen oleh Aris) langsung berkurang jumlah sisa tugasnya secara real-time.
  - Menambahkan parameter anti-cache CDN (`_t=timestamp` dan header `Cache-Control: no-cache`) pada `fetch_single_sheet` agar data CSV yang diunduh selalu kondisi paling mutakhir dari Google Spreadsheet.
- **Sinkronisasi Realtime Google Spreadsheet Sebelum List Wilayah (`tambah_remark_27a.py`)**:
  - Sebelum menampilkan daftar Kecamatan, ID Desa, Batch, maupun Semua Data di terminal, bot otomatis mengambil status pengerjaan terbaru (`Status Pengerjaan`) dari Google Spreadsheet secara realtime menggunakan query berkecepatan tinggi dengan bypass cache CDN (`_t=timestamp`).
  - Mengupdate status `raw_data` in-memory dan memperbarui file cache JSON lokal (`processed_tambah_remark.json`) secara instan, sehingga cache lokal selalu sinkron dengan apa yang dikerjakan oleh operator lain di Google Spreadsheet.
  - Jumlah sisa data dan filter auto-hide pada daftar kecamatan dan desa selalu 100% akurat dan realtime.
- **Pengecekan Status Realtime Tiap Assignment Sebelum Dijalankan (`tambah_remark_27a.py`)**:
  - Di dalam loop eksekusi, tepat sebelum membuka link assignment di browser Fasih-SM, bot melakukan verifikasi live langsung ke Google Spreadsheet.
  - Jika assignment sudah selesai dikerjakan oleh operator/rekan lain (status `SELESAI`, `DONE`, `SUDAH`, dsb), bot otomatis melewatinya (`[SKIP REALTIME]`), menambahkan ke cache lokal, dan langsung melompat ke assignment berikutnya.
  - Mencegah duplikasi pengerjaan / tabrakan antar pengguna saat banyak orang menjalankan bot bersamaan.
- **Optimasi Google Apps Script Sync v3 (`google_apps_script_sync.js`)**:
  - Menambahkan dukungan request cepat `check_status` dan `get_completed`.
  - Operasi read-only kini tidak mengunci script lock (`lock.tryLock`) sehingga pengecekan status berlangsung non-blocking dan instan dari banyak PC sekaligus. Operasi penulisan status dan pewarnaan hijau tetap terlindungi ScriptLock.
- **Pembaruan Paket Zip Distribusi (`Bot-Tambah-Remark-SE2026.zip`) & Panduan Multi-Operator**:
  - Mengemas ulang seluruh kode utama, launcher, script sinkronisasi, template konfigurasi operator, dan panduan penggunaan ke dalam file zip mandiri `Bot-Tambah-Remark-SE2026.zip` yang siap dibagikan ke rekan-rekan kerja.

### 2026-10-21
- **Paket Distribusi Siap Pakai untuk Rekan Kerja (`Bot-Tambah-Remark-SE2026.zip`)**:
  - Membuat bundle zip mandiri khusus bot tambah remark yang siap dibagikan ke rekan kerja.
  - Sudah menyertakan link Google Apps Script live sync, data backup offline, dan master kode wilayah, dengan `operator_name` kosong sehingga teman Anda langsung diminta memasukkan namanya saat pertama kali dijalankan.
  - Menyertakan file panduan `PANDUAN_PENGGUNAAN.txt` di dalam zip.
- **Penamaan Sumber Data: Prioritas 1 & Prioritas 2 (`tambah_remark_27a.py`)**:
  - Mengubah label sumber spreadsheet di terminal menjadi:
    - `[1] Prioritas 1: Data Detail Anomali (gid: 0) [DEFAULT]`
    - `[2] Prioritas 2: Data NTB (gid: 293925488)`
    - `[3] GABUNGKAN SEMUA SHEET (Prioritas 1 + Prioritas 2)`
  - Menyediakan script utilitas `push_local_cache_to_sheets.py` untuk mensinkronkan seluruh riwayat data lokal yang sudah pernah sukses dikerjakan ke Google Sheets.
- **Dukungan Multi-Spreadsheet & Keamanan Script URL (`sheet_sync_config.json` & `.gitignore`)**:
  - Menambahkan dukungan sinkronisasi ke multi-spreadsheet (`gas_webapp_urls`): jika pengguna mengerjakan *Data detail anomali* atau *Dtata_NTB*, bot secara cerdas mengarahkan request ke URL Apps Script spreadsheet yang bersangkutan.
  - Memasukkan `sheet_sync_config.json` ke `.gitignore` sehingga URL Web App privat dan nama operator **tidak akan pernah ter-push ke GitHub**. Disediakan `sheet_sync_config.example.json` sebagai template repositori.
  - **Penyempurnaan Google Apps Script v2**: Apps Script kini mendeteksi tab target secara otomatis via parameter `gid` (misal `gid=0` atau `gid=293925488`), pencocokan nama tab non-case-sensitive, maupun deteksi kolom header `link_fasih`, sehingga 100% tepat sasaran meskipun nama tab diubah atau diurutkan ulang.
- **Registrasi Nama Operator Otomatis (`tambah_remark_27a.py`)**:
  - Jika nama operator di `sheet_sync_config.json` masih kosong, terminal akan otomatis meminta nama/panggilan pengguna hanya pada saat run pertama kali.
  - Nama tersebut otomatis disimpan di konfigurasi lokal sehingga tidak akan ditanyakan lagi pada run berikutnya.
  - Nama operator otomatis dicantumkan di kolom spreadsheet: `SELESAI (Nama Operator)` agar rekan kerja mengetahui siapa yang memproses baris tersebut.
- **Filter Berdasarkan Kecamatan & Auto-Hide Wilayah Selesai (`tambah_remark_27a.py`)**:
  - Menambahkan menu filter baru **`[1] Filter Berdasarkan Kecamatan`** yang mengelompokkan data berdasarkan kode & nama resmi kecamatan di Kab. Demak (3321).
  - **Sistem Auto-Hide Wilayah Selesai**: Kecamatan, Desa, atau Batch yang seluruh datanya sudah berstatus selesai (di cache lokal maupun spreadsheet) otomatis disembunyikan dari daftar pilihan terminal sehingga pengguna tidak akan memilih wilayah kosong/sudah beres.
  - Menambahkan opsi sub-wilayah saat memilih kecamatan: memproses seluruh desa sekaligus atau memilih desa tertentu yang belum selesai di kecamatan tersebut.
- **Sinkronisasi Live Google Sheets & Pewarnaan Hijau Otomatis (`google_apps_script_sync.js` & `sheet_sync_config.json`)**:
  - Menyediakan script **Google Apps Script Web App** (`google_apps_script_sync.js`) siap pasang untuk Google Spreadsheet.
  - Setiap kali bot selesai memproses baris data di Fasih-SM, bot secara otomatis menambahkan/memperbarui kolom **`Status Pengerjaan`** menjadi **`SELESAI`** dan memberi warna latar **HIJAU LEMBUT (`#d9ead3`)** pada baris tersebut di Google Sheet secara live.
  - Mendukung kolaborasi tim: saat teman Anda menjalankan bot di komputernya, bot otomatis mendeteksi baris yang sudah ditandai selesai di Google Sheet dan melewatinya (skip) tanpa memproses ulang.
  - Dilengkapi sistem antrean aman (`LockService`) pada Apps Script agar tidak terjadi tabrakan data saat banyak bot berjalan serentak.
  - Dual-layer cache: tetap menyimpan cache lokal offline (`processed_tambah_remark.json`) sebagai cadangan aman jika koneksi terputus.
- **Isolasi Log & Cache Progres Lokal (`.gitignore`)**:
  - Menambahkan `processed_tambah_remark.json` dan `laporan_tambah_remark.json` ke `.gitignore` serta menghapusnya dari git tracking (`git rm --cached`).
  - Memastikan seluruh histori pengerjaan dan log rekap bot tetap tersimpan aman di komputer lokal tanpa ter-push ke repository.
- **Perbaikan Kritis Menu BLOK II, Scoping 27.a vs 26.f, dan Feedback Submit (`tambah_remark_27a.py`)**:
  - **Fix Buka Menu BLOK II**: Menghapus pre-check palsu yang sebelumnya langsung menganggap BLOK II aktif karena teks sidebar. Sekarang bot selalu mengklik menu `BLOK II` secara eksplisit dan mengekspansi accordion `SE2026 - L` jika tertutup.
  - **Fix Isolasi Catatan 27.a (Anti Pertanyaan 26.f)**: Memperbaiki pemilihan tombol Catatan agar tidak pernah mengklik Pertanyaan `26. f. Total pengeluaran (a+b+c+d+e)`. Menambahkan filter arah vertikal ketat (`top >= targetRect.top - 10`) sehingga tombol yang terletak di atas 27.a (milik 26.f) ditolak mentah-mentah, serta memfilter container agar tidak mengandung teks `26.` atau `total pengeluaran`.
  - **Fix Feedback Kirim / Submit Kuesioner**: Memperbaiki `wait_for_submit_finish` agar tidak keluar sebelum server memberikan feedback. Bot kini menunggu request network selesai (*networkidle*), memantau redirect keluar dari `/edit`, dan menunggu toast notifikasi sukses atau perubahan status formulir sebelum berganti ke assignment berikutnya.
- **Penanganan Loading Submit / Konfirmasi Kirim Kuesioner (`tambah_remark_27a.py`)**:
  - Menambahkan fungsi `wait_for_submit_finish` untuk menangani proses loading saat menekan tombol konfirmasi submit kuesioner (*"Ya, kirim"* / *"Konfirmasi"*).
  - Bot secara dinamis menunggu hingga respon server selesai, loading spinner pengiriman hilang, notifikasi sukses muncul, atau halaman dialihkan (*redirect*).
  - Mencegah bot terburu-buru berpindah ke assignment berikutnya sebelum server Fasih-SM selesai menyimpan data kiriman formulir.
- **Optimasi Deteksi Kuesioner Siap & Penghilangan False Positive Overlay (`tambah_remark_27a.py`)**:
  - Memperbaiki bug false positive deteksi overlay backdrop (`.fixed.inset-0` + `svg`) yang menyebabkan bot macet berputar lama menunggu padahal formulir kuesioner sudah tampil di layar.
  - Mengimplementasikan `wait_for_kuesioner_ready` khusus saat awal membuka link assignment `/edit`. Bot secara cerdas mendeteksi kemunculan elemen kuesioner (`SE2026`, `BLOK`, atau navigasi sidebar). Begitu kuesioner ter-render dan loading splash screen hilang, bot langsung melanjutkan proses seketika tanpa menunggu timeout.
  - Mengubah mode navigasi `page.goto` ke `domcontentloaded` untuk mencegah hanging akibat request background yang lambat.
  - Menghapus pengecekan loading berulang yang tidak perlu di dalam menu BLOK II, pemilihan kartu usaha, dialog popover, dan loop scroll pemindaian pertanyaan 27.a sehingga bot bekerja jauh lebih cepat dan responsif.
- **Deteksi Loading Dinamis & Penanganan Website Lemot (`tambah_remark_27a.py`)**:
  - Menambahkan fungsi `wait_for_loading_finish` yang secara otomatis mendeteksi animasi spinner (`animate-spin`), skeleton, shimmer, teks *'Memuat...' / 'Loading...'*, maupun overlay backdrop.
  - Bot secara adaptif menunggu hingga proses pengambilan data dari server Fasih-SM selesai sepenuhnya sebelum mencoba berinteraksi dengan menu, kartu usaha, atau pertanyaan formulir.
  - Mengatasi kendala koneksi/server lemot di mana rincian pertanyaan 27.a sebelumnya belum selesai ter-render saat bot melakukan pemindaian.
  - Menambahkan sistem *two-pass scanning* (hingga 35 langkah scroll dan pemindaian ulang dari atas) untuk memastikan pertanyaan 27.a yang dimuat secara asinkron/lazy-load dapat ditemukan dengan sempurna.
- **Perbaikan Scoping Tombol Catatan & Navigasi Pertanyaan 27.a (`tambah_remark_27a.py`)**:
  - Memperbaiki bug di mana bot sebelumnya mengklik tombol Catatan pada Pertanyaan 8 (seperti 8.d Jenis Kawasan) karena scoping selector container luas yang salah mencocokkan tombol pertama.
  - Menambahkan scanning scroll aktif bertahap (hingga 24 langkah) untuk memastikan form/modal rincian digulir sampai Pertanyaan **27. a. Nilai produksi/pendapatan/penjualan barang dan jasa** benar-benar termuat dan berada di tengah layar.
  - Menerapkan penguncian seleksi tombol Catatan berbasis isolasi kontainer dan kedekatan geometris vertikal (*spatial proximity* < 120px dari label 27.a), sehingga 100% mustahil salah mengklik tombol pertanyaan lain.
  - Memperbaiki alur ekspansi sidebar `SE2026 - L` -> `BLOK II` dan pemilihan kartu usaha via tombol `Lihat Rincian`.
  - Menambahkan verifikasi penyimpanan catatan di dalam popover sebelum ditutup agar catatan `#0` terkonfirmasi tersimpan di server Fasih-SM.
- **Dukungan Multi-Sumber Google Sheets (`tambah_remark_27a.py`)**:
  - Menambahkan dukungan sumber data baru: Google Sheet **`Dtata_NTB`** (gid: `293925488`).
  - Menambahkan menu pilihan sumber data pada saat bot dijalankan:
    - `[1] Dtata_NTB` *(Default)*
    - `[2] Data detail anomali`
    - `[3] GABUNGKAN SEMUA SHEET (Dtata_NTB + Data detail anomali)`
  - Kolom `Sumber Sheet` otomatis dicatat pada rekap laporan Excel `laporan_tambah_remark.xlsx`.
- **Bot Baru: Tambah Remark 27.a Nilai Produksi (`tambah_remark_27a.py` & `run_tambah_remark.bat`)**:
  - Mengambil data anomali live secara otomatis langsung dari Google Sheets (*Dtata_NTB* dan *Data detail anomali*) dengan default filter kolom `Petugas Cleaning = 0`.
  - Dilengkapi menu interaktif di terminal untuk filter pengerjaan (Semua data, filter berdasarkan ID Desa, filter berdasarkan Batch, dsb).
  - Alur otomatisasi:
    1. Membuka link assignment Fasih-SM (`/edit`).
    2. Masuk ke menu sidebar **SE2026 - L BLOK II: KETERANGAN USAHA/PERUSAHAAN**.
    3. Mencocokkan nama usaha pada data sheet dengan submenu di sidebar maupun kartu usaha di form.
    4. Mengklik tombol Catatan (speech bubble) pada pertanyaan **27. a. Nilai produksi/pendapatan/penjualan barang dan jasa**.
    5. Mengisikan teks remark **`#0`**, menyimpan catatan, dan menutup popup.
    6. Mengklik tombol simpan pada modal rincian jika ada.
    7. Masuk ke menu **Catatan** di sidebar.
    8. Memastikan switch/checkbox **Tampilkan Anomali Usaha dan Keluarga** bernilai **TRUE**.
    9. Mengklik tombol **Kirim** dan menyelesaikan konfirmasi submit bertingkat (konfirm 2x / Submit Paksa).
    10. Lanjut ke assignment berikutnya secara mulus.
  - Dilengkapi cache resume (`processed_tambah_remark.json`), laporan rekap otomatis (`laporan_tambah_remark.xlsx`), penanganan sesi terputus/VPN disconnect dengan auto-pause & login ulang, serta pemulihan deteksi bot/WAF.
  - Menambahkan menu pilihan **`[5] Bot Tambah Remark 27.a Nilai Produksi`** pada `run.bat` dan file launcher instan **`run_tambah_remark.bat`**.

### 2026-10-20
- **Penanganan Otomatis Sesi Terputus / VPN Disconnect & Auto-Pause (`ubah_jenis_kawasan.py`)**:
  - Menambahkan deteksi otomatis gangguan jaringan, sesi logout SSO BPS, atau VPN terputus (`ERR_INTERNET_DISCONNECTED`, `Checking the connection`, redirect ke halaman login, atau halaman blank).
  - Ketika gangguan terjadi, bot **TIDAK** memasukkan link tersebut ke cache (`processed_jenis_kawasan.json`) dan **TIDAK** mencatatnya sebagai kegagalan permanen.
  - Bot secara otomatis **menghentikan sementara (pause)** proses dan memunculkan notifikasi di terminal agar pengguna dapat memeriksa koneksi VPN dan login ulang di browser.
  - Setelah pengguna selesai login ulang dan menekan `ENTER`, bot langsung melanjutkan dan mencoba ulang (*retry*) link tersebut tanpa ada data yang terlewat.
  - Membersihkan pencatatan gagal sementara dari file laporan saat terjadi pemutusan koneksi.
- **Penambahan Alur Remark / Catatan Pertanyaan 8.d & Presisi Switch Anomali (`ubah_jenis_kawasan.py`)**:
  - Menambahkan penanganan otomatis pengisian catatan/remark pada rincian pertanyaan **8. d. Jenis kawasan beroperasi**:
    1. Mengklik tombol ikon catatan (`button[title="Catatan"]`).
    2. Menuliskan catatan `#1` pada field catatan yang muncul.
    3. Mengklik tombol Simpan (`button[title="Simpan"]` ikon pesawat kertas).
    4. Menutup dialog catatan.
  - Memperbarui pendeteksian dan aktivasi switch **Tampilkan Anomali Usaha dan Keluarga** pada menu Catatan menggunakan struktur DOM switch (`input[role="switch"]` dan `id$="-control"` dengan pengecekan `aria-checked="true"`).
  - Menyempurnakan selector tombol oranye **Lihat Rincian** dengan penargetan spesifik elemen `<button>` (`no_viewport=True` dan `force=True`) agar tidak terhalang tampilan responsive Tailwind `tw:sm:flex`.

### 2026-10-19
- **Fitur Bot Ubah Jenis Kawasan ke '10. Diluar kawasan' (`ubah_jenis_kawasan.py`)**:
  - Menambahkan script otomasi baru `ubah_jenis_kawasan.py` untuk menangani anomali isian kawasan pada data SE2026 (seperti file `sqllab_untitled_query_17_20261019T113729.xlsx`).
  - Mengotomatisasi perubahan rincian pertanyaan **8. d. Jenis kawasan beroperasi** dari kode **1. Kawasan Ekonomi Khusus (KEK)** menjadi kode **10. Diluar kawasan** (atau Di luar kawasan).
  - Alur otomatisasi:
    1. Membuka link assignment `/edit`.
    2. Membuka menu **SE2026 - L BLOK II: KETERANGAN USAHA/PERUSAHAAN**.
    3. Mengklik tombol **Lihat Rincian** pada usaha yang sesuai.
    4. Melakukan scroll bertahap ke bawah pada form/modal rincian agar blok pertanyaan 8 termuat sepenuhnya di DOM.
    5. Mencari pertanyaan 8.d dan memilih opsi `10. Diluar kawasan` (mendukung radio button, native select, maupun custom combobox/dropdown).
    6. Membersihkan field *Nama kawasan* (8.e) jika masih terisi/aktif.
    7. Menyimpan rincian usaha.
    8. Masuk ke menu **Catatan** dan memastikan checkbox **Tampilkan Anomali Usaha dan Keluarga** bernilai `true` (tercentang).
    9. Mengklik tombol **Kirim** / **Submit** dan menyelesaikan konfirmasi modal bertingkat (konfirm 2x).
  - Dilengkapi fitur pemilihan filter wilayah Kecamatan interaktif (misal Bonang, Sayung, atau Semua).
  - Dilengkapi cache progres mandiri `processed_jenis_kawasan.json` agar pengerjaan yang terputus dapat dilanjutkan tanpa mengulang link yang sukses.
  - Dilengkapi pencatatan laporan rekap Excel `laporan_jenis_kawasan.xlsx` dan JSON `laporan_jenis_kawasan.json`.
  - Menambahkan launcher khusus **`run_ubah_kawasan.bat`** serta mengintegrasikan opsi `[4] Bot Ubah Jenis Kawasan ke Diluar Kawasan` pada menu utama **`run.bat`**.

### 2026-09-28
- **Pemisahan Folder Data Otomatis (`data/Anomali` & `data/BKU Ditautkan`)**:
  - Bot Anomali (`reject_anomali.py` & `edit_anomali_admin.py`) kini otomatis memprioritaskan membaca file dari folder `data/Anomali` dan mengabaikan file BKU serta file laporan Excel.
  - Bot Ubah Status Ditemukan (`ubah_status_ditemukan.py`) otomatis memprioritaskan membaca file dari folder `data/BKU Ditautkan` dan mengabaikan file anomali serta file laporan Excel.
  - Memperbarui seluruh runner (`run.bat`, `run_edit_admin.bat`, `run_status_ditemukan.bat`) agar mendeteksi ketersediaan file Excel secara rekursif di seluruh subfolder dan mengorganisir penempatan file otomatis.
- **Fitur Bot Edit Anomali by Admin (`edit_anomali_admin.py`)**:
  - Mengubah alur penanganan anomali dari yang semula me-reject tugas menjadi langsung konfirmasi dan penyelesaian oleh Admin.
  - Alur otomatisasi:
    1. Membuka link assignment `/edit`.
    2. Masuk ke menu **Catatan** dan mencentang checkbox **Tampilkan Anomali Usaha dan Keluarga** (langsung lanjut tanpa submit).
    3. Memeriksa menu baru yang muncul: **Anomali Usaha** dan **Anomali Keluarga**.
    4. Mendeteksi apakah ada anomali yang perlu ditindaklanjuti.
    5. Mencentang opsi **Ya, Sesuai Kondisi Lapangan** untuk setiap anomali yang muncul.
    6. Mengisi field **Penjelasan Anomali** dengan teks: *"sudah dikonfirmasi ke petugas lapangan. dan sudah sesuai kondisi lapangan"*.
    7. Setelah seluruh menu anomali selesai diperiksa dan diisi, baru mengklik tombol **Kirim** dan menyelesaikan konfirmasi submit satu kali di akhir.
  - Memperbarui `run.bat` dengan opsi default [2] Edit by Admin (`1. reject, 2. edit by admin [default], 3. ubah status ditemukan`).
  - Menambahkan launcher khusus `run_edit_admin.bat` untuk eksekusi langsung bot edit by admin.
  - Menyimpan cache progres secara mandiri pada `processed_edit_admin.json`.

### 2026-09-30
- **Dukungan Multi-File BKU Ditautkan & Pemrosesan Gabungan**:
  - Memperbarui sistem pemilihan file pada bot `ubah_status_ditemukan.py`: menghilangkan pembatasan hardcoded yang sebelumnya hanya memilih file 'Status Ganda'.
  - Menambahkan opsi default `[0] SEMUA FILE BKU` untuk menggabungkan seluruh file BKU sekaligus (Status Tutup, Status Ganda, dan Tidak Ditemukan) dengan deduplikasi link otomatis.
  - Tetap menyediakan opsi interaktif untuk memilih file individual jika diinginkan.
- **Filter Data BKU Ditautkan Khusus Kabupaten Demak (3321)**:
  - Menyaring seluruh data pada file `11. BKU Ditautkan tapi Status Tutup`, `12. BKU Ditautkan tapi Status Ganda`, dan `BKU ditautkan tp tdk ditemukan` khusus untuk wilayah Kabupaten Demak (kode `3321`).
  - Mengamankan file data provinsi awal ke dalam subfolder `data/BKU Ditautkan/_semua_prov_backup/`.
- **Rekap Otomatis Usaha Sudah Diganti (`laporan_usaha_sudah_diganti.xlsx`)**:
  - Menambahkan file laporan Excel khusus `laporan_usaha_sudah_diganti.xlsx` (dan `.json`) untuk merekap seluruh kasus penautan yang gagal diubah statusnya karena usaha sudah diganti pada keluarga tersebut.
  - Memuat metadata lengkap dari data BKU (Assignment ID, Level 6 Full Code, ID SBR Usaha, Nama Bangunan/Usaha, Status Awal, Keluarga Ditautkan Awal, Keluarga Pengganti di Fasih, Pesan Peringatan Lengkap, Link Fasih, dan File Sumber).
  - Mengikutsertakan seluruh data yang telah diproses sebelumnya (36 data awal) serta memperbarui laporan secara *real-time* saat bot sedang berjalan.
- **Penyederhanaan Log & Akselerasi Skip Modal Titik Tiga**:
  - Menghilangkan log dump diagnostik tombol DOM yang panjang saat tombol titik tiga tidak ditemukan, diganti dengan pesan status bersih dan ringkas.
  - Menambahkan selector presisi `border-l` (split button di samping tombol Kirim) dan memangkas waktu percobaan pencarian menjadi maksimal 2x percobaan cepat.
  - Menangani kondisi ketiadaan menu 'Submit Paksa' sebagai status `SKIP - TIDAK ADA SUBMIT PAKSA` yang otomatis tersimpan di cache dan laporan tanpa memicu crash/error berulang.
  - Mempercepat jeda (*delay*) perpindahan link untuk data yang di-skip menjadi 1.5 - 3 detik (sebelumnya 6 - 11 detik).
- **Fix Deteksi Tombol 'Submit Paksa' Non-Modal & Penghapusan Asumsi Dialog**:
  - Memperbaiki alur penemuan tombol **Submit Paksa** agar memeriksa langsung pada halaman/form (non-modal), tombol split dropdown (titik tiga), maupun setelah tombol Kirim diklik.
  - Menghilangkan bug `offsetParent !== null` pada JavaScript yang sebelumnya menyebabkan elemen ber-style `position: fixed` (seperti floating action bar Fasih-SM dan popover Radix UI) terabaikan dan dianggap tidak terlihat (*invisible*).
  - Mengganti deteksi keterlihatan elemen dengan `checkVisibility()` standar W3C dan evaluasi *computed style* + *bounding client rect* yang akurat di browser Chromium.
  - Menambahkan pemulihan otomatis cache progres dan laporan: 692 data yang sebelumnya keliru di-skip sebagai `SKIP - TIDAK ADA SUBMIT PAKSA` otomatis dipulihkan ke antrean aktif (`Sisa Diproses`) agar dapat diproses kembali hingga tuntas.
- **Deteksi Cepat & Skip Instan Usaha Sudah Diganti pada Menu SE2026 - P**:
  - Mengintegrasikan pengecekan peringatan "Usaha ini pernah ditautkan... sudah diganti" langsung saat halaman menu `SE2026 - P` terbuka, selama proses render, dan sebelum memilih radio atau mengklik Kirim.
  - Memanfaatkan Playwright locator DOM langsung (`text=/pernah ditautkan/i`) yang bekerja instan (<50ms) dan andal mendeteksi alert banner.
  - Meniadakan langkah memilih radio, klik Kirim, dan menunggu modal titik tiga pada assignment yang usahanya sudah diganti sehingga langsung di-skip seketika dan tercatat rapi ke Excel.
- **Fix Unpack Variable & Inisialisasi Status pada Main Loop**:
  - Menyelaraskan seluruh pemanggilan `check_usaha_sudah_diganti` (pada halaman awal dan modal) agar selalu menerima 3 nilai kembalian (`is_diganti, msg_diganti, fam_diganti`) guna mencegah `ValueError: too many values to unpack`.
  - Menginisialisasi variabel `status = None` sebelum blok `try` pada *loop* utama sehingga penentuan jeda aman dari `UnboundLocalError`.

### 2026-09-23
- **Fix Force Close Launcher Batch & Kompatibilitas Multi-PC**:
  - Memperbaiki parsing karakter khusus `&` pada baris menu `run.bat` dan `title` pada `run_status_ditemukan.bat` yang sebelumnya menyebabkan command prompt langsung *force close* (crash) seketika saat dibuka di komputer lain.
  - Memperbaiki penanganan validasi virtual environment `(venv)` menggunakan label alur `goto` agar aman dari error sintaks parser batch Windows CMD (`... was unexpected at this time`).
  - Menambahkan pengecekan otomatis validitas virtual environment: jika folder dicopy dari PC lain dengan path Python berbeda, runner akan otomatis membersihkan dan membuat ulang `venv` yang kompatibel dengan PC tersebut.
- **Eliminasi Ketergantungan Pandas & Standalone Auto-Installer**:
  - Menghapus ketergantungan `pandas` pada `ubah_status_ditemukan.py` dan menggantinya dengan `openpyxl` murni untuk pembacaan data dan pembuatan laporan rekapitulasi Excel.
  - Menambahkan auto-installer dependencies langsung pada script Python (`ubah_status_ditemukan.py`) sehingga script dapat dieksekusi secara mandiri/standalone di terminal maupun VS Code.
- **Deteksi Peringatan Usaha Sudah Diganti & Auto-Reporting**:
  - Menambahkan deteksi otomatis peringatan bahwa usaha pernah ditautkan namun sudah diganti pada keluarga tersebut sehingga assignment tidak ditemukan lagi.
  - Melewati (skip) assignment yang memuat peringatan tersebut secara otomatis tanpa membuang waktu mencoba mengubah status atau submit paksa.
  - Menyimpan data riwayat pemrosesan lengkap ke file rekap laporan **`laporan_status_ditemukan.xlsx`** dan **`laporan_status_ditemukan.json`** yang mencatat status "BERHASIL", "SKIP - USAHA SUDAH DIGANTI", dan "SKIP - BUKAN OTORISASI" beserta pesan sistemnya.
  - Menyimpan link yang di-skip ke cache agar tidak dikerjakan ulang pada eksekusi berikutnya.

### 2026-09-22
- **Bot Ubah Status ke '1. Ditemukan' & Submit Paksa**:
  - Menambahkan script baru `ubah_status_ditemukan.py` untuk mengotomatisasi perubahan status keberadaan bangunan usaha dari data Excel (seperti `12. BKU Ditautkan tapi Status Ganda.xlsx`).
  - Alur otomatisasi: membuka link `/edit`, memilih menu **SE2026 - P**, memilih radio **1. Ditemukan** pada pertanyaan *Keberadaan Bangunan Lainnya/ Usaha*, mengklik tombol **Kirim**, mengklik tombol split dropdown (titik tiga), dan memilih **Submit Paksa**.
  - Dilengkapi fitur pembacaan data Excel streaming cepat (`read_only=True`) dengan kemampuan mendeteksi filter Excel (AutoFilter / hidden rows) secara instan via XML parser (< 1 detik). Baris yang disembunyikan/di-filter di Excel otomatis dilewati.
  - Menambahkan fitur filter interaktif per Kabupaten/Kota (`kab`) atau memproses seluruh data.
  - Menyimpan cache progres secara mandiri pada `processed_status_ditemukan.json` sehingga pengerjaan dapat dilanjutkan kapan saja tanpa mengulang link yang sudah selesai.
  - Menambahkan launcher khusus `run_status_ditemukan.bat` dan memperbarui `run.bat` dengan menu pilihan interaktif untuk memilih bot yang ingin dijalankan.
  - **Fix Klik Menu SE2026 - P & Pemilihan Radio 1. Ditemukan**: Memperbaiki selector menu tab agar tidak mengklik generic container `div`, menggunakan koordinat mouse click dan JS deep event simulation, menambahkan verifikasi render form sebelum memilih opsi, serta pencarian radio input `value="1"` langsung.
  - **Fix Tombol Titik Tiga Modal & Submit Paksa**: Mengganti event click tombol split dropdown dengan Playwright native mouse click dan simulasi event `pointerdown` + `mousedown` lengkap (dibutuhkan komponen Radix UI/Tailwind) serta verifikasi terbukanya menu item 'Submit Paksa'.

### 2026-08-28
- **Rekapan Rincian Per Kecamatan**: Menambahkan ringkasan statistik per kecamatan sebelum browser dibuka yang menampilkan kolom Nomor, Nama Kecamatan, Total Link, Jumlah Sudah Diproses (Cache), Sisa yang Akan Diproses, serta Persentase Progress pengerjaan.
- **Deteksi Dinamis Kolom Kecamatan**: Memperbarui parser Excel agar secara otomatis mendeteksi kolom `Nama Kecamatan` / `Kecamatan`.
- **Indikator Kecamatan pada Log Pemrosesan**: Menyematkan nama kecamatan pada setiap baris log pemrosesan link (contoh: `[1/111] [MRANGGEN] Memproses: ...`).

### 2026-08-06
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
