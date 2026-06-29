# Anomali SE2026

**Anomali SE2026** adalah sebuah tool penarik dan penampil (dashboard) data anomali Sensus Ekonomi 2026. Aplikasi ini secara otomatis mengambil data dari dashboard BPS, menelusuri dari level kecamatan hingga level usaha/mikro, serta menyediakan antarmuka web yang bersih untuk memonitor daftar usaha anomali.

## Fitur Utama
1. **Penarikan Data Berjenjang**: Otomatis menarik data anomali dari level Kecamatan -> Desa -> SLS -> Sub SLS -> Kasus Mikro.
2. **Resume-Friendly Cache**: Menggunakan cache harian JSON sehingga jika terputus, tarikan bisa dilanjutkan tanpa harus menembak ulang API.
3. **Web Dashboard Interaktif**: Menampilkan data dengan pagination, pencarian, dan tautan langsung ke FASIH SM.
4. **Anti-Session Expired**: Memberitahu Anda untuk login ulang saat sesi di BPS habis, lalu melanjutkan tarikan data tanpa putus.
5. **Indikator Dinamis**: Tambah/kurang jenis anomali cukup dengan mengedit `config.json`.

---

## Prasyarat (*Requirements*)
- **Python 3.8+**
- **Google Chrome Browser**
- Library: `flask`, `requests`, `undetected_chromedriver`, `selenium`

---

## 1. Instalasi dan Setup

1. **Clone repositori** ini:
   ```bash
   git clone https://github.com/muhshi/anomali_se2026.git
   cd anomali_se2026
   ```

2. **Install dependency**:
   ```bash
   pip install flask requests undetected_chromedriver selenium
   ```

3. **Konfigurasi Database & API**:
   Copy file `config.example.json` menjadi `config.json` dan sesuaikan kredensial Anda:
   ```bash
   cp config.example.json config.json
   ```
   *(Pastikan untuk mengedit file `config.json` dan menyesuaikan konfigurasi database jika Anda menggunakan MySQL, atau biarkan default jika menggunakan SQLite)*.

---

## 2. Cara Menjalankan Program

Cara termudah untuk menjalankan Dashboard Web dan Penarik Data secara bersamaan adalah menggunakan script runner yang telah disediakan.

### **Pengguna Windows**
Klik ganda (`Double Click`) pada file **`run.bat`** (atau jalankan `run.bat` di terminal/CMD).

### **Pengguna Mac/Linux**
Buka Terminal, masuk ke direktori folder ini, lalu jalankan:
```bash
chmod +x run.sh
./run.sh
```

**Alur Kerja Runner:**
1. Menjalankan Dashboard Web (`web_anomali.py`) di latar belakang (dapat diakses di `http://localhost:5050`).
2. Membuka browser Chrome otomatis untuk Anda login ke web BPS.
3. Menjalankan script penarik data (`tarik_anomali.py`) di terminal depan.

*(Catatan: Anda juga bisa menjalankan `python web_anomali.py` dan `python tarik_anomali.py` di dua terminal yang berbeda secara manual jika diperlukan).*

## Changelog
- **2026-06-26**: Mengubah engine database dari `mysql` menjadi `sqlite` pada konfigurasi untuk menyesuaikan keadaan server yang tidak dapat diakses.
- **2026-06-26**: Optimasi penarikan data - setiap level (kecamatan, desa, SLS) kini hanya fetch ke child jika ada `total_value > 0`, sehingga ribuan request ke sub-SLS yang tidak ada anomalinya dapat dihindari.
- **2026-06-26**: Tambah cache harian berbasis file JSON (`cache_YYYY-MM-DD.json`). Script kini bisa dilanjutkan (resume) setelah dihentikan — URL yang sudah di-fetch sebelumnya langsung diambil dari cache tanpa delay dan tanpa hit API ulang.
- **2026-06-26**: Tambah `lihat_anomali.py` — script CLI interaktif untuk melihat hasil tarik data: ringkasan per kecamatan, detail per kecamatan, daftar SLS beranoali, rekap per jenis anomali, dan daftar kasus mikro.
- **2026-06-26**: Tambah deteksi session expired — kalau muncul halaman HTML 3x berturut-turut, script otomatis pause dan minta user login ulang di browser sebelum melanjutkan.
- **2026-06-26**: Pembaruan UI Dashboard Anomali dengan tema terang (light mode) beraksen oranye ala Sensus Ekonomi, pagination, dan filter terpusat di Dashboard.
- **2026-06-26**: Implementasi penyimpanan `raw_data` JSON pada `kasus_anomali_mikro` untuk menangkap detail nama KRT/Usaha di Web, serta pembaruan format ID kecamatan menjadi `[KODE] NAMA`.
- **2026-06-26**: Perbaikan silent error (HTTP 500) pada API BPS dengan menyisipkan parameter `type` dan `anomali_no` secara otomatis berdasarkan jenis indikator.
- **2026-06-26**: Dinamisasi konfigurasi — daftar indikator dipindahkan dari hardcode Python ke `config.json` agar mudah ditambah jika pusat mengeluarkan indikator baru.
- **2026-06-26**: Menambahkan `run.bat` dan `run.sh` untuk mempermudah eksekusi script.
- **2026-06-29**: Perbaikan `UnicodeDecodeError` saat memuat `config.json` di Windows dengan menambahkan `encoding='utf-8'` secara eksplisit.

