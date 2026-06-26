# Anomali-SE2026

## Changelog
- **2026-06-26**: Mengubah engine database dari `mysql` menjadi `sqlite` pada konfigurasi untuk menyesuaikan keadaan server yang tidak dapat diakses.
- **2026-06-26**: Optimasi penarikan data - setiap level (kecamatan, desa, SLS) kini hanya fetch ke child jika ada `total_value > 0`, sehingga ribuan request ke sub-SLS yang tidak ada anomalinya dapat dihindari.
- **2026-06-26**: Tambah cache harian berbasis file JSON (`cache_YYYY-MM-DD.json`). Script kini bisa dilanjutkan (resume) setelah dihentikan — URL yang sudah di-fetch sebelumnya langsung diambil dari cache tanpa delay dan tanpa hit API ulang.
- **2026-06-26**: Tambah `lihat_anomali.py` — script CLI interaktif untuk melihat hasil tarik data: ringkasan per kecamatan, detail per kecamatan, daftar SLS beranoali, rekap per jenis anomali, dan daftar kasus mikro.
- **2026-06-26**: Tambah deteksi session expired — kalau muncul halaman HTML 3x berturut-turut, script otomatis pause dan minta user login ulang di browser sebelum melanjutkan.
