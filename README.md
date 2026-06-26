# Anomali-SE2026

## Changelog
- **2026-06-26**: Mengubah engine database dari `mysql` menjadi `sqlite` pada konfigurasi untuk menyesuaikan keadaan server yang tidak dapat diakses.
- **2026-06-26**: Optimasi penarikan data - setiap level (kecamatan, desa, SLS) kini hanya fetch ke child jika ada `total_value > 0`, sehingga ribuan request ke sub-SLS yang tidak ada anomalinya dapat dihindari.
