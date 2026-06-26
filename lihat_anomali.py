import sqlite3
import os
import sys
from datetime import datetime

# ========================================
# Lihat Hasil Tarik Data Anomali SE2026
# Jalankan: python lihat_anomali.py
# atau:      python lihat_anomali.py [tanggal]  (format: YYYY-MM-DD)
# ========================================

DB_FILE = 'anomali_se2026.db'

# Mapping kode indikator → nama pendek
NAMA_INDIKATOR = {
    "40":  "A1-Biaya Produksi Dominan (✓ selesai)",
    "128": "A1-Biaya Produksi Dominan (✗ belum)",
    "41":  "A2-Keuntungan Usaha (✓ selesai)",
    "129": "A2-Keuntungan Usaha (✗ belum)",
    "42":  "A3-Modal Korporasi Non-Badan Usaha (✓ selesai)",
    "130": "A3-Modal Korporasi Non-Badan Usaha (✗ belum)",
    "43":  "A4-Data Keuangan MBG (✓ selesai)",
    "131": "A4-Data Keuangan MBG (✗ belum)",
    "44":  "A5-Omset Didominasi Satu Jenis (✓ selesai)",
    "132": "A5-Omset Didominasi Satu Jenis (✗ belum)",
    "45":  "A6-Tenaga Kerja Tidak Dibayar (✓ selesai)",
    "133": "A6-Tenaga Kerja Tidak Dibayar (✗ belum)",
    "46":  "A7-Pengeluaran Tidak Wajar (✓ selesai)",
    "134": "A7-Pengeluaran Tidak Wajar (✗ belum)",
    "135": "A8-Anomali Lainnya (✗ belum)",
}

def conn_db():
    if not os.path.exists(DB_FILE):
        print(f"[!] File database tidak ditemukan: {DB_FILE}")
        sys.exit(1)
    return sqlite3.connect(DB_FILE)

def nama_ind(kode):
    return NAMA_INDIKATOR.get(str(kode), f"Indikator {kode}")

def get_tanggal(con):
    cur = con.cursor()
    cur.execute("SELECT DISTINCT tanggal_tarik FROM agregat_anomali ORDER BY tanggal_tarik DESC")
    rows = cur.fetchall()
    return [r[0] for r in rows]

def sep(char='=', n=70):
    print(char * n)

def header(judul):
    sep()
    print(f"  {judul}")
    sep()

# ─────────────────────────────────────────
# 1. RINGKASAN PER KECAMATAN
# ─────────────────────────────────────────
def ringkasan_kecamatan(con, tgl):
    cur = con.cursor()
    cur.execute("""
        SELECT nama_wilayah, kode_indikator, total_value
        FROM agregat_anomali
        WHERE tanggal_tarik = ? AND level_wilayah = 'kecamatan' AND total_value > 0
        ORDER BY nama_wilayah, kode_indikator
    """, (tgl,))
    rows = cur.fetchall()

    if not rows:
        print("  (tidak ada data kecamatan)")
        return

    from collections import defaultdict
    per_kec = defaultdict(dict)
    for nama, kode, val in rows:
        per_kec[nama][kode] = val

    # Total per kecamatan (jumlah semua indikator)
    print(f"\n{'Kecamatan':<20} {'Total Anomali':>14}  Indikator Aktif")
    print("-" * 70)
    grand_total = 0
    for kec in sorted(per_kec):
        total = sum(per_kec[kec].values())
        grand_total += total
        inds = ", ".join([f"Ind{k}={v}" for k, v in sorted(per_kec[kec].items())])
        print(f"  {kec:<18} {total:>14,}  {inds}")
    print("-" * 70)
    print(f"  {'TOTAL':<18} {grand_total:>14,}")

# ─────────────────────────────────────────
# 2. DETAIL KECAMATAN TERTENTU (Sub-SLS)
# ─────────────────────────────────────────
def detail_kecamatan(con, tgl, kode_kec_prefix):
    cur = con.cursor()
    # Cari wilayah yang id-nya berawalan kode kecamatan (7 digit pertama)
    cur.execute("""
        SELECT id_wilayah, nama_wilayah, level_wilayah, kode_indikator, total_value
        FROM agregat_anomali
        WHERE tanggal_tarik = ? 
          AND id_wilayah LIKE ?
          AND total_value > 0
        ORDER BY level_wilayah, id_wilayah, kode_indikator
    """, (tgl, f"{kode_kec_prefix}%"))
    rows = cur.fetchall()

    if not rows:
        print(f"  (tidak ada data untuk kecamatan {kode_kec_prefix})")
        return

    cur_level = None
    for id_w, nama, level, kode, val in rows:
        if level != cur_level:
            print(f"\n  [{level.upper()}]")
            cur_level = level
        print(f"    {nama:<30} {id_w}  |  {nama_ind(kode):<40}  = {val:,}")

# ─────────────────────────────────────────
# 3. DAFTAR SLS/SUB-SLS DENGAN ANOMALI
# ─────────────────────────────────────────
def daftar_sls_anomali(con, tgl):
    cur = con.cursor()
    cur.execute("""
        SELECT id_wilayah, nama_wilayah, kode_indikator, SUM(total_value) as jml
        FROM agregat_anomali
        WHERE tanggal_tarik = ? 
          AND level_wilayah IN ('sls', 'sub_sls')
          AND total_value > 0
        GROUP BY id_wilayah, kode_indikator
        ORDER BY jml DESC, id_wilayah
        LIMIT 50
    """, (tgl,))
    rows = cur.fetchall()

    if not rows:
        print("  (tidak ada data SLS/sub-SLS)")
        return

    print(f"\n{'Kode Wilayah':<20} {'Nama':<25} {'Indikator':<42} {'Jml':>5}")
    print("-" * 96)
    for id_w, nama, kode, jml in rows:
        print(f"  {id_w:<18} {(nama or '-'):<25} {nama_ind(kode):<42} {jml:>5,}")

# ─────────────────────────────────────────
# 4. REKAP JENIS ANOMALI (SEMUA WILAYAH)
# ─────────────────────────────────────────
def rekap_per_indikator(con, tgl):
    cur = con.cursor()
    cur.execute("""
        SELECT kode_indikator, SUM(total_value)
        FROM agregat_anomali
        WHERE tanggal_tarik = ? AND level_wilayah = 'kecamatan' AND total_value > 0
        GROUP BY kode_indikator
        ORDER BY CAST(kode_indikator AS INTEGER)
    """, (tgl,))
    rows = cur.fetchall()

    if not rows:
        print("  (tidak ada data)")
        return

    print(f"\n{'Kode':>5}  {'Jenis Anomali':<46} {'Total':>8}")
    print("-" * 65)
    grand = 0
    for kode, total in rows:
        grand += total or 0
        print(f"  {kode:>4}   {nama_ind(kode):<46} {total:>8,}")
    print("-" * 65)
    print(f"  {'':>4}   {'TOTAL':<46} {grand:>8,}")

# ─────────────────────────────────────────
# 5. DAFTAR KASUS MIKRO
# ─────────────────────────────────────────
def daftar_mikro(con, tgl, limit=30):
    cur = con.cursor()
    cur.execute("""
        SELECT kode_wilayah, anomali_no, anomali_title, assignment_id, is_resolved, link_fasih
        FROM kasus_anomali_mikro
        WHERE tanggal_tarik = ?
        ORDER BY kode_wilayah, anomali_no
        LIMIT ?
    """, (tgl, limit))
    rows = cur.fetchall()

    if not rows:
        print("  (tidak ada kasus mikro)")
        return

    cur.execute("SELECT COUNT(*) FROM kasus_anomali_mikro WHERE tanggal_tarik = ?", (tgl,))
    total = cur.fetchone()[0]
    print(f"  Menampilkan {len(rows)} dari {total} kasus mikro\n")

    for kode, ano_no, title, assign_id, resolved, link in rows:
        status = "✓ selesai" if resolved else "✗ belum"
        print(f"  [{status}] {kode}  Anomali-{ano_no}")
        print(f"           {title or '-'}")
        print(f"           Assignment: {assign_id or '-'}")
        if link:
            print(f"           Link: {link}")
        print()

# ─────────────────────────────────────────
# MAIN MENU
# ─────────────────────────────────────────
def main():
    con = conn_db()

    # Tentukan tanggal
    if len(sys.argv) > 1:
        tgl = sys.argv[1]
    else:
        tanggal_list = get_tanggal(con)
        if not tanggal_list:
            print("[!] Database kosong, belum ada data.")
            sys.exit(0)
        tgl = tanggal_list[0]
        if len(tanggal_list) > 1:
            print(f"Tanggal tersedia: {', '.join(tanggal_list)}")
            print(f"Menggunakan tanggal terbaru: {tgl}\n")

    print(f"\n{'='*70}")
    print(f"  LAPORAN ANOMALI SE2026 - Kabupaten DEMAK")
    print(f"  Tanggal data : {tgl}")
    print(f"  Dijalankan   : {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"{'='*70}")

    while True:
        print("\n[MENU]")
        print("  1. Ringkasan anomali per kecamatan")
        print("  2. Detail wilayah per kecamatan (masukkan kode kec)")
        print("  3. Daftar SLS/Sub-SLS dengan anomali (top 50)")
        print("  4. Rekap per jenis anomali (indikator)")
        print("  5. Daftar kasus mikro")
        print("  0. Keluar")
        pilihan = input("\nPilih menu: ").strip()

        if pilihan == '1':
            header(f"RINGKASAN PER KECAMATAN  [{tgl}]")
            ringkasan_kecamatan(con, tgl)

        elif pilihan == '2':
            kode = input("Masukkan 7 digit kode kecamatan (mis: 3321010): ").strip()
            header(f"DETAIL KECAMATAN {kode}  [{tgl}]")
            detail_kecamatan(con, tgl, kode)

        elif pilihan == '3':
            header(f"DAFTAR SLS/SUB-SLS DENGAN ANOMALI  [{tgl}]")
            daftar_sls_anomali(con, tgl)

        elif pilihan == '4':
            header(f"REKAP PER JENIS ANOMALI  [{tgl}]")
            rekap_per_indikator(con, tgl)

        elif pilihan == '5':
            limit_input = input("Tampilkan berapa kasus? [default: 30]: ").strip()
            limit = int(limit_input) if limit_input.isdigit() else 30
            header(f"KASUS MIKRO  [{tgl}]")
            daftar_mikro(con, tgl, limit)

        elif pilihan == '0':
            print("Keluar.")
            break

        else:
            print("Pilihan tidak valid.")

    con.close()

if __name__ == "__main__":
    main()
