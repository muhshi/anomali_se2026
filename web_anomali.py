"""
Web viewer untuk data Anomali SE2026
Jalankan: python web_anomali.py
Buka di browser: http://localhost:5050
"""

import sqlite3
import os
from flask import Flask, render_template_string, request, g
from datetime import datetime

app = Flask(__name__)
DB_FILE = 'anomali_se2026.db'

NAMA_INDIKATOR = {
    "40":  "A1 - Biaya Produksi Dominan",
    "128": "A1 - Biaya Produksi Dominan",
    "41":  "A2 - Keuntungan Usaha",
    "129": "A2 - Keuntungan Usaha",
    "42":  "A3 - Modal Korporasi Non-Badan Usaha",
    "130": "A3 - Modal Korporasi Non-Badan Usaha",
    "43":  "A4 - Data Keuangan MBG",
    "131": "A4 - Data Keuangan MBG",
    "44":  "A5 - Konsistensi Aset/Pekerja/Produksi",
    "132": "A5 - Konsistensi Aset/Pekerja/Produksi",
    "45":  "A6 - UMB Tanpa Internet",
    "133": "A6 - UMB Tanpa Internet",
    "46":  "A7 - UMB Tanpa Laporan Keuangan",
    "134": "A7 - UMB Tanpa Laporan Keuangan",
    "135": "A8 - Anomali Lainnya",
}

# Deskripsi lengkap per jenis anomali
DESKRIPSI_ANOMALI = [
    {
        "kode": "A1", "ind_belum": "128", "ind_sudah": "40",
        "judul": "Biaya Produksi Dominan",
        "deskripsi": "Usaha yang memiliki total biaya produksi lebih besar dari nilai produksi (omset), sehingga secara matematis usaha tersebut selalu merugi. Kondisi ini dicurigai sebagai kesalahan pencatatan karena tidak wajar secara ekonomi.",
        "contoh": "Usaha dengan omset Rp 10 juta/bulan tetapi biaya bahan baku, tenaga kerja, dan operasional mencapai Rp 15 juta/bulan.",
        "icon": "💰"
    },
    {
        "kode": "A2", "ind_belum": "129", "ind_sudah": "41",
        "judul": "Keuntungan Usaha",
        "deskripsi": "Usaha yang melaporkan keuntungan (laba bersih) secara tidak wajar — terlalu tinggi atau negatif (rugi besar) secara konsisten. Anomali ini menandakan kemungkinan kesalahan dalam pengisian data keuangan usaha.",
        "contoh": "Usaha kecil dengan omset Rp 5 juta/bulan tetapi melaporkan keuntungan Rp 50 juta/bulan, atau rugi terus-menerus tanpa penjelasan.",
        "icon": "📈"
    },
    {
        "kode": "A3", "ind_belum": "130", "ind_sudah": "42",
        "judul": "Bukan Badan Usaha tapi Ada Penyertaan Modal Korporasi",
        "deskripsi": "Usaha yang dicatat sebagai usaha perorangan (bukan PT/CV/koperasi dll) tetapi memiliki penyertaan modal dari korporasi. Hal ini secara hukum tidak konsisten dan menunjukkan kemungkinan kesalahan klasifikasi bentuk badan usaha.",
        "contoh": "Usaha dicatat sebagai 'usaha perseorangan' namun ada modal yang bersumber dari perusahaan lain.",
        "icon": "🏢"
    },
    {
        "kode": "A4", "ind_belum": "131", "ind_sudah": "43",
        "judul": "Data Keuangan MBG",
        "deskripsi": "Usaha kategori Menengah-Besar-Gabungan (MBG) yang memiliki isian data keuangan yang tidak konsisten atau tidak wajar, seperti aset, pendapatan, dan pengeluaran yang saling bertentangan.",
        "contoh": "Usaha skala menengah dengan aset ratusan juta tetapi pendapatan hanya jutaan rupiah, atau sebaliknya.",
        "icon": "📊"
    },
    {
        "kode": "A5", "ind_belum": "132", "ind_sudah": "44",
        "judul": "Konsistensi Aset, Pekerja, dan Produksi Usaha",
        "deskripsi": "Ketidaksesuaian antara skala aset usaha, jumlah tenaga kerja, dan nilai produksi. Ketiga variabel ini seharusnya proporsional; anomali terjadi bila salah satu terlalu besar atau kecil dibanding yang lain.",
        "contoh": "Usaha dengan 50 tenaga kerja dan aset Rp 1 miliar tetapi hanya memproduksi barang senilai Rp 2 juta/bulan.",
        "icon": "⚖️"
    },
    {
        "kode": "A6", "ind_belum": "133", "ind_sudah": "45",
        "judul": "Usaha Menengah & Besar Tanpa Internet",
        "deskripsi": "Usaha berskala Menengah atau Besar yang tidak menggunakan internet untuk kegiatan usahanya. Di era digital, kondisi ini tidak wajar untuk usaha skala tersebut dan perlu dikonfirmasi kebenarannya.",
        "contoh": "Pabrik dengan 100 karyawan dan omset miliaran rupiah tetapi mengisi 'tidak menggunakan internet' untuk kegiatan usaha.",
        "icon": "🌐"
    },
    {
        "kode": "A7", "ind_belum": "134", "ind_sudah": "46",
        "judul": "Usaha Menengah & Besar Tanpa Laporan Keuangan",
        "deskripsi": "Usaha berskala Menengah atau Besar yang tidak memiliki laporan keuangan. Secara regulasi dan praktik bisnis, usaha skala ini seharusnya memiliki pencatatan keuangan formal.",
        "contoh": "Perusahaan dengan aset di atas Rp 500 juta atau 20+ tenaga kerja yang mengisi 'tidak memiliki laporan keuangan'.",
        "icon": "📋"
    },
    {
        "kode": "A8", "ind_belum": "135", "ind_sudah": None,
        "judul": "Anomali Lainnya",
        "deskripsi": "Kategori anomali tambahan yang tidak termasuk dalam A1-A7. Mencakup berbagai inkonsistensi data lain yang terdeteksi oleh sistem secara otomatis dan perlu diverifikasi langsung oleh petugas lapangan.",
        "contoh": "Inkonsistensi data yang tidak terklasifikasi secara spesifik, misalnya data usaha yang tiba-tiba berubah drastis dibanding periode sebelumnya.",
        "icon": "❓"
    },
]

# Indikator yang "belum ditindaklanjuti"
BELUM = {"128","129","130","131","132","133","134","135"}
SUDAH = {"40","41","42","43","44","45","46"}

BASE_STYLE = """
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Inter',sans-serif;background:#fdf7f2;color:#1a1a2e;min-height:100vh}
  a{color:#c2440e;text-decoration:none}
  a:hover{text-decoration:underline}
  .nav{background:linear-gradient(135deg,#e85d04,#f48c06);border-bottom:none;padding:14px 32px;display:flex;align-items:center;gap:24px;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(232,93,4,.25)}
  .nav-brand{font-size:16px;font-weight:700;color:#fff;letter-spacing:.3px}
  .nav-brand span{color:#fde68a}
  .nav a{color:rgba(255,255,255,.8);font-size:13px;font-weight:500;padding:6px 12px;border-radius:6px;transition:.15s}
  .nav a:hover,.nav a.active{background:rgba(255,255,255,.2);color:#fff;text-decoration:none}
  .container{max-width:1200px;margin:0 auto;padding:32px 24px}
  h1{font-size:22px;font-weight:700;color:#7c2d12;margin-bottom:4px}
  .subtitle{font-size:13px;color:#9a6a4e;margin-bottom:24px}
  .cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin-bottom:32px}
  .card{background:#fff;border:1px solid #fcd9b6;border-radius:12px;padding:20px;transition:.2s;box-shadow:0 1px 4px rgba(232,93,4,.06)}
  .card:hover{border-color:#f48c06;transform:translateY(-2px);box-shadow:0 4px 16px rgba(232,93,4,.12)}
  .card-label{font-size:12px;color:#9a6a4e;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px}
  .card-value{font-size:28px;font-weight:700;color:#1a1a2e}
  .card-value.red{color:#dc2626}
  .card-value.green{color:#16a34a}
  .card-value.blue{color:#1d4ed8}
  .card-value.orange{color:#ea580c}
  .table-wrap{background:#fff;border:1px solid #fcd9b6;border-radius:12px;overflow:hidden;margin-bottom:24px;box-shadow:0 1px 4px rgba(232,93,4,.06)}
  .table-head{padding:16px 20px;border-bottom:1px solid #fde8d0;display:flex;align-items:center;justify-content:space-between;background:#fff8f3}
  .table-head h2{font-size:14px;font-weight:600;color:#7c2d12}
  table{width:100%;border-collapse:collapse}
  th{padding:12px 16px;text-align:left;font-size:11px;font-weight:600;color:#9a6a4e;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid #fde8d0;background:#fff8f3}
  td{padding:11px 16px;font-size:13px;border-bottom:1px solid #fef3ea;vertical-align:top;color:#2d2d2d}
  tr:last-child td{border-bottom:none}
  tr:hover td{background:#fff8f3}
  .badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600}
  .badge-red{background:#fef2f2;color:#dc2626;border:1px solid #fca5a5}
  .badge-green{background:#f0fdf4;color:#16a34a;border:1px solid #86efac}
  .badge-orange{background:#fff7ed;color:#ea580c;border:1px solid #fdba74}
  .badge-gray{background:#f8fafc;color:#64748b;border:1px solid #e2e8f0}
  .num{font-variant-numeric:tabular-nums;font-weight:600}
  .num-big{font-size:15px;color:#dc2626;font-weight:700}
  .select-bar{display:flex;gap:12px;align-items:center;margin-bottom:24px;flex-wrap:wrap}
  select,input{background:#fff;border:1px solid #fcd9b6;color:#1a1a2e;padding:8px 14px;border-radius:8px;font-size:13px;font-family:inherit}
  select:focus,input:focus{outline:none;border-color:#f48c06;box-shadow:0 0 0 3px rgba(244,140,6,.15)}
  button{background:#e85d04;color:#fff;border:none;padding:9px 18px;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;font-family:inherit;transition:.15s}
  button:hover{background:#c2440e}
  .empty{padding:40px;text-align:center;color:#9a6a4e;font-size:13px}
  .kec-link{color:#c2440e;font-weight:600}
  .progress-bar{height:6px;background:#fde8d0;border-radius:3px;margin-top:6px;overflow:hidden}
  .progress-fill{height:100%;border-radius:3px;background:linear-gradient(90deg,#f97316,#fbbf24);transition:.3s}
  .tag-ind{font-size:11px;color:#9a6a4e}
  .date-sel{display:flex;gap:8px;align-items:center}
  .date-sel label{font-size:12px;color:#9a6a4e}
  .hero-strip{background:linear-gradient(135deg,#fff7ed,#fef3ea);border:1px solid #fcd9b6;border-radius:12px;padding:20px 24px;margin-bottom:28px}
  .hero-title{font-size:18px;font-weight:700;color:#7c2d12}
  .hero-sub{font-size:12px;color:#9a6a4e;margin-top:2px}
</style>
"""


def get_db():
    db = getattr(g, '_database', None)
    if db is None:
        db = g._database = sqlite3.connect(DB_FILE)
        db.row_factory = sqlite3.Row
    return db

@app.teardown_appcontext
def close_connection(exception):
    db = getattr(g, '_database', None)
    if db is not None:
        db.close()

def get_tanggal_list():
    db = get_db()
    rows = db.execute("SELECT DISTINCT tanggal_tarik FROM agregat_anomali ORDER BY tanggal_tarik DESC").fetchall()
    return [r[0] for r in rows]

def nav(active=''):
    return f"""
    <nav class="nav">
      <span class="nav-brand">Anomali <span>SE2026</span> &middot; Demak</span>
      <a href="/" class="{'active' if active=='home' else ''}">Dashboard</a>
      <a href="/kecamatan" class="{'active' if active=='kec' else ''}">Per Kecamatan</a>
      <a href="/sls" class="{'active' if active=='sls' else ''}">SLS/Sub-SLS</a>
      <a href="/indikator" class="{'active' if active=='ind' else ''}">Per Jenis</a>
      <a href="/mikro" class="{'active' if active=='mikro' else ''}">Kasus Mikro</a>
      <a href="/panduan" class="{'active' if active=='panduan' else ''}">Panduan Anomali</a>
    </nav>"""

# ─────────────────────────────────────────
# HOME - DASHBOARD
# ─────────────────────────────────────────
@app.route('/')
def home():
    if not os.path.exists(DB_FILE):
        return "<h2 style='color:white;padding:40px'>Database belum ada. Jalankan tarik_anomali.py dulu.</h2>"
    
    tanggal_list = get_tanggal_list()
    if not tanggal_list:
        return f"{nav('home')}<div class='container'><div class='empty'>Belum ada data di database.</div></div>"
    
    tgl = request.args.get('tgl', tanggal_list[0])
    db = get_db()

    # Total agregat belum + sudah
    r = db.execute("""
        SELECT 
          SUM(CASE WHEN kode_indikator IN ('128','129','130','131','132','133','134','135') THEN total_value ELSE 0 END) as belum,
          SUM(CASE WHEN kode_indikator IN ('40','41','42','43','44','45','46') THEN total_value ELSE 0 END) as sudah,
          COUNT(DISTINCT id_wilayah) as n_wilayah
        FROM agregat_anomali
        WHERE tanggal_tarik=? AND level_wilayah='kecamatan' AND total_value > 0
    """, (tgl,)).fetchone()
    
    total_belum = r['belum'] or 0
    total_sudah = r['sudah'] or 0
    total_kasus = db.execute("SELECT COUNT(*) FROM kasus_anomali_mikro WHERE tanggal_tarik=?", (tgl,)).fetchone()[0]
    pct = round(total_sudah / (total_sudah + total_belum) * 100, 1) if (total_sudah + total_belum) > 0 else 0

    # Top 5 kecamatan
    top_kec = db.execute("""
        SELECT nama_wilayah, SUM(total_value) as total
        FROM agregat_anomali
        WHERE tanggal_tarik=? AND level_wilayah='kecamatan'
          AND kode_indikator IN ('128','129','130','131','132','133','134','135')
        GROUP BY nama_wilayah ORDER BY total DESC LIMIT 5
    """, (tgl,)).fetchall()

    max_val = top_kec[0]['total'] if top_kec else 1

    kec_rows = ""
    for i, row in enumerate(top_kec):
        pct_bar = round(row['total'] / max_val * 100)
        kec_rows += f"""
        <tr>
          <td><span style="color:#64748b;font-size:12px">#{i+1}</span>&nbsp; 
              <a class="kec-link" href="/kecamatan?id={row['nama_wilayah']}&tgl={tgl}">{row['nama_wilayah']}</a></td>
          <td><span class="num-big">{row['total']:,}</span>
            <div class="progress-bar"><div class="progress-fill" style="width:{pct_bar}%"></div></div>
          </td>
        </tr>"""

    tgl_options = "".join([f"<option value='{t}' {'selected' if t==tgl else ''}>{t}</option>" for t in tanggal_list])

    return render_template_string(f"""<!DOCTYPE html><html>
    <head><title>Anomali SE2026 - Dashboard</title>{BASE_STYLE}</head>
    <body>
    {nav('home')}
    <div class="container">
      <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:24px">
        <div>
          <h1>Dashboard Anomali SE2026</h1>
          <div class="subtitle">Kabupaten Demak &middot; Data per {tgl}</div>
        </div>
        <form method="get" class="date-sel">
          <label>Tanggal:</label>
          <select name="tgl" onchange="this.form.submit()">{tgl_options}</select>
        </form>
      </div>

      <div class="cards">
        <div class="card">
          <div class="card-label">Belum Ditindaklanjuti</div>
          <div class="card-value red">{total_belum:,}</div>
        </div>
        <div class="card">
          <div class="card-label">Sudah Ditindaklanjuti</div>
          <div class="card-value green">{total_sudah:,}</div>
        </div>
        <div class="card">
          <div class="card-label">Progress Penyelesaian</div>
          <div class="card-value blue">{pct}%</div>
        </div>
        <div class="card">
          <div class="card-label">Kasus Mikro Tersimpan</div>
          <div class="card-value orange">{total_kasus:,}</div>
        </div>
      </div>

      <div class="table-wrap">
        <div class="table-head"><h2>🏆 Top Kecamatan (Anomali Belum Ditindaklanjuti)</h2></div>
        <table><tbody>{kec_rows or "<tr><td class='empty'>Tidak ada data</td></tr>"}</tbody></table>
      </div>
    </div>
    </body></html>""")

# ─────────────────────────────────────────
# PER KECAMATAN
# ─────────────────────────────────────────
@app.route('/kecamatan')
def kecamatan():
    tanggal_list = get_tanggal_list()
    if not tanggal_list:
        return f"{nav('kec')}<div class='container'><div class='empty'>Belum ada data.</div></div>"
    
    tgl = request.args.get('tgl', tanggal_list[0])
    kec_filter = request.args.get('id', '')
    db = get_db()

    # Daftar kecamatan
    kec_list = db.execute("""
        SELECT DISTINCT nama_wilayah, id_wilayah FROM agregat_anomali
        WHERE tanggal_tarik=? AND level_wilayah='kecamatan'
        ORDER BY nama_wilayah
    """, (tgl,)).fetchall()

    tgl_options = "".join([f"<option value='{t}' {'selected' if t==tgl else ''}>{t}</option>" for t in tanggal_list])
    kec_options = "<option value=''>-- Semua Kecamatan --</option>" + "".join([
        f"<option value='{r['nama_wilayah']}' {'selected' if r['nama_wilayah']==kec_filter else ''}>{r['nama_wilayah']}</option>"
        for r in kec_list])

    # Query
    if kec_filter:
        rows = db.execute("""
            SELECT nama_wilayah, kode_indikator, total_value
            FROM agregat_anomali
            WHERE tanggal_tarik=? AND level_wilayah='kecamatan' AND nama_wilayah=?
            ORDER BY CAST(kode_indikator AS INTEGER)
        """, (tgl, kec_filter)).fetchall()
    else:
        rows = db.execute("""
            SELECT nama_wilayah, kode_indikator, SUM(total_value) as total_value
            FROM agregat_anomali
            WHERE tanggal_tarik=? AND level_wilayah='kecamatan' AND total_value > 0
            GROUP BY nama_wilayah, kode_indikator
            ORDER BY nama_wilayah, CAST(kode_indikator AS INTEGER)
        """, (tgl,)).fetchall()

    # Group by kecamatan
    from collections import defaultdict
    per_kec = defaultdict(list)
    for r in rows:
        if r['total_value'] and r['total_value'] > 0:
            per_kec[r['nama_wilayah']].append(r)

    table_rows = ""
    for kec_name in sorted(per_kec.keys()):
        data = per_kec[kec_name]
        total_b = sum(r['total_value'] for r in data if str(r['kode_indikator']) in BELUM)
        total_s = sum(r['total_value'] for r in data if str(r['kode_indikator']) in SUDAH)
        ind_badges = " ".join([
            f"<span class='badge {'badge-red' if str(r['kode_indikator']) in BELUM else 'badge-green'}'>"
            f"{NAMA_INDIKATOR.get(str(r['kode_indikator']),r['kode_indikator']).split(' - ')[0]}: {r['total_value']:,}</span>"
            for r in sorted(data, key=lambda x: int(x['kode_indikator']))
        ])
        table_rows += f"""
        <tr>
          <td><strong>{kec_name}</strong></td>
          <td><span class="num" style="color:#f87171">{total_b:,}</span></td>
          <td><span class="num" style="color:#4ade80">{total_s:,}</span></td>
          <td style="line-height:1.8">{ind_badges}</td>
        </tr>"""

    return render_template_string(f"""<!DOCTYPE html><html>
    <head><title>Per Kecamatan - Anomali SE2026</title>{BASE_STYLE}</head>
    <body>
    {nav('kec')}
    <div class="container">
      <h1>Anomali per Kecamatan</h1>
      <div class="subtitle">Data per {tgl}</div>
      <form method="get" class="select-bar">
        <select name="tgl" onchange="this.form.submit()">{tgl_options}</select>
        <select name="id">{kec_options}</select>
        <button type="submit">Filter</button>
        <a href="/kecamatan?tgl={tgl}"><button type="button" style="background:#334155">Reset</button></a>
      </form>
      <div class="table-wrap">
        <table>
          <thead><tr>
            <th>Kecamatan</th><th>Belum ✗</th><th>Sudah ✓</th><th>Detail per Indikator</th>
          </tr></thead>
          <tbody>{table_rows or "<tr><td colspan='4' class='empty'>Tidak ada data anomali</td></tr>"}</tbody>
        </table>
      </div>
    </div></body></html>""")

# ─────────────────────────────────────────
# SLS / SUB-SLS
# ─────────────────────────────────────────
@app.route('/sls')
def sls_view():
    tanggal_list = get_tanggal_list()
    if not tanggal_list:
        return f"{nav('sls')}<div class='container'><div class='empty'>Belum ada data.</div></div>"
    
    tgl = request.args.get('tgl', tanggal_list[0])
    level_filter = request.args.get('level', 'sub_sls')
    db = get_db()

    rows = db.execute("""
        SELECT id_wilayah, nama_wilayah, level_wilayah, kode_indikator, total_value
        FROM agregat_anomali
        WHERE tanggal_tarik=? AND level_wilayah=? AND total_value > 0
        ORDER BY total_value DESC
        LIMIT 200
    """, (tgl, level_filter)).fetchall()

    tgl_options = "".join([f"<option value='{t}' {'selected' if t==tgl else ''}>{t}</option>" for t in tanggal_list])
    level_options = "".join([
        f"<option value='{lv}' {'selected' if lv==level_filter else ''}>{lv.upper()}</option>"
        for lv in ['sls', 'sub_sls', 'desa', 'kecamatan']
    ])

    table_rows = ""
    for r in rows:
        kode = str(r['kode_indikator'])
        cls = 'badge-red' if kode in BELUM else 'badge-green'
        table_rows += f"""
        <tr>
          <td><code style="font-size:12px;color:#94a3b8">{r['id_wilayah']}</code></td>
          <td>{r['nama_wilayah'] or '-'}</td>
          <td><span class="badge badge-gray">{r['level_wilayah']}</span></td>
          <td><span class="badge {cls}">{NAMA_INDIKATOR.get(kode, kode)}</span></td>
          <td class="num" style="{'color:#f87171' if kode in BELUM else 'color:#4ade80'}">{r['total_value']:,}</td>
        </tr>"""

    return render_template_string(f"""<!DOCTYPE html><html>
    <head><title>SLS - Anomali SE2026</title>{BASE_STYLE}</head>
    <body>
    {nav('sls')}
    <div class="container">
      <h1>Data per SLS / Sub-SLS</h1>
      <div class="subtitle">Menampilkan maks 200 baris teratas &middot; Data per {tgl}</div>
      <form method="get" class="select-bar">
        <select name="tgl" onchange="this.form.submit()">{tgl_options}</select>
        <select name="level">{level_options}</select>
        <button type="submit">Tampilkan</button>
      </form>
      <div class="table-wrap">
        <table>
          <thead><tr>
            <th>Kode Wilayah</th><th>Nama</th><th>Level</th><th>Jenis Anomali</th><th>Jumlah</th>
          </tr></thead>
          <tbody>{table_rows or "<tr><td colspan='5' class='empty'>Tidak ada data</td></tr>"}</tbody>
        </table>
      </div>
    </div></body></html>""")

# ─────────────────────────────────────────
# PER INDIKATOR
# ─────────────────────────────────────────
@app.route('/indikator')
def indikator_view():
    tanggal_list = get_tanggal_list()
    if not tanggal_list:
        return f"{nav('ind')}<div class='container'><div class='empty'>Belum ada data.</div></div>"
    
    tgl = request.args.get('tgl', tanggal_list[0])
    db = get_db()

    rows = db.execute("""
        SELECT kode_indikator, SUM(total_value) as total
        FROM agregat_anomali
        WHERE tanggal_tarik=? AND level_wilayah='kecamatan' AND total_value > 0
        GROUP BY kode_indikator
        ORDER BY total DESC
    """, (tgl,)).fetchall()

    grand = sum(r['total'] for r in rows if r['total'])
    tgl_options = "".join([f"<option value='{t}' {'selected' if t==tgl else ''}>{t}</option>" for t in tanggal_list])

    table_rows = ""
    for r in rows:
        kode = str(r['kode_indikator'])
        total = r['total'] or 0
        cls = 'badge-red' if kode in BELUM else 'badge-green'
        status = "✗ Belum" if kode in BELUM else "✓ Sudah"
        pct_bar = round(total / grand * 100) if grand else 0
        table_rows += f"""
        <tr>
          <td><span class="badge badge-gray">{kode}</span></td>
          <td>{NAMA_INDIKATOR.get(kode, kode)}</td>
          <td><span class="badge {cls}">{status}</span></td>
          <td>
            <span class="num" style="{'color:#f87171' if kode in BELUM else 'color:#4ade80'}">{total:,}</span>
            <div class="progress-bar" style="width:120px;display:inline-block;margin-left:8px;vertical-align:middle">
              <div class="progress-fill" style="width:{pct_bar}%;{'background:#f87171' if kode in BELUM else 'background:#4ade80'}"></div>
            </div>
            <span style="font-size:11px;color:#64748b;margin-left:6px">{pct_bar}%</span>
          </td>
        </tr>"""

    return render_template_string(f"""<!DOCTYPE html><html>
    <head><title>Per Indikator - Anomali SE2026</title>{BASE_STYLE}</head>
    <body>
    {nav('ind')}
    <div class="container">
      <h1>Rekap per Jenis Anomali</h1>
      <div class="subtitle">Total: {grand:,} kasus &middot; Data per {tgl}</div>
      <form method="get" class="select-bar">
        <select name="tgl" onchange="this.form.submit()">{tgl_options}</select>
      </form>
      <div class="table-wrap">
        <table>
          <thead><tr><th>Kode</th><th>Jenis Anomali</th><th>Status</th><th>Jumlah (dari kecamatan)</th></tr></thead>
          <tbody>{table_rows or "<tr><td colspan='4' class='empty'>Tidak ada data</td></tr>"}</tbody>
        </table>
      </div>
    </div></body></html>""")

# ─────────────────────────────────────────
# KASUS MIKRO
# ─────────────────────────────────────────
@app.route('/mikro')
def mikro_view():
    tanggal_list = get_tanggal_list()
    if not tanggal_list:
        return f"{nav('mikro')}<div class='container'><div class='empty'>Belum ada data.</div></div>"
    
    tgl = request.args.get('tgl', tanggal_list[0])
    page = int(request.args.get('page', 1))
    per_page = 50
    offset = (page - 1) * per_page
    db = get_db()

    total_kasus = db.execute("SELECT COUNT(*) FROM kasus_anomali_mikro WHERE tanggal_tarik=?", (tgl,)).fetchone()[0]
    rows = db.execute("""
        SELECT kode_wilayah, anomali_no, anomali_title, assignment_id, is_resolved, link_fasih
        FROM kasus_anomali_mikro
        WHERE tanggal_tarik=?
        ORDER BY kode_wilayah, anomali_no
        LIMIT ? OFFSET ?
    """, (tgl, per_page, offset)).fetchall()

    tgl_options = "".join([f"<option value='{t}' {'selected' if t==tgl else ''}>{t}</option>" for t in tanggal_list])
    total_pages = (total_kasus + per_page - 1) // per_page

    table_rows = ""
    for r in rows:
        resolved = r['is_resolved']
        badge = "<span class='badge badge-green'>✓ Selesai</span>" if resolved else "<span class='badge badge-red'>✗ Belum</span>"
        link_html = f"<a href='{r['link_fasih']}' target='_blank' style='font-size:11px'>Buka FASIH</a>" if r['link_fasih'] else "-"
        table_rows += f"""
        <tr>
          <td><code style="font-size:11px;color:#94a3b8">{r['kode_wilayah']}</code></td>
          <td><span class="badge badge-gray">Anomali-{r['anomali_no']}</span></td>
          <td style="max-width:300px;font-size:12px">{r['anomali_title'] or '-'}</td>
          <td style="font-size:12px;color:#94a3b8">{r['assignment_id'] or '-'}</td>
          <td>{badge}</td>
          <td>{link_html}</td>
        </tr>"""

    pagi_links = ""
    for p in range(max(1, page-2), min(total_pages+1, page+3)):
        cls = "style='background:#1e40af'" if p == page else ""
        pagi_links += f"<a href='/mikro?tgl={tgl}&page={p}'><button type='button' {cls}>{p}</button></a> "

    return render_template_string(f"""<!DOCTYPE html><html>
    <head><title>Kasus Mikro - Anomali SE2026</title>{BASE_STYLE}</head>
    <body>
    {nav('mikro')}
    <div class="container">
      <h1>Daftar Kasus Anomali Mikro</h1>
      <div class="subtitle">Total: {total_kasus:,} kasus &middot; Halaman {page}/{total_pages} &middot; Data per {tgl}</div>
      <form method="get" class="select-bar">
        <select name="tgl" onchange="this.form.submit()">{tgl_options}</select>
      </form>
      <div class="table-wrap">
        <table>
          <thead><tr>
            <th>Kode Wilayah</th><th>Anomali</th><th>Keterangan</th><th>Assignment ID</th><th>Status</th><th>Link</th>
          </tr></thead>
          <tbody>{table_rows or "<tr><td colspan='6' class='empty'>Tidak ada kasus mikro tersimpan</td></tr>"}</tbody>
        </table>
      </div>
      <div style="display:flex;gap:8px;align-items:center;margin-top:16px">{pagi_links}</div>
    </div></body></html>""")

# ─────────────────────────────────────────
# PANDUAN ANOMALI
# ─────────────────────────────────────────
@app.route('/panduan')
def panduan_view():
    cards_html = ""
    for a in DESKRIPSI_ANOMALI:
        ind_info = f"Ind. {a['ind_belum']} (belum)"
        if a['ind_sudah']:
            ind_info += f" &middot; Ind. {a['ind_sudah']} (sudah)"
        
        cards_html += f"""
        <div class="panduan-card">
          <div class="panduan-header">
            <span class="panduan-icon">{a['icon']}</span>
            <div>
              <div class="panduan-kode">{a['kode']}</div>
              <div class="panduan-judul">{a['judul']}</div>
            </div>
            <span class="panduan-indkode">{ind_info}</span>
          </div>
          <div class="panduan-body">
            <p class="panduan-desc">{a['deskripsi']}</p>
            <div class="panduan-contoh">
              <span class="contoh-label">📌 Contoh:</span>
              <span class="contoh-text">{a['contoh']}</span>
            </div>
          </div>
        </div>"""

    style_extra = """
    <style>
      .panduan-grid{display:grid;gap:20px}
      .panduan-card{background:#fff;border:1px solid #fcd9b6;border-radius:14px;overflow:hidden;box-shadow:0 1px 4px rgba(232,93,4,.06);transition:.2s}
      .panduan-card:hover{border-color:#f48c06;box-shadow:0 4px 16px rgba(232,93,4,.12);transform:translateY(-2px)}
      .panduan-header{display:flex;align-items:center;gap:16px;padding:18px 20px;background:linear-gradient(135deg,#fff8f3,#fff);border-bottom:1px solid #fde8d0}
      .panduan-icon{font-size:28px;flex-shrink:0}
      .panduan-kode{font-size:11px;font-weight:700;color:#ea580c;text-transform:uppercase;letter-spacing:.5px}
      .panduan-judul{font-size:15px;font-weight:700;color:#7c2d12;margin-top:2px}
      .panduan-indkode{margin-left:auto;font-size:11px;color:#9a6a4e;background:#fff7ed;border:1px solid #fcd9b6;padding:4px 10px;border-radius:20px;white-space:nowrap;flex-shrink:0}
      .panduan-body{padding:18px 20px}
      .panduan-desc{font-size:13.5px;line-height:1.7;color:#374151;margin-bottom:14px}
      .panduan-contoh{background:#fffbf5;border-left:3px solid #f97316;padding:10px 14px;border-radius:0 8px 8px 0}
      .contoh-label{font-size:12px;font-weight:600;color:#ea580c;margin-right:6px}
      .contoh-text{font-size:12.5px;color:#6b5a4e;line-height:1.6}
      .info-box{background:linear-gradient(135deg,#fff7ed,#fef3ea);border:1px solid #fcd9b6;border-radius:12px;padding:18px 22px;margin-bottom:28px;display:flex;gap:14px;align-items:flex-start}
      .info-box-icon{font-size:22px;flex-shrink:0;margin-top:2px}
      .info-box-text{font-size:13px;color:#7c2d12;line-height:1.7}
      .info-box-text strong{color:#ea580c}
    </style>"""

    return render_template_string(f"""<!DOCTYPE html><html>
    <head><title>Panduan Anomali - SE2026</title>{BASE_STYLE}{style_extra}</head>
    <body>
    {nav('panduan')}
    <div class="container">
      <h1>Panduan Jenis Anomali SE2026</h1>
      <div class="subtitle">Penjelasan dan kriteria setiap jenis anomali yang ditindaklanjuti</div>

      <div class="info-box">
        <div class="info-box-icon">ℹ️</div>
        <div class="info-box-text">
          Setiap anomali memiliki <strong>dua kode indikator</strong>:
          indikator <strong>"belum" (128-135)</strong> untuk kasus yang belum diselesaikan,
          dan indikator <strong>"sudah" (40-46)</strong> untuk kasus yang sudah ditindaklanjuti oleh petugas lapangan.
          Tujuan pemantauan adalah menekan angka "belum" dan meningkatkan angka "sudah".
        </div>
      </div>

      <div class="panduan-grid">
        {cards_html}
      </div>
    </div>
    </body></html>""")


if __name__ == '__main__':
    if not os.path.exists(DB_FILE):
        print(f"[!] File {DB_FILE} tidak ditemukan. Jalankan tarik_anomali.py dulu.")
    else:
        print("="*50)
        print("  Web Anomali SE2026")
        print("  Buka di browser: http://localhost:5050")
        print("  Tekan Ctrl+C untuk berhenti")
        print("="*50)
    app.run(debug=False, port=5050, host='0.0.0.0')
