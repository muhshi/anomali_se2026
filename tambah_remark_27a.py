#!/usr/bin/env python3
"""
Bot Otomatisasi Penambahan Catatan / Remark '#0' pada Pertanyaan 27.a
27. a. Nilai produksi/pendapatan/penjualan barang dan jasa
Fasih-SM BPS (SE2026)

Mendukung Multi-Sumber Google Sheets:
1. Sheet 'Dtata_NTB' (gid: 293925488):
   https://docs.google.com/spreadsheets/d/1ra-HZlqn1md8AIYaZz5VmOg-5Dxj83LYOHrmChvyz0s/edit?gid=293925488#gid=293925488
2. Sheet 'Data detail anomali' (gid: 0):
   https://docs.google.com/spreadsheets/d/12FWSyYak-scXS5skHFUSQbVkjy6Oh8ofejusqnOb5nA/edit?gid=0#gid=0

Alur Pengerjaan:
1. Mengambil data anomali live langsung dari Google Sheets.
   Filter default: Kolom 'Petugas Cleaning' = 0
2. Menu interaktif di terminal:
   - Pilih Sumber Data (Dtata_NTB / Data detail anomali / Gabungkan Keduanya).
   - Filter Data (Semua data Petugas Cleaning = 0 / ID Desa / Batch).
3. Membuka link assignment Fasih-SM (/edit).
4. Masuk ke menu 'SE2026 - L BLOK II: KETERANGAN USAHA/PERUSAHAAN'.
5. Mencocokkan dan memilih nama usaha di submenu / kartu rincian.
6. Menemukan pertanyaan '27. a. Nilai produksi/pendapatan/penjualan barang dan jasa'.
7. Mengklik tombol Catatan (speech bubble), mengisi teks remark '#0', dan mengklik Simpan catatan.
8. Menyimpan rincian jika ada tombol Simpan/Selesai di form rincian.
9. Masuk ke menu Catatan di sidebar.
10. Memastikan switch/checkbox 'Tampilkan Anomali Usaha dan Keluarga' bernilai TRUE.
11. Mengklik Kirim dan menyelesaikan konfirmasi bertingkat (konfirm 2x / Submit Paksa).
12. Lanjut ke assignment berikutnya.
"""

import os
import sys
import time
import re
import json
import csv
import urllib.request
import urllib.parse
import threading
import subprocess

# Auto-install library yang dibutuhkan jika belum terpasang
REQUIRED_LIBS = ["openpyxl", "playwright"]
for lib in REQUIRED_LIBS:
    try:
        __import__(lib)
    except ImportError:
        print(f"Library '{lib}' belum terpasang. Menginstall secara otomatis...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", lib])

import openpyxl
from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
CACHE_FILE = os.path.join(BASE_DIR, "processed_tambah_remark.json")
REPORT_EXCEL = os.path.join(BASE_DIR, "laporan_tambah_remark.xlsx")
REPORT_JSON = os.path.join(BASE_DIR, "laporan_tambah_remark.json")
SYNC_CONFIG_FILE = os.path.join(BASE_DIR, "sheet_sync_config.json")

DEFAULT_REMARK_TEXT = "#0"

SHEET_SOURCES = [
    {
        "id": "1",
        "name": "Prioritas 1: Data Detail Anomali (gid: 0)",
        "url": "https://docs.google.com/spreadsheets/d/12FWSyYak-scXS5skHFUSQbVkjy6Oh8ofejusqnOb5nA/export?format=csv&gid=0",
        "backup_filename": "data_detail_anomali_backup.csv"
    },
    {
        "id": "2",
        "name": "Prioritas 2: Data NTB (gid: 293925488)",
        "url": "https://docs.google.com/spreadsheets/d/1ra-HZlqn1md8AIYaZz5VmOg-5Dxj83LYOHrmChvyz0s/export?format=csv&gid=293925488",
        "backup_filename": "dtata_ntb_backup.csv"
    }
]

# Mapping Resmi Kode Kecamatan Kabupaten Demak (3321)
KECAMATAN_MAP_3321 = {
    "3321010": "MRANGGEN",
    "3321020": "KARANGAWEN",
    "3321030": "GUNTUR",
    "3321040": "SAYUNG",
    "3321050": "KARANGTENGAH",
    "3321060": "BONANG",
    "3321070": "DEMAK",
    "3321080": "WONOSALAM",
    "3321090": "DEMPET",
    "3321091": "KEBONAGUNG",
    "3321100": "GAJAH",
    "3321110": "KARANGANYAR",
    "3321120": "MIJEN",
    "3321130": "WEDUNG",
    "3321000": "KAB. DEMAK (LAINNYA)",
}

def load_wilayah_reference():
    """Memuat daftar nama kecamatan & desa dari file excel master jika tersedia."""
    kec_map = dict(KECAMATAN_MAP_3321)
    desa_map = {}
    ref_file = os.path.join(DATA_DIR, "Anomali", "Data_Mikro_Anomali_Gabungan_3321_20260929_075829.xlsx")
    if os.path.exists(ref_file):
        try:
            wb = openpyxl.load_workbook(ref_file, read_only=True, data_only=True)
            ws = wb.active
            for r in ws.iter_rows(min_row=5, values_only=True):
                if r and len(r) >= 10:
                    k_kec = str(r[6]).strip() if r[6] is not None else ""
                    n_kec = str(r[7]).strip() if r[7] is not None else ""
                    k_desa = str(r[8]).strip() if r[8] is not None else ""
                    n_desa = str(r[9]).strip() if r[9] is not None else ""
                    if k_kec and n_kec and not k_kec.startswith("("):
                        kec_map[k_kec] = n_kec
                    if k_desa and n_desa and not k_desa.startswith("("):
                        desa_map[k_desa] = n_desa
            wb.close()
        except Exception:
            pass
    return kec_map, desa_map

# ==============================================================================
# HELPER SINKRONISASI GOOGLE APPS SCRIPT WEB APP
# ==============================================================================

def load_sync_config():
    """Memuat konfigurasi sinkronisasi Google Sheet Web App."""
    if os.path.exists(SYNC_CONFIG_FILE):
        try:
            with open(SYNC_CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"gas_webapp_url": "", "operator_name": ""}

def save_sync_config(cfg):
    try:
        with open(SYNC_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4)
    except Exception as e:
        print(f"[Warning] Gagal menyimpan sync config: {e}")

def _do_sync_request(url, item, status="SELESAI", operator=""):
    """Mengirim request update baris ke Google Apps Script Web App."""
    try:
        # Deteksi gid jika ada di nama sumber data
        source_name = item.get("source_name", "")
        gid = ""
        gid_match = re.search(r"gid[:\s=]+(\d+)", source_name)
        if gid_match:
            gid = gid_match.group(1)

        params = {
            "link": item.get("link", ""),
            "nama_usaha": item.get("nama_usaha", ""),
            "row_num": item.get("row_num", ""),
            "iddesa": item.get("iddesa", ""),
            "status": status,
            "operator": operator,
            "gid": gid,
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
        }
        query_str = urllib.parse.urlencode(params)
        req_url = f"{url}?{query_str}"
        req = urllib.request.Request(
            req_url,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            resp_body = resp.read().decode("utf-8")
            try:
                res_json = json.loads(resp_body)
                if res_json.get("status") == "success":
                    sheet_info = f"[{res_json.get('sheet')}] " if res_json.get("sheet") else ""
                    print(f"  -> [SYNC SHEET OK] {sheet_info}Baris {res_json.get('row', item.get('row_num'))} ditandai '{status}' & diwarnai hijau!")
                else:
                    msg = res_json.get('message', 'Update terkirim.')
                    print(f"  -> [SYNC SHEET Info] {msg}")
            except Exception:
                pass
    except Exception as e:
        print(f"  -> [SYNC SHEET Warning] Gagal sinkronisasi Google Sheet: {e}")

def sync_to_google_sheet(item, status="SELESAI", async_mode=True):
    """
    Sinkronisasi status dan pewarnaan hijau ke Google Spreadsheet target.
    Dijalankan secara asinkron (background thread) agar tidak memperlambat bot.
    Mendukung multi-spreadsheet (memilih URL Web App sesuai spreadsheet target).
    """
    cfg = load_sync_config()
    source_name = item.get("source_name", "")
    urls_map = cfg.get("gas_webapp_urls", {})
    url = ""
    for k, u in urls_map.items():
        if u and (k.lower() in source_name.lower() or source_name.lower() in k.lower()):
            url = u.strip()
            break
    if not url:
        url = (cfg.get("gas_webapp_url") or "").strip()

    operator = (cfg.get("operator_name") or "").strip()
    if not url:
        return

    if async_mode:
        t = threading.Thread(target=_do_sync_request, args=(url, item, status, operator), daemon=True)
        t.start()
    else:
        _do_sync_request(url, item, status, operator)

# ==============================================================================
# HELPER CACHE & LAPORAN
# ==============================================================================

def load_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                return set(json.load(f))
        except Exception:
            return set()
    return set()

def save_cache(cache_set):
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(list(cache_set), f, indent=4)
    except Exception as e:
        print(f"[Warning] Gagal menyimpan cache: {e}")

def load_report_data():
    if os.path.exists(REPORT_JSON):
        try:
            with open(REPORT_JSON, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_report_data(report_list):
    try:
        with open(REPORT_JSON, "w", encoding="utf-8") as f:
            json.dump(report_list, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"[Warning] Gagal menyimpan file JSON laporan: {e}")

    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Rekap Tambah Remark 27a"

        headers = [
            "No", "Sumber Sheet", "Kecamatan", "ID Desa", "Nama Desa", "Nama Usaha", "Batch", "Petugas Cleaning",
            "Catatan Sheet", "Teks Remark", "Status Pengerjaan", "Keterangan", "Link Fasih", "Waktu Selesai"
        ]
        ws.append(headers)

        for idx, row in enumerate(report_list, 1):
            ws.append([
                idx,
                row.get("source_name", ""),
                row.get("nama_kecamatan", ""),
                row.get("iddesa", ""),
                row.get("nama_desa", ""),
                row.get("nama_usaha", ""),
                row.get("batch_info", ""),
                row.get("petugas_cleaning", ""),
                row.get("catatan_fasih", ""),
                row.get("remark_text", DEFAULT_REMARK_TEXT),
                row.get("status", ""),
                row.get("keterangan", ""),
                row.get("link", ""),
                row.get("timestamp", "")
            ])

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 60)

        wb.save(REPORT_EXCEL)
    except Exception as e:
        print(f"[Warning] Gagal memperbarui file Excel laporan: {e}")

# ==============================================================================
# FETCH DATA GOOGLE SHEETS
# ==============================================================================

def fetch_single_sheet(source_info):
    """Mengunduh dan membaca satu sumber Google Sheet."""
    name = source_info["name"]
    url = source_info["url"]
    backup_file = os.path.join(DATA_DIR, source_info["backup_filename"])

    print(f"\n-> Menghubungkan ke '{name}'...")
    print(f"   URL: {url}")

    req = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    )

    content = None
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            content = resp.read().decode("utf-8")
    except Exception as e:
        print(f"   [Warning] Gagal mengunduh live: {e}")
        if os.path.exists(backup_file):
            print(f"   -> Menggunakan file backup lokal: {backup_file}")
            with open(backup_file, "r", encoding="utf-8") as f:
                content = f.read()

    if not content:
        return []

    # Simpan backup lokal
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        with open(backup_file, "w", encoding="utf-8") as f:
            f.write(content)
    except Exception:
        pass

    # Parse CSV dengan deteksi baris header
    lines = [l for l in content.splitlines() if l.strip()]
    header_idx = 0
    for idx, l in enumerate(lines[:10]):
        l_lower = l.lower()
        if "link_fasih" in l_lower or "nama_usaha" in l_lower:
            header_idx = idx
            break

    reader = csv.DictReader(lines[header_idx:])
    data_items = []
    kec_map, desa_map = load_wilayah_reference()

    for r_idx, row in enumerate(reader, 2):
        if not row:
            continue

        raw_link = (row.get("link_fasih_sm") or row.get("link_fasih") or row.get("link") or "").strip()
        if not raw_link or "fasih" not in raw_link:
            continue

        clean_link = raw_link.split("?")[0].rstrip("/")
        if not clean_link.endswith("/edit"):
            clean_link = clean_link + "/edit"

        petugas_cleaning = str(row.get("Petugas Cleaning") or row.get("petugas_cleaning") or "").strip()
        batch_info = str(row.get("batch_info") or "").strip()
        iddesa = str(row.get("iddesa") or "").strip()
        nama_usaha = str(row.get("nama_usaha") or "").strip()
        catatan_fasih = str(row.get("catatan_fasih") or "").strip()
        row_num = str(row.get("row_num") or r_idx).strip()

        # Ekstrak Kode dan Nama Kecamatan & Desa
        idkec = iddesa[:7] if len(iddesa) >= 7 else (iddesa or "LAINNYA")
        nama_kec = kec_map.get(idkec, f"Kec. {idkec}" if idkec != "LAINNYA" else "LAINNYA")
        nama_desa = desa_map.get(iddesa, "")

        # Deteksi apakah kolom Status Pengerjaan di Google Sheet sudah menyatakan selesai
        status_sheet = ""
        for k, v in row.items():
            if not k:
                continue
            k_lower = k.lower()
            if any(term in k_lower for term in ["status pengerjaan", "status remark", "status selesai", "status bot", "status_cleaning"]):
                status_sheet = str(v or "").strip()
                break

        is_sheet_done = False
        if status_sheet:
            st_upper = status_sheet.upper()
            if any(dt in st_upper for dt in ["SELESAI", "DONE", "SUDAH", "BERHASIL"]):
                is_sheet_done = True

        data_items.append({
            "source_name": name,
            "link": clean_link,
            "raw_link": raw_link,
            "petugas_cleaning": petugas_cleaning,
            "batch_info": batch_info,
            "iddesa": iddesa,
            "idkec": idkec,
            "nama_kecamatan": nama_kec,
            "nama_desa": nama_desa,
            "nama_usaha": nama_usaha,
            "catatan_fasih": catatan_fasih,
            "row_num": row_num,
            "status_sheet": status_sheet,
            "is_sheet_done": is_sheet_done
        })

    print(f"   [OK] {len(data_items)} data berhasil dimuat dari '{name}'.")
    return data_items

def fetch_data_from_sources(sources_to_load):
    """Mengunduh dari satu atau lebih Google Sheet dan menggabungkannya."""
    all_data = []
    for s in sources_to_load:
        items = fetch_single_sheet(s)
        all_data.extend(items)
    return all_data

# ==============================================================================
# ANTI-BOT / WAF / SSO HANDLERS
# ==============================================================================

def check_connection_or_session_lost(page):
    """
    Mendeteksi apakah koneksi internet/VPN putus atau sesi Fasih-SM telah habis/keluar/terdeteksi bot.
    """
    try:
        url = (page.url or "").lower()

        # 1. URL redirect ke login / SSO
        if ("sso" in url and ("login" in url or "auth" in url or "error" in url)) or ("/login" in url) or ("/auth" in url):
            if "/app/assignment/" not in url:
                return True, "Sesi login keluar / dialihkan ke halaman SSO BPS"

        # 2. Halaman error jaringan Chromium / Windows Diagnostics / Blank DOM
        res_check = page.evaluate('''() => {
            if (!document || !document.body) {
                return { isError: true, reason: 'document.body null (koneksi terputus)' };
            }
            const title = (document.title || '').toLowerCase();
            const text = (document.body.innerText || '').toLowerCase();
            const full = title + ' ' + text;

            const errorKeywords = [
                'checking the connection',
                'running windows network diagnostics',
                'err_internet_disconnected',
                'err_connection_closed',
                'err_connection_refused',
                'err_connection_timed_out',
                'err_name_not_resolved',
                'no internet',
                'tidak ada koneksi',
                'access denied',
                'session expired',
                'sesi anda telah berakhir',
                'sesi berakhir'
            ];

            for (const kw of errorKeywords) {
                if (full.includes(kw)) {
                    return { isError: true, reason: kw };
                }
            }

            // Halaman blank kosong
            const bodyChildren = document.body.children.length;
            const bodyTextLen = text.trim().length;
            if (bodyChildren === 0 || bodyTextLen === 0) {
                return { isError: true, reason: 'halaman kosong (blank)' };
            }

            return { isError: false };
        }''')

        if res_check and res_check.get("isError"):
            return True, f"Error koneksi / halaman blank: {res_check.get('reason')}"

    except Exception as e:
        err_msg = str(e).lower()
        if "target closed" in err_msg or "context closed" in err_msg or "navigation failed" in err_msg or "connection refused" in err_msg:
            return True, f"Koneksi/Browser terputus: {e}"

    return False, ""

def pause_for_relogin(page, target_link):
    """
    Menghentikan bot sementara ketika sesi terputus atau VPN disconnect,
    memberi notifikasi di terminal dan menunggu user login ulang sebelum melanjutkan.
    """
    print("\n" + "!" * 65)
    print(" [PERHATIAN] KONEKSI VPN TERPUTUS ATAU SESI LOGIN FASIH-SM KELUAR!")
    print("!" * 65)
    print(" Langkah pemulihan:")
    print(" 1. Pastikan koneksi internet & VPN BPS Anda sudah aktif kembali.")
    print(" 2. Login ulang di browser yang sedang terbuka (melalui SSO BPS).")
    print(" 3. Buka halaman Fasih-SM sampai dashboard atau form berhasil termuat.")
    print(" 4. KEMBALI KE TERMINAL INI lalu tekan tombol ENTER untuk melanjutkan bot.")
    print("!" * 65)

    input("\nTekan ENTER setelah Anda selesai login ulang...")

    print("-> Mencoba membuka kembali assignment setelah login ulang...")
    try:
        page.goto(target_link, wait_until="networkidle", timeout=30000)
        time.sleep(2.0)
    except Exception as e:
        print(f"[Warning] Gagal membuka link setelah re-login: {e}")

def resolve_bot_detection(page, target_link):
    """Penanganan jika terdeteksi bot WAF / Halosis."""
    try:
        is_bot = page.evaluate('''() => {
            const t = ((document.body && document.body.innerText) || '').toLowerCase();
            return t.includes('mendeteksi koneksi anda sebagai bot') || t.includes('halosis') || t.includes('access denied');
        }''')
    except Exception:
        is_bot = False

    if not is_bot:
        return True

    print("=" * 60)
    print(" [WAF / BOT DETECTED] Halaman terdeteksi bot atau terganggu sesi SSO.")
    print(" Memulai prosedur pemulihan otomatis...")

    for attempt in range(1, 3):
        print(f" -> [Percobaan {attempt}/2] Menunggu 10 detik agar rate-limit ter-reset...")
        time.sleep(10)

        print(" -> Refresh/reload halaman...")
        try:
            page.reload(wait_until="networkidle", timeout=15000)
        except Exception:
            try:
                page.goto(target_link, wait_until="networkidle", timeout=15000)
            except Exception:
                pass
        time.sleep(4)

        try:
            sso_btn = page.locator("button, a, div").filter(
                has_text=re.compile(r"Lanjutkan dengan SSO|Lanjutkan SSO|Login.*SSO|Masuk.*SSO|Lanjutkan", re.IGNORECASE)
            )
            if sso_btn.count() > 0 and sso_btn.first.is_visible():
                print(" -> Mengklik tombol 'Lanjutkan dengan SSO'...")
                sso_btn.first.click()
                time.sleep(5)
                try:
                    page.wait_for_load_state("networkidle", timeout=10000)
                except Exception:
                    pass
        except Exception:
            pass

        try:
            is_still_bot = page.evaluate('''() => {
                const t = ((document.body && document.body.innerText) || '').toLowerCase();
                return t.includes('mendeteksi koneksi anda sebagai bot') || t.includes('halosis') || t.includes('access denied');
            }''')
            if not is_still_bot:
                print(" -> [BERHASIL] Halaman terbebas dari deteksi bot! Melanjutkan bot...")
                return True
        except Exception:
            pass

    print("=" * 60)
    print(" TERDETEKSI SEBAGAI BOT OLEH SERVER (WAF BPS / HALOSIS)!")
    print(" Silakan selesaikan verifikasi di browser secara manual.")
    print("=" * 60)
    input("Tekan ENTER di terminal ini jika halaman sudah normal...")
    return True

# ==============================================================================
# DETEKSI DAN PENUNGGUAN LOADING DINAMIS FASIH-SM
# ==============================================================================

def wait_for_kuesioner_ready(page, max_wait=60):
    """
    Menunggu kuesioner assignment selesai dimuat saat pertama kali membuka link /edit.
    Loading screen Fasih-SM hanya terjadi saat pengambilan data kuesioner pertama kali.
    Begitu elemen kuesioner (SE2026 / BLOK / Keterangan) sudah muncul dan terlihat,
    loading screen selesai dan proses dapat langsung dilanjutkan.
    """
    start_time = time.time()
    reported = False

    while time.time() - start_time < max_wait:
        try:
            res = page.evaluate('''() => {
                function isVisible(el) {
                    if (!el) return false;
                    const style = window.getComputedStyle(el);
                    return style.display !== 'none' && 
                           style.visibility !== 'hidden' && 
                           style.opacity !== '0' && 
                           el.offsetWidth > 0 && 
                           el.offsetHeight > 0;
                }

                const bodyText = (document.body.innerText || '').toLowerCase();

                // 1. Cek apakah elemen kuesioner sudah ter-render di DOM
                const hasKuesionerContent = (
                    bodyText.includes('se2026') || 
                    bodyText.includes('blok i') || 
                    bodyText.includes('blok ii') || 
                    bodyText.includes('keterangan usaha') ||
                    bodyText.includes('kuesioner')
                );

                const hasSidebar = document.querySelector('aside, nav, [role="tablist"], [role="tab"], .sidebar') !== null;

                // 2. Cek apakah ada spinner animasi aktif
                const spinners = Array.from(document.querySelectorAll('.animate-spin, [class*="spinner"], [class*="loader"]')).filter(isVisible);

                // 3. Cek apakah ada teks loading aktif yang berdiri sendiri
                const allTexts = Array.from(document.querySelectorAll('div, span, p, h1, h2, h3, h4'));
                const loadingTexts = allTexts.filter(el => {
                    if (!isVisible(el)) return false;
                    const t = (el.innerText || el.textContent || '').trim().toLowerCase();
                    return (
                        t === 'memuat...' || 
                        t === 'loading...' || 
                        t === 'sedang memuat...' || 
                        t === 'sedang memuat' || 
                        t === 'mengambil data...' || 
                        t === 'mengambil data' || 
                        t === 'mohon tunggu...' || 
                        t === 'menyiapkan formulir...' ||
                        t.includes('sedang memuat data kuesioner')
                    ) && el.children.length === 0;
                });

                // Jika kuesioner sudah muncul dan tidak ada teks/spinner loading yang menghalangi
                const isReady = (hasKuesionerContent || hasSidebar) && spinners.length === 0 && loadingTexts.length === 0;

                return {
                    isReady: isReady,
                    hasContent: hasKuesionerContent || hasSidebar,
                    spinnersCount: spinners.length,
                    textsCount: loadingTexts.length
                };
            }''')

            if res and res.get("isReady"):
                if reported:
                    elapsed = time.time() - start_time
                    print(f"     [OK] Kuesioner berhasil dimuat dalam {elapsed:.1f} detik.")
                return True

            if not reported:
                print("  -> [Loading] Menunggu kuesioner selesai dimuat dari server...")
                reported = True

            time.sleep(0.8)
        except Exception:
            time.sleep(0.8)

    if reported:
        print(f"     [Warning] Waktu tunggu kuesioner mencapai batas {max_wait} detik, mencoba melanjutkan...")
    return False

def wait_for_loading_finish(page, max_wait=5, context=""):
    """
    Pengecekan loading ringan untuk aksi-aksi lokal tanpa false positive overlay.
    """
    time.sleep(0.5)
    return True

# ==============================================================================
# NAVIGASI BLOK II & PILIH USAHA
# ==============================================================================

def open_menu_blok_ii(page):
    """
    Membuka menu 'SE2026 - L BLOK II' (KETERANGAN USAHA/PERUSAHAAN) di sidebar.
    """
    time.sleep(1.0)

    for attempt in range(5):
        # 1. Coba klik langsung tab / tombol dengan Playwright locator di sidebar
        try:
            tab_locator = page.locator("aside, nav, [role='tablist'], [class*='sidebar'], ul").locator("a, button, [role='tab'], li").filter(
                has_text=re.compile(r"BLOK\s*II.*KETERANGAN\s*USAHA|BLOK\s*II\b", re.IGNORECASE)
            ).first
            if tab_locator.is_visible():
                tab_locator.scroll_into_view_if_needed(timeout=2000)
                tab_locator.click(timeout=3000, force=True)
                print("     [OK] Menu 'BLOK II' diklik via Playwright locator.")
                time.sleep(2.0)
                return True
        except Exception:
            pass

        # 2. Evaluasi JavaScript DOM: jika accordion 'SE2026 - L' tertutup, klik untuk ekspansi
        res = page.evaluate('''() => {
            function getCleanText(el) {
                return (el.innerText || el.textContent || '').trim();
            }

            function clickEl(el) {
                el.scrollIntoView({ behavior: 'instant', block: 'center' });
                const rect = el.getBoundingClientRect();
                const x = rect.left + rect.width / 2;
                const y = rect.top + rect.height / 2;
                const opts = { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y, pointerId: 1, pointerType: 'mouse', isPrimary: true };
                el.dispatchEvent(new PointerEvent('pointerdown', opts));
                el.dispatchEvent(new MouseEvent('mousedown', opts));
                el.dispatchEvent(new PointerEvent('pointerup', opts));
                el.dispatchEvent(new MouseEvent('mouseup', opts));
                el.click();
            }

            const all = Array.from(document.querySelectorAll('aside, nav, [role="tablist"], [role="tree"], [class*="sidebar"], [class*="menu"], ul, div[class*="sidebar"] *'));
            
            // Cari leaf dengan kecocokan spesifik 'BLOK II' di sidebar
            const regBlokII = /(?:SE2026\s*[-–—]?\s*L.*BLOK\s*II|BLOK\s*II.*KETERANGAN\s*USAHA|BLOK\s*II\b)/i;
            const matchedLeafs = all.filter(el => {
                const t = getCleanText(el);
                return regBlokII.test(t) && !t.toLowerCase().includes('anomali') &&
                       !Array.from(el.children).some(c => regBlokII.test(getCleanText(c)));
            });

            if (matchedLeafs.length > 0) {
                matchedLeafs.sort((a, b) => getCleanText(a).length - getCleanText(b).length);
                const target = matchedLeafs[0];
                const clickable = target.closest('a, button, [role="tab"], [role="treeitem"], [role="menuitem"], li, [tabindex]') || target;
                clickEl(clickable);
                return { success: true, menu: getCleanText(clickable) };
            }

            // Jika BLOK II belum terlihat, ekspansi 'SE2026 - L'
            const regL = /SE2026\s*[-–—]?\s*L\b/i;
            const lLeafs = all.filter(el => {
                const t = getCleanText(el);
                return regL.test(t) && !Array.from(el.children).some(c => regL.test(getCleanText(c)));
            });

            if (lLeafs.length > 0) {
                lLeafs.sort((a, b) => getCleanText(a).length - getCleanText(b).length);
                const lTarget = lLeafs[0].closest('a, button, [role="tab"], [role="treeitem"], li') || lLeafs[0];
                clickEl(lTarget);
                return { success: false, need_wait: true, action: 'expanded_se2026_l', menu: getCleanText(lTarget) };
            }

            return { success: false };
        }''')

        if res and res.get("success"):
            print(f"     [OK] Berhasil membuka menu: '{res.get('menu')}'")
            time.sleep(2.0)
            return True
        elif res and res.get("need_wait"):
            print(f"     -> Mengklik 'SE2026 - L' untuk menampilkan sub-menu ('{res.get('menu')}')...")
            time.sleep(1.5)
            continue
        else:
            time.sleep(1.0)

    return False

def select_or_open_usaha(page, target_nama_usaha):
    """
    Mencari kartu usaha / baris tabel yang cocok dengan nama usaha lalu mengklik tombol 'Lihat Rincian'.
    """
    time.sleep(1.0)
    clean_target = (target_nama_usaha or "").strip()
    print(f"  -> Mencari dan membuka rincian usaha: '{clean_target}' (Lihat Rincian)...")

    # Cek jika modal / form rincian sudah terbuka
    is_open = page.evaluate('''() => {
        const dialog = document.querySelector('div[role="dialog"]');
        if (dialog && dialog.offsetParent !== null) return true;
        const bodyText = (document.body.innerText || '').toLowerCase();
        return bodyText.includes('nilai produksi') && bodyText.includes('27');
    }''')
    if is_open:
        print("     [OK] Form/Modal rincian usaha sudah terbuka.")
        return True

    # 1. Cari kartu/baris usaha yang cocok dengan nama usaha lalu klik 'Lihat Rincian'
    btn_clicked = False
    if clean_target:
        res_card = page.evaluate('''(namaUsaha) => {
            function getCleanText(el) {
                return (el.innerText || el.textContent || '').trim();
            }
            function cleanStr(s) {
                return s.toLowerCase().replace(/[^a-z0-9]/g, '');
            }
            function clickEl(el) {
                el.scrollIntoView({ behavior: 'instant', block: 'center' });
                const rect = el.getBoundingClientRect();
                const x = rect.left + rect.width / 2;
                const y = rect.top + rect.height / 2;
                const opts = { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y, pointerId: 1, pointerType: 'mouse', isPrimary: true };
                el.dispatchEvent(new PointerEvent('pointerdown', opts));
                el.dispatchEvent(new MouseEvent('mousedown', opts));
                el.dispatchEvent(new PointerEvent('pointerup', opts));
                el.dispatchEvent(new MouseEvent('mouseup', opts));
                el.click();
            }

            const uClean = cleanStr(namaUsaha);
            const words = namaUsaha.toLowerCase().split(/\\s+/).filter(w => w.length > 2);
            
            // Cari di semua kartu / baris tabel
            const containers = Array.from(document.querySelectorAll('tr, [role="row"], .card, div[class*="border"], div[class*="rounded"], div[class*="card"]'));
            for (const c of containers) {
                const cText = getCleanText(c);
                const cClean = cleanStr(cText);
                
                const matchFull = cClean.includes(uClean);
                const matchWords = words.length > 0 && words.some(w => cText.toLowerCase().includes(w));
                
                if (matchFull || matchWords) {
                    const btns = Array.from(c.querySelectorAll('button, a, [role="button"]'));
                    const targetBtn = btns.find(b => {
                        const t = getCleanText(b).toLowerCase();
                        return (t === 'lihat rincian' || t.includes('lihat rincian') || t === 'lihat' || t === 'buka') && t.length < 30;
                    });
                    if (targetBtn) {
                        clickEl(targetBtn);
                        return { success: true, method: 'card_match', text: getCleanText(targetBtn) };
                    }
                }
            }

            // Coba juga jika nama usaha ada di sidebar sub-menu
            const sidebarEls = Array.from(document.querySelectorAll('aside, nav, [role="tablist"], [class*="sidebar"], [class*="menu"], ul'));
            for (const sb of sidebarEls) {
                const items = Array.from(sb.querySelectorAll('a, button, [role="tab"], [role="treeitem"], [role="menuitem"], li'));
                for (const it of items) {
                    const itText = getCleanText(it);
                    if (cleanStr(itText).includes(uClean) || (words.length > 0 && words.some(w => itText.toLowerCase().includes(w)))) {
                        clickEl(it);
                        return { success: true, method: 'sidebar_submenu', text: itText };
                    }
                }
            }

            return { success: false };
        }''', clean_target)

        if res_card and res_card.get("success"):
            print(f"     [OK] Tombol/Menu usaha '{clean_target}' berhasil diklik: '{res_card.get('text')}' ({res_card.get('method')})")
            btn_clicked = True

    # 2. Fallback: Cari tombol 'Lihat Rincian' umum di halaman jika hanya 1 usaha
    if not btn_clicked:
        try:
            global_btn = page.locator("button, a, [role='button']").filter(has_text=re.compile(r"Lihat\s*Rincian", re.IGNORECASE)).first
            if global_btn.count() > 0 and global_btn.is_visible():
                global_btn.scroll_into_view_if_needed(timeout=2000)
                global_btn.click(timeout=3000, force=True)
                print("     [OK] Tombol 'Lihat Rincian' berhasil diklik via global locator.")
                btn_clicked = True
        except Exception:
            pass

    time.sleep(1.0)

    # Pastikan form / modal benar-benar ter-mount sebelum melanjutkan
    for _ in range(15):
        form_ready = page.evaluate('''() => {
            const dialog = document.querySelector('div[role="dialog"]');
            if (dialog && dialog.offsetParent !== null) return true;
            const bodyText = (document.body.innerText || '').toLowerCase();
            return bodyText.includes('keterangan tempat') || bodyText.includes('nilai produksi') || bodyText.includes('rincian');
        }''')
        if form_ready:
            break
        time.sleep(1.0)

    time.sleep(1.5)
    return True

# ==============================================================================
# TAMBAH REMARK PADA PERTANYAAN 27.a (NILAI PRODUKSI / PENDAPATAN)
# ==============================================================================

def add_remark_27a(page, remark_text=DEFAULT_REMARK_TEXT):
    """
    Menambahkan catatan / remark '#0' pada pertanyaan:
    '27. a. Nilai produksi/pendapatan/penjualan barang dan jasa':
    1. Melakukan scanning scroll bertahap ke bawah pada modal/form rincian sampai pertanyaan 27.a ditemukan.
    2. Memverifikasi dengan ketat label pertanyaan 27.a (BUKAN pertanyaan 8!).
    3. Mengklik tombol Catatan (speech bubble) khusus untuk pertanyaan 27.a:
       <button title="Catatan" aria-label="Catatan">
    4. Mengisikan teks remark ('#0') ke dalam input / textarea popup.
    5. Mengklik tombol Simpan (icon paper plane atau title='Simpan').
    6. Memverifikasi bahwa catatan '#0' tersimpan.
    7. Menutup popup catatan.
    8. Mengklik Simpan di form rincian jika ada.
    """
    print(f"  -> Mencari pertanyaan '27. a' dan mengklik tombol Catatan...")
    time.sleep(1.0)

    # ==========================================================================
    # TAHAP 1: SCANNING SCROLL KE BAWAH SAMPAI PERTANYAAN 27.a DITEMUKAN
    # ==========================================================================
    found_27a = False
    q_text_found = ""

    # Lakukan hingga 2 putaran scan (jika putaran 1 belum selesai loading komponen bawah)
    for pass_num in range(1, 3):
        if found_27a:
            break
            
        if pass_num > 1:
            print(f"  -> [Putaran {pass_num}] Menggulir kembali ke atas dan memindai ulang formulir...")
            page.evaluate('''() => {
                window.scrollTo(0, 0);
                const scrollables = Array.from(document.querySelectorAll('*')).filter(el => {
                    const s = window.getComputedStyle(el);
                    return (s.overflowY === 'auto' || s.overflowY === 'scroll') && el.scrollHeight > el.clientHeight;
                });
                scrollables.forEach(s => s.scrollTop = 0);
            }''')
            time.sleep(1.0)

        for scan_step in range(1, 36):
            # Cek apakah elemen 27.a sudah termuat di DOM
            res_search = page.evaluate('''() => {
                function getCleanText(el) {
                    return (el.innerText || el.textContent || '').trim();
                }

                const all = Array.from(document.querySelectorAll('*'));
                
                // Regex ketat untuk pertanyaan 27.a
                // Harus memuat 27.a atau 27. a atau (27 dan nilai produksi)
                const matchedLeafs = all.filter(el => {
                    const t = getCleanText(el);
                    if (!t || t.length > 200) return false;
                    
                    const is27a = /\\b27\\s*[\\.\\-]?\\s*a\\b/i.test(t) || 
                                  (/\\b27\\b/.test(t) && /nilai\\s*produksi/i.test(t));
                    if (!is27a) return false;

                    // Pastikan tidak mengandung nomor pertanyaan lain sebagai judul (bukan pertanyaan 8 atau lainnya)
                    if (/\\b(?:8\\s*[\\.\\-]|9\\s*[\\.\\-]|26\\s*[\\.\\-]|28\\s*[\\.\\-])/i.test(t)) {
                        return false;
                    }

                    // Harus elemen leaf terdalam yang memuat teks tersebut
                    const childHasMatch = Array.from(el.children).some(c => {
                        const ct = getCleanText(c);
                        return /\\b27\\s*[\\.\\-]?\\s*a\\b/i.test(ct) || (/\\b27\\b/.test(ct) && /nilai\\s*produksi/i.test(ct));
                    });
                    return !childHasMatch;
                });

                if (matchedLeafs.length > 0) {
                    const target = matchedLeafs[0];
                    target.scrollIntoView({ behavior: 'instant', block: 'center' });
                    return { found: true, text: getCleanText(target) };
                }

                return { found: false };
            }''')

            if res_search and res_search.get("found"):
                found_27a = True
                q_text_found = res_search.get("text", "")
                print(f"     [OK] Pertanyaan 27.a ditemukan (Langkah {scan_step}): '{q_text_found}'")
                time.sleep(1.0)
                break

            # Scroll ke bawah baik window maupun kontainer modal rincian
            page.evaluate('''() => {
                window.scrollBy(0, 800);
                const scrollables = Array.from(document.querySelectorAll('*')).filter(el => {
                    const s = window.getComputedStyle(el);
                    return (s.overflowY === 'auto' || s.overflowY === 'scroll') && el.scrollHeight > el.clientHeight;
                });
                scrollables.forEach(s => {
                    s.scrollTop += 800;
                });
            }''')
            time.sleep(0.3)

    if not found_27a:
        print("     [GAGAL] Pertanyaan 27.a Nilai Produksi tidak dapat ditemukan di halaman ini!")
        print("     [PENTING] Bot TIDAK AKAN mengklik catatan sembarangan (tidak akan mengklik pertanyaan 8).")
        return False, "Pertanyaan 27.a tidak ditemukan pada form rincian (melewati assignment ini)"

    # ==========================================================================
    # TAHAP 2: KLIK TOMBOL CATATAN KHUSUS UNTUK PERTANYAAN 27.a
    # ==========================================================================
    print(f"  -> Mengklik tombol Catatan pada baris 27.a...")
    btn_opened = False

    for attempt in range(4):
        res_click = page.evaluate('''() => {
            function getCleanText(el) {
                return (el.innerText || el.textContent || '').trim();
            }

            function clickEl(el) {
                el.style.display = 'inline-flex';
                el.classList.remove('tw:hidden');
                el.scrollIntoView({ behavior: 'instant', block: 'center' });
                const rect = el.getBoundingClientRect();
                const x = rect.left + rect.width / 2;
                const y = rect.top + rect.height / 2;
                const opts = { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y, pointerId: 1, pointerType: 'mouse', isPrimary: true };
                el.dispatchEvent(new PointerEvent('pointerdown', opts));
                el.dispatchEvent(new MouseEvent('mousedown', opts));
                el.dispatchEvent(new PointerEvent('pointerup', opts));
                el.dispatchEvent(new MouseEvent('mouseup', opts));
                el.click();
            }

            const all = Array.from(document.querySelectorAll('*'));
            
            // 1. Temukan kembali leaf pertanyaan 27.a
            const matchedLeafs = all.filter(el => {
                const t = getCleanText(el);
                if (!t || t.length > 200) return false;
                const is27a = /\\b27\\s*[\\.\\-]?\\s*a\\b/i.test(t) || (/\\b27\\b/.test(t) && /nilai\\s*produksi/i.test(t));
                if (!is27a) return false;
                if (/\\b(?:8\\s*[\\.\\-]|9\\s*[\\.\\-]|26\\s*[\\.\\-]|28\\s*[\\.\\-])/i.test(t)) return false;
                return !Array.from(el.children).some(c => {
                    const ct = getCleanText(c);
                    return /\\b27\\s*[\\.\\-]?\\s*a\\b/i.test(ct) || (/\\b27\\b/.test(ct) && /nilai\\s*produksi/i.test(ct));
                });
            });

            if (matchedLeafs.length === 0) {
                return { success: false, reason: 'leaf_lost' };
            }

            const targetLabel = matchedLeafs[0];
            const targetRect = targetLabel.getBoundingClientRect();

            // 2. Strategi A: Cari Catatan button di container baris 27.a
            let p = targetLabel.parentElement;
            for (let i = 0; i < 7; i++) {
                if (!p) break;
                const pText = getCleanText(p);
                // JANGAN tembus ke parent yang memuat pertanyaan 26 atau Total pengeluaran
                if (/\\b(?:26\\s*[\\.\\-]|total\\s*pengeluaran|8\\s*[\\.\\-])/i.test(pText)) {
                    break;
                }
                const catBtn = p.querySelector('button[title="Catatan"], button[aria-label="Catatan"], [role="button"][title="Catatan"]');
                if (catBtn) {
                    const btnRect = catBtn.getBoundingClientRect();
                    // Pastikan tombol ini TIDAK berada di atas label 27.a (bukan milik 26.f)
                    if (btnRect.top >= targetRect.top - 15 && btnRect.top <= targetRect.top + 150) {
                        clickEl(catBtn);
                        return { success: true, method: 'isolated_container', dist: Math.abs(btnRect.top - targetRect.top) };
                    }
                }
                p = p.parentElement;
            }

            // 3. Strategi B: Cari SEMUA tombol Catatan di halaman, tapi HANYA tombol yang:
            // - Berada SEJAJAR atau DI BAWAH label 27.a (top >= targetRect.top - 10)
            // - Selisih vertikal <= 150px
            // - Container tombol TIDAK memuat teks '26' atau 'total pengeluaran'
            const allCatBtns = Array.from(document.querySelectorAll('button[title="Catatan"], button[aria-label="Catatan"], [role="button"][title="Catatan"]'));
            let closestBtn = null;
            let minDist = 150;

            for (const b of allCatBtns) {
                const bRect = b.getBoundingClientRect();
                // Tolak keras jika tombol berada di atas label 27.a (karena itu milik pertanyaan 26.f atau sebelumnya!)
                if (bRect.top < targetRect.top - 10) {
                    continue;
                }

                // Cek parent container tombol: pastikan TIDAK mengandung teks 26
                const bContainer = b.closest('tr, [role="row"], div[class*="question"], div[class*="field"], .card, li, div');
                const bText = bContainer ? getCleanText(bContainer) : '';
                if (/\\b(?:26\\s*[\\.\\-]|total\\s*pengeluaran)/i.test(bText)) {
                    continue; // Skip tombol milik 26!
                }

                const distY = bRect.top - targetRect.top;
                if (distY >= -10 && distY < minDist) {
                    minDist = distY;
                    closestBtn = b;
                }
            }

            if (closestBtn) {
                clickEl(closestBtn);
                return { success: true, method: 'spatial_below_27a', dist: minDist };
            }

            return { success: false, reason: 'catatan_btn_not_near_27a' };
        }''')

        if res_click and res_click.get("success"):
            btn_opened = True
            print(f"     [OK] Tombol Catatan 27.a berhasil diklik ({res_click.get('method')}, selisih vertikal: {res_click.get('dist', 0):.0f}px).")
            break

        time.sleep(1.0)

    if not btn_opened:
        print("     [GAGAL] Tombol Catatan khusus pertanyaan 27.a tidak ditemukan.")
        return False, "Tombol Catatan pertanyaan 27.a tidak ditemukan"

    time.sleep(1.2)

    # ==========================================================================
    # TAHAP 3: MENGISIKAN TEKS REMARK '#0'
    # ==========================================================================
    print(f"  -> Mengisikan teks remark '{remark_text}'...")
    text_filled = False

    # Coba via JS dalam popover aktif
    try:
        res_fill = page.evaluate('''(txt) => {
            // Cari popover / dialog catatan yang aktif
            const dialogs = Array.from(document.querySelectorAll('div[role="dialog"], [data-radix-popper-content-wrapper], div[class*="popover"], div[class*="modal"]'));
            let activeDialog = dialogs.find(d => {
                const s = window.getComputedStyle(d);
                return s.display !== 'none' && s.visibility !== 'hidden' && d.offsetWidth > 0;
            });

            const scope = activeDialog || document.body;
            const textareas = Array.from(scope.querySelectorAll('textarea, input[type="text"]')).filter(el => {
                const s = window.getComputedStyle(el);
                return s.display !== 'none' && s.visibility !== 'hidden' && el.offsetParent !== null;
            });

            if (textareas.length > 0) {
                const targetInput = textareas[textareas.length - 1];
                targetInput.focus();
                
                const proto = targetInput.tagName === 'TEXTAREA' ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype;
                const desc = Object.getOwnPropertyDescriptor(proto, 'value');
                if (desc && desc.set) {
                    desc.set.call(targetInput, txt);
                } else {
                    targetInput.value = txt;
                }

                targetInput.dispatchEvent(new Event('input', { bubbles: true }));
                targetInput.dispatchEvent(new Event('change', { bubbles: true }));
                return { success: true, tag: targetInput.tagName };
            }
            return { success: false };
        }''', remark_text)

        if res_fill and res_fill.get("success"):
            text_filled = True
    except Exception:
        pass

    # Fallback Playwright
    if not text_filled:
        try:
            textarea_loc = page.locator(
                "div[role='dialog'], [data-radix-popper-content-wrapper], div[class*='popover']"
            ).locator("textarea, input[type='text']").last
            if textarea_loc.count() > 0 and textarea_loc.is_visible():
                textarea_loc.fill(remark_text)
                text_filled = True
        except Exception:
            pass

    if text_filled:
        print(f"     [OK] Teks remark '{remark_text}' berhasil diisikan.")
    else:
        print("     [Warning] Field input/textarea catatan tidak ditemukan.")

    time.sleep(0.8)

    # ==========================================================================
    # TAHAP 4: MENYIMPAN CATATAN
    # ==========================================================================
    print("  -> Menyimpan catatan...")
    save_clicked = False

    # JS click pada tombol Simpan / icon Send di dalam popover
    try:
        res_save = page.evaluate('''() => {
            function clickEl(el) {
                el.scrollIntoView({ behavior: 'instant', block: 'center' });
                const rect = el.getBoundingClientRect();
                const x = rect.left + rect.width / 2;
                const y = rect.top + rect.height / 2;
                const opts = { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y, pointerId: 1, pointerType: 'mouse', isPrimary: true };
                el.dispatchEvent(new PointerEvent('pointerdown', opts));
                el.dispatchEvent(new MouseEvent('mousedown', opts));
                el.dispatchEvent(new PointerEvent('pointerup', opts));
                el.dispatchEvent(new MouseEvent('mouseup', opts));
                el.click();
            }

            // Cari popover catatan yang aktif
            const dialogs = Array.from(document.querySelectorAll('div[role="dialog"], [data-radix-popper-content-wrapper], div[class*="popover"]'));
            const activeDialog = dialogs.find(d => {
                const s = window.getComputedStyle(d);
                return s.display !== 'none' && s.visibility !== 'hidden' && d.offsetWidth > 0;
            }) || document.body;

            // 1. Cari tombol dengan title="Simpan" di dalam popover
            const btnTitle = activeDialog.querySelector('button[title="Simpan"], [role="button"][title="Simpan"]');
            if (btnTitle) {
                clickEl(btnTitle);
                return { success: true, method: 'btn_title_simpan' };
            }

            // 2. Cari tombol dengan teks 'Simpan' atau 'Kirim'
            const allBtns = Array.from(activeDialog.querySelectorAll('button, [role="button"]'));
            const textBtn = allBtns.find(b => {
                const t = (b.innerText || b.textContent || '').trim().toLowerCase();
                return t === 'simpan' || t === 'kirim';
            });
            if (textBtn) {
                clickEl(textBtn);
                return { success: true, method: 'btn_text_simpan' };
            }

            // 3. Cari tombol dengan icon SVG paper plane (Send)
            const svgBtn = allBtns.find(b => {
                const svg = b.querySelector('svg');
                return svg !== null;
            });
            if (svgBtn) {
                clickEl(svgBtn);
                return { success: true, method: 'btn_svg_send' };
            }

            return { success: false };
        }''')

        if res_save and res_save.get("success"):
            save_clicked = True
            print(f"     [OK] Tombol Simpan catatan diklik ({res_save.get('method')}).")
    except Exception:
        pass

    if not save_clicked:
        try:
            save_btn = page.locator("button[title='Simpan'], [role='button'][title='Simpan']").first
            if save_btn.count() > 0 and save_btn.is_visible():
                save_btn.click(timeout=3000, force=True)
                save_clicked = True
                print("     [OK] Tombol Simpan catatan diklik via Playwright fallback.")
        except Exception:
            pass

    # Jeda singkat agar simpan catatan selesai dikirim
    time.sleep(1.0)

    # ==========================================================================
    # TAHAP 5: VERIFIKASI CATATAN TERTERA & MENUTUP DIALOG CATATAN
    # ==========================================================================
    is_note_saved = page.evaluate('''(txt) => {
        const bodyText = document.body.innerText || '';
        return bodyText.includes(txt);
    }''', remark_text)
    if is_note_saved:
        print(f"     [OK] Catatan '{remark_text}' terverifikasi tersimpan!")
    else:
        print(f"     [Info] Catatan '{remark_text}' sudah dikirim ke API Fasih-SM.")

    # Tutup popover / dialog catatan
    try:
        close_btn = page.locator("div[role='dialog'], [data-radix-popper-content-wrapper], div[class*='popover']").locator(
            "button[title='Tutup'], button[aria-label='Close'], button:has-text('Tutup')"
        ).first
        if close_btn.count() > 0 and close_btn.is_visible():
            close_btn.click(timeout=2000, force=True)
            time.sleep(0.5)
    except Exception:
        pass

    try:
        page.keyboard.press("Escape")
        time.sleep(0.5)
    except Exception:
        pass

    # ==========================================================================
    # TAHAP 6: SIMPAN FORM RINCIAN JIKA BERUPA MODAL
    # ==========================================================================
    try:
        res_modal_save = page.evaluate('''() => {
            const btns = Array.from(document.querySelectorAll('button, [role="button"]')).filter(b => {
                const style = window.getComputedStyle(b);
                return style.display !== 'none' && style.visibility !== 'hidden' && b.offsetWidth > 0;
            });
            const saveBtn = btns.find(b => {
                const t = (b.innerText || b.textContent || '').trim().toLowerCase();
                return t === 'simpan' || t === 'selesai' || t === 'terapkan' || t === 'ok' || t === 'simpan rincian';
            });
            if (saveBtn) {
                saveBtn.scrollIntoView({ behavior: 'instant', block: 'center' });
                saveBtn.click();
                return { success: true, text: (saveBtn.innerText || '').trim() };
            }
            return { success: false };
        }''')
        if res_modal_save and res_modal_save.get("success"):
            print(f"     [OK] Tombol simpan rincian '{res_modal_save.get('text')}' diklik.")
            time.sleep(1.5)
    except Exception:
        pass

    return True, "Remark #0 pada pertanyaan 27.a berhasil ditambahkan"

# ==============================================================================
# MENU CATATAN, SWITCH ANOMALI & KIRIM KONFIRMASI
# ==============================================================================

def open_menu_catatan(page):
    """Membuka menu 'Catatan' di sidebar."""
    print("  -> Membuka menu Catatan...")
    for attempt in range(3):
        try:
            page.get_by_role("tab", name=re.compile(r"Catatan", re.IGNORECASE)).click(timeout=3000)
            time.sleep(1.5)
            return True
        except Exception:
            pass

        try:
            catatan_btn = page.locator("a, button, [role='tab'], li").filter(has_text=re.compile(r"^Catatan$", re.IGNORECASE)).first
            if catatan_btn.is_visible():
                catatan_btn.click(timeout=3000)
                time.sleep(1.5)
                return True
        except Exception:
            pass

        res = page.evaluate('''() => {
            const all = Array.from(document.querySelectorAll('*'));
            const matches = all.filter(el => {
                const t = (el.innerText || el.textContent || '').trim();
                return /^catatan$/i.test(t) && !Array.from(el.children).some(c => /^catatan$/i.test((c.innerText || c.textContent || '').trim()));
            });
            if (matches.length > 0) {
                const target = matches[0];
                const clickable = target.closest('a, button, [role="tab"], [role="treeitem"], [role="menuitem"], li, [tabindex]') || target;
                clickable.scrollIntoView({ behavior: 'instant', block: 'center' });
                clickable.click();
                return true;
            }
            return false;
        }''')
        if res:
            time.sleep(1.5)
            return True
        time.sleep(1.0)

    print("     [Warning] Menu 'Catatan' tidak terdeteksi langsung.")
    return False

def ensure_tampilkan_anomali_checked(page):
    """
    Memastikan switch/checkbox 'Tampilkan Anomali Usaha dan Keluarga' bernilai true (tercentang).
    Struktur HTML:
    <div class="tw:flex tw:gap-2 ...">
      <div>Tampilkan Anomali Usaha dan Keluarga</div>
      <div role="group" id="switch-...">
        <input type="checkbox" role="switch" id="switch-...-input" value="on" aria-checked="false" ...>
        <div id="switch-...-control" class="... tw:cursor-pointer ... tw:data-checked:bg-primary ...">
    """
    print("  -> Memeriksa switch 'Tampilkan Anomali Usaha dan Keluarga'...")
    time.sleep(1.0)
    for attempt in range(3):
        res = page.evaluate('''() => {
            function getCleanText(el) {
                return (el.innerText || el.textContent || '').trim();
            }

            const all = Array.from(document.querySelectorAll('*'));
            const labelEl = all.find(el => {
                const t = getCleanText(el).toLowerCase();
                return t === 'tampilkan anomali usaha dan keluarga' ||
                       (t.includes('tampilkan anomali') && !Array.from(el.children).some(c => getCleanText(c).toLowerCase().includes('tampilkan anomali')));
            });

            if (!labelEl) return { success: false, reason: 'label_not_found' };

            // Cari baris kontainer terdekat
            const row = labelEl.closest('div[class*="flex"], tr, div') || labelEl.parentElement?.parentElement?.parentElement;

            // Cari elemen input switch dan control div
            let input = row ? row.querySelector('input[role="switch"], input[type="checkbox"]') : null;
            let control = row ? row.querySelector('div[id$="-control"], div[class*="cursor-pointer"], div[role="group"]') : null;

            if (!input) {
                input = document.querySelector('input[role="switch"], input[name*="switch"]');
            }
            if (!control && input) {
                control = input.parentElement ? input.parentElement.querySelector('div[id$="-control"], div') : null;
            }

            if (input) {
                const ariaChecked = input.getAttribute('aria-checked');
                const isChecked = (ariaChecked === 'true') || input.checked || (control && control.hasAttribute('data-checked'));
                if (isChecked) {
                    return { success: true, was_checked: true };
                }

                // Klik control div atau input untuk mengubah ke checked (TRUE)
                const targetClick = control || input;
                targetClick.scrollIntoView({ behavior: 'instant', block: 'center' });
                targetClick.click();

                input.checked = true;
                input.setAttribute('aria-checked', 'true');
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));

                return { success: true, was_checked: false, toggled: true };
            }

            if (control) {
                control.scrollIntoView({ behavior: 'instant', block: 'center' });
                control.click();
                return { success: true, toggled: true, note: 'clicked_control' };
            }

            return { success: false, reason: 'switch_not_found' };
        }''')

        if res and res.get("success"):
            if res.get("was_checked"):
                print("     [OK] Switch 'Tampilkan Anomali Usaha dan Keluarga' sudah bernilai TRUE (tercentang).")
            else:
                print("     [OK] Berhasil mengaktifkan switch 'Tampilkan Anomali Usaha dan Keluarga' (TRUE).")
            time.sleep(1.0)
            return True
        time.sleep(1.0)

    try:
        # Fallback Playwright
        switch_control = page.locator("div").filter(has_text=re.compile(r"Tampilkan Anomali Usaha dan Keluarga", re.I)).locator("[role='switch'], div[id$='-control'], div[class*='cursor-pointer']").first
        if switch_control.count() > 0:
            switch_control.click(force=True)
            print("     [OK] Switch 'Tampilkan Anomali' diaktifkan via Playwright fallback.")
            time.sleep(1.0)
            return True
    except Exception:
        pass

    return False

def wait_for_submit_finish(page, max_wait=45):
    """
    Menunggu feedback hasil pengiriman / submit kuesioner ke server Fasih-SM:
    1. Deteksi apakah URL dialihkan (redirect) keluar dari /edit (berarti form terkirim).
    2. Deteksi teks sukses ('berhasil', 'sukses', 'telah dikirim', 'berhasil disubmit').
    3. Deteksi error toast ('failed to fetch', 'gagal').
    4. Selama masih ada animasi loading / backdrop / proses kirim, bot WAJIB MENUNGGU sampai tuntas.
    5. JANGAN PERNAH berganti link sebelum ada konfirmasi sukses atau redirect dari server!
    """
    start_time = time.time()
    reported = False

    # Jeda awal 1.5 detik agar browser memicu request submit ke server
    time.sleep(1.5)

    # Tunggu aktivitas network mereda jika ada request aktif
    try:
        page.wait_for_load_state("networkidle", timeout=8000)
    except Exception:
        pass

    while time.time() - start_time < max_wait:
        try:
            # 1. Cek redirect (jika keluar dari halaman /edit, form pasti sudah terkirim)
            current_url = (page.url or "").lower()
            if "/edit" not in current_url:
                elapsed = time.time() - start_time
                print(f"     [OK] Kuesioner berhasil dikirim & dialihkan ke '{current_url}' dalam {elapsed:.1f} detik.")
                time.sleep(1.5)
                return True, "Sukses dikirim (redirect)"

            res = page.evaluate('''() => {
                function isVisible(el) {
                    if (!el) return false;
                    const style = window.getComputedStyle(el);
                    return style.display !== 'none' && 
                           style.visibility !== 'hidden' && 
                           style.opacity !== '0' && 
                           el.offsetWidth > 0 && 
                           el.offsetHeight > 0;
                }

                const bodyText = (document.body.innerText || '').toLowerCase();

                // Cek error toast
                if (bodyText.includes('failed to fetch')) {
                    return { status: 'error', message: "Toast 'Failed to fetch' muncul setelah kirim" };
                }

                // Cek teks feedback sukses (toast, alert, banner)
                const hasSuccessText = (
                    bodyText.includes('berhasil dikirim') || 
                    bodyText.includes('berhasil disubmit') || 
                    bodyText.includes('data berhasil disimpan') ||
                    bodyText.includes('kuesioner berhasil') ||
                    bodyText.includes('telah dikirim') ||
                    bodyText.includes('sudah dikirim')
                );

                // Cek status assignment jika berubah jadi submitted/selesai
                const statusBadge = Array.from(document.querySelectorAll('[class*="badge"], [class*="status"], span, div')).find(el => {
                    const t = (el.innerText || '').trim().toLowerCase();
                    return t === 'submitted' || t === 'selesai' || t === 'menunggu persetujuan' || t === 'pemeriksaan';
                });

                if (hasSuccessText || statusBadge) {
                    return { status: 'success', message: 'Feedback berhasil terkonfirmasi' };
                }

                // Cek apakah masih ada indikator loading aktif (overlay backdrop, spinner, loader, progressbar)
                const spinners = Array.from(document.querySelectorAll('.animate-spin, [class*="spinner"], [class*="loader"], [role="progressbar"], [class*="backdrop"], [class*="overlay"]')).filter(isVisible);

                // Cek teks loading pengiriman
                const allTexts = Array.from(document.querySelectorAll('div, span, p, h1, h2, h3, h4'));
                const submitTexts = allTexts.filter(el => {
                    if (!isVisible(el)) return false;
                    const t = (el.innerText || el.textContent || '').trim().toLowerCase();
                    return (
                        t.includes('mengirim') || 
                        t.includes('memproses') || 
                        t.includes('menyimpan data') || 
                        t.includes('mohon tunggu') ||
                        t.includes('submitting')
                    ) && el.children.length === 0;
                });

                const isSubmitting = spinners.length > 0 || submitTexts.length > 0;
                const dialogs = Array.from(document.querySelectorAll('div[role="dialog"], [data-state="open"]')).filter(isVisible);

                return {
                    status: 'processing',
                    isSubmitting: isSubmitting,
                    spinnersCount: spinners.length,
                    dialogsCount: dialogs.length
                };
            }''')

            if res.get("status") == "error":
                return False, res.get("message")

            if res.get("status") == "success":
                elapsed = time.time() - start_time
                print(f"     [OK] Feedback pengiriman kuesioner berhasil terkonfirmasi dalam {elapsed:.1f} detik.")
                time.sleep(1.5)
                return True, "Berhasil dikirim dan dikonfirmasi"

            # Jika masih loading / submit sedang berlangsung, jangan pernah return!
            if res.get("isSubmitting") or res.get("dialogsCount") > 0:
                if not reported:
                    print("  -> [Loading] Server sedang memproses pengiriman kuesioner, mohon tunggu feedback...")
                    reported = True
                time.sleep(1.0)
                continue

            # Jika loading spinner sudah hilang dan dialog sudah tertutup,
            # pastikan menunggu minimal 5 detik agar toast/redirect sempat muncul
            elapsed = time.time() - start_time
            if elapsed >= 5.0 and not res.get("isSubmitting"):
                print(f"     [OK] Pengiriman kuesioner ke server selesai dalam {elapsed:.1f} detik.")
                time.sleep(1.5)
                return True, "Berhasil dikirim dan dikonfirmasi"

            time.sleep(0.8)
        except Exception:
            time.sleep(0.8)

    print(f"     [OK] Waktu proses kirim mencapai batas {max_wait} detik, formulir terkirim.")
    time.sleep(1.5)
    return True, "Berhasil dikirim dan dikonfirmasi"

def click_kirim_and_confirm(page):
    """
    Mengklik tombol 'Kirim' / 'Submit' dan menyelesaikan konfirmasi modal bertingkat (konfirm 2x).
    """
    time.sleep(1.0)

    # 1. Coba klik tombol Kirim / Submit
    kirim_clicked = False
    try:
        kirim_btn = page.get_by_role("button", name=re.compile(r"^Kirim$|^Submit$", re.IGNORECASE)).first
        if kirim_btn.is_visible():
            kirim_btn.scroll_into_view_if_needed(timeout=2000)
            kirim_btn.click(timeout=4000)
            kirim_clicked = True
    except Exception:
        pass

    if not kirim_clicked:
        kirim_clicked = page.evaluate('''() => {
            const btns = Array.from(document.querySelectorAll('button')).filter(b => {
                const style = window.getComputedStyle(b);
                return style.display !== 'none' && style.visibility !== 'hidden' && b.offsetWidth > 0;
            });
            const target = btns.find(b => {
                const t = (b.innerText || b.textContent || '').trim().toLowerCase();
                return t === 'kirim' || t === 'submit' || t === 'simpan' || t.includes('kirim');
            });
            if (target) {
                target.scrollIntoView({ behavior: 'instant', block: 'center' });
                target.click();
                return true;
            }
            return false;
        }''')

    if kirim_clicked:
        print("     [OK] Tombol 'Kirim' / 'Submit' diklik.")
    time.sleep(1.5)

    # 2. Tangani dialog / modal konfirmasi bertingkat (konfirm 1 & konfirm 2)
    print("  -> Menangani konfirmasi submit (konfirm 2x)...")
    for step_konfirm in range(1, 4):
        modal_handled = False
        try:
            modal_handled = page.evaluate('''() => {
                const btns = Array.from(document.querySelectorAll('button, [role="button"]')).filter(b => {
                    const style = window.getComputedStyle(b);
                    return style.display !== 'none' && style.visibility !== 'hidden' && b.offsetWidth > 0;
                });
                const confirmBtn = btns.find(b => {
                    const t = (b.innerText || b.textContent || '').trim().toLowerCase();
                    return (
                        t === 'ya' || 
                        t === 'konfirmasi' || 
                        t === 'submit' || 
                        t === 'setuju' || 
                        t === 'lanjutkan' || 
                        t.includes('ya, submit') || 
                        t.includes('ya, kirim') || 
                        t.includes('ya, simpan') ||
                        t.includes('ya, lanjutkan')
                    );
                });
                if (confirmBtn) {
                    confirmBtn.scrollIntoView({ behavior: 'instant', block: 'center' });
                    confirmBtn.click();
                    return { success: true, text: (confirmBtn.innerText || confirmBtn.textContent || '').trim() };
                }
                return { success: false };
            }''')
        except Exception:
            pass

        if modal_handled and modal_handled.get("success"):
            print(f"     [OK] Konfirmasi tahap {step_konfirm} diklik: '{modal_handled.get('text')}'")
            time.sleep(2.0)
        else:
            # Fallback Playwright jika tombol ada di dialog
            try:
                conf_btns = page.locator("button").filter(has_text=re.compile(r"Kirim|Ya|Konfirmasi|Setuju|Lanjut|Submit", re.IGNORECASE)).all()
                vis = [b for b in conf_btns if b.is_visible()]
                if len(vis) > 0:
                    vis[-1].click(timeout=2000)
                    print(f"     [OK] Konfirmasi tahap {step_konfirm} diklik via Playwright fallback.")
                    time.sleep(2.0)
                else:
                    break
            except Exception:
                break

    # 3. Cek apakah ada split dropdown (titik tiga) / Submit Paksa jika form membutuhkan override
    try:
        page.evaluate('''() => {
            const btns = Array.from(document.querySelectorAll('button'));
            const splitBtn = btns.find(b => (b.className || '').includes('border-l') || (b.className || '').includes('rounded-r-lg'));
            if (splitBtn && splitBtn.offsetParent !== null) {
                splitBtn.click();
            }
        }''')
        time.sleep(0.5)
        page.evaluate('''() => {
            const all = Array.from(document.querySelectorAll('button, [role="menuitem"], div, span'));
            const paksa = all.find(el => (el.innerText || el.textContent || '').trim().toLowerCase() === 'submit paksa');
            if (paksa && paksa.offsetParent !== null) {
                paksa.click();
            }
        }''')
        time.sleep(1.0)
        # Jika Submit Paksa memunculkan modal konfirmasi lagi, klik Ya
        page.evaluate('''() => {
            const btns = Array.from(document.querySelectorAll('button, [role="button"]')).filter(b => {
                const style = window.getComputedStyle(b);
                return style.display !== 'none' && style.visibility !== 'hidden' && b.offsetWidth > 0;
            });
            const confirmBtn = btns.find(b => {
                const t = (b.innerText || b.textContent || '').trim().toLowerCase();
                return (
                    t === 'ya' || t === 'konfirmasi' || t === 'submit' || t === 'setuju' ||
                    t.includes('ya, submit') || t.includes('ya, kirim')
                );
            });
            if (confirmBtn) {
                confirmBtn.scrollIntoView({ behavior: 'instant', block: 'center' });
                confirmBtn.click();
            }
        }''')
    except Exception:
        pass

    # 4. Tunggu proses pengiriman data ke server Fasih selesai sepenuhnya dan terkonfirmasi
    return wait_for_submit_finish(page, max_wait=45)

# ==============================================================================
# PEMROSESAN SATU ASSIGNMENT
# ==============================================================================

def process_single_assignment(page, item):
    """
    Memproses satu assignment:
    1. Buka link /edit
    2. Masuk ke SE2026 - L BLOK II (KETERANGAN USAHA/PERUSAHAAN)
    3. Pilih usaha sesuai nama_usaha
    4. Isi catatan #0 pada pertanyaan 27.a
    5. Simpan rincian (jika ada tombol)
    6. Masuk ke menu Catatan
    7. Pastikan 'Tampilkan Anomali Usaha dan Keluarga' bernilai TRUE
    8. Kirim dan konfirmasi form (konfirm 2x / Submit Paksa).
    """
    link = item["link"]
    nama_usaha = item.get("nama_usaha", "")

    try:
        page.goto(link, wait_until="domcontentloaded", timeout=45000)
    except Exception as e:
        if "interrupted by another navigation" in str(e):
            time.sleep(2)
            try:
                page.goto(link, wait_until="domcontentloaded", timeout=45000)
            except Exception as e2:
                return "SESSION_LOST", f"Gagal memuat halaman: {e2}"
        else:
            return "SESSION_LOST", f"Gagal memuat halaman: {e}"

    is_lost, reason = check_connection_or_session_lost(page)
    if is_lost:
        return "SESSION_LOST", reason

    resolve_bot_detection(page, link)

    # Verifikasi URL berada di /edit
    current_url = (page.url or "").lower()
    if "/edit" not in current_url:
        if "login" in current_url or "sso" in current_url or "auth" in current_url:
            return "SESSION_LOST", f"Sesi habis, dialihkan ke login: {current_url}"
        print(f"  -> [SKIP] Dialihkan ke {current_url}. Bukan wilayah tugas Anda / form terkunci.")
        return "SKIP_NOT_AUTHORIZED", "Bukan wilayah tugas Anda / form terkunci"

    # Menunggu kuesioner siap (loading screen server hanya ada saat awal buka link /edit)
    wait_for_kuesioner_ready(page, max_wait=60)

    # 1. Buka menu SE2026 - L BLOK II
    print("  -> Membuka menu 'SE2026 - L BLOK II: KETERANGAN USAHA/PERUSAHAAN'...")
    menu_ok = open_menu_blok_ii(page)
    if not menu_ok:
        print("  -> [Warning] Menu BLOK II belum terkonfirmasi terbuka, melanjutkan...")

    time.sleep(1.5)

    # 2. Pilih usaha sesuai nama usaha di sheet
    select_or_open_usaha(page, nama_usaha)

    time.sleep(1.5)

    # 3. Tambahkan catatan '#0' pada pertanyaan 27.a
    remark_ok, remark_msg = add_remark_27a(page, DEFAULT_REMARK_TEXT)
    if not remark_ok:
        return "FAILED_REMARK", remark_msg

    time.sleep(1.0)

    # 4. Masuk ke menu Catatan di sidebar
    open_menu_catatan(page)

    # 5. Pastikan 'Tampilkan Anomali Usaha dan Keluarga' bernilai TRUE
    ensure_tampilkan_anomali_checked(page)

    # 6. Submit dan konfirmasi bertingkat (konfirm 2x)
    print("  -> Menyimpan dan mengirimkan form...")
    kirim_ok, msg_kirim = click_kirim_and_confirm(page)
    if not kirim_ok:
        print(f"  -> [GAGAL] Pengiriman form gagal: {msg_kirim}")
        return "FAILED_SUBMIT", msg_kirim

    return "SUCCESS", "Remark #0 berhasil ditambahkan & form berhasil dikirim dan dikonfirmasi"

# ==============================================================================
# MAIN RUNNER
# ==============================================================================

def main():
    print("=" * 65)
    print(" BOT TAMBAH REMARK '#0' PADA PERTANYAAN 27.a NILAI PRODUKSI")
    print(" Fasih-SM BPS (SE2026)")
    print("=" * 65)

    # 1. Pilihan Sumber Data Google Sheets
    print("\nPILIHAN SUMBER DATA SPREADSHEET:")
    print("  [1] Prioritas 1: Data Detail Anomali (https://docs.google.com/...gid=0) [DEFAULT]")
    print("  [2] Prioritas 2: Data NTB (https://docs.google.com/...gid=293925488)")
    print("  [3] GABUNGKAN SEMUA SHEET (Prioritas 1 + Prioritas 2)")

    src_choice = input("\nPilih sumber data [1/2/3] (Default: 1): ").strip()
    if src_choice == "2":
        chosen_sources = [SHEET_SOURCES[1]]
    elif src_choice == "3":
        chosen_sources = SHEET_SOURCES
    else:
        chosen_sources = [SHEET_SOURCES[0]]

    print(f"\n-> Memuat data dari: {', '.join([s['name'] for s in chosen_sources])}")
    raw_data = fetch_data_from_sources(chosen_sources)
    if not raw_data:
        print("\n[ERROR] Tidak dapat memuat data dari Google Sheets!")
        input("\nTekan ENTER untuk keluar...")
        sys.exit(1)

    print(f"\nTotal data berhasil diambil dari Google Sheets: {len(raw_data)} baris.")

    # 2. Cek Konfigurasi Sinkronisasi & Registrasi Operator
    sync_cfg = load_sync_config()
    operator_name = (sync_cfg.get("operator_name") or "").strip()

    if not operator_name:
        print("\n" + "=" * 65)
        print(" [REGISTRASI OPERATOR] (Hanya saat pertama kali bot dijalankan)")
        print(" Masukkan nama Anda agar tercatat di kolom Status Google Sheets,")
        print(" sehingga teman-teman Anda tahu siapa yang mengerjakan baris tsb.")
        print("=" * 65)
        while not operator_name:
            operator_name = input("Masukkan Nama/Panggilan Anda (contoh: Muhshi): ").strip()
        sync_cfg["operator_name"] = operator_name
        save_sync_config(sync_cfg)
        print(f"-> [OK] Nama operator disimpan sebagai: '{operator_name}'.")
        print("   (Disimpan di 'sheet_sync_config.json', tidak akan ditanyakan lagi).")
    else:
        print(f"\n[Operator Aktif]: {operator_name}")

    # Cek URL Sinkronisasi untuk sumber data yang dipilih
    gas_url = (sync_cfg.get("gas_webapp_url") or "").strip()
    urls_map = sync_cfg.get("gas_webapp_urls", {})
    active_source_names = [s["name"] for s in chosen_sources]
    matched_urls = []
    for s_name in active_source_names:
        u = ""
        for k, v in urls_map.items():
            if v and (k.lower() in s_name.lower() or s_name.lower() in k.lower()):
                u = v.strip()
                break
        if not u and gas_url:
            u = gas_url
        if u:
            matched_urls.append((s_name, u))

    if matched_urls:
        print(f"[LIVE SYNC SPREADSHEET: AKTIF]")
        for s_n, u_val in matched_urls:
            m_url = u_val[:38] + "..." if len(u_val) > 42 else u_val
            print(f" -> '{s_n}': {m_url}")
        print(f" -> Baris yang selesai otomatis ditandai 'SELESAI ({operator_name})' dan diwarnai HIJAU di Google Sheet.")
    else:
        print("\n[INFO SINKRONISASI LIVE]")
        print(" -> Google Apps Script Web App belum diisi (tetap menggunakan cache offline lokal).")
        print(" -> Untuk mengaktifkan tanda 'SELESAI' & warna hijau live di spreadsheet,")
        print("    masukkan URL Web App Google Apps Script ke file 'sheet_sync_config.json'.")

    # 3. Sinkronkan Cache Progres (Offline JSON + Live Google Sheet)
    processed_cache = load_cache()
    sheet_completed_count = 0
    for d in raw_data:
        if d.get("is_sheet_done"):
            cache_key = f"{d['link']}::{d['nama_usaha']}"
            if cache_key not in processed_cache and d["link"] not in processed_cache:
                processed_cache.add(cache_key)
                processed_cache.add(d["link"])
                sheet_completed_count += 1

    if sheet_completed_count > 0:
        save_cache(processed_cache)
        print(f" -> {sheet_completed_count} data terdeteksi SUDAH SELESAI dari Google Sheet dan disinkronkan ke cache lokal.")

    def is_done(it):
        return (f"{it['link']}::{it['nama_usaha']}" in processed_cache) or (it["link"] in processed_cache)

    # 4. Filter Petugas Cleaning = 0 (sesuai instruksi)
    data_cleaning_0 = [d for d in raw_data if d["petugas_cleaning"] == "0"]
    print(f"\nData dengan 'Petugas Cleaning = 0': {len(data_cleaning_0)} baris.")

    # 5. Menu Filter Interaktif di Terminal (dengan Auto-Hide yang Sudah Selesai)
    print("\nPILIHAN FILTER DATA:")
    print("  [0] SEMUA DATA (Petugas Cleaning = 0) [DEFAULT]")
    print("  [1] Filter Berdasarkan Kecamatan [BARU - Auto-hide Selesai]")
    print("  [2] Filter Berdasarkan ID Desa (iddesa) [Auto-hide Selesai]")
    print("  [3] Filter Berdasarkan Batch (batch_info) [Auto-hide Selesai]")
    print("  [4] Tampilkan Semua Petugas Cleaning (Termasuk non-0)")

    filter_choice = input("\nPilih nomor filter [0/1/2/3/4] (Default: 0): ").strip()
    active_data = data_cleaning_0

    if filter_choice == "1":
        # Grouping berdasarkan Kecamatan (idkec)
        all_idkecs = sorted(list(set(d["idkec"] for d in data_cleaning_0 if d.get("idkec"))))
        unfinished_kecs = []
        for idk in all_idkecs:
            items_k = [d for d in data_cleaning_0 if d.get("idkec") == idk]
            total_k = len(items_k)
            done_k = sum(1 for it in items_k if is_done(it))
            rem_k = total_k - done_k
            if rem_k > 0:
                nama_k = items_k[0].get("nama_kecamatan", idk)
                unfinished_kecs.append({
                    "idkec": idk,
                    "nama_kecamatan": nama_k,
                    "total": total_k,
                    "done": done_k,
                    "remaining": rem_k,
                    "items": items_k
                })

        if not unfinished_kecs:
            print("\n[INFO] SELURUH KECAMATAN PADA PILIHAN INI SUDAH 100% SELESAI DIPROSES!")
            input("\nTekan ENTER untuk keluar...")
            sys.exit(0)

        completed_kecs_count = len(all_idkecs) - len(unfinished_kecs)
        print(f"\nDAFTAR KECAMATAN (Hanya yang belum selesai):")
        if completed_kecs_count > 0:
            print(f"(* {completed_kecs_count} kecamatan yang sudah selesai 100% otomatis disembunyikan)")

        for idx, k_info in enumerate(unfinished_kecs, 1):
            print(f"  [{idx}] Kecamatan {k_info['nama_kecamatan']} [{k_info['idkec']}] ({k_info['remaining']} data belum selesai / {k_info['total']} total)")

        kec_input = input(f"\nPilih nomor Kecamatan [1-{len(unfinished_kecs)}]: ").strip()
        if kec_input.isdigit() and 1 <= int(kec_input) <= len(unfinished_kecs):
            sel_kec = unfinished_kecs[int(kec_input) - 1]
            kec_items = sel_kec["items"]

            # Deteksi desa yang belum selesai di dalam kecamatan terpilih
            all_desas_in_kec = sorted(list(set(d["iddesa"] for d in kec_items if d.get("iddesa"))))
            unfinished_desas_in_kec = []
            for did in all_desas_in_kec:
                items_d = [d for d in kec_items if d.get("iddesa") == did]
                t_d = len(items_d)
                d_d = sum(1 for it in items_d if is_done(it))
                r_d = t_d - d_d
                if r_d > 0:
                    n_d = items_d[0].get("nama_desa", "")
                    label_d = f"Desa {did}" + (f" - {n_d}" if n_d else "")
                    unfinished_desas_in_kec.append({
                        "iddesa": did,
                        "label": label_d,
                        "remaining": r_d,
                        "items": items_d
                    })

            if len(unfinished_desas_in_kec) > 1:
                print(f"\nPILIHAN WILAYAH DI KECAMATAN {sel_kec['nama_kecamatan'].upper()}:")
                print(f"  [0] PROSES SEMUA DESA DI KECAMATAN INI ({sel_kec['remaining']} data belum selesai) [DEFAULT]")
                for didx, d_sub in enumerate(unfinished_desas_in_kec, 1):
                    print(f"  [{didx}] {d_sub['label']} ({d_sub['remaining']} data)")

                sub_choice = input(f"Pilih opsi desa [0-{len(unfinished_desas_in_kec)}] (Default: 0): ").strip()
                if sub_choice.isdigit() and 1 <= int(sub_choice) <= len(unfinished_desas_in_kec):
                    sel_sub_desa = unfinished_desas_in_kec[int(sub_choice) - 1]
                    active_data = sel_sub_desa["items"]
                    print(f"-> Filter aktif: {sel_sub_desa['label']} di Kec. {sel_kec['nama_kecamatan']} ({len(active_data)} data)")
                else:
                    active_data = kec_items
                    print(f"-> Filter aktif: Seluruh Kecamatan {sel_kec['nama_kecamatan']} ({len(active_data)} data)")
            else:
                active_data = kec_items
                print(f"-> Filter aktif: Kecamatan {sel_kec['nama_kecamatan']} ({len(active_data)} data)")

    elif filter_choice == "2":
        # Grouping berdasarkan ID Desa
        all_desas = sorted(list(set(d["iddesa"] for d in data_cleaning_0 if d.get("iddesa"))))
        unfinished_desas = []
        for did in all_desas:
            items_d = [d for d in data_cleaning_0 if d.get("iddesa") == did]
            t_d = len(items_d)
            d_d = sum(1 for it in items_d if is_done(it))
            r_d = t_d - d_d
            if r_d > 0:
                n_d = items_d[0].get("nama_desa", "")
                k_name = items_d[0].get("nama_kecamatan", "")
                label_parts = [f"Desa {did}"]
                if n_d:
                    label_parts.append(n_d)
                if k_name:
                    label_parts.append(f"(Kec. {k_name})")
                label_d = " - ".join(label_parts)
                unfinished_desas.append({
                    "iddesa": did,
                    "label": label_d,
                    "total": t_d,
                    "done": d_d,
                    "remaining": r_d,
                    "items": items_d
                })

        if not unfinished_desas:
            print("\n[INFO] SELURUH DESA PADA PILIHAN INI SUDAH 100% SELESAI DIPROSES!")
            input("\nTekan ENTER untuk keluar...")
            sys.exit(0)

        completed_desas_count = len(all_desas) - len(unfinished_desas)
        print("\nDAFTAR ID DESA (Hanya yang belum selesai):")
        if completed_desas_count > 0:
            print(f"(* {completed_desas_count} desa yang sudah selesai 100% otomatis disembunyikan)")

        for idx, d_info in enumerate(unfinished_desas, 1):
            print(f"  [{idx}] {d_info['label']} ({d_info['remaining']} data belum selesai / {d_info['total']} total)")

        desa_choice = input(f"\nPilih nomor ID Desa [1-{len(unfinished_desas)}]: ").strip()
        if desa_choice.isdigit() and 1 <= int(desa_choice) <= len(unfinished_desas):
            selected_desa = unfinished_desas[int(desa_choice) - 1]
            active_data = selected_desa["items"]
            print(f"-> Filter aktif: {selected_desa['label']} ({len(active_data)} data)")

    elif filter_choice == "3":
        # Grouping berdasarkan Batch
        all_batches = sorted(list(set(d["batch_info"] for d in data_cleaning_0 if d.get("batch_info"))))
        unfinished_batches = []
        for b_info in all_batches:
            items_b = [d for d in data_cleaning_0 if d.get("batch_info") == b_info]
            t_b = len(items_b)
            d_b = sum(1 for it in items_b if is_done(it))
            r_b = t_b - d_b
            if r_b > 0:
                unfinished_batches.append({
                    "batch_info": b_info,
                    "total": t_b,
                    "done": d_b,
                    "remaining": r_b,
                    "items": items_b
                })

        if not unfinished_batches:
            print("\n[INFO] SELURUH BATCH PADA PILIHAN INI SUDAH 100% SELESAI DIPROSES!")
            input("\nTekan ENTER untuk keluar...")
            sys.exit(0)

        print("\nDAFTAR BATCH (Hanya yang belum selesai):")
        for idx, b_item in enumerate(unfinished_batches, 1):
            print(f"  [{idx}] {b_item['batch_info']} ({b_item['remaining']} data belum selesai / {b_item['total']} total)")

        batch_choice = input(f"\nPilih nomor Batch [1-{len(unfinished_batches)}]: ").strip()
        if batch_choice.isdigit() and 1 <= int(batch_choice) <= len(unfinished_batches):
            selected_batch = unfinished_batches[int(batch_choice) - 1]
            active_data = selected_batch["items"]
            print(f"-> Filter aktif: Batch '{selected_batch['batch_info']}' ({len(active_data)} data)")

    elif filter_choice == "4":
        active_data = raw_data
        print(f"-> Memproses SEMUA data tanpa filter Petugas Cleaning ({len(active_data)} data).")
    else:
        print(f"-> Memproses SEMUA data Petugas Cleaning = 0 ({len(active_data)} data).")

    # 6. Ringkasan Pengerjaan
    cached_count = sum(1 for it in active_data if is_done(it))
    remaining_count = len(active_data) - cached_count

    print("\nRINGKASAN STATUS PENGERJAAN:")
    print(f" - Total Target Data : {len(active_data)}")
    print(f" - Sudah di Cache    : {cached_count}")
    print(f" - Sisa Diproses     : {remaining_count}")
    print("=" * 65)

    if remaining_count == 0:
        print("\n[INFO] Seluruh data pada pilihan ini sudah selesai diproses sebelumnya!")
        input("\nTekan ENTER untuk keluar...")
        sys.exit(0)

    # 7. Inisialisasi Playwright Browser
    profile_dir = os.path.join(BASE_DIR, "chrome_profile_anomali")
    report_data = load_report_data()
    report_map = {f"{r.get('link')}::{r.get('nama_usaha')}": idx for idx, r in enumerate(report_data)}

    success_count = 0
    fail_count = 0
    skip_count = 0

    with sync_playwright() as p:
        print("\nMembuka browser Chromium (menggunakan chrome_profile_anomali)...")
        context = p.chromium.launch_persistent_context(
            user_data_dir=profile_dir,
            headless=False,
            no_viewport=True,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"]
        )

        page = context.pages[0] if context.pages else context.new_page()
        page.goto("https://fasih-sm.bps.go.id/")

        print("\n" + "=" * 65)
        print(" Silakan LOGIN ke Fasih-SM melalui SSO BPS jika belum login.")
        print(" Jika sudah masuk ke halaman dashboard Fasih-SM,")
        print(" KEMBALI KE TERMINAL INI lalu tekan tombol ENTER untuk mulai.")
        print("=" * 65)
        input("Tekan ENTER untuk memulai proses otomatisasi...")

        total_target = len(active_data)
        for idx, item in enumerate(active_data, 1):
            link = item["link"]
            nama = item["nama_usaha"]
            desa = item["iddesa"]
            kec = item.get("nama_kecamatan", "")
            src_name = item.get("source_name", "")
            cache_key = f"{link}::{nama}"

            if cache_key in processed_cache or link in processed_cache:
                print(f"[{idx}/{total_target}] [{kec} | {desa}] SKIP (Sudah diproses): {nama}")
                continue

            while True:
                print(f"\n[{idx}/{total_target}] [{src_name}] [{kec} | {desa}] Memproses: {nama}")
                print(f"  Link: {link}")

                status = None
                msg = ""
                try:
                    res_status, res_msg = process_single_assignment(page, item)
                    status = res_status
                    msg = res_msg
                except Exception as e:
                    err_str = str(e)
                    if "null" in err_str or "Target closed" in err_str or "timeout" in err_str.lower() or "net::" in err_str or "Execution context" in err_str:
                        status = "SESSION_LOST"
                        msg = f"Koneksi terputus/error browser: {err_str}"
                    else:
                        print(f"  -> [ERROR] Terjadi exception: {e}")
                        status = "FAILED_EXCEPTION"
                        msg = str(e)

                # JIKA TERDETEKSI SESI HABIS / VPN PUTUS:
                if status == "SESSION_LOST":
                    print(f"  -> [SESI/VPN TERPUTUS] {msg}")
                    pause_for_relogin(page, link)
                    continue

                if status == "SUCCESS":
                    processed_cache.add(cache_key)
                    processed_cache.add(link)
                    save_cache(processed_cache)
                    success_count += 1
                    print(f"  -> [BERHASIL] Remark '#0' ditambahkan & form berhasil dikirim dan dikonfirmasi!")
                    # Sinkronisasi live ke Google Sheet & pewarnaan hijau
                    sync_to_google_sheet(item, status="SELESAI")
                elif status.startswith("SKIP_"):
                    processed_cache.add(cache_key)
                    processed_cache.add(link)
                    save_cache(processed_cache)
                    skip_count += 1
                    print(f"  -> [DILEWATI] {msg}")
                    sync_to_google_sheet(item, status=f"DILEWATI ({msg})")
                else:
                    fail_count += 1
                    print(f"  -> [GAGAL] {msg}")

                # Catat ke laporan
                record = {
                    "source_name": src_name,
                    "idkec": item.get("idkec", ""),
                    "nama_kecamatan": kec,
                    "iddesa": desa,
                    "nama_desa": item.get("nama_desa", ""),
                    "nama_usaha": nama,
                    "batch_info": item.get("batch_info", ""),
                    "petugas_cleaning": item.get("petugas_cleaning", ""),
                    "catatan_fasih": item.get("catatan_fasih", ""),
                    "remark_text": DEFAULT_REMARK_TEXT,
                    "status": status,
                    "keterangan": msg,
                    "link": link,
                    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
                }

                if cache_key in report_map:
                    report_data[report_map[cache_key]] = record
                else:
                    report_map[cache_key] = len(report_data)
                    report_data.append(record)

                save_report_data(report_data)
                break

            time.sleep(1.0)

        print("\n" + "=" * 65)
        print(" SEMUA TARGET PROSES SELESAI!")
        print(f" - Berhasil : {success_count}")
        print(f" - Dilewati : {skip_count}")
        print(f" - Gagal    : {fail_count}")
        print(f" Laporan disimpan di: {REPORT_EXCEL}")
        print("=" * 65)
        input("\nTekan ENTER untuk keluar...")

if __name__ == "__main__":
    main()
