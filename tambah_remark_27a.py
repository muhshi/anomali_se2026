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

DEFAULT_REMARK_TEXT = "#0"

SHEET_SOURCES = [
    {
        "id": "1",
        "name": "Dtata_NTB (Sheet NTB)",
        "url": "https://docs.google.com/spreadsheets/d/1ra-HZlqn1md8AIYaZz5VmOg-5Dxj83LYOHrmChvyz0s/export?format=csv&gid=293925488",
        "backup_filename": "dtata_ntb_backup.csv"
    },
    {
        "id": "2",
        "name": "Data detail anomali (Sheet 1)",
        "url": "https://docs.google.com/spreadsheets/d/12FWSyYak-scXS5skHFUSQbVkjy6Oh8ofejusqnOb5nA/export?format=csv&gid=0",
        "backup_filename": "data_detail_anomali_backup.csv"
    }
]

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
            "No", "Sumber Sheet", "ID Desa", "Nama Usaha", "Batch", "Petugas Cleaning",
            "Catatan Sheet", "Teks Remark", "Status Pengerjaan", "Keterangan", "Link Fasih", "Waktu Selesai"
        ]
        ws.append(headers)

        for idx, row in enumerate(report_list, 1):
            ws.append([
                idx,
                row.get("source_name", ""),
                row.get("iddesa", ""),
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

        data_items.append({
            "source_name": name,
            "link": clean_link,
            "raw_link": raw_link,
            "petugas_cleaning": petugas_cleaning,
            "batch_info": batch_info,
            "iddesa": iddesa,
            "nama_usaha": nama_usaha,
            "catatan_fasih": catatan_fasih,
            "row_num": row_num
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
# NAVIGASI BLOK II & PILIH USAHA
# ==============================================================================

def open_menu_blok_ii(page):
    """
    Membuka menu 'SE2026 - L BLOK II' (KETERANGAN USAHA/PERUSAHAAN) di sidebar.
    """
    time.sleep(1.0)
    for attempt in range(4):
        # 1. Coba klik langsung tab / tombol dengan Playwright locator
        try:
            tab_locator = page.get_by_role("tab", name=re.compile(r"(?:SE2026\s*[-–—]?\s*L.*BLOK\s*II|BLOK\s*II.*KETERANGAN\s*USAHA|BLOK\s*II)", re.IGNORECASE)).first
            if tab_locator.is_visible():
                tab_locator.scroll_into_view_if_needed(timeout=2000)
                tab_locator.click(timeout=3000, force=True)
                print("     [OK] Menu 'SE2026 - L BLOK II' diklik via Playwright locator.")
                time.sleep(1.5)
                return True
        except Exception:
            pass

        # 2. Coba evaluasi JavaScript DOM
        res = page.evaluate('''() => {
            function getCleanText(el) {
                return (el.innerText || el.textContent || '').trim();
            }

            const all = Array.from(document.querySelectorAll('*'));
            
            // Ambil semua sidebar menu items
            const sidebars = Array.from(document.querySelectorAll('aside, nav, [role="tablist"], [role="tree"], [class*="sidebar"], [class*="menu"], ul'));
            let allMenuTexts = [];
            sidebars.forEach(s => {
                const children = Array.from(s.querySelectorAll('a, button, [role="tab"], [role="treeitem"], [role="menuitem"], li, div[class*="item"]'));
                children.forEach(c => {
                    const t = getCleanText(c);
                    if (t && t.length < 70 && !allMenuTexts.includes(t)) allMenuTexts.push(t);
                });
            });

            // Cari leaf dengan kecocokan spesifik 'SE2026 - L' + 'BLOK II' atau 'BLOK II'
            const regFull = /(?:SE2026\s*[-–—]?\s*L.*BLOK\s*II|BLOK\s*II.*KETERANGAN\s*USAHA)/i;
            const regBlokII = /BLOK\s*II\b/i;
            const regKetUsaha = /KETERANGAN\s*USAHA/i;

            let matchedLeafs = all.filter(el => {
                const t = getCleanText(el);
                return (regFull.test(t) || regBlokII.test(t) || (regKetUsaha.test(t) && !t.toLowerCase().includes('anomali'))) &&
                       !Array.from(el.children).some(c => regFull.test(getCleanText(c)) || regBlokII.test(getCleanText(c)));
            });

            // Jika belum ada, coba ekspansi 'SE2026 - L' jika berupa accordion yang tertutup
            if (matchedLeafs.length === 0) {
                const lLeafs = all.filter(el => {
                    const t = getCleanText(el);
                    return /SE2026\s*[-–—]?\s*L/i.test(t) &&
                           !Array.from(el.children).some(c => /SE2026\s*[-–—]?\s*L/i.test(getCleanText(c)));
                });
                if (lLeafs.length > 0) {
                    const lTarget = lLeafs[0].closest('a, button, [role="tab"], [role="treeitem"], li') || lLeafs[0];
                    lTarget.scrollIntoView({ behavior: 'instant', block: 'center' });
                    lTarget.click();
                    return { success: false, need_wait: true, action: 'expanded_se2026_l', allMenus: allMenuTexts.slice(0, 20) };
                }
            }

            if (matchedLeafs.length > 0) {
                matchedLeafs.sort((a, b) => getCleanText(a).length - getCleanText(b).length);
                const target = matchedLeafs[0];
                const clickable = target.closest('a, button, [role="tab"], [role="treeitem"], [role="menuitem"], li, [tabindex]') || target;
                clickable.scrollIntoView({ behavior: 'instant', block: 'center' });
                clickable.click();
                return {
                    success: true,
                    menu: getCleanText(clickable),
                    allMenus: allMenuTexts.slice(0, 20)
                };
            }

            return { success: false, allMenus: allMenuTexts.slice(0, 20) };
        }''')

        if res and res.get("success"):
            print(f"     [OK] Berhasil membuka menu: '{res.get('menu')}'")
            time.sleep(1.5)
            return True
        elif res and res.get("need_wait"):
            print("     -> Mengklik 'SE2026 - L' untuk menampilkan sub-menu...")
            time.sleep(1.5)
            continue
        else:
            time.sleep(1.0)

    return False

def select_or_open_usaha(page, target_nama_usaha):
    """
    Mencari dan mengklik submenu nama usaha di sidebar atau tombol Lihat Rincian kartu usaha
    yang cocok dengan nama usaha di sheet.
    """
    time.sleep(1.0)
    clean_target = (target_nama_usaha or "").strip()
    print(f"  -> Mencari dan memilih usaha: '{clean_target}'...")

    # 1. Coba klik submenu nama usaha di sidebar kiri (di bawah SE2026 - L BLOK II)
    if clean_target:
        try:
            res_sub = page.evaluate('''(namaUsaha) => {
                function getCleanText(el) {
                    return (el.innerText || el.textContent || '').trim();
                }

                function cleanStr(s) {
                    return s.toLowerCase().replace(/[^a-z0-9]/g, '');
                }

                const uClean = cleanStr(namaUsaha);
                const sidebarEls = Array.from(document.querySelectorAll('aside, nav, [role="tablist"], [role="tree"], [class*="sidebar"], [class*="menu"], ul'));
                
                for (const sb of sidebarEls) {
                    const items = Array.from(sb.querySelectorAll('a, button, [role="tab"], [role="treeitem"], [role="menuitem"], li, div[class*="item"]'));
                    for (const it of items) {
                        const itText = getCleanText(it);
                        const itClean = cleanStr(itText);
                        if (itClean && (itClean === uClean || itClean.includes(uClean) || (uClean.length > 5 && uClean.includes(itClean)))) {
                            it.scrollIntoView({ behavior: 'instant', block: 'center' });
                            it.click();
                            return { success: true, method: 'sidebar_submenu_click', text: itText };
                        }
                    }
                }
                return { success: false };
            }''', clean_target)

            if res_sub and res_sub.get("success"):
                print(f"     [OK] Submenu usaha di sidebar berhasil diklik: '{res_sub.get('text')}'")
                time.sleep(2.0)
                return True
        except Exception:
            pass

    # 2. Coba klik tombol 'Lihat Rincian' pada baris / kartu usaha yang cocok
    if clean_target:
        try:
            res_card = page.evaluate('''(namaUsaha) => {
                function getCleanText(el) {
                    return (el.innerText || el.textContent || '').trim();
                }
                function cleanStr(s) {
                    return s.toLowerCase().replace(/[^a-z0-9]/g, '');
                }

                const uClean = cleanStr(namaUsaha);
                const containers = Array.from(document.querySelectorAll('tr, [role="row"], .card, div[class*="border"], div[class*="rounded"]'));
                
                for (const c of containers) {
                    const cText = getCleanText(c);
                    const cClean = cleanStr(cText);
                    if (cClean.includes(uClean)) {
                        const btns = Array.from(c.querySelectorAll('button, a, [role="button"]'));
                        const targetBtn = btns.find(b => {
                            const t = getCleanText(b).toLowerCase();
                            return (t === 'lihat rincian' || t.includes('lihat rincian') || t === 'lihat' || t === 'buka') && t.length < 30;
                        });
                        if (targetBtn) {
                            targetBtn.scrollIntoView({ behavior: 'instant', block: 'center' });
                            targetBtn.click();
                            return { success: true, method: 'card_button_match', text: getCleanText(targetBtn) };
                        }
                    }
                }
                return { success: false };
            }''', clean_target)

            if res_card and res_card.get("success"):
                print(f"     [OK] Tombol kartu usaha '{clean_target}' berhasil diklik: '{res_card.get('text')}'")
                time.sleep(2.0)
                return True
        except Exception:
            pass

    # 3. Fallback: jika hanya ada 1 usaha atau tombol 'Lihat Rincian' umum di halaman
    try:
        global_btn = page.locator("button, a, [role='button']").filter(has_text=re.compile(r"Lihat\s*Rincian", re.IGNORECASE)).first
        if global_btn.count() > 0 and global_btn.is_visible():
            global_btn.scroll_into_view_if_needed(timeout=2000)
            global_btn.click(timeout=3000, force=True)
            print("     [OK] Tombol 'Lihat Rincian' berhasil diklik via fallback locator.")
            time.sleep(2.0)
            return True
    except Exception:
        pass

    print("     [Info] Form rincian usaha diasumsikan sudah terbuka langsung.")
    return True

# ==============================================================================
# TAMBAH REMARK PADA PERTANYAAN 27.a
# ==============================================================================

def add_remark_27a(page, remark_text=DEFAULT_REMARK_TEXT):
    """
    Menambahkan catatan / remark '#0' pada pertanyaan:
    '27. a. Nilai produksi/pendapatan/penjualan barang dan jasa':
    1. Mencari field / label pertanyaan 27.a
    2. Menemukan dan mengklik tombol Catatan (speech bubble) di sebelah field tersebut:
       <button title="Catatan" aria-label="Catatan">
    3. Mengisikan teks remark ('#0') ke dalam input / textarea popup.
    4. Mengklik tombol Simpan (icon paper plane atau title='Simpan').
    5. Menutup popup catatan.
    6. Mengklik Simpan di form rincian jika ada.
    """
    print(f"  -> Mencari pertanyaan '27. a' dan mengklik tombol Catatan...")
    time.sleep(1.0)

    # 1. Scroll bertahap ke bawah pada form agar pertanyaan 27.a termuat
    for scroll_step in range(1, 4):
        page.evaluate('''(step) => {
            window.scrollBy(0, 800 * step);
            const scrollables = Array.from(document.querySelectorAll('*')).filter(el => {
                const s = window.getComputedStyle(el);
                return (s.overflowY === 'auto' || s.overflowY === 'scroll') && el.scrollHeight > el.clientHeight;
            });
            scrollables.forEach(s => s.scrollTop += 800 * step);
        }''', scroll_step)
        time.sleep(0.4)

    # 2. Temukan dan klik tombol Catatan pertanyaan 27.a
    btn_opened = False

    # A. Evaluasi DOM JavaScript presisi untuk mencari tombol Catatan pada baris 27.a
    for attempt in range(3):
        try:
            res_click = page.evaluate('''() => {
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

                const all = Array.from(document.querySelectorAll('*'));
                
                // Cari label / teks pertanyaan 27.a
                const qRegex = /(?:27\s*[\.\-]?\s*a|nilai\s*produksi|pendapatan\s*\/\s*penjualan)/i;
                const matchedLabels = all.filter(el => {
                    const t = getCleanText(el);
                    return qRegex.test(t) && !Array.from(el.children).some(c => qRegex.test(getCleanText(c)));
                });

                if (matchedLabels.length > 0) {
                    // Urutkan yang memuat 27.a paling spesifik
                    matchedLabels.sort((a, b) => {
                        const aHas27 = /27\s*[\.\-]?\s*a/i.test(getCleanText(a)) ? 0 : 1;
                        const bHas27 = /27\s*[\.\-]?\s*a/i.test(getCleanText(b)) ? 0 : 1;
                        return aHas27 - bHas27;
                    });

                    const targetLabel = matchedLabels[0];
                    targetLabel.scrollIntoView({ behavior: 'instant', block: 'center' });

                    // Telusuri parent container terdekat
                    let p = targetLabel.parentElement;
                    for (let i = 0; i < 7; i++) {
                        if (!p) break;
                        const catBtn = p.querySelector('button[title="Catatan"], button[aria-label="Catatan"], [role="button"][title="Catatan"]');
                        if (catBtn) {
                            clickEl(catBtn);
                            return { success: true, method: 'closest_parent_catatan', qText: getCleanText(targetLabel) };
                        }
                        p = p.parentElement;
                    }
                }

                // Fallback: cari input angka atau field 27.a lalu cari tombol Catatan di barisnya
                const inputs = Array.from(document.querySelectorAll('input'));
                for (const inp of inputs) {
                    const row = inp.closest('div[class*="flex"], tr, div[class*="grid"], fieldset') || inp.parentElement?.parentElement;
                    if (row) {
                        const rText = getCleanText(row);
                        if (qRegex.test(rText)) {
                            const btn = row.querySelector('button[title="Catatan"], button[aria-label="Catatan"], [role="button"][title="Catatan"]');
                            if (btn) {
                                clickEl(btn);
                                return { success: true, method: 'input_row_catatan', qText: rText.slice(0, 30) };
                            }
                        }
                    }
                }

                return { success: false };
            }''')

            if res_click and res_click.get("success"):
                btn_opened = True
                print(f"     [OK] Tombol Catatan 27.a berhasil diklik ({res_click.get('method')}).")
                break
        except Exception:
            pass

        time.sleep(1.0)

    # B. Fallback Playwright Locator jika JS belum menemukan
    if not btn_opened:
        try:
            row_loc = page.locator("div, tr, fieldset, section").filter(
                has_text=re.compile(r"27\s*[\.\-]?\s*a|nilai\s*produksi|pendapatan.*penjualan", re.IGNORECASE)
            ).filter(
                has=page.locator("button[title='Catatan'], button[aria-label='Catatan'], [role='button'][title='Catatan']")
            ).first

            if row_loc.count() > 0:
                cat_btn = row_loc.locator("button[title='Catatan'], button[aria-label='Catatan'], [role='button'][title='Catatan']").first
                if cat_btn.count() > 0:
                    cat_btn.scroll_into_view_if_needed(timeout=2000)
                    cat_btn.click(timeout=3000, force=True)
                    btn_opened = True
                    print("     [OK] Tombol Catatan diklik via Playwright row locator.")
        except Exception:
            pass

    if not btn_opened:
        print("     [GAGAL] Tombol Catatan pertanyaan 27.a tidak dapat ditemukan.")
        return False, "Tombol Catatan 27.a tidak ditemukan"

    time.sleep(1.2)

    # 3. Isi teks remark '#0' ke dalam input / textarea popup
    print(f"  -> Mengisikan teks remark '{remark_text}'...")
    text_filled = False

    try:
        textarea_loc = page.locator(
            "div[role='dialog'], [data-radix-popper-content-wrapper], div[class*='popover'], div[class*='dialog'], body"
        ).locator("textarea, input[type='text']").last

        if textarea_loc.count() > 0 and textarea_loc.is_visible():
            textarea_loc.fill(remark_text)
            text_filled = True
    except Exception:
        pass

    if not text_filled:
        try:
            res_fill = page.evaluate('''(txt) => {
                const inputs = Array.from(document.querySelectorAll('textarea, input[type="text"]')).filter(el => {
                    const style = window.getComputedStyle(el);
                    return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
                });
                if (inputs.length > 0) {
                    const targetInput = inputs[inputs.length - 1];
                    targetInput.value = txt;
                    targetInput.dispatchEvent(new Event('input', { bubbles: true }));
                    targetInput.dispatchEvent(new Event('change', { bubbles: true }));
                    return true;
                }
                return false;
            }''', remark_text)
            if res_fill:
                text_filled = True
        except Exception:
            pass

    if text_filled:
        print(f"     [OK] Teks remark '{remark_text}' berhasil diisikan.")
    else:
        print("     [Warning] Field input/textarea catatan tidak ditemukan.")

    time.sleep(0.8)

    # 4. Klik tombol Simpan catatan (icon paper plane atau title='Simpan')
    print("  -> Menyimpan catatan...")
    save_clicked = False
    try:
        save_btn = page.locator("button[title='Simpan'], [role='button'][title='Simpan']").first
        if save_btn.count() > 0 and save_btn.is_visible():
            save_btn.scroll_into_view_if_needed(timeout=2000)
            save_btn.click(timeout=3000, force=True)
            save_clicked = True
            print("     [OK] Tombol Simpan catatan diklik via Playwright.")
    except Exception:
        pass

    if not save_clicked:
        try:
            res_save = page.evaluate('''() => {
                function clickEl(el) {
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

                const btn = document.querySelector('button[title="Simpan"], [role="button"][title="Simpan"]');
                if (btn) {
                    clickEl(btn);
                    return true;
                }
                const svgBtn = Array.from(document.querySelectorAll('button')).find(b => {
                    const path = b.querySelector('svg path[d*="M10 14l11 -11"], svg path[d*="M21 3l-6.5 18"]');
                    return path !== null;
                });
                if (svgBtn) {
                    clickEl(svgBtn);
                    return true;
                }
                return false;
            }''')
            if res_save:
                save_clicked = True
                print("     [OK] Tombol Simpan catatan diklik via JS DOM.")
        except Exception:
            pass

    time.sleep(1.2)

    # 5. Tutup dialog popover catatan
    try:
        close_btn = page.locator("div[role='dialog'], [data-radix-popper-content-wrapper], div[class*='popover']").locator("button[title='Tutup'], button[aria-label='Close'], button:has-text('Tutup')").first
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

    # 6. Jika ada tombol Simpan di dalam modal rincian, klik untuk simpan rincian
    try:
        page.evaluate('''() => {
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
                return true;
            }
            return false;
        }''')
        time.sleep(1.5)
    except Exception:
        pass

    return True, "Remark 27.a berhasil ditambahkan"

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
    except Exception:
        pass

    # 4. Tunggu pengiriman data ke server Fasih
    time.sleep(2.0)

    # Cek toast error 'failed to fetch'
    has_fetch_error = False
    try:
        has_fetch_error = page.evaluate('''() => {
            const t = ((document.body.innerText || '') + ' ' + (document.body.textContent || '')).toLowerCase();
            return t.includes('failed to fetch');
        }''')
    except Exception:
        pass

    if has_fetch_error:
        return False, "Toast 'Failed to fetch' muncul setelah kirim"

    return True, "Berhasil dikirim dan dikonfirmasi"

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
        page.goto(link, wait_until="networkidle", timeout=30000)
    except Exception as e:
        if "interrupted by another navigation" in str(e):
            time.sleep(2)
            try:
                page.goto(link, wait_until="networkidle", timeout=30000)
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

    time.sleep(1.0)

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
    print("  [1] Dtata_NTB (https://docs.google.com/...gid=293925488) [DEFAULT]")
    print("  [2] Data detail anomali (https://docs.google.com/...gid=0)")
    print("  [3] GABUNGKAN SEMUA SHEET (Dtata_NTB + Data detail anomali)")

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

    # 2. Filter Petugas Cleaning = 0 (sesuai instruksi user)
    data_cleaning_0 = [d for d in raw_data if d["petugas_cleaning"] == "0"]
    print(f"Data dengan 'Petugas Cleaning = 0': {len(data_cleaning_0)} baris.")

    # 3. Pilihan Menu Filter Interaktif di Terminal
    print("\nPILIHAN FILTER DATA:")
    print("  [0] SEMUA DATA (Petugas Cleaning = 0) [DEFAULT]")
    print("  [1] Filter Berdasarkan ID Desa (iddesa)")
    print("  [2] Filter Berdasarkan Batch (batch_info)")
    print("  [3] Tampilkan Semua Petugas Cleaning (Termasuk non-0)")

    filter_choice = input("\nPilih nomor filter [0/1/2/3] (Default: 0): ").strip()
    active_data = data_cleaning_0

    if filter_choice == "1":
        all_desas = sorted(list(set(d["iddesa"] for d in data_cleaning_0 if d["iddesa"])))
        print("\nDAFTAR ID DESA:")
        for idx, d_id in enumerate(all_desas, 1):
            count_d = sum(1 for it in data_cleaning_0 if it["iddesa"] == d_id)
            print(f"  [{idx}] Desa {d_id} ({count_d} data)")
        desa_choice = input(f"Pilih nomor ID Desa [1-{len(all_desas)}]: ").strip()
        if desa_choice.isdigit() and 1 <= int(desa_choice) <= len(all_desas):
            selected_desa = all_desas[int(desa_choice) - 1]
            active_data = [d for d in data_cleaning_0 if d["iddesa"] == selected_desa]
            print(f"-> Filter aktif: ID Desa {selected_desa} ({len(active_data)} data)")
    elif filter_choice == "2":
        all_batches = sorted(list(set(d["batch_info"] for d in data_cleaning_0 if d["batch_info"])))
        print("\nDAFTAR BATCH:")
        for idx, b_info in enumerate(all_batches, 1):
            count_b = sum(1 for it in data_cleaning_0 if it["batch_info"] == b_info)
            print(f"  [{idx}] {b_info} ({count_b} data)")
        batch_choice = input(f"Pilih nomor Batch [1-{len(all_batches)}]: ").strip()
        if batch_choice.isdigit() and 1 <= int(batch_choice) <= len(all_batches):
            selected_batch = all_batches[int(batch_choice) - 1]
            active_data = [d for d in data_cleaning_0 if d["batch_info"] == selected_batch]
            print(f"-> Filter aktif: Batch '{selected_batch}' ({len(active_data)} data)")
    elif filter_choice == "3":
        active_data = raw_data
        print(f"-> Memproses SEMUA data tanpa filter Petugas Cleaning ({len(active_data)} data).")
    else:
        print(f"-> Memproses SEMUA data Petugas Cleaning = 0 ({len(active_data)} data).")

    # 4. Cek Cache Progres
    processed_cache = load_cache()
    cached_count = sum(1 for it in active_data if f"{it['link']}::{it['nama_usaha']}" in processed_cache or it["link"] in processed_cache)
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

    # 5. Inisialisasi Playwright Browser
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
            src_name = item.get("source_name", "")
            cache_key = f"{link}::{nama}"

            if cache_key in processed_cache or link in processed_cache:
                print(f"[{idx}/{total_target}] [{desa}] SKIP (Sudah diproses): {nama}")
                continue

            while True:
                print(f"\n[{idx}/{total_target}] [{src_name}] [{desa}] Memproses: {nama}")
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
                elif status.startswith("SKIP_"):
                    processed_cache.add(cache_key)
                    processed_cache.add(link)
                    save_cache(processed_cache)
                    skip_count += 1
                    print(f"  -> [DILEWATI] {msg}")
                else:
                    fail_count += 1
                    print(f"  -> [GAGAL] {msg}")

                # Catat ke laporan
                record = {
                    "source_name": src_name,
                    "iddesa": desa,
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
