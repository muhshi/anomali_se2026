#!/usr/bin/env python3
"""
Bot Otomatisasi Perubahan '8. d. Jenis kawasan beroperasi' ke '10. Diluar kawasan'
Fasih-SM BPS (SE2026)

Fungsi:
1. Membaca file data Excel anomali (khususnya sqllab_untitled_query_17_20261019T113729.xlsx atau file kawasan lainnya).
2. Membuka link assignment /edit di Fasih-SM.
3. Berpindah ke menu 'SE2026 - P'.
4. Mengubah isian pertanyaan '8. d. Jenis kawasan beroperasi' dari:
   '1. Kawasan Ekonomi Khusus (KEK)' menjadi '10. Diluar kawasan' (atau '10. Di luar kawasan').
5. Membersihkan field 'Nama kawasan' jika masih aktif/terisi.
6. Menyimpan / Mengirimkan (Kirim + Konfirmasi modal / Submit Paksa jika diperlukan).
7. Mencatat cache progres ('processed_jenis_kawasan.json') dan laporan ('laporan_jenis_kawasan.xlsx').
"""

import os
import sys
import time
import re
import json
import random
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
CACHE_FILE = os.path.join(BASE_DIR, "processed_jenis_kawasan.json")
REPORT_EXCEL = os.path.join(BASE_DIR, "laporan_jenis_kawasan.xlsx")
REPORT_JSON = os.path.join(BASE_DIR, "laporan_jenis_kawasan.json")

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
        print(f"  -> [Warning] Gagal menyimpan cache: {e}")

def load_report_data():
    if os.path.exists(REPORT_JSON):
        try:
            with open(REPORT_JSON, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_report_data(reports):
    try:
        with open(REPORT_JSON, "w", encoding="utf-8") as f:
            json.dump(reports, f, indent=4)
    except Exception as e:
        print(f"  -> [Warning] Gagal menyimpan laporan JSON: {e}")

    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Laporan Ubah Kawasan"
        headers = [
            "No", "Waktu", "Kecamatan", "Desa", "Nama SLS", "Nama Usaha",
            "Assignment ID", "Status Awal", "Jenis Kawasan Awal", "Hasil",
            "Keterangan", "Detail Pesan", "Link Fasih"
        ]
        ws.append(headers)
        for idx, r in enumerate(reports, 1):
            ws.append([
                idx,
                r.get("waktu", ""),
                r.get("kecamatan", ""),
                r.get("desa", ""),
                r.get("nama_sls", ""),
                r.get("nama_usaha", ""),
                r.get("assignment_id", ""),
                r.get("status_awal", ""),
                r.get("jenis_kawasan_awal", ""),
                r.get("status_hasil", ""),
                r.get("keterangan", ""),
                r.get("detail_pesan", ""),
                r.get("link", "")
            ])
        wb.save(REPORT_EXCEL)
    except Exception as e:
        print(f"  -> [Warning] Gagal menyimpan laporan Excel: {e}")

# ==============================================================================
# PEMBACAAN FILE EXCEL
# ==============================================================================

def find_excel_files():
    """Mencari file Excel yang relevan dengan anomali kawasan di folder data."""
    matched = []
    if not os.path.exists(DATA_DIR):
        return matched

    # Telusuri semua file .xlsx di folder data dan subfoldernya
    for root, _, files in os.walk(DATA_DIR):
        for f in files:
            if not f.endswith(".xlsx") or f.startswith("~$") or "laporan" in f.lower():
                continue
            full_path = os.path.join(root, f)
            # Prioritaskan jika nama file memuat sqllab, kawasan, atau query
            matched.append(full_path)

    # Urutkan agar file query sqllab / kawasan berada di urutan teratas
    def sort_score(path):
        name = os.path.basename(path).lower()
        if "sqllab" in name and "kawasan" in name:
            return 0
        if "sqllab" in name:
            return 1
        if "kawasan" in name:
            return 2
        return 10

    matched.sort(key=sort_score)
    return matched

def load_data_from_excel(file_path):
    """Membaca data anomali jenis kawasan dari file Excel."""
    wb = openpyxl.load_workbook(file_path, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    header = [str(col).strip().lower() if col is not None else "" for col in rows[0]]
    
    # Deteksi index kolom
    def get_col_idx(candidates):
        for c in candidates:
            if c in header:
                return header.index(c)
        for i, h in enumerate(header):
            for c in candidates:
                if c in h:
                    return i
        return None

    link_idx = get_col_idx(["link_fasih", "link", "url", "link fasih"])
    kec_idx = get_col_idx(["kecamatan", "nama_kecamatan", "nama kecamatan"])
    desa_idx = get_col_idx(["desa", "kelurahan", "nama_desa"])
    sls_idx = get_col_idx(["nama_sls", "sls", "nama sls", "id_sub_sls"])
    usaha_idx = get_col_idx(["nama_usaha", "nama usaha", "nama_bangunan", "nama bangunan"])
    assign_idx = get_col_idx(["assignment_id", "assignment id", "id_assignment"])
    status_assign_idx = get_col_idx(["status_assignment", "status assignment"])
    kawasan_idx = get_col_idx(["jenis_kawasan", "jenis kawasan"])
    nama_kawasan_idx = get_col_idx(["nama_kawasan", "nama kawasan"])

    data_items = []
    for r_idx, row in enumerate(rows[1:], 2):
        if not row or all(v is None for v in row):
            continue

        raw_link = str(row[link_idx]).strip() if link_idx is not None and row[link_idx] else ""
        if not raw_link or "fasih" not in raw_link:
            continue

        # Format URL /edit
        clean_link = raw_link.split("?")[0].rstrip("/")
        if not clean_link.endswith("/edit"):
            clean_link = clean_link + "/edit"

        kec = str(row[kec_idx]).strip().upper() if kec_idx is not None and row[kec_idx] else "LAINNYA"
        desa = str(row[desa_idx]).strip() if desa_idx is not None and row[desa_idx] else ""
        sls = str(row[sls_idx]).strip() if sls_idx is not None and row[sls_idx] else ""
        usaha = str(row[usaha_idx]).strip() if usaha_idx is not None and row[usaha_idx] else "-"
        assign_id = str(row[assign_idx]).strip() if assign_idx is not None and row[assign_idx] else ""
        status_assign = str(row[status_assign_idx]).strip() if status_assign_idx is not None and row[status_assign_idx] else ""
        j_kawasan = str(row[kawasan_idx]).strip() if kawasan_idx is not None and row[kawasan_idx] else ""
        n_kawasan = str(row[nama_kawasan_idx]).strip() if nama_kawasan_idx is not None and row[nama_kawasan_idx] else ""

        data_items.append({
            "link": clean_link,
            "raw_link": raw_link,
            "kecamatan": kec,
            "desa": desa,
            "nama_sls": sls,
            "nama_usaha": usaha,
            "assignment_id": assign_id,
            "status_awal": status_assign,
            "jenis_kawasan_awal": j_kawasan,
            "nama_kawasan_awal": n_kawasan,
            "source_file": os.path.basename(file_path),
            "row_num": r_idx
        })

    return data_items

# ==============================================================================
# ANTI-BOT / WAF / SSO HANDLERS
# ==============================================================================

def check_is_bot_or_blocked(page):
    """Mengecek apakah halaman saat ini menunjukkan pesan bot, WAF, atau error SSO."""
    try:
        if page.locator("text=/mendeteksi koneksi anda sebagai bot/i").count() > 0:
            return True
        if page.locator("text=/HaloSIS/i").count() > 0:
            return True
        if page.locator("text=/Lanjutkan dengan SSO/i").count() > 0:
            return True
        if page.locator("text=/Access Denied/i").count() > 0:
            return True
        url = page.url.lower()
        if "sso" in url and ("error" in url or "login" in url or "block" in url):
            return True
    except Exception:
        pass
    return False

def resolve_bot_detection(page, target_link):
    """Prosedur pemulihan otomatis jika terdeteksi bot atau rate-limit."""
    if not check_is_bot_or_blocked(page):
        return True

    print("=" * 60)
    print(" [WAF / BOT DETECTED] Halaman terdeteksi bot atau terganggu sesi SSO.")
    print(" Memulai prosedur pemulihan otomatis...")
    
    max_retries = 2
    for attempt in range(1, max_retries + 1):
        print(f" -> [Percobaan {attempt}/{max_retries}] Menunggu 10 detik agar rate-limit ter-reset...")
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

        if not check_is_bot_or_blocked(page):
            print(" -> [BERHASIL] Halaman terbebas dari deteksi bot! Melanjutkan...")
            print("=" * 60)
            return True

    print("=" * 60)
    print(" MASIH TERDETEKSI BOT / HALOSIS!")
    print(" Silakan verifikasi di browser secara manual jika perlu.")
    print("=" * 60)
    time.sleep(5)
    return True

# ==============================================================================
# LOGIKA INTI PERUBAHAN FORM FASIH-SM
# ==============================================================================

def open_menu_blok_ii(page):
    """
    Membuka menu 'SE2026 - L BLOK II' (KETERANGAN USAHA/PERUSAHAAN).
    Pencarian:
    - 'SE2026 - L BLOK II'
    - 'BLOK II'
    - 'KETERANGAN USAHA/PERUSAHAAN' atau 'KETERANGAN USAHA'
    Jika 'SE2026 - L' adalah menu accordion yang tertutup, akan diklik terlebih dahulu untuk mengekspansi.
    """
    for attempt in range(2):
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

            // 1. Cari leaf dengan kecocokan spesifik 'SE2026 - L' + 'BLOK II' atau 'BLOK II'
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
            all_m = res.get("allMenus", []) if res else []
            print(f"     [Info] Daftar menu sidebar terdeteksi: {all_m}")
            time.sleep(1.0)

    return False

def click_lihat_rincian(page, nama_usaha=""):
    """
    Mengklik tombol oranye 'Lihat Rincian' pada menu SE2026 - L BLOK II.
    Mendukung:
    - Tombol 'Lihat Rincian' pada baris tabel/kartu sesuai nama usaha
    - Tombol 'Lihat Rincian' langsung pada halaman via Playwright & DOM
    - Fallback sub-menu nama usaha di sidebar kiri
    """
    time.sleep(1.0)
    clean_usaha = nama_usaha.strip()

    # 1. Coba klik langsung dengan Playwright locator berbasis card/container
    if clean_usaha:
        try:
            # Cari card yang memuat nama usaha lalu cari button 'Lihat Rincian'
            card_btn = page.locator("div, tr, li").filter(has_text=re.compile(re.escape(clean_usaha), re.IGNORECASE)).locator("button").filter(has_text=re.compile(r"Lihat\s*Rincian", re.IGNORECASE)).first
            if card_btn.count() > 0:
                card_btn.scroll_into_view_if_needed(timeout=2000)
                card_btn.click(timeout=3000, force=True)
                print(f"     [OK] Tombol oranye 'Lihat Rincian' pada usaha '{clean_usaha}' berhasil diklik via Playwright card locator.")
                time.sleep(2.5)
                return True
        except Exception:
            pass

    # 2. Coba klik langsung tombol dengan teks 'Lihat Rincian' di seluruh halaman
    try:
        global_btn = page.locator("button").filter(has_text=re.compile(r"Lihat\s*Rincian", re.IGNORECASE)).first
        if global_btn.count() > 0:
            global_btn.scroll_into_view_if_needed(timeout=2000)
            global_btn.click(timeout=3000, force=True)
            print("     [OK] Tombol oranye 'Lihat Rincian' berhasil diklik via Playwright global button locator.")
            time.sleep(2.5)
            return True
    except Exception:
        pass

    # 3. Evaluasi JavaScript DOM yang presisi (HANYA mencari tag button, a, [role='button'], BUKAN container div)
    for attempt in range(3):
        res = page.evaluate('''(namaUsaha) => {
            function getCleanText(el) {
                return (el.innerText || el.textContent || '').trim();
            }

            function clickElement(el) {
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

            // A. Cari tombol 'Lihat Rincian' di dalam kartu usaha yang cocok
            if (namaUsaha) {
                const uClean = namaUsaha.toLowerCase().trim();
                const containers = Array.from(document.querySelectorAll('tr, [role="row"], .card, div[class*="border"], div[class*="rounded"]'));
                for (const c of containers) {
                    const cText = getCleanText(c).toLowerCase();
                    if (cText.includes(uClean)) {
                        // HANYA query button, a, role=button (JANGAN span atau div agar tidak mengklik kartu luar)
                        const btns = Array.from(c.querySelectorAll('button, a, [role="button"]'));
                        const targetBtn = btns.find(b => {
                            const t = getCleanText(b).toLowerCase();
                            return (t === 'lihat rincian' || t.includes('lihat rincian') || t === 'lihat') && t.length < 30;
                        });
                        if (targetBtn) {
                            clickElement(targetBtn);
                            return { success: true, method: 'card_button_match', text: getCleanText(targetBtn) };
                        }
                    }
                }
            }

            // B. Cari SEMUA tombol di halaman yang bertuliskan 'Lihat Rincian'
            const allBtns = Array.from(document.querySelectorAll('button, a, [role="button"]'));
            const rincianBtn = allBtns.find(b => {
                const t = getCleanText(b).toLowerCase();
                return (t === 'lihat rincian' || t.includes('lihat rincian') || t.includes('lihat')) && t.length < 30;
            });
            if (rincianBtn) {
                clickElement(rincianBtn);
                return { success: true, method: 'global_button_match', text: getCleanText(rincianBtn) };
            }

            // C. Fallback: Cari nama usaha di sub-menu sidebar (di bawah SE2026 - L BLOK II)
            if (namaUsaha) {
                const uClean = namaUsaha.toLowerCase().trim();
                const sidebarEls = Array.from(document.querySelectorAll('aside, nav, [role="tablist"], [class*="sidebar"], [class*="menu"], ul'));
                for (const sb of sidebarEls) {
                    const items = Array.from(sb.querySelectorAll('*')).filter(el => {
                        const t = getCleanText(el).toLowerCase();
                        return (t === uClean || t.includes(uClean)) && !Array.from(el.children).some(c => getCleanText(c).toLowerCase().includes(uClean));
                    });
                    if (items.length > 0) {
                        const clickEl = items[0].closest('a, button, [role="tab"], [role="treeitem"], [role="menuitem"], li') || items[0];
                        clickElement(clickEl);
                        return { success: true, method: 'sidebar_submenu_nama_usaha', text: getCleanText(clickEl) };
                    }
                }
            }

            return { success: false };
        }''', clean_usaha)

        if res and res.get("success"):
            print(f"     [OK] Berhasil mengklik tombol oranye '{res.get('text')}' ({res.get('method')})")
            time.sleep(2.0)
            return True
        time.sleep(1.0)

    print("     [Warning] Tombol 'Lihat Rincian' tidak terdeteksi langsung, melanjutkan pencarian pertanyaan...")
    return False

def open_menu_usaha_bawah_p(page, item):
    """Fallback membuka menu di bawah 'SE2026 - P' jika BLOK II tidak ditemukan terpisah."""
    nama_usaha = item.get("nama_usaha", "").strip()

    # Dapatkan daftar semua menu di sidebar untuk diagnostik dan pencarian
    menu_info = page.evaluate('''(namaUsaha) => {
        function getCleanText(el) {
            return (el.innerText || el.textContent || '').trim();
        }

        const all = Array.from(document.querySelectorAll('*'));
        
        // 1. Ambil daftar semua teks item di sidebar / nav / aside
        const sidebarEls = Array.from(document.querySelectorAll('aside, nav, [role="tablist"], [role="tree"], [class*="sidebar"], [class*="menu"], ul'));
        let menuTexts = [];
        sidebarEls.forEach(s => {
            const children = Array.from(s.querySelectorAll('a, button, [role="tab"], [role="treeitem"], [role="menuitem"], li, div[class*="item"]'));
            children.forEach(c => {
                const t = getCleanText(c);
                if (t && t.length < 60 && !menuTexts.includes(t)) {
                    menuTexts.push(t);
                }
            });
        });

        // 2. Coba cocokkan dengan nama usaha langsung
        if (namaUsaha) {
            const uClean = namaUsaha.toLowerCase().trim();
            const matchedLeafs = all.filter(el => {
                const t = getCleanText(el).toLowerCase();
                return (t === uClean || t.includes(uClean)) &&
                       !Array.from(el.children).some(c => getCleanText(c).toLowerCase().includes(uClean));
            });
            if (matchedLeafs.length > 0) {
                const target = matchedLeafs[0];
                const clickable = target.closest('a, button, [role="tab"], [role="treeitem"], [role="menuitem"], li, [tabindex]') || target;
                clickable.scrollIntoView({ behavior: 'instant', block: 'center' });
                clickable.click();
                return {
                    success: true,
                    menu: getCleanText(clickable),
                    method: 'by_nama_usaha',
                    allMenus: menuTexts.slice(0, 20)
                };
            }
        }

        // 3. Cari elemen 'SE2026 - P'
        const pLeafs = all.filter(el => {
            const t = getCleanText(el);
            return /SE2026\s*[-–—]?\s*P/i.test(t) &&
                   !Array.from(el.children).some(c => /SE2026\s*[-–—]?\s*P/i.test(getCleanText(c)));
        });

        if (pLeafs.length > 0) {
            pLeafs.sort((a, b) => getCleanText(a).length - getCleanText(b).length);
            const pLeaf = pLeafs[0];
            const pClickable = pLeaf.closest('a, button, [role="tab"], [role="treeitem"], [role="menuitem"], li, div[class*="item"]') || pLeaf;

            // Cari container sidebar tempat P berada
            const sidebar = pLeaf.closest('aside, nav, [role="tablist"], [role="tree"], [class*="sidebar"], [class*="menu"], ul') || document.body;
            const sidebarItems = Array.from(sidebar.querySelectorAll('a, button, [role="tab"], [role="treeitem"], [role="menuitem"], li, div[tabindex]')).filter(el => {
                const t = getCleanText(el);
                return t.length > 0 && t.length < 70 && el.offsetParent !== null;
            });

            // Temukan index P dalam sidebarItems
            const pIndex = sidebarItems.findIndex(el => /SE2026\s*[-–—]?\s*P/i.test(getCleanText(el)));
            if (pIndex !== -1 && pIndex + 1 < sidebarItems.length) {
                const nextItem = sidebarItems[pIndex + 1];
                const targetClick = nextItem.querySelector('a, button, [role="tab"], [role="treeitem"], [role="menuitem"]') || nextItem;
                targetClick.scrollIntoView({ behavior: 'instant', block: 'center' });
                targetClick.click();
                return {
                    success: true,
                    menu: getCleanText(targetClick),
                    method: 'item_below_p',
                    allMenus: menuTexts.slice(0, 20)
                };
            }

            // Jika P adalah accordion yang tertutup, coba klik P sekali lalu cari lagi child/sibling berikutnya
            let sibling = pClickable.nextElementSibling || pClickable.parentElement?.nextElementSibling;
            if (sibling) {
                const targetClick = sibling.querySelector('a, button, [role="tab"], [role="treeitem"], [role="menuitem"]') || sibling;
                targetClick.scrollIntoView({ behavior: 'instant', block: 'center' });
                targetClick.click();
                return {
                    success: true,
                    menu: getCleanText(targetClick),
                    method: 'sibling_below_p',
                    allMenus: menuTexts.slice(0, 20)
                };
            }
        }

        // 4. Fallback: Cari menu dengan pola '1. ' atau 'Usaha' atau 'SE2026 - U'
        const candidateLeafs = all.filter(el => {
            const t = getCleanText(el);
            return (/^(?:1\.|usaha|se2026\s*[-–—]?\s*u)/i.test(t) || /rincian\s+usaha/i.test(t)) &&
                   !Array.from(el.children).some(c => (/^(?:1\.|usaha|se2026\s*[-–—]?\s*u)/i.test(getCleanText(c))));
        });
        if (candidateLeafs.length > 0) {
            candidateLeafs.sort((a, b) => getCleanText(a).length - getCleanText(b).length);
            const target = candidateLeafs[0];
            const clickable = target.closest('a, button, [role="tab"], [role="treeitem"], [role="menuitem"], li') || target;
            clickable.scrollIntoView({ behavior: 'instant', block: 'center' });
            clickable.click();
            return {
                success: true,
                menu: getCleanText(clickable),
                method: 'fallback_candidate',
                allMenus: menuTexts.slice(0, 20)
            };
        }

        return { success: false, allMenus: menuTexts.slice(0, 20) };
    }''', nama_usaha)

    if menu_info and menu_info.get("success"):
        menu_name = menu_info.get("menu", "")
        method = menu_info.get("method", "")
        print(f"     [OK] Berhasil mengklik menu bawah: '{menu_name}' ({method})")
        time.sleep(1.5)
        return True
    else:
        all_m = menu_info.get("allMenus", []) if menu_info else []
        print(f"     [Info] Menu sidebar terdeteksi: {all_m}")
        return False

def scroll_form_down(page):
    """
    Scroll ke bawah pada form / halaman Fasih-SM agar pertanyaan 8.d termuat dan tampil.
    Melakukan scroll bertahap pada window maupun elemen container scrollable.
    """
    for step in range(1, 5):
        page.evaluate('''(stepNum) => {
            // Scroll window bertahap
            window.scrollBy(0, 1000 * stepNum);
            
            // Scroll semua elemen container scrollable (div overflow, form, main)
            const scrollables = Array.from(document.querySelectorAll('*')).filter(el => {
                const style = window.getComputedStyle(el);
                return (style.overflowY === 'auto' || style.overflowY === 'scroll') && el.scrollHeight > el.clientHeight;
            });
            scrollables.forEach(s => {
                s.scrollTop = s.scrollHeight * (stepNum / 4);
            });
        }''', step)
        time.sleep(0.4)

    # Scroll penuh ke bawah sekali lagi
    page.evaluate('''() => {
        window.scrollTo({ top: document.body.scrollHeight, behavior: 'instant' });
        const scrollables = Array.from(document.querySelectorAll('*')).filter(el => {
            const style = window.getComputedStyle(el);
            return (style.overflowY === 'auto' || style.overflowY === 'scroll') && el.scrollHeight > el.clientHeight;
        });
        scrollables.forEach(s => {
            s.scrollTop = s.scrollHeight;
        });
    }''')
    time.sleep(0.6)

def change_jenis_kawasan_to_10(page):
    """
    Mengubah pilihan pertanyaan 8. d. Jenis kawasan beroperasi ke '10. Diluar kawasan'.
    Mendukung berbagai struktur UI:
    - Radio button (native / custom)
    - Custom Dropdown / Combobox (Radix UI / Shadcn)
    - Native <select>
    """
    time.sleep(1.0)

    # Evaluasi JS cerdas untuk menemukan dan memilih opsi 10 Diluar kawasan
    res = page.evaluate('''() => {
        function triggerEvents(el) {
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
        }

        function clickElement(el) {
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

        const allElements = Array.from(document.querySelectorAll('*'));

        // Regex untuk mencocokkan teks opsi '10. Diluar kawasan' atau '10. Di luar kawasan'
        const optRegex = /(?:^|\s)10\s*[\.\-]?\s*(?:di\s*luar|diluar)\s*kawasan|(?:di\s*luar|diluar)\s*kawasan/i;
        const qRegex = /(?:8\s*[\.\-]?\s*d|jenis\s+kawasan|kawasan.*beroperasi)/i;

        // 1. Cek apakah radio button dengan value "10" atau teks terkait sudah ada di form
        const radios = Array.from(document.querySelectorAll('input[type="radio"]'));
        
        // Cari radio dengan value 10
        let targetRadio = radios.find(r => {
            const v = (r.value || '').trim();
            return v === '10' || optRegex.test(v);
        });

        // Jika tidak langsung dari value, cari radio yang label/induknya mengandung teks opsi
        if (!targetRadio) {
            for (const r of radios) {
                const parent = r.closest('label, tr, fieldset, div[role="radiogroup"], div');
                if (parent) {
                    const txt = (parent.innerText || parent.textContent || '').trim();
                    if (optRegex.test(txt)) {
                        targetRadio = r;
                        break;
                    }
                }
            }
        }

        if (targetRadio) {
            const wasChecked = targetRadio.checked;
            clickElement(targetRadio);
            targetRadio.checked = true;
            triggerEvents(targetRadio);

            const parentLabel = targetRadio.closest('label');
            if (parentLabel) parentLabel.click();

            return {
                success: true,
                method: 'radio_input',
                already_selected: wasChecked,
                details: targetRadio.value || 'radio'
            };
        }

        // 2. Cek elemen leaf teks yang memuat '10. Diluar kawasan' / '10. Di luar kawasan'
        const matchedTextLeafs = allElements.filter(el => {
            const t = (el.innerText || el.textContent || '').trim();
            if (!optRegex.test(t)) return false;
            return !Array.from(el.children).some(c => optRegex.test((c.innerText || c.textContent || '').trim()));
        });

        if (matchedTextLeafs.length > 0) {
            matchedTextLeafs.sort((a, b) => (a.innerText || '').length - (b.innerText || '').length);
            const leaf = matchedTextLeafs[0];
            const clickable = leaf.closest('label, [role="radio"], button, tr, div[class*="cursor-pointer"], div') || leaf;
            
            clickElement(clickable);
            const innerRadio = clickable.querySelector('input[type="radio"]');
            if (innerRadio) {
                innerRadio.checked = true;
                triggerEvents(innerRadio);
            }
            return {
                success: true,
                method: 'text_leaf_click',
                already_selected: false,
                details: (leaf.innerText || leaf.textContent || '').trim().substring(0, 40)
            };
        }

        // 3. Cek apakah berupa custom Select / Combobox (Radix UI / Popover)
        // Cari elemen pemicu dropdown kawasan yang saat ini menampilkan '1. Kawasan Ekonomi Khusus (KEK)' atau berada di blok pertanyaan 8.d
        const kekLeafs = allElements.filter(el => {
            const t = (el.innerText || el.textContent || '').trim();
            return /(?:kawasan\s+ekonomi\s+khusus|kek)/i.test(t) &&
                   !Array.from(el.children).some(c => /(?:kawasan\s+ekonomi\s+khusus|kek)/i.test((c.innerText || c.textContent || '').trim()));
        });

        if (kekLeafs.length > 0) {
            const kekLeaf = kekLeafs[0];
            const dropdownTrigger = kekLeaf.closest('button, [role="combobox"], [role="button"], div[aria-haspopup="listbox"], div[tabindex]') || kekLeaf;
            clickElement(dropdownTrigger);
            return {
                success: false,
                need_wait_dropdown: true,
                method: 'opened_combobox'
            };
        }

        // 4. Cek native <select>
        const selects = Array.from(document.querySelectorAll('select'));
        for (const sel of selects) {
            const opts = Array.from(sel.options);
            const targetOpt = opts.find(o => optRegex.test(o.text) || o.value === '10');
            if (targetOpt) {
                const wasChecked = (sel.value === targetOpt.value);
                sel.value = targetOpt.value;
                triggerEvents(sel);
                return {
                    success: true,
                    method: 'native_select',
                    already_selected: wasChecked,
                    details: targetOpt.text
                };
            }
        }

        return { success: false, reason: 'not_found' };
    }''')

    if res and res.get("need_wait_dropdown"):
        time.sleep(1.0)
        # Cari pilihan '10. Diluar kawasan' di popover/listbox yang baru terbuka
        pop_res = page.evaluate('''() => {
            function triggerEvents(el) {
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
            }
            const optRegex = /(?:^|\s)10\s*[\.\-]?\s*(?:di\s*luar|diluar)\s*kawasan|(?:di\s*luar|diluar)\s*kawasan/i;
            const popItems = Array.from(document.querySelectorAll('[role="option"], [role="menuitem"], li, div[class*="item"], div, span'));
            const matched = popItems.filter(el => {
                const t = (el.innerText || el.textContent || '').trim();
                return optRegex.test(t) && el.offsetParent !== null;
            });
            if (matched.length > 0) {
                matched.sort((a, b) => (a.innerText || '').length - (b.innerText || '').length);
                const target = matched[0];
                target.scrollIntoView({ behavior: 'instant', block: 'center' });
                target.click();
                return { success: true, method: 'dropdown_item_click', text: (target.innerText || '').trim() };
            }
            return { success: false };
        }''')
        if pop_res and pop_res.get("success"):
            res = pop_res

    # Cek dan bersihkan field 8. e. Nama kawasan jika masih aktif / terisi
    try:
        page.evaluate('''() => {
            const all = Array.from(document.querySelectorAll('*'));
            const namaKawasanLabel = all.find(el => {
                const t = (el.innerText || el.textContent || '').trim().toLowerCase();
                return (t.includes('nama kawasan') || t.includes('8. e') || t.includes('8.e')) &&
                       !Array.from(el.children).some(c => (c.innerText || '').toLowerCase().includes('nama kawasan'));
            });
            if (namaKawasanLabel) {
                const container = namaKawasanLabel.closest('tr, fieldset, form > div, div[class*="border"], div') || namaKawasanLabel.parentElement;
                const inputEl = container.querySelector('input[type="text"], textarea');
                if (inputEl && !inputEl.disabled && inputEl.offsetParent !== null) {
                    if (inputEl.value && inputEl.value.trim() !== '') {
                        inputEl.value = '';
                        inputEl.dispatchEvent(new Event('input', { bubbles: true }));
                        inputEl.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                }
            }
        }''')
    except Exception:
        pass

    return res

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
    """Memastikan checkbox 'Tampilkan Anomali Usaha dan Keluarga' bernilai true (tercentang)."""
    print("  -> Memeriksa checkbox 'Tampilkan Anomali Usaha dan Keluarga'...")
    time.sleep(1.0)
    for _ in range(3):
        res = page.evaluate('''() => {
            function triggerChange(el) {
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
            }

            const all = Array.from(document.querySelectorAll('*'));
            const targetEl = all.find(el => {
                const t = (el.innerText || el.textContent || '').toLowerCase();
                return t.includes('tampilkan anomali') && !Array.from(el.children).some(c => (c.innerText || '').toLowerCase().includes('tampilkan anomali'));
            });

            if (targetEl) {
                let parent = targetEl.parentElement;
                for (let i = 0; i < 5; i++) {
                    if (!parent) break;
                    const cb = parent.querySelector('input[type="checkbox"]');
                    if (cb) {
                        const was_checked = cb.checked;
                        if (!was_checked) {
                            cb.scrollIntoView({ behavior: 'instant', block: 'center' });
                            cb.click();
                            cb.checked = true;
                            triggerChange(cb);
                            return { success: true, toggled: true, was_checked: false };
                        }
                        return { success: true, toggled: false, was_checked: true };
                    }
                    parent = parent.parentElement;
                }

                // Jika tidak ada input checkbox eksplisit, klik teks/label langsung
                targetEl.scrollIntoView({ behavior: 'instant', block: 'center' });
                targetEl.click();
                return { success: true, toggled: true, was_checked: false, note: "clicked_label" };
            }
            return { success: false };
        }''')

        if res and res.get("success"):
            if res.get("was_checked"):
                print("     [OK] Checkbox 'Tampilkan Anomali' sudah bernilai TRUE (tercentang).")
            else:
                print("     [OK] Berhasil mencentang checkbox 'Tampilkan Anomali Usaha dan Keluarga' (TRUE).")
            time.sleep(1.0)
            return True
        time.sleep(1.0)

    try:
        page.locator("text=/Tampilkan Anomali/i").first.click(timeout=3000, force=True)
        print("     [OK] Checkbox 'Tampilkan Anomali' dicentang via Playwright fallback.")
        time.sleep(1.0)
        return True
    except Exception:
        print("     [Info] Checkbox 'Tampilkan Anomali' tidak terdeteksi atau sudah aktif.")
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

def process_single_assignment(page, item):
    """
    Memproses satu assignment:
    1. Buka link /edit
    2. Buka menu 'SE2026 - L BLOK II' (KETERANGAN USAHA/PERUSAHAAN)
    3. Klik tombol 'Lihat Rincian'
    4. Scroll form ke bawah agar pertanyaan 8.d termuat
    5. Ubah 8. d. Jenis kawasan beroperasi ke '10. Diluar kawasan'
    6. Bersihkan nama kawasan (8.e) jika masih terisi
    7. Simpan rincian dan kirim form
    """
    link = item["link"]
    nama_usaha = item.get("nama_usaha", "")

    try:
        page.goto(link, wait_until="networkidle", timeout=30000)
    except Exception as e:
        if "interrupted by another navigation" in str(e):
            time.sleep(2)
            page.goto(link, wait_until="networkidle", timeout=30000)
        else:
            raise e

    resolve_bot_detection(page, link)

    # Verifikasi berada di halaman /edit
    current_url = page.url
    if "/edit" not in current_url:
        print(f"  -> [SKIP] Dialihkan ke {current_url}. Bukan wilayah tugas Anda / form terkunci.")
        return "SKIP_NOT_AUTHORIZED", "Bukan wilayah tugas Anda / form terkunci"

    time.sleep(1.0)

    # 1. Buka menu 'SE2026 - L BLOK II' (KETERANGAN USAHA/PERUSAHAAN)
    print(f"  -> Membuka menu 'SE2026 - L BLOK II: KETERANGAN USAHA/PERUSAHAAN'...")
    menu_ok = open_menu_blok_ii(page)
    if not menu_ok:
        print("  -> [Warning] Menu BLOK II belum terkonfirmasi terbuka. Mencoba fallback pencarian menu...")
        open_menu_usaha_bawah_p(page, item)

    time.sleep(1.5)

    # 2. Klik tombol 'Lihat Rincian'
    print(f"  -> Mencari dan mengklik tombol 'Lihat Rincian' (Usaha: '{nama_usaha}')...")
    click_lihat_rincian(page, nama_usaha)

    time.sleep(1.5)

    # 3. Scroll ke bawah pada form / modal rincian
    print("  -> Melakukan scroll ke bawah pada form rincian...")
    scroll_form_down(page)

    # 4. Ubah pertanyaan 8.d Jenis Kawasan ke '10. Diluar kawasan'
    print("  -> Mencari pertanyaan '8. d. Jenis kawasan beroperasi' dan memilih '10. Diluar kawasan'...")
    res_change = change_jenis_kawasan_to_10(page)

    if not res_change or not res_change.get("success"):
        # Jika belum ditemukan, coba scroll lebih lanjut dan coba lagi
        print("  -> Mencoba scroll ulang dan mencari kembali opsi 10...")
        scroll_form_down(page)
        res_change = change_jenis_kawasan_to_10(page)

    if not res_change or not res_change.get("success"):
        diag = page.evaluate('''() => {
            const sidebars = Array.from(document.querySelectorAll('aside, nav, [role="tablist"], [class*="sidebar"], [class*="menu"], ul'))
                .flatMap(s => Array.from(s.querySelectorAll('a, button, [role="tab"], li')))
                .map(e => (e.innerText || '').trim())
                .filter(t => t.length > 0 && t.length < 50);
            const titles = Array.from(document.querySelectorAll('h1, h2, h3, h4, h5, label, legend, p'))
                .map(e => (e.innerText || '').trim())
                .filter(t => t.length > 0 && (t.includes('8') || t.includes('kawasan') || t.includes('usaha')))
                .slice(0, 15);
            return { currentUrl: window.location.href, sidebars: Array.from(new Set(sidebars)).slice(0, 10), titles: titles };
        }''')
        print(f"  -> [GAGAL] Opsi '10. Diluar kawasan' tidak ditemukan pada form.")
        print(f"     Diagnostik: Menu Sidebar={diag.get('sidebars')}, Teks Form={diag.get('titles')}")
        return "FAILED_LOCATING_OPTION", "Tidak dapat menemukan opsi '10. Diluar kawasan'"

    method = res_change.get("method", "unknown")
    already = res_change.get("already_selected", False)
    if already:
        print(f"     [OK] Opsi '10. Diluar kawasan' sudah terpilih sebelumnya ({method}).")
    else:
        print(f"     [OK] Berhasil memilih opsi '10. Diluar kawasan' ({method}).")

    time.sleep(0.5)

    # 5. Jika ada tombol Simpan di dalam modal rincian, klik untuk simpan rincian
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

    # 6. Masuk ke menu Catatan
    open_menu_catatan(page)

    # 7. Pastikan 'Tampilkan Anomali Usaha dan Keluarga' bernilai TRUE
    ensure_tampilkan_anomali_checked(page)

    # 8. Submit dan konfirmasi bertingkat (konfirm 2x)
    print("  -> Menyimpan dan mengirimkan form...")
    kirim_ok, msg_kirim = click_kirim_and_confirm(page)
    if not kirim_ok:
        print(f"  -> [GAGAL] Pengiriman form gagal: {msg_kirim}")
        return "FAILED_SUBMIT", msg_kirim

    return "SUCCESS", "Jenis kawasan berhasil diubah ke 10. Diluar kawasan"

# ==============================================================================
# MAIN RUNNER
# ==============================================================================

def main():
    print("=" * 65)
    print(" BOT UBAH 8. d. JENIS KAWASAN BEROPERASI KE '10. DILUAR KAWASAN'")
    print(" Fasih-SM BPS (SE2026)")
    print("=" * 65)

    # 1. Cari file Excel data
    excel_files = find_excel_files()
    if not excel_files:
        print("\n[ERROR] Tidak ditemukan file Excel di folder 'data'!")
        print("Silakan letakkan file Excel (seperti 'sqllab_untitled_query_17_20261019T113729.xlsx') di folder 'data'.")
        input("\nTekan ENTER untuk keluar...")
        sys.exit(1)

    print("\nFile Excel yang terdeteksi:")
    for idx, f in enumerate(excel_files, 1):
        print(f"  [{idx}] {os.path.basename(f)}")

    selected_file = excel_files[0]
    if len(excel_files) > 1:
        choice = input(f"\nPilih nomor file [1-{len(excel_files)}] (Default: 1): ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(excel_files):
            selected_file = excel_files[int(choice) - 1]

    print(f"\n-> Membaca data dari: {os.path.basename(selected_file)}")
    rows_data = load_data_from_excel(selected_file)
    if not rows_data:
        print("[ERROR] Tidak ada baris data atau link Fasih yang valid di dalam file tersebut!")
        input("\nTekan ENTER untuk keluar...")
        sys.exit(1)

    # 2. Opsi Filter Kecamatan
    all_kecs = sorted(list(set(item["kecamatan"] for item in rows_data if item["kecamatan"])))
    print(f"\nTotal data ditemukan: {len(rows_data)} baris.")
    print("PILIHAN FILTER KECAMATAN:")
    print("  [0] SEMUA KECAMATAN (Proses seluruh data)")
    for i, k in enumerate(all_kecs, 1):
        count_k = sum(1 for it in rows_data if it["kecamatan"] == k)
        print(f"  [{i}] Kecamatan {k} ({count_k} data)")

    kec_choice = input(f"Pilih nomor kecamatan [0-{len(all_kecs)}] (Default: 0): ").strip()
    active_data = rows_data
    if kec_choice.isdigit() and int(kec_choice) > 0 and int(kec_choice) <= len(all_kecs):
        target_kec = all_kecs[int(kec_choice) - 1]
        active_data = [it for it in rows_data if it["kecamatan"] == target_kec]
        print(f"-> Filter aktif: Kecamatan {target_kec} ({len(active_data)} data)")
    else:
        print(f"-> Memproses SEMUA Kecamatan ({len(active_data)} data).")

    # 3. Cek Cache Progres
    processed_cache = load_cache()
    cached_count = sum(1 for it in active_data if it["link"] in processed_cache)
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

    # 4. Inisialisasi Playwright Browser
    profile_dir = os.path.join(BASE_DIR, "chrome_profile_anomali")
    report_data = load_report_data()
    report_map = {r.get("link"): idx for idx, r in enumerate(report_data)}

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
            kec = item["kecamatan"]
            nama = item["nama_usaha"]

            if link in processed_cache:
                print(f"[{idx}/{total_target}] [{kec}] SKIP (Sudah diproses): {nama}")
                continue

            print(f"\n[{idx}/{total_target}] [{kec}] Memproses: {nama}")
            print(f"  Link: {link}")

            status = None
            msg = ""
            try:
                res_status, res_msg = process_single_assignment(page, item)
                status = res_status
                msg = res_msg

                if status == "SUCCESS":
                    processed_cache.add(link)
                    save_cache(processed_cache)
                    success_count += 1
                    print(f"  -> [BERHASIL] 8.d Jenis Kawasan diubah ke '10. Diluar kawasan'!")

                    # Hapus dari laporan gagal sebelumnya jika ada
                    if link in report_map:
                        report_data.pop(report_map[link])
                        report_map = {r.get("link"): i for i, r in enumerate(report_data)}
                        save_report_data(report_data)

                elif status == "FAILED_SUBMIT" or status == "FAILED_LOCATING_OPTION":
                    fail_count += 1
                    rec = {
                        "waktu": time.strftime("%Y-%m-%d %H:%M:%S"),
                        "kecamatan": kec,
                        "desa": item.get("desa", ""),
                        "nama_sls": item.get("nama_sls", ""),
                        "nama_usaha": nama,
                        "assignment_id": item.get("assignment_id", ""),
                        "status_awal": item.get("status_awal", ""),
                        "jenis_kawasan_awal": item.get("jenis_kawasan_awal", ""),
                        "status_hasil": f"GAGAL - {status}",
                        "keterangan": "Gagal mengubah atau mengirim form",
                        "detail_pesan": msg,
                        "link": link
                    }
                    if link in report_map:
                        report_data[report_map[link]] = rec
                    else:
                        report_map[link] = len(report_data)
                        report_data.append(rec)
                    save_report_data(report_data)
                    print(f"  -> [GAGAL] {msg}")

                elif status.startswith("SKIP"):
                    skip_count += 1
                    processed_cache.add(link)
                    save_cache(processed_cache)
                    rec = {
                        "waktu": time.strftime("%Y-%m-%d %H:%M:%S"),
                        "kecamatan": kec,
                        "desa": item.get("desa", ""),
                        "nama_sls": item.get("nama_sls", ""),
                        "nama_usaha": nama,
                        "assignment_id": item.get("assignment_id", ""),
                        "status_awal": item.get("status_awal", ""),
                        "jenis_kawasan_awal": item.get("jenis_kawasan_awal", ""),
                        "status_hasil": status,
                        "keterangan": msg,
                        "detail_pesan": msg,
                        "link": link
                    }
                    if link in report_map:
                        report_data[report_map[link]] = rec
                    else:
                        report_map[link] = len(report_data)
                        report_data.append(rec)
                    save_report_data(report_data)
                    print(f"  -> [SKIP] {msg}")

            except Exception as e:
                fail_count += 1
                print(f"  -> [ERROR] Terjadi exception: {e}")
                if check_is_bot_or_blocked(page):
                    resolve_bot_detection(page, link)

            # Jeda antar link untuk simulasi manusia dan keamanan WAF
            if status and status.startswith("SKIP"):
                delay = random.uniform(1.5, 3.0)
            else:
                delay = random.uniform(4.5, 7.5)
            time.sleep(delay)

            # Cooling down setiap 15 link
            if idx % 15 == 0 and idx < total_target:
                rest_time = random.uniform(30, 45)
                print("=" * 60)
                print(f" [COOLING DOWN] Telah memproses {idx} link. Jeda sejenak {int(rest_time)} detik...")
                print("=" * 60)
                time.sleep(rest_time)

        print("\n" + "=" * 65)
        print(" PROSES OTOMATISASI SELESAI!")
        print(f" - Sukses Diproses  : {success_count}")
        print(f" - Di-skip          : {skip_count}")
        print(f" - Gagal/Error      : {fail_count}")
        print(f" - File Laporan     : {REPORT_EXCEL}")
        print("=" * 65)
        context.close()

if __name__ == "__main__":
    main()
