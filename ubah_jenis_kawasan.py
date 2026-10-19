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

def open_menu_se2026_p(page):
    """Membuka atau memastikan form 'SE2026 - P' aktif."""
    def is_form_rendered():
        try:
            return page.evaluate('''() => {
                const text = (document.body.innerText || '').toLowerCase();
                return text.includes("keberadaan bangunan") || 
                       text.includes("jenis kawasan") || 
                       text.includes("kawasan beroperasi") ||
                       text.includes("1. ditemukan");
            }''')
        except Exception:
            return False

    if is_form_rendered():
        return True

    pattern = re.compile(r"SE2026\s*[-–—]?\s*P", re.IGNORECASE)
    candidates = [
        page.get_by_role("tab", name=pattern),
        page.get_by_role("link", name=pattern),
        page.get_by_role("button", name=pattern),
        page.locator("a, button, [role='tab'], [role='menuitem'], [role='treeitem'], li").filter(has_text=pattern),
        page.locator("[title*='SE2026 - P'], [title*='SE2026-P'], [title*='SE2026']")
    ]
    for cand in candidates:
        try:
            cnt = cand.count()
            if cnt > 0:
                for i in range(cnt):
                    el = cand.nth(i)
                    if el.is_visible():
                        box = el.bounding_box()
                        if box:
                            page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                        else:
                            el.click(timeout=3000)
                        time.sleep(1.5)
                        if is_form_rendered():
                            return True
        except Exception:
            pass

    # Fallback JS Deep Click
    page.evaluate('''() => {
        const all = Array.from(document.querySelectorAll('*'));
        const matches = all.filter(el => {
            const txt = (el.innerText || el.textContent || '').trim();
            return /SE2026\s*[-–—]?\s*P/i.test(txt);
        });
        if (matches.length > 0) {
            matches.sort((a, b) => (a.innerText || '').length - (b.innerText || '').length);
            const leaf = matches[0];
            const clickable = leaf.closest('a, button, [role="tab"], [role="treeitem"], [role="menuitem"], li, [tabindex]') || leaf;
            clickable.scrollIntoView({ behavior: 'instant', block: 'center' });
            clickable.click();
        }
    }''')
    time.sleep(2)
    return is_form_rendered()

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

def click_kirim_and_confirm(page):
    """
    Mengklik tombol 'Kirim' dan menangani modal konfirmasi jika muncul,
    serta tombol split dropdown 'Submit Paksa' jika diperlukan.
    """
    time.sleep(1.0)

    # 1. Coba klik tombol Kirim
    kirim_clicked = False
    try:
        kirim_btn = page.get_by_role("button", name=re.compile(r"^Kirim$", re.IGNORECASE)).first
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
                return t === 'kirim' || t === 'simpan' || t.includes('kirim');
            });
            if (target) {
                target.scrollIntoView({ behavior: 'instant', block: 'center' });
                target.click();
                return true;
            }
            return false;
        }''')

    time.sleep(1.5)

    # 2. Tangani dialog / modal konfirmasi jika muncul (Tombol Kirim / Ya / Konfirmasi)
    modal_handled = False
    try:
        modal_handled = page.evaluate('''() => {
            const btns = Array.from(document.querySelectorAll('button')).filter(b => {
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
                    t.includes('ya, simpan')
                );
            });
            if (confirmBtn) {
                confirmBtn.scrollIntoView({ behavior: 'instant', block: 'center' });
                confirmBtn.click();
                return true;
            }
            return false;
        }''')
    except Exception:
        pass

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
    """Memproses satu assignment: buka /edit, buka SE2026-P, ubah 8.d ke kode 10, simpan."""
    link = item["link"]

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

    # 1. Buka menu 'SE2026 - P'
    print("  -> Memeriksa dan membuka tab menu 'SE2026 - P'...")
    menu_ok = open_menu_se2026_p(page)
    if not menu_ok:
        print("  -> [Warning] Tampilan belum terkonfirmasi berubah ke 'SE2026 - P'. Mencoba mencari pertanyaan...")

    time.sleep(1.0)

    # 2. Ubah pertanyaan 8.d Jenis Kawasan ke '10. Diluar kawasan'
    print("  -> Mencari pertanyaan '8. d. Jenis kawasan beroperasi' dan memilih '10. Diluar kawasan'...")
    res_change = change_jenis_kawasan_to_10(page)

    if not res_change or not res_change.get("success"):
        # Coba scroll ke bawah form dan coba sekali lagi jika form lazy-loaded
        page.evaluate("window.scrollTo(0, document.body.scrollHeight / 2)")
        time.sleep(1.0)
        res_change = change_jenis_kawasan_to_10(page)

    if not res_change or not res_change.get("success"):
        diag = page.evaluate('''() => {
            const text = (document.body.innerText || '').substring(0, 500);
            return { currentUrl: window.location.href, preview: text };
        }''')
        print(f"  -> [GAGAL] Opsi '10. Diluar kawasan' tidak ditemukan pada form. (URL: {diag.get('currentUrl')})")
        return "FAILED_LOCATING_OPTION", "Tidak dapat menemukan opsi '10. Diluar kawasan'"

    method = res_change.get("method", "unknown")
    already = res_change.get("already_selected", False)
    if already:
        print(f"     [OK] Opsi '10. Diluar kawasan' sudah terpilih sebelumnya ({method}).")
    else:
        print(f"     [OK] Berhasil memilih opsi '10. Diluar kawasan' ({method}).")

    # 3. Klik Kirim dan Konfirmasi
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
            viewport={"width": 1366, "height": 768},
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
