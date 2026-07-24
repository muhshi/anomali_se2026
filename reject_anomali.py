import pandas as pd
from playwright.sync_api import sync_playwright
import re
import time
import json
import os
import random

CACHE_FILE = "processed_links.json"

def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r") as f:
            return set(json.load(f))
    return set()

def save_cache(cache_set):
    with open(CACHE_FILE, "w") as f:
        json.dump(list(cache_set), f, indent=4)

def main():
    data_dir = os.path.join(os.getcwd(), "data")
    
    # Buat folder data jika belum ada
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)
        print(f"Folder 'data' telah dibuat di {data_dir}")
        print("Silakan masukkan file Excel anomali ke dalam folder tersebut lalu jalankan ulang script.")
        return
        
    excel_files = [f for f in os.listdir(data_dir) if f.endswith(".xlsx")]
    if not excel_files:
        print("TIDAK ADA FILE EXCEL DITEMUKAN!")
        print(f"Silakan masukkan file Excel anomali (.xlsx) ke dalam folder: {data_dir}")
        return

    # Looping semua file excel yang ada di folder data
    raw_links = []
    for file_name in excel_files:
        excel_file = os.path.join(data_dir, file_name)
        print(f"Membaca file Excel: {file_name}...")
        try:
            # Membaca kolom R (link) secara langsung
            df = pd.read_excel(excel_file, usecols="R")
            links = df.iloc[:, 0].dropna().astype(str).tolist()
            raw_links.extend(links)
        except Exception as e:
            print(f"  -> Gagal membaca file {file_name}: {e}")

    if not raw_links:
        print("Tidak ada data yang berhasil dibaca dari file Excel mana pun.")
        return

    # Filter dan bentuk ulang link menjadi format /edit
    kegiatan_id = "fd68e454-ba45-4b85-8205-f3bf777ded24"
    edit_links = []
    
    for link in raw_links:
        link = link.strip()
        # Ekstrak ID dinamis dari link awal
        match = re.search(r"assignment-detail/([a-zA-Z0-9\-]+)", link)
        if match:
            dynamic_id = match.group(1)
            # Buat link edit baru
            new_link = f"https://fasih-sm.bps.go.id/app/assignment/{kegiatan_id}/{dynamic_id}/edit"
            edit_links.append(new_link)

    print(f"Ditemukan {len(edit_links)} link anomali dalam file Excel.")

    if not edit_links:
        print("Tidak ada link valid yang ditemukan di file Excel.")
        return

    processed_cache = load_cache()
    
    # Filter link yang belum diproses untuk mengetahui sisa pekerjaan
    pending_links = [link for link in edit_links if link not in processed_cache]
    
    print(f"Total link: {len(edit_links)}")
    print(f"Sudah diproses (dari cache): {len(processed_cache)}")
    print(f"Sisa yang akan diproses: {len(pending_links)}")
    
    if not pending_links:
        print("Semua data untuk kodekec ini sudah berhasil diproses!")
        return

    print("="*60)
    print("Membuka browser Playwright...")
    
    with sync_playwright() as p:
        user_data_dir = os.path.join(os.getcwd(), "chrome_profile_anomali")
        context = p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=False,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--start-maximized"
            ],
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            no_viewport=True
        )
        
        # Gunakan page pertama yang terbuka otomatis
        page = context.pages[0] if context.pages else context.new_page()
        
        # Bypass webdriver sesuai screenshotBot.py
        page.add_init_script("delete navigator.__proto__.webdriver;")

        page.goto("https://fasih-sm.bps.go.id/")
        
        print("Silakan login ke Fasih-SM (SSO BPS) jika belum login.")
        print("="*60)
        input("Tekan ENTER di terminal ini jika sudah berhasil login dan siap memulai...")

        for idx, link in enumerate(edit_links, 1):
            if link in processed_cache:
                print(f"[{idx}/{len(edit_links)}] SKIP (Sudah diproses): {link}")
                continue
                
            print(f"[{idx}/{len(edit_links)}] Memproses: {link}")
            try:
                # Coba navigasi ke link, tangani jika terpotong oleh redirect SSO (biasanya di link pertama)
                try:
                    page.goto(link, wait_until="networkidle")
                except Exception as e:
                    if "interrupted by another navigation" in str(e):
                        print("  -> [Info] Navigasi link pertama dialihkan ke Dashboard (SSO Refresh). Mencoba ulang...")
                        time.sleep(3) # Tunggu redirect selesai
                        page.goto(link, wait_until="networkidle")
                    else:
                        raise e
                
                # Cek apakah terkena blokir bot (WAF)
                if page.locator("text=/mendeteksi koneksi anda sebagai bot/i").count() > 0 or page.locator("text=/HaloSIS/i").count() > 0:
                    print("="*60)
                    print(" TERDETEKSI SEBAGAI BOT OLEH SERVER (WAF BPS)!")
                    print(" Ini biasanya karena kita membuka terlalu banyak halaman dalam waktu singkat.")
                    print(" Silakan selesaikan tantangan/tunggu di browser, lalu tekan ENTER di sini untuk lanjut.")
                    print("="*60)
                    input("Tekan ENTER jika sudah lepas dari halaman blokir bot...")
                    page.goto(link, wait_until="networkidle")
                
                # 1. Masuk ke menu catatan
                try:
                    # Perpanjang timeout menjadi 10 detik karena kadang loading lambat
                    page.get_by_role("tab", name="Catatan", exact=False).click(timeout=10000)
                except:
                    page.get_by_text("Catatan", exact=False).first.click(timeout=10000)
                time.sleep(1)
                
                # 2. Centang checkbox "Tampilkan Anomali Usaha dan Keluarga"
                # Kita gunakan JavaScript agar bisa menembus elemen input yang mungkin disembunyikan (visually hidden) oleh sistem UI (seperti React/MUI)
                check_result = page.evaluate('''() => {
                    // Cari semua elemen dan temukan yang teksnya mengandung "Tampilkan Anomali" dan tidak punya child element (elemen teks terbawah)
                    const elements = Array.from(document.querySelectorAll('*'));
                    const targetEl = elements.find(el => el.textContent && el.textContent.toLowerCase().includes('tampilkan anomali') && el.children.length === 0);
                    
                    if (targetEl) {
                        // Telusuri parent elemennya ke atas untuk mencari input checkbox
                        let parent = targetEl.parentElement;
                        for(let i=0; i<5; i++) {
                            if(!parent) break;
                            const cb = parent.querySelector('input[type="checkbox"]');
                            if(cb) {
                                const was_checked = cb.checked;
                                if (!was_checked) {
                                    cb.click(); // Klik checkboxnya secara paksa lewat JS
                                }
                                return { found: true, was_checked: was_checked };
                            }
                            parent = parent.parentElement;
                        }
                        
                        // Jika tidak ketemu inputnya, paksa klik elemen teksnya saja sebagai gantinya
                        targetEl.click();
                        return { found: true, was_checked: false, note: "clicked_text_only" };
                    }
                    return { found: false, was_checked: false };
                }''')
                
                if check_result and check_result.get('found'):
                    if check_result.get('was_checked'):
                        print("  -> Checkbox sudah tercentang dari awal. Melewati klik 'Kirim'.")
                    else:
                        # 3. Klik Kirim (Tombol utama)
                        page.get_by_role("button", name="Kirim").first.click()
                        time.sleep(1.5) # Tunggu popup muncul
                        
                        # 4. Handle Konfirmasi (Bisa muncul 2x popup beruntun)
                        for _ in range(3):
                            try:
                                # Cari tombol yang berpotensi menjadi konfirmasi (Kirim, Ya, Konfirmasi, dll)
                                btns = page.locator("button").filter(has_text=re.compile(r"Kirim|Ya|Konfirmasi|Setuju", re.IGNORECASE)).all()
                                visible_btns = [b for b in btns if b.is_visible()]
                                
                                # Jika ada > 1 tombol (berarti ada 1 tombol utama + tombol di modal)
                                if len(visible_btns) > 1:
                                    print("     -> Mengklik tombol konfirmasi pada dialog...")
                                    visible_btns[-1].click(timeout=3000)
                                    time.sleep(1.5) # Tunggu dialog berikutnya (jika ada) atau loading
                                else:
                                    break # Tidak ada dialog lagi
                            except Exception as e:
                                break
                        
                        time.sleep(2)
                else:
                    # Fallback Playwright murni jika script JS gagal menemukan teks
                    try:
                        print("  -> (Fallback) Mencoba klik paksa teks 'Tampilkan Anomali'...")
                        page.locator("text=/Tampilkan Anomali/i").first.click(timeout=5000, force=True)
                        page.get_by_role("button", name="Kirim").first.click()
                        time.sleep(1)
                        page.get_by_role("button", name="Kirim").last.click()
                        time.sleep(2)
                    except:
                        pass
                
                # --- FASE 2: REJECT ---
                # Hapus "/edit" dari link untuk kembali ke halaman assignment detail
                base_link = link.replace("/edit", "")
                print(f"  -> Membuka halaman detail untuk Reject: {base_link}")
                
                try:
                    page.goto(base_link, wait_until="networkidle")
                except Exception as e:
                    if "interrupted by another navigation" in str(e):
                        time.sleep(2)
                        page.goto(base_link, wait_until="networkidle")
                    else:
                        raise e
                        
                time.sleep(2)
                
                # Klik tombol Reject atau X
                print("  -> Mengklik tombol Reject / X...")
                reject_success = False
                for _ in range(10): # Coba cari tombolnya selama 10 detik (loading kadang lama)
                    clicked_reject = page.evaluate('''() => {
                        const buttons = Array.from(document.querySelectorAll('button'));
                        
                        // 1. Cari berdasarkan teks, aria-label, atau title
                        let target = buttons.find(b => {
                            const t = (b.textContent || '').toLowerCase().trim();
                            const a = (b.getAttribute('aria-label') || '').toLowerCase();
                            const title = (b.getAttribute('title') || '').toLowerCase();
                            return t.includes('reject') || t.includes('tolak') || t === 'x' || t === '×' || 
                                   a.includes('reject') || a.includes('tolak') || title.includes('reject');
                        });
                        
                        // 2. Jika tidak ketemu, cari tombol warna merah atau dengan class danger/reject/destructive
                        if (!target) {
                            const dangerBtns = buttons.filter(b => {
                                const c = (b.className || '').toLowerCase();
                                return c.includes('reject') || c.includes('danger') || c.includes('red-') || c.includes('destructive');
                            });
                            
                            // Jika ada tombol destructive, ambil yang paling bawah posisinya di layar (biasanya tombol floating/di pojok)
                            if (dangerBtns.length > 0) {
                                target = dangerBtns.reduce((prev, curr) => {
                                    return (curr.getBoundingClientRect().bottom > prev.getBoundingClientRect().bottom) ? curr : prev;
                                });
                            }
                        }
                        
                        // 3. Jika tidak ketemu, cari tombol yang melayang di pojok kanan bawah
                        if (!target) {
                            target = buttons.find(b => {
                                const rect = b.getBoundingClientRect();
                                const isBottomRight = rect.left > window.innerWidth * 0.7 && rect.top > window.innerHeight * 0.7;
                                // Kadang tombol ada di dalam container yang fixed, jadi kita longgarkan syarat fixed
                                return isBottomRight;
                            });
                        }
                        
                        if (target) {
                            target.click();
                            return true;
                        }
                        return false;
                    }''')
                    
                    if clicked_reject:
                        reject_success = True
                        break
                    time.sleep(1) # Tunggu 1 detik sebelum coba lagi
                
                if not reject_success:
                    raise Exception("Gagal menemukan tombol Reject/X di halaman setelah menunggu 10 detik.")
                
                time.sleep(1.5) # Tunggu animasi dialog konfirmasi muncul
                
                # Klik tombol Konfirmasi di dialog
                print("  -> Mengklik Konfirmasi...")
                confirm_success = False
                for _ in range(10): # Coba cari selama 10 detik
                    clicked_confirm = page.evaluate('''() => {
                        // Dialog sering kali dirender di akhir body, kita ambil semua tombol
                        const buttons = Array.from(document.querySelectorAll('button'));
                        
                        // Cari tombol yang mengandung kata konfirmasi, ya, setuju
                        let target = buttons.find(b => {
                            const t = (b.textContent || '').toLowerCase().trim();
                            return t.includes('konfirmasi') || t.includes('ya') || t.includes('setuju') || t.includes('lanjut');
                        });
                        
                        // Alternatif: klik tombol terakhir (biasanya tombol aksi utama di modal) jika ada elemen dialog/modal yang aktif
                        if (!target) {
                            const modal = document.querySelector('.modal, [role="dialog"], [data-state="open"]');
                            if (modal) {
                                const modalBtns = Array.from(modal.querySelectorAll('button'));
                                if (modalBtns.length > 0) {
                                    target = modalBtns[modalBtns.length - 1]; // Tombol terakhir
                                }
                            }
                        }
                        
                        if (target) {
                            target.click();
                            return true;
                        }
                        return false;
                    }''')
                    
                    if clicked_confirm:
                        confirm_success = True
                        break
                    time.sleep(1)
                
                if not confirm_success:
                    print("     [Warning] Tombol Konfirmasi tidak ditemukan. Melanjutkan (mungkin langsung tersimpan otomatis).")
                
                time.sleep(2.5) # Tunggu proses loading reject selesai ke server
                
                # 5. Jika SEMUA LANGKAH sukses, baru simpan ke cache
                processed_cache.add(link)
                save_cache(processed_cache)
                print(f"  -> Sukses dikirim, direject, dan disimpan ke cache.")
                
            except Exception as e:
                print(f"  -> Terjadi error pada link {link}:")
                print(f"     {e}")
                print("     Lanjut ke link berikutnya...")

            # Jeda acak 3-7 detik antar link untuk mensimulasikan kecepatan manusia
            # Ini sangat penting agar IP tidak di-blacklist oleh firewall server BPS
            delay = random.uniform(3, 7)
            time.sleep(delay)

        print("Proses otomatisasi selesai!")
        context.close()

if __name__ == "__main__":
    main()
