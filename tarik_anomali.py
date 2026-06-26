import json
import time
import os
import random
from datetime import datetime
import undetected_chromedriver as uc
from selenium.webdriver.chrome.options import Options

# ========================================
# Cache Harian (resume-friendly)
# ========================================

def get_cache_path(date_str):
    return f"cache_{date_str}.json"

def load_cache(date_str):
    cache_path = get_cache_path(date_str)
    if os.path.exists(cache_path):
        with open(cache_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        print(f"[i] Cache ditemukan: {len(data)} URL tersimpan dari sesi sebelumnya.")
        return data
    return {}

def save_cache(date_str, cache):
    cache_path = get_cache_path(date_str)
    with open(cache_path, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False)

# ========================================
# Penarikan Data Anomali SE2026
# ========================================

def load_config():
    config_path = 'config.json'
    if not os.path.exists(config_path):
        print(f"Error: {config_path} tidak ditemukan!")
        exit(1)
    with open(config_path, 'r') as f:
        return json.load(f)

def init_db(config):
    db_cfg = config.get('database', {})
    engine = db_cfg.get('engine', 'sqlite').lower()
    
    if engine == 'sqlite':
        import sqlite3
        db_file = db_cfg.get('sqlite_file', 'anomali_se2026.db')
        print(f"Menggunakan SQLite: {db_file}")
        connection = sqlite3.connect(db_file)
        cursor = connection.cursor()
        
        # Tabel Agregat
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS agregat_anomali (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal_tarik DATE NOT NULL,
            id_wilayah VARCHAR(20),
            nama_wilayah VARCHAR(100),
            level_wilayah VARCHAR(20),
            kode_indikator VARCHAR(20),
            total_value INT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE (tanggal_tarik, id_wilayah, kode_indikator)
        )
        """)
        
        # Tabel Mikro
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS kasus_anomali_mikro (
            id_db INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal_tarik DATE NOT NULL,
            id_kasus VARCHAR(200),
            source_table VARCHAR(100),
            source_type VARCHAR(50),
            anomali_no INT,
            id_indikator INT,
            anomali_title TEXT,
            assignment_id VARCHAR(100),
            link_fasih TEXT,
            kode_wilayah VARCHAR(20),
            is_resolved INT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE (tanggal_tarik, id_kasus)
        )
        """)
        connection.commit()
        return connection, 'sqlite'
        
    elif engine == 'mysql':
        import mysql.connector
        print("Menggunakan MySQL...")
        connection = mysql.connector.connect(
            host=db_cfg['host'],
            database=db_cfg['database_name'],
            user=db_cfg['user'],
            password=db_cfg['password']
        )
        cursor = connection.cursor()
        
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS agregat_anomali (
            id BIGINT AUTO_INCREMENT PRIMARY KEY,
            tanggal_tarik DATE NOT NULL,
            id_wilayah VARCHAR(20),
            nama_wilayah VARCHAR(100),
            level_wilayah VARCHAR(20),
            kode_indikator VARCHAR(20),
            total_value INT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            UNIQUE KEY unique_agregat (tanggal_tarik, id_wilayah, kode_indikator)
        )
        """)
        
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS kasus_anomali_mikro (
            id_db BIGINT AUTO_INCREMENT PRIMARY KEY,
            tanggal_tarik DATE NOT NULL,
            id_kasus VARCHAR(200),
            source_table VARCHAR(100),
            source_type VARCHAR(50),
            anomali_no INT,
            id_indikator INT,
            anomali_title TEXT,
            assignment_id VARCHAR(100),
            link_fasih TEXT,
            kode_wilayah VARCHAR(20),
            is_resolved INT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            UNIQUE KEY unique_mikro (tanggal_tarik, id_kasus)
        )
        """)
        connection.commit()
        return connection, 'mysql'
    else:
        print("Engine database tidak dikenali")
        exit(1)

def open_browser_and_login():
    print("\n" + "="*60)
    print("  MEMBUKA BROWSER CHROME...")
    print("="*60)
    
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument("--start-maximized")
    
    driver = uc.Chrome(options=chrome_options, version_main=149)
    driver.get("https://dashboard-se2026.apps.bps.go.id")
    time.sleep(5)
    
    print("\n" + "="*60)
    print("  INSTRUKSI:")
    print("  1. Login ke Dashboard SE2026 di browser")
    print("  2. Pastikan sudah masuk ke halaman utama")
    print("  3. Kembali ke terminal ini")
    print("  4. Tekan ENTER untuk mulai menarik data anomali")
    print("="*60)
    input("\n>>> Tekan ENTER setelah login berhasil... ")
    
    return driver

# Counter error HTML berturut-turut (login expired detection)
_html_error_count = 0
_HTML_ERROR_THRESHOLD = 3


def fetch_api_get(driver, url, max_retries=3):
    global _html_error_count
    js_script = """
    var callback = arguments[arguments.length - 1];
    var url = arguments[0];
    
    fetch(url, {
        method: 'GET',
        headers: {
            'Accept': '*/*'
        },
        credentials: 'same-origin'
    })
    .then(function(response) {
        return response.text().then(function(text) {
            callback({status: response.status, body: text});
        });
    })
    .catch(function(err) {
        callback({error: err.message || err.toString()});
    });
    """
    
    for attempt in range(1, max_retries + 1):
        try:
            driver.set_script_timeout(120)
            result = driver.execute_async_script(js_script, url)
            
            if not result or "error" in result:
                print(f"    [!] Fetch error: {result.get('error', 'Unknown')}")
                return None
            
            status = result.get("status", 0)
            body = result.get("body", "")
            
            if "<!DOCTYPE" in body or "<html>" in body.lower():
                _html_error_count += 1
                print(f"\n    [!] Halaman HTML terdeteksi - kemungkinan LOGIN EXPIRED [{_html_error_count}/{_HTML_ERROR_THRESHOLD}]")
                
                if _html_error_count >= _HTML_ERROR_THRESHOLD:
                    print("\n" + "!"*60)
                    print("  SESSION EXPIRED!")
                    print("  Silakan login ulang di browser Chrome yang terbuka,")
                    print("  pastikan sudah masuk ke halaman dashboard,")
                    print("  lalu kembali ke sini dan tekan ENTER untuk melanjutkan.")
                    print("!"*60)
                    input("\n>>> Tekan ENTER setelah login ulang... ")
                    _html_error_count = 0  # reset counter
                    continue  # coba ulang request yang sama
                return None
                
            if status >= 500:
                time.sleep(attempt * 3)
                continue
            if status != 200:
                print(f"    [!] HTTP {status}")
                return None
            
            # Sukses - reset counter HTML error
            _html_error_count = 0
            return json.loads(body)
            
        except Exception as e:
            time.sleep(attempt * 3)
            
    return None


def fetch_with_cache(driver, url, cache, date_str):
    """Fetch dengan cache harian. Kalau URL sudah ada di cache, langsung return tanpa delay."""
    if url in cache:
        return cache[url]  # Hit cache, skip API call
    
    result = fetch_api_get(driver, url)
    if result is not None:
        cache[url] = result
        save_cache(date_str, cache)  # Simpan setelah setiap fetch berhasil
    return result

def upsert_agregat(connection, engine_type, current_date, data_list):
    if not data_list: return 0
    rows = []
    for d in data_list:
        rows.append((
            current_date, 
            d.get("id_wilayah", ""), 
            d.get("nama_wilayah", ""), 
            d.get("level_wilayah", ""),
            str(d.get("kode_indikator", "")),
            d.get("total_value") or 0
        ))
        
    cursor = connection.cursor()
    try:
        if engine_type == 'mysql':
            q = """
            INSERT INTO agregat_anomali (
                tanggal_tarik, id_wilayah, nama_wilayah, level_wilayah, kode_indikator, total_value
            ) VALUES (%s, %s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE 
                nama_wilayah = VALUES(nama_wilayah),
                level_wilayah = VALUES(level_wilayah),
                total_value = VALUES(total_value),
                updated_at = NOW()
            """
        else:
            q = """
            INSERT INTO agregat_anomali (
                tanggal_tarik, id_wilayah, nama_wilayah, level_wilayah, kode_indikator, total_value
            ) VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(tanggal_tarik, id_wilayah, kode_indikator) DO UPDATE SET
                nama_wilayah = excluded.nama_wilayah,
                level_wilayah = excluded.level_wilayah,
                total_value = excluded.total_value,
                updated_at = CURRENT_TIMESTAMP
            """
        cursor.executemany(q, rows)
        connection.commit()
        return len(rows)
    except Exception as e:
        print(f"    [!] Error upsert agregat: {e}")
        return 0

def upsert_mikro(connection, engine_type, current_date, data_list):
    if not data_list: return 0
    rows = []
    for d in data_list:
        rows.append((
            current_date,
            d.get("id", ""),
            d.get("source_table", ""),
            d.get("source_type", ""),
            d.get("anomali_no", 0),
            d.get("id_indikator", 0),
            d.get("anomali_title", ""),
            d.get("assignment_id", ""),
            d.get("link_fasih", ""),
            d.get("kode_wilayah", ""),
            1 if d.get("is_resolved") else 0
        ))
        
    cursor = connection.cursor()
    try:
        if engine_type == 'mysql':
            q = """
            INSERT INTO kasus_anomali_mikro (
                tanggal_tarik, id_kasus, source_table, source_type, anomali_no, 
                id_indikator, anomali_title, assignment_id, link_fasih, kode_wilayah, is_resolved
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE 
                is_resolved = VALUES(is_resolved),
                updated_at = NOW()
            """
        else:
            q = """
            INSERT INTO kasus_anomali_mikro (
                tanggal_tarik, id_kasus, source_table, source_type, anomali_no, 
                id_indikator, anomali_title, assignment_id, link_fasih, kode_wilayah, is_resolved
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(tanggal_tarik, id_kasus) DO UPDATE SET
                is_resolved = excluded.is_resolved,
                updated_at = CURRENT_TIMESTAMP
            """
        cursor.executemany(q, rows)
        connection.commit()
        return len(rows)
    except Exception as e:
        print(f"    [!] Error upsert mikro: {e}")
        return 0

def delay():
    time.sleep(random.uniform(2, 4))

def main():
    config = load_config()
    driver = open_browser_and_login()
    
    print("\nMenghubungkan ke Database...")
    connection, engine_type = init_db(config)
    
    current_date = datetime.now().strftime('%Y-%m-%d')
    indikator_list = "40,128,41,129,42,130,43,131,44,132,45,133,46,134,135"
    kabupaten_code = "3321" # DEMAK
    
    base_url = "https://dashboard-se2026.apps.bps.go.id/api/agregat/fasih"
    mikro_url = "https://dashboard-se2026.apps.bps.go.id/api/mikro/anomali-case"
    
    print(f"\nMulai menarik data anomali untuk tanggal: {current_date}")
    print("-" * 50)
    
    # Load cache harian untuk resume
    cache = load_cache(current_date)
    
    try:
        # LEVEL KECAMATAN
        print("[*] Fetch Level Kecamatan...")
        url_kec = f"{base_url}?level=kecamatan&jenis=kualitas&indikator={indikator_list}&kabupaten={kabupaten_code}"
        data_kec = fetch_with_cache(driver, url_kec, cache, current_date)
        
        if not data_kec:
            print("Gagal fetch kecamatan.")
            return
            
        upsert_agregat(connection, engine_type, current_date, data_kec)
        
        # Hanya proses kecamatan yang punya setidaknya 1 total_value > 0
        kec_ids = set([d["id_wilayah"] for d in data_kec if d.get("total_value") and int(d.get("total_value", 0)) > 0])
        print(f"  [i] {len(kec_ids)} kecamatan punya anomali (dari {len(set(d['id_wilayah'] for d in data_kec))} total)")
        
        for kec in sorted(kec_ids):  # sorted untuk urutan konsisten saat resume
            if not kec or len(kec) < 7: continue
            
            url_desa = f"{base_url}?level=desa&jenis=kualitas&indikator={indikator_list}&kecamatan={kec[:7]}"
            cached_desa = url_desa in cache
            print(f"  [-] Fetch Desa di Kecamatan {kec}{' [cache]' if cached_desa else ''}")
            if not cached_desa: delay()
            data_desa = fetch_with_cache(driver, url_desa, cache, current_date)
            
            if data_desa:
                upsert_agregat(connection, engine_type, current_date, data_desa)
                # Hanya proses desa yang punya setidaknya 1 total_value > 0
                desa_ids = set([d["id_wilayah"] for d in data_desa if d.get("total_value") and int(d.get("total_value", 0)) > 0])
                print(f"    [i] {len(desa_ids)} desa punya anomali di kecamatan {kec[:7]}")
                
                for desa in sorted(desa_ids):
                    if not desa or len(desa) < 10: continue
                    url_sls = f"{base_url}?level=sls&jenis=kualitas&indikator={indikator_list}&desa={desa[:10]}"
                    cached_sls = url_sls in cache
                    print(f"    [-] Fetch SLS di Desa {desa}{' [cache]' if cached_sls else ''}")
                    if not cached_sls: delay()
                    data_sls = fetch_with_cache(driver, url_sls, cache, current_date)
                    
                    if data_sls:
                        upsert_agregat(connection, engine_type, current_date, data_sls)
                        # Hanya proses SLS yang punya setidaknya 1 total_value > 0
                        sls_ids = set([d["id_wilayah"] for d in data_sls if d.get("total_value") and int(d.get("total_value", 0)) > 0])
                        print(f"      [i] {len(sls_ids)} SLS punya anomali di desa {desa[:10]}")
                        
                        for sls in sorted(sls_ids):
                            if not sls or len(sls) < 14: continue
                            url_subsls = f"{base_url}?level=sub_sls&jenis=kualitas&indikator={indikator_list}&sls={sls[:14]}"
                            cached_subsls = url_subsls in cache
                            print(f"      [-] Fetch Sub-SLS di SLS {sls}{' [cache]' if cached_subsls else ''}")
                            if not cached_subsls: delay()
                            data_subsls = fetch_with_cache(driver, url_subsls, cache, current_date)
                            
                            if data_subsls:
                                upsert_agregat(connection, engine_type, current_date, data_subsls)
                                
                                # Cek Sub-SLS mana saja yang BENAR-BENAR ada total_value > 0
                                # Karena anomali-case butuh 'kode_wilayah' (16 digit) dan 'indikator' spesifik
                                for item in data_subsls:
                                    total_val = item.get("total_value")
                                    if total_val and int(total_val) > 0:
                                        subsls_code = item["id_wilayah"]
                                        ind = item["kode_indikator"]
                                        
                                        url_mikro = f"{mikro_url}?kode_wilayah={subsls_code}&indikator={ind}"
                                        
                                        # Pasangan indikator (belum ditindaklanjuti dgn sudah ditindaklanjuti)
                                        pasangan = {
                                            "128": "40", "129": "41", "130": "42", 
                                            "131": "43", "132": "44", "133": "45", 
                                            "134": "46", "135": "47"
                                        }
                                        if ind in pasangan:
                                            url_mikro += f"&sudah_indikator={pasangan[ind]}"
                                        
                                        cached_mikro = url_mikro in cache
                                        print(f"        [>] Fetch Kasus Anomali di Sub-SLS {subsls_code} Indikator {ind}{' [cache]' if cached_mikro else ''}")
                                        if not cached_mikro: delay()
                                            
                                        data_kasus = fetch_with_cache(driver, url_mikro, cache, current_date)
                                        if data_kasus:
                                            inserted = upsert_mikro(connection, engine_type, current_date, data_kasus)
                                            print(f"            + Disimpan {inserted} kasus mikro.")
                                            
    except Exception as e:
        import traceback
        print(f"Error utama: {e}")
        traceback.print_exc()
    finally:
        if 'connection' in locals():
            try:
                connection.close()
            except:
                pass
        print(f"\nSelesai. Cache tersimpan di: {get_cache_path(current_date)}")
        print("Browser tetap terbuka.")

if __name__ == "__main__":
    main()
