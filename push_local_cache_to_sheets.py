#!/usr/bin/env python3
"""
Script Utilitas: Push Seluruh Histori Pengerjaan Lokal ke Google Sheets
Mencocokkan data dari laporan_tambah_remark.json, laporan_tambah_remark (2).json,
dan processed_tambah_remark*.json ke Google Sheets Priority 1 & 2.
Mendukung Batch Update (v3 Ultra Fast) dan fallback single-update.
"""

import os
import sys
import json
import time
import re
import urllib.request
import urllib.parse
import csv

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_1 = os.path.join(BASE_DIR, "laporan_tambah_remark.json")
REPORT_2 = os.path.join(BASE_DIR, "laporan_tambah_remark (2).json")
SYNC_CONFIG_FILE = os.path.join(BASE_DIR, "sheet_sync_config.json")

P1_URL = "https://docs.google.com/spreadsheets/d/12FWSyYak-scXS5skHFUSQbVkjy6Oh8ofejusqnOb5nA/export?format=csv&gid=0"
P2_URL = "https://docs.google.com/spreadsheets/d/1ra-HZlqn1md8AIYaZz5VmOg-5Dxj83LYOHrmChvyz0s/export?format=csv&gid=293925488"

def clean_u(link):
    return str(link or "").split("?")[0].replace("/edit", "").rstrip("/")

def clean_nama(nama):
    return re.sub(r'[^a-z0-9]', '', str(nama or "").lower())

def load_json(filepath):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None
    return None

def fetch_sheet_csv(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0", "Cache-Control": "no-cache"})
    with urllib.request.urlopen(req, timeout=35) as resp:
        lines = resp.read().decode("utf-8").splitlines()
    return list(csv.DictReader(lines))

def main():
    print("=" * 65)
    print(" PUSH SELURUH HISTORI PENGERJAAN LOKAL KE GOOGLE SHEETS")
    print("=" * 65)

    cfg = load_json(SYNC_CONFIG_FILE) or {}
    operator = cfg.get("operator_name") or "Muhshi"
    urls_map = cfg.get("gas_webapp_urls", {})
    fallback_url = cfg.get("gas_webapp_url", "")

    p1_gas = urls_map.get("Prioritas 1: Data Detail Anomali (gid: 0)") or fallback_url
    p2_gas = urls_map.get("Prioritas 2: Data NTB (gid: 293925488)") or urls_map.get("Dtata_NTB") or urls_map.get("Data_NTB") or fallback_url

    # 1. Kumpulkan seluruh item selesai dari laporan JSON
    all_reports = []
    for fp in [REPORT_1, REPORT_2]:
        data = load_json(fp) or []
        for r in data:
            if r.get("status") in ["SUCCESS", "DILEWATI"]:
                all_reports.append(r)

    print(f"Total data selesai ditemukan di laporan lokal: {len(all_reports)}")

    # 2. Unduh CSV Google Sheets Priority 1 dan Priority 2
    print("\n-> Mengunduh data live Google Sheets...")
    p1_rows = fetch_sheet_csv(P1_URL)
    p2_rows = fetch_sheet_csv(P2_URL)
    print(f"   Priority 1 (Detail Anomali): {len(p1_rows)} baris")
    print(f"   Priority 2 (Data NTB)      : {len(p2_rows)} baris")

    # Buat peta pencarian baris live: (clean_link, clean_nama) -> (row_idx, status, iddesa, nama_usaha, raw_link)
    p1_map = {}
    for idx, r in enumerate(p1_rows, 2):
        lk = clean_u(r.get("link_fasih_sm") or r.get("link") or "")
        nm = clean_nama(r.get("nama_usaha") or "")
        st = str(r.get("Status Pengerjaan") or "").strip()
        p1_map[(lk, nm)] = (idx, st, r.get("iddesa", ""), r.get("nama_usaha", ""), r.get("link_fasih_sm", ""))

    p2_map = {}
    for idx, r in enumerate(p2_rows, 2):
        lk = clean_u(r.get("link_fasih_sm") or r.get("link") or "")
        nm = clean_nama(r.get("nama_usaha") or "")
        st = str(r.get("Status Pengerjaan") or "").strip()
        p2_map[(lk, nm)] = (idx, st, r.get("iddesa", ""), r.get("nama_usaha", ""), r.get("link_fasih_sm", ""))

    # 3. Cari baris yang BELUM SELESAI di Google Sheets
    items_to_push = []
    seen_keys = set()

    for r in all_reports:
        lk = clean_u(r.get("link", ""))
        nm = clean_nama(r.get("nama_usaha", ""))
        key = (lk, nm)
        if key in seen_keys or not lk:
            continue
        seen_keys.add(key)

        src_name = r.get("source_name", "").lower()
        is_p2 = ("ntb" in src_name) or ("prioritas 2" in src_name) or ("293925488" in src_name)

        target_info = None
        target_gid = "0"
        target_gas = p1_gas
        target_priority = "P1"

        if is_p2 or (lk, nm) in p2_map:
            if (lk, nm) in p2_map:
                target_info = p2_map[(lk, nm)]
                target_gid = "293925488"
                target_gas = p2_gas
                target_priority = "P2"

        if not target_info and ((lk, nm) in p1_map):
            target_info = p1_map[(lk, nm)]
            target_gid = "0"
            target_gas = p1_gas
            target_priority = "P1"

        if target_info:
            row_idx, cur_status, iddesa, orig_nama, orig_link = target_info
            if not any(k in cur_status.upper() for k in ["SELESAI", "DONE", "SUDAH", "BERHASIL"]):
                items_to_push.append({
                    "priority": target_priority,
                    "gas_url": target_gas,
                    "gid": target_gid,
                    "row_num": row_idx,
                    "link": orig_link or r.get("link", ""),
                    "nama_usaha": orig_nama or r.get("nama_usaha", ""),
                    "iddesa": iddesa or r.get("iddesa", ""),
                    "operator": operator
                })

    print(f"\nDitemukan {len(items_to_push)} baris yang sudah Anda kerjakan tapi BELUM berstatus SELESAI di Google Sheets.")
    if not items_to_push:
        print("[OK] Semua data pengerjaan lokal sudah berstatus SELESAI di Google Sheets!")
        return

    # Kelompokkan berdasarkan gas_url & gid untuk mencoba Batch Update
    grouped = {}
    for it in items_to_push:
        k = (it["gas_url"], it["gid"], it["priority"])
        if k not in grouped:
            grouped[k] = []
        grouped[k].append(it)

    print("\nMemulai pengiriman pembaruan status ke Google Sheets...")
    total_success = 0
    total_fail = 0

    for (gas_url, gid, p_name), items in grouped.items():
        print(f"\n--- Mengirim {len(items)} baris untuk {p_name} (gid: {gid}) ---")

        # Coba batch update terlebih dahulu (Didukung oleh GAS v3)
        batch_payload = {
            "action": "batch_update",
            "gid": gid,
            "operator": operator,
            "items": [
                {
                    "row_num": it["row_num"],
                    "link": it["link"],
                    "nama_usaha": it["nama_usaha"],
                    "status": "SELESAI",
                    "operator": operator
                }
                for it in items
            ]
        }

        try_batch = True
        try:
            req_data = json.dumps(batch_payload).encode("utf-8")
            req = urllib.request.Request(
                gas_url,
                data=req_data,
                headers={"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"}
            )
            t0 = time.time()
            with urllib.request.urlopen(req, timeout=60) as resp:
                resp_json = json.loads(resp.read().decode("utf-8"))
                if resp_json.get("status") == "success" and resp_json.get("updated_count", 0) > 0:
                    cnt = resp_json.get("updated_count")
                    print(f"  -> [BATCH UPDATE SUKSES] {cnt} baris berhasil diperbarui dan diwarnai hijau dalam {time.time()-t0:.2f}s!")
                    total_success += cnt
                    total_fail += resp_json.get("failed_count", 0)
                    try_batch = False
        except Exception as e:
            # Batch update mungkin belum terdeploy jika user belum update GAS v3
            print(f"  -> [Batch Update Belum Tersedia / Menunggu Deploy v3: {e}]")
            print("  -> Melakukan update per baris (sequential)...")
            try_batch = True

        if try_batch:
            for idx, item in enumerate(items, 1):
                row_num = item["row_num"]
                nama = item["nama_usaha"]
                link = item["link"]
                iddesa = item["iddesa"]

                params = {
                    "link": link,
                    "nama_usaha": nama,
                    "row_num": row_num,
                    "iddesa": iddesa,
                    "status": "SELESAI",
                    "operator": operator,
                    "gid": gid,
                    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
                }

                query_str = urllib.parse.urlencode(params)
                req_url = f"{gas_url}?{query_str}"
                print(f"  [{idx}/{len(items)}] Baris {row_num}: {nama} ({iddesa})...", end=" ", flush=True)

                req_success = False
                for retry in range(1, 4):
                    try:
                        r_req = urllib.request.Request(req_url, headers={"User-Agent": "Mozilla/5.0"})
                        t_item = time.time()
                        with urllib.request.urlopen(r_req, timeout=45) as resp:
                            res_json = json.loads(resp.read().decode("utf-8"))
                            if res_json.get("status") == "success":
                                print(f"[OK ({time.time()-t_item:.1f}s)]")
                                total_success += 1
                                req_success = True
                                break
                            else:
                                msg = res_json.get("message", "Gagal")
                                if "sibuk" in msg.lower() and retry < 3:
                                    time.sleep(2)
                                    continue
                                print(f"[{res_json.get('status')}] {msg}")
                                break
                    except Exception as e:
                        if retry < 3:
                            time.sleep(2)
                            continue
                        print(f"[ERROR: {e}]")
                        break

                if not req_success:
                    total_fail += 1

                time.sleep(0.3)

    print("\n" + "=" * 65)
    print(" SINKRONISASI PUSH SELESAI!")
    print(f" - Berhasil ditandai di Google Sheets : {total_success}")
    print(f" - Gagal / Not Found                : {total_fail}")
    print("=" * 65)

if __name__ == "__main__":
    main()
