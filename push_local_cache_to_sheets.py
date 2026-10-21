#!/usr/bin/env python3
"""
Script Utilitas: Sinkronkan Histori Pengerjaan Lokal ke Google Sheets
Digunakan untuk menandai baris yang sudah pernah dikerjakan sebelumnya
agar di Google Spreadsheet langsung tertulis SELESAI dan berwarna HIJAU.
"""

import os
import sys
import json
import time
import re
import urllib.request
import urllib.parse

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_JSON = os.path.join(BASE_DIR, "laporan_tambah_remark.json")
CACHE_FILE = os.path.join(BASE_DIR, "processed_tambah_remark.json")
SYNC_CONFIG_FILE = os.path.join(BASE_DIR, "sheet_sync_config.json")

def load_json(filepath):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None
    return None

def main():
    print("=" * 65)
    print(" SINKRONISASI HISTORI PENGERJAAN LOKAL KE GOOGLE SHEETS")
    print("=" * 65)

    cfg = load_json(SYNC_CONFIG_FILE) or {}
    operator = cfg.get("operator_name") or "Muhshi"
    urls_map = cfg.get("gas_webapp_urls", {})
    fallback_url = cfg.get("gas_webapp_url", "")

    report_data = load_json(REPORT_JSON) or []
    # Ambil data yang berhasil
    completed_items = [r for r in report_data if r.get("status") in ["SUCCESS", "DILEWATI"]]

    if not completed_items:
        print("\n[INFO] Belum ada data berstatus 'SUCCESS' di laporan_tambah_remark.json.")
        return

    print(f"\nDitemukan {len(completed_items)} data yang sudah sukses dikerjakan di komputer ini.")
    print(f"Operator : {operator}")

    success_sync = 0
    fail_sync = 0

    for idx, item in enumerate(completed_items, 1):
        src = item.get("source_name", "")
        nama = item.get("nama_usaha", "")
        link = item.get("link", "")
        iddesa = item.get("iddesa", "")

        # Pilih URL Web App sesuai sumber (Prioritas 1 vs Prioritas 2)
        url = ""
        for k, u in urls_map.items():
            if not u:
                continue
            k_low = k.lower()
            src_low = src.lower()
            if (k_low in src_low) or (src_low in k_low) or \
               ("prioritas 1" in src_low and ("detail" in k_low or "anomali" in k_low or "1" in k_low)) or \
               ("prioritas 2" in src_low and ("ntb" in k_low or "2" in k_low)) or \
               ("ntb" in src_low and "ntb" in k_low) or \
               ("anomali" in src_low and ("anomali" in k_low or "detail" in k_low)):
                url = u.strip()
                break
        if not url:
            url = fallback_url.strip()

        if not url:
            print(f"[{idx}/{len(completed_items)}] [SKIP - URL KOSONG] {src} : {nama}")
            fail_sync += 1
            continue

        # Ekstrak gid
        gid = ""
        gid_match = re.search(r"gid[:\s=]+(\d+)", src)
        if gid_match:
            gid = gid_match.group(1)

        params = {
            "link": link,
            "nama_usaha": nama,
            "iddesa": iddesa,
            "status": "SELESAI",
            "operator": operator,
            "gid": gid,
            "timestamp": item.get("timestamp") or time.strftime("%Y-%m-%d %H:%M:%S")
        }

        query_str = urllib.parse.urlencode(params)
        req_url = f"{url}?{query_str}"
        print(f"[{idx}/{len(completed_items)}] Mengirim: {nama} ({iddesa})...", end=" ", flush=True)

        try:
            req = urllib.request.Request(
                req_url,
                headers={"User-Agent": "Mozilla/5.0"}
            )
            with urllib.request.urlopen(req, timeout=20) as resp:
                resp_data = resp.read().decode("utf-8")
                res_json = json.loads(resp_data)
                if res_json.get("status") == "success":
                    sheet_tab = res_json.get("sheet", "")
                    row_idx = res_json.get("row", "")
                    print(f"[OK] Baris {row_idx} ({sheet_tab}) -> SELESAI & HIJAU")
                    success_sync += 1
                else:
                    msg = res_json.get("message", "Gagal")
                    print(f"[{res_json.get('status', 'FAIL')}] {msg}")
                    fail_sync += 1
        except Exception as e:
            print(f"[ERROR] {e}")
            fail_sync += 1

        time.sleep(0.3)

    print("\n" + "=" * 65)
    print(" SINKRONISASI SELESAI!")
    print(f" - Berhasil ditandai di Google Sheets : {success_sync}")
    print(f" - Gagal / Not Found                : {fail_sync}")
    print("=" * 65)

if __name__ == "__main__":
    main()
