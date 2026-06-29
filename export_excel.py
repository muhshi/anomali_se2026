import sqlite3
import pandas as pd
import json

def export_to_excel():
    db_file = "anomali_se2026.db"
    
    # 1. Connect to database
    conn = sqlite3.connect(db_file)
    
    # 2. Ambil data untuk tanggal '2026-06-26'
    query = """
    SELECT 
        tanggal_tarik, 
        kode_wilayah, 
        anomali_title, 
        assignment_id, 
        is_resolved,
        link_fasih,
        raw_data
    FROM kasus_anomali_mikro
    WHERE tanggal_tarik = '2026-06-26'
    ORDER BY kode_wilayah, anomali_title
    """
    
    df = pd.read_sql_query(query, conn)
    
    if df.empty:
        print("Tidak ada data untuk tanggal 2026-06-26.")
        return
    
    # 3. Ekstrak data tambahan dari raw_data json
    def extract_level(row, level_key):
        try:
            data = json.loads(row['raw_data'])
            return data.get(level_key, "")
        except:
            return ""
            
    df['Kecamatan'] = df.apply(lambda row: extract_level(row, 'level_3_code'), axis=1)
    df['Desa'] = df.apply(lambda row: extract_level(row, 'level_4_code'), axis=1)
    df['SLS'] = df.apply(lambda row: extract_level(row, 'level_6_code'), axis=1)
    
    # Status Teks
    df['Status'] = df['is_resolved'].apply(lambda x: 'Sudah Ditindaklanjuti' if x else 'Belum Ditindaklanjuti')
    
    # 4. Atur urutan dan nama kolom
    columns_order = [
        'tanggal_tarik', 'Kecamatan', 'Desa', 'SLS', 'kode_wilayah', 
        'anomali_title', 'Status', 'assignment_id', 'link_fasih'
    ]
    df_final = df[columns_order]
    
    # Rename kolom
    df_final.columns = [
        'Tanggal Tarik', 'Kecamatan', 'Desa', 'SLS', 'Kode Wilayah', 
        'Jenis Anomali', 'Status', 'ID Assignment', 'Link FASIH'
    ]
    
    # 5. Export ke Excel
    output_file = "Laporan_Anomali_2026-06-26.xlsx"
    df_final.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Berhasil membuat file laporan: {output_file} (Total: {len(df_final)} baris)")

if __name__ == "__main__":
    export_to_excel()
