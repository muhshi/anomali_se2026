def hitung_statistik(deret_angka):
    # Validasi jika deret kosong
    if not deret_angka:
        return None
    
    nilai_min = min(deret_angka)
    nilai_max = max(deret_angka)
    
    # Rata-rata adalah total jumlah dibagi banyaknya elemen
    rata_rata = sum(deret_angka) / len(deret_angka)
    
    return nilai_min, nilai_max, rata_rata

# --- Contoh Penggunaan ---
input_angka = [15, 2, 8, 24, 10, 5]
minimum, maximum, rata2 = hitung_statistik(input_angka)

print(f"Deret Angka: {input_angka}")
print(f"Nilai Minimum: {minimum}")
print(f"Nilai Maksimum: {maximum}")
print(f"Rata-rata: {rata2}")