#!/bin/bash
echo "=================================================="
echo "  Menjalankan Penarik Data Anomali SE2026"
echo "=================================================="

if [ ! -f config.json ]; then
    echo "[ERROR] config.json tidak ditemukan!"
    echo "Silakan copy config.example.json menjadi config.json dan isi kredensial database Anda."
    exit 1
fi

echo "Memulai web dashboard di latar belakang..."
python web_anomali.py &
WEB_PID=$!

echo "Menunggu 3 detik..."
sleep 3

echo "Memulai penarik data anomali..."
python tarik_anomali.py

echo ""
echo "Menghentikan web dashboard..."
kill $WEB_PID
echo "Selesai."
