# Sistem Fuzzy Logic untuk Pemilihan Restoran Terbaik di kota Bandung

## Persyaratan wajib install
- Python 3.x
- Library openpyxl (untuk operasi Excel) diclaimer ini pakai library karena hanya untuk membaca operasi exel!
- 

## Instalasi
1. Pastikan Python terinstal di sistem Anda
2. Install library yang diperlukan:
   pip install openpyxl

## Cara Menjalankan
1. Letakkan file data restoran.xlsx di folder /data
2. Jalankan program utama:
   python main.py

## Struktur File Input
File restoran.xlsx harus memiliki format:
- Kolom 1: ID Pelanggan (integer)
- Kolom 2: Servis/Pelayanan (1-100)
- Kolom 3: Harga (25000-55000)

## Output
Program akan menghasilkan:
1. File peringkat.xlsx di folder /data berisi 10 restoran terbaik
2. Tampilan di console dengan 10 restoran terbaik

## Kontak
Nama: [Ahmad Junaidi]
Email: [ahmdjunaidibeds@gmail.com]