# Eksplorasi Data Penjualan 10.000 Transaksi

Praktikum eksplorasi data: analisis deskriptif, segmentasi pelanggan RFM + K-Means,
peramalan SARIMA, dan dashboard interaktif.

Pemilik: **Putri Okta Fitria A** — NRP **3325600084**

## Isi Folder

| Folder | Isi |
|---|---|
| `01-dataset/` | `soal.txt` dan dataset `Dataset_Transaksi_Penjualan_10000_*.xlsx` (sheet `Data_Transaksi` 10.000 × 18, `Master_Produk` 25 × 7) |
| `02-analisis/` | `analisis.py` (pipeline), `kualitas.json` (hasil pemeriksaan data), `ringkasan.json`, `clustering.json`, `forecast.json`, `grafik/` (4 PNG) |
| `03-dashboard/` | `index.html` (dashboard mandiri, ±2,5 MB), `template.html`, `buat-dashboard.py`, `indonesia-provinces.geojson` |
| `04-laporan/` | `laporan-praktikum.docx` + `.pdf`, `screenshot/` (13 tangkapan layar), `langkah/` (halaman log terminal + `log/`), `sampul.png`, `buat-laporan.py`, `buat-langkah.py` |

## Cara Membuka

- Laporan: `04-laporan/laporan-praktikum.pdf` (25 halaman) atau versi `.docx`.
- Dashboard: buka `03-dashboard/index.html` di peramban — tanpa server, tanpa internet.
  Tiga tab: Dashboard Penjualan, Data Transaksi (filter + tambah), Master Produk.

## Cara Menjalankan Ulang

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt

.venv/bin/python 02-analisis/analisis.py
.venv/bin/python 03-dashboard/buat-dashboard.py
.venv/bin/python 04-laporan/buat-langkah.py
.venv/bin/python 04-laporan/buat-laporan.py

# PDF (butuh LibreOffice)
cd 04-laporan && soffice --headless --convert-to pdf laporan-praktikum.docx --outdir .
```

Skrip memakai path relatif terhadap folder proyek, jadi bisa dijalankan dari mana saja
selama urutan di atas diikuti. `04-laporan/screenshot/` sudah tersedia di repo; kalau
ingin menangkap ulang gambar langkah, buka `04-laporan/langkah/*.html` di peramban lalu
salin tangkapan layarnya ke nama berkas yang sama.

## Ringkasan Hasil

- Total penjualan Rp 4.311.513.600 (bersih setelah pembatalan Rp 4.174.773.550), profit Rp 1.757.346.600, margin 40,76%.
- Jakarta 25,74% penjualan; Elektronik 28,77%; Shopee 34,72%.
- RFM + K-Means k=4 (silhouette 0,2824; DBI 1,0774). Segmen "Pelanggan Bernilai Tinggi" = 26,9% pelanggan, 40,63% penjualan.
- Peramalan terbaik SARIMA(0,1,1)(0,1,0,12): MAPE uji 3,02%, backtest 4,13%.
  Ramalan Nov 2026 Rp 239,5 jt, Des 2026 Rp 274,8 jt, Jan 2027 Rp 214,9 jt.
