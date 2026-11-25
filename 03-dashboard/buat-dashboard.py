"""Bangun 03-dashboard/index.html dari dataset Excel + hasil analisis JSON.

Jalankan dari root proyek:  .venv/bin/python 03-dashboard/buat-dashboard.py
Seluruh data ditanam inline supaya dashboard tetap jalan saat dibuka lewat file://
tanpa server dan tanpa akses internet.
"""
import json
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent.parent
DASH = BASE / "03-dashboard"
XLSX = BASE / "01-dataset" / "Dataset_Transaksi_Penjualan_10000_1795569666632_1795589876097.xlsx"
if not XLSX.exists():
    XLSX = BASE / "Dataset_Transaksi_Penjualan_10000_1795569666632_1795589876097.xlsx"

KOTA_PROVINSI = {
    "Jakarta": "Jakarta Raya", "Surabaya": "Jawa Timur", "Bandung": "Jawa Barat",
    "Semarang": "Jawa Tengah", "Medan": "Sumatera Utara", "Makassar": "Sulawesi Selatan",
    "Yogyakarta": "Yogyakarta", "Denpasar": "Bali", "Malang": "Jawa Timur",
    "Balikpapan": "Kalimantan Timur",
}
WARNA = ["#4f46e5", "#0ea5e9", "#14b8a6", "#f59e0b"]

trx = pd.read_excel(XLSX, "Data_Transaksi")
prd = pd.read_excel(XLSX, "Master_Produk")
trx["Tanggal"] = pd.to_datetime(trx["Tanggal"]).dt.strftime("%Y-%m-%d")

kolom = ["ID_Transaksi", "Tanggal", "Nama_Pelanggan", "Segmen_Pelanggan", "Kota",
         "Channel_Penjualan", "Kategori_Produk", "Nama_Produk", "Qty", "Harga_Satuan",
         "Diskon_Pct", "Subtotal", "Nilai_Diskon", "Total_Penjualan", "HPP_Satuan",
         "Profit", "Metode_Pembayaran", "Status_Transaksi"]
rows = [[(float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else v)
         for v in r] for r in trx[kolom].itertuples(index=False)]

produk = [[p["Kode_Produk"], p["Kategori_Produk"], p["Nama_Produk"],
           float(p["Harga_Satuan (Rp)"]), float(p["HPP_Satuan (Rp)"]),
           float(p["Margin_Unit (Rp)"]), float(p["Margin (%)"])]
          for _, p in prd.iterrows()]

cl = json.loads((BASE / "02-analisis/clustering.json").read_text())
fc = json.loads((BASE / "02-analisis/forecast.json").read_text())
geo = json.loads((DASH / "indonesia-provinces.geojson").read_text())

cluster = {
    "nama": [p["segmen"] for p in cl["profil"]],
    "colors": WARNA[: len(cl["profil"])],
    "profil": cl["profil"],
    "titik": [[r["Recency"], r["Monetary"], r["Cluster"]] for r in cl["sebaran"]],
    "k": cl["k_terpilih"],
    "silhouette": cl["silhouette"],
    "dbi": cl["dbi"],
}

data = {
    "transaksi": rows,
    "produk": produk,
    "geojson": geo,
    "kota_provinsi": KOTA_PROVINSI,
    "cluster": cluster,
    "forecast": {
        "metode_terpilih": fc["metode_terpilih"],
        "mape_uji": fc["mape_uji"],
        "mape_backtest": fc["mape_backtest"],
        "backtest": fc["backtest"],
        "kandidat": fc["kandidat"],
        "aktual": fc["aktual"],
        "label_bulan": fc["label_bulan"],
        "prediksi": fc["prediksi"],
        "label_prediksi": fc["label_prediksi"],
    },
}

tpl = (DASH / "template.html").read_text()
payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
html = tpl.replace("/*__DATA__*/", payload)
out = DASH / "index.html"
out.write_text(html)
print(f"index.html ditulis: {out.stat().st_size/1e6:.2f} MB")
print(f"  transaksi : {len(rows)} baris")
print(f"  produk    : {len(produk)} SKU")
print(f"  cluster   : k={cluster['k']} sil={cluster['silhouette']} titik={len(cluster['titik'])}")
print(f"  forecast  : {fc['metode_terpilih']} MAPE={fc['mape_uji']}%")
