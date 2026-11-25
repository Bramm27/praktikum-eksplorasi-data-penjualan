"""Bangun halaman HTML bertema terminal dari log pada langkah/log/*.txt.

Halaman ini yang ditangkap layar menjadi gambar langkah di laporan.
Jalankan dari root proyek:  python 04-laporan/buat-langkah.py
"""
from pathlib import Path

LAP = Path(__file__).resolve().parent
LOG = LAP / "langkah" / "log"
OUT = LAP / "langkah"
OUT.mkdir(parents=True, exist_ok=True)

HALAMAN = [
    ("langkah-01-lingkungan.html", "Langkah 1 — Menyiapkan lingkungan Python",
     "1-instal.txt", "Virtualenv proyek dibuat, lalu paket yang belum tersedia dipasang."),
    ("langkah-02-dataset.html", "Langkah 2 — Memeriksa struktur dataset",
     "2-dataset.txt", "Dua sheet: Data_Transaksi (10.000 baris, 18 kolom) dan Master_Produk (25 SKU)."),
    ("langkah-03-analisis.html", "Langkah 3 — Menjalankan skrip analisis",
     "3-analisis.txt", "Statistik deskriptif, RFM + K-Means, dan peramalan SARIMA dalam satu skrip."),
    ("langkah-04-dashboard.html", "Langkah 4 — Membangun dashboard",
     "4-dashboard.txt", "Seluruh data ditanam inline supaya dashboard bisa dibuka tanpa server."),
]

TPL = """<!DOCTYPE html>
<html lang="id"><head><meta charset="utf-8">
<style>
body{{margin:0;background:#eef1f6;font-family:"Inter","Segoe UI",system-ui,sans-serif;padding:26px}}
.box{{max-width:1180px;margin:0 auto;background:#fff;border-radius:16px;padding:22px 24px;
  box-shadow:0 10px 30px -18px rgba(15,23,42,.35)}}
h2{{margin:0 0 4px;font-size:16px;color:#0f172a}}
p{{margin:0 0 14px;font-size:12.5px;color:#64748b}}
.term{{background:#0b1120;border-radius:12px;padding:16px 18px;overflow:hidden;
  box-shadow:inset 0 0 0 1px #1e293b}}
.bar{{display:flex;gap:7px;margin-bottom:12px}}
.bar i{{width:10px;height:10px;border-radius:50%;display:block}}
pre{{margin:0;color:#e2e8f0;font-family:"JetBrains Mono","Fira Code",Consolas,"DejaVu Sans Mono",monospace;
  font-size:12.2px;line-height:1.62;white-space:pre-wrap;word-break:break-word}}
.cmd{{color:#7dd3fc;font-weight:700}}
.err{{color:#fca5a5}}
.meta{{font-size:11px;color:#94a3b8;margin-top:10px}}
</style></head><body>
<div class="box">
  <h2>{judul}</h2>
  <p>{ket}</p>
  <div class="term">
    <div class="bar"><i style="background:#ef4444"></i><i style="background:#f59e0b"></i><i style="background:#22c55e"></i></div>
    <pre>{isi}</pre>
  </div>
  <div class="meta">Keluaran terminal dari sesi praktikum.</div>
</div>
</body></html>
"""

for nama, judul, sumber, ket in HALAMAN:
    teks = (LOG / sumber).read_text()
    baris = []
    for ln in teks.rstrip("\n").split("\n"):
        if ln.startswith("$ "):
            baris.append(f'<span class="cmd">{ln}</span>')
        elif ln.startswith("Traceback") or ln.startswith("  File") or ln.startswith("StopIteration"):
            baris.append(f'<span class="err">{ln}</span>')
        else:
            baris.append(ln.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    html = TPL.format(judul=judul, ket=ket, isi="\n".join(baris))
    (OUT / nama).write_text(html)
    print("tulis", nama, len(html), "byte")
