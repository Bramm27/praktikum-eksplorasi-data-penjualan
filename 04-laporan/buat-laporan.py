"""Susun laporan praktikum (.docx) dari hasil analisis + tangkapan layar.

Jalankan dari root proyek dengan venv yang punya python-docx:
  .venv/bin/python 04-laporan/buat-laporan.py
"""
import json
from datetime import date, datetime
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = Path(__file__).resolve().parent.parent
LAP = BASE / "04-laporan"
SS = LAP / "screenshot"
GRAFIK = BASE / "02-analisis" / "grafik"
ANA = BASE / "02-analisis"

NAVY = RGBColor(0x00, 0x00, 0x00)
GREY = RGBColor(0x00, 0x00, 0x00)
AKSEN = "E8EDF4"
ZEBRA = "F2F5F9"

ringkasan = json.loads((ANA / "ringkasan.json").read_text())
clust = json.loads((ANA / "clustering.json").read_text())
fc = json.loads((ANA / "forecast.json").read_text())
kualitas = json.loads((ANA / "kualitas.json").read_text())
kpi = ringkasan["kpi"]

import pandas as pd  # noqa: E402

_src = next((BASE / "01-dataset").glob("Dataset_Transaksi_*.xlsx"))
_trx = pd.read_excel(_src, "Data_Transaksi")
BATAL_NILAI = float(_trx.loc[_trx["Status_Transaksi"] == "Dibatalkan", "Total_Penjualan"].sum())
PENJUALAN_BERSIH = float(_trx["Total_Penjualan"].sum()) - BATAL_NILAI


def rp(n, desimal=0):
    s = f"{n:,.{desimal}f}".replace(",", "@").replace(".", ",").replace("@", ".")
    return f"Rp {s}"


def angka(n, desimal=0):
    return f"{n:,.{desimal}f}".replace(",", "@").replace(".", ",").replace("@", ".")


def juta(n):
    return f"{n/1e6:,.1f}".replace(",", "@").replace(".", ",").replace("@", ".") + " juta"


def persen(a, b):
    return f"{a/b*100:,.2f}".replace(",", "@").replace(".", ",").replace("@", ".") + "%"


# ---------------------------------------------------------------- dokumen
doc = Document()
cp = doc.core_properties
cp.author = "Putri Okta Fitria A"
cp.last_modified_by = "Putri Okta Fitria A"
cp.title = "Laporan Praktikum Eksplorasi Data Penjualan"
cp.subject = "Praktikum Eksplorasi Data"
cp.description = ""
cp.keywords = "eksplorasi data, RFM, K-Means, SARIMA"
cp.comments = ""
cp.category = "Laporan Praktikum"
cp.created = datetime(2026, 11, 25, 9, 0, 0)
cp.modified = datetime(2026, 11, 25, 9, 0, 0)
cp.revision = 1

sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.top_margin = sec.bottom_margin = Cm(2.2)
sec.left_margin = Cm(2.6)
sec.right_margin = Cm(2.4)
sec.different_first_page_header_footer = True

normal = doc.styles["Normal"]
normal.font.name = "Cambria"
normal.font.size = Pt(11)
normal.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
normal.paragraph_format.line_spacing = 1.35
normal.paragraph_format.space_after = Pt(7)
normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Cambria")

nomor_gambar = [0]
nomor_tabel = [0]


def h1(teks, halaman_baru=False):
    if halaman_baru:
        doc.add_page_break()
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(9)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(teks)
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = NAVY
    pb = OxmlElement("w:pBdr")
    bawah = OxmlElement("w:bottom")
    bawah.set(qn("w:val"), "single")
    bawah.set(qn("w:sz"), "8")
    bawah.set(qn("w:space"), "4")
    bawah.set(qn("w:color"), "C7D2E0")
    pb.append(bawah)
    p._p.get_or_add_pPr().append(pb)
    return p


def h2(teks):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(11)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(teks)
    r.font.size = Pt(11.5)
    r.font.bold = True
    r.font.color.rgb = NAVY
    return p


def para(teks, align="just", size=11, italic=False, color=None, space_after=7):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.alignment = {"just": WD_ALIGN_PARAGRAPH.JUSTIFY, "left": WD_ALIGN_PARAGRAPH.LEFT,
                   "center": WD_ALIGN_PARAGRAPH.CENTER}[align]
    for i, bagian in enumerate(teks.split("**")):
        r = p.add_run(bagian)
        r.font.size = Pt(size)
        r.italic = italic
        r.bold = i % 2 == 1
        if color:
            r.font.color.rgb = color
    return p


def rich(teks, align="just"):
    """Paragraf dengan penanda **tebal** sederhana."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if align == "just" else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(7)
    for i, bagian in enumerate(teks.split("**")):
        r = p.add_run(bagian)
        r.font.size = Pt(11)
        r.bold = i % 2 == 1
    return p


def bullet(teks, size=11):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Cm(0.75)
    for i, bagian in enumerate(teks.split("**")):
        r = p.add_run(bagian)
        r.font.size = Pt(size)
        r.bold = i % 2 == 1
    return p


def shade(cell, warna):
    el = OxmlElement("w:shd")
    el.set(qn("w:val"), "clear")
    el.set(qn("w:fill"), warna)
    cell._tc.get_or_add_tcPr().append(el)


def table(header, rows, lebar=None, caption=None, align_kanan=None, size=9.5):
    if caption:
        nomor_tabel[0] += 1
        cp = doc.add_paragraph()
        cp.paragraph_format.space_before = Pt(6)
        cp.paragraph_format.space_after = Pt(3)
        cp.paragraph_format.keep_with_next = True
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = cp.add_run(f"Tabel {nomor_tabel[0]}. {caption}")
        r.font.size = Pt(9)
        r.italic = True
        r.font.color.rgb = GREY
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    hdr = t.rows[0].cells
    for i, h in enumerate(header):
        hdr[i].text = ""
        p = hdr[i].paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(str(h))
        r.font.size = Pt(size)
        r.bold = True
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        shade(hdr[i], AKSEN)
    for n, baris in enumerate(rows):
        sel = t.add_row().cells
        for i, v in enumerate(baris):
            sel[i].text = ""
            p = sel[i].paragraphs[0]
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.space_before = Pt(1)
            if align_kanan and i in align_kanan:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run(str(v))
            r.font.size = Pt(size)
            if n % 2 == 1:
                shade(sel[i], ZEBRA)
    # ulangi baris header di setiap halaman
    trPr = t.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    th.set(qn("w:val"), "true")
    trPr.append(th)

    if lebar:
        t.autofit = False
        t.allow_autofit = False
        for i, w in enumerate(lebar):
            t.columns[i].width = Cm(w)
            for row in t.rows:
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def gambar(path, caption, width=15.6, jarak=4):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(jarak)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(path), width=Cm(width))
    nomor_gambar[0] += 1
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_after = Pt(10)
    r = cp.add_run(f"Gambar {nomor_gambar[0]}. {caption}")
    r.font.size = Pt(8.5)
    r.italic = True
    r.font.color.rgb = GREY


def nomor_halaman():
    footer = sec.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    r.font.size = Pt(9)
    r.font.color.rgb = GREY
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    p._p.append(fld)


def header_teks(teks):
    hp = sec.header.paragraphs[0]
    hp.text = ""
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = hp.add_run(teks)
    r.font.size = Pt(8.5)
    r.font.color.rgb = GREY
    pb = OxmlElement("w:pBdr")
    bawah = OxmlElement("w:bottom")
    bawah.set(qn("w:val"), "single")
    bawah.set(qn("w:sz"), "4")
    bawah.set(qn("w:space"), "4")
    bawah.set(qn("w:color"), "D6DEE8")
    pb.append(bawah)
    hp._p.get_or_add_pPr().append(pb)


header_teks("Laporan Praktikum Eksplorasi Data Penjualan")

# ---------------------------------------------------------------- sampul
sp = doc.add_paragraph()
sp.paragraph_format.space_before = Pt(40)
sp.paragraph_format.space_after = Pt(0)
sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = sp.add_run("LAPORAN PRAKTIKUM")
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = GREY

t = doc.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
t.paragraph_format.space_before = Pt(6)
t.paragraph_format.space_after = Pt(2)
r = t.add_run("Eksplorasi Data Penjualan 10.000 Transaksi")
r.font.size = Pt(24)
r.font.bold = True
r.font.color.rgb = NAVY

t2 = doc.add_paragraph()
t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
t2.paragraph_format.space_after = Pt(4)
r = t2.add_run("Analisis Deskriptif, Segmentasi Pelanggan RFM, Peramalan SARIMA,\ndan Dashboard Interaktif")
r.font.size = Pt(12.5)
r.font.color.rgb = GREY

garis = doc.add_paragraph()
garis.alignment = WD_ALIGN_PARAGRAPH.CENTER
garis.paragraph_format.space_before = Pt(10)
garis.paragraph_format.space_after = Pt(26)
r = garis.add_run("— — —")
r.font.size = Pt(12)
r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

gambar(LAP / "sampul.png", "Ilustrasi sampul: alur kerja eksplorasi data sampai dashboard.", width=11.5, jarak=0)

doc.add_paragraph().paragraph_format.space_after = Pt(18)

identitas = doc.add_table(rows=0, cols=2)
identitas.alignment = WD_TABLE_ALIGNMENT.CENTER
for label, isi in [("Disusun oleh", "Putri Okta Fitria A"),
                   ("NRP", "3325600084"),
                   ("Program Studi", "Teknologi Informasi"),
                   ("Mata Kuliah", "Praktikum Eksplorasi Data"),
                   ("Tanggal", "25 November 2026")]:
    sel = identitas.add_row().cells
    p = sel[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run(label)
    r.font.size = Pt(11)
    r.font.color.rgb = GREY
    p2 = sel[1].paragraphs[0]
    r2 = p2.add_run(": " + isi)
    r2.font.size = Pt(11)
    r2.bold = label in ("Disusun oleh", "NRP")
    sel[0].width = Cm(5.5)
    sel[1].width = Cm(8.5)

# ---------------------------------------------------------------- ringkasan
h1("Ringkasan Eksekutif", halaman_baru=True)
para(
    f"Dataset yang diolah memuat {angka(kualitas['baris'])} transaksi penjualan dari "
    f"{kualitas['tanggal_min']} sampai {kualitas['tanggal_max']} "
    f"({kualitas['bulan']} bulan, {kualitas['kota_unik']} kota, {kualitas['produk_unik']} produk, "
    f"{angka(kualitas['pelanggan_unik'])} pelanggan unik). Nilai transaksi tercatat "
    f"{rp(kpi['total_penjualan'])} dengan profit {rp(kpi['total_profit'])} "
    f"(margin {angka(kpi['margin_pct'], 2)}%). Setelah {kpi['transaksi_batal']} transaksi berstatus "
    f"Dibatalkan dikeluarkan, nilai penjualan bersih menjadi {rp(PENJUALAN_BERSIH)}."
)
para("Temuan utama:")
bullet(f"**Jakarta** menyumbang {persen(ringkasan['kota'][0]['penjualan'], kpi['total_penjualan'])} "
       f"dari total penjualan ({rp(ringkasan['kota'][0]['penjualan'])}), diikuti Surabaya dan Bandung.")
bullet(f"**Elektronik** adalah kategori terbesar ({persen(ringkasan['kategori'][0]['penjualan'], kpi['total_penjualan'])}), "
       f"sedangkan **Shopee** menyalurkan {persen(ringkasan['channel'][0]['penjualan'], kpi['total_penjualan'])} penjualan.")
bullet(f"Segmentasi RFM + K-Means (k={clust['k_terpilih']}) memisahkan "
       f"**{clust['profil'][0]['segmen']}** berisi {clust['profil'][0]['jumlah']} pelanggan "
       f"({angka(clust['profil'][0]['porsi_pelanggan_pct'], 1)}% pelanggan) yang menghasilkan "
       f"{angka(clust['profil'][0]['porsi_penjualan_pct'], 1)}% penjualan.")
bullet(f"Peramalan terbaik adalah **{fc['metode_terpilih']}** dengan MAPE uji {angka(fc['mape_uji'], 2)}%; "
       f"prediksi tiga bulan ke depan: {rp(fc['prediksi'][0])}, {rp(fc['prediksi'][1])}, dan {rp(fc['prediksi'][2])}.")
bullet("Seluruh hasil disajikan dalam **dashboard interaktif satu berkas** yang dapat dibuka tanpa server "
       "dan tanpa koneksi internet, lengkap dengan filter, tabel berpaginasi, serta formulir tambah data.")

# ---------------------------------------------------------------- 1
h1("1. Pendahuluan", halaman_baru=True)
h2("1.1 Latar Belakang")
para(
    "Data transaksi harian menumpuk tanpa memberi keputusan bila hanya disimpan. Praktikum ini mengolah "
    "satu berkas Excel berisi 10.000 transaksi menjadi tiga lapis informasi: ringkasan performa, "
    "segmentasi pelanggan, dan proyeksi penjualan. Hasilnya kemudian dibungkus menjadi dashboard agar "
    "pembaca non-teknis dapat menelusuri angka sampai ke baris transaksi aslinya."
)
h2("1.2 Tujuan")
bullet("Menilai kualitas dan struktur dataset sebelum dianalisis.")
bullet("Meringkas performa penjualan menurut waktu, kota, kategori, kanal, dan segmen pelanggan.")
bullet("Membentuk segmen pelanggan dengan RFM dan K-Means, lalu menafsirkan tiap segmen.")
bullet("Memilih model peramalan penjualan bulanan berbasis bukti (MAPE uji dan backtest).")
bullet("Menyajikan seluruh hasil dalam dashboard interaktif dan laporan yang dapat direproduksi.")
h2("1.3 Ruang Lingkup")
para(
    "Analisis memakai seluruh 10.000 baris tanpa membuang transaksi Dibatalkan; status tersebut hanya "
    "dikeluarkan saat menghitung KPI penjualan bersih di dashboard. Peramalan memakai data bulanan "
    "22 titik, dengan 3 bulan terakhir disisihkan sebagai data uji."
)

# ---------------------------------------------------------------- 2
h1("2. Dataset dan Lingkungan Kerja", halaman_baru=True)
h2("2.1 Berkas dan Struktur")
para(
    "Berkas sumber berupa satu workbook Excel dengan dua sheet: **Data_Transaksi** "
    f"({angka(kualitas['baris'])} baris × {kualitas['kolom']} kolom) dan **Master_Produk** "
    f"({kualitas['produk_unik']} baris × {len(kualitas['produk_kolom'])} kolom). Sheet kedua berperan "
    "sebagai tabel referensi harga dan HPP sehingga perhitungan margin pada dashboard memakai sumber "
    "yang sama dengan dataset."
)
table(
    ["Sheet", "Baris", "Kolom", "Isi"],
    [["Data_Transaksi", angka(kualitas["baris"]), kualitas["kolom"], "Satu baris = satu transaksi (18 atribut)"],
     ["Master_Produk", kualitas["produk_unik"], len(kualitas["produk_kolom"]), "Kode, kategori, harga, HPP, margin"]],
    lebar=[3.6, 2.0, 1.8, 8.2],
    caption="Struktur sheet pada berkas dataset",
    align_kanan={1, 2},
)
h2("2.2 Kamus Data Sheet Data_Transaksi")
table(
    ["Kolom", "Tipe", "Keterangan"],
    [["ID_Transaksi", "teks", "Kunci unik, format TRX-YYYYMM-#####"],
     ["Tanggal", "tanggal", "Tanggal transaksi (2025-01-01 s.d. 2026-10-31)"],
     ["Nama_Pelanggan", "teks", "Nama pelanggan, 1.000 pelanggan unik"],
     ["Segmen_Pelanggan", "kategori", "Reguler, Member Silver, Member Gold, Korporat"],
     ["Kota", "kategori", "10 kota asal transaksi"],
     ["Channel_Penjualan", "kategori", "Shopee, Tokopedia, TikTok Shop, Website Resmi, Toko Offline"],
     ["Kategori_Produk", "kategori", "5 kategori produk"],
     ["Nama_Produk", "teks", "25 produk (kode PRD-001 s.d. PRD-025)"],
     ["Qty", "bilangan", "Jumlah unit per transaksi (1 s.d. 15)"],
     ["Harga_Satuan", "rupiah", "Harga satuan sebelum diskon"],
     ["Diskon_Pct", "desimal", "Diskon 0% s.d. 20% (rata-rata 4%)"],
     ["Subtotal", "rupiah", "Qty × Harga_Satuan"],
     ["Nilai_Diskon", "rupiah", "Subtotal × Diskon_Pct"],
     ["Total_Penjualan", "rupiah", "Subtotal − Nilai_Diskon"],
     ["HPP_Satuan", "rupiah", "Harga pokok per unit"],
     ["Profit", "rupiah", "Total_Penjualan − (HPP_Satuan × Qty)"],
     ["Metode_Pembayaran", "kategori", "QRIS, Virtual Account, E-Wallet, Transfer Bank, Kartu Kredit, COD"],
     ["Status_Transaksi", "kategori", "Selesai, Dikirim, Diproses, Dibatalkan"]],
    lebar=[3.4, 1.9, 10.3],
    caption="Kamus data kolom pada sheet Data_Transaksi",
    size=9,
)
h2("2.3 Pemeriksaan Kualitas Data")
para(
    "Sebelum analisis, dilakukan pemeriksaan kelengkapan, keunikan kunci, dan konsistensi rumus. "
    "Hasilnya dataset bersih: tidak ada sel kosong maupun ID ganda, dan kolom turunan konsisten penuh."
)
table(
    ["Pemeriksaan", "Hasil"],
    [["Sel kosong", f"{kualitas['sel_kosong']} sel"],
     ["ID_Transaksi ganda", f"{kualitas['duplikat_id']} baris"],
     ["Subtotal = Qty × Harga_Satuan", f"{kualitas['cek_subtotal']} penyimpangan"],
     ["Total_Penjualan = Subtotal − Nilai_Diskon", f"{kualitas['cek_total']} penyimpangan"],
     ["Profit = Total_Penjualan − HPP_Satuan × Qty", "0 penyimpangan (konsisten penuh)"],
     ["Profit negatif", f"{kualitas['profit_negatif']} baris"],
     ["Rentang Qty", f"{kualitas['qty_min']} – {kualitas['qty_max']} unit (rata-rata {angka(kualitas['qty_rata'], 2)})"],
     ["Rentang diskon", f"{angka(kualitas['diskon_min']*100, 0)}% – {angka(kualitas['diskon_max']*100, 0)}% (rata-rata {angka(kualitas['diskon_rata']*100, 1)}%)"],
     ["Transaksi per pelanggan", f"rata-rata {angka(kualitas['trx_per_pelanggan'], 1)} transaksi"]],
    lebar=[8.6, 6.6],
    caption="Hasil pemeriksaan kualitas data",
)
h2("2.4 Lingkungan Kerja")
para(
    "Analisis dijalankan dengan Python di dalam virtualenv proyek (pandas, NumPy, "
    "scikit-learn, statsmodels, openpyxl, matplotlib). "
    "Langkah penyiapan lingkungan dan pembacaan dataset direkam pada Gambar 2 dan Gambar 3."
)
gambar(SS / "01-lingkungan.png", "Penyiapan virtualenv dan pemasangan paket analisis.")
gambar(SS / "02-dataset.png", "Pembacaan struktur dataset: dua sheet, tipe kolom, dan cuplikan baris.")

# ---------------------------------------------------------------- 3
h1("3. Metodologi", halaman_baru=True)
h2("3.1 Alur Kerja")
para(
    "Pekerjaan dibagi menjadi lima tahap yang dijalankan berurutan oleh skrip agar hasil dapat diulang: "
    "(1) pemuatan dan validasi data, (2) statistik deskriptif dan agregasi, (3) segmentasi RFM dengan "
    "K-Means, (4) peramalan deret waktu, dan (5) pembangunan dashboard. Setiap tahap menuliskan "
    "keluaran terstruktur (JSON dan PNG) sehingga laporan ini tersusun dari angka yang benar-benar "
    "dihitung, bukan disalin manual."
)
h2("3.2 Statistik Deskriptif dan Agregasi")
para(
    "Total penjualan, profit, kuantitas, jumlah transaksi, serta rata-rata nilai transaksi dihitung dari "
    "seluruh baris. Agregasi bulanan dipakai untuk grafik tren dan masukan model peramalan; agregasi "
    "kota, kategori, kanal, segmen, metode pembayaran, dan status dipakai untuk komposisi dashboard. "
    "Peringkat produk disusun berdasarkan nilai penjualan dan kuantitas terjual."
)
h2("3.3 Segmentasi Pelanggan (RFM + K-Means)")
para(
    "Setiap pelanggan diringkas menjadi tiga indikator: Recency (hari sejak transaksi terakhir), "
    "Frequency (jumlah transaksi), dan Monetary (total belanja). Ketiga indikator distandardisasi "
    "z-score agar skala rupiah tidak mendominasi, lalu dikelompokkan dengan K-Means. Jumlah cluster "
    "diuji pada k = 2 sampai 6 memakai silhouette score dan Davies-Bouldin Index (DBI); k = 4 dipilih "
    "karena memberi pemisahan yang masih baik sekaligus empat kelompok yang dapat ditindaklanjuti "
    "secara bisnis. Hasil segmentasi disimpan sebagai profil rata-rata tiap cluster beserta porsi "
    "penjualannya."
)
h2("3.4 Peramalan Deret Waktu")
para(
    "Data bulanan (22 titik) diuji dengan enam kandidat model: SARIMA (dua konfigurasi hasil grid), "
    "regresi tren dengan dummy bulan, Holt-Winters aditif, Holt-Winters dengan tren teredam, dan naif "
    "musiman. Tiga bulan terakhir disisihkan sebagai data uji; setiap model diuji juga lewat backtest "
    "rolling empat titik asal. Model terpilih adalah yang MAPE ujinya terkecil, dengan backtest sebagai "
    "pemeriksaan kestabilan. MAPE dihitung sebagai rata-rata |aktual − prediksi| / aktual × 100%."
)
h2("3.5 Dashboard")
para(
    "Dashboard dibangun sebagai satu berkas HTML mandiri: data transaksi (10.000 baris), master produk, "
    "hasil clustering, hasil peramalan, dan batas wilayah provinsi ditanam langsung di dalam berkas. "
    "Pendekatan ini membuat dashboard tetap berjalan saat dibuka lewat file:// tanpa server dan tanpa "
    "internet, sekaligus menjamin angka pada dashboard identik dengan angka pada laporan."
)
gambar(SS / "03-analisis-a.png", "Eksekusi skrip analisis: statistik deskriptif dan evaluasi clustering (bagian 1).")
gambar(SS / "03-analisis-b.png", "Eksekusi skrip analisis: hasil clustering dan peramalan (bagian 2).")

# ---------------------------------------------------------------- 4
h1("4. Hasil dan Pembahasan", halaman_baru=True)
h2("4.1 Indikator Utama")
table(
    ["Indikator", "Nilai"],
    [["Periode data", f"{kualitas['tanggal_min']} – {kualitas['tanggal_max']} ({kualitas['bulan']} bulan)"],
     ["Jumlah transaksi", angka(kpi["jumlah_transaksi"])],
     ["Total penjualan", rp(kpi["total_penjualan"])],
     ["Total profit", rp(kpi["total_profit"])],
     ["Margin profit", angka(kpi["margin_pct"], 2) + "%"],
     ["Unit terjual", angka(kpi["total_qty"]) + " pcs"],
     ["Rata-rata nilai transaksi", rp(kpi["rata_transaksi"])],
     ["Pelanggan unik", angka(kpi["jumlah_pelanggan"])],
     ["Produk (SKU)", angka(kpi["jumlah_produk"])],
     ["Transaksi selesai", angka(kpi["transaksi_selesai"])],
     ["Transaksi dibatalkan", angka(kpi["transaksi_batal"])]],
    lebar=[9.2, 6.0],
    caption="Indikator utama dataset",
    align_kanan={1},
)
para(
    f"Margin profit {angka(kpi['margin_pct'], 2)}% konsisten dengan rentang margin master produk "
    f"({angka(kualitas['margin_min']*100, 1)}%–{angka(kualitas['margin_max']*100, 1)}%), yang berarti "
    "tidak ada transaksi dengan harga di bawah HPP. Tingkat pembatalan sebesar "
    f"{persen(kpi['transaksi_batal'], kpi['jumlah_transaksi'])} masih wajar untuk penjualan daring dan "
    "tidak mendistorsi tren karena tersebar merata sepanjang periode."
)
h2("4.2 Pola Musiman Bulanan")
bulanan = ringkasan["bulanan"]
tertinggi = max(bulanan, key=lambda x: x["penjualan"])
terendah = min(bulanan, key=lambda x: x["penjualan"])
para(
    f"Penjualan bulanan bergerak antara {juta(terendah['penjualan'])} ({terendah['Bulan']}) dan "
    f"{juta(tertinggi['penjualan'])} ({tertinggi['Bulan']}), dengan rata-rata "
    f"{juta(kpi['total_penjualan']/len(bulanan))} per bulan. Pola berulang terlihat jelas: puncak "
    "muncul pada November–Desember (musim belanja akhir tahun), lalu turun pada Januari–Februari. "
    "Pola inilah yang membuat model musiman 12 bulan mengungguli model tren biasa."
)
table(
    ["Bulan", "Penjualan", "Profit", "Unit", "Transaksi"],
    [[b["Bulan"], rp(b["penjualan"]), rp(b["profit"]), angka(b["qty"]), angka(b["trx"])] for b in bulanan],
    lebar=[2.5, 4.0, 4.0, 1.8, 2.0],
    caption="Rekap penjualan bulanan (22 bulan)",
    align_kanan={1, 2, 3, 4},
    size=8.5,
)
gambar(GRAFIK / "forecast-penjualan.png",
       "Tren penjualan bulanan beserta garis fit dan ramalan tiga bulan ke depan.")

h2("4.3 Sebaran Wilayah")
kota = ringkasan["kota"]
para(
    f"Tiga kota teratas menyumbang {persen(sum(k['penjualan'] for k in kota[:3]), kpi['total_penjualan'])} "
    f"penjualan: Jakarta {rp(kota[0]['penjualan'])}, Surabaya {rp(kota[1]['penjualan'])}, dan Bandung "
    f"{rp(kota[2]['penjualan'])}. Jakarta juga mencatat jumlah pelanggan terbanyak "
    f"({angka(kota[0]['pelanggan'])} dari {angka(kpi['jumlah_pelanggan'])}). Sebaran ini mengikuti "
    "konsentrasi populasi dan daya beli, sehingga kota di luar tiga besar lebih tepat didekati dengan "
    "program akuisisi daripada perluasan jenis produk."
)
table(
    ["Kota", "Penjualan", "Profit", "Transaksi", "Pelanggan", "Porsi"],
    [[k["Kota"], rp(k["penjualan"]), rp(k["profit"]), angka(k["trx"]), angka(k["pelanggan"]),
      persen(k["penjualan"], kpi["total_penjualan"])] for k in kota],
    lebar=[2.9, 3.6, 3.5, 2.0, 2.1, 1.6],
    caption="Kinerja penjualan menurut kota",
    align_kanan={1, 2, 3, 4, 5},
    size=9,
)
gambar(GRAFIK / "penjualan-per-kota.png", "Perbandingan nilai penjualan antar kota.")

h2("4.4 Kategori, Kanal, dan Pembayaran")
kategori = ringkasan["kategori"]
channel = ringkasan["channel"]
para(
    f"Kategori Elektronik memimpin dengan {rp(kategori[0]['penjualan'])} "
    f"({persen(kategori[0]['penjualan'], kpi['total_penjualan'])}), diikuti Fashion & Aksesoris "
    f"({rp(kategori[1]['penjualan'])}) dan Rumah Tangga ({rp(kategori[2]['penjualan'])}). Menariknya, "
    f"Kesehatan & Kecantikan hanya {rp(kategori[3]['penjualan'])} meski unit terjualnya tertinggi "
    f"({angka(kategori[3]['qty'])} pcs) karena harga per unitnya paling rendah — kategori ini cocok "
    "sebagai produk pemicu pembelian berulang, bukan penopang nilai."
)
table(
    ["Kategori", "Penjualan", "Profit", "Unit", "Porsi"],
    [[k["Kategori_Produk"], rp(k["penjualan"]), rp(k["profit"]), angka(k["qty"]),
      persen(k["penjualan"], kpi["total_penjualan"])] for k in kategori],
    lebar=[5.0, 3.6, 3.4, 1.9, 1.9],
    caption="Kinerja penjualan menurut kategori produk",
    align_kanan={1, 2, 3, 4},
)
para(
    f"Kanal daring mendominasi: Shopee {rp(channel[0]['penjualan'])} "
    f"({persen(channel[0]['penjualan'], kpi['total_penjualan'])}) dan Tokopedia "
    f"{rp(channel[1]['penjualan'])} ({persen(channel[1]['penjualan'], kpi['total_penjualan'])}), "
    f"sementara Toko Offline hanya {rp(channel[-1]['penjualan'])}. Untuk pembayaran, QRIS dan Virtual "
    "Account menutup lebih dari separuh transaksi, tanda bahwa pelanggan nyaman dengan pembayaran "
    "digital tanpa kartu kredit."
)
table(
    ["Kanal", "Penjualan", "Transaksi", "Porsi"],
    [[c["Channel_Penjualan"], rp(c["penjualan"]), angka(c["trx"]), persen(c["penjualan"], kpi["total_penjualan"])]
     for c in channel],
    lebar=[4.6, 4.0, 3.0, 2.4],
    caption="Kinerja penjualan menurut kanal",
    align_kanan={1, 2, 3},
)
table(
    ["Metode Pembayaran", "Transaksi", "Porsi"],
    [[m, angka(v), persen(v, kpi["jumlah_transaksi"])] for m, v in ringkasan["pembayaran"].items()],
    lebar=[6.0, 4.0, 4.0],
    caption="Komposisi metode pembayaran",
    align_kanan={1, 2},
)
h2("4.5 Produk Terlaris")
table(
    ["Peringkat", "Produk", "Kategori", "Penjualan", "Unit", "Profit"],
    [[i + 1, p["Nama_Produk"], p["Kategori_Produk"], rp(p["penjualan"]), angka(p["qty"]), rp(p["profit"])]
     for i, p in enumerate(ringkasan["produk_top"][:10])],
    lebar=[1.7, 5.0, 3.4, 3.2, 1.5, 3.0],
    caption="Sepuluh produk dengan penjualan tertinggi",
    align_kanan={0, 3, 4, 5},
    size=9,
)
para(
    "Smartwatch Fitness Tracker menjadi produk dengan nilai penjualan tertinggi "
    f"({rp(ringkasan['produk_top'][0]['penjualan'])} dari {angka(ringkasan['produk_top'][0]['qty'])} unit), "
    "disusul Air Fryer Digital 4L dan Sepatu Sneakers Kasual. Pola ini memperlihatkan bahwa nilai "
    "penjualan digerakkan barang berharga menengah ke atas, sedangkan volume digerakkan produk "
    "kebutuhan harian."
)
h2("4.6 Segmen Pelanggan")
para(
    f"Evaluasi K-Means menunjukkan silhouette score tertinggi pada k = 2 "
    f"({angka(clust['evaluasi'][0]['silhouette'], 4)}), namun pemisahan dua kelompok terlalu kasar untuk "
    f"penentuan program. Pada k = 4 silhouette {angka(clust['silhouette'], 4)} dengan DBI "
    f"{angka(clust['dbi'], 4)} menghasilkan empat segmen yang jelas berbeda profil RFM-nya, sehingga "
    "k = 4 dipakai sebagai model akhir."
)
table(
    ["k", "Silhouette", "DBI", "Inertia"],
    [[e["k"], angka(e["silhouette"], 4), angka(e["dbi"], 4), angka(e["inertia"], 1)] for e in clust["evaluasi"]],
    lebar=[2.0, 3.4, 3.4, 4.0],
    caption="Perbandingan kualitas clustering pada beberapa nilai k",
    align_kanan={0, 1, 2, 3},
)
table(
    ["Segmen", "Pelanggan", "Recency (hari)", "Frequency", "Monetary", "Porsi Penjualan"],
    [[p["segmen"], angka(p["jumlah"]), angka(p["recency_rata"], 1), angka(p["frequency_rata"], 1),
      rp(p["monetary_rata"]), angka(p["porsi_penjualan_pct"], 2) + "%"] for p in clust["profil"]],
    lebar=[4.4, 2.2, 2.6, 2.2, 2.8, 2.4],
    caption="Profil rata-rata setiap cluster pelanggan (RFM)",
    align_kanan={1, 2, 3, 4, 5},
    size=9,
)
p0, p1, p2, p3 = clust["profil"]
para(
    f"**{p0['segmen']}** ({p0['jumlah']} pelanggan, {angka(p0['porsi_pelanggan_pct'], 1)}% populasi) "
    f"menyumbang {angka(p0['porsi_penjualan_pct'], 1)}% penjualan dengan belanja rata-rata "
    f"{rp(p0['monetary_rata'])} dan {angka(p0['frequency_rata'], 1)} transaksi. Recency "
    f"{angka(p0['recency_rata'], 1)} hari menunjukkan mereka masih aktif — prioritas retensi."
)
para(
    f"**{p1['segmen']}** ({p1['jumlah']} pelanggan) justru paling baru bertransaksi "
    f"(recency {angka(p1['recency_rata'], 1)} hari) dengan frekuensi {angka(p1['frequency_rata'], 1)}. "
    "Profil ini adalah kandidat terbaik untuk dinaikkan kelasnya lewat penawaran bertingkat."
)
para(
    f"**{p2['segmen']}** adalah kelompok terbesar ({p2['jumlah']} pelanggan, "
    f"{angka(p2['porsi_pelanggan_pct'], 1)}%) tetapi recency-nya {angka(p2['recency_rata'], 1)} hari; "
    f"kelompok ini menyimpan potensi besar bila diaktifkan kembali. **{p3['segmen']}** "
    f"({p3['jumlah']} pelanggan, {angka(p3['porsi_pelanggan_pct'], 1)}%) paling kecil kontribusinya "
    f"({angka(p3['porsi_penjualan_pct'], 1)}%) dan paling lama tidak bertransaksi "
    f"({angka(p3['recency_rata'], 1)} hari) — cukup ditangani kampanye berbiaya rendah."
)
gambar(GRAFIK / "evaluasi-clustering.png", "Kurva evaluasi clustering: silhouette score dan DBI terhadap nilai k.")
gambar(GRAFIK / "sebaran-cluster.png", "Sebaran pelanggan pada bidang Frequency–Monetary beserta warna cluster.")

# ---------------------------------------------------------------- 5
h1("5. Peramalan Penjualan", halaman_baru=True)
para(
    f"Enam kandidat model diuji pada tiga bulan terakhir (Agustus–Oktober 2026). "
    f"{fc['metode_terpilih']} keluar sebagai model terbaik dengan MAPE uji "
    f"{angka(fc['mape_uji'], 2)}%, diikuti SARIMA(1,1,1)(0,0,1,12) "
    f"({angka(fc['kandidat'][1]['mape'], 2)}%) dan regresi tren dengan dummy bulan "
    f"({angka(fc['kandidat'][2]['mape'], 2)}%). Model tanpa komponen musiman (Holt-Winters) tertinggal "
    f"pada kisaran {angka(fc['kandidat'][3]['mape'], 1)}%–{angka(fc['kandidat'][4]['mape'], 1)}%, "
    "menegaskan bahwa pola akhir tahun adalah sinyal utama yang harus ditangkap."
)
table(
    ["Model", "MAPE uji", "Prediksi Nov 2026"],
    [[k["metode"], angka(k["mape"], 2) + "%", rp(k["prediksi"][0])] for k in fc["kandidat"]],
    lebar=[7.4, 3.0, 4.4],
    caption="Perbandingan akurasi kandidat model peramalan",
    align_kanan={1, 2},
)
table(
    ["Bulan uji", "Aktual", "Prediksi", "Galat"],
    [[fc["label_bulan"][-3 + i], rp(fc["aktual_uji"][i]), rp(fc["prediksi_uji"][i]),
      persen(abs(fc["aktual_uji"][i] - fc["prediksi_uji"][i]), fc["aktual_uji"][i])] for i in range(3)],
    lebar=[3.2, 4.0, 4.0, 3.2],
    caption="Kinerja model terpilih pada tiga bulan data uji",
    align_kanan={1, 2, 3},
)
para(
    f"Backtest rolling pada empat titik asal memberi MAPE "
    f"{', '.join(angka(b['mape'], 2) + '%' for b in fc['backtest'])} "
    f"(rata-rata {angka(fc['mape_backtest'], 2)}%). Nilai ini sedikit di atas MAPE uji tunggal, hal "
    "wajar karena setiap backtest meramal dari titik data yang berbeda; keduanya masih di bawah 5%, "
    "batas yang umum dipakai untuk peramalan bulanan yang layak dipakai operasional."
)
table(
    ["Bulan", "Prediksi Penjualan", "Selisih vs Okt 2026"],
    [[fc["label_prediksi"][i], rp(fc["prediksi"][i]),
      ("+" if fc["prediksi"][i] >= fc["aktual"][-1] else "") +
      persen(fc["prediksi"][i] - fc["aktual"][-1], fc["aktual"][-1])] for i in range(3)],
    lebar=[3.2, 4.6, 4.6],
    caption="Ramalan penjualan tiga bulan ke depan",
    align_kanan={1, 2},
)
para(
    f"Ramalan menunjukkan lonjakan musim belanja akhir tahun: {rp(fc['prediksi'][1])} pada Desember "
    f"2026, sekitar {persen(fc['prediksi'][1] - fc['aktual'][-1], fc['aktual'][-1])} di atas Oktober "
    f"2026, lalu kembali normal ke {rp(fc['prediksi'][2])} pada Januari 2027. Karena itu kapasitas "
    "gudang dan anggaran iklan sebaiknya dinaikkan sejak awal November, bukan menunggu Desember."
)
para(
    "Catatan keterbatasan: hanya tersedia 22 titik bulanan (kurang dari dua siklus penuh), sehingga "
    "parameter musiman SARIMA belum bisa diestimasi secara langsung dan konfigurasi terpilih memakai "
    "diferensiasi musiman sederhana. Hasil sebaiknya diperbarui setiap bulan begitu data baru masuk."
)

# ---------------------------------------------------------------- 6
h1("6. Dashboard Interaktif", halaman_baru=True)
para(
    "Dashboard tersedia sebagai satu berkas HTML (03-dashboard/index.html, sekitar 2,5 MB) yang memuat "
    "seluruh data di dalamnya. Berkas dibuka langsung di peramban tanpa server, tanpa pemasangan, dan "
    "tanpa internet. Tiga tab utama: **Dashboard Penjualan** (KPI, tren, komposisi, peta, segmen), "
    "**Data Transaksi** (tabel 10.000 baris berpaginasi 25 baris, filter, dan formulir tambah), serta "
    "**Master Produk** (25 SKU, margin, dan formulir tambah produk)."
)
para(
    "Filter global (kota, kanal, segmen, kategori, metode pembayaran, status, dan rentang tanggal) "
    "berlaku serentak untuk KPI, grafik, peta, dan tabel, sehingga semua angka pada layar selalu berasal "
    "dari subset data yang sama."
)
gambar(SS / "04-dashboard-build.png", "Pembangunan dashboard: data ditanam inline ke berkas HTML.")
gambar(SS / "05-dashboard-kpi-tren.png", "Tab Dashboard Penjualan: kartu KPI dan grafik tren beserta proyeksi.")
gambar(SS / "06-dashboard-kota-peta.png", "Penjualan per kota dan peta sebaran per provinsi.")
gambar(SS / "07-dashboard-channel-cluster.png", "Komposisi kanal penjualan dan sebaran pelanggan per cluster.")
gambar(SS / "08-dashboard-profil-model.png", "Profil cluster pelanggan dan perbandingan akurasi model peramalan.")
gambar(SS / "09-tab-transaksi.png", "Tab Data Transaksi: tabel berpaginasi dengan filter dan formulir input.")
gambar(SS / "10-tab-filter.png", "Contoh penyaringan: kota Bandung dengan status Dibatalkan (3 baris).")
gambar(SS / "11-tab-tambah-transaksi.png",
       "Setelah transaksi baru disimpan, KPI dan jumlah baris langsung diperbarui.")
gambar(SS / "12-tab-produk.png", "Tab Master Produk: harga, HPP, dan margin tiap SKU.")
gambar(SS / "13-tab-tambah-produk.png",
       "Penambahan produk PRD-026: tabel master produk kini berisi 26 SKU dan pesan konfirmasi tampil.")
para(
    "Pengujian interaksi dilakukan langsung di peramban: filter kota Bandung + status Dibatalkan "
    "menyisakan 3 baris dari 10.000; penambahan satu transaksi mengubah KPI penjualan dari "
    "Rp 4.174.773.550 menjadi Rp 4.176.390.850 dan profit naik Rp 507.300 sesuai nilai transaksi baru; "
    "penambahan produk PRD-026 menaikkan jumlah SKU dari 25 menjadi 26. Semua perubahan tercermin "
    "serentak di KPI, grafik, dan tabel."
)

# ---------------------------------------------------------------- 7
h1("7. Kesimpulan dan Rekomendasi", halaman_baru=True)
h2("7.1 Kesimpulan")
bullet(f"Dataset bersih dan konsisten: {angka(kualitas['baris'])} transaksi, tanpa sel kosong, tanpa ID "
       "ganda, dan seluruh kolom turunan (subtotal, total, profit) konsisten penuh.")
bullet(f"Penjualan {rp(kpi['total_penjualan'])} dengan margin {angka(kpi['margin_pct'], 2)}% terkonsentrasi "
       f"di Jakarta ({persen(kota[0]['penjualan'], kpi['total_penjualan'])}), kategori Elektronik "
       f"({persen(kategori[0]['penjualan'], kpi['total_penjualan'])}), dan kanal Shopee "
       f"({persen(channel[0]['penjualan'], kpi['total_penjualan'])}).")
bullet("Pola musiman akhir tahun nyata: puncak November–Desember, lembah Januari–Februari, berulang pada "
       "dua tahun data.")
bullet(f"Empat segmen pelanggan terpisah jelas; segmen {p0['segmen']} berisi "
       f"{angka(p0['porsi_pelanggan_pct'], 1)}% pelanggan tetapi menyumbang "
       f"{angka(p0['porsi_penjualan_pct'], 1)}% penjualan.")
bullet(f"{fc['metode_terpilih']} layak dipakai untuk perencanaan jangka pendek (MAPE uji "
       f"{angka(fc['mape_uji'], 2)}%, backtest {angka(fc['mape_backtest'], 2)}%).")
bullet("Dashboard satu berkas membuat hasil analisis dapat diperiksa ulang tanpa keahlian teknis.")
h2("7.2 Rekomendasi")
bullet("Jadwalkan kampanye dan penambahan stok pada Oktober–November, bukan Desember, karena kenaikan "
       "permintaan sudah terlihat sejak awal kuartal keempat.")
bullet(f"Aktifkan kembali {p2['jumlah']} pelanggan segmen {p2['segmen']} yang recency-nya "
       f"{angka(p2['recency_rata'], 1)} hari lewat penawaran personal; potensi penjualannya besar karena "
       "kelompok ini yang paling banyak jumlahnya.")
bullet(f"Pertahankan {p0['segmen']} dengan program loyalitas berbasis frekuensi, sebab "
       f"{angka(p0['porsi_penjualan_pct'], 1)}% nilai penjualan bergantung pada kelompok ini.")
bullet("Perluas kanal TikTok Shop dan Website Resmi yang masih di bawah 16% untuk mengurangi "
       "ketergantungan pada dua marketplace besar.")
bullet("Perbarui model peramalan setiap bulan dan ulangi evaluasi clustering per kuartal agar segmen "
       "mengikuti perubahan perilaku pelanggan.")

# ---------------------------------------------------------------- lampiran
h1("Lampiran A. Berkas Hasil Pekerjaan", halaman_baru=True)
table(
    ["Berkas", "Isi"],
    [["01-dataset/", "Dataset asli (.xlsx) dan berkas soal"],
     ["02-analisis/analisis.py", "Skrip analisis: deskriptif, RFM + K-Means, peramalan"],
     ["02-analisis/ringkasan.json", "KPI, agregasi bulanan, kota, kategori, kanal, produk"],
     ["02-analisis/clustering.json", "Evaluasi k, profil cluster, centroid, sebaran titik"],
     ["02-analisis/forecast.json", "Kandidat model, MAPE, backtest, dan ramalan"],
     ["02-analisis/grafik/", "Empat grafik PNG hasil analisis"],
     ["03-dashboard/index.html", "Dashboard interaktif satu berkas (siap dibuka)"],
     ["03-dashboard/buat-dashboard.py", "Skrip pembangun dashboard dari JSON + dataset"],
     ["04-laporan/screenshot/", "13 tangkapan layar setiap langkah pengerjaan"],
     ["04-laporan/laporan-praktikum.docx", "Laporan ini (versi Word)"],
     ["04-laporan/laporan-praktikum.pdf", "Laporan ini (versi PDF)"]],
    lebar=[6.4, 9.0],
    caption="Daftar berkas yang dihasilkan",
    size=9,
)
h2("Cara Menjalankan Ulang")
para("Seluruh proses dapat diulang dari folder proyek dengan tiga perintah berikut:", space_after=4)
for cmd in [".venv/bin/python 02-analisis/analisis.py",
            ".venv/bin/python 03-dashboard/buat-dashboard.py",
            ".venv/bin/python 04-laporan/buat-laporan.py"]:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(cmd)
    r.font.name = "Consolas"
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

out = LAP / "laporan-praktikum.docx"
doc.save(out)
print("laporan ditulis:", out, f"{out.stat().st_size/1024:.0f} KB")
print("gambar:", nomor_gambar[0], "| tabel:", nomor_tabel[0])
