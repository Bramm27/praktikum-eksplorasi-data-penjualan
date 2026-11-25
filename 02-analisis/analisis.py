"""Analisis dataset transaksi penjualan: EDA, clustering RFM, forecasting penjualan.

Input : ../Dataset_Transaksi_Penjualan_10000_*.xlsx
Output: ringkasan.json, clustering.json, forecast.json, grafik/*.png
"""
import json
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score, davies_bouldin_score
from sklearn.preprocessing import StandardScaler
from statsmodels.tsa.holtwinters import ExponentialSmoothing
from statsmodels.tsa.statespace.sarimax import SARIMAX

warnings.filterwarnings("ignore")
BASE = Path(__file__).resolve().parent
ROOT = BASE.parent
OUT = BASE
GRAFIK = BASE / "grafik"
GRAFIK.mkdir(exist_ok=True)

KANDIDAT = sorted(ROOT.glob("01-dataset/Dataset_Transaksi_Penjualan_*.xlsx")) or \
    sorted(ROOT.glob("Dataset_Transaksi_Penjualan_*.xlsx"))
if not KANDIDAT:
    raise SystemExit("Dataset Excel tidak ditemukan di 01-dataset/ maupun root proyek.")
XLSX = KANDIDAT[0]
trx = pd.read_excel(XLSX, sheet_name="Data_Transaksi")
prd = pd.read_excel(XLSX, sheet_name="Master_Produk")

trx["Tanggal"] = pd.to_datetime(trx["Tanggal"])
trx["Bulan"] = trx["Tanggal"].dt.to_period("M").astype(str)
selesai = trx[trx["Status_Transaksi"] == "Selesai"].copy()

# ---------------------------------------------------------------- ringkasan
bulanan = trx.groupby("Bulan").agg(
    penjualan=("Total_Penjualan", "sum"),
    profit=("Profit", "sum"),
    qty=("Qty", "sum"),
    trx=("ID_Transaksi", "count"),
).reset_index()

kota = trx.groupby("Kota").agg(
    penjualan=("Total_Penjualan", "sum"),
    profit=("Profit", "sum"),
    qty=("Qty", "sum"),
    trx=("ID_Transaksi", "count"),
    pelanggan=("Nama_Pelanggan", "nunique"),
).sort_values("penjualan", ascending=False).reset_index()

kategori = trx.groupby("Kategori_Produk").agg(
    penjualan=("Total_Penjualan", "sum"),
    profit=("Profit", "sum"),
    qty=("Qty", "sum"),
).sort_values("penjualan", ascending=False).reset_index()

channel = trx.groupby("Channel_Penjualan").agg(
    penjualan=("Total_Penjualan", "sum"),
    trx=("ID_Transaksi", "count"),
).sort_values("penjualan", ascending=False).reset_index()

produk = trx.groupby(["Nama_Produk", "Kategori_Produk"]).agg(
    penjualan=("Total_Penjualan", "sum"),
    qty=("Qty", "sum"),
    profit=("Profit", "sum"),
).sort_values("penjualan", ascending=False).reset_index()

ringkasan = {
    "periode": {"awal": str(trx["Tanggal"].min().date()), "akhir": str(trx["Tanggal"].max().date()),
                "jumlah_bulan": int(trx["Bulan"].nunique())},
    "kpi": {
        "total_penjualan": int(trx["Total_Penjualan"].sum()),
        "total_profit": int(trx["Profit"].sum()),
        "total_qty": int(trx["Qty"].sum()),
        "jumlah_transaksi": int(len(trx)),
        "jumlah_pelanggan": int(trx["Nama_Pelanggan"].nunique()),
        "jumlah_produk": int(trx["Nama_Produk"].nunique()),
        "rata_transaksi": int(trx["Total_Penjualan"].mean()),
        "margin_pct": round(trx["Profit"].sum() / trx["Total_Penjualan"].sum() * 100, 2),
        "transaksi_selesai": int((trx["Status_Transaksi"] == "Selesai").sum()),
        "transaksi_batal": int((trx["Status_Transaksi"] == "Dibatalkan").sum()),
    },
    "bulanan": bulanan.to_dict("records"),
    "kota": kota.to_dict("records"),
    "kategori": kategori.to_dict("records"),
    "channel": channel.to_dict("records"),
    "produk_top": produk.head(10).to_dict("records"),
    "segmen": trx["Segmen_Pelanggan"].value_counts().to_dict(),
    "status": trx["Status_Transaksi"].value_counts().to_dict(),
    "pembayaran": trx["Metode_Pembayaran"].value_counts().to_dict(),
}
json.dump(ringkasan, open(OUT / "ringkasan.json", "w"), indent=1, default=str)

# ------------------------------------------------------------------ clustering RFM
acuan = trx["Tanggal"].max() + pd.Timedelta(days=1)
rfm = trx.groupby("Nama_Pelanggan").agg(
    Recency=("Tanggal", lambda s: (acuan - s.max()).days),
    Frequency=("ID_Transaksi", "count"),
    Monetary=("Total_Penjualan", "sum"),
)
rfm["Monetary"] = rfm["Monetary"].astype(float)
X = rfm[["Recency", "Frequency", "Monetary"]].copy()
X["Recency"] = np.log1p(X["Recency"])
X["Frequency"] = np.log1p(X["Frequency"])
X["Monetary"] = np.log1p(X["Monetary"])
Xs = StandardScaler().fit_transform(X)

evaluasi = []
for k in range(2, 7):
    km = KMeans(n_clusters=k, n_init=20, random_state=42).fit(Xs)
    evaluasi.append({
        "k": k,
        "silhouette": round(float(silhouette_score(Xs, km.labels_)), 4),
        "dbi": round(float(davies_bouldin_score(Xs, km.labels_)), 4),
        "inertia": round(float(km.inertia_), 1),
    })
best = max(evaluasi, key=lambda e: e["silhouette"])
# aturan pemilihan: k terkecil dengan DBI < 1.10 dan silhouette >= 0.28 (granularitas bisnis)
kandidat_k = [e for e in evaluasi if e["dbi"] < 1.10 and e["silhouette"] >= 0.28]
best = kandidat_k[0] if kandidat_k else best
km = KMeans(n_clusters=best["k"], n_init=50, random_state=42).fit(Xs)
rfm["Cluster"] = km.labels_

profil = rfm.groupby("Cluster").agg(
    jumlah=("Recency", "size"),
    recency_rata=("Recency", "mean"),
    frequency_rata=("Frequency", "mean"),
    monetary_rata=("Monetary", "mean"),
    monetary_total=("Monetary", "sum"),
).reset_index()
profil["porsi_pelanggan_pct"] = (profil["jumlah"] / len(rfm) * 100).round(2)
profil["porsi_penjualan_pct"] = (profil["monetary_total"] / rfm["Monetary"].sum() * 100).round(2)

# urutkan cluster: monetary tertinggi = segmen terbaik
urut = profil.sort_values("monetary_rata", ascending=False)["Cluster"].tolist()
label = ["Pelanggan Bernilai Tinggi", "Pelanggan Potensial", "Pelanggan Bernilai Sedang",
         "Pelanggan Berisiko Churn", "Pelanggan Pasif", "Pelanggan Low Value"]
peta = {c: label[i] for i, c in enumerate(urut)}
rfm["Segmen"] = rfm["Cluster"].map(peta)
profil["segmen"] = profil["Cluster"].map(peta)
profil = profil.sort_values("monetary_rata", ascending=False)

centro = pd.DataFrame(
    StandardScaler().fit(X).inverse_transform(km.cluster_centers_),
    columns=["Recency", "Frequency", "Monetary"],
)
centro = np.expm1(centro)
centro["Cluster"] = range(len(centro))

clustering = {
    "metode": "K-Means pada RFM (log1p + StandardScaler)",
    "evaluasi": evaluasi,
    "k_terpilih": best["k"],
    "silhouette": best["silhouette"],
    "dbi": best["dbi"],
    "profil": json.loads(profil.round(2).to_json(orient="records")),
    "centroid": json.loads(centro.round(2).to_json(orient="records")),
    "sebaran": json.loads(
        rfm.reset_index()[["Nama_Pelanggan", "Recency", "Frequency", "Monetary", "Cluster", "Segmen"]]
        .round(0).to_json(orient="records")
    ),
}
json.dump(clustering, open(OUT / "clustering.json", "w"), indent=1)

# ------------------------------------------------------------------ forecasting
seri = bulanan.set_index("Bulan")["penjualan"].astype(float)
seri.index = pd.PeriodIndex(seri.index, freq="M")
HORIZON = 3
train, test = seri[:-HORIZON], seri[-HORIZON:]


def mape(a, f):
    a, f = np.asarray(a, float), np.asarray(f, float)
    return float(np.mean(np.abs((a - f) / a)) * 100)


def regresi_tren_musiman(s):
    """Regresi linear: tren waktu + dummy bulan (musiman deterministik)."""
    from sklearn.linear_model import LinearRegression
    idx = pd.DataFrame({"t": np.arange(len(s)), "m": s.index.month})
    Xd = pd.get_dummies(idx["m"], prefix="m", drop_first=True).astype(float)
    Xd["t"] = idx["t"]
    lr = LinearRegression().fit(Xd.iloc[:-HORIZON], s.values[:-HORIZON])
    return lr.predict(Xd)


def prediksi(nama, s):
    """Kembalikan (prediksi_h, fitted) untuk horizon HORIZON dari data latih s."""
    n = len(s)
    if nama == "Holt-Winters (trend aditif, tanpa musiman)":
        m = ExponentialSmoothing(s, trend="add", seasonal=None, initialization_method="estimated").fit()
        return np.asarray(m.forecast(HORIZON), float), np.asarray(m.fittedvalues, float)
    if nama == "Holt-Winters (trend teredam)":
        m = ExponentialSmoothing(s, trend="add", damped_trend=True, seasonal=None,
                                 initialization_method="estimated").fit()
        return np.asarray(m.forecast(HORIZON), float), np.asarray(m.fittedvalues, float)
    if nama == "SARIMA(0,1,1)(0,1,0,12)":
        m = SARIMAX(s, order=(0, 1, 1), seasonal_order=(0, 1, 0, 12), trend="n",
                    enforce_stationarity=False, enforce_invertibility=False).fit(disp=False)
        return np.asarray(m.forecast(HORIZON), float), np.asarray(m.fittedvalues, float)
    if nama == "SARIMA(1,1,1)(0,0,1,12)":
        m = SARIMAX(s, order=(1, 1, 1), seasonal_order=(0, 0, 1, 12), trend="n",
                    enforce_stationarity=False, enforce_invertibility=False).fit(disp=False)
        return np.asarray(m.forecast(HORIZON), float), np.asarray(m.fittedvalues, float)
    if nama == "Regresi tren + dummy bulan":
        p = regresi_tren_musiman(s)
        return p[-HORIZON:], p[:-HORIZON]
    if nama == "Naif musiman (bulan sama tahun lalu)":
        p = s.values[-12:][:HORIZON]
        return np.asarray(p, float), np.asarray(s.values, float)
    raise ValueError(nama)


METODE = [
    "Holt-Winters (trend aditif, tanpa musiman)",
    "Holt-Winters (trend teredam)",
    "SARIMA(0,1,1)(0,1,0,12)",
    "SARIMA(1,1,1)(0,0,1,12)",
    "Regresi tren + dummy bulan",
    "Naif musiman (bulan sama tahun lalu)",
]

kandidat = []
for nama in METODE:
    try:
        pred, _ = prediksi(nama, train)
        kandidat.append({"metode": nama, "mape": round(mape(test.values, pred), 2),
                         "prediksi": [round(float(v)) for v in pred]})
    except Exception as exc:  # noqa: BLE001
        kandidat.append({"metode": nama, "mape": None, "error": str(exc)[:120], "prediksi": []})

kandidat = [k for k in kandidat if k["mape"] is not None]
kandidat.sort(key=lambda k: k["mape"])
terpilih = kandidat[0]

# backtest rolling-origin: 4 titik asal, horizon 3 bulan
backtest = []
for asal in range(len(seri) - HORIZON - 3, len(seri) - HORIZON + 1):
    latih, uji = seri[:asal], seri[asal:asal + HORIZON]
    try:
        p, _ = prediksi(terpilih["metode"], latih)
        backtest.append({"asal": str(latih.index[-1]), "mape": round(mape(uji.values, p), 2)})
    except Exception:  # noqa: BLE001
        continue
mape_backtest = round(float(np.mean([b["mape"] for b in backtest])), 2) if backtest else None

# model final dilatih pada seluruh data
pred_final, fit_values = prediksi(terpilih["metode"], seri)
idx_final = pd.period_range(seri.index[-1] + 1, periods=HORIZON, freq="M")

forecast = {
    "horizon": HORIZON,
    "periode_latih": [str(seri.index[0]), str(train.index[-1])],
    "periode_uji": [str(test.index[0]), str(test.index[-1])],
    "kandidat": kandidat,
    "metode_terpilih": terpilih["metode"],
    "mape_uji": terpilih["mape"],
    "backtest": backtest,
    "mape_backtest": mape_backtest,
    "aktual": [round(float(v)) for v in seri.values],
    "label_bulan": [str(p) for p in seri.index],
    "fit": [round(float(v)) for v in fit_values],
    "prediksi": [round(float(v)) for v in pred_final],
    "label_prediksi": [str(p) for p in idx_final],
    "aktual_uji": [round(float(v)) for v in test.values],
    "prediksi_uji": terpilih["prediksi"],
}
json.dump(forecast, open(OUT / "forecast.json", "w"), indent=1)

# ------------------------------------------------------------------ grafik
plt.rcParams.update({"font.size": 10, "axes.grid": True, "grid.alpha": .25,
                     "axes.spines.top": False, "axes.spines.right": False,
                     "figure.dpi": 150,
                     # seluruh teks grafik hitam
                     "text.color": "black", "axes.labelcolor": "black",
                     "axes.titlecolor": "black", "axes.edgecolor": "black",
                     "xtick.color": "black", "ytick.color": "black",
                     "legend.labelcolor": "black"})

fig, ax = plt.subplots(figsize=(10, 3.6))
x = np.arange(len(seri))
ax.plot(x, seri.values / 1e6, marker="o", color="#12355b", label="Aktual")
ax.plot(x, fit_values / 1e6, ls="--", color="#7d93b0", label="Fit model")
xf = np.arange(len(seri), len(seri) + HORIZON)
ax.plot(xf, pred_final / 1e6, marker="s", color="#17a398", label="Ramalan")
ax.set_xticks(list(x) + list(xf))
ax.set_xticklabels(list(seri.index) + list(idx_final), rotation=60, fontsize=7)
ax.set_ylabel("Penjualan (juta Rp)")
ax.set_title(f"Forecast Penjualan Bulanan — {terpilih['metode']} (MAPE uji {terpilih['mape']:.2f}%)")
ax.legend()
fig.tight_layout()
fig.savefig(GRAFIK / "forecast-penjualan.png")
plt.close(fig)

fig, ax = plt.subplots(figsize=(5.2, 3.4))
ks = [e["k"] for e in evaluasi]
ax.plot(ks, [e["silhouette"] for e in evaluasi], marker="o", color="#12355b", label="Silhouette")
ax.set_xlabel("Jumlah cluster (k)")
ax.set_ylabel("Silhouette")
ax2 = ax.twinx()
ax2.plot(ks, [e["dbi"] for e in evaluasi], marker="s", color="#17a398", label="DBI")
ax2.set_ylabel("Davies-Bouldin Index")
ax2.grid(False)
ax.set_title("Evaluasi Kualitas Clustering")
fig.tight_layout()
fig.savefig(GRAFIK / "evaluasi-clustering.png")
plt.close(fig)

fig, ax = plt.subplots(figsize=(6.4, 4.2))
warna = ["#12355b", "#2f8fd6", "#17a398", "#d98b23", "#8e6bb3"]
for i, (nama, g) in enumerate(rfm.groupby("Segmen")):
    ax.scatter(g["Frequency"], g["Monetary"] / 1e6, s=18, alpha=.65, label=nama, color=warna[i])
ax.set_xlabel("Frequency (jumlah transaksi)")
ax.set_ylabel("Monetary (juta Rp)")
ax.set_title(f"Sebaran Pelanggan per Cluster (k={best['k']})")
ax.legend(fontsize=7)
fig.tight_layout()
fig.savefig(GRAFIK / "sebaran-cluster.png")
plt.close(fig)

fig, ax = plt.subplots(figsize=(8, 3.6))
ax.bar(kota["Kota"], kota["penjualan"] / 1e6, color="#2f8fd6", edgecolor="#12355b", linewidth=.4)
ax.set_ylabel("Penjualan (juta Rp)")
ax.set_title("Penjualan per Wilayah (Kota)")
ax.tick_params(axis="x", rotation=35)
fig.tight_layout()
fig.savefig(GRAFIK / "penjualan-per-kota.png")
plt.close(fig)

print(json.dumps({
    "ringkasan": ringkasan["kpi"],
    "clustering": {"k": best["k"], "silhouette": best["silhouette"], "dbi": best["dbi"],
                   "profil": clustering["profil"]},
    "forecast": {"terpilih": terpilih["metode"], "mape": terpilih["mape"], "kandidat": kandidat,
                 "prediksi": forecast["prediksi"], "label": forecast["label_prediksi"]},
}, indent=1, default=str))
