# -*- coding: utf-8 -*-
"""
ops_parser.py — Backend pembacaan & parsing Excel OPS Bulanan Sumbagut.

Isi modul ini DISALIN dari `dashboard_ops_sumbagut.py` (bagian 0 konstanta,
1. Fungsi Format, 2. Load & Parsing Data, serta helper turunan `agg_growth`,
`ppdp_frames`, dan `pm_wide`) TANPA pemanggilan Streamlit. Logika parsing
tidak diubah, sehingga hasilnya identik dengan dashboard.

Karena bebas Streamlit, modul ini dapat di-import oleh:
  * infografis_ops.py   (Infographic Generator), dan
  * dashboard_ops_sumbagut.py (opsional: ganti blok parsing di dashboard dengan
    `from ops_parser import *` agar hanya ada SATU sumber kebenaran).
"""
from __future__ import annotations

import io
import re
import warnings
from datetime import date, datetime

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =========================================================
# 0. KONSTANTA STRUKTUR EXCEL
# =========================================================
PROVINSI = ["Sumut", "Aceh", "Riau", "Sumbar", "Kepri"]

SHEETS = {
    "umum": "BU-Kinerja Umum",
    "jenis": "BU-Jenis Kredit",
    "skala": "BU-Skala Kredit",
    "sektor": "Kredit BU Per Sektor",
}
# Sheet modul tambahan (opsional: jika tidak ada, hanya modul terkait yang nonaktif)
SHEETS_LAIN = {
    "bpd": "Bank Imut",
    "bpr": "BPR-Kinerja Umum",
    "pvml": "PVML",
    "ppdp": "PPDP",
    "pm": "PM",
}
SUMBER_KOREKSI = "Tabel koreksi (bawah)"
SUMBER_COGNOS = "Tabel Cognos (atas)"

# Posisi blok data tiap provinsi di Excel (sheet Bank Umum)
KOLOM_UMUM = {"Sumut": 0, "Aceh": 8, "Riau": 16, "Sumbar": 24, "Kepri": 32}
KOLOM_JENIS = {"Sumut": 0, "Aceh": 5, "Riau": 10, "Sumbar": 15, "Kepri": 20}
KOLOM_SKALA = {"Sumut": 0, "Aceh": 6, "Riau": 12, "Sumbar": 18, "Kepri": 24}
BARIS_SEKTOR = {"Sumut": 2, "Aceh": 9, "Riau": 16, "Sumbar": 23, "Kepri": 30}

# Pengenal nama provinsi pada header sheet (urutan penting: yang spesifik dulu)
PROV_KEYS = [
    ("kepulauan riau", "Kepri"), ("kep. riau", "Kepri"), ("kepri", "Kepri"),
    ("sumatera utara", "Sumut"), ("sumut", "Sumut"),
    ("sumatera barat", "Sumbar"), ("sumbar", "Sumbar"),
    ("aceh", "Aceh"), ("riau", "Riau"),
]
# Pemetaan bank daerah -> provinsi wilayah operasi utama
BPD_KEYS = [
    ("riau kepri", ["Riau", "Kepri"]), ("brk", ["Riau", "Kepri"]),
    ("mestika", ["Sumut"]), ("sumut", ["Sumut"]),
    ("aceh", ["Aceh"]), ("nagari", ["Sumbar"]), ("sumbar", ["Sumbar"]),
    ("riau", ["Riau", "Kepri"]), ("kepri", ["Kepri"]),
]

BULAN = ["", "Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]


# =========================================================
# 1. FUNGSI FORMAT
# =========================================================
def _id(s: str) -> str:
    """Konversi format angka ke gaya Indonesia (1.234,56)."""
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def format_triliun(nilai) -> str:
    if pd.isna(nilai):
        return "-"
    return "Rp" + _id(f"{nilai / 1000:,.2f}") + " T"


def format_rupiah(nilai) -> str:
    """Input Rp miliar. >= 1.000 M ditampilkan dalam triliun, selain itu miliar."""
    if pd.isna(nilai):
        return "-"
    if abs(nilai) >= 1000:
        return format_triliun(nilai)
    return "Rp" + _id(f"{nilai:,.1f}") + " M"


def format_ribu(nilai, satuan: str = "rb") -> str:
    if pd.isna(nilai):
        return "-"
    return _id(f"{nilai:,.1f}") + f" {satuan}"


def format_persen(nilai, desimal: int = 2) -> str:
    if pd.isna(nilai):
        return "-"
    return _id(f"{nilai * 100:.{desimal}f}") + "%"


def format_pp(nilai) -> str:
    """Selisih dalam percentage point, bertanda (+/-)."""
    if pd.isna(nilai):
        return "-"
    return _id(f"{nilai * 100:+.2f}") + " pp"


# =========================================================
# 2. LOAD & PARSING DATA
# =========================================================
def _num(v) -> float:
    """Konversi isi sel menjadi float. Menangani teks seperti '21,10%', '5.25%', '-', 'N/a'."""
    if v is None or isinstance(v, bool):
        return np.nan
    if isinstance(v, (datetime, date, pd.Timestamp)):
        return np.nan
    if isinstance(v, (int, float, np.number)):
        return float(v)
    s = str(v).strip().replace(" ", "")
    if s.lower() in ("", "-", "n/a", "na", "nan", "#n/a", "none"):
        return np.nan
    pct = s.endswith("%")
    s = s.rstrip("%")
    if "," in s and "." in s:          # 1.234,56
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:                      # 21,10
        s = s.replace(",", ".")
    try:
        x = float(s)
    except ValueError:
        return np.nan
    return x / 100 if pct else x


def _cell(df: pd.DataFrame, r: int, c: int):
    try:
        return _num(df.iat[r, c])
    except IndexError:
        return np.nan


def _raw(df: pd.DataFrame, r: int, c: int):
    try:
        v = df.iat[r, c]
    except IndexError:
        return None
    return None if (not isinstance(v, str) and pd.isna(v)) else v


def _txt(df: pd.DataFrame, r: int, c: int) -> str:
    v = _raw(df, r, c)
    return "" if v is None else str(v).strip()


def _norm(df: pd.DataFrame, r: int, c: int) -> str:
    return re.sub(r"\s+", " ", _txt(df, r, c).lower())


def _is_date(v) -> bool:
    return isinstance(v, (datetime, date, pd.Timestamp)) and not pd.isna(v)


def _is_label(v) -> bool:
    return isinstance(v, str) and v.strip() != "" and not _is_date(v)


def _periode(ts) -> str:
    if ts is None or pd.isna(ts):
        return "-"
    ts = pd.Timestamp(ts)
    return f"{BULAN[ts.month]} {ts.year}"


def _to_prov(text: str) -> str | None:
    t = str(text).lower()
    for key, val in PROV_KEYS:
        if key in t:
            return val
    return None


def _bpd_provs(text: str) -> list[str]:
    t = str(text).lower()
    for key, val in BPD_KEYS:
        if key in t:
            return val
    p = _to_prov(text)
    return [p] if p else []


def _nice_name(text: str) -> str:
    text = str(text).strip()
    if text.isupper():
        return " ".join(w if w in ("BPD", "BPR", "BRK", "BSI") else w.capitalize() for w in text.split())
    return text


def _blocks(raw: pd.DataFrame, row: int) -> list[tuple[int, int, str]]:
    """Daftar blok (kolom_awal, kolom_akhir, label) berdasarkan teks pada baris header."""
    starts = [(c, _txt(raw, row, c)) for c in range(raw.shape[1]) if _is_label(_raw(raw, row, c))]
    out = []
    for i, (c, lab) in enumerate(starts):
        end = starts[i + 1][0] if i + 1 < len(starts) else raw.shape[1]
        out.append((c, end, lab))
    return out


def _header_rows(raw: pd.DataFrame, matcher, min_hit: int = 3) -> list[int]:
    """Baris yang memuat >= min_hit nama entitas/provinsi (header blok)."""
    rows = []
    for r in range(raw.shape[0]):
        hits = sum(1 for c in range(raw.shape[1])
                   if _is_label(_raw(raw, r, c)) and matcher(_txt(raw, r, c)))
        if hits >= min_hit:
            rows.append(r)
    return rows


# --------------------- Bank Umum ---------------------
def _parse_umum(raw: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for prov in PROVINSI:
        k = KOLOM_UMUM[prov]
        rows.append({
            "Provinsi": prov,
            "Aset": _cell(raw, 3, k + 4), "YoY Aset": _cell(raw, 3, k + 6),
            "DPK": _cell(raw, 4, k + 4), "YoY DPK": _cell(raw, 4, k + 6),
            "Kredit": _cell(raw, 5, k + 4), "YoY Kredit": _cell(raw, 5, k + 6),
            "NPL Gross": _cell(raw, 6, k + 4), "NPL Net": _cell(raw, 7, k + 4),
            "LaR": _cell(raw, 8, k + 4), "LDR": _cell(raw, 9, k + 4),
        })
    return pd.DataFrame(rows)


def _parse_jenis(raw: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for prov in PROVINSI:
        k = KOLOM_JENIS[prov]
        for b in range(3, 6):
            rows.append({
                "Provinsi": prov,
                "Jenis Kredit": _txt(raw, b, k + 1),
                "Share": _cell(raw, b, k + 2),
                "YoY": _cell(raw, b, k + 3),
            })
    return pd.DataFrame(rows)


def _parse_skala(raw: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for prov in PROVINSI:
        k = KOLOM_SKALA[prov]
        for b in range(3, 6):
            rows.append({
                "Provinsi": prov,
                "Skala": _txt(raw, b, k + 1),
                "Share": _cell(raw, b, k + 2),
                "YoY": _cell(raw, b, k + 3),
                "NPL": _cell(raw, b, k + 4),
            })
    return pd.DataFrame(rows)


def _parse_sektor(raw: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for prov in PROVINSI:
        awal = BARIS_SEKTOR[prov]
        for i in range(5):
            b = awal + i
            rows.append({
                "Provinsi": prov,
                "Sektor": _txt(raw, b, 8),
                "NPL": _cell(raw, b, 9),
                "Share": _cell(raw, b, 10),
                "YoY": _cell(raw, b, 11),
            })
    return pd.DataFrame(rows)


# --------------------- BPD & BPR (deteksi otomatis berbasis label) ---------------------
# (nama kolom, alias label di Excel, jenis) — "nom" = nominal Rp miliar, "rasio" = fraksi
BANK_IND = [
    ("Aset", ["aset", "total aset"], "nom"),
    ("DPK", ["dpk"], "nom"),
    ("Kredit", ["kredit", "pembiayaan"], "nom"),
    ("Modal Inti", ["modal inti"], "nom"),
    ("CAR", ["car", "kpmm"], "rasio"),
    ("NPL Gross", ["npl gross", "npl", "npf", "npf gross"], "rasio"),
    ("NPL Net", ["npl net", "npl nett", "npf net", "npf nett"], "rasio"),
    ("LaR", ["lar", "loan at risk"], "rasio"),
    ("LDR", ["ldr", "fdr"], "rasio"),
    ("AL/DPK", ["al/dpk"], "rasio"),
    ("AL/NCD", ["al/ncd"], "rasio"),
]
BANK_COLS = [c for key, _, kind in BANK_IND
             for c in ([key, f"YoY {key}", f"Ytd {key}"] if kind == "nom" else [key, f"Δ {key}"])]


def _pick_periods(cols: list[tuple[int, pd.Timestamp]]):
    """Dari kolom bertanggal pilih: periode terkini, periode sama tahun lalu, Desember tahun lalu."""
    if not cols:
        return None, None, None
    c_cur, t_cur = max(cols, key=lambda x: x[1])
    c_prev = c_dec = None
    best = None
    for c, t in cols:
        lag = (t_cur - t).days
        if 300 <= lag <= 430 and (best is None or abs(lag - 365) < best):
            c_prev, best = c, abs(lag - 365)
        if t.month == 12 and t.year == t_cur.year - 1:
            c_dec = c
    return (c_cur, t_cur), c_prev, c_dec


def _parse_bank_lokal(raw: pd.DataFrame, mode: str, sumber: str):
    """Parser sheet 'Bank Imut' (mode='bpd') dan 'BPR-Kinerja Umum' (mode='bpr').
    Sheet memiliki dua tabel: atas (Cognos) dan bawah (koreksi manual)."""
    matcher = (lambda t: bool(_bpd_provs(t))) if mode == "bpd" else (lambda t: _to_prov(t) is not None)
    headers = _header_rows(raw, matcher)
    if not headers:
        raise ValueError("Header blok (nama bank/provinsi) tidak ditemukan.")
    h = headers[-1] if sumber == SUMBER_KOREKSI else headers[0]

    # Baris tanggal = baris pertama setelah header yang memuat tanggal
    dr = next((r for r in range(h + 1, min(h + 4, raw.shape[0]))
               if any(_is_date(_raw(raw, r, c)) for c in range(raw.shape[1]))), None)
    if dr is None:
        raise ValueError("Baris periode (tanggal) tidak ditemukan.")

    rows, t_any, t_prev_any = [], None, None
    for s, e, lab in _blocks(raw, h):
        if not matcher(lab):
            continue
        dcols = [(c, pd.Timestamp(_raw(raw, dr, c))) for c in range(s, e) if _is_date(_raw(raw, dr, c))]
        cur, c_prev, c_dec = _pick_periods(dcols)
        if cur is None:
            continue
        c_cur, t_cur = cur
        t_any = t_any or t_cur
        if c_prev is not None and t_prev_any is None:
            t_prev_any = pd.Timestamp(_raw(raw, dr, c_prev))

        rec = {}
        if mode == "bpd":
            rec["Bank"] = _nice_name(lab)
            rec["Provinsi"] = ", ".join(_bpd_provs(lab))
        else:
            rec["Provinsi"] = _to_prov(lab)

        started = False
        for r in range(dr + 1, min(dr + 25, raw.shape[0])):
            lbl = _norm(raw, r, s)
            if not lbl:
                if started:
                    break
                continue
            started = True
            hit = next(((k, kind) for k, al, kind in BANK_IND if lbl in al), None)
            if hit is None or hit[0] in rec:
                continue
            key, kind = hit
            v_cur = _cell(raw, r, c_cur)
            v_prev = _cell(raw, r, c_prev) if c_prev is not None else np.nan
            v_dec = _cell(raw, r, c_dec) if c_dec is not None else np.nan
            rec[key] = v_cur
            if kind == "nom":
                rec[f"YoY {key}"] = v_cur / v_prev - 1 if (pd.notna(v_prev) and v_prev) else np.nan
                rec[f"Ytd {key}"] = v_cur / v_dec - 1 if (pd.notna(v_dec) and v_dec) else np.nan
            else:
                rec[f"Δ {key}"] = v_cur - v_prev
        rows.append(rec)

    df = pd.DataFrame(rows)
    id_cols = ["Bank", "Provinsi"] if mode == "bpd" else ["Provinsi"]
    df = df.reindex(columns=id_cols + BANK_COLS)
    meta = {"periode": _periode(t_any), "pembanding": _periode(t_prev_any), "sumber": sumber}
    return df, meta


# --------------------- PVML ---------------------
def _parse_pvml(raw: pd.DataFrame):
    h = _header_rows(raw, lambda t: _to_prov(t) is not None)[0]
    rows = []
    for s, e, lab in _blocks(raw, h):
        prov = _to_prov(lab)
        if prov is None:
            continue
        # Baris judul kolom: memuat 'Rp Miliar'
        hr = next((r for r in range(h + 1, min(h + 6, raw.shape[0]))
                   if any("miliar" in _norm(raw, r, c) for c in range(s, e))), None)
        if hr is None:
            continue
        cols = {_norm(raw, hr, c): c for c in range(s, e) if _norm(raw, hr, c)}
        c_nom = next((c for t, c in cols.items() if "miliar" in t), None)
        c_yoy = next((c for t, c in cols.items() if "yoy" in t), None)
        c_npf = next((c for t, c in cols.items() if t.startswith(("npf", "twp", "npl"))), None)
        for r in range(hr + 1, raw.shape[0]):
            name = _raw(raw, r, s)
            if not _is_label(name):
                continue
            m = re.match(r"^(.*?)\s*\(([^)]*)\)\s*$", str(name).strip())
            ind, per = (m.group(1), m.group(2)) if m else (str(name).strip(), "")
            rows.append({
                "Provinsi": prov, "Industri": ind, "Periode": per,
                "Nominal": _cell(raw, r, c_nom) if c_nom is not None else np.nan,
                "YoY": _cell(raw, r, c_yoy) if c_yoy is not None else np.nan,
                "NPF": _cell(raw, r, c_npf) if c_npf is not None else np.nan,
            })
    df = pd.DataFrame(rows)
    # NPF identik di semua provinsi => kemungkinan angka nasional (mis. TWP90 fintech)
    nas = df.groupby("Industri")["NPF"].transform(
        lambda s: s.notna().sum() >= 3 and s.dropna().nunique() == 1)
    df["NPF Nasional"] = nas.astype(bool)
    pers = ", ".join(f"{i}: {p}" for i, p in df.drop_duplicates("Industri")[["Industri", "Periode"]].values if p)
    return df, {"periode": pers or "-"}


# --------------------- PPDP ---------------------
def _parse_ppdp(raw: pd.DataFrame):
    h = _header_rows(raw, lambda t: _to_prov(t) is not None)[0]
    rows = []
    for s, e, lab in _blocks(raw, h):
        prov = _to_prov(lab)
        if prov is None:
            continue
        lini = per = c_val = c_yoy = None
        for r in range(h + 1, raw.shape[0]):
            vals = [_raw(raw, r, c) for c in range(s, e)]
            if all(v is None for v in vals):              # baris kosong menutup seksi
                lini = None
                continue
            head = _raw(raw, r, s)
            nxt = _raw(raw, r + 1, s)
            if _is_label(head) and _is_date(nxt):          # judul seksi (lini usaha)
                lini, per = str(head).strip(), pd.Timestamp(nxt)
                hdr = {_norm(raw, r + 1, c): c for c in range(s, e) if _norm(raw, r + 1, c)}
                c_val = next((c for t, c in hdr.items() if "miliar" in t), s + 1)
                c_yoy = next((c for t, c in hdr.items() if "yoy" in t), s + 2)
                continue
            if lini is None or _is_date(head):
                continue
            metrik = str(head).strip() if _is_label(head) else "Nilai"
            v = _cell(raw, r, c_val)
            rows.append({"Provinsi": prov, "Lini": lini, "Metrik": metrik,
                         "Periode": _periode(per), "Nominal": v, "YoY": _cell(raw, r, c_yoy)})
    df = pd.DataFrame(rows)
    pers = ", ".join(f"{l}: {p}" for l, p in df.drop_duplicates("Lini")[["Lini", "Periode"]].values)
    return df, {"periode": pers or "-"}


# --------------------- Pasar Modal ---------------------
def _parse_pm(raw: pd.DataFrame):
    h = _header_rows(raw, lambda t: _to_prov(t) is not None)[0]
    rows, t_any = [], None
    for s, e, lab in _blocks(raw, h):
        prov = _to_prov(lab)
        if prov is None:
            continue
        tabel = per = c_val = c_yoy = None
        for r in range(h + 1, raw.shape[0]):
            head = _raw(raw, r, s)
            if _is_date(head):
                per = pd.Timestamp(head)
                t_any = t_any or per
                tabel = None
                continue
            if _norm(raw, r, s) == "no":                     # header sub-tabel
                hv = _norm(raw, r, s + 2)
                tabel = ("Investor (SID)" if "sid" in hv else
                         "Kepemilikan Saham" if "kep" in hv else
                         "Transaksi Saham" if ("nilai" in hv or "jual" in _norm(raw, r, s + 1)) else
                         _txt(raw, r, s + 1))
                c_val, c_yoy = s + 2, next((c for c in range(s, e) if "yoy" in _norm(raw, r, c)), s + 3)
                continue
            if tabel is None:
                continue
            if all(_raw(raw, r, c) is None for c in range(s, e)):
                tabel = None
                continue
            kat = "Total" if _norm(raw, r, s) == "total" else _txt(raw, r, s + 1)
            if not kat:
                continue
            rows.append({"Provinsi": prov, "Tabel": tabel, "Kategori": kat, "Periode": _periode(per),
                         "Nilai": _cell(raw, r, c_val), "YoY": _cell(raw, r, c_yoy)})
    df = pd.DataFrame(rows)
    # Lengkapi YoY 'Total' yang kosong dari komponennya (back-cast)
    for (pv, tb), g in df.groupby(["Provinsi", "Tabel"]):
        tot = g[g["Kategori"] == "Total"]
        if tot.empty or pd.notna(tot["YoY"].iloc[0]):
            continue
        comp = g[g["Kategori"] != "Total"].dropna(subset=["Nilai", "YoY"])
        prior = (comp["Nilai"] / (1 + comp["YoY"])).sum()
        if prior:
            df.loc[tot.index, "YoY"] = comp["Nilai"].sum() / prior - 1
    return df, {"periode": _periode(t_any)}


PARSER_LAIN = {
    "bpd": lambda raw, sumber: _parse_bank_lokal(raw, "bpd", sumber),
    "bpr": lambda raw, sumber: _parse_bank_lokal(raw, "bpr", sumber),
    "pvml": lambda raw, sumber: _parse_pvml(raw),
    "ppdp": lambda raw, sumber: _parse_ppdp(raw),
    "pm": lambda raw, sumber: _parse_pm(raw),
}


def load_workbook(file_bytes: bytes, sumber_bank: str = SUMBER_KOREKSI):
    """Baca seluruh sheet sekali saja.
    Return: (df_umum, df_jenis, df_skala, df_sektor, DATA, META, ERRORS)
    Sheet Bank Umum wajib; 5 sheet modul tambahan opsional (gagal parse -> modul nonaktif).
    (Caching dilakukan oleh pemanggil, mis. st.cache_data di aplikasi Streamlit.)"""
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None, engine="openpyxl")
    hilang = [s for s in SHEETS.values() if s not in raw]
    if hilang:
        raise KeyError(f"Sheet wajib tidak ditemukan: {', '.join(hilang)}")
    bu = (
        _parse_umum(raw[SHEETS["umum"]]),
        _parse_jenis(raw[SHEETS["jenis"]]),
        _parse_skala(raw[SHEETS["skala"]]),
        _parse_sektor(raw[SHEETS["sektor"]]),
    )
    data, meta, errors = {}, {}, {}
    for key, sheet in SHEETS_LAIN.items():
        data[key], meta[key] = None, {}
        if sheet not in raw:
            errors[key] = f"Sheet `{sheet}` tidak ditemukan pada file."
            continue
        try:
            df, m = PARSER_LAIN[key](raw[sheet], sumber_bank)
            if df.empty:
                raise ValueError("tidak ada baris data yang terbaca")
            data[key], meta[key] = df, m
        except Exception as e:  # noqa: BLE001
            errors[key] = f"Sheet `{sheet}` gagal dibaca: {e}"
    return (*bu, data, meta, errors)


# =========================================================
# 3. HELPER TURUNAN (dari dashboard)
# =========================================================
def agg_growth(nilai: pd.Series, yoy: pd.Series) -> float:
    """Pertumbuhan agregat dari beberapa komponen (back-cast nilai tahun lalu)."""
    m = nilai.notna() & yoy.notna() & (yoy > -1)
    if not m.any():
        return np.nan
    prior = (nilai[m] / (1 + yoy[m])).sum()
    return nilai[m].sum() / prior - 1 if prior else np.nan


def ppdp_frames(df: pd.DataFrame):
    """Pisahkan PPDP menjadi (asuransi: Premi/Klaim/Rasio per lini) dan (lainnya: Dapen, Penjaminan)."""
    ins_rows, lain_rows = [], []
    for (pv, lini), g in df.groupby(["Provinsi", "Lini"], sort=False):
        m = g.assign(_m=g["Metrik"].str.lower()).set_index("_m")
        if "premi" in m.index:
            pr, kl = m.loc["premi"], m.loc["klaim"] if "klaim" in m.index else None
            ins_rows.append({
                "Provinsi": pv, "Lini": lini, "Periode": pr["Periode"],
                "Premi": pr["Nominal"], "YoY Premi": pr["YoY"],
                "Klaim": kl["Nominal"] if kl is not None else np.nan,
                "YoY Klaim": kl["YoY"] if kl is not None else np.nan,
            })
        else:
            r = g.iloc[0]
            lain_rows.append({"Provinsi": pv, "Lini": lini, "Periode": r["Periode"],
                              "Nominal": r["Nominal"], "YoY": r["YoY"]})
    ins = pd.DataFrame(ins_rows)
    if not ins.empty:
        ins["Rasio Klaim"] = np.where(ins["Premi"] > 0, ins["Klaim"] / ins["Premi"], np.nan)
    return ins, pd.DataFrame(lain_rows)


def pm_wide(df: pd.DataFrame) -> pd.DataFrame:
    """Ringkasan Pasar Modal satu baris per provinsi."""
    def get(g, tabel_kw, kat_kw, col):
        x = g[g["Tabel"].str.lower().str.contains(tabel_kw) & g["Kategori"].str.lower().str.contains(kat_kw)]
        return x[col].iloc[0] if not x.empty else np.nan

    rows = []
    for pv, g in df.groupby("Provinsi", sort=False):
        rec = {"Provinsi": pv}
        for kat in ["Total", "Saham", "SBN", "Reksadana"]:
            rec[f"SID {kat}"] = get(g, "sid", kat.lower(), "Nilai")
            rec[f"YoY SID {kat}"] = get(g, "sid", kat.lower(), "YoY")
        for kat in ["Penjualan", "Pembelian"]:
            rec[kat] = get(g, "transaksi", kat.lower(), "Nilai")
            rec[f"YoY {kat}"] = get(g, "transaksi", kat.lower(), "YoY")
        for kat in ["Total", "Individu", "Institusi"]:
            rec[f"Kepemilikan {kat}"] = get(g, "kepemilikan", kat.lower(), "Nilai")
            rec[f"YoY Kepemilikan {kat}"] = get(g, "kepemilikan", kat.lower(), "YoY")
        rows.append(rec)
    w = pd.DataFrame(rows)
    w["Net Beli"] = w["Pembelian"] - w["Penjualan"]
    w["Porsi Institusi"] = w["Kepemilikan Institusi"] / w["Kepemilikan Total"]
    return w