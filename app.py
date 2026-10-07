# -*- coding: utf-8 -*-
"""
Infographic Generator OPS Sumbagut — ONE-PAGE DASHBOARD (2 Kolom)

Dua pilihan tata letak, sama-sama satu halaman & dua kolom utama:

  Layout A — Portrait (A4/A3)              Layout B — Landscape (A4/A3)
  ┌────────── HEADER ──────────┐           ┌──────────────────── HEADER ────────────────────┐
  │ A. Bank Umum │ B. Lainnya  │           │ A. Bank Umum            │ B. Lainnya            │
  │  01          │  05         │           │  01                     │  05                   │
  │  02          │  06         │           │  02        │  03        │  06        │  07      │
  │  03          │  07         │           │  04                     │  08        │  09      │
  │  04          │  08         │           └──────────────────── FOOTER ────────────────────┘
  │              │  09         │
  └────────── FOOTER ──────────┘

Kanvas desain (unit = px pada skala 1):
  Portrait  1240 × 1754  → A4 300 dpi = skala 2 (2480 × 3508), A3 300 dpi = skala 2,83
  Landscape 1754 × 1240  → A4 300 dpi = skala 2 (3508 × 2480), A3 300 dpi = skala 2,83
Ekspor PNG & PDF vektor.

Tema warna: lihat THEMES (dictionary) — header, judul section, kartu, border, aksen,
grafik, dekorasi & footer mengikuti tema. Tema "Template Merah/Biru (Ilustrasi)" mengikuti
template desain dan memakai file tema_merah_ilustrasi.png / tema_biru_ilustrasi.png
(simpan satu folder dengan app.py; bila tidak ada, tema tetap jalan tanpa ilustrasi). Warna STATUS (aman/waspada/merah) dan warna
arah perubahan (NPL/LaR naik = merah, turun = hijau) TIDAK ikut tema.

Data: hasil parsing riil dari backend dashboard (`ops_parser.py`, wajib satu folder).
Opsional: logo_ojk.png (latar transparan) → kiri atas header.
Ikon footer: Font Awesome Free 7.1 (https://fontawesome.com) — ikon CC BY 4.0,
path vektor tertanam di kode (tanpa file ikon/font tambahan).

Jalankan :  streamlit run app.py
"""
from __future__ import annotations

import io
import math
from dataclasses import astuple, dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path

import matplotlib

matplotlib.use("Agg")  # backend non-GUI untuk server Streamlit

import matplotlib.colors as mcolors  # noqa: E402
import matplotlib.font_manager  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
from matplotlib.backends.backend_agg import FigureCanvasAgg  # noqa: E402
from matplotlib.figure import Figure  # noqa: E402
from matplotlib.patches import Circle, FancyBboxPatch, PathPatch, Rectangle, RegularPolygon  # noqa: E402
from matplotlib.path import Path as MPath  # noqa: E402
from matplotlib.transforms import Affine2D  # noqa: E402

from ops_parser import (  # noqa: E402
    PROVINSI, SUMBER_COGNOS, SUMBER_KOREKSI, agg_growth, format_persen, format_pp,
    format_ribu, format_rupiah, format_triliun, load_workbook, pm_wide, ppdp_frames,
)

# =============================================================================
# A. KONFIGURASI
# =============================================================================
NAMA_PROV = {"Sumut": "Sumatera Utara", "Aceh": "Aceh", "Riau": "Riau",
             "Sumbar": "Sumatera Barat", "Kepri": "Kepulauan Riau"}
ORG_LABEL = "OTORITAS JASA KEUANGAN  •  KANTOR REGIONAL SUMATERA BAGIAN UTARA"
HEADER_CHIP = "BANK UMUM • BPD • BPR • PVML • PPDP • PASAR MODAL"

# --- Logo header --------------------------------------------------------------
# File logo (PNG disarankan, latar transparan). Dicari di folder kerja, lalu di folder app.py.
# Bila tidak ditemukan / gagal dibaca → logo dilewati, teks header kembali ke margin kiri.
LOGO_PATH = "logo_ojk.png"
LOGO_BOX = (230, 120)   # area maksimum logo (lebar, tinggi) dalam unit desain; rasio asli dipertahankan
LOGO_PANEL = True       # True = logo di atas panel putih membulat (kontras di header berwarna)
LOGO_PAD = 12           # padding panel di sekitar logo
LOGO_TEXT_GAP = 26      # jarak horizontal logo → teks header
LOGO_TRIM = True        # pangkas otomatis margin kosong (transparan/putih) di sekeliling logo

PVML_URUTAN = ["pembiayaan", "ventura", "gadai", "lkm", "mikro"]  # urutan tampil industri PVML

# --- Footer: kanal resmi OJK ------------------------------------------------------
# (ikon, label kecil, teks utama). Ikon: kunci FA_ICONS; tuple = beberapa ikon berdampingan.
FOOTER_CHANNELS = [
    (("globe",), "Website", "www.ojk.go.id"),
    (("instagram", "x"), "Instagram & X", "@ojkindonesia"),
    (("facebook",), "Facebook", "Otoritas Jasa Keuangan"),
    (("youtube",), "YouTube", "Otoritas Jasa Keuangan"),
    (("tiktok",), "TikTok", "@ojk_indonesia"),
    (("phone",), "Layanan Konsumen", "Kontak OJK 157"),
    (("instagram",), "OJK Sumatera Utara", "@ojk_sumut"),
]

# --- Tema warna -----------------------------------------------------------------
# Setiap tema WAJIB memiliki kunci berikut (tambah tema baru cukup menyalin satu blok):
#   primary      : warna utama (header, nomor section, aksen kartu, seri grafik pertama)
#   dark         : varian gelap primary (gradasi header)
#   accent       : aksen dekoratif (garis bawah header/section, waktu generate)
#   accent_light : teks terang di atas header/area gelap (nama provinsi, label)
#   ink          : teks utama & elemen gelap
#   page         : latar halaman
#   footer       : latar footer;  footer_muted: teks sekunder footer
#   banner_sub   : teks subjudul banner aplikasi Streamlit
#   chart        : 6 warna seri grafik (3 pertama dipakai donut/porsi)
STATUS = {  # warna status & arah perubahan — TIDAK dapat ditimpa tema
    "good": "#0F766E", "warn": "#B7791F", "bad": "#C81E3A", "neutral": "#9CA3AF",
}
BASE = {"white": "#FFFFFF", "soft": "#6B7280", "line": "#E5E7EB", "track": "#EDEFF2"}
THEMES = {
    "Default (Merah OJK)": dict(
        primary="#B71234", dark="#7A0C22", accent="#C2A24D", accent_light="#F1DFA8", ink="#1F2937",
        page="#F3F4F6", footer="#1F2937", footer_muted="#CBD5E1", banner_sub="#F3D3DA",
        chart=["#B71234", "#C2A24D", "#1F2937", "#9CA3AF", "#E5484D", "#5FA8A0"]),
    "Blue (Biru Korporat)": dict(
        primary="#1D4E89", dark="#0F2C52", accent="#D9A43B", accent_light="#F6DFA4", ink="#1B2638",
        page="#F1F4F8", footer="#13233B", footer_muted="#C3CFDF", banner_sub="#D2DEEE",
        chart=["#1D4E89", "#D9A43B", "#1B2638", "#8FA3BF", "#4F86C6", "#5FA8A0"]),
    "Green (Hijau Emerald)": dict(
        primary="#17694A", dark="#0B3D2B", accent="#C9A23F", accent_light="#F0E0A8", ink="#1C2A24",
        page="#F1F5F3", footer="#13261E", footer_muted="#C2D3CA", banner_sub="#CFE5DA",
        chart=["#17694A", "#C9A23F", "#1C2A24", "#8FA89C", "#3E9B73", "#6B8FB5"]),
    "Orange (Oranye Terakota)": dict(
        primary="#B5501B", dark="#6E2E0C", accent="#E3A33B", accent_light="#FFE2C6", ink="#2A211C",
        page="#F6F3F0", footer="#2A211C", footer_muted="#DCCFC6", banner_sub="#FBDCC8",
        chart=["#B5501B", "#2F4A63", "#E3A33B", "#9CA3AF", "#D9773D", "#5FA8A0"]),
    "Purple (Ungu Elegan)": dict(
        primary="#5B2A86", dark="#331650", accent="#C9A23F", accent_light="#F3E2A9", ink="#231C33",
        page="#F4F2F7", footer="#231C33", footer_muted="#D3CCE0", banner_sub="#E3D6F0",
        chart=["#5B2A86", "#C9A23F", "#231C33", "#A99BBF", "#8E5CC2", "#5FA8A0"]),
    "Slate (Abu Elegan)": dict(
        primary="#334155", dark="#1E293B", accent="#C2A24D", accent_light="#F1DFA8", ink="#111827",
        page="#F3F4F6", footer="#0F172A", footer_muted="#CBD5E1", banner_sub="#D7DEE8",
        chart=["#334155", "#C2A24D", "#8FA1B6", "#64748B", "#B4C0CE", "#5FA8A0"]),
}
# --- Tema berbasis template desain (gambar ilustrasi header) -------------------------
# Kunci tambahan "template":
#   image     : file ilustrasi (dicari di folder kerja, lalu folder app.py)
#   style     : "band"  → pita header berwarna primary, teks putih (template merah)
#               "light" → header putih, teks gelap, garis primary di bawah header (template biru)
#   ref_band  : tinggi header pada gambar template asli (px) — dasar skala ilustrasi
#   gap, line : tebal jarak putih & garis di bawah header, relatif terhadap tinggi header
# Ilustrasi diletakkan di pojok kanan atas dengan skala tinggi header, sehingga desain tetap
# proporsional di Layout A (portrait) maupun B (landscape). Bila file ilustrasi tidak ada,
# tema tetap berjalan dengan warna & gaya header yang sama (tanpa ilustrasi).
THEMES.update({
    "Template Merah (Ilustrasi)": dict(
        primary="#A62524", dark="#7E1A19", accent="#D4A64A", accent_light="#F8E3B0", ink="#1F2937",
        page="#FAFAFA", footer="#A62524", footer_muted="#F3C9C8", banner_sub="#F6D3D2",
        chart=["#A62524", "#D4A64A", "#1F2937", "#9CA3AF", "#E5484D", "#5FA8A0"],
        template=dict(image="tema_merah_ilustrasi.png", style="band", ref_band=503,
                      gap=24 / 503, line=27 / 503, footer_strip=False)),
    "Template Biru (Ilustrasi)": dict(
        primary="#1E589C", dark="#0E478A", accent="#E8961E", accent_light="#FFE2B0", ink="#16263D",
        page="#F3FFFF", footer="#1E589C", footer_muted="#CFE0F2", banner_sub="#D2E3F5",
        chart=["#1E589C", "#E8961E", "#16263D", "#8FA3BF", "#4F86C6", "#5FA8A0"],
        template=dict(image="tema_biru_ilustrasi.png", style="light", ref_band=681,
                      line=19 / 681, gap=36 / 681, footer_strip=False)),
})
DEFAULT_THEME = "Default (Merah OJK)"


def build_colors(theme: str) -> dict:
    """Gabungkan warna dasar + tema + status. STATUS diterapkan terakhir sehingga
    warna status/arah perubahan tidak pernah tertimpa oleh tema."""
    t = THEMES.get(theme, THEMES[DEFAULT_THEME])
    c = {**BASE, **{k: v for k, v in t.items() if k not in ("chart", "template")}, **STATUS}
    c["template"] = t.get("template")
    c["pal6"] = list(t["chart"])
    c["pal3"] = list(t["chart"][:3])
    return c


def on_color(col: str, c: dict) -> str:
    """Warna teks yang kontras di atas `col` (putih untuk latar gelap, ink untuk latar terang)."""
    r, g, b = mcolors.to_rgb(col)
    lum = 0.2126 * r + 0.7152 * g + 0.0722 * b
    return c["ink"] if lum > 0.55 else c["white"]


@dataclass(frozen=True)
class Batas:
    """Batas acuan rasio (fraksi) — default sama dengan Dashboard OPS Sumbagut."""
    npl: float = 0.05
    lar: float = 0.15
    ldr_min: float = 0.84
    ldr_max: float = 0.94
    car: float = 0.12
    bpr_npl: float = 0.10
    bpr_lar: float = 0.30
    bpr_ldr_min: float = 0.75
    bpr_ldr_max: float = 1.10
    npf: float = 0.05
    klaim: float = 0.80


# --- Geometri & tata letak --------------------------------------------------------
M = 36                 # margin luar kiri/kanan
GUTTER = 24            # jarak antarkolom utama
SUBGAP = 12            # jarak antarpanel berdampingan dalam satu baris
CAPTION_H = 34         # label "A. PERBANKAN" / "B. ..." di atas kolom
PANEL_GAP = 12         # jarak vertikal antarbaris panel
PANEL_HEAD = 58        # tinggi judul di dalam panel
PAD = 16               # padding dalam panel
EMPTY_W = 92           # bobot panel tanpa data (menyusut jadi strip tipis)


@dataclass(frozen=True)
class Geo:
    """Geometri satu tata letak. rows_left/rows_right: daftar baris, tiap baris = nomor segmen
    yang diletakkan berdampingan (lebar dibagi rata)."""
    key: str
    label: str
    W: int
    H: int
    header_h: int
    col_w_l: float
    rows_left: tuple
    rows_right: tuple

    @property
    def col_w_r(self) -> float:
        return self.W - 2 * M - GUTTER - self.col_w_l

    @property
    def x_left(self) -> float:
        return M

    @property
    def x_right(self) -> float:
        return M + self.col_w_l + GUTTER


LAYOUTS = {
    "A": Geo("A", "Layout A — Portrait 2 kolom", 1240, 1754, 160, 548,
             rows_left=((1,), (2,), (3,), (4,)), rows_right=((5,), (6,), (7,), (8,), (9,))),
    "B": Geo("B", "Layout B — Landscape 2 kolom", 1754, 1240, 140, (1754 - 2 * M - GUTTER) / 2,
             rows_left=((1,), (2, 3), (4,)), rows_right=((5,), (6, 7), (8, 9))),
}

# Ukuran output: (sisi pendek, sisi panjang) dalam piksel. Skala = sisi yang sesuai orientasi / lebar kanvas.
EXPORT_SIZES = {"A4 cetak 300 dpi": (2480, 3508), "A3 cetak 300 dpi": (3508, 4961), "Preview layar": None}


def export_scale(size_key: str, geo: Geo) -> float:
    dims = EXPORT_SIZES[size_key]
    if dims is None:
        return 1.0
    target_w = dims[0] if geo.W <= geo.H else dims[1]
    return target_w / geo.W

_FONT_PREF = ["Inter", "Segoe UI", "Roboto", "Arial", "Helvetica", "DejaVu Sans"]
_FONT_AVAIL = {f.name for f in matplotlib.font_manager.fontManager.ttflist}
plt.rcParams["font.family"] = [f for f in _FONT_PREF if f in _FONT_AVAIL] or ["DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["pdf.fonttype"] = 42  # font tertanam di PDF (teks dapat dicari/dipilih)


def _nan(v) -> bool:
    try:
        return v is None or pd.isna(v)
    except (TypeError, ValueError):
        return False


def _id(s: str) -> str:
    """Format angka gaya Indonesia (1.234,56)."""
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def format_jt_from_ribu(v_ribu) -> str:
    """Jumlah (dalam ribu) → juta, mis. 4.250 rb → '4,25 JT'."""
    if _nan(v_ribu):
        return "-"
    return _id(f"{v_ribu / 1000:,.2f}") + " JT"


def format_bps(delta) -> str:
    """Selisih rasio (fraksi) → basis poin bertanda, mis. 0,0069 → '+69 bps'."""
    if _nan(delta):
        return "-"
    v = round(delta * 10000)
    return ("+" if v > 0 else "") + _id(f"{v:,.0f}") + " bps"


def now_wib() -> datetime:
    return datetime.now(timezone.utc) + timedelta(hours=7)


# --- Status rasio terhadap batas ---------------------------------------------
def st_max(v, lim):
    """Semakin kecil semakin baik (NPL, LaR, NPF, rasio klaim)."""
    if _nan(v):
        return "neutral"
    return "bad" if v >= lim else ("warn" if v >= 0.8 * lim else "good")


def st_min(v, lim):
    """Semakin besar semakin baik (CAR)."""
    if _nan(v):
        return "neutral"
    return "bad" if v < lim else ("warn" if v < lim + 0.02 else "good")


def st_range(v, lo, hi):
    """Koridor (LDR)."""
    if _nan(v):
        return "neutral"
    return "good" if lo <= v <= hi else "warn"


STATUS_TXT = {"good": "Aman", "warn": "Waspada", "bad": "Di atas batas", "neutral": "n/a"}


def change_color(c: dict, delta, rule: str, base_col: str | None = None) -> str:
    """Warna perubahan rasio berdasarkan ARAH perubahan (bukan level):
       'up_bad'  : naik → merah, turun → hijau (NPL Gross, NPL Net, LaR)
       'up_good' : naik → hijau, turun → merah (mis. CAR)
       'neutral' : selalu base_col (LDR: sama dengan warna kotak indikatornya)
    Tidak berubah / tidak ada data → abu-abu netral."""
    if rule == "neutral":
        return base_col or c["soft"]
    if _nan(delta) or abs(delta) < 5e-6:
        return c["neutral"]
    up = delta > 0
    return c["bad"] if (up == (rule == "up_bad")) else c["good"]


# =============================================================================
# B. IKON (Font Awesome Free 7.1, CC BY 4.0) — path vektor tertanam
#    Format: kunci → (lebar viewBox, tinggi viewBox, "M x y L x y C x1 y1 x2 y2 x y Z ...")
# =============================================================================
FA_ICONS = {
    "globe": (512, 512, (
        "M 351.9 280 L 161 280 C 163.9 344.5 178.2 403.9 198.5 447.4 C 209.9 471.9 222.2 489.2 233.6 499.8 C "
        "244.8 510.3 252.5 512 256.5 512 C 260.5 512 268.2 510.3 279.4 499.8 C 290.8 489.2 303.1 471.8 314.5 "
        "447.4 C 334.8 403.9 349.1 344.5 352 280 Z M 160.9 232 L 351.8 232 C 349 167.5 334.7 108.1 314.4 64.6 "
        "C 303 40.2 290.7 22.8 279.3 12.2 C 268.1 1.7 260.4 0 256.4 0 C 252.4 0 244.7 1.7 233.5 12.2 C 222.1 "
        "22.8 209.8 40.2 198.4 64.6 C 178.1 108.1 163.8 167.5 160.9 232 M 112.9 232 C 116.4 146.4 138.5 66.9 "
        "170.8 14.7 C 78.7 47.3 10.9 131.2 1.5 232 Z M 1.5 280 C 10.9 380.8 78.7 464.7 170.8 497.3 C 138.5 "
        "445.1 116.4 365.6 112.9 280 Z M 399.9 280 C 396.4 365.6 374.3 445.1 342 497.3 C 434.1 464.6 501.9 "
        "380.8 511.3 280 Z M 511.3 232 C 501.9 131.2 434.1 47.3 342 14.7 C 374.3 66.9 396.4 146.4 399.9 232 Z "
    )),
    "instagram": (448, 512, (
        "M 224.3 141 C 183.2 140.9 145.2 162.7 124.6 198.2 C 103.9 233.8 103.8 277.6 124.3 313.2 C 144.7 "
        "348.9 182.6 370.9 223.7 371 C 264.8 371.1 302.8 349.3 323.4 313.8 C 344.1 278.2 344.2 234.4 323.7 "
        "198.8 C 303.3 163.1 265.4 141.1 224.3 141 M 223.7 181.4 C 250.4 181.3 275 195.4 288.5 218.4 C 301.9 "
        "241.5 302 269.9 288.8 293 C 275.5 316.2 251 330.5 224.3 330.6 C 183.1 330.8 149.6 297.5 149.4 256.3 "
        "C 149.2 215.1 182.5 181.6 223.7 181.4 M 317.1 136.3 C 317.1 121.5 329.1 109.5 343.9 109.5 C 358.7 "
        "109.5 370.7 121.5 370.7 136.3 C 370.7 151.1 358.7 163.1 343.9 163.1 C 329.1 163.1 317.1 151.1 317.1 "
        "136.3 M 446.8 163.5 C 445.1 127.6 436.9 95.8 410.6 69.6 C 384.4 43.4 352.6 35.2 316.7 33.4 C 279.7 "
        "31.3 168.8 31.3 131.8 33.4 C 96 35.1 64.2 43.3 37.9 69.5 C 11.6 95.7 3.5 127.5 1.7 163.4 C -0.4 "
        "200.4 -0.4 311.3 1.7 348.3 C 3.4 384.2 11.6 416 37.9 442.2 C 64.2 468.4 95.9 476.6 131.8 478.4 C "
        "168.8 480.5 279.7 480.5 316.7 478.4 C 352.6 476.7 384.4 468.5 410.6 442.2 C 436.8 416 445 384.2 "
        "446.8 348.3 C 448.9 311.3 448.9 200.5 446.8 163.5 M 399 388 C 391.2 407.6 376.1 422.7 356.4 430.6 C "
        "326.9 442.3 256.9 439.6 224.3 439.6 C 191.7 439.6 121.6 442.2 92.2 430.6 C 72.6 422.8 57.5 407.7 "
        "49.6 388 C 37.9 358.5 40.6 288.5 40.6 255.9 C 40.6 223.3 38 153.2 49.6 123.8 C 57.4 104.2 72.5 89.1 "
        "92.2 81.2 C 121.7 69.5 191.7 72.2 224.3 72.2 C 256.9 72.2 327 69.6 356.4 81.2 C 376 89 391.1 104.1 "
        "399 123.8 C 410.7 153.3 408 223.3 408 255.9 C 408 288.5 410.7 358.6 399 388 "
    )),
    "x": (448, 512, (
        "M 357.2 48 L 427.8 48 L 273.6 224.2 L 455 464 L 313 464 L 201.7 318.6 L 74.5 464 L 3.8 464 L 168.7 "
        "275.5 L -5.2 48 L 140.4 48 L 240.9 180.9 Z M 332.4 421.8 L 371.5 421.8 L 119.1 88 L 77.1 88 Z "
    )),
    "facebook": (512, 512, (
        "M 512 256 C 512 114.6 397.4 0 256 0 C 114.6 0 0 114.6 0 256 C 0 376 82.7 476.8 194.2 504.5 L 194.2 "
        "334.2 L 141.4 334.2 L 141.4 256 L 194.2 256 L 194.2 222.3 C 194.2 135.2 233.6 94.8 319.2 94.8 C "
        "335.4 94.8 363.4 98 374.9 101.2 L 374.9 172 C 368.9 171.4 358.4 171 345.3 171 C 303.3 171 287.1 "
        "186.9 287.1 228.2 L 287.1 256 L 370.7 256 L 356.3 334.2 L 287 334.2 L 287 510.1 C 413.8 494.8 512 "
        "386.9 512 256 "
    )),
    "youtube": (576, 512, (
        "M 549.7 124.1 C 543.5 100.4 524.9 81.8 501.4 75.5 C 458.9 64 288.1 64 288.1 64 C 288.1 64 117.3 64 "
        "74.7 75.5 C 51.2 81.8 32.7 100.4 26.4 124.1 C 15 167 15 256.4 15 256.4 C 15 256.4 15 345.8 26.4 "
        "388.7 C 32.7 412.3 51.2 430.2 74.7 436.5 C 117.3 448 288.1 448 288.1 448 C 288.1 448 458.9 448 501.5 "
        "436.5 C 525 430.2 543.5 412.3 549.8 388.7 C 561.2 345.8 561.2 256.4 561.2 256.4 C 561.2 256.4 561.2 "
        "167 549.8 124.1 Z M 232.2 337.6 L 232.2 175.2 L 374.9 256.4 Z "
    )),
    "tiktok": (448, 512, (
        "M 448.5 209.9 C 404.5 210 361.5 196.3 325.7 170.7 L 325.7 349.4 C 325.7 382.5 315.6 414.8 296.7 442 "
        "C 277.8 469.2 251.1 490 220.1 501.6 C 189.1 513.2 155.3 515.1 123.2 506.9 C 91.1 498.7 62.3 481 40.5 "
        "456.1 C 18.7 431.2 5.2 400.1 1.5 367.2 C -2.2 334.3 4.4 301.1 20.1 272 C 35.8 242.9 60.1 219.3 89.7 "
        "204.3 C 119.3 189.3 152.6 183.8 185.4 188.3 L 185.4 278.2 C 170.4 273.5 154.3 273.6 139.4 278.6 C "
        "124.5 283.6 111.5 293.2 102.4 305.9 C 93.3 318.6 88.4 334 88.5 349.8 C 88.6 365.6 93.7 380.8 103 "
        "393.5 C 112.3 406.2 125.4 415.6 140.4 420.4 C 155.4 425.2 171.5 425.2 186.4 420.3 C 201.3 415.4 "
        "214.4 405.9 223.6 393.2 C 232.8 380.5 237.8 365.1 237.8 349.4 L 237.8 0 L 325.8 0 C 325.7 7.4 326.4 "
        "14.9 327.7 22.2 C 330.8 38.5 337.1 54.1 346.4 67.9 C 355.7 81.7 367.7 93.5 381.6 102.5 C 401.5 115.6 "
        "424.8 122.6 448.6 122.6 L 448.6 210 Z "
    )),
    "phone": (512, 512, (
        "M 160.2 25 C 152.3 6.1 131.7 -3.9 112.1 1.4 L 106.6 2.9 C 42 20.5 -13.2 83.1 2.9 159.3 C 40 334.3 "
        "177.7 472 352.7 509.1 C 429 525.3 491.5 470 509.1 405.4 L 510.6 399.9 C 516 380.2 505.9 359.6 487.1 "
        "351.8 L 389.8 311.3 C 373.3 304.4 354.2 309.2 342.8 323.1 L 304.2 370.3 C 233.9 335.4 177.3 277 "
        "144.8 205.3 L 189 169.3 C 202.9 158 207.6 138.9 200.8 122.3 Z "
    )),
}

_ICON_PATHS: dict = {}


def _icon_path(name: str):
    if name not in _ICON_PATHS:
        vw, vh, s = FA_ICONS[name]
        verts, codes, tok, i = [], [], s.split(), 0
        while i < len(tok):
            c = tok[i]
            i += 1
            if c in ("M", "L"):
                verts.append((float(tok[i]), float(tok[i + 1])))
                codes.append(MPath.MOVETO if c == "M" else MPath.LINETO)
                i += 2
            elif c == "C":
                for _ in range(3):
                    verts.append((float(tok[i]), float(tok[i + 1])))
                    codes.append(MPath.CURVE4)
                    i += 2
            elif c == "Z":
                verts.append(verts[-1] if verts else (0.0, 0.0))
                codes.append(MPath.CLOSEPOLY)
        _ICON_PATHS[name] = (vw, vh, MPath(verts, codes))
    return _ICON_PATHS[name]


# =============================================================================
# C. KANVAS & KOMPONEN VISUAL
# =============================================================================
class Poster:
    """Kanvas berkoordinat unit-desain. Origin KIRI-ATAS, sumbu Y ke bawah.
    Semua elemen ditempatkan dalam kotak (x, y, w, h). Memakai Figure (bukan pyplot)
    agar aman dipakai bersamaan oleh beberapa sesi Streamlit."""

    def __init__(self, geo: Geo, colors: dict):
        self.geo, self.C = geo, colors
        self.W, self.H = geo.W, geo.H
        self.fig = Figure(figsize=(self.W / 100, self.H / 100), dpi=100)
        FigureCanvasAgg(self.fig)
        self.fig.patch.set_facecolor(colors["page"])
        self.ax = self.fig.add_axes([0, 0, 1, 1])
        self.ax.set_xlim(0, self.W)
        self.ax.set_ylim(self.H, 0)
        self.ax.axis("off")
        self.renderer = self.fig.canvas.get_renderer()

    # ---------- primitif ----------
    def sub_axes(self, x, y, w, h):
        a = self.fig.add_axes([x / self.W, (self.H - y - h) / self.H, w / self.W, h / self.H])
        a.set_facecolor("none")
        return a

    def text(self, x, y, s, size=9, color=None, weight="normal", ha="left", va="center", z=5, **kw):
        return self.ax.text(x, y, s, fontsize=size, color=color or self.C["ink"], weight=weight,
                            ha=ha, va=va, zorder=z, **kw)

    def width_of(self, s, size, weight="normal") -> float:
        t = self.ax.text(0, 0, s, fontsize=size, weight=weight)
        w = t.get_window_extent(self.renderer).width
        t.remove()
        return w

    def fit(self, s, size, max_w, weight="normal") -> str:
        s = str(s)
        if self.width_of(s, size, weight) <= max_w:
            return s
        while len(s) > 1 and self.width_of(s + "…", size, weight) > max_w:
            s = s[:-1]
        return s.rstrip() + "…"

    def fit_size(self, s, size, max_w, weight="normal", min_size=None) -> float:
        """Ukuran font terbesar (≤ size) agar teks muat di max_w (batas bawah min_size)."""
        min_size = min_size or size * 0.7
        while size > min_size and self.width_of(s, size, weight) > max_w:
            size -= 0.5
        return size

    def wrap(self, s, size, max_w, max_lines=2, weight="normal") -> str:
        words, lines, cur = str(s).split(), [], ""
        for w_ in words:
            trial = f"{cur} {w_}".strip()
            if self.width_of(trial, size, weight) <= max_w or not cur:
                cur = trial
            else:
                lines.append(cur)
                cur = w_
        lines.append(cur)
        if len(lines) > max_lines:
            lines = lines[:max_lines - 1] + [self.fit(" ".join(lines[max_lines - 1:]), size, max_w, weight)]
        return "\n".join(lines)

    def box(self, x, y, w, h, fc, ec="none", lw=0, r=8, z=1, alpha=1.0, shadow=False):
        if shadow:
            self.ax.add_patch(FancyBboxPatch((x + 1, y + 3), w, h, boxstyle=f"round,pad=0,rounding_size={r}",
                                             fc="#000000", ec="none", alpha=0.045, zorder=z - 0.1))
        p = FancyBboxPatch((x, y), w, h, boxstyle=f"round,pad=0,rounding_size={r}",
                           fc=fc, ec=ec, lw=lw, alpha=alpha, zorder=z)
        self.ax.add_patch(p)
        return p

    def rect(self, x, y, w, h, fc, z=1, alpha=1.0, clip=None):
        p = Rectangle((x, y), w, h, fc=fc, ec="none", zorder=z, alpha=alpha)
        self.ax.add_patch(p)
        if clip is not None:
            p.set_clip_path(clip)
        return p

    def card(self, x, y, w, h, accent=None, side="top", r=8, fc=None, z=2):
        c = self.box(x, y, w, h, fc=fc or self.C["white"], ec=self.C["line"], lw=0.8, r=r, z=z, shadow=True)
        if accent:
            if side == "top":
                self.rect(x, y, w, 3.5, accent, z=z + 0.5, clip=c)
            else:
                self.rect(x, y, 4.5, h, accent, z=z + 0.5, clip=c)
        return c

    def dot(self, x, y, color, r=3.2, z=6):
        self.ax.add_patch(Circle((x, y), r, fc=color, ec="none", zorder=z))

    def arrow(self, x, y, up, color, size=4.2, z=6):
        """Segitiga naik/turun (shape). up=None → garis datar."""
        if up is None:
            self.rect(x - size, y - 1, size * 2, 2, color, z=z)
            return
        self.ax.add_patch(RegularPolygon((x, y + (0.4 if up else -0.4)), 3, radius=size,
                                         orientation=np.pi if up else 0, fc=color, ec="none", zorder=z))

    def icon(self, name, cx, cy, size, color, z=7):
        """Gambar ikon vektor FA_ICONS berpusat di (cx, cy), muat dalam kotak size × size."""
        vw, vh, path = _icon_path(name)
        sc = size / max(vw, vh)
        tr = Affine2D().translate(-vw / 2, -vh / 2).scale(sc).translate(cx, cy) + self.ax.transData
        self.ax.add_patch(PathPatch(path, fc=color, ec="none", transform=tr, zorder=z))

    def pill(self, x, y, text, color, size=7.5, anchor="left", alpha=0.12, icon=None, tcolor=None):
        """Label pill. icon: None|'up'|'down'|'flat'|'dot'. Return x tepi kanan."""
        tw = self.width_of(text, size, "bold")
        pl = 19 if icon else 8
        bw, bh = tw + pl + 8, size * 1.39 + 8
        bx = x if anchor == "left" else (x - bw if anchor == "right" else x - bw / 2)
        self.box(bx, y - bh / 2, bw, bh, fc=color, r=bh / 2, z=4, alpha=alpha)
        tc = tcolor or color
        self.text(bx + pl, y, text, size=size, color=tc, weight="bold", z=6)
        if icon in ("up", "down"):
            self.arrow(bx + 11, y, icon == "up", tc, size=3.8)
        elif icon == "flat":
            self.arrow(bx + 11, y, None, tc, size=3.5)
        elif icon == "dot":
            self.dot(bx + 11, y, tc, r=3)
        return bx + bw

    def growth(self, x, y, g, size=7.5, anchor="left", invert=False, plain=False, suffix="YoY"):
        """Badge pertumbuhan (fraksi). invert=True bila kenaikan = memburuk."""
        C = self.C
        if _nan(g):
            return self.pill(x, y, f"{suffix} n/a", C["neutral"], size, anchor, icon="flat")
        up = g > 0
        flat = abs(g) < 5e-5
        color = C["neutral"] if flat else (C["good"] if (up != invert) else C["bad"])
        txt = f"{format_persen(abs(g))} {suffix}".strip()
        if plain:
            tw = self.width_of(txt, size, "bold")
            bx = x if anchor == "left" else (x - tw - 12 if anchor == "right" else x - (tw + 12) / 2)
            self.arrow(bx + 4, y, None if flat else up, color, size=3.6)
            self.text(bx + 12, y, txt, size=size, color=color, weight="bold")
            return bx + tw + 12
        return self.pill(x, y, txt, color, size, anchor, icon="flat" if flat else ("up" if up else "down"))

    def change_bps(self, x, y, delta, color, size=7.2, anchor="left", suffix="YoY"):
        """Teks perubahan rasio dalam bps dengan panah arah, mis. '▲ +69 bps YoY'.
        Warna ditentukan pemanggil (change_color). Return lebar yang dipakai."""
        if _nan(delta):
            txt, up = f"{suffix} n/a", None
        else:
            v = round(delta * 10000)
            txt, up = f"{format_bps(delta)} {suffix}".strip(), (None if v == 0 else v > 0)
        tw = self.width_of(txt, size, "bold")
        bx = x if anchor == "left" else (x - tw - 12 if anchor == "right" else x - (tw + 12) / 2)
        self.arrow(bx + 4, y, up, color, size=3.6)
        self.text(bx + 12, y, txt, size=size, color=color, weight="bold")
        return tw + 12

    def bar(self, x, y, w, h, frac, color, z=3):
        self.box(x, y, w, h, fc=self.C["track"], r=h / 2, z=z)
        if not _nan(frac) and frac > 0:
            self.box(x, y, max(w * min(frac, 1.0), h), h, fc=color, r=h / 2, z=z + 0.2)

    def gauge(self, x, y, w, v, lim, status, scale_min=1.6, h=6):
        """Bar nilai vs batas, dengan penanda garis batas."""
        scale = max(lim * scale_min, (0 if _nan(v) else v) * 1.12) or 1
        self.bar(x, y, w, h, np.nan if _nan(v) else v / scale, self.C[status])
        tx = x + w * lim / scale
        self.rect(tx - 0.8, y - 3, 1.6, h + 6, self.C["ink"], z=4)
        return tx


# ---------- komponen tingkat tinggi ----------
def nominal_tile(p: Poster, x, y, w, h, label, value, g, vsize=15):
    C = p.C
    p.card(x, y, w, h, accent=C["primary"])
    ls = p.fit_size(label, 7.5, w - 20, "bold", 6)
    p.text(x + 12, y + 20, p.fit(label, ls, w - 20, "bold"), size=ls, color=C["soft"], weight="bold")
    vs = p.fit_size(value, vsize, w - 20, "bold")
    p.text(x + 11, y + h * 0.52, p.fit(value, vs, w - 20, "bold"), size=vs, weight="bold")
    p.growth(x + 12, y + h - 17, g, size=7.2)


def ratio_tile(p: Poster, x, y, w, h, label, v, status, note, delta=None, delta_rule="up_bad", vsize=14,
               growth=None, growth_rule="up_bad"):
    """Kartu rasio.
    delta  : perubahan (fraksi) yang ditampilkan ringkas 'Δ+0,26' (pp) di baris catatan (BPR).
    growth : perubahan (fraksi) yang ditampilkan sebagai baris 'Pertumbuhan' dalam bps (Bank Umum).
    *_rule : aturan warna arah perubahan — lihat change_color()."""
    C = p.C
    col = C[status]
    p.card(x, y, w, h, accent=col, side="left")
    val_col = C["ink"] if status in ("good", "neutral") else col
    ls = p.fit_size(label, 7.5, w - 20, "bold", 6)
    p.text(x + 13, y + 18, p.fit(label, ls, w - 20, "bold"), size=ls, color=C["soft"], weight="bold")

    if growth is not None:  # ----- kartu dengan baris pertumbuhan (bps) -----
        gcol = change_color(C, growth, growth_rule, base_col=col)
        gtxt = f"{format_bps(growth)} YoY" if not _nan(growth) else "YoY n/a"
        gw = p.width_of(gtxt, 7.2, "bold") + 12
        vs = p.fit_size(format_persen(v), vsize, w - 24, "bold")
        vtw = p.width_of(format_persen(v), vs, "bold")
        ny = y + h - 14
        if vtw + gw + 34 <= w:          # lebar cukup → nilai & pertumbuhan sebaris
            vy = (y + 18 + ny) / 2 + 1
            p.text(x + 12, vy, format_persen(v), size=vs, weight="bold", color=val_col)
            p.change_bps(x + w - 9, vy, growth, gcol, size=7.2, anchor="right")
            p.text(x + 13, ny, p.fit(note, 6.6, w - 22), size=6.6, color=C["soft"])
        elif h >= 96:                    # sempit tetapi tinggi → bertumpuk
            p.text(x + 12, y + h * 0.40, format_persen(v), size=vs, weight="bold", color=val_col)
            p.change_bps(x + 13, y + h * 0.66, growth, gcol, size=p.fit_size(gtxt, 7.2, w - 30, "bold", 6))
            p.text(x + 13, ny, p.fit(note, 6.6, w - 22), size=6.6, color=C["soft"])
        else:                            # sempit & pendek → catatan batas diganti baris pertumbuhan
            p.text(x + 12, (y + 18 + ny) / 2 + 1, format_persen(v), size=vs, weight="bold", color=val_col)
            p.change_bps(x + 13, ny, growth, gcol, size=p.fit_size(gtxt, 7.2, w - 30, "bold", 6))
        return

    vs = p.fit_size(format_persen(v), vsize, w - 20, "bold", 8)
    has_d = delta is not None and not _nan(delta)
    dtxt = ("Δ" + format_pp(delta).replace(" pp", "")) if has_d else ""
    dw = p.width_of(dtxt, 6.6, "bold") + 6 if has_d else 0
    dc = change_color(C, delta, delta_rule, base_col=col) if has_d else None
    if w - 22 - dw >= 18 or not has_d:      # catatan & Δ sebaris
        p.text(x + 12, y + h * 0.5, format_persen(v), size=vs, weight="bold", color=val_col)
        p.text(x + 13, y + h - 15, p.fit(note, 6.6, w - 22 - dw), size=6.6, color=C["soft"])
        if has_d:
            p.text(x + w - 7, y + h - 15, dtxt, size=6.6, color=dc, weight="bold", ha="right")
    elif h >= 88:                            # kartu sempit tapi tinggi → catatan & Δ bertumpuk
        p.text(x + 12, y + h * 0.43, format_persen(v), size=vs, weight="bold", color=val_col)
        p.text(x + 13, y + h - 29, p.fit(note, 6.4, w - 20), size=6.4, color=C["soft"])
        p.text(x + 13, y + h - 14, dtxt, size=6.6, color=dc, weight="bold")
    else:                                    # sempit & pendek → hanya Δ
        p.text(x + 12, y + h * 0.5, format_persen(v), size=vs, weight="bold", color=val_col)
        p.text(x + 13, y + h - 15, dtxt, size=6.6, color=dc, weight="bold")


def empty_note(p: Poster, box, msg):
    x, y, w, h = box
    lines = max(1, min(3, int(h // 14)))
    p.text(x + w / 2, y + h / 2 - 2, p.wrap(f"Data tidak tersedia — {msg}", 8, w - 30, lines),
           size=8, color=p.C["soft"], style="italic", ha="center", linespacing=1.3)


def panel(p: Poster, x, y, w, h, no, title, sub):
    """Bingkai panel segmen. Return kotak konten (x, y, w, h)."""
    C = p.C
    p.card(x, y, w, h, r=10)
    p.ax.add_patch(Circle((x + PAD + 12, y + 26), 13, fc=C["primary"], ec="none", zorder=4))
    p.ax.add_patch(Circle((x + PAD + 12, y + 26), 15.5, fc="none", ec=C["accent"], lw=1.2, zorder=4))
    p.text(x + PAD + 12, y + 26.5, f"{no:02d}", size=8.5, color=C["white"], weight="bold", ha="center", z=5)
    p.text(x + PAD + 36, y + 19, p.fit(title, 11.5, w - PAD * 2 - 40, "bold"), size=11.5, weight="bold")
    p.text(x + PAD + 36, y + 37, p.fit(sub, 7.2, w - PAD * 2 - 40), size=7.2, color=C["soft"])
    p.rect(x + PAD, y + PANEL_HEAD - 5, w - 2 * PAD, 0.8, C["line"], z=3)
    p.rect(x + PAD, y + PANEL_HEAD - 6, 34, 2.4, C["accent"], z=3.2)
    return x + PAD, y + PANEL_HEAD + 6, w - 2 * PAD, h - PANEL_HEAD - 6 - PAD + 2


# =============================================================================
# D. DATA SATU PROVINSI
# =============================================================================
@dataclass
class ProvData:
    prov: str
    periode: str          # label periode opsional dari pengguna ("" = tidak ditampilkan)
    umum: pd.Series | None
    jenis: pd.DataFrame
    skala: pd.DataFrame
    sektor: pd.DataFrame
    bpd: pd.DataFrame | None
    bpr: pd.Series | None
    pvml: pd.DataFrame | None
    ins: pd.DataFrame | None
    lain: pd.DataFrame | None
    pm: pd.Series | None
    meta: dict
    errors: dict


def _first(df, prov):
    if df is None or df.empty:
        return None
    d = df[df["Provinsi"] == prov]
    return None if d.empty else d.iloc[0]


def collect(prov: str, periode: str, parsed) -> ProvData:
    """Filter seluruh hasil load_workbook() untuk satu provinsi."""
    df_umum, df_jenis, df_skala, df_sektor, DATA, META, ERRORS = parsed
    df_bpd, df_bpr, df_pvml, df_ppdp, df_pm = (DATA.get(k) for k in ("bpd", "bpr", "pvml", "ppdp", "pm"))

    bpd = None
    if df_bpd is not None:
        bpd = df_bpd[df_bpd["Provinsi"].astype(str).str.split(",").apply(
            lambda xs: prov in [s.strip() for s in xs])].reset_index(drop=True)

    pvml = None
    if df_pvml is not None:
        pvml = df_pvml[df_pvml["Provinsi"] == prov].copy()
        rank = lambda nm: next((i for i, k in enumerate(PVML_URUTAN) if k in str(nm).lower()), len(PVML_URUTAN))  # noqa: E731
        pvml = pvml.assign(_r=pvml["Industri"].map(rank)).sort_values("_r", kind="stable") \
                   .drop(columns="_r").reset_index(drop=True)

    ins = lain = None
    if df_ppdp is not None:
        ins_all, lain_all = ppdp_frames(df_ppdp)
        ins = ins_all[ins_all["Provinsi"] == prov].reset_index(drop=True) if not ins_all.empty else ins_all
        lain = lain_all[lain_all["Provinsi"] == prov].reset_index(drop=True) if not lain_all.empty else lain_all

    sektor = (df_sektor[df_sektor["Provinsi"] == prov].dropna(subset=["Share"]).query("Sektor != ''")
              .sort_values("Share", ascending=False).head(5).reset_index(drop=True))

    return ProvData(
        prov=prov, periode=(periode or "").strip(), umum=_first(df_umum, prov),
        jenis=df_jenis[df_jenis["Provinsi"] == prov].dropna(subset=["Share"]).reset_index(drop=True),
        skala=df_skala[df_skala["Provinsi"] == prov].dropna(subset=["Share"]).reset_index(drop=True),
        sektor=sektor, bpd=bpd, bpr=_first(df_bpr, prov), pvml=pvml, ins=ins, lain=lain,
        pm=_first(pm_wide(df_pm), prov) if df_pm is not None else None, meta=META, errors=ERRORS,
    )


def _get(row, key):
    """Ambil nilai kolom dari Series; NaN bila kolom tidak ada (mis. parser versi lama)."""
    try:
        return row[key] if key in row.index else np.nan
    except Exception:  # noqa: BLE001
        return np.nan


# =============================================================================
# E. SEGMEN — setiap draw_segment_N(p, box, d, b) menggambar di dalam kotak
#    konten (x, y, w, h) yang diberikan oleh layout grid.
# =============================================================================
# ---------- 01 Kinerja Umum Bank Umum ----------
def draw_segment_1(p: Poster, box, d: ProvData, b: Batas):
    x, y, w, h = box
    r = d.umum
    if r is None:
        return empty_note(p, box, "baris Bank Umum provinsi ini kosong.")
    nh, gap = round(h * 0.52), 10
    tw = (w - 2 * gap) / 3
    for i, (lab, key) in enumerate([("TOTAL ASET", "Aset"), ("DPK", "DPK"), ("TOTAL KREDIT", "Kredit")]):
        nominal_tile(p, x + i * (tw + gap), y, tw, nh, lab, format_triliun(r[key]), r[f"YoY {key}"])
    ry, rh = y + nh + gap, h - nh - gap
    rw = (w - 3 * gap) / 4
    # (label, nilai, status level, catatan batas, perubahan YoY, aturan warna perubahan)
    tiles = [
        ("NPL GROSS", r["NPL Gross"], st_max(r["NPL Gross"], b.npl), f"Batas {format_persen(b.npl, 0)}",
         _get(r, "Δ NPL Gross"), "up_bad"),
        ("NPL NET", r["NPL Net"], st_max(r["NPL Net"], b.npl), "Setelah CKPN",
         _get(r, "Δ NPL Net"), "up_bad"),
        ("LOAN AT RISK", r["LaR"], st_max(r["LaR"], b.lar), f"Waspada ≥{format_persen(b.lar, 0)}",
         _get(r, "Δ LaR"), "up_bad"),
        ("LDR", r["LDR"], st_range(r["LDR"], b.ldr_min, b.ldr_max),
         f"Koridor {format_persen(b.ldr_min, 0)}–{format_persen(b.ldr_max, 0)}",
         _get(r, "Δ LDR"), "neutral"),
    ]
    for i, (lab, v, s, note, g, rule) in enumerate(tiles):
        ratio_tile(p, x + i * (rw + gap), ry, rw, rh, lab, v, s, note, growth=g, growth_rule=rule)


# ---------- 02 Jenis Kredit ----------
def draw_segment_2(p: Poster, box, d: ProvData, b: Batas):
    C = p.C
    x, y, w, h = box
    dj = d.jenis
    if dj.empty:
        return empty_note(p, box, "sheet BU-Jenis Kredit kosong untuk provinsi ini.")
    tot = d.umum["Kredit"] if d.umum is not None else np.nan
    shares = dj["Share"].clip(lower=0).to_numpy()
    colors = [C["pal3"][i % 3] for i in range(len(dj))]

    size = min(h, w * 0.40 if w >= 480 else w * 0.44)
    ax = p.sub_axes(x, y + (h - size) / 2, size, size)
    wedges, _ = ax.pie(shares / shares.sum() if shares.sum() else shares, colors=colors, startangle=90,
                       counterclock=False, wedgeprops=dict(width=0.44, edgecolor=C["white"], linewidth=2.5))
    ax.set_aspect("equal")
    for w_, s, col in zip(wedges, shares, colors):
        if s >= 0.07:
            a = np.deg2rad((w_.theta1 + w_.theta2) / 2)
            ax.text(0.78 * np.cos(a), 0.78 * np.sin(a), format_persen(s, 0), ha="center", va="center",
                    fontsize=7.5, weight="bold", color=on_color(col, C))
    hole_w = size * 0.36                      # lebar aman di dalam lubang donut
    cfs = p.fit_size(format_triliun(tot), 9.5, hole_w, "bold", 6)
    ax.text(0, 0.11, format_triliun(tot), ha="center", va="center", fontsize=cfs, weight="bold", color=C["ink"])
    ax.text(0, -0.15, "Total Kredit", ha="center", va="center", fontsize=min(6.8, cfs * 0.75), color=C["soft"])

    x0, x1 = x + size + 22, x + w
    rh = h / max(len(dj), 3)
    for i, row in dj.iterrows():
        ry = y + i * rh + (rh - 58) / 2
        p.box(x0, ry + 6, 11, 11, fc=colors[i], r=2.5, z=3)
        nm = str(row["Jenis Kredit"]) or "-"
        nm = nm if nm.lower().startswith("kredit") else f"Kredit {nm}"
        nfs = p.fit_size(nm, 9.5, x1 - x0 - 66, "bold", 7.5)
        p.text(x0 + 18, ry + 12, p.fit(nm, nfs, x1 - x0 - 66, "bold"), size=nfs, weight="bold")
        p.text(x1, ry + 12, format_persen(row["Share"], 1), size=11, weight="bold", ha="right")
        est = row["Share"] * tot if not _nan(tot) else np.nan
        p.text(x0 + 18, ry + 33, f"≈ {format_triliun(est)}", size=9.5, weight="bold", color=C["soft"])
        p.growth(x1, ry + 33, row["YoY"], size=7.3, anchor="right", plain=True)
        p.bar(x0 + 18, ry + 47, x1 - x0 - 18, 5, row["Share"], colors[i])


# ---------- 03 Skala Kredit ----------
def draw_segment_3(p: Poster, box, d: ProvData, b: Batas):
    C = p.C
    x, y, w, h = box
    ds = d.skala
    if ds.empty:
        return empty_note(p, box, "sheet BU-Skala Kredit kosong untuk provinsi ini.")
    tot = d.umum["Kredit"] if d.umum is not None else np.nan
    colors = [C["pal3"][i % 3] for i in range(len(ds))]

    # Baris judul + TOTAL (seperti total pada Kredit per Jenis Penggunaan)
    p.text(x, y + 6, "PORSI KREDIT PER SKALA USAHA", size=7, color=C["soft"], weight="bold")
    tv = format_triliun(tot)
    p.text(x + w, y + 6, tv, size=8.5, weight="bold", ha="right")
    p.text(x + w - p.width_of(tv, 8.5, "bold") - 6, y + 6, "TOTAL KREDIT", size=7, color=C["soft"],
           weight="bold", ha="right")

    by, bh = y + 16, 30
    clip = p.box(x, by, w, bh, fc=C["track"], r=7, z=2)
    total, cx = ds["Share"].clip(lower=0).sum() or 1, x
    for (_, row), col in zip(ds.iterrows(), colors):
        sw = w * max(row["Share"], 0) / total
        p.rect(cx, by, sw, bh, col, z=2.5, clip=clip)
        tc = on_color(col, C)
        lab = f"{row['Skala']}  {format_persen(row['Share'], 1)}"
        if p.width_of(lab, 8, "bold") < sw - 12:
            p.text(cx + sw / 2, by + bh / 2, lab, size=8, weight="bold", color=tc, ha="center")
        elif p.width_of(format_persen(row["Share"], 0), 7.5, "bold") < sw - 6:
            p.text(cx + sw / 2, by + bh / 2, format_persen(row["Share"], 0), size=7.5, weight="bold",
                   color=tc, ha="center")
        cx += sw

    n, gap = len(ds), 10
    cw, cy = (w - (n - 1) * gap) / n, by + bh + 12
    ch = y + h - cy
    compact = ch < 175
    for i, row in ds.iterrows():
        cxx = x + i * (cw + gap)
        s = st_max(row["NPL"], b.npl)
        npl_col = C["ink"] if s in ("good", "neutral") else C[s]
        p.card(cxx, cy, cw, ch, accent=colors[i])
        p.box(cxx + 12, cy + 15, 9, 9, fc=colors[i], r=2, z=3)
        p.text(cxx + 26, cy + 20, p.fit(str(row["Skala"]).upper(), 8.5, cw - 34, "bold"), size=8.5, weight="bold")
        if not compact:
            p.text(cxx + 12, cy + 42, "NPL", size=7, color=C["soft"], weight="bold")
            p.text(cxx + 11, cy + 64, format_persen(row["NPL"]), size=16, weight="bold", color=npl_col)
            gy = cy + 88
        else:
            p.text(cxx + 12, cy + 42, "NPL", size=7, color=C["soft"], weight="bold")
            p.text(cxx + 34, cy + 42, format_persen(row["NPL"]), size=13.5, weight="bold", color=npl_col)
            gy = cy + 58
        tx = p.gauge(cxx + 12, gy, cw - 24, row["NPL"], b.npl, s)
        p.text(tx, gy + 16, f"batas {format_persen(b.npl, 0)}", size=6, color=C["soft"], ha="center")

        # Nominal kredit (≈ porsi × total kredit BU) + porsi tetap ditampilkan
        line_y = gy + 26
        p.rect(cxx + 12, line_y, cw - 24, 0.8, C["line"], z=3)
        bottom_y = cy + ch - 18
        my = line_y + (bottom_y - 10 - line_y) / 2
        est = row["Share"] * tot if not _nan(tot) else np.nan
        p.text(cxx + 12, my - 9, "NOMINAL", size=6.6, color=C["soft"], weight="bold")
        p.text(cxx + cw - 12, my - 9, format_persen(row["Share"], 1), size=7, color=C["soft"], weight="bold",
               ha="right")
        nv = f"≈ {format_triliun(est)}"
        ns = p.fit_size(nv, 11.5 if compact else 12.5, cw - 24, "bold", 8)
        p.text(cxx + 11, my + 9, p.fit(nv, ns, cw - 24, "bold"), size=ns, weight="bold")
        p.text(cxx + 12, bottom_y, "Kredit", size=7, color=C["soft"])
        p.growth(cxx + 42, bottom_y, row["YoY"], size=7)


# ---------- 04 Kredit per Sektor (Top 5) ----------
def draw_segment_4(p: Poster, box, d: ProvData, b: Batas):
    C = p.C
    x, y, w, h = box
    ds = d.sektor
    if ds.empty:
        return empty_note(p, box, "sheet Kredit BU Per Sektor kosong untuk provinsi ini.")
    c_name, c_npl = x + 40, x + w - 8
    c_bar, w_bar, c_yoy = x + w * 0.47, w * 0.22, x + w - 112
    name_w = c_bar - c_name - 14
    for t, xx, ha in [("LAPANGAN USAHA", c_name, "left"), ("PANGSA", c_bar, "left"),
                      ("YoY", c_yoy, "center"), ("NPL", c_npl - 26, "center")]:
        p.text(xx, y + 6, t, size=6.8, color=C["soft"], weight="bold", ha=ha)
    vmax = ds["Share"].max() or 1
    rh = (h - 18) / max(len(ds), 5)
    for i, row in ds.iterrows():
        ry = y + 18 + i * rh
        p.card(x, ry, w, rh - 6, r=7)
        mid = ry + (rh - 6) / 2
        p.dot(x + 18, mid, C["primary"] if i == 0 else C["ink"], r=10)
        p.text(x + 18, mid + 0.5, str(i + 1), size=8, color=C["white"], weight="bold", ha="center", z=7)
        p.text(c_name, mid, p.wrap(row["Sektor"], 8.2, name_w, 2, "bold"), size=8.2, weight="bold", linespacing=1.15)
        p.text(c_bar, mid - 7, format_persen(row["Share"]), size=9.5, weight="bold")
        p.bar(c_bar, mid + 5, w_bar, 5, row["Share"] / vmax, C["pal6"][0] if i == 0 else C["pal6"][1])
        p.growth(c_yoy, mid, row["YoY"], size=7.3, anchor="center", plain=True, suffix="")
        p.pill(c_npl, mid, format_persen(row["NPL"]), C[st_max(row["NPL"], b.npl)], size=7.5,
               anchor="right", icon="dot")


# ---------- 05 Bank Daerah (Bank Imut) ----------
def draw_segment_5(p: Poster, box, d: ProvData, b: Batas):
    C = p.C
    x, y, w, h = box
    if d.bpd is None:
        return empty_note(p, box, d.errors.get("bpd", "sheet Bank Imut tidak tersedia."))
    if d.bpd.empty:
        return empty_note(p, box, f"tidak ada bank daerah yang dipetakan ke {d.prov}.")
    n, gap = len(d.bpd), 8
    bh = min((h - (n - 1) * gap) / n, 128)
    y0 = y + (h - (n * bh + (n - 1) * gap)) / 2
    cw = (w - 16) / 6
    for i, r in d.bpd.iterrows():
        by = y0 + i * (bh + gap)
        c = p.card(x, by, w, bh, r=7)
        p.rect(x, by, w, 26, C["primary"], z=2.6, clip=c)
        p.rect(x, by + 26, w, 1.8, C["accent"], z=2.7, clip=c)
        p.text(x + 12, by + 13, p.fit(r["Bank"], 9.5, w - 160, "bold"), size=9.5, color=C["white"], weight="bold", z=6)
        p.pill(x + w - 8, by + 13, f"Wilayah: {r['Provinsi']}", C["white"], size=6.5, anchor="right",
               alpha=0.18, tcolor=C["white"])
        gy = by + 26 + (bh - 26) / 2
        # (label, kolom, status level, aturan warna Δ)
        cells = [("ASET", "Aset", None, None), ("DPK", "DPK", None, None), ("KREDIT", "Kredit", None, None),
                 ("CAR", "CAR", st_min(r["CAR"], b.car), "plain"),
                 ("NPL GROSS", "NPL Gross", st_max(r["NPL Gross"], b.npl), "up_bad"),
                 ("LDR", "LDR", st_range(r["LDR"], b.ldr_min, b.ldr_max), "plain")]
        for j, (lab, key, s, rule) in enumerate(cells):
            cx = x + 10 + j * cw
            if j == 3:
                p.rect(cx - 6, gy - 26, 0.8, 52, C["line"], z=3)
            p.text(cx, gy - 18, lab, size=6.6, color=C["soft"], weight="bold")
            if s is None:
                p.text(cx, gy + 1, p.fit(format_rupiah(r[key]), 10.5, cw - 6, "bold"), size=10.5, weight="bold")
                p.growth(cx, gy + 20, r[f"YoY {key}"], size=6.6, plain=True)
            else:
                p.text(cx, gy + 1, format_persen(r[key]), size=10.5, weight="bold",
                       color=C["ink"] if s in ("good", "neutral") else C[s])
                p.dot(cx + 3, gy + 20, C[s], r=2.6)
                dl = r[f"Δ {key}"]
                dcol = C["soft"] if rule == "plain" else change_color(C, dl, rule)
                p.text(cx + 10, gy + 20, ("Δ " + format_pp(dl)) if not _nan(dl) else STATUS_TXT[s],
                       size=6.4, color=dcol, weight="bold" if rule != "plain" and not _nan(dl) else "normal")


# ---------- 06 BPR Kinerja Umum ----------
def draw_segment_6(p: Poster, box, d: ProvData, b: Batas):
    x, y, w, h = box
    r = d.bpr
    if r is None:
        return empty_note(p, box, d.errors.get("bpr", f"data BPR {d.prov} tidak ditemukan."))
    gap = 8
    nh = round(h * 0.46)
    tw = (w - 2 * gap) / 3
    for i, key in enumerate(["Aset", "DPK", "Kredit"]):
        nominal_tile(p, x + i * (tw + gap), y, tw, nh, f"{key.upper()} BPR", format_rupiah(r[key]),
                     r[f"YoY {key}"], vsize=13)
    ry, rh = y + nh + gap, h - nh - gap
    rw = (w - 4 * gap) / 5
    tiles = [  # (label, kolom, status, catatan, aturan warna Δ)
        ("NPL GROSS", "NPL Gross", st_max(r["NPL Gross"], b.bpr_npl), f"≤{format_persen(b.bpr_npl, 0)}", "up_bad"),
        ("NPL NET", "NPL Net", st_max(r["NPL Net"], b.bpr_npl), "Net", "up_bad"),
        ("LaR", "LaR", st_max(r["LaR"], b.bpr_lar), f"≤{format_persen(b.bpr_lar, 0)}", "up_bad"),
        ("LDR", "LDR", st_range(r["LDR"], b.bpr_ldr_min, b.bpr_ldr_max),
         f"{format_persen(b.bpr_ldr_min, 0)}–{format_persen(b.bpr_ldr_max, 0)}", "neutral"),
        ("CAR", "CAR", st_min(r["CAR"], b.car), f"≥{format_persen(b.car, 0)}", "up_good"),
    ]
    for i, (lab, key, s, note, rule) in enumerate(tiles):
        ratio_tile(p, x + i * (rw + gap), ry, rw, rh, lab, r[key], s, note, delta=r[f"Δ {key}"],
                   delta_rule=rule, vsize=12)


# ---------- 07 PVML ----------
def draw_segment_7(p: Poster, box, d: ProvData, b: Batas):
    C = p.C
    x, y, w, h = box
    dv = d.pvml
    if dv is None or dv.empty:
        return empty_note(p, box, d.errors.get("pvml", f"data PVML {d.prov} tidak ditemukan."))
    c_name, c_npf = x + 12, x + w - 6
    c_val, w_bar = x + w * 0.36, w * 0.21
    name_w = c_val - c_name - 12

    def _npf_txt(row):
        if _nan(row["NPF"]):
            return "n/a"
        return f"{format_persen(row['NPF'])}{' nas.' if bool(row.get('NPF Nasional', False)) else ''}"
    pill_w = max(p.width_of(_npf_txt(r_), 7, "bold") + (35 if not _nan(r_["NPF"]) else 16)
                 for _, r_ in dv.iterrows())
    c_yoy = (c_val + w_bar + (c_npf - pill_w)) / 2       # tengah antara akhir bar & awal pill NPF
    for t, xx, ha in [("INDUSTRI", c_name, "left"), ("OUTSTANDING", c_val, "left"),
                      ("YoY", c_yoy, "center"), ("NPF", c_npf - 30, "center")]:
        p.text(xx, y + 5, t, size=6.8, color=C["soft"], weight="bold", ha=ha)
    tot_h = 30
    rh = min((h - 16 - tot_h - 6) / len(dv), 46)
    vmax = dv["Nominal"].max() if dv["Nominal"].notna().any() else 1
    two_line = rh >= 32
    for i, row in dv.iterrows():
        ry = y + 16 + i * rh
        col = C["pal6"][i % 6]
        p.card(x, ry, w, rh - 4, r=6, accent=col, side="left")
        mid = ry + (rh - 4) / 2
        if two_line:
            nfs = p.fit_size(row["Industri"], 8.3, name_w, "bold", 7)
            p.text(c_name, mid - 6, p.fit(row["Industri"], nfs, name_w, "bold"), size=nfs, weight="bold")
            p.text(c_name, mid + 8, f"Posisi {row['Periode']}" if row["Periode"] else "", size=6.3, color=C["soft"])
        else:
            p.text(c_name, mid, p.fit(f"{row['Industri']} ({row['Periode']})", 7.8, name_w, "bold"),
                   size=7.8, weight="bold")
        p.text(c_val, mid - 5, format_rupiah(row["Nominal"]), size=9, weight="bold")
        p.bar(c_val, mid + 6, w_bar, 4, (row["Nominal"] / vmax) if vmax else np.nan, col)
        p.growth(c_yoy, mid, row["YoY"], size=7.2, anchor="center", plain=True, suffix="")
        if _nan(row["NPF"]):
            p.pill(c_npf, mid, "n/a", C["neutral"], size=7, anchor="right")
        else:
            p.pill(c_npf, mid, _npf_txt(row), C[st_max(row["NPF"], b.npf)], size=7, anchor="right", icon="dot")
    ty = y + h - tot_h
    tot, g = dv["Nominal"].sum(), agg_growth(dv["Nominal"], dv["YoY"])
    p.box(x, ty, w, tot_h, fc=C["ink"], r=7, z=2)
    p.text(c_name, ty + tot_h / 2, "TOTAL PVML", size=8, color=C["accent_light"], weight="bold")
    p.text(c_val, ty + tot_h / 2, format_rupiah(tot), size=10, color=C["white"], weight="bold")
    if not _nan(g):
        p.pill(c_yoy, ty + tot_h / 2, f"{format_persen(abs(g))}", C["white"], size=7, anchor="center",
               icon="up" if g > 0 else "down", alpha=0.15, tcolor=C["white"])
    p.text(c_npf, ty + tot_h / 2, f"batas NPF {format_persen(b.npf, 0)}", size=6.5, color="#D1D5DB", ha="right")


# ---------- 08 PPDP ----------
def draw_segment_8(p: Poster, box, d: ProvData, b: Batas):
    C = p.C
    x, y, w, h = box
    if d.ins is None:
        return empty_note(p, box, d.errors.get("ppdp", "sheet PPDP tidak tersedia."))
    gap = 10
    cw = (w - gap) / 2
    n_ins = len(d.ins)
    n_lain = 0 if d.lain is None else len(d.lain)
    lain_rows = math.ceil(n_lain / 2)
    lh = 62 if lain_rows else 0
    ins_rows = math.ceil(n_ins / 2)
    ins_area = h - lain_rows * (lh + gap)
    ch = min((ins_area - (ins_rows - 1) * gap) / max(ins_rows, 1), 134)   # tinggi kartu dibatasi
    block_h = ins_rows * ch + (ins_rows - 1) * gap + lain_rows * (lh + gap)
    y = y + max(0, (h - block_h) / 2)                                      # isi diletakkan di tengah panel
    for i, r in d.ins.iterrows():
        cx, cy = x + (i % 2) * (cw + gap), y + (i // 2) * (ch + gap)
        p.card(cx, cy, cw, ch, accent=C["primary"] if i % 2 == 0 else C["accent"])
        ptxt = f"Posisi {r['Periode']}" if cw >= 240 else str(r["Periode"])
        pr = p.pill(cx + cw - 8, cy + 18, ptxt, C["soft"], size=6, anchor="right", alpha=0.1)
        pw_ = p.width_of(ptxt, 6, "bold") + 16
        lname = str(r["Lini"]).upper()
        lfs = p.fit_size(lname, 8.5, cw - 30 - pw_, "bold", 7)
        p.text(cx + 11, cy + 18, p.fit(lname, lfs, cw - 30 - pw_, "bold"), size=lfs, weight="bold")
        for j, (lab, v, g) in enumerate([("PREMI", r["Premi"], r["YoY Premi"]), ("KLAIM", r["Klaim"], r["YoY Klaim"])]):
            xx = cx + 11 + j * (cw / 2)
            p.text(xx, cy + 38, lab, size=6.6, color=C["soft"], weight="bold")
            vt = format_rupiah(v)
            vfs = p.fit_size(vt, 11, cw / 2 - 16, "bold", 8)
            p.text(xx, cy + 56, p.fit(vt, vfs, cw / 2 - 16, "bold"), size=vfs, weight="bold")
            p.growth(xx, cy + 75, g, size=6.6, plain=True, invert=(lab == "KLAIM"))
        s = st_max(r["Rasio Klaim"], b.klaim)
        p.text(cx + 11, cy + ch - 26, "Rasio klaim", size=6.8, color=C["soft"], weight="bold")
        p.text(cx + cw - 11, cy + ch - 26, format_persen(r["Rasio Klaim"], 1), size=8.5, weight="bold", ha="right",
               color=C["ink"] if s in ("good", "neutral") else C[s])
        p.gauge(cx + 11, cy + ch - 15, cw - 22, r["Rasio Klaim"], b.klaim, s, scale_min=1.4, h=5)
    ly = y + ins_rows * (ch + gap) if lain_rows else None
    for i, r in (d.lain.iterrows() if n_lain else []):
        tx, ty = x + (i % 2) * (cw + gap), ly + (i // 2) * (lh + gap)
        p.card(tx, ty, cw, lh, accent=C["ink"], side="left")
        lab = str(r["Lini"]).upper()
        g = r["YoY"]
        gtxt = "YoY n/a" if _nan(g) else f"{format_persen(abs(g))} YoY"
        gw = p.width_of(gtxt, 7, "bold") + 30
        vt = format_rupiah(r["Nominal"])
        if cw >= 240:   # lebar: label + periode sebaris, nilai + pertumbuhan sebaris
            ptxt = f"Posisi {r['Periode']}"
            ptw = p.width_of(ptxt, 6) + 12
            p.text(tx + 13, ty + 15, p.fit(lab, 7, cw - 24 - ptw, "bold"), size=7, color=C["soft"], weight="bold")
            p.text(tx + cw - 8, ty + 15, ptxt, size=6, color=C["soft"], ha="right")
            vfs = p.fit_size(vt, 13, cw - 24 - gw, "bold", 9)
            p.text(tx + 12, ty + 41, p.fit(vt, vfs, cw - 24 - gw, "bold"), size=vfs, weight="bold")
            p.growth(tx + cw - 8, ty + 41, g, size=7, anchor="right")
        else:           # sempit: label penuh, nilai + pertumbuhan, periode kecil di bawah
            lfs = p.fit_size(lab, 7, cw - 22, "bold", 6)
            p.text(tx + 13, ty + 13, p.fit(lab, lfs, cw - 22, "bold"), size=lfs, color=C["soft"], weight="bold")
            vfs = p.fit_size(vt, 12, cw - 24 - gw, "bold", 9)
            p.text(tx + 12, ty + 33, p.fit(vt, vfs, cw - 24 - gw, "bold"), size=vfs, weight="bold")
            p.growth(tx + cw - 8, ty + 33, g, size=6.8, anchor="right")
            p.text(tx + 13, ty + lh - 11, f"Posisi {r['Periode']}", size=6, color=C["soft"])


# ---------- 09 Pasar Modal ----------
def draw_segment_9(p: Poster, box, d: ProvData, b: Batas):
    C = p.C
    x, y, w, h = box
    r = d.pm
    if r is None:
        return empty_note(p, box, d.errors.get("pm", f"data Pasar Modal {d.prov} tidak ditemukan."))
    gap = 8
    th = round(h * 0.5)
    hw = w * 0.31
    p.box(x, y, hw, th, fc=C["primary"], r=8, z=2, shadow=True)
    p.ax.add_patch(Circle((x + hw - 14, y + 8), 50, fc=C["white"], alpha=0.07, ec="none", zorder=2.2))
    hl = "TOTAL INVESTOR (SID)"
    hfs = p.fit_size(hl, 7, hw - 22, "bold", 5.6)
    p.text(x + 12, y + 17, p.fit(hl, hfs, hw - 22, "bold"), size=hfs, color=C["accent_light"], weight="bold")
    jt = format_jt_from_ribu(r["SID Total"])             # langsung dalam satuan JT (juta)
    js = p.fit_size(jt, 17, hw - 20, "bold", 11)
    p.text(x + 11, y + th * 0.5, jt, size=js, color=C["white"], weight="bold")
    g = r["YoY SID Total"]
    if not _nan(g):
        p.pill(x + 11, y + th - 16, f"{format_persen(abs(g))} YoY", C["white"], size=6.8,
               icon="up" if g > 0 else "down", alpha=0.18, tcolor=C["white"])
    tw = (w - hw - 3 * gap) / 3
    for i, (kat, col) in enumerate(zip(["Saham", "SBN", "Reksadana"], C["pal3"])):
        tx = x + hw + gap + i * (tw + gap)
        p.card(tx, y, tw, th, accent=col)
        sl = f"SID {kat.upper()}"
        sfs = p.fit_size(sl, 7, tw - 16, "bold", 5.6)
        p.text(tx + 10, y + 18, p.fit(sl, sfs, tw - 16, "bold"), size=sfs, color=C["soft"], weight="bold")
        sv = format_ribu(r[f"SID {kat}"])
        vfs = p.fit_size(sv, 12.5, tw - 18, "bold", 8)
        p.text(tx + 9, y + th * 0.46, p.fit(sv, vfs, tw - 18, "bold"), size=vfs, weight="bold")
        tot = r["SID Total"]
        p.bar(tx + 10, y + th * 0.64, tw - 20, 4, r[f"SID {kat}"] / tot if (not _nan(tot) and tot) else np.nan, col)
        p.growth(tx + 10, y + th - 14, r[f"YoY SID {kat}"], size=6.6, plain=True)

    ty, tth = y + th + gap, h - th - gap
    p.card(x, ty, w, tth)
    p.text(x + 12, ty + 15, "TRANSAKSI SAHAM INVESTOR PROVINSI", size=7, color=C["soft"], weight="bold")
    vals = [("Pembelian", r["Pembelian"], r["YoY Pembelian"], C["pal6"][0]),
            ("Penjualan", r["Penjualan"], r["YoY Penjualan"], C["pal6"][1])]
    ok = [v for _, v, _, _ in vals if not _nan(v)]
    vmax = max(ok) if ok else 1
    net = r["Net Beli"]
    nb_w = min(112, w * 0.24)
    val_w, yoy_w = 74, 52
    bar_x = x + 74
    bar_w = max(30, w - 74 - val_w - yoy_w - nb_w - 22)
    for i, (lab, v, gg, col) in enumerate(vals):
        ry = ty + 36 + i * ((tth - 44) / 2)
        p.text(x + 12, ry, lab, size=8, weight="bold")
        p.bar(bar_x, ry - 5, bar_w, 10, (v / vmax) if not _nan(v) else np.nan, col)
        p.text(bar_x + bar_w + 8, ry, format_rupiah(v), size=8.5, weight="bold")
        p.growth(bar_x + bar_w + 8 + val_w, ry, gg, size=6.6, plain=True, suffix="")
    if not _nan(net):
        col = C["good"] if net >= 0 else C["bad"]
        bx = x + w - nb_w - 10
        p.box(bx, ty + 22, nb_w, tth - 32, fc=col, r=7, z=3, alpha=0.1)
        p.text(bx + nb_w / 2, ty + 22 + (tth - 32) * 0.32, "NET BELI" if net >= 0 else "NET JUAL", size=7,
               color=col, weight="bold", ha="center")
        p.text(bx + nb_w / 2, ty + 22 + (tth - 32) * 0.68, p.fit(format_rupiah(abs(net)), 10, nb_w - 10, "bold"),
               size=10, color=col, weight="bold", ha="center")


# ---------- Registri: nomor → (judul, subjudul(d), fungsi gambar) ----------
def _sub_umum(d: ProvData) -> str:
    base = "Aset, DPK, kredit & rasio utama • perubahan rasio dalam bps (YoY)"
    return f"{base} • Posisi {d.periode}" if d.periode else base


SEGMENTS = {
    1: ("Kinerja Umum Bank Umum", _sub_umum, draw_segment_1),
    2: ("Kredit per Jenis Penggunaan", lambda d: "Porsi & YoY • ≈ nominal = porsi × total kredit BU", draw_segment_2),
    3: ("Kredit per Skala Usaha", lambda d: "Porsi, ≈ nominal (porsi × total kredit BU) & NPL per skala usaha",
        draw_segment_3),
    4: ("Top 5 Kredit per Lapangan Usaha", lambda d: "Sektor dengan pangsa kredit terbesar beserta NPL",
        draw_segment_4),
    5: ("Bank Daerah (Bank Imut)", lambda d: f"Posisi {d.meta.get('bpd', {}).get('periode', '-')} vs "
                                             f"{d.meta.get('bpd', {}).get('pembanding', '-')} • Δ = perubahan YoY",
        draw_segment_5),
    6: ("BPR — Kinerja Umum", lambda d: f"Agregat BPR provinsi • Posisi {d.meta.get('bpr', {}).get('periode', '-')}"
                                        f" • Δ = perubahan YoY (pp)", draw_segment_6),
    7: ("PVML", lambda d: "Pembiayaan, Modal Ventura, Pergadaian, LKM & lainnya • outstanding & NPF", draw_segment_7),
    8: ("PPDP — Asuransi, Dapen & Penjaminan", lambda d: "Premi & klaim asuransi, investasi dapen, penjaminan",
        draw_segment_8),
    9: ("Pasar Modal", lambda d: f"Investor (SID) & transaksi saham • Posisi {d.meta.get('pm', {}).get('periode', '-')}",
        draw_segment_9),
}
LEFT_SEGMENTS = [(n, SEGMENTS[n][0]) for n in (1, 2, 3, 4)]
RIGHT_SEGMENTS = [(n, SEGMENTS[n][0]) for n in (5, 6, 7, 8, 9)]


# =============================================================================
# F. LOGO, HEADER, FOOTER
# =============================================================================
def logo_candidates() -> list[Path]:
    """Lokasi yang diperiksa untuk file logo (folder kerja, lalu folder app.py)."""
    cands = [Path(LOGO_PATH)]
    try:
        cands.append(Path(__file__).resolve().parent / LOGO_PATH)
    except NameError:  # __file__ tidak tersedia (mis. dijalankan interaktif)
        pass
    return cands


def logo_file() -> Path | None:
    """Path logo yang ditemukan, atau None."""
    return next((c for c in logo_candidates() if c.is_file()), None)


def asset_candidates(name: str) -> list[Path]:
    """Lokasi yang diperiksa untuk file aset (folder kerja, lalu folder app.py)."""
    cands = [Path(name)]
    try:
        cands.append(Path(__file__).resolve().parent / name)
    except NameError:
        pass
    return cands


def asset_file(name: str) -> Path | None:
    return next((c for c in asset_candidates(name) if c.is_file()), None)


def _sig(f: Path | None) -> str:
    if f is None:
        return "none"
    st_ = f.stat()
    return f"{f.resolve()}:{st_.st_mtime_ns}:{st_.st_size}"


def assets_signature(theme: str) -> str:
    """Kunci cache: logo + ilustrasi template tema terpilih (ganti file → render ulang)."""
    tpl = THEMES.get(theme, {}).get("template")
    return _sig(logo_file()) + "|" + (_sig(asset_file(tpl["image"])) if tpl else "-")


_TPL_CACHE: dict = {}


def load_template_image(name: str):
    """Baca ilustrasi template (RGBA/RGB float). None bila tidak ada / gagal dibaca."""
    try:
        f = asset_file(name)
        if f is None:
            return None
        key = _sig(f)
        if key not in _TPL_CACHE:
            _TPL_CACHE.clear()
            _TPL_CACHE[key] = plt.imread(str(f))
        return _TPL_CACHE[key]
    except Exception:  # noqa: BLE001 — ilustrasi bersifat opsional
        return None


def logo_signature() -> str:
    """Penanda versi file logo (path + waktu ubah + ukuran) untuk kunci cache Streamlit,
    agar menambah/mengganti logo langsung memicu render ulang."""
    f = logo_file()
    if f is None:
        return "none"
    st_ = f.stat()
    return f"{f.resolve()}:{st_.st_mtime_ns}:{st_.st_size}"


def _trim_logo(img):
    """Pangkas margin kosong: piksel transparan atau hampir putih di sekeliling logo."""
    rgb = img[..., :3] if img.ndim == 3 else np.repeat(img[..., None], 3, axis=2)
    if rgb.dtype.kind in "ui":            # JPG dibaca sebagai uint8 (0–255)
        rgb = rgb / 255.0
    mask = (rgb < 0.94).any(axis=2)        # bukan putih
    if img.ndim == 3 and img.shape[2] == 4:
        a = img[..., 3] / (255.0 if img.dtype.kind in "ui" else 1.0)
        mask &= a > 0.05                   # bukan transparan
    ys, xs = np.where(mask)
    if len(xs) == 0:
        return img
    m = 4                                  # sisakan sedikit ruang napas
    y0, y1 = max(ys.min() - m, 0), min(ys.max() + m + 1, img.shape[0])
    x0, x1 = max(xs.min() - m, 0), min(xs.max() + m + 1, img.shape[1])
    return img[y0:y1, x0:x1]


def _load_logo():
    """Baca file logo. Return array gambar, atau None bila file tidak ada / rusak / format
    tidak didukung — sehingga render header tidak pernah gagal karena logo."""
    try:
        path = logo_file()
        if path is None:
            return None
        img = plt.imread(str(path))
        if img.ndim not in (2, 3) or img.shape[0] == 0 or img.shape[1] == 0:
            return None
        return _trim_logo(img) if LOGO_TRIM else img
    except Exception:  # noqa: BLE001 — logo bersifat opsional
        return None


def draw_logo(p: Poster, panel_on: bool | None = None) -> float:
    """Gambar logo di pojok kiri atas header (proporsional, di tengah vertikal header).
    Return koordinat x tempat teks header dimulai (M bila logo tidak tersedia)."""
    panel_on = LOGO_PANEL if panel_on is None else panel_on
    img = _load_logo()
    if img is None:
        return M
    try:
        hh = p.geo.header_h
        ih, iw = img.shape[:2]
        box_w, box_h = LOGO_BOX[0], min(LOGO_BOX[1], hh - 20)
        pad = LOGO_PAD if panel_on else 0
        s = min((box_w - 2 * pad) / iw, (box_h - 2 * pad) / ih)   # skala tanpa distorsi
        lw, lh = iw * s, ih * s
        pw, ph = lw + 2 * pad, lh + 2 * pad
        px, py = M, (hh - ph) / 2
        if panel_on:
            p.box(px, py, pw, ph, fc=p.C["white"], r=12, z=2.5, shadow=True)
        ax = p.sub_axes(px + pad, py + pad, lw, lh)
        ax.imshow(img, cmap="gray" if img.ndim == 2 else None, interpolation="antialiased")
        ax.set_axis_off()
        return px + pw + LOGO_TEXT_GAP
    except Exception:  # noqa: BLE001 — gagal menggambar logo → teks tetap di margin kiri
        return M


def _header_texts(p: Poster, d: ProvData, tx: float, right_limit: float, colors: dict):
    """Teks header (label OJK, judul, provinsi, chip) mulai di x=tx, tidak melewati right_limit."""
    C, hh = p.C, p.geo.header_h
    text_w = right_limit - tx
    y_org, y_title, y_prov, y_chip = hh * 0.20, hh * 0.45, hh * 0.6875, hh * 0.88
    p.text(tx, y_org, p.fit(ORG_LABEL, 8.5, text_w, "bold"), size=8.5, color=colors["org"], weight="bold")
    tfs = p.fit_size("Kinerja Sektor Jasa Keuangan", 27, text_w, "bold", 20)
    p.text(tx, y_title, p.fit("Kinerja Sektor Jasa Keuangan", tfs, text_w, "bold"), size=tfs,
           color=colors["title"], weight="bold")
    p.text(tx, y_prov, p.fit(f"Provinsi {NAMA_PROV.get(d.prov, d.prov)}", 19, text_w, "bold"), size=19,
           color=colors["prov"], weight="bold")
    # Nama file Excel TIDAK ditampilkan. Label periode hanya bila diisi pengguna secara eksplisit.
    xr = tx
    if d.periode:
        xr = p.pill(tx, y_chip, f"PERIODE  {d.periode.upper()}", colors["chip1"], size=8.5, alpha=colors["chip_a"],
                    tcolor=colors["chip1"]) + 8
    if xr + p.width_of(HEADER_CHIP, 8.5, "bold") + 16 <= right_limit:
        p.pill(xr, y_chip, HEADER_CHIP, colors["chip2"], size=8.5, alpha=colors["chip_a"], tcolor=colors["chip2"])


def draw_header_template(p: Poster, d: ProvData, tpl: dict):
    """Header mengikuti template desain (pita/garis + ilustrasi di kanan atas).
    Return (y_bawah_header, kotak_ilustrasi | None)."""
    C, W, hh = p.C, p.W, p.geo.header_h
    if tpl["style"] == "band":                       # template merah: pita merah, celah putih, garis merah
        g, ln = tpl["gap"] * hh, tpl["line"] * hh
        p.rect(0, 0, W, hh, C["primary"], z=1)
        p.rect(0, hh, W, g, C["white"], z=1)
        p.rect(0, hh + g, W, ln, C["primary"], z=1)
        bottom = hh + g + ln
        colors = dict(org=C["accent_light"], title=C["white"], prov=C["accent_light"],
                      chip1=C["white"], chip2=C["accent_light"], chip_a=0.16)
        panel_on = True
    else:                                            # template biru: header putih, garis biru di bawah
        ln, g = tpl["line"] * hh, tpl["gap"] * hh
        p.rect(0, 0, W, hh, C["white"], z=1)
        p.rect(0, hh, W, ln, C["dark"], z=1)
        p.rect(0, hh + ln, W, g, C["white"], z=1)
        bottom = hh + ln + g
        colors = dict(org=C["primary"], title=C["ink"], prov=C["primary"],
                      chip1=C["primary"], chip2=C["primary"], chip_a=0.10)
        panel_on = False

    ill_box, right_limit = None, W - M
    img = load_template_image(tpl["image"])
    if img is not None:
        try:
            s = hh / tpl["ref_band"]
            ih, iw = img.shape[:2]
            iw_u, ih_u = iw * s, ih * s
            ax = p.sub_axes(W - iw_u, 0, iw_u, ih_u)
            ax.imshow(img, interpolation="antialiased", aspect="auto")
            ax.set_axis_off()
            ill_box = (W - iw_u, 0.0, iw_u, ih_u)
            right_limit = W - iw_u - 24
        except Exception:  # noqa: BLE001
            ill_box = None
    tx = draw_logo(p, panel_on=panel_on)
    _header_texts(p, d, tx, right_limit, colors)
    return bottom, ill_box


def draw_header(p: Poster, d: ProvData):
    """Header standar (tema warna) atau header template. Return (y_bawah_header, kotak_ilustrasi | None)."""
    tpl = p.C.get("template")
    if tpl:
        return draw_header_template(p, d, tpl)
    C, W, hh = p.C, p.W, p.geo.header_h
    hdr = p.rect(0, 0, W, hh, C["primary"], z=1)
    p.rect(0, hh * 0.55, W, hh * 0.45, C["dark"], z=1.1, alpha=0.35, clip=hdr)
    for cx, cy, r, a in [(W - 90, 10, 170, 0.07), (W - 200, hh + 5, 110, 0.06), (W - 10, hh * 0.75, 70, 0.09)]:
        c = Circle((cx, cy), r, fc=C["white"], ec="none", alpha=a, zorder=1.5)
        p.ax.add_patch(c)
        c.set_clip_path(hdr)
    p.rect(0, hh, W, 6, C["accent"], z=2)

    # Logo di kiri atas; semua teks header dimulai setelah logo (tx)
    tx = draw_logo(p)
    _header_texts(p, d, tx, W - M, dict(org=C["accent_light"], title=C["white"], prov=C["accent_light"],
                                        chip1=C["white"], chip2=C["accent_light"], chip_a=0.16))
    return hh + 6, None


def draw_column_caption(p: Poster, x, y, code, text, col_w):
    C = p.C
    col_w = max(col_w, 40)
    p.rect(x, y + 4, 4, 18, C["primary"], z=3)
    p.text(x + 12, y + 13, code, size=9.5, color=C["primary"], weight="bold")
    p.text(x + 12 + p.width_of(code, 9.5, "bold") + 6, y + 13, text, size=9.5, weight="bold")
    p.rect(x, y + CAPTION_H - 4, col_w, 1, C["line"], z=2)


# ---------- Footer: kanal resmi OJK + waktu generate ----------
FOOT_LABEL_FS, FOOT_VALUE_FS = 6.2, 7.6
FOOT_ICON_R, FOOT_ICON_GAP = 10.5, 4
FOOT_ROW_H = 46


def _footer_item_width(p: Poster, item) -> float:
    icons, label, value = item
    iw = len(icons) * 2 * FOOT_ICON_R + (len(icons) - 1) * FOOT_ICON_GAP
    tw = max(p.width_of(label, FOOT_LABEL_FS), p.width_of(value, FOOT_VALUE_FS, "bold"))
    return iw + 7 + tw


def footer_plan(p: Poster):
    """Susun kanal footer dalam 1 baris (atau 2 baris bila tidak muat). Return (rows, tinggi, lebar_waktu)."""
    gen_w = max(p.width_of("Waktu Generate", FOOT_LABEL_FS), p.width_of("00-00-0000  00:00 WIB", 8.5, "bold"))
    avail = p.W - 2 * M - gen_w - 34
    items = [(it, _footer_item_width(p, it)) for it in FOOTER_CHANNELS]
    min_gap = 18
    if sum(w for _, w in items) + min_gap * (len(items) - 1) <= avail:
        rows = [items]
    else:  # bagi dua baris seimbang
        tot, acc, cut = sum(w for _, w in items), 0, len(items)
        for i, (_, w) in enumerate(items):
            acc += w
            if acc >= tot / 2:
                cut = i + 1
                break
        rows = [items[:cut], items[cut:]]
    height = 18 + FOOT_ROW_H * len(rows)
    return rows, height, gen_w, avail


def draw_footer(p: Poster, plan):
    """Footer resmi: ikon & kanal media sosial OJK (kiri) + waktu generate (kanan)."""
    C, W, H = p.C, p.W, p.H
    rows, fh, gen_w, avail = plan
    y0 = H - fh
    if (C.get("template") or {}).get("footer_strip", True):
        p.rect(0, y0, W, 4, C["accent"], z=2)
    p.rect(0, y0 + 4, W, fh - 4, C["footer"], z=1)
    for ri, row in enumerate(rows):
        cy = y0 + 4 + 7 + FOOT_ROW_H * ri + FOOT_ROW_H / 2
        n = len(row)
        used = sum(w for _, w in row)
        gap = min(52, (avail - used) / (n - 1)) if n > 1 else 0
        x = M
        for k, ((icons, label, value), w) in enumerate(row):
            ix = x + FOOT_ICON_R
            for name in icons:
                p.ax.add_patch(Circle((ix, cy), FOOT_ICON_R, fc=C["white"], alpha=0.13, ec="none", zorder=3))
                p.icon(name, ix, cy, FOOT_ICON_R * 1.12, C["white"])
                ix += 2 * FOOT_ICON_R + FOOT_ICON_GAP
            tx = ix - FOOT_ICON_R - FOOT_ICON_GAP + 7
            p.text(tx, cy - 7, label, size=FOOT_LABEL_FS, color=C["footer_muted"])
            p.text(tx, cy + 6, value, size=FOOT_VALUE_FS, color=C["white"], weight="bold")
            x += w
            if k < n - 1:   # pemisah tipis antar kanal
                p.rect(x + gap / 2 - 0.4, cy - 11, 0.8, 22, C["white"], z=3, alpha=0.18)
                x += gap
    # Waktu generate (kanan, tengah vertikal footer)
    gy = y0 + 4 + (fh - 4) / 2
    gen = now_wib()
    p.text(W - M, gy - 8, "Waktu Generate", size=FOOT_LABEL_FS + 0.4, color=C["footer_muted"], ha="right")
    p.text(W - M, gy + 7, f"{gen:%d-%m-%Y  %H:%M} WIB", size=8.5, color=C["accent"], weight="bold", ha="right")


# =============================================================================
# G. LAYOUT GRID & RENDER
# =============================================================================
def panel_weights(d: ProvData) -> dict:
    """Bobot tinggi tiap panel (nomor segmen → bobot) untuk provinsi ini."""
    def has(df):
        return df is not None and not getattr(df, "empty", False)
    n_bpd = len(d.bpd) if has(d.bpd) else 0
    n_pv = len(d.pvml) if has(d.pvml) else 0
    return {
        1: 320 if d.umum is not None else EMPTY_W,
        2: 300 if has(d.jenis) else EMPTY_W,
        3: 330 if has(d.skala) else EMPTY_W,
        4: 436 if has(d.sektor) else EMPTY_W,
        5: 75 + 100 * min(n_bpd, 3) if n_bpd else EMPTY_W,
        6: 230 if d.bpr is not None else EMPTY_W,
        7: 110 + 38 * min(n_pv, 7) if n_pv else EMPTY_W,
        8: 270 if d.ins is not None else EMPTY_W,
        9: 262 if d.pm is not None else EMPTY_W,
    }


def layout_rows(top, bottom, x, w, rows, weights) -> list:
    """Bagi kolom [top, bottom] × [x, x+w] menjadi baris panel. Tinggi baris ∝ bobot terbesar
    panel di baris itu; panel dalam satu baris berbagi lebar rata. Return [(no, (x, y, w, h))]."""
    row_w = [max(weights[n] for n in row) for row in rows]
    avail = bottom - top - PANEL_GAP * (len(rows) - 1)
    out, y = [], top
    for row, rw in zip(rows, row_w):
        h = avail * rw / sum(row_w)
        pw = (w - SUBGAP * (len(row) - 1)) / len(row)
        for j, no in enumerate(row):
            out.append((no, (x + j * (pw + SUBGAP), y, pw, h)))
        y += h + PANEL_GAP
    return out


def render_infographic(d: ProvData, batas: Batas = Batas(), scale: float = 2.0, fmt: str = "png",
                       layout: str = "A", theme: str = DEFAULT_THEME) -> bytes:
    """Render one-page dashboard.
    layout : 'A' (portrait, layout lama) | 'B' (landscape 2 kolom)
    theme  : nama tema pada THEMES
    fmt    : 'png' (pakai scale) | 'pdf' (vektor)"""
    geo = LAYOUTS.get(layout, LAYOUTS["A"])
    p = Poster(geo, build_colors(theme))
    header_bottom, ill = draw_header(p, d)
    plan = footer_plan(p)

    cap_y = header_bottom + 12
    top, bottom = cap_y + CAPTION_H + 6, geo.H - plan[1] - 16

    def col_top_and_caption(x_col, col_w):
        """Bila ilustrasi template menjorok ke bawah header di atas kolom ini, panel kolom
        dimulai di bawah ilustrasi dan garis caption dipendekkan agar tidak menabrak."""
        if ill is None:
            return top, col_w
        ix, iy, iw, ih = ill
        if x_col + col_w <= ix or ih <= cap_y:
            return top, col_w
        return max(top, iy + ih + 10), (ix - 14 - x_col if ih > cap_y else col_w)

    t_left, cw_left = col_top_and_caption(geo.x_left, geo.col_w_l)
    t_right, cw_right = col_top_and_caption(geo.x_right, geo.col_w_r)
    draw_column_caption(p, geo.x_left, cap_y, "A.", "PERBANKAN — BANK UMUM", cw_left)
    draw_column_caption(p, geo.x_right, cap_y, "B.", "BANK DAERAH, BPR, IKNB & PASAR MODAL", cw_right)

    weights = panel_weights(d)
    for x_col, col_w, rows, col_top in [(geo.x_left, geo.col_w_l, geo.rows_left, t_left),
                                        (geo.x_right, geo.col_w_r, geo.rows_right, t_right)]:
        for no, (px, py, pw, ph) in layout_rows(col_top, bottom, x_col, col_w, rows, weights):
            title, sub_fn, draw_fn = SEGMENTS[no]
            try:
                sub = sub_fn(d)
            except Exception:  # noqa: BLE001
                sub = ""
            content = panel(p, px, py, pw, ph, no, title, sub)
            try:
                draw_fn(p, content, d, batas)
            except Exception as e:  # noqa: BLE001 — satu segmen gagal tidak menggagalkan halaman
                empty_note(p, content, f"segmen gagal digambar ({e})")

    draw_footer(p, plan)
    buf = io.BytesIO()
    p.fig.savefig(buf, format=fmt, dpi=100 * scale, facecolor=p.C["page"])
    return buf.getvalue()


# =============================================================================
# H. APLIKASI STREAMLIT
# =============================================================================
def main():
    import streamlit as st

    st.set_page_config(page_title="Infografis OPS Sumbagut", page_icon="🖼️", layout="wide")
    _ver = tuple(int(v) for v in st.__version__.split(".")[:2] if v.isdigit())
    stretch = {"width": "stretch"} if _ver >= (1, 50) else {"use_container_width": True}

    @st.cache_data(show_spinner="Membaca dan memproses file Excel…")
    def _load(file_bytes: bytes, sumber: str):
        return load_workbook(file_bytes, sumber)

    @st.cache_data(show_spinner=False)
    def _render(file_bytes: bytes, sumber: str, prov: str, periode: str, scale: float, batas_t: tuple, fmt: str,
                layout: str, theme: str, logo_sig: str):  # logo_sig hanya kunci cache
        d = collect(prov, periode, _load(file_bytes, sumber))
        return render_infographic(d, Batas(*batas_t), scale, fmt, layout, theme)

    with st.sidebar:
        st.subheader("1. Data")
        up = st.file_uploader("File Excel OPS (.xlsx / .xlsm)", type=["xlsx", "xlsm"])
        periode = st.text_input("Label periode di header (opsional)", value="",
                                placeholder="mis. Agustus 2026",
                                help="Kosongkan agar header bersih. Nama file Excel tidak pernah ditampilkan.")
        sumber = st.radio("Sumber tabel BPD/BPR", [SUMBER_KOREKSI, SUMBER_COGNOS],
                          help="Sama dengan pilihan di dashboard: tabel koreksi (bawah) atau Cognos (atas).")
        st.subheader("2. Provinsi")
        prov = st.selectbox("Provinsi", PROVINSI, format_func=lambda k: f"{k} — {NAMA_PROV[k]}")
        st.subheader("3. Tampilan")
        layout = st.radio("Tata letak", list(LAYOUTS), format_func=lambda k: LAYOUTS[k].label)
        theme = st.selectbox("Color theme", list(THEMES), index=list(THEMES).index(DEFAULT_THEME))
        tc = build_colors(theme)
        st.markdown("".join(
            f'<span style="display:inline-block;width:26px;height:14px;border-radius:3px;margin-right:4px;'
            f'background:{col}"></span>' for col in [tc["primary"], tc["dark"], tc["accent"], tc["ink"]]
            + tc["pal6"][3:5]), unsafe_allow_html=True)
        geo = LAYOUTS[layout]
        size_lbl = {k: f"{k} — {round(geo.W * export_scale(k, geo))} × {round(geo.H * export_scale(k, geo))} px"
                    for k in EXPORT_SIZES}
        res = st.radio("Ukuran output PNG", list(EXPORT_SIZES), format_func=lambda k: size_lbl[k])
        scale = export_scale(res, geo)
        tpl = THEMES[theme].get("template")
        if tpl:
            if load_template_image(tpl["image"]) is not None:
                st.caption(f"🎨 Ilustrasi template terdeteksi: `{asset_file(tpl['image']).resolve()}`")
            else:
                st.warning(f"File ilustrasi `{tpl['image']}` tidak ditemukan — tema tetap dipakai tanpa ilustrasi. "
                           "Simpan file tersebut satu folder dengan app.py.", icon="🎨")
        lf = logo_file()
        if lf is not None and _load_logo() is not None:
            st.caption(f"🖼️ Logo terdeteksi: `{lf.resolve()}`")
        elif lf is not None:
            st.warning(f"File logo ditemukan tetapi gagal dibaca: `{lf}`. Pastikan berformat PNG/JPG valid.", icon="🖼️")
        else:
            st.warning("Logo tidak ditemukan. Lokasi yang diperiksa:\n\n"
                       + "\n".join(f"- `{c.resolve()}`" for c in logo_candidates())
                       + "\n\nPeriksa nama file (huruf kecil, bukan `logo_ojk.png.png`).", icon="🖼️")
        with st.expander("⚙️ Batas acuan warna status"):
            st.caption("Default sama dengan Dashboard OPS Sumbagut.")
            npl = st.number_input("Batas NPL Bank Umum/BPD (%)", 1.0, 10.0, 5.0, 0.1)
            lar = st.number_input("Batas waspada LaR (%)", 5.0, 40.0, 15.0, 0.5)
            ldr = st.slider("Koridor LDR Bank Umum/BPD (%)", 50, 120, (84, 94))
            car = st.number_input("Batas aman CAR/KPMM (%)", 8.0, 20.0, 12.0, 0.5)
            bnpl = st.number_input("Batas NPL BPR (%)", 3.0, 20.0, 10.0, 0.5)
            blar = st.number_input("Batas LaR BPR (%)", 10.0, 50.0, 30.0, 1.0)
            bldr = st.slider("Koridor LDR BPR (%)", 50, 150, (75, 110))
            npf = st.number_input("Batas NPF PVML (%)", 1.0, 10.0, 5.0, 0.1)
            klaim = st.number_input("Batas rasio klaim (%)", 40, 120, 80, 5)
        batas = Batas(npl / 100, lar / 100, ldr[0] / 100, ldr[1] / 100, car / 100, bnpl / 100, blar / 100,
                      bldr[0] / 100, bldr[1] / 100, npf / 100, klaim / 100)

    st.markdown(
        f"""<div style="background:linear-gradient(115deg,{tc['primary']},{tc['dark']});padding:18px 24px;
        border-radius:14px;border-bottom:4px solid {tc['accent']};margin-bottom:16px">
        <div style="color:{tc['accent_light']};font-size:12px;font-weight:700;letter-spacing:1px">OJK • SUMBAGUT</div>
        <div style="color:white;font-size:25px;font-weight:800">Infographic Generator — One-Page Dashboard</div>
        <div style="color:{tc['banner_sub']};font-size:14px">Unggah Excel OPS → pilih provinsi, tata letak &amp; tema
        → unduh infografis 1 halaman</div></div>""",
        unsafe_allow_html=True,
    )

    if up is None:
        st.info("Unggah file Excel OPS pada sidebar untuk memulai. Format sheet sama dengan Dashboard OPS Sumbagut.",
                icon="⬅️")
        st.stop()

    file_bytes = up.getvalue()
    try:
        parsed = _load(file_bytes, sumber)
    except Exception as e:  # noqa: BLE001
        st.error(f"Gagal membaca file. Pastikan sheet Bank Umum tersedia dan formatnya sesuai.\n\nDetail: {e}")
        st.stop()
    if parsed[-1]:
        st.warning("Sebagian modul tidak dapat dimuat (panel terkait ditandai 'data tidak tersedia'): "
                   + "; ".join(parsed[-1].values()), icon="⚠️")

    per = (periode or "").strip()
    args = (file_bytes, sumber, prov, per)
    with st.spinner(f"Menyusun infografis {prov}…"):
        png = _render(*args, scale, astuple(batas), "png", layout, theme, assets_signature(theme))

    c1, c2 = st.columns([3, 2] if layout == "A" else [4, 2], gap="large")
    with c1:
        st.markdown(f"##### Preview — {geo.label}")
        st.image(png, width=620 if layout == "A" else 860)
    with c2:
        st.markdown("##### Unduh")
        base = f"Infografis_OPS_{prov}_{'Portrait' if layout == 'A' else 'Landscape'}"
        if per:
            base += "_" + per.replace(" ", "_")
        st.download_button("📥 Download PNG", data=png, file_name=f"{base}.png", mime="image/png",
                           type="primary", **stretch)
        st.caption(f"{size_lbl[res]} • {len(png) / 1e6:,.1f} MB")
        pdf = _render(*args, 1.0, astuple(batas), "pdf", layout, theme, assets_signature(theme))
        st.download_button("📄 Download PDF (vektor, siap cetak)", data=pdf, file_name=f"{base}.pdf",
                           mime="application/pdf", **stretch)
        st.markdown("**Kolom A — Bank Umum:**\n" + "\n".join(f"- {n:02d} {t}" for n, t in LEFT_SEGMENTS))
        st.markdown("**Kolom B — Lembaga lainnya:**\n" + "\n".join(f"- {n:02d} {t}" for n, t in RIGHT_SEGMENTS))

        umum = parsed[0]
        row = umum[umum["Provinsi"] == prov]
        dcols = [c for c in ["Δ NPL Gross", "Δ NPL Net", "Δ LaR", "Δ LDR"] if c in umum.columns]
        if not row.empty and dcols:
            with st.expander("🔎 Cek pertumbuhan rasio Bank Umum (bps YoY)"):
                st.dataframe(pd.DataFrame({
                    "Rasio": [c.replace("Δ ", "") for c in dcols],
                    "Posisi": [format_persen(row.iloc[0][c.replace("Δ ", "")]) for c in dcols],
                    "Perubahan YoY": [format_bps(row.iloc[0][c]) for c in dcols],
                }), hide_index=True, **stretch)
                st.caption("Dibaca dari kolom posisi tahun lalu (bila terdeteksi) atau kolom YoY baris rasio pada "
                           "sheet BU-Kinerja Umum. Bila satuannya keliru, atur BU_RASIO_YOY_MODE di ops_parser.py.")


if __name__ == "__main__":
    main()