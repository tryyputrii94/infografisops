# -*- coding: utf-8 -*-
"""
Infographic Generator OPS Sumbagut — Long Poster 9 Segmen (tema OJK)

Alur:
  1. Unggah file Excel OPS Bulanan (format sama dengan Dashboard OPS Sumbagut).
  2. Pilih SATU provinsi.
  3. Data diparsing dengan backend dashboard (modul `ops_parser.py`):
     df_umum, df_jenis, df_skala, df_sektor, df_bpd, df_bpr, df_pvml, PPDP, PM.
  4. Matplotlib merender SATU poster portrait (lebar 1080 px, tinggi dinamis
     ±5.000–6.500 px) berisi Header + 9 segmen + Footer, lalu diunduh sebagai PNG.

Jalankan :  streamlit run infografis_ops.py
Dependensi:  streamlit, pandas, numpy, matplotlib, openpyxl
File wajib di folder yang sama: ops_parser.py
Opsional  :  logo_ojk.png (latar transparan) → tampil di header.

Struktur kode:
  A. Konfigurasi tema & batas rasio
  B. Kelas Poster (kanvas koordinat piksel, sumbu Y ke bawah) + komponen visual
  C. Pengumpulan data satu provinsi (ProvData)
  D. draw_header, draw_segment_1 … draw_segment_9, draw_footer
  E. render_infographic (hitung tinggi → gambar → PNG)
  F. Aplikasi Streamlit
"""
from __future__ import annotations

import io
import os
from dataclasses import astuple, dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path

import matplotlib

matplotlib.use("Agg")  # backend non-GUI untuk server Streamlit

import matplotlib.font_manager  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
from matplotlib.patches import Circle, FancyBboxPatch, Rectangle, RegularPolygon  # noqa: E402

from ops_parser import (  # noqa: E402
    PROVINSI, SUMBER_COGNOS, SUMBER_KOREKSI, agg_growth, format_persen, format_pp,
    format_ribu, format_rupiah, format_triliun, load_workbook, pm_wide, ppdp_frames,
)

# =============================================================================
# A. KONFIGURASI
# =============================================================================
# Palet korporat OJK (selaras dengan Dashboard OPS Sumbagut)
C = {
    "red": "#B71234",        # Merah OJK
    "maroon": "#7A0C22",
    "gold": "#C2A24D",       # Emas OJK (aksen)
    "gold_light": "#F1DFA8",
    "gold_text": "#8A6D1F",  # emas gelap untuk teks di atas putih
    "ink": "#1F2937",        # teks utama
    "soft": "#6B7280",       # teks sekunder
    "line": "#E5E7EB",
    "band": "#F7F7F9",       # latar selang-seling segmen
    "track": "#EEF0F3",
    "white": "#FFFFFF",
    "good": "#0F766E",       # teal: aman / membaik
    "warn": "#B7791F",       # amber: waspada
    "bad": "#C81E3A",        # merah: melewati batas / memburuk
    "neutral": "#9CA3AF",
}
PAL3 = [C["red"], C["gold"], C["ink"]]
PAL6 = [C["red"], C["gold"], C["ink"], "#9CA3AF", "#E5484D", "#5FA8A0"]

NAMA_PROV = {"Sumut": "Sumatera Utara", "Aceh": "Aceh", "Riau": "Riau",
             "Sumbar": "Sumatera Barat", "Kepri": "Kepulauan Riau"}
ORG_LABEL = "OTORITAS JASA KEUANGAN  •  KANTOR REGIONAL SUMATERA BAGIAN UTARA"
LOGO_PATH = "logo_ojk.png"

# Urutan tampil industri PVML (kata kunci, huruf kecil). Industri lain ditaruh di akhir.
PVML_URUTAN = ["pembiayaan", "ventura", "gadai", "lkm", "mikro"]


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


W = 1080            # lebar poster (px pada skala 1x)
X0, CW = 60, 960    # margin kiri & lebar konten

_FONT_PREF = ["Inter", "Segoe UI", "Roboto", "Arial", "Helvetica", "DejaVu Sans"]
_FONT_AVAIL = {f.name for f in matplotlib.font_manager.fontManager.ttflist}
plt.rcParams["font.family"] = [f for f in _FONT_PREF if f in _FONT_AVAIL] or ["DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False


def _nan(v) -> bool:
    try:
        return v is None or pd.isna(v)
    except (TypeError, ValueError):
        return False


def now_wib() -> datetime:
    return datetime.now(timezone.utc) + timedelta(hours=7)


# --- Status rasio terhadap batas → warna -------------------------------------
def st_max(v, lim):
    """Rasio 'semakin kecil semakin baik' (NPL, LaR, NPF, rasio klaim)."""
    if _nan(v):
        return "neutral"
    return "bad" if v >= lim else ("warn" if v >= 0.8 * lim else "good")


def st_min(v, lim):
    """Rasio 'semakin besar semakin baik' (CAR)."""
    if _nan(v):
        return "neutral"
    return "bad" if v < lim else ("warn" if v < lim + 0.02 else "good")


def st_range(v, lo, hi):
    """Rasio dengan koridor (LDR)."""
    if _nan(v):
        return "neutral"
    return "good" if lo <= v <= hi else "warn"


# =============================================================================
# B. KANVAS POSTER & KOMPONEN VISUAL
# =============================================================================
class Poster:
    """Kanvas berkoordinat piksel. Origin di KIRI-ATAS, sumbu Y mengarah ke bawah,
    sehingga segmen dapat ditumpuk dengan 'kursor' y yang terus bertambah."""

    def __init__(self, height: int):
        self.H = height
        self.fig = plt.figure(figsize=(W / 100, height / 100), dpi=100)
        self.fig.patch.set_facecolor(C["white"])
        self.ax = self.fig.add_axes([0, 0, 1, 1])
        self.ax.set_xlim(0, W)
        self.ax.set_ylim(height, 0)   # sumbu Y terbalik
        self.ax.axis("off")
        self.renderer = self.fig.canvas.get_renderer()

    # ---------- primitif ----------
    def sub_axes(self, x, y, w, h):
        """Axes chart pada kotak piksel (x, y=atas, w, h)."""
        a = self.fig.add_axes([x / W, (self.H - y - h) / self.H, w / W, h / self.H])
        a.set_facecolor("none")
        return a

    def text(self, x, y, s, size=13, color=None, weight="normal", ha="left", va="center", z=5, **kw):
        return self.ax.text(x, y, s, fontsize=size, color=color or C["ink"], weight=weight,
                            ha=ha, va=va, zorder=z, **kw)

    def width_of(self, s, size, weight="normal") -> float:
        t = self.ax.text(0, 0, s, fontsize=size, weight=weight)
        w = t.get_window_extent(self.renderer).width
        t.remove()
        return w

    def fit(self, s, size, max_w, weight="normal") -> str:
        """Potong teks dengan elipsis agar muat di lebar max_w piksel."""
        s = str(s)
        if self.width_of(s, size, weight) <= max_w:
            return s
        while len(s) > 1 and self.width_of(s + "…", size, weight) > max_w:
            s = s[:-1]
        return s.rstrip() + "…"

    def wrap(self, s, size, max_w, max_lines=2, weight="normal") -> str:
        """Bungkus teks per kata berdasarkan lebar terukur (maks. max_lines baris)."""
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

    def box(self, x, y, w, h, fc, ec="none", lw=0, r=14, z=1, alpha=1.0, shadow=False):
        if shadow:
            self.ax.add_patch(FancyBboxPatch(
                (x + 2, y + 5), w, h, boxstyle=f"round,pad=0,rounding_size={r}",
                fc="#000000", ec="none", alpha=0.05, zorder=z - 0.1))
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

    def card(self, x, y, w, h, accent=None, accent_side="top", r=14, fc=None):
        """Kartu putih ber-border + bayangan, dengan garis aksen opsional."""
        c = self.box(x, y, w, h, fc=fc or C["white"], ec=C["line"], lw=1.1, r=r, z=2, shadow=True)
        if accent:
            if accent_side == "top":
                self.rect(x, y, w, 5, accent, z=2.5, clip=c)
            else:
                self.rect(x, y, 7, h, accent, z=2.5, clip=c)
        return c

    def arrow(self, x, y, up, color, size=6.5, z=6):
        """Segitiga naik/turun (shape, tidak bergantung font). up=None → garis datar."""
        if up is None:
            self.rect(x - size, y - 1.5, size * 2, 3, color, z=z)
            return
        # Sumbu Y terbalik: orientasi pi = menunjuk ke ATAS secara visual
        self.ax.add_patch(RegularPolygon((x, y + (0.5 if up else -0.5)), 3, radius=size,
                                         orientation=np.pi if up else 0,
                                         fc=color, ec="none", zorder=z))

    def pill(self, x, y, text, color, size=11, anchor="left", fill_alpha=0.12, icon=None,
             text_color=None, fc=None):
        """Label pill. icon: None | 'up' | 'down' | 'flat' | 'dot'. Return (x_kiri, x_kanan)."""
        tw = self.width_of(text, size, "bold")
        pad_l = 30 if icon else 12
        bw, bh = tw + pad_l + 12, size * 1.39 + 13
        bx = x if anchor == "left" else (x - bw if anchor == "right" else x - bw / 2)
        self.box(bx, y - bh / 2, bw, bh, fc=fc or color, r=bh / 2, z=4, alpha=fill_alpha)
        self.text(bx + pad_l, y, text, size=size, color=text_color or color, weight="bold", z=6)
        ic = bx + 17
        if icon == "up":
            self.arrow(ic, y, True, text_color or color, size=6)
        elif icon == "down":
            self.arrow(ic, y, False, text_color or color, size=6)
        elif icon == "flat":
            self.arrow(ic, y, None, text_color or color, size=5)
        elif icon == "dot":
            self.ax.add_patch(Circle((ic, y), 4.5, fc=text_color or color, ec="none", zorder=6))
        return bx, bx + bw

    def growth(self, x, y, g, size=11, anchor="left", invert=False, suffix="YoY", plain=False):
        """Badge pertumbuhan (g = fraksi). invert=True bila kenaikan berarti memburuk."""
        if _nan(g):
            return self.pill(x, y, f"{suffix} n/a", C["neutral"], size, anchor, icon="flat")
        up = g > 0
        good = (not up) if invert else up
        color = C["good"] if good else C["bad"]
        if abs(g) < 5e-5:
            color, icon = C["neutral"], "flat"
        else:
            icon = "up" if up else "down"
        txt = f"{format_persen(abs(g))} {suffix}".strip()
        if plain:  # versi tanpa latar (untuk baris tabel)
            tw = self.width_of(txt, size, "bold")
            bx = x if anchor == "left" else (x - tw - 18 if anchor == "right" else x - (tw + 18) / 2)
            self.arrow(bx + 6, y, None if icon == "flat" else up, color, size=5.5)
            self.text(bx + 18, y, txt, size=size, color=color, weight="bold")
            return bx, bx + tw + 18
        return self.pill(x, y, txt, color, size, anchor, icon=icon)

    def bar_track(self, x, y, w, h, frac, color, z=3):
        """Progress bar membulat (frac 0..1)."""
        self.box(x, y, w, h, fc=C["track"], r=h / 2, z=z)
        if not _nan(frac) and frac > 0:
            self.box(x, y, max(w * min(frac, 1.0), h), h, fc=color, r=h / 2, z=z + 0.2)


# ---------- komponen tingkat tinggi ----------
STATUS_TXT = {"good": "Aman", "warn": "Waspada", "bad": "Di atas batas", "neutral": "Info"}


def nominal_card(p: Poster, x, y, w, h, label, value, g, value_size=25):
    p.card(x, y, w, h, accent=C["red"])
    p.text(x + 22, y + 34, label, size=11.5, color=C["soft"], weight="bold")
    p.text(x + 21, y + 78, p.fit(value, value_size, w - 40, "bold"), size=value_size, weight="bold")
    p.growth(x + 22, y + h - 30, g, size=10.5)


def ratio_tile(p: Poster, x, y, w, h, label, v, status, note, delta=None, invert=True, value_size=22):
    """Tile rasio: garis status di kiri, nilai, keterangan batas, Δ pp opsional."""
    col = C[status]
    p.card(x, y, w, h, accent=col, accent_side="left")
    p.text(x + 22, y + 28, label, size=11, color=C["soft"], weight="bold")
    p.text(x + 21, y + 63, format_persen(v), size=value_size, weight="bold",
           color=C["ink"] if status in ("good", "neutral") else col)
    has_d = delta is not None and not _nan(delta)
    p.text(x + 22, y + h - 22, p.fit(note, 9.5, w - (90 if has_d else 34)), size=9.5, color=C["soft"])
    if has_d:  # Δ YoY (pp) di pojok kanan bawah
        worse = (delta > 0) if invert else (delta < 0)
        dc = C["neutral"] if abs(delta) < 5e-5 else (C["bad"] if worse else C["good"])
        p.text(x + w - 12, y + h - 22, f"{format_pp(delta)}", size=9, color=dc, weight="bold", ha="right")


def placeholder(p: Poster, y, msg):
    p.card(X0, y, CW, 80)
    p.text(X0 + 28, y + 40, p.fit(f"Data tidak tersedia — {msg}", 12, CW - 56),
           size=12, color=C["soft"], style="italic")


# =============================================================================
# C. DATA SATU PROVINSI
# =============================================================================
@dataclass
class ProvData:
    prov: str
    periode: str
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
    """Filter seluruh hasil parsing (load_workbook) untuk satu provinsi."""
    df_umum, df_jenis, df_skala, df_sektor, DATA, META, ERRORS = parsed
    df_bpd, df_bpr, df_pvml, df_ppdp, df_pm = (DATA.get(k) for k in ("bpd", "bpr", "pvml", "ppdp", "pm"))

    bpd = None
    if df_bpd is not None:
        bpd = df_bpd[df_bpd["Provinsi"].astype(str).str.split(",").apply(
            lambda xs: prov in [s.strip() for s in xs])].reset_index(drop=True)

    pvml = None
    if df_pvml is not None:
        pvml = df_pvml[df_pvml["Provinsi"] == prov].copy()

        def _rank(nm):
            t = str(nm).lower()
            return next((i for i, k in enumerate(PVML_URUTAN) if k in t), len(PVML_URUTAN))
        pvml["_r"] = pvml["Industri"].map(_rank)
        pvml = pvml.sort_values(["_r"], kind="stable").drop(columns="_r").reset_index(drop=True)

    ins = lain = None
    if df_ppdp is not None:
        ins_all, lain_all = ppdp_frames(df_ppdp)
        ins = ins_all[ins_all["Provinsi"] == prov].reset_index(drop=True) if not ins_all.empty else ins_all
        lain = lain_all[lain_all["Provinsi"] == prov].reset_index(drop=True) if not lain_all.empty else lain_all

    pm = _first(pm_wide(df_pm), prov) if df_pm is not None else None

    sektor = (df_sektor[df_sektor["Provinsi"] == prov].dropna(subset=["Share"])
              .query("Sektor != ''").sort_values("Share", ascending=False).head(5).reset_index(drop=True))

    return ProvData(
        prov=prov, periode=periode, umum=_first(df_umum, prov),
        jenis=df_jenis[df_jenis["Provinsi"] == prov].reset_index(drop=True),
        skala=df_skala[df_skala["Provinsi"] == prov].reset_index(drop=True),
        sektor=sektor, bpd=bpd, bpr=_first(df_bpr, prov), pvml=pvml, ins=ins, lain=lain, pm=pm,
        meta=META, errors=ERRORS,
    )


# =============================================================================
# D. SEGMEN-SEGMEN POSTER
#    Setiap segmen punya:  h_segment_N(d) -> tinggi isi (px)
#                          draw_segment_N(p, y, d, b) -> menggambar isi mulai y
# =============================================================================
HEADER_H = 330
SEG_HEAD_H = 92     # judul segmen
SEG_GAP = 64        # jarak antarsegmen
FOOTER_H = 190
MISSING_H = 80


def segment_heading(p: Poster, y, no, title, sub):
    """Judul segmen: lingkaran nomor + judul + subjudul + garis emas."""
    p.ax.add_patch(Circle((X0 + 28, y + 30), 28, fc=C["red"], ec="none", zorder=4))
    p.ax.add_patch(Circle((X0 + 28, y + 30), 33, fc="none", ec=C["gold"], lw=2, zorder=4))
    p.text(X0 + 28, y + 31, f"{no:02d}", size=15, color=C["white"], weight="bold", ha="center", z=5)
    p.text(X0 + 78, y + 18, title, size=21, weight="bold")
    p.text(X0 + 79, y + 50, p.fit(sub, 11.5, CW - 90), size=11.5, color=C["soft"])
    p.rect(X0 + 79, y + 72, 56, 4, C["gold"], z=3)


# ---------- SEGMEN 1: Kinerja Umum Bank Umum ----------
def h_segment_1(d):
    return 296 if d.umum is not None else MISSING_H


def draw_segment_1(p: Poster, y, d: ProvData, b: Batas):
    r = d.umum
    if r is None:
        return placeholder(p, y, "baris Bank Umum provinsi ini kosong.")
    cw3, gap = (CW - 2 * 20) / 3, 20
    for i, (lab, key) in enumerate([("TOTAL ASET", "Aset"), ("DANA PIHAK KETIGA (DPK)", "DPK"),
                                    ("TOTAL KREDIT", "Kredit")]):
        nominal_card(p, X0 + i * (cw3 + gap), y, cw3, 156, lab, format_triliun(r[key]), r[f"YoY {key}"])
    y2, cw4, g4 = y + 176, (CW - 3 * 16) / 4, 16
    tiles = [
        ("NPL GROSS", r["NPL Gross"], st_max(r["NPL Gross"], b.npl), f"Batas wajar {format_persen(b.npl)}"),
        ("NPL NET", r["NPL Net"], st_max(r["NPL Net"], b.npl), "Setelah dikurangi CKPN"),
        ("LOAN AT RISK", r["LaR"], st_max(r["LaR"], b.lar), f"Batas waspada {format_persen(b.lar)}"),
        ("LDR", r["LDR"], st_range(r["LDR"], b.ldr_min, b.ldr_max),
         f"Koridor {format_persen(b.ldr_min, 0)}–{format_persen(b.ldr_max, 0)}"),
    ]
    for i, (lab, v, s, note) in enumerate(tiles):
        ratio_tile(p, X0 + i * (cw4 + g4), y2, cw4, 120, lab, v, s, note)


# ---------- SEGMEN 2: Jenis Kredit ----------
def h_segment_2(d):
    return 360 if not d.jenis.dropna(subset=["Share"]).empty else MISSING_H


def draw_segment_2(p: Poster, y, d: ProvData, b: Batas):
    dj = d.jenis.dropna(subset=["Share"]).reset_index(drop=True)
    if dj.empty:
        return placeholder(p, y, "sheet BU-Jenis Kredit kosong untuk provinsi ini.")
    p.card(X0, y, CW, 340)
    tot_kredit = d.umum["Kredit"] if d.umum is not None else np.nan
    shares = dj["Share"].clip(lower=0).to_numpy()
    s_norm = shares / shares.sum() if shares.sum() else shares
    colors = [PAL3[i % 3] for i in range(len(dj))]

    ax = p.sub_axes(X0 + 24, y + 22, 300, 300)
    wedges, _ = ax.pie(s_norm, colors=colors, startangle=90, counterclock=False,
                       wedgeprops=dict(width=0.46, edgecolor=C["white"], linewidth=4))
    ax.set_aspect("equal")
    for w_, s, col in zip(wedges, shares, colors):
        if s >= 0.06:
            ang = np.deg2rad((w_.theta1 + w_.theta2) / 2)
            ax.text(0.77 * np.cos(ang), 0.77 * np.sin(ang), format_persen(s, 1), ha="center",
                    va="center", fontsize=10, weight="bold",
                    color=C["ink"] if col == C["gold"] else C["white"])
    ax.text(0, 0.10, format_triliun(tot_kredit), ha="center", va="center", fontsize=12.5,
            weight="bold", color=C["ink"])
    ax.text(0, -0.14, "Total Kredit BU", ha="center", va="center", fontsize=9.5, color=C["soft"])

    x0, x1 = X0 + 360, X0 + CW - 30
    row_h = 290 / max(len(dj), 3)
    for i, row in dj.iterrows():
        ry = y + 40 + i * row_h
        col = colors[i]
        p.box(x0, ry - 9, 18, 18, fc=col, r=4, z=3)
        nama = str(row["Jenis Kredit"]) or "-"
        nama = nama if nama.lower().startswith("kredit") else f"Kredit {nama}"
        p.text(x0 + 30, ry, p.fit(nama, 14.5, 330, "bold"), size=14.5, weight="bold")
        p.text(x1, ry, format_persen(row["Share"], 1), size=17, weight="bold", ha="right")
        est = row["Share"] * tot_kredit if not _nan(tot_kredit) else np.nan
        p.text(x0 + 30, ry + 32, f"≈ {format_triliun(est)}", size=15, weight="bold", color=C["soft"])
        p.growth(x1, ry + 32, row["YoY"], size=10.5, anchor="right", plain=True)
        p.bar_track(x0 + 30, ry + 54, x1 - x0 - 30, 8, row["Share"], col)
    p.text(x0, y + 326, "≈ nominal estimasi = porsi × total kredit Bank Umum provinsi",
           size=9, color=C["soft"], style="italic", va="bottom")


# ---------- SEGMEN 3: Skala Kredit ----------
def h_segment_3(d):
    return 300 if not d.skala.dropna(subset=["Share"]).empty else MISSING_H


def draw_segment_3(p: Poster, y, d: ProvData, b: Batas):
    ds = d.skala.dropna(subset=["Share"]).reset_index(drop=True)
    if ds.empty:
        return placeholder(p, y, "sheet BU-Skala Kredit kosong untuk provinsi ini.")
    colors = [PAL3[i % 3] for i in range(len(ds))]

    # Batang 100% porsi kredit
    p.text(X0, y + 10, "PORSI KREDIT PER SKALA USAHA", size=11, color=C["soft"], weight="bold")
    total = ds["Share"].clip(lower=0).sum() or 1
    bx, by, bh = X0, y + 30, 52
    clip = p.box(bx, by, CW, bh, fc=C["track"], r=12, z=2)
    cx = bx
    for (_, row), col in zip(ds.iterrows(), colors):
        wseg = CW * max(row["Share"], 0) / total
        p.rect(cx, by, wseg, bh, col, z=2.5, clip=clip)
        lab = f"{row['Skala']}  {format_persen(row['Share'], 1)}"
        tc = C["ink"] if col == C["gold"] else C["white"]
        if p.width_of(lab, 12, "bold") < wseg - 20:
            p.text(cx + wseg / 2, by + bh / 2, lab, size=12, weight="bold", color=tc, ha="center")
        elif p.width_of(format_persen(row["Share"], 0), 11, "bold") < wseg - 8:
            p.text(cx + wseg / 2, by + bh / 2, format_persen(row["Share"], 0), size=11,
                   weight="bold", color=tc, ha="center")
        cx += wseg

    # Kartu NPL per skala
    n = len(ds)
    gap = 20
    cw = (CW - (n - 1) * gap) / n
    cy, ch = y + 110, 186
    for i, row in ds.iterrows():
        x = X0 + i * (cw + gap)
        s = st_max(row["NPL"], b.npl)
        p.card(x, cy, cw, ch, accent=colors[i])
        p.box(x + 20, cy + 26, 14, 14, fc=colors[i], r=3, z=3)
        p.text(x + 42, cy + 33, p.fit(str(row["Skala"]).upper(), 12, cw - 60, "bold"), size=12, weight="bold")
        p.text(x + 20, cy + 64, "NPL", size=10, color=C["soft"], weight="bold")
        p.text(x + 20, cy + 94, format_persen(row["NPL"]), size=23, weight="bold",
               color=C["ink"] if s == "good" else C[s])
        # mini gauge NPL vs batas
        scale = max(b.npl * 1.6, (row["NPL"] if not _nan(row["NPL"]) else 0) * 1.15)
        gx, gy, gw = x + 20, cy + 122, cw - 40
        p.bar_track(gx, gy, gw, 8, (row["NPL"] / scale) if not _nan(row["NPL"]) else np.nan, C[s] if s != "neutral" else C["neutral"])
        tx = gx + gw * b.npl / scale
        p.rect(tx - 1, gy - 5, 2, 18, C["ink"], z=4)
        p.text(tx, gy - 9, f"batas {format_persen(b.npl, 0)}", size=8, color=C["soft"], ha="center", va="bottom")
        p.text(x + 20, cy + ch - 26, "Kredit", size=10, color=C["soft"])
        p.growth(x + 66, cy + ch - 26, row["YoY"], size=10)


# ---------- SEGMEN 4: Kredit per Sektor (Top 5) ----------
SEK_ROW = 78


def h_segment_4(d):
    return 44 + len(d.sektor) * SEK_ROW if not d.sektor.empty else MISSING_H


def draw_segment_4(p: Poster, y, d: ProvData, b: Batas):
    ds = d.sektor
    if ds.empty:
        return placeholder(p, y, "sheet Kredit BU Per Sektor kosong untuk provinsi ini.")
    cx_name, cx_bar, w_bar, cx_yoy, cx_npl = X0 + 64, X0 + 410, 300, X0 + 790, X0 + CW - 16
    for txt, x, ha in [("LAPANGAN USAHA", cx_name, "left"), ("PANGSA KREDIT", cx_bar, "left"),
                       ("PERTUMBUHAN", cx_yoy, "center"), ("NPL", cx_npl - 50, "center")]:
        p.text(x, y + 12, txt, size=10, color=C["soft"], weight="bold", ha=ha)
    vmax = ds["Share"].max() or 1
    for i, row in ds.iterrows():
        ry = y + 36 + i * SEK_ROW
        p.card(X0, ry, CW, SEK_ROW - 10, r=12)
        mid = ry + (SEK_ROW - 10) / 2
        p.ax.add_patch(Circle((X0 + 32, mid), 17, fc=C["red"] if i == 0 else C["ink"], ec="none", zorder=4))
        p.text(X0 + 32, mid + 1, str(i + 1), size=12, color=C["white"], weight="bold", ha="center")
        p.text(cx_name, mid, p.wrap(row["Sektor"], 12, 320, 2, "bold"), size=12, weight="bold",
               linespacing=1.15)
        p.bar_track(cx_bar, mid + 8, w_bar, 10, row["Share"] / vmax, C["red"] if i == 0 else C["gold"])
        p.text(cx_bar, mid - 10, format_persen(row["Share"]), size=13, weight="bold")
        p.growth(cx_yoy, mid, row["YoY"], size=11, anchor="center", plain=True)
        s = st_max(row["NPL"], b.npl)
        p.pill(cx_npl, mid, format_persen(row["NPL"]), C[s], size=11, anchor="right", icon="dot")


# ---------- SEGMEN 5: Bank Imut / BPD ----------
BPD_CARD = 262


def h_segment_5(d):
    if d.bpd is None:
        return MISSING_H
    return max(len(d.bpd), 1) * (BPD_CARD + 20) - 20 if len(d.bpd) else MISSING_H


def draw_segment_5(p: Poster, y, d: ProvData, b: Batas):
    if d.bpd is None:
        return placeholder(p, y, d.errors.get("bpd", "sheet Bank Imut tidak tersedia."))
    if d.bpd.empty:
        return placeholder(p, y, f"tidak ada bank daerah yang dipetakan ke {d.prov}.")
    for i, r in d.bpd.iterrows():
        cy = y + i * (BPD_CARD + 20)
        c = p.card(X0, cy, CW, BPD_CARD)
        p.rect(X0, cy, CW, 54, C["red"], z=2.6, clip=c)
        p.rect(X0, cy + 54, CW, 3, C["gold"], z=2.7, clip=c)
        p.text(X0 + 24, cy + 28, p.fit(r["Bank"], 17, 560, "bold"), size=17, color=C["white"], weight="bold", z=6)
        p.pill(X0 + CW - 20, cy + 28, f"Wilayah: {r['Provinsi']}", C["white"], size=10, anchor="right",
               fill_alpha=0.18, text_color=C["white"])
        # Baris nominal
        cw3 = (CW - 48) / 3
        for j, key in enumerate(["Aset", "DPK", "Kredit"]):
            x = X0 + 24 + j * cw3
            p.text(x, cy + 84, key.upper(), size=10.5, color=C["soft"], weight="bold")
            p.text(x, cy + 114, format_rupiah(r[key]), size=20, weight="bold")
            p.growth(x, cy + 146, r[f"YoY {key}"], size=10)
        p.rect(X0 + 24, cy + 168, CW - 48, 1.2, C["line"], z=3)
        # Baris rasio
        ratios = [("CAR", r["CAR"], st_min(r["CAR"], b.car), f"min. {format_persen(b.car, 0)}", r["Δ CAR"], False),
                  ("NPL GROSS", r["NPL Gross"], st_max(r["NPL Gross"], b.npl), f"batas {format_persen(b.npl, 0)}",
                   r["Δ NPL Gross"], True),
                  ("LDR", r["LDR"], st_range(r["LDR"], b.ldr_min, b.ldr_max),
                   f"koridor {format_persen(b.ldr_min, 0)}–{format_persen(b.ldr_max, 0)}", r["Δ LDR"], None)]
        for j, (lab, v, s, note, dl, inv) in enumerate(ratios):
            x = X0 + 24 + j * cw3
            p.text(x, cy + 192, lab, size=10.5, color=C["soft"], weight="bold")
            p.text(x, cy + 222, format_persen(v), size=19, weight="bold",
                   color=C["ink"] if s in ("good", "neutral") else C[s])
            vx = x + p.width_of(format_persen(v), 19, "bold") + 12
            p.pill(vx, cy + 222, STATUS_TXT[s] if s != "neutral" else "n/a", C[s], size=9, icon="dot")
            dtxt = f"{note}" + ("" if _nan(dl) else f"  •  {format_pp(dl)} YoY")
            p.text(x, cy + 246, dtxt, size=9, color=C["soft"])


# ---------- SEGMEN 6: BPR Kinerja Umum ----------
def h_segment_6(d):
    return 296 if d.bpr is not None else MISSING_H


def draw_segment_6(p: Poster, y, d: ProvData, b: Batas):
    r = d.bpr
    if r is None:
        return placeholder(p, y, d.errors.get("bpr", f"data BPR {d.prov} tidak ditemukan."))
    cw3, gap = (CW - 2 * 20) / 3, 20
    for i, key in enumerate(["Aset", "DPK", "Kredit"]):
        nominal_card(p, X0 + i * (cw3 + gap), y, cw3, 156, f"{key.upper()} BPR", format_rupiah(r[key]),
                     r[f"YoY {key}"])
    y2, g5 = y + 176, 14
    cw5 = (CW - 4 * g5) / 5
    tiles = [
        ("NPL GROSS", r["NPL Gross"], st_max(r["NPL Gross"], b.bpr_npl), f"Batas {format_persen(b.bpr_npl, 0)}",
         r["Δ NPL Gross"], True),
        ("NPL NET", r["NPL Net"], st_max(r["NPL Net"], b.bpr_npl), "Setelah CKPN", r["Δ NPL Net"], True),
        ("LaR", r["LaR"], st_max(r["LaR"], b.bpr_lar), f"Batas {format_persen(b.bpr_lar, 0)}", r["Δ LaR"], True),
        ("LDR", r["LDR"], st_range(r["LDR"], b.bpr_ldr_min, b.bpr_ldr_max),
         f"{format_persen(b.bpr_ldr_min, 0)}–{format_persen(b.bpr_ldr_max, 0)}", r["Δ LDR"], True),
        ("CAR", r["CAR"], st_min(r["CAR"], b.car), f"Min. {format_persen(b.car, 0)}", r["Δ CAR"], False),
    ]
    for i, (lab, v, s, note, dl, inv) in enumerate(tiles):
        ratio_tile(p, X0 + i * (cw5 + g5), y2, cw5, 120, lab, v, s, note, delta=dl, invert=inv, value_size=19)


# ---------- SEGMEN 7: PVML / IKNB ----------
PV_ROW = 82


def h_segment_7(d):
    if d.pvml is None or d.pvml.empty:
        return MISSING_H
    return 40 + len(d.pvml) * PV_ROW + 86


def draw_segment_7(p: Poster, y, d: ProvData, b: Batas):
    dv = d.pvml
    if dv is None or dv.empty:
        return placeholder(p, y, d.errors.get("pvml", f"data PVML {d.prov} tidak ditemukan."))
    cx_name, cx_val, w_bar, cx_yoy, cx_npf = X0 + 24, X0 + 340, 260, X0 + 690, X0 + CW - 16
    for txt, x, ha in [("INDUSTRI", cx_name, "left"), ("OUTSTANDING PEMBIAYAAN", cx_val, "left"),
                       ("PERTUMBUHAN", cx_yoy, "center"), ("NPF", cx_npf - 55, "center")]:
        p.text(x, y + 12, txt, size=10, color=C["soft"], weight="bold", ha=ha)
    vmax = dv["Nominal"].max() if dv["Nominal"].notna().any() else 1
    for i, row in dv.iterrows():
        ry = y + 34 + i * PV_ROW
        p.card(X0, ry, CW, PV_ROW - 12, r=12, accent=PAL6[i % 6], accent_side="left")
        mid = ry + (PV_ROW - 12) / 2
        p.text(cx_name + 4, mid - 10, p.fit(row["Industri"], 13.5, 310, "bold"), size=13.5, weight="bold")
        p.text(cx_name + 4, mid + 14, f"Posisi {row['Periode']}" if row["Periode"] else "", size=9.5, color=C["soft"])
        p.text(cx_val, mid - 10, format_rupiah(row["Nominal"]), size=15, weight="bold")
        p.bar_track(cx_val, mid + 12, w_bar, 8, (row["Nominal"] / vmax) if vmax else np.nan, PAL6[i % 6])
        p.growth(cx_yoy, mid, row["YoY"], size=11, anchor="center", plain=True)
        if _nan(row["NPF"]):
            p.pill(cx_npf, mid, "n/a", C["neutral"], size=10.5, anchor="right")
        else:
            s = st_max(row["NPF"], b.npf)
            nas = " (nas.)" if bool(row.get("NPF Nasional", False)) else ""
            p.pill(cx_npf, mid, f"{format_persen(row['NPF'])}{nas}", C[s], size=10.5, anchor="right", icon="dot")
    # Total
    ty = y + 34 + len(dv) * PV_ROW + 4
    tot, g = dv["Nominal"].sum(), agg_growth(dv["Nominal"], dv["YoY"])
    p.box(X0, ty, CW, 64, fc=C["ink"], r=12, z=2)
    p.text(X0 + 28, ty + 32, "TOTAL PVML", size=12, color=C["gold_light"], weight="bold")
    p.text(cx_val, ty + 32, format_rupiah(tot), size=17, color=C["white"], weight="bold")
    if not _nan(g):
        p.pill(cx_yoy, ty + 32, f"{format_persen(abs(g))} YoY", C["white"], size=10.5, anchor="center",
               icon="up" if g > 0 else "down", fill_alpha=0.15, text_color=C["white"])
    p.text(cx_npf, ty + 32, f"NPF batas {format_persen(b.npf, 0)}", size=9.5, color="#D1D5DB", ha="right")


# ---------- SEGMEN 8: PPDP ----------
INS_CARD = 232


def h_segment_8(d):
    if d.ins is None:
        return MISSING_H
    n_ins = len(d.ins) if not d.ins.empty else 0
    n_lain = len(d.lain) if d.lain is not None and not d.lain.empty else 0
    h = ((n_ins + 1) // 2) * (INS_CARD + 20)
    h += ((n_lain + 1) // 2) * 140
    return max(h - 20, MISSING_H)


def draw_segment_8(p: Poster, y, d: ProvData, b: Batas):
    if d.ins is None:
        return placeholder(p, y, d.errors.get("ppdp", "sheet PPDP tidak tersedia."))
    cw2, gap = (CW - 20) / 2, 20
    cy = y
    for i, r in d.ins.iterrows():
        x = X0 + (i % 2) * (cw2 + gap)
        cy = y + (i // 2) * (INS_CARD + 20)
        p.card(x, cy, cw2, INS_CARD, accent=C["red"] if i % 2 == 0 else C["gold"])
        p.text(x + 22, cy + 34, p.fit(str(r["Lini"]).upper(), 13, cw2 - 170, "bold"), size=13, weight="bold")
        p.pill(x + cw2 - 18, cy + 34, f"Posisi {r['Periode']}", C["soft"], size=9, anchor="right", fill_alpha=0.10)
        for j, (lab, v, g) in enumerate([("PREMI", r["Premi"], r["YoY Premi"]),
                                         ("KLAIM", r["Klaim"], r["YoY Klaim"])]):
            xx = x + 22 + j * (cw2 / 2)
            p.text(xx, cy + 74, lab, size=10.5, color=C["soft"], weight="bold")
            p.text(xx, cy + 104, p.fit(format_rupiah(v), 19, cw2 / 2 - 30, "bold"), size=19, weight="bold")
            p.growth(xx, cy + 136, g, size=10, invert=(lab == "KLAIM"))
        s = st_max(r["Rasio Klaim"], b.klaim)
        p.text(x + 22, cy + 176, "Rasio klaim", size=10.5, color=C["soft"], weight="bold")
        p.text(x + cw2 - 22, cy + 176, format_persen(r["Rasio Klaim"], 1), size=13, weight="bold",
               ha="right", color=C["ink"] if s in ("good", "neutral") else C[s])
        scale = max(1.2, (r["Rasio Klaim"] if not _nan(r["Rasio Klaim"]) else 0) * 1.1)
        gx, gw = x + 22, cw2 - 44
        p.bar_track(gx, cy + 196, gw, 10, (r["Rasio Klaim"] / scale) if not _nan(r["Rasio Klaim"]) else np.nan,
                    C[s] if s != "neutral" else C["neutral"])
        tx = gx + gw * b.klaim / scale
        p.rect(tx - 1, cy + 191, 2, 20, C["ink"], z=4)
        p.text(tx, cy + 219, f"batas {format_persen(b.klaim, 0)}", size=8, color=C["soft"], ha="center")
    yl = y + ((len(d.ins) + 1) // 2) * (INS_CARD + 20) if not d.ins.empty else y
    if d.lain is not None and not d.lain.empty:
        for i, r in d.lain.iterrows():
            x = X0 + (i % 2) * (cw2 + gap)
            ty = yl + (i // 2) * 140
            p.card(x, ty, cw2, 120, accent=C["ink"], accent_side="left")
            lab = str(r["Lini"])
            p.text(x + 26, ty + 30, p.fit(lab.upper(), 11.5, cw2 - 150, "bold"), size=11.5, color=C["soft"], weight="bold")
            p.text(x + cw2 - 18, ty + 30, f"Posisi {r['Periode']}", size=9, color=C["soft"], ha="right")
            p.text(x + 25, ty + 70, format_rupiah(r["Nominal"]), size=22, weight="bold")
            p.growth(x + cw2 - 18, ty + 70, r["YoY"], size=10.5, anchor="right")


# ---------- SEGMEN 9: Pasar Modal ----------
def h_segment_9(d):
    return 370 if d.pm is not None else MISSING_H


def draw_segment_9(p: Poster, y, d: ProvData, b: Batas):
    r = d.pm
    if r is None:
        return placeholder(p, y, d.errors.get("pm", f"data Pasar Modal {d.prov} tidak ditemukan."))
    # Kartu hero SID total
    hw, hh = 300, 180
    p.box(X0, y, hw, hh, fc=C["red"], r=16, z=2, shadow=True)
    p.ax.add_patch(Circle((X0 + hw - 30, y + 20), 90, fc=C["white"], alpha=0.06, ec="none", zorder=2.2))
    p.text(X0 + 24, y + 34, "TOTAL INVESTOR (SID)", size=11.5, color=C["gold_light"], weight="bold")
    p.text(X0 + 22, y + 88, format_ribu(r["SID Total"]), size=30, color=C["white"], weight="bold")
    p.text(X0 + 24, y + 118, "investor domisili provinsi (ribu SID)", size=9.5, color="#F3D3DA")
    g = r["YoY SID Total"]
    if not _nan(g):
        p.pill(X0 + 22, y + 150, f"{format_persen(abs(g))} YoY", C["white"], size=10.5,
               icon="up" if g > 0 else "down", fill_alpha=0.18, text_color=C["white"])
    # Tiles per instrumen
    gap = 16
    tw = (CW - hw - 20 - 2 * gap) / 3
    for i, (kat, col) in enumerate([("Saham", C["red"]), ("SBN", C["gold"]), ("Reksadana", C["ink"])]):
        x = X0 + hw + 20 + i * (tw + gap)
        p.card(x, y, tw, hh, accent=col)
        p.text(x + 20, y + 36, f"SID {kat.upper()}", size=11, color=C["soft"], weight="bold")
        p.text(x + 19, y + 84, format_ribu(r[f"SID {kat}"]), size=22, weight="bold")
        tot = r["SID Total"]
        frac = r[f"SID {kat}"] / tot if (not _nan(tot) and tot) else np.nan
        p.bar_track(x + 20, y + 110, tw - 40, 7, frac, col)
        p.growth(x + 20, y + 148, r[f"YoY SID {kat}"], size=10)

    # Transaksi saham
    ty, th = y + hh + 20, 170
    p.card(X0, ty, CW, th)
    p.text(X0 + 24, ty + 32, "NILAI TRANSAKSI SAHAM INVESTOR DOMISILI PROVINSI", size=11.5,
           color=C["soft"], weight="bold")
    vals = [("Pembelian", r["Pembelian"], r["YoY Pembelian"], C["red"]),
            ("Penjualan", r["Penjualan"], r["YoY Penjualan"], C["gold"])]
    vmax = np.nanmax([v for _, v, _, _ in vals]) if any(not _nan(v) for _, v, _, _ in vals) else 1
    for i, (lab, v, g, col) in enumerate(vals):
        ry = ty + 72 + i * 50
        p.text(X0 + 24, ry, lab, size=12.5, weight="bold")
        p.bar_track(X0 + 130, ry - 9, 330, 18, (v / vmax) if not _nan(v) else np.nan, col)
        p.text(X0 + 476, ry, format_rupiah(v), size=13.5, weight="bold")
        p.growth(X0 + 600, ry, g, size=10.5, plain=True)
    net = r["Net Beli"]
    if not _nan(net):
        col = C["good"] if net >= 0 else C["bad"]
        bx = X0 + CW - 190
        p.box(bx, ty + 62, 166, 88, fc=col, r=12, z=3, alpha=0.10)
        p.text(bx + 83, ty + 88, "NET BELI" if net >= 0 else "NET JUAL", size=10.5, color=col,
               weight="bold", ha="center")
        p.text(bx + 83, ty + 122, format_rupiah(abs(net)), size=15, color=col, weight="bold", ha="center")


# ---------- Registri segmen ----------
SEGMENTS = [
    (1, "Kinerja Umum Bank Umum", h_segment_1, draw_segment_1,
     lambda d: f"Aset, DPK, kredit & rasio utama perbankan • Posisi {d.periode}"),
    (2, "Kredit Berdasarkan Jenis Penggunaan", h_segment_2, draw_segment_2,
     lambda d: "Komposisi & pertumbuhan Modal Kerja, Investasi, Konsumsi • Bank Umum"),
    (3, "Kredit Berdasarkan Skala Usaha", h_segment_3, draw_segment_3,
     lambda d: "Porsi, pertumbuhan & kualitas kredit (NPL) per skala usaha • Bank Umum"),
    (4, "Top 5 Kredit per Lapangan Usaha", h_segment_4, draw_segment_4,
     lambda d: "Sektor ekonomi dengan pangsa kredit terbesar beserta NPL • Bank Umum"),
    (5, "Bank Pembangunan Daerah (Bank Imut)", h_segment_5, draw_segment_5,
     lambda d: f"Bank daerah yang beroperasi di {NAMA_PROV.get(d.prov, d.prov)} • Posisi "
               f"{d.meta.get('bpd', {}).get('periode', '-')} vs {d.meta.get('bpd', {}).get('pembanding', '-')}"),
    (6, "Bank Perekonomian Rakyat (BPR)", h_segment_6, draw_segment_6,
     lambda d: f"Agregat BPR di provinsi • Posisi {d.meta.get('bpr', {}).get('periode', '-')} "
               f"• Δ = perubahan YoY (pp)"),
    (7, "PVML (Pembiayaan, Ventura, Gadai, LKM)", h_segment_7, draw_segment_7,
     lambda d: "Outstanding pembiayaan, pertumbuhan YoY & NPF per industri • Rp miliar"),
    (8, "Asuransi, Dana Pensiun & Penjaminan (PPDP)", h_segment_8, draw_segment_8,
     lambda d: "Premi & klaim asuransi, investasi dana pensiun, outstanding penjaminan"),
    (9, "Pasar Modal", h_segment_9, draw_segment_9,
     lambda d: f"Jumlah investor (SID) & nilai transaksi saham • Posisi {d.meta.get('pm', {}).get('periode', '-')}"),
]


# ---------- Header & Footer ----------
def draw_header(p: Poster, d: ProvData):
    hdr = p.rect(0, 0, W, HEADER_H, C["red"], z=1)
    p.rect(0, HEADER_H * 0.55, W, HEADER_H * 0.45, C["maroon"], z=1.1, alpha=0.35, clip=hdr)
    for cx, cy, r, a in [(1000, 30, 240, 0.07), (900, 300, 150, 0.06), (1070, 230, 90, 0.09)]:
        c = Circle((cx, cy), r, fc=C["white"], ec="none", alpha=a, zorder=1.5)
        p.ax.add_patch(c)
        c.set_clip_path(hdr)
    p.rect(0, HEADER_H, W, 10, C["gold"], z=2)

    p.text(X0, 52, ORG_LABEL, size=11.5, color=C["gold_light"], weight="bold")
    p.text(X0, 118, "Kinerja Sektor Jasa Keuangan", size=36, color=C["white"], weight="bold")
    p.text(X0, 180, f"Provinsi {NAMA_PROV.get(d.prov, d.prov)}", size=28, color=C["gold_light"], weight="bold")
    _, xr = p.pill(X0, 248, f"PERIODE  {str(d.periode).upper()}", C["white"], size=12, fill_alpha=0.16,
                   text_color=C["white"])
    p.pill(xr + 12, 248, "PERBANKAN • IKNB • PASAR MODAL", C["gold_light"], size=12, fill_alpha=0.16,
           text_color=C["gold_light"])
    p.text(X0, 296, "Ringkasan kinerja 9 segmen sektor jasa keuangan berdasarkan Laporan OPS Bulanan",
           size=11, color="#F3D3DA")

    if LOGO_PATH and os.path.exists(LOGO_PATH):
        try:
            img = plt.imread(LOGO_PATH)
            la = p.sub_axes(W - X0 - 150, 40, 150, 110)
            la.imshow(img)
            la.axis("off")
        except Exception:  # logo opsional
            pass


def draw_footer(p: Poster, d: ProvData, y):
    p.rect(0, y, W, 6, C["gold"], z=2)
    p.rect(0, y + 6, W, FOOTER_H - 6, C["ink"], z=1)
    p.text(X0, y + 46, "Sumber Data: Otoritas Jasa Keuangan (OJK) — Laporan OPS Bulanan (diolah)",
           size=13, color=C["white"], weight="bold")
    per = [f"BPD {d.meta.get('bpd', {}).get('periode', '-')}", f"BPR {d.meta.get('bpr', {}).get('periode', '-')}",
           f"Pasar Modal {d.meta.get('pm', {}).get('periode', '-')}"]
    p.text(X0, y + 80, p.fit(f"Periode Bank Umum: {d.periode}  •  " + "  •  ".join(per), 10, CW),
           size=10, color="#CBD5E1")
    p.text(X0, y + 104, p.fit("Periode PVML & PPDP tercantum di tiap kartu. Warna status: hijau = aman, "
                              "kuning = waspada, merah = melewati batas acuan.", 10, CW),
           size=10, color="#CBD5E1")
    p.rect(X0, y + 128, CW, 1, "#374151", z=2)
    gen = now_wib()
    p.text(X0, y + 156, "Infographic Generator OPS Sumbagut", size=10.5, color="#94A3B8")
    p.text(W - X0, y + 156, f"Digenerate {gen:%d-%m-%Y %H:%M} WIB", size=11, color=C["gold"],
           weight="bold", ha="right")


# =============================================================================
# E. RENDER
# =============================================================================
def poster_height(d: ProvData) -> int:
    body = sum(SEG_HEAD_H + h(d) + SEG_GAP for _, _, h, _, _ in SEGMENTS)
    return HEADER_H + 10 + 50 + body + FOOTER_H


def render_infographic(d: ProvData, batas: Batas = Batas(), scale: float = 1.0) -> bytes:
    """Gambar header + 9 segmen + footer → PNG bytes. scale=2 → lebar 2160 px."""
    H = poster_height(d)
    p = Poster(H)
    draw_header(p, d)
    y = HEADER_H + 10 + 50
    for idx, (no, title, h_fn, draw_fn, sub_fn) in enumerate(SEGMENTS):
        h = h_fn(d)
        if idx % 2 == 1:   # pita latar selang-seling agar ritme poster jelas
            p.rect(0, y - SEG_GAP / 2, W, SEG_HEAD_H + h + SEG_GAP, C["band"], z=0.5)
        try:
            sub = sub_fn(d)
        except Exception:  # noqa: BLE001
            sub = ""
        segment_heading(p, y, no, title, sub)
        try:
            draw_fn(p, y + SEG_HEAD_H, d, batas)
        except Exception as e:  # noqa: BLE001 — satu segmen gagal tidak menggagalkan poster
            placeholder(p, y + SEG_HEAD_H, f"segmen gagal digambar ({e})")
        y += SEG_HEAD_H + h + SEG_GAP
    draw_footer(p, d, H - FOOTER_H)

    buf = io.BytesIO()
    p.fig.savefig(buf, format="png", dpi=100 * scale, facecolor=C["white"])
    plt.close(p.fig)
    return buf.getvalue()


# =============================================================================
# F. APLIKASI STREAMLIT
# =============================================================================
def main():
    import streamlit as st

    st.set_page_config(page_title="Infografis OPS Sumbagut", page_icon="🖼️", layout="wide")
    # Kompatibilitas versi Streamlit (sama dengan dashboard): use_container_width -> width="stretch"
    _ver = tuple(int(x) for x in st.__version__.split(".")[:2] if x.isdigit())
    stretch = {"width": "stretch"} if _ver >= (1, 50) else {"use_container_width": True}

    @st.cache_data(show_spinner="Membaca dan memproses file Excel…")
    def _load(file_bytes: bytes, sumber: str):
        return load_workbook(file_bytes, sumber)

    @st.cache_data(show_spinner=False)
    def _render(file_bytes: bytes, sumber: str, prov: str, periode: str, scale: float, batas_t: tuple):
        d = collect(prov, periode, _load(file_bytes, sumber))
        return render_infographic(d, Batas(*batas_t), scale), poster_height(d)

    st.markdown(
        f"""<div style="background:linear-gradient(115deg,{C['red']},{C['maroon']});padding:20px 26px;
        border-radius:14px;border-bottom:4px solid {C['gold']};margin-bottom:16px">
        <div style="color:{C['gold_light']};font-size:12px;font-weight:700;letter-spacing:1px">OJK • SUMBAGUT</div>
        <div style="color:white;font-size:26px;font-weight:800">Infographic Generator — Long Poster 9 Segmen</div>
        <div style="color:#F3D3DA;font-size:14px">Unggah Excel OPS → pilih provinsi → unduh poster PNG</div></div>""",
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.subheader("1. Data")
        up = st.file_uploader("File Excel OPS (.xlsx / .xlsm)", type=["xlsx", "xlsm"])
        periode = st.text_input("Label periode laporan", value=Path(up.name).stem if up else "",
                                placeholder="mis. Agustus 2026")
        sumber = st.radio("Sumber tabel BPD/BPR", [SUMBER_KOREKSI, SUMBER_COGNOS],
                          help="Sama dengan pilihan di dashboard: tabel koreksi (bawah) atau Cognos (atas).")
        st.subheader("2. Provinsi")
        prov = st.selectbox("Provinsi", PROVINSI, format_func=lambda k: f"{k} — {NAMA_PROV[k]}")
        st.subheader("3. Output")
        res = st.radio("Resolusi PNG", ["Standar (lebar 1080 px)", "Tinggi (lebar 2160 px)"], index=1)
        scale = 2.0 if res.startswith("Tinggi") else 1.0
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
    errors = parsed[-1]
    if errors:
        st.warning("Sebagian modul tidak dapat dimuat (segmen terkait ditampilkan sebagai 'data tidak tersedia'): "
                   + "; ".join(errors.values()), icon="⚠️")

    with st.spinner(f"Menyusun poster {prov}…"):
        png, h = _render(file_bytes, sumber, prov, periode or "-", scale, astuple(batas))

    c1, c2 = st.columns([3, 2], gap="large")
    with c1:
        st.markdown("##### Preview Poster")
        st.image(png, width=460)
    with c2:
        st.markdown("##### Unduh")
        nama = f"Infografis_OPS_{prov}_{(periode or 'periode').replace(' ', '_')}.png"
        st.download_button("📥 Download Poster (PNG)", data=png, file_name=nama, mime="image/png",
                           type="primary", **stretch)
        st.caption(f"Ukuran: {int(W * scale)} × {int(h * scale)} px • {len(png) / 1e6:,.1f} MB")
        st.markdown("**Isi poster:**\n\n" + "\n".join(f"{no}. {t}" for no, t, *_ in SEGMENTS))


if __name__ == "__main__":
    main()