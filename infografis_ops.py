# -*- coding: utf-8 -*-
"""
Infographic Generator OPS Sumbagut — ONE-PAGE DASHBOARD (A4 Portrait, 2 Kolom)

Tata letak (satu halaman, proporsi A4 portrait):
    ┌──────────────────────── HEADER (full width) ────────────────────────┐
    │  KOLOM KIRI — Bank Umum        │  KOLOM KANAN — Lembaga Lainnya      │
    │   01 Kinerja Umum              │   05 Bank Daerah (Bank Imut)        │
    │   02 Jenis Kredit              │   06 BPR Kinerja Umum               │
    │   03 Skala Kredit              │   07 PVML                           │
    │   04 Kredit per Sektor         │   08 PPDP                           │
    │                                │   09 Pasar Modal                    │
    └──────────────────────── FOOTER (full width) ────────────────────────┘

Kanvas didesain pada 1240 × 1754 unit (= A4 pada 150 dpi) lalu diekspor:
    scale 1 → 1240 × 1754 px (preview)
    scale 2 → 2480 × 3508 px (A4 300 dpi, siap cetak)   ← default
    scale 2,83 → 3508 × 4961 px (A3 300 dpi)
Juga tersedia ekspor PDF vektor.

Data: hasil parsing riil dari backend dashboard (`ops_parser.py`, wajib berada
di folder yang sama): df_umum, df_jenis, df_skala, df_sektor, df_bpd, df_bpr,
df_pvml, PPDP (ppdp_frames) dan Pasar Modal (pm_wide).

Jalankan :  streamlit run app.py
Opsional  :  logo_ojk.png (latar transparan) → tampil di kiri atas header (lihat LOGO_*).
"""
from __future__ import annotations

import io
import math
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
C = {
    "red": "#B71234", "maroon": "#7A0C22", "gold": "#C2A24D", "gold_light": "#F1DFA8",
    "ink": "#1F2937", "soft": "#6B7280", "line": "#E5E7EB", "page": "#F3F4F6",
    "track": "#EDEFF2", "white": "#FFFFFF",
    "good": "#0F766E", "warn": "#B7791F", "bad": "#C81E3A", "neutral": "#9CA3AF",
}
PAL3 = [C["red"], C["gold"], C["ink"]]
PAL6 = [C["red"], C["gold"], C["ink"], "#9CA3AF", "#E5484D", "#5FA8A0"]

NAMA_PROV = {"Sumut": "Sumatera Utara", "Aceh": "Aceh", "Riau": "Riau",
             "Sumbar": "Sumatera Barat", "Kepri": "Kepulauan Riau"}
ORG_LABEL = "OTORITAS JASA KEUANGAN  •  KANTOR REGIONAL SUMATERA BAGIAN UTARA"

# --- Logo header --------------------------------------------------------------
# File logo (PNG disarankan, latar transparan). Dicari di folder kerja, lalu di folder app.py.
# Bila tidak ditemukan / gagal dibaca → logo dilewati, teks header kembali ke margin kiri.
LOGO_PATH = "logo_ojk.png"
LOGO_BOX = (230, 120)   # area maksimum logo (lebar, tinggi) dalam unit desain; rasio asli dipertahankan
LOGO_PANEL = True       # True = logo di atas panel putih membulat (kontras di header merah)
LOGO_PAD = 12           # padding panel di sekitar logo
LOGO_TEXT_GAP = 26      # jarak horizontal logo → teks header
LOGO_TRIM = True        # pangkas otomatis margin kosong (transparan/putih) di sekeliling logo

PVML_URUTAN = ["pembiayaan", "ventura", "gadai", "lkm", "mikro"]  # urutan tampil industri PVML


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


# --- Geometri halaman (unit desain = px pada skala 1) ------------------------
W, H = 1240, 1754                 # proporsi A4 (1 : √2)
M = 36                            # margin luar kiri/kanan
GUTTER = 24                       # jarak antarkolom
COL_W_L = 548                     # lebar kolom kiri  (Bank Umum)
COL_W_R = W - 2 * M - GUTTER - COL_W_L   # lebar kolom kanan (lembaga lain, lebih padat) = 596
X_LEFT = M                        # x kolom kiri
X_RIGHT = M + COL_W_L + GUTTER    # x kolom kanan
HEADER_H = 160
FOOTER_H = 62
CAPTION_H = 34                    # label "A. PERBANKAN" / "B. ..." di atas kolom
PANEL_GAP = 12                    # jarak vertikal antarpanel
PANEL_HEAD = 58                   # tinggi judul di dalam panel
PAD = 16                          # padding dalam panel

# Bobot tinggi panel tiap kolom (dibagi proporsional terhadap ruang tersedia).
# Bobot dihitung ulang per provinsi oleh panel_weights(): panel tanpa data menyusut
# menjadi strip tipis (EMPTY_W) dan BPD/PVML menyesuaikan jumlah baris.
LEFT_WEIGHTS = [320, 300, 330, 436]          # segmen 1–4
RIGHT_WEIGHTS = [290, 230, 320, 270, 262]    # segmen 5–9 (acuan: 2 bank daerah, 5 industri PVML)
EMPTY_W = 92

EXPORT_SCALES = {
    "A4 cetak 300 dpi — 2480 × 3508 px": 2.0,
    "A3 cetak 300 dpi — 3508 × 4961 px": 3508 / W,
    "Preview layar — 1240 × 1754 px": 1.0,
}

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


# =============================================================================
# B. KANVAS & KOMPONEN VISUAL
# =============================================================================
class Poster:
    """Kanvas A4 berkoordinat unit-desain. Origin KIRI-ATAS, sumbu Y ke bawah.
    Semua elemen ditempatkan dalam kotak (x, y, w, h) — tidak ada kursor 1-D."""

    def __init__(self):
        self.fig = plt.figure(figsize=(W / 100, H / 100), dpi=100)
        self.fig.patch.set_facecolor(C["page"])
        self.ax = self.fig.add_axes([0, 0, 1, 1])
        self.ax.set_xlim(0, W)
        self.ax.set_ylim(H, 0)
        self.ax.axis("off")
        self.renderer = self.fig.canvas.get_renderer()

    # ---------- primitif ----------
    def sub_axes(self, x, y, w, h):
        a = self.fig.add_axes([x / W, (H - y - h) / H, w / W, h / H])
        a.set_facecolor("none")
        return a

    def text(self, x, y, s, size=9, color=None, weight="normal", ha="left", va="center", z=5, **kw):
        return self.ax.text(x, y, s, fontsize=size, color=color or C["ink"], weight=weight,
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
        c = self.box(x, y, w, h, fc=fc or C["white"], ec=C["line"], lw=0.8, r=r, z=z, shadow=True)
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

    def bar(self, x, y, w, h, frac, color, z=3):
        self.box(x, y, w, h, fc=C["track"], r=h / 2, z=z)
        if not _nan(frac) and frac > 0:
            self.box(x, y, max(w * min(frac, 1.0), h), h, fc=color, r=h / 2, z=z + 0.2)

    def gauge(self, x, y, w, v, lim, status, scale_min=1.6, h=6):
        """Bar nilai vs batas, dengan penanda garis batas."""
        scale = max(lim * scale_min, (0 if _nan(v) else v) * 1.12) or 1
        self.bar(x, y, w, h, np.nan if _nan(v) else v / scale, C[status])
        tx = x + w * lim / scale
        self.rect(tx - 0.8, y - 3, 1.6, h + 6, C["ink"], z=4)
        return tx


# ---------- komponen tingkat tinggi ----------
def nominal_tile(p: Poster, x, y, w, h, label, value, g, vsize=15):
    p.card(x, y, w, h, accent=C["red"])
    p.text(x + 12, y + 20, p.fit(label, 7.5, w - 20, "bold"), size=7.5, color=C["soft"], weight="bold")
    p.text(x + 11, y + h * 0.52, p.fit(value, vsize, w - 20, "bold"), size=vsize, weight="bold")
    p.growth(x + 12, y + h - 17, g, size=7.2)


def ratio_tile(p: Poster, x, y, w, h, label, v, status, note, delta=None, invert=True, vsize=14):
    col = C[status]
    p.card(x, y, w, h, accent=col, side="left")
    p.text(x + 13, y + 18, p.fit(label, 7.5, w - 18, "bold"), size=7.5, color=C["soft"], weight="bold")
    p.text(x + 12, y + h * 0.5, format_persen(v), size=vsize, weight="bold",
           color=C["ink"] if status in ("good", "neutral") else col)
    has_d = delta is not None and not _nan(delta)
    dtxt = ("Δ" + format_pp(delta).replace(" pp", "")) if has_d else ""
    dw = p.width_of(dtxt, 6.6, "bold") + 6 if has_d else 0
    p.text(x + 13, y + h - 15, p.fit(note, 6.6, w - 22 - dw), size=6.6, color=C["soft"])
    if has_d:
        worse = (delta > 0) if invert else (delta < 0)
        dc = C["neutral"] if abs(delta) < 5e-5 else (C["bad"] if worse else C["good"])
        p.text(x + w - 7, y + h - 15, dtxt, size=6.6, color=dc, weight="bold", ha="right")


def empty_note(p: Poster, box, msg):
    x, y, w, h = box
    lines = max(1, min(3, int(h // 14)))
    p.text(x + w / 2, y + h / 2 - 2, p.wrap(f"Data tidak tersedia — {msg}", 8, w - 30, lines),
           size=8, color=C["soft"], style="italic", ha="center", linespacing=1.3)


def panel(p: Poster, x, y, w, h, no, title, sub):
    """Bingkai panel segmen. Return kotak konten (x, y, w, h)."""
    p.card(x, y, w, h, r=10)
    p.ax.add_patch(Circle((x + PAD + 12, y + 26), 13, fc=C["red"], ec="none", zorder=4))
    p.ax.add_patch(Circle((x + PAD + 12, y + 26), 15.5, fc="none", ec=C["gold"], lw=1.2, zorder=4))
    p.text(x + PAD + 12, y + 26.5, f"{no:02d}", size=8.5, color=C["white"], weight="bold", ha="center", z=5)
    p.text(x + PAD + 36, y + 19, p.fit(title, 11.5, w - PAD * 2 - 40, "bold"), size=11.5, weight="bold")
    p.text(x + PAD + 36, y + 37, p.fit(sub, 7.2, w - PAD * 2 - 40), size=7.2, color=C["soft"])
    p.rect(x + PAD, y + PANEL_HEAD - 5, w - 2 * PAD, 0.8, C["line"], z=3)
    p.rect(x + PAD, y + PANEL_HEAD - 6, 34, 2.4, C["gold"], z=3.2)
    return x + PAD, y + PANEL_HEAD + 6, w - 2 * PAD, h - PANEL_HEAD - 6 - PAD + 2


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
        prov=prov, periode=periode, umum=_first(df_umum, prov),
        jenis=df_jenis[df_jenis["Provinsi"] == prov].dropna(subset=["Share"]).reset_index(drop=True),
        skala=df_skala[df_skala["Provinsi"] == prov].dropna(subset=["Share"]).reset_index(drop=True),
        sektor=sektor, bpd=bpd, bpr=_first(df_bpr, prov), pvml=pvml, ins=ins, lain=lain,
        pm=_first(pm_wide(df_pm), prov) if df_pm is not None else None, meta=META, errors=ERRORS,
    )


# =============================================================================
# D. SEGMEN — setiap draw_segment_N(p, box, d, b) menggambar di dalam kotak
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
    tiles = [
        ("NPL GROSS", r["NPL Gross"], st_max(r["NPL Gross"], b.npl), f"Batas {format_persen(b.npl, 0)}"),
        ("NPL NET", r["NPL Net"], st_max(r["NPL Net"], b.npl), "Setelah CKPN"),
        ("LOAN AT RISK", r["LaR"], st_max(r["LaR"], b.lar), f"Waspada ≥{format_persen(b.lar, 0)}"),
        ("LDR", r["LDR"], st_range(r["LDR"], b.ldr_min, b.ldr_max),
         f"Koridor {format_persen(b.ldr_min, 0)}–{format_persen(b.ldr_max, 0)}"),
    ]
    for i, (lab, v, s, note) in enumerate(tiles):
        ratio_tile(p, x + i * (rw + gap), ry, rw, rh, lab, v, s, note)


# ---------- 02 Jenis Kredit ----------
def draw_segment_2(p: Poster, box, d: ProvData, b: Batas):
    x, y, w, h = box
    dj = d.jenis
    if dj.empty:
        return empty_note(p, box, "sheet BU-Jenis Kredit kosong untuk provinsi ini.")
    tot = d.umum["Kredit"] if d.umum is not None else np.nan
    shares = dj["Share"].clip(lower=0).to_numpy()
    colors = [PAL3[i % 3] for i in range(len(dj))]

    size = min(h, w * 0.40)
    ax = p.sub_axes(x, y + (h - size) / 2, size, size)
    wedges, _ = ax.pie(shares / shares.sum() if shares.sum() else shares, colors=colors, startangle=90,
                       counterclock=False, wedgeprops=dict(width=0.44, edgecolor=C["white"], linewidth=2.5))
    ax.set_aspect("equal")
    for w_, s, col in zip(wedges, shares, colors):
        if s >= 0.07:
            a = np.deg2rad((w_.theta1 + w_.theta2) / 2)
            ax.text(0.78 * np.cos(a), 0.78 * np.sin(a), format_persen(s, 0), ha="center", va="center",
                    fontsize=7.5, weight="bold", color=C["ink"] if col == C["gold"] else C["white"])
    ax.text(0, 0.11, format_triliun(tot), ha="center", va="center", fontsize=9.5, weight="bold", color=C["ink"])
    ax.text(0, -0.15, "Total Kredit", ha="center", va="center", fontsize=6.8, color=C["soft"])

    x0, x1 = x + size + 22, x + w
    rh = h / max(len(dj), 3)
    for i, row in dj.iterrows():
        ry = y + i * rh + (rh - 58) / 2
        p.box(x0, ry + 6, 11, 11, fc=colors[i], r=2.5, z=3)
        nm = str(row["Jenis Kredit"]) or "-"
        nm = nm if nm.lower().startswith("kredit") else f"Kredit {nm}"
        p.text(x0 + 18, ry + 12, p.fit(nm, 9.5, x1 - x0 - 70, "bold"), size=9.5, weight="bold")
        p.text(x1, ry + 12, format_persen(row["Share"], 1), size=11, weight="bold", ha="right")
        est = row["Share"] * tot if not _nan(tot) else np.nan
        p.text(x0 + 18, ry + 33, f"≈ {format_triliun(est)}", size=9.5, weight="bold", color=C["soft"])
        p.growth(x1, ry + 33, row["YoY"], size=7.3, anchor="right", plain=True)
        p.bar(x0 + 18, ry + 47, x1 - x0 - 18, 5, row["Share"], colors[i])


# ---------- 03 Skala Kredit ----------
def draw_segment_3(p: Poster, box, d: ProvData, b: Batas):
    x, y, w, h = box
    ds = d.skala
    if ds.empty:
        return empty_note(p, box, "sheet BU-Skala Kredit kosong untuk provinsi ini.")
    colors = [PAL3[i % 3] for i in range(len(ds))]
    p.text(x, y + 6, "PORSI KREDIT PER SKALA USAHA", size=7, color=C["soft"], weight="bold")
    by, bh = y + 16, 30
    clip = p.box(x, by, w, bh, fc=C["track"], r=7, z=2)
    total, cx = ds["Share"].clip(lower=0).sum() or 1, x
    for (_, row), col in zip(ds.iterrows(), colors):
        sw = w * max(row["Share"], 0) / total
        p.rect(cx, by, sw, bh, col, z=2.5, clip=clip)
        tc = C["ink"] if col == C["gold"] else C["white"]
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
    for i, row in ds.iterrows():
        cxx = x + i * (cw + gap)
        s = st_max(row["NPL"], b.npl)
        p.card(cxx, cy, cw, ch, accent=colors[i])
        p.box(cxx + 12, cy + 15, 9, 9, fc=colors[i], r=2, z=3)
        p.text(cxx + 26, cy + 20, p.fit(str(row["Skala"]).upper(), 8.5, cw - 34, "bold"), size=8.5, weight="bold")
        p.text(cxx + 12, cy + 42, "NPL", size=7, color=C["soft"], weight="bold")
        p.text(cxx + 11, cy + 64, format_persen(row["NPL"]), size=16, weight="bold",
               color=C["ink"] if s in ("good", "neutral") else C[s])
        gy = cy + 88
        tx = p.gauge(cxx + 12, gy, cw - 24, row["NPL"], b.npl, s)
        p.text(tx, gy + 16, f"batas {format_persen(b.npl, 0)}", size=6, color=C["soft"], ha="center")
        my = gy + 22 + (cy + ch - 32 - gy - 22) / 2
        p.rect(cxx + 12, gy + 26, cw - 24, 0.8, C["line"], z=3)
        p.text(cxx + 12, my - 8, "PORSI KREDIT", size=6.6, color=C["soft"], weight="bold")
        p.text(cxx + 11, my + 10, format_persen(row["Share"], 1), size=12.5, weight="bold", color=C["soft"])
        p.text(cxx + 12, cy + ch - 18, "Kredit", size=7, color=C["soft"])
        p.growth(cxx + 42, cy + ch - 18, row["YoY"], size=7)


# ---------- 04 Kredit per Sektor (Top 5) ----------
def draw_segment_4(p: Poster, box, d: ProvData, b: Batas):
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
        p.dot(x + 18, mid, C["red"] if i == 0 else C["ink"], r=10)
        p.text(x + 18, mid + 0.5, str(i + 1), size=8, color=C["white"], weight="bold", ha="center", z=7)
        p.text(c_name, mid, p.wrap(row["Sektor"], 8.2, name_w, 2, "bold"), size=8.2, weight="bold", linespacing=1.15)
        p.text(c_bar, mid - 7, format_persen(row["Share"]), size=9.5, weight="bold")
        p.bar(c_bar, mid + 5, w_bar, 5, row["Share"] / vmax, C["red"] if i == 0 else C["gold"])
        p.growth(c_yoy, mid, row["YoY"], size=7.3, anchor="center", plain=True, suffix="")
        p.pill(c_npl, mid, format_persen(row["NPL"]), C[st_max(row["NPL"], b.npl)], size=7.5,
               anchor="right", icon="dot")


# ---------- 05 Bank Daerah (Bank Imut) ----------
def draw_segment_5(p: Poster, box, d: ProvData, b: Batas):
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
        p.rect(x, by, w, 26, C["red"], z=2.6, clip=c)
        p.rect(x, by + 26, w, 1.8, C["gold"], z=2.7, clip=c)
        p.text(x + 12, by + 13, p.fit(r["Bank"], 9.5, w - 160, "bold"), size=9.5, color=C["white"], weight="bold", z=6)
        p.pill(x + w - 8, by + 13, f"Wilayah: {r['Provinsi']}", C["white"], size=6.5, anchor="right",
               alpha=0.18, tcolor=C["white"])
        gy = by + 26 + (bh - 26) / 2
        cells = [("ASET", "Aset", None), ("DPK", "DPK", None), ("KREDIT", "Kredit", None),
                 ("CAR", "CAR", st_min(r["CAR"], b.car)), ("NPL GROSS", "NPL Gross", st_max(r["NPL Gross"], b.npl)),
                 ("LDR", "LDR", st_range(r["LDR"], b.ldr_min, b.ldr_max))]
        for j, (lab, key, s) in enumerate(cells):
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
                p.text(cx + 10, gy + 20, ("Δ " + format_pp(dl)) if not _nan(dl) else STATUS_TXT[s],
                       size=6.4, color=C["soft"])


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
    tiles = [
        ("NPL GROSS", "NPL Gross", st_max(r["NPL Gross"], b.bpr_npl), f"≤{format_persen(b.bpr_npl, 0)}", True),
        ("NPL NET", "NPL Net", st_max(r["NPL Net"], b.bpr_npl), "Net", True),
        ("LaR", "LaR", st_max(r["LaR"], b.bpr_lar), f"≤{format_persen(b.bpr_lar, 0)}", True),
        ("LDR", "LDR", st_range(r["LDR"], b.bpr_ldr_min, b.bpr_ldr_max),
         f"{format_persen(b.bpr_ldr_min, 0)}–{format_persen(b.bpr_ldr_max, 0)}", True),
        ("CAR", "CAR", st_min(r["CAR"], b.car), f"≥{format_persen(b.car, 0)}", False),
    ]
    for i, (lab, key, s, note, inv) in enumerate(tiles):
        ratio_tile(p, x + i * (rw + gap), ry, rw, rh, lab, r[key], s, note, delta=r[f"Δ {key}"],
                   invert=inv, vsize=12)


# ---------- 07 PVML ----------
def draw_segment_7(p: Poster, box, d: ProvData, b: Batas):
    x, y, w, h = box
    dv = d.pvml
    if dv is None or dv.empty:
        return empty_note(p, box, d.errors.get("pvml", f"data PVML {d.prov} tidak ditemukan."))
    c_name, c_npf = x + 12, x + w - 6
    c_val, w_bar, c_yoy = x + w * 0.36, w * 0.21, x + w * 0.70
    name_w = c_val - c_name - 12
    for t, xx, ha in [("INDUSTRI", c_name, "left"), ("OUTSTANDING", c_val, "left"),
                      ("YoY", c_yoy, "center"), ("NPF", c_npf - 30, "center")]:
        p.text(xx, y + 5, t, size=6.8, color=C["soft"], weight="bold", ha=ha)
    tot_h = 30
    rh = min((h - 16 - tot_h - 6) / len(dv), 46)
    vmax = dv["Nominal"].max() if dv["Nominal"].notna().any() else 1
    two_line = rh >= 32
    for i, row in dv.iterrows():
        ry = y + 16 + i * rh
        col = PAL6[i % 6]
        p.card(x, ry, w, rh - 4, r=6, accent=col, side="left")
        mid = ry + (rh - 4) / 2
        if two_line:
            p.text(c_name, mid - 6, p.fit(row["Industri"], 8.3, name_w, "bold"), size=8.3, weight="bold")
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
            nas = " nas." if bool(row.get("NPF Nasional", False)) else ""
            p.pill(c_npf, mid, f"{format_persen(row['NPF'])}{nas}", C[st_max(row["NPF"], b.npf)],
                   size=7, anchor="right", icon="dot")
    ty = y + h - tot_h
    tot, g = dv["Nominal"].sum(), agg_growth(dv["Nominal"], dv["YoY"])
    p.box(x, ty, w, tot_h, fc=C["ink"], r=7, z=2)
    p.text(c_name, ty + tot_h / 2, "TOTAL PVML", size=8, color=C["gold_light"], weight="bold")
    p.text(c_val, ty + tot_h / 2, format_rupiah(tot), size=10, color=C["white"], weight="bold")
    if not _nan(g):
        p.pill(c_yoy, ty + tot_h / 2, f"{format_persen(abs(g))}", C["white"], size=7, anchor="center",
               icon="up" if g > 0 else "down", alpha=0.15, tcolor=C["white"])
    p.text(c_npf, ty + tot_h / 2, f"batas NPF {format_persen(b.npf, 0)}", size=6.5, color="#D1D5DB", ha="right")


# ---------- 08 PPDP ----------
def draw_segment_8(p: Poster, box, d: ProvData, b: Batas):
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
        p.card(cx, cy, cw, ch, accent=C["red"] if i % 2 == 0 else C["gold"])
        p.text(cx + 11, cy + 18, p.fit(str(r["Lini"]).upper(), 8.5, cw - 92, "bold"), size=8.5, weight="bold")
        p.pill(cx + cw - 8, cy + 18, f"Posisi {r['Periode']}", C["soft"], size=6, anchor="right", alpha=0.1)
        for j, (lab, v, g) in enumerate([("PREMI", r["Premi"], r["YoY Premi"]), ("KLAIM", r["Klaim"], r["YoY Klaim"])]):
            xx = cx + 11 + j * (cw / 2)
            p.text(xx, cy + 38, lab, size=6.6, color=C["soft"], weight="bold")
            p.text(xx, cy + 56, p.fit(format_rupiah(v), 11, cw / 2 - 18, "bold"), size=11, weight="bold")
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
        p.text(tx + 13, ty + 15, p.fit(str(r["Lini"]).upper(), 7, cw - 80, "bold"), size=7, color=C["soft"], weight="bold")
        p.text(tx + cw - 8, ty + 15, f"Posisi {r['Periode']}", size=6, color=C["soft"], ha="right")
        p.text(tx + 12, ty + 41, format_rupiah(r["Nominal"]), size=13, weight="bold")
        p.growth(tx + cw - 8, ty + 41, r["YoY"], size=7, anchor="right")


# ---------- 09 Pasar Modal ----------
def draw_segment_9(p: Poster, box, d: ProvData, b: Batas):
    x, y, w, h = box
    r = d.pm
    if r is None:
        return empty_note(p, box, d.errors.get("pm", f"data Pasar Modal {d.prov} tidak ditemukan."))
    gap = 8
    th = round(h * 0.5)
    hw = w * 0.31
    p.box(x, y, hw, th, fc=C["red"], r=8, z=2, shadow=True)
    p.ax.add_patch(Circle((x + hw - 14, y + 8), 50, fc=C["white"], alpha=0.07, ec="none", zorder=2.2))
    p.text(x + 12, y + 17, "TOTAL INVESTOR (SID)", size=7, color=C["gold_light"], weight="bold")
    p.text(x + 11, y + th * 0.5, p.fit(format_ribu(r["SID Total"]), 17, hw - 20, "bold"), size=17,
           color=C["white"], weight="bold")
    g = r["YoY SID Total"]
    if not _nan(g):
        p.pill(x + 11, y + th - 16, f"{format_persen(abs(g))} YoY", C["white"], size=6.8,
               icon="up" if g > 0 else "down", alpha=0.18, tcolor=C["white"])
    tw = (w - hw - 3 * gap) / 3
    for i, (kat, col) in enumerate([("Saham", C["red"]), ("SBN", C["gold"]), ("Reksadana", C["ink"])]):
        tx = x + hw + gap + i * (tw + gap)
        p.card(tx, y, tw, th, accent=col)
        p.text(tx + 10, y + 18, f"SID {kat.upper()}", size=7, color=C["soft"], weight="bold")
        p.text(tx + 9, y + th * 0.46, p.fit(format_ribu(r[f"SID {kat}"]), 12.5, tw - 16, "bold"), size=12.5,
               weight="bold")
        tot = r["SID Total"]
        p.bar(tx + 10, y + th * 0.64, tw - 20, 4, r[f"SID {kat}"] / tot if (not _nan(tot) and tot) else np.nan, col)
        p.growth(tx + 10, y + th - 14, r[f"YoY SID {kat}"], size=6.6, plain=True)

    ty, tth = y + th + gap, h - th - gap
    p.card(x, ty, w, tth)
    p.text(x + 12, ty + 15, "TRANSAKSI SAHAM INVESTOR PROVINSI", size=7, color=C["soft"], weight="bold")
    vals = [("Pembelian", r["Pembelian"], r["YoY Pembelian"], C["red"]),
            ("Penjualan", r["Penjualan"], r["YoY Penjualan"], C["gold"])]
    ok = [v for _, v, _, _ in vals if not _nan(v)]
    vmax = max(ok) if ok else 1
    net = r["Net Beli"]
    nb_w = 112
    bar_w = w - 12 - 62 - 82 - 70 - nb_w - 24
    for i, (lab, v, gg, col) in enumerate(vals):
        ry = ty + 36 + i * ((tth - 44) / 2)
        p.text(x + 12, ry, lab, size=8, weight="bold")
        p.bar(x + 74, ry - 5, bar_w, 10, (v / vmax) if not _nan(v) else np.nan, col)
        p.text(x + 82 + bar_w, ry, format_rupiah(v), size=8.5, weight="bold")
        p.growth(x + 82 + bar_w + 78, ry, gg, size=6.6, plain=True, suffix="")
    if not _nan(net):
        col = C["good"] if net >= 0 else C["bad"]
        bx = x + w - nb_w - 10
        p.box(bx, ty + 22, nb_w, tth - 32, fc=col, r=7, z=3, alpha=0.1)
        p.text(bx + nb_w / 2, ty + 22 + (tth - 32) * 0.32, "NET BELI" if net >= 0 else "NET JUAL", size=7,
               color=col, weight="bold", ha="center")
        p.text(bx + nb_w / 2, ty + 22 + (tth - 32) * 0.68, p.fit(format_rupiah(abs(net)), 10, nb_w - 10, "bold"),
               size=10, color=col, weight="bold", ha="center")


# ---------- Registri: (nomor, judul, subjudul(d), fungsi gambar) ----------
LEFT_SEGMENTS = [
    (1, "Kinerja Umum Bank Umum", lambda d: f"Aset, DPK, kredit & rasio utama • Posisi {d.periode}", draw_segment_1),
    (2, "Kredit per Jenis Penggunaan", lambda d: "Porsi & YoY • ≈ nominal = porsi × total kredit BU",
     draw_segment_2),
    (3, "Kredit per Skala Usaha", lambda d: "Porsi, pertumbuhan & NPL per skala usaha", draw_segment_3),
    (4, "Top 5 Kredit per Lapangan Usaha", lambda d: "Sektor dengan pangsa kredit terbesar beserta NPL",
     draw_segment_4),
]
RIGHT_SEGMENTS = [
    (5, "Bank Daerah (Bank Imut)", lambda d: f"Posisi {d.meta.get('bpd', {}).get('periode', '-')} vs "
                                             f"{d.meta.get('bpd', {}).get('pembanding', '-')} • Δ = perubahan YoY",
     draw_segment_5),
    (6, "BPR — Kinerja Umum", lambda d: f"Agregat BPR provinsi • Posisi {d.meta.get('bpr', {}).get('periode', '-')} • Δ = perubahan YoY (pp)",
     draw_segment_6),
    (7, "PVML", lambda d: "Pembiayaan, Modal Ventura, Pergadaian, LKM & lainnya • outstanding & NPF",
     draw_segment_7),
    (8, "PPDP — Asuransi, Dapen & Penjaminan", lambda d: "Premi & klaim asuransi, investasi dapen, penjaminan",
     draw_segment_8),
    (9, "Pasar Modal", lambda d: f"Investor (SID) & transaksi saham • Posisi {d.meta.get('pm', {}).get('periode', '-')}",
     draw_segment_9),
]


# =============================================================================
# E. HEADER, FOOTER & LAYOUT GRID
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


def draw_logo(p: Poster) -> float:
    """Gambar logo di pojok kiri atas header (proporsional, di tengah vertikal header).
    Return koordinat x tempat teks header dimulai (M bila logo tidak tersedia)."""
    img = _load_logo()
    if img is None:
        return M
    try:
        ih, iw = img.shape[:2]
        box_w, box_h = LOGO_BOX
        pad = LOGO_PAD if LOGO_PANEL else 0
        s = min((box_w - 2 * pad) / iw, (box_h - 2 * pad) / ih)   # skala tanpa distorsi
        lw, lh = iw * s, ih * s
        pw, ph = lw + 2 * pad, lh + 2 * pad
        px, py = M, (HEADER_H - ph) / 2
        if LOGO_PANEL:
            p.box(px, py, pw, ph, fc=C["white"], r=12, z=2.5, shadow=True)
        ax = p.sub_axes(px + pad, py + pad, lw, lh)
        ax.imshow(img, cmap="gray" if img.ndim == 2 else None, interpolation="antialiased")
        ax.set_axis_off()
        return px + pw + LOGO_TEXT_GAP
    except Exception:  # noqa: BLE001 — gagal menggambar logo → teks tetap di margin kiri
        return M


def draw_header(p: Poster, d: ProvData):
    hdr = p.rect(0, 0, W, HEADER_H, C["red"], z=1)
    p.rect(0, HEADER_H * 0.55, W, HEADER_H * 0.45, C["maroon"], z=1.1, alpha=0.35, clip=hdr)
    for cx, cy, r, a in [(1150, 10, 170, 0.07), (1040, 165, 110, 0.06), (1230, 120, 70, 0.09)]:
        c = Circle((cx, cy), r, fc=C["white"], ec="none", alpha=a, zorder=1.5)
        p.ax.add_patch(c)
        c.set_clip_path(hdr)
    p.rect(0, HEADER_H, W, 6, C["gold"], z=2)

    # Logo di kiri atas; semua teks header dimulai setelah logo (tx)
    tx = draw_logo(p)
    text_w = W - M - tx   # ruang horizontal tersisa untuk teks
    p.text(tx, 32, p.fit(ORG_LABEL, 8.5, text_w, "bold"), size=8.5, color=C["gold_light"], weight="bold")
    p.text(tx, 72, p.fit("Kinerja Sektor Jasa Keuangan", 27, text_w, "bold"), size=27, color=C["white"],
           weight="bold")
    p.text(tx, 110, p.fit(f"Provinsi {NAMA_PROV.get(d.prov, d.prov)}", 19, text_w, "bold"), size=19,
           color=C["gold_light"], weight="bold")
    xr = p.pill(tx, 141, f"PERIODE  {str(d.periode).upper()}", C["white"], size=8.5, alpha=0.16, tcolor=C["white"])
    chip = "BANK UMUM • BPD • BPR • PVML • PPDP • PASAR MODAL"
    if xr + 8 + p.width_of(chip, 8.5, "bold") + 16 <= W - M:   # chip kedua hanya jika muat
        p.pill(xr + 8, 141, chip, C["gold_light"], size=8.5, alpha=0.16, tcolor=C["gold_light"])


def draw_column_caption(p: Poster, x, y, code, text, col_w):
    p.rect(x, y + 4, 4, 18, C["red"], z=3)
    p.text(x + 12, y + 13, code, size=9.5, color=C["red"], weight="bold")
    p.text(x + 12 + p.width_of(code, 9.5, "bold") + 6, y + 13, text, size=9.5, weight="bold")
    p.rect(x, y + CAPTION_H - 4, col_w, 1, C["line"], z=2)


def draw_footer(p: Poster, d: ProvData):
    y = H - FOOTER_H
    p.rect(0, y, W, 4, C["gold"], z=2)
    p.rect(0, y + 4, W, FOOTER_H - 4, C["ink"], z=1)
    p.text(M, y + 22, "Sumber Data: Otoritas Jasa Keuangan (OJK) — Laporan OPS Bulanan (diolah)",
           size=8.5, color=C["white"], weight="bold")
    per = (f"Periode: Bank Umum {d.periode} • BPD {d.meta.get('bpd', {}).get('periode', '-')} • "
           f"BPR {d.meta.get('bpr', {}).get('periode', '-')} • Pasar Modal {d.meta.get('pm', {}).get('periode', '-')} "
           f"• PVML & PPDP per kartu.   Warna: hijau aman, kuning waspada, merah melewati batas.")
    p.text(M, y + 43, p.fit(per, 6.8, W - 2 * M - 230), size=6.8, color="#CBD5E1")
    gen = now_wib()
    p.text(W - M, y + 22, "Waktu Generate", size=7, color="#94A3B8", ha="right")
    p.text(W - M, y + 41, f"{gen:%d-%m-%Y  %H:%M} WIB", size=9, color=C["gold"], weight="bold", ha="right")


def panel_weights(d: ProvData):
    """Bobot tinggi panel kiri & kanan untuk provinsi ini."""
    def has(df):
        return df is not None and not getattr(df, "empty", False)
    left = [w if ok else EMPTY_W for w, ok in zip(
        LEFT_WEIGHTS, [d.umum is not None, has(d.jenis), has(d.skala), has(d.sektor)])]
    n_bpd = len(d.bpd) if has(d.bpd) else 0
    n_pv = len(d.pvml) if has(d.pvml) else 0
    right = [
        75 + 100 * min(n_bpd, 3) if n_bpd else EMPTY_W,
        RIGHT_WEIGHTS[1] if d.bpr is not None else EMPTY_W,
        110 + 38 * min(n_pv, 7) if n_pv else EMPTY_W,
        RIGHT_WEIGHTS[3] if d.ins is not None else EMPTY_W,
        RIGHT_WEIGHTS[4] if d.pm is not None else EMPTY_W,
    ]
    return left, right


def layout_column(top, bottom, weights):
    """Bagi ruang vertikal [top, bottom] menjadi panel sesuai bobot. Return list (y, h)."""
    avail = bottom - top - PANEL_GAP * (len(weights) - 1)
    tot = sum(weights)
    out, y = [], top
    for wgt in weights:
        h = avail * wgt / tot
        out.append((y, h))
        y += h + PANEL_GAP
    return out


def render_infographic(d: ProvData, batas: Batas = Batas(), scale: float = 2.0, fmt: str = "png") -> bytes:
    """Render one-page dashboard. fmt='png' (pakai scale) atau 'pdf' (vektor)."""
    p = Poster()
    draw_header(p, d)

    cap_y = HEADER_H + 6 + 12
    draw_column_caption(p, X_LEFT, cap_y, "A.", "PERBANKAN — BANK UMUM", COL_W_L)
    draw_column_caption(p, X_RIGHT, cap_y, "B.", "BANK DAERAH, BPR, IKNB & PASAR MODAL", COL_W_R)
    top, bottom = cap_y + CAPTION_H + 6, H - FOOTER_H - 16

    w_left, w_right = panel_weights(d)
    for x_col, col_w, segs, weights in [(X_LEFT, COL_W_L, LEFT_SEGMENTS, w_left),
                                        (X_RIGHT, COL_W_R, RIGHT_SEGMENTS, w_right)]:
        for (no, title, sub_fn, draw_fn), (py, ph) in zip(segs, layout_column(top, bottom, weights)):
            try:
                sub = sub_fn(d)
            except Exception:  # noqa: BLE001
                sub = ""
            content = panel(p, x_col, py, col_w, ph, no, title, sub)
            try:
                draw_fn(p, content, d, batas)
            except Exception as e:  # noqa: BLE001 — satu segmen gagal tidak menggagalkan halaman
                empty_note(p, content, f"segmen gagal digambar ({e})")

    draw_footer(p, d)
    buf = io.BytesIO()
    p.fig.savefig(buf, format=fmt, dpi=100 * scale, facecolor=C["page"])
    plt.close(p.fig)
    return buf.getvalue()


# =============================================================================
# F. APLIKASI STREAMLIT
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
                logo_sig: str):  # logo_sig hanya kunci cache: logo baru/berubah → render ulang
        d = collect(prov, periode, _load(file_bytes, sumber))
        return render_infographic(d, Batas(*batas_t), scale, fmt)

    st.markdown(
        f"""<div style="background:linear-gradient(115deg,{C['red']},{C['maroon']});padding:18px 24px;
        border-radius:14px;border-bottom:4px solid {C['gold']};margin-bottom:16px">
        <div style="color:{C['gold_light']};font-size:12px;font-weight:700;letter-spacing:1px">OJK • SUMBAGUT</div>
        <div style="color:white;font-size:25px;font-weight:800">Infographic Generator — One-Page Dashboard A4</div>
        <div style="color:#F3D3DA;font-size:14px">Unggah Excel OPS → pilih provinsi → unduh infografis 1 halaman</div>
        </div>""",
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
        res = st.radio("Ukuran PNG", list(EXPORT_SCALES))
        lf = logo_file()
        if lf is not None and _load_logo() is not None:
            st.caption(f"🖼️ Logo terdeteksi: `{lf.resolve()}`")
        elif lf is not None:
            st.warning(f"File logo ditemukan tetapi gagal dibaca: `{lf}`. Pastikan berformat PNG/JPG valid.", icon="🖼️")
        else:
            st.warning("Logo tidak ditemukan. Lokasi yang diperiksa:\n\n"
                       + "\n".join(f"- `{c.resolve()}`" for c in logo_candidates())
                       + "\n\nPeriksa nama file (huruf kecil, bukan `logo_ojk.png.png`).", icon="🖼️")
        scale = EXPORT_SCALES[res]
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
    if parsed[-1]:
        st.warning("Sebagian modul tidak dapat dimuat (panel terkait ditandai 'data tidak tersedia'): "
                   + "; ".join(parsed[-1].values()), icon="⚠️")

    per = periode or "-"
    with st.spinner(f"Menyusun infografis {prov}…"):
        png = _render(file_bytes, sumber, prov, per, scale, astuple(batas), "png", logo_signature())

    c1, c2 = st.columns([3, 2], gap="large")
    with c1:
        st.markdown("##### Preview (1 halaman A4)")
        st.image(png, width=620)
    with c2:
        st.markdown("##### Unduh")
        base = f"Infografis_OPS_{prov}_{(periode or 'periode').replace(' ', '_')}"
        st.download_button("📥 Download PNG", data=png, file_name=f"{base}.png", mime="image/png",
                           type="primary", **stretch)
        st.caption(f"{res} • {len(png) / 1e6:,.1f} MB")
        pdf = _render(file_bytes, sumber, prov, per, 1.0, astuple(batas), "pdf", logo_signature())
        st.download_button("📄 Download PDF (vektor, siap cetak)", data=pdf, file_name=f"{base}.pdf",
                           mime="application/pdf", **stretch)
        st.markdown("**Kolom kiri — Bank Umum:**\n" + "\n".join(f"- {n:02d} {t}" for n, t, *_ in LEFT_SEGMENTS))
        st.markdown("**Kolom kanan — Lembaga lainnya:**\n" + "\n".join(f"- {n:02d} {t}" for n, t, *_ in RIGHT_SEGMENTS))


if __name__ == "__main__":
    main()