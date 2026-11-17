# -*- coding: utf-8 -*-
"""
Calendari de maneig de la campanya 2026.

Sortides:
  figures/fig17_calendari.png   figura completa per al TFM
  calendari_maneig_2026.xlsx    full editable amb les dades i un gràfic natiu
  calendari_maneig_2026.pptx    versió editable en diapositiva

Color: les dues magnituds ordinals porten rampa d'un sol to, validades amb
l'script del mètode de visualització (monotonia de lluminositat, salt entre
passos i contrast de l'extrem clar). Les files d'identitat porten to
categòric i van sempre etiquetades, de manera que el color no és l'únic
portador del significat.
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle
from PIL import Image

BASE = os.path.dirname(os.path.abspath(__file__))
FIGDIR = os.path.join(BASE, "figures")
os.makedirs(FIGDIR, exist_ok=True)

MESOS = ["Març", "Abril", "Maig", "Juny", "Juliol", "Agost", "Setembre", "Octubre"]
NSET = 32

BROTACIO, DESBROT, VEROLAT, DESFULLAT = 3, 12, 18, 18
VEREMA = (23, 26)
VOLS = [("03/06", 12), ("30/06", 15), ("22/07", 19), ("12/08", 21)]
KC_SPAN, KC = (8, 27), "0,45"
REG = [((8, 18), "100 %"), ((19, 26), "25 %"), ((27, 27), "100 %")]
TRACT = [("SE", "Sense esporga", 1.4),
         ("EL", "Esporga lleugera", 1.1),
         ("EM", "Esporga mitjana", 1.0),
         ("ES", "Esporga severa", 0.7),
         ("ES+DV", "Esporga severa + desfullat", 0.7)]

INK, MUTED, RULE, FILL = "#1a2230", "#52514e", "#d4d7dc", "#f4f6f8"
C_FEN, C_TAS, C_KC = "#1baf7a", "#6250d6", "#3d4451"
KR_RAMP = {0.7: ("#e7a384", INK), 1.0: ("#d9825a", INK),
           1.1: ("#c56437", "white"), 1.4: ("#9e441b", "white")}
REG_RAMP = {"25 %": ("#86b6ef", INK), "100 %": ("#1c5cab", "white")}

plt.rcParams.update({"font.family": "DejaVu Sans", "text.color": INK})

fig, ax = plt.subplots(figsize=(10.6, 6.5))
X0 = -11.2
ax.set_xlim(X0, NSET + 0.3)
ax.axis("off")

H, HK = 0.70, 0.52
y = 0.0
rows = []                       # (clau, y)
for key, h, gap in [("fen", H, 0.16), ("tas", H, 0.16), ("vol", H, 0.30),
                    ("kc", H, 0.34)]:
    y -= h
    rows.append((key, y))
    y -= gap
y -= 0.76                       # espai per al títol del bloc Kr
for code, _, _ in TRACT:
    y -= HK
    rows.append(("kr_" + code, y))
    y -= 0.07
y -= 0.30
y -= H
rows.append(("reg", y))
BOTTOM = y
POS = dict(rows)
TOP = 0.55
ax.set_ylim(BOTTOM - 1.95, TOP + 1.25)


def band(yy, h=H):
    ax.add_patch(Rectangle((0, yy), NSET, h, fc=FILL, ec=RULE, lw=0.7, zorder=1))


def label(yy, text, h=H, bold=True, size=9.4, sub=None):
    ax.text(-0.5, yy + h / 2 + (0.11 if sub else 0), text, ha="right",
            va="center", fontsize=size, weight="bold" if bold else "normal",
            color=INK)
    if sub:
        ax.text(-0.5, yy + h / 2 - 0.15, sub, ha="right", va="center",
                fontsize=7.6, color=MUTED)


def chip(yy, a, b, fc, text="", tc="white", h=H, fs=8.2, r=0.10):
    ax.add_patch(FancyBboxPatch((a, yy + 0.07), b - a + 1, h - 0.14,
                                boxstyle=f"round,pad=0,rounding_size={r}",
                                fc=fc, ec="white", lw=1.1, zorder=3))
    if text:
        ax.text((a + b + 1) / 2, yy + h / 2, text, ha="center", va="center",
                fontsize=fs, weight="bold", color=tc, zorder=4)


# --- capçalera de mesos i setmanes ----------------------------------------
for m, mes in enumerate(MESOS):
    x = m * 4
    ax.add_patch(Rectangle((x, TOP + 0.42), 4, 0.44,
                           fc="#e9edf1" if m % 2 == 0 else "#f5f7f9",
                           ec=RULE, lw=0.7, zorder=1))
    ax.text(x + 2, TOP + 0.64, mes.upper(), ha="center", va="center",
            fontsize=8.4, weight="bold", color=INK)
    for s in range(4):
        ax.text(x + s + 0.5, TOP + 0.20, str(s + 1), ha="center", va="center",
                fontsize=6.8, color=MUTED)
for m in range(1, 8):
    ax.plot([m * 4, m * 4], [BOTTOM, TOP + 0.42], color=RULE, lw=0.7, zorder=0)

# --- fenologia -------------------------------------------------------------
yy = POS["fen"]; band(yy); label(yy, "Fenologia")
for idx, txt in ((BROTACIO, "Brotació"), (VEROLAT, "Verolat")):
    chip(yy, idx, idx, C_FEN)
    ax.text(idx + 1.35, yy + H / 2, txt, ha="left", va="center",
            fontsize=8.4, color=INK)

# --- tasques ---------------------------------------------------------------
yy = POS["tas"]; band(yy); label(yy, "Operacions")
chip(yy, DESBROT, DESBROT, C_TAS)
ax.text(DESBROT - 0.35, yy + H / 2, "Desbrot i tall en altura", ha="right",
        va="center", fontsize=8.4, color=INK)
chip(yy, DESFULLAT, DESFULLAT, C_TAS)
ax.text(DESFULLAT + 1.35, yy + H / 2, "Desfullat", ha="left", va="center",
        fontsize=8.4, color=INK)
chip(yy, *VEREMA, fc=C_TAS, text="Verema")

# --- vols ------------------------------------------------------------------
yy = POS["vol"]; band(yy); label(yy, "Vols de dron")
for n, (dat, idx) in enumerate(VOLS, start=1):
    chip(yy, idx, idx, INK, str(n), "white", fs=8.4, r=0.14)
    ax.text(idx + 0.5, yy - 0.16, dat, ha="center", va="top", fontsize=7.2,
            color=MUTED)

# --- Kc --------------------------------------------------------------------
yy = POS["kc"]; band(yy)
label(yy, "Kc", sub="coeficient de cultiu")
chip(yy, KC_SPAN[0], KC_SPAN[1], C_KC, KC)

# --- Kr --------------------------------------------------------------------
ytop = POS["kr_" + TRACT[0][0]] + HK
ax.text(-0.5, ytop + 0.30, "Kr", ha="right", va="center", fontsize=9.4,
        weight="bold", color=INK)
ax.text(-0.5, ytop + 0.07, "coeficient de reducció per capçada",
        ha="right", va="center", fontsize=7.6, color=MUTED)
for code, nom, kr in TRACT:
    yy = POS["kr_" + code]
    band(yy, HK)
    ax.text(-0.5, yy + HK / 2, f"{nom}  ({code})", ha="right", va="center",
            fontsize=8.0, color=INK)
    fc1, tc1 = KR_RAMP[1.0]
    chip(yy, 8, 11, fc1, "1,0", tc1, h=HK, fs=7.6, r=0.08)
    fc2, tc2 = KR_RAMP[kr]
    chip(yy, 12, 27, fc2, str(kr).replace(".", ","), tc2, h=HK, fs=7.6, r=0.08)

# --- reg -------------------------------------------------------------------
yy = POS["reg"]; band(yy)
label(yy, "Reg aplicat", sub="fracció de la dosi")
for (a, b), txt in REG:
    fc, tc = REG_RAMP[txt]
    chip(yy, a, b, fc, txt if b - a >= 3 else "", tc)
    if b - a < 3:
        ax.text(b + 1.35, yy + H / 2, txt, ha="left", va="center",
                fontsize=7.6, weight="bold", color=INK)

# --- llegenda i notes ------------------------------------------------------
ly = BOTTOM - 0.55
ax.text(0, ly, "Kr:", ha="left", va="center", fontsize=8.0, weight="bold",
        color=INK)
x = 1.6
for kr in (0.7, 1.0, 1.1, 1.4):
    fc, _ = KR_RAMP[kr]
    ax.add_patch(FancyBboxPatch((x, ly - 0.16), 0.75, 0.32,
                                boxstyle="round,pad=0,rounding_size=0.07",
                                fc=fc, ec="white", lw=0.9))
    ax.text(x + 1.0, ly, str(kr).replace(".", ","), ha="left", va="center",
            fontsize=7.8, color=INK)
    x += 2.0
ax.text(x + 0.6, ly, "Reg:", ha="left", va="center", fontsize=8.0,
        weight="bold", color=INK)
x += 2.4
for txt, desc, adv in (("100 %", "sense restricció", 9.2),
                       ("25 %", "deficitari controlat", 0)):
    fc, _ = REG_RAMP[txt]
    ax.add_patch(FancyBboxPatch((x, ly - 0.16), 0.75, 0.32,
                                boxstyle="round,pad=0,rounding_size=0.07",
                                fc=fc, ec="white", lw=0.9))
    ax.text(x + 1.0, ly, f"{txt}, {desc}", ha="left", va="center",
            fontsize=7.8, color=INK)
    x += adv

for k, txt in enumerate([
        "Desbrot i tall en altura el 02/06; desfullat del tractament ES+DV el 20/07, "
        "coincidint amb el verolat.",
        "El Kc es manté constant de la primera setmana de maig a l'última de setembre. "
        "El reg deficitari controlat s'aplica des del verolat fins a la verema."]):
    ax.text(0, BOTTOM - 1.05 - k * 0.33, txt, ha="left", va="center",
            fontsize=7.6, color=MUTED)

fig.savefig(os.path.join(FIGDIR, "fig17_calendari.png"), dpi=230,
            bbox_inches="tight", facecolor="white")
plt.close(fig)
print("fig17_calendari.png",
      Image.open(os.path.join(FIGDIR, "fig17_calendari.png")).size)
