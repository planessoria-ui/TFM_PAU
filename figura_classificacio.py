# -*- coding: utf-8 -*-
"""
Figura de la classificació de píxels del processament ET4DROUGHT: llindar
d'alçada (0,5 m) per separar dosser i sòl, i llindar de SAVI (0,1) per quedar-se
amb els píxels purs de cada classe. Estil sobri, coherent amb la resta de
figures del TFM.
"""
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch
from PIL import Image

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "figures")
os.makedirs(OUT, exist_ok=True)

INK = '#1a2230'
MUTED = '#5b6675'
RULE = '#b9c0c9'
FILL = '#f4f6f8'
ACCENT = '#1f5c8b'
WARM = '#9a6b1f'

plt.rcParams.update({
    'font.family': 'DejaVu Sans',
    'text.color': INK,
})

fig, ax = plt.subplots(figsize=(7.4, 7.0))
ax.set_xlim(0, 10)
ax.set_ylim(0, 10.2)
ax.axis('off')


def box(x, y, w, h, title, body=None, edge=RULE, face=FILL, bar=None, fs=9.4):
    ax.add_patch(FancyBboxPatch((x, y), w, h,
                 boxstyle="round,pad=0.015,rounding_size=0.09",
                 fc=face, ec=edge, lw=0.95, zorder=2))
    if bar:
        ax.add_patch(FancyBboxPatch((x, y), 0.075, h,
                     boxstyle="square,pad=0", fc=bar, ec='none', zorder=3))
    ty = y + h - 0.30 if body else y + h / 2
    ax.text(x + w / 2, ty, title, ha='center',
            va='center' if not body else 'top',
            fontsize=fs, weight='bold', color=INK, zorder=4)
    if body:
        ax.text(x + w / 2, y + h - 0.62, body, ha='center', va='top',
                fontsize=8.2, color=MUTED, linespacing=1.5, zorder=4)


def arrow(x0, y0, x1, y1, label=None, lx=0.0, italic=True):
    ax.add_patch(FancyArrowPatch((x0, y0), (x1, y1), arrowstyle='-|>',
                 mutation_scale=11, color=RULE, lw=1.1, zorder=1))
    if label:
        ax.text((x0 + x1) / 2 + lx, (y0 + y1) / 2, label, ha='center',
                va='center', fontsize=8.4, color=ACCENT,
                style='italic' if italic else 'normal', zorder=5,
                bbox=dict(boxstyle='round,pad=0.26', fc='white',
                          ec='none'))


# --- entrada ---------------------------------------------------------------
box(0.9, 8.85, 8.2, 1.20,
    "Capes del vol remostrejades a la graella del NIR",
    "cinc bandes de reflectància · model digital de vegetació (hc)\ntemperatura de superfície (LST)",
    bar=ACCENT)

# --- primer llindar: alçada ------------------------------------------------
arrow(5.0, 8.85, 5.0, 8.20)
ax.text(5.0, 8.00, "Llindar d'alçada:  hc = 0,5 m", ha='center', va='center',
        fontsize=9.0, color=INK, weight='bold', zorder=5)

arrow(5.0, 7.80, 2.9, 7.05, "hc ≥ 0,5 m", lx=-0.55)
arrow(5.0, 7.80, 7.1, 7.05, "hc < 0,5 m", lx=0.55)

box(1.0, 6.15, 3.8, 0.90, "CANOPY", "píxels de dosser")
box(5.2, 6.15, 3.8, 0.90, "SOIL", "píxels d'entrefilera")

# --- segon llindar: SAVI ---------------------------------------------------
arrow(2.9, 6.15, 2.9, 5.25, "SAVI ≥ 0,1", lx=-0.92)
arrow(7.1, 6.15, 7.1, 5.25, "SAVI < 0,1", lx=0.92)

box(1.0, 4.35, 3.8, 0.90, "CANOPY_GREEN", "dosser fotosintèticament actiu",
    bar=ACCENT)
box(5.2, 4.35, 3.8, 0.90, "SOIL_GREEN", "sòl i coberta de l'entrefilera",
    bar=WARM)

# --- sortides --------------------------------------------------------------
arrow(2.9, 4.35, 2.9, 3.60)
arrow(7.1, 4.35, 7.1, 3.60)

box(0.55, 1.05, 4.25, 2.55, "Variables del dosser",
    "NDVI, SAVI, MTVI, NDWI i CIG\n"
    "(mitjana i desviació per cep)\n"
    "alçada màxima (hc)\n"
    "àrea i volum de copa\n"
    "fracció de coberta (fc)\n"
    "temperatura de dosser (Tc)", fs=9.0)
box(5.20, 1.05, 4.25, 2.55, "Variables de l'entrefilera",
    "NDVI, SAVI, MTVI, NDWI i CIG\n"
    "de la superfície del carrer\n"
    "àrea de sòl i de coberta\n"
    "fracció de coberta del carrer\n"
    "temperatura de sòl (Ts)", fs=9.0)

# --- nota de peu -----------------------------------------------------------
ax.text(5.0, 0.48, "Estadístics zonals calculats cep a cep sobre la capa de polígons TREES",
        ha='center', va='center', fontsize=8.4, color=MUTED, style='italic')

fig.savefig(os.path.join(OUT, 'fig15_classificacio.png'), dpi=220,
            bbox_inches='tight', facecolor='white')
plt.close(fig)
im = Image.open(os.path.join(OUT, 'fig15_classificacio.png'))
print('fig15_classificacio.png', im.size)
