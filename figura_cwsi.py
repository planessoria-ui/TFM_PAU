# -*- coding: utf-8 -*-
"""
Figura 3: fonament del càlcul del CWSI.

Representa els dos límits de referència de la diferència entre la temperatura
del dosser i la de l'aire en funció del dèficit de pressió de vapor, tal com
els defineixen Idso et al. (1981) i Jackson et al. (1981), i la construcció
geomètrica de l'índex. Les rectes són il·lustratives: els coeficients de la
recta de referència depenen del cultiu, de la varietat i de l'estat fenològic.
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from PIL import Image

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "figures")
os.makedirs(OUT, exist_ok=True)

INK, MUTED, RULE = "#1a2230", "#52514e", "#d4d7dc"
LOW, UP = "#2a78d6", "#eb6834"        # paleta categòrica validada, slots 1 i 2

plt.rcParams.update({"font.family": "DejaVu Sans", "text.color": INK,
                     "axes.labelcolor": INK, "xtick.color": MUTED,
                     "ytick.color": MUTED, "axes.edgecolor": RULE})

fig, ax = plt.subplots(figsize=(7.6, 5.0))

vpd = np.linspace(0.5, 5.0, 200)
lower = 2.6 - 1.55 * vpd          # dosser ben regat, transpiració màxima
upper = np.full_like(vpd, 3.1)    # dosser que no transpira

ax.fill_between(vpd, lower, upper, color=LOW, alpha=0.05, zorder=1)
ax.plot(vpd, upper, color=UP, lw=2.2, zorder=3)
ax.plot(vpd, lower, color=LOW, lw=2.2, zorder=3)

# etiquetes directes als dos límits, dins de l'àrea del gràfic
ax.text(0.68, 3.28, "Límit superior · dosser que no transpira", ha="left",
        va="bottom", fontsize=9.4, color=UP, weight="bold")
ax.text(0.68, 3.28 - 0.52, "CWSI = 1", ha="left", va="bottom",
        fontsize=8.8, color=UP)
ax.text(3.72, 2.6 - 1.55 * 3.72 - 0.34, "Límit inferior · cultiu ben regat",
        ha="right", va="top", fontsize=9.4, color=LOW, weight="bold",
        rotation=-18.5, rotation_mode="anchor")
ax.text(3.80, 2.6 - 1.55 * 3.80 + 0.18, "CWSI = 0", ha="left", va="bottom",
        fontsize=8.8, color=LOW, rotation=-18.5, rotation_mode="anchor")

# construcció de l'índex sobre una mesura concreta
x0 = 2.45
y_lo = 2.6 - 1.55 * x0
y_up = 3.1
y_m = y_lo + 0.52 * (y_up - y_lo)

ax.plot([x0, x0], [y_lo, y_up], color=MUTED, lw=0.8, ls=(0, (3, 3)), zorder=2)
ax.annotate("", xy=(x0, y_m), xytext=(x0, y_lo),
            arrowprops=dict(arrowstyle="<->", color=INK, lw=1.1))
ax.annotate("", xy=(x0, y_up), xytext=(x0, y_m),
            arrowprops=dict(arrowstyle="<->", color=INK, lw=1.1))
ax.text(x0 - 0.13, (y_lo + y_m) / 2, "a", ha="right", va="center",
        fontsize=11.5, style="italic", color=INK)
ax.text(x0 - 0.13, (y_m + y_up) / 2, "b", ha="right", va="center",
        fontsize=11.5, style="italic", color=INK)

ax.plot([x0], [y_m], "o", ms=8, color=INK, mec="white", mew=1.4, zorder=5)
ax.annotate("(Tc − Ta) mesurat", xy=(x0 + 0.08, y_m), xytext=(x0 + 0.62, y_m + 0.95),
            fontsize=9.4, color=INK, ha="left", va="center",
            arrowprops=dict(arrowstyle="-", color=MUTED, lw=0.9,
                            shrinkA=3, shrinkB=5))

ax.text(0.68, -4.55,
        "CWSI =  a / (a + b)  =  1 − ETa / ETp",
        ha="left", va="center", fontsize=11, color=INK, zorder=6,
        bbox=dict(boxstyle="round,pad=0.5", fc="white", ec=RULE, lw=0.9))

ax.set_xlabel("Dèficit de pressió de vapor, VPD (kPa)", fontsize=10.5, labelpad=8)
ax.set_ylabel("Tc − Ta (°C)", fontsize=10.5, labelpad=8)
ax.set_xlim(0.5, 5.0)
ax.set_ylim(-5.4, 4.3)
ax.grid(True, color=RULE, lw=0.55, alpha=0.7)
ax.set_axisbelow(True)
for s in ("top", "right"):
    ax.spines[s].set_visible(False)

fig.tight_layout()
p = os.path.join(OUT, "fig3_cwsi.png")
fig.savefig(p, dpi=230, bbox_inches="tight", facecolor="white")
plt.close(fig)
print("fig3_cwsi.png", Image.open(p).size)
