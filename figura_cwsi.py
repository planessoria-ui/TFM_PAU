# -*- coding: utf-8 -*-
"""
Figura 3: límits de referència del CWSI en Tempranillo.

Les quatre rectes són les publicades a la Taula 3 de Bellvert, J., Marsal, J.,
Girona, J., i Zarco-Tejada, P. J. (2015), "Seasonal evolution of crop water
stress index in grapevine varieties determined with high-resolution remote
sensing thermal imagery", Irrigation Science, 33(2), 81-93,
https://doi.org/10.1007/s00271-014-0456-y, obtingudes en vinyes de Raimat.

  Tempranillo, estadi I      LL: y = -1,438x + 1,098   UL: y = 0,261x + 3,977
  Tempranillo, estadi II-III LL: y = -1,780x + 1,253   UL: y = 0,466x + 5,317

on y és Tc - Ta en graus Celsius i x el dèficit de pressió de vapor en kPa.
Els autors les van desenvolupar per a un VPD d'entre 1 i 5 kPa.
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
LOW, UP, PAST = "#2a78d6", "#eb6834", "#9aa3ae"

# Taula 3 de Bellvert et al. (2015), varietat Tempranillo
LL_I, UL_I = (-1.438, 1.098), (0.261, 3.977)
LL_II, UL_II = (-1.780, 1.253), (0.466, 5.317)

plt.rcParams.update({"font.family": "DejaVu Sans", "text.color": INK,
                     "axes.labelcolor": INK, "xtick.color": MUTED,
                     "ytick.color": MUTED, "axes.edgecolor": RULE})

fig, ax = plt.subplots(figsize=(7.8, 5.4))
x = np.linspace(1.0, 5.0, 200)


def line(c):
    return c[0] * x + c[1]


# estadi I, de referència, en segon pla
ax.plot(x, line(LL_I), color=PAST, lw=1.3, ls=(0, (5, 3)), zorder=2)
ax.plot(x, line(UL_I), color=PAST, lw=1.3, ls=(0, (5, 3)), zorder=2)
ax.text(4.30, LL_I[0] * 4.30 + LL_I[1] + 0.22, "estadi I", fontsize=8.4,
        color=PAST, style="italic", rotation=-17.0, rotation_mode="anchor")
ax.text(2.00, UL_I[0] * 2.00 + UL_I[1] + 0.24, "estadi I", fontsize=8.4,
        color=PAST, style="italic", rotation=2.6, rotation_mode="anchor")

# estadi II-III, el del període de l'assaig
ax.fill_between(x, line(LL_II), line(UL_II), color=LOW, alpha=0.055, zorder=1)
ax.plot(x, line(UL_II), color=UP, lw=2.3, zorder=3)
ax.plot(x, line(LL_II), color=LOW, lw=2.3, zorder=3)

ax.text(1.12, UL_II[0] * 1.12 + UL_II[1] + 0.34,
        "Límit superior · dosser que no transpira",
        fontsize=9.3, color=UP, weight="bold", rotation=4.6,
        rotation_mode="anchor")
ax.text(1.12, LL_II[0] * 1.12 + LL_II[1] - 0.46,
        "Límit inferior · transpiració potencial",
        fontsize=9.3, color=LOW, weight="bold", rotation=-17.0,
        rotation_mode="anchor")
ax.text(4.96, UL_II[0] * 4.96 + UL_II[1] + 0.26,
        "CWSI = 1   ·   y = 0,466x + 5,317",
        fontsize=8.6, color=UP, ha="right", va="bottom")
ax.text(4.96, LL_II[0] * 4.96 + LL_II[1] - 0.26,
        "CWSI = 0   ·   y = −1,780x + 1,253",
        fontsize=8.6, color=LOW, ha="right", va="top")

# construcció de l'índex sobre una mesura concreta
x0 = 3.0
y_lo = LL_II[0] * x0 + LL_II[1]
y_up = UL_II[0] * x0 + UL_II[1]
y_m = y_lo + 0.35 * (y_up - y_lo)

ax.plot([x0, x0], [y_lo, y_up], color=MUTED, lw=0.8, ls=(0, (3, 3)), zorder=2)
ax.annotate("", xy=(x0, y_m), xytext=(x0, y_lo),
            arrowprops=dict(arrowstyle="<->", color=INK, lw=1.1))
ax.annotate("", xy=(x0, y_up), xytext=(x0, y_m),
            arrowprops=dict(arrowstyle="<->", color=INK, lw=1.1))
ax.text(x0 - 0.10, (y_lo + y_m) / 2, "a", ha="right", va="center",
        fontsize=12, style="italic", color=INK)
ax.text(x0 - 0.10, (y_m + y_up) / 2, "b", ha="right", va="center",
        fontsize=12, style="italic", color=INK)
ax.plot([x0], [y_m], "o", ms=8, color=INK, mec="white", mew=1.4, zorder=5)
ax.annotate("(Tc − Ta) mesurat", xy=(x0 + 0.07, y_m), xytext=(x0 + 0.55, y_m + 1.6),
            fontsize=9.4, color=INK, ha="left", va="center",
            arrowprops=dict(arrowstyle="-", color=MUTED, lw=0.9,
                            shrinkA=3, shrinkB=5))

ax.text(1.12, -7.3, "CWSI  =  a / (a + b)  =  1 − ETa / ETp",
        ha="left", va="center", fontsize=11, color=INK, zorder=6,
        bbox=dict(boxstyle="round,pad=0.5", fc="white", ec=RULE, lw=0.9))

ax.set_xlabel("Dèficit de pressió de vapor, VPD (kPa)", fontsize=10.5, labelpad=8)
ax.set_ylabel("Tc − Ta (°C)", fontsize=10.5, labelpad=8)
ax.set_xlim(1.0, 5.0)
ax.set_ylim(-8.6, 8.4)
ax.set_yticks(range(-8, 9, 2))
ax.grid(True, color=RULE, lw=0.55, alpha=0.7)
ax.set_axisbelow(True)
for s in ("top", "right"):
    ax.spines[s].set_visible(False)

fig.tight_layout()
p = os.path.join(OUT, "fig3_cwsi.png")
fig.savefig(p, dpi=230, bbox_inches="tight", facecolor="white")
plt.close(fig)
print("fig3_cwsi.png", Image.open(p).size)
