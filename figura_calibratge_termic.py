# -*- coding: utf-8 -*-
"""
Recta de calibratge de la banda tèrmica del vol del 3 de juny de 2026,
ajustada sobre els cinc elements de referència del full Temp_CAL.

Una sola sèrie, de manera que no porta llegenda: el títol del peu la nomena.
Cada punt va etiquetat amb l'element que representa.
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from PIL import Image

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "figures")
os.makedirs(OUT, exist_ok=True)

INK, MUTED, RULE, ACC = "#1a2230", "#52514e", "#d4d7dc", "#2a78d6"

# full Temp_CAL, vol del 03/06/2026
PUNTS = [("bk", "panell negre", 32444.728784397, 55.3895),
         ("wh", "panell blanc", 31172.4420239519, 32.9339),
         ("bs", "sòl nu", 31687.8572668651, 43.0701),
         ("cc", "coberta de l'entrefilera", 31554.5809966739, 37.1797),
         ("vg", "vegetació", 30633.5962245904, 26.0561)]
# desplaçament de l'etiqueta de cada punt, en punts tipogràfics
DESP = {"bk": (-12, 0), "wh": (12, 0), "bs": (12, 0), "cc": (-12, 0), "vg": (12, 0)}

x = np.array([p[2] for p in PUNTS])
y = np.array([p[3] for p in PUNTS])
m, n = np.polyfit(x, y, 1)
pred = m * x + n
r2 = 1 - ((y - pred) ** 2).sum() / ((y - y.mean()) ** 2).sum()

plt.rcParams.update({"font.family": "DejaVu Sans", "text.color": INK,
                     "axes.labelcolor": INK, "xtick.color": MUTED,
                     "ytick.color": MUTED, "axes.edgecolor": RULE})

fig, ax = plt.subplots(figsize=(7.0, 4.6))

xx = np.linspace(x.min() - 150, x.max() + 150, 100)
ax.plot(xx, m * xx + n, color=ACC, lw=1.8, zorder=2)
ax.plot(x, y, "o", ms=9, color=ACC, mec="white", mew=1.5, zorder=3, clip_on=False)

for codi, _nom, xi, yi in PUNTS:
    dx, dy = DESP[codi]
    ax.annotate(codi, xy=(xi, yi), xytext=(dx, dy), textcoords="offset points",
                fontsize=9, color=INK, weight="bold",
                ha="right" if dx < 0 else "left", va="center")

ax.text(0.97, 0.10,
        f"T = {m:.4f} · B6 − {abs(n):.2f}".replace(".", ","),
        transform=ax.transAxes, ha="right", va="bottom", fontsize=10.5, color=INK)
ax.text(0.97, 0.03, f"R² = {r2:.4f}".replace(".", ","),
        transform=ax.transAxes, ha="right", va="bottom", fontsize=10.5, color=MUTED)

ax.set_xlabel("Valor de la banda 6 de l'ortomosaic (B6)", fontsize=10.5, labelpad=8)
ax.set_ylabel("Temperatura mesurada amb l'IRT (°C)", fontsize=10.5, labelpad=8)
ax.set_xlim(30450, 32650)
ax.set_ylim(22, 60)
ax.set_xticks(range(30500, 32700, 400))
ax.get_xaxis().set_major_formatter(
    matplotlib.ticker.FuncFormatter(lambda v, _p: f"{int(v):,}".replace(",", ".")))
ax.grid(True, color=RULE, lw=0.55, alpha=0.75)
ax.set_axisbelow(True)
for s in ("top", "right"):
    ax.spines[s].set_visible(False)

fig.tight_layout()
p = os.path.join(OUT, "fig_calibratge_termic_0306.png")
fig.savefig(p, dpi=230, bbox_inches="tight", facecolor="white")
plt.close(fig)
print("fig_calibratge_termic_0306.png", Image.open(p).size)
print(f"ajust: T = {m:.6f}·B6 {n:+.2f}   R2 = {r2:.4f}")
