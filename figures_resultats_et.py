# -*- coding: utf-8 -*-
"""
Figures del capítol de resultats sobre evapotranspiració i estat hídric.

  fig27_kc_vs_fipar.png   Kc real contra fIPAR, una sèrie per vol
  fig28_eta_kc.png        ETa i Kc real per tractament al llarg de la campanya
  fig29_psi_cwsi.png      potencial de tija i CWSI per tractament
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from PIL import Image

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "figures")
X = "/tmp/claude-0/-home-user-TFM-PAU/446d5a3f-407d-5d51-8cea-50e8348ad898/xls/"
os.makedirs(OUT, exist_ok=True)

DATES = ["03/06", "30/06", "23/07", "12/08"]
ETIQ = [d + "/2026" for d in DATES]
ORDRE = ["ES+DV", "ES", "EM", "ELL", "C"]
NOM = {"ES+DV": "Esporga severa + desfullat", "ES": "Esporga severa",
       "EM": "Esporga mitjana", "ELL": "Esporga lleugera", "C": "Sense esporga"}
COLOR = {"C": "#2a78d6", "ELL": "#eb6834", "EM": "#1baf7a",
         "ES": "#eda100", "ES+DV": "#e87ba4"}
MARCA = {"C": "o", "ELL": "s", "EM": "^", "ES": "D", "ES+DV": "X"}
# rampa ordinal per als quatre vols, validada: L monòtona
COLVOL = ["#bcd2ee", "#7aa6d8", "#3d6fb0", "#143d73"]
MARCVOL = ["o", "s", "^", "D"]

INK, MUTED, RULE = "#1a2230", "#52514e", "#d4d7dc"
plt.rcParams.update({"font.family": "DejaVu Sans", "text.color": INK,
                     "axes.labelcolor": INK, "xtick.color": MUTED,
                     "ytick.color": MUTED, "axes.edgecolor": RULE})

m = pd.read_csv(X + "merged.csv")
m["kc_real"] = m.ETa_DAILY_2T / m.Eto
PV = {v: m.pivot_table(index="t", columns="data", values=v).loc[ORDRE, DATES]
      for v in ("kc_real", "fipar_pct", "ETa_DAILY_2T", "CWSI")}
SD = m.pivot_table(index="t", columns="data", values="ETa_DAILY_2T",
                   aggfunc="std").loc[ORDRE, DATES]
PSI = pd.DataFrame(
    {"03/06": [-4.325, -4.725, -4.65, -4.675, -5.075],
     "30/06": [-7.075, -7.025, -7.275, -7.188, -7.625],
     "23/07": [-9.088, -9.863, -9.287, -8.275, -9.85],
     "12/08": [-9.586, -10.5, -11.03, -10.43, -9.6]}, index=ORDRE) / 10.0


def desa(fig, nom):
    fig.tight_layout()
    p = os.path.join(OUT, nom)
    fig.savefig(p, dpi=230, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(nom, Image.open(p).size)


def net(ax, eixy=True):
    ax.grid(True, axis="y" if eixy else "both", color=RULE, lw=0.55, alpha=0.8)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)


# --- 27. Kc real contra fIPAR, una sèrie per vol
fig, ax = plt.subplots(figsize=(7.4, 5.0))
x = PV["fipar_pct"].values / 100.0
y = PV["kc_real"].values
for j, d in enumerate(DATES):
    ax.plot(x[:, j], y[:, j], ls="none", marker=MARCVOL[j], ms=8,
            color=COLVOL[j], mec="white", mew=1.2, label=ETIQ[j], zorder=3)
xf, yf = x.ravel(), y.ravel()
a, b = np.polyfit(xf, yf, 1)
r2 = np.corrcoef(xf, yf)[0, 1] ** 2
xx = np.linspace(0.20, 0.58, 50)
ax.plot(xx, a * xx + b, color=MUTED, lw=1.3, ls=(0, (5, 3)), zorder=2,
        label="conjunt dels quatre vols")
ax.text(0.575, 0.462,
        (f"Kc = {a:.3f}·fIPAR + {b:.3f}\nR² = {r2:.3f}").replace(".", ","),
        fontsize=9, color=MUTED, ha="right", va="bottom")
ax.set_xlabel("fIPAR (–)", fontsize=10.5, labelpad=8)
ax.set_ylabel("Kc real (ETa / ETo)", fontsize=10.5, labelpad=8)
ax.set_xlim(0.20, 0.58)
ax.set_ylim(0.45, 0.80)
net(ax, eixy=False)
ax.legend(frameon=False, fontsize=9, loc="upper left")
desa(fig, "fig27_kc_vs_fipar.png")

# --- 28. ETa i Kc real al llarg de la campanya
fig, axs = plt.subplots(1, 2, figsize=(11.4, 4.4))
xi = np.arange(4)
for ax, key, ylab, lim in (
        (axs[0], "ETa_DAILY_2T", "ETa (mm d⁻¹)", (2.8, 5.3)),
        (axs[1], "kc_real", "Kc real (ETa / ETo)", (0.45, 0.80))):
    for tr in ORDRE:
        v = PV[key].loc[tr].values
        if key == "ETa_DAILY_2T":
            ax.errorbar(xi, v, yerr=SD.loc[tr].values, color=COLOR[tr], lw=1.9,
                        marker=MARCA[tr], ms=6, mec="white", mew=1.2,
                        capsize=3, elinewidth=0.9, label=NOM[tr], zorder=3)
        else:
            ax.plot(xi, v, color=COLOR[tr], lw=1.9, marker=MARCA[tr], ms=6,
                    mec="white", mew=1.2, label=NOM[tr], zorder=3)
    ax.set_xticks(xi)
    ax.set_xticklabels([d for d in DATES], fontsize=9)
    ax.set_ylabel(ylab, fontsize=10.5, labelpad=8)
    ax.set_ylim(*lim)
    net(ax)
h, l = axs[0].get_legend_handles_labels()
fig.legend(h, l, frameon=False, fontsize=9, ncol=5, loc="lower center",
           bbox_to_anchor=(0.5, -0.06))
desa(fig, "fig28_eta_kc.png")

# --- 29. potencial de tija i CWSI
fig, axs = plt.subplots(1, 2, figsize=(11.4, 4.4))
for ax, dat, ylab, lim in ((axs[0], PSI, "Potencial de tija (MPa)", (-1.20, -0.35)),
                           (axs[1], PV["CWSI"], "CWSI (–)", (0.0, 0.42))):
    for tr in ORDRE:
        ax.plot(xi, dat.loc[tr].values, color=COLOR[tr], lw=1.9,
                marker=MARCA[tr], ms=6, mec="white", mew=1.2, label=NOM[tr],
                zorder=3)
    ax.set_xticks(xi)
    ax.set_xticklabels([d for d in DATES], fontsize=9)
    ax.set_ylabel(ylab, fontsize=10.5, labelpad=8)
    ax.set_ylim(*lim)
    net(ax)
    ax.axvline(2.5, color="#c0392b", lw=1.0, ls=(0, (4, 3)), zorder=2)
axs[1].text(2.45, 0.40, "reg deficitari", fontsize=8.6, color="#c0392b",
            ha="right", va="top")
h, l = axs[0].get_legend_handles_labels()
fig.legend(h, l, frameon=False, fontsize=9, ncol=5, loc="lower center",
           bbox_to_anchor=(0.5, -0.06))
desa(fig, "fig29_psi_cwsi.png")
