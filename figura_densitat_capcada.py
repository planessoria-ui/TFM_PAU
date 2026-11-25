# -*- coding: utf-8 -*-
"""
Les tres figures de la relació entre fracció de coberta i àrea foliar, en PNG,
amb els mateixos colors per tractament que la resta de figures del treball.

  fig_densitat_paret.png   densitat LAI/fc per tractament i vol
  fig_lai_vs_fc.png        núvol de LAI contra fc
  fig_reg_vs_consum.png    dosi de reg, estructura i consum, relatius al control
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
D = "/tmp/claude-0/-home-user-TFM-PAU/446d5a3f-407d-5d51-8cea-50e8348ad898/csvdata/"
os.makedirs(OUT, exist_ok=True)

FS = {"03/06": "790d17a6-inputs_03062026.csv", "30/06": "c64943ce-inputs_30062026.csv",
      "23/07": "4a8253d5-inputs_23072026.csv", "12/08": "244b6521-inputs_12082026.csv"}
DATES = list(FS)
ORDRE = ["ES+DV", "ES", "EM", "ELL", "C"]
NOM = {"ES+DV": "Esporga severa + desfullat", "ES": "Esporga severa",
       "EM": "Esporga mitjana", "ELL": "Esporga lleugera", "C": "Sense esporga"}
COLOR = {"C": "#2a78d6", "ELL": "#eb6834", "EM": "#1baf7a",
         "ES": "#eda100", "ES+DV": "#e87ba4"}
MARCA = {"C": "o", "ELL": "s", "EM": "^", "ES": "D", "ES+DV": "X"}
KR = {"ES+DV": 0.70, "ES": 0.70, "EM": 1.00, "ELL": 1.10, "C": 1.40}

INK, MUTED, RULE = "#1a2230", "#52514e", "#d4d7dc"
plt.rcParams.update({"font.family": "DejaVu Sans", "text.color": INK,
                     "axes.labelcolor": INK, "xtick.color": MUTED,
                     "ytick.color": MUTED, "axes.edgecolor": RULE})

m = pd.read_csv(X + "merged.csv")
inp = pd.concat([pd.read_csv(D + f, index_col=0).assign(data=d, id=range(1, 119))
                 for d, f in FS.items()])
m = m.merge(inp[["data", "id", "fc"]], on=["data", "id"])
g = (m.groupby(["t", "data"])
       .agg(fc=("fc", "mean"), LAI=("lai_bo", "mean"), ETa=("ETa_DAILY_2T", "mean"))
       .reset_index())
P = {v: g.pivot_table(index="t", columns="data", values=v).loc[ORDRE, DATES]
     for v in ("fc", "LAI", "ETa")}
DEN = P["LAI"] / P["fc"]
camp = m.groupby("t").agg(fc=("fc", "mean"), LAI=("lai_bo", "mean"),
                          ETa=("ETa_DAILY_2T", "mean")).loc[ORDRE]


def desa(fig, nom):
    fig.tight_layout()
    p = os.path.join(OUT, nom)
    fig.savefig(p, dpi=230, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(nom, Image.open(p).size)


def net(ax):
    ax.grid(True, axis="y", color=RULE, lw=0.55, alpha=0.8)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)


# --- 1. densitat de la paret vegetal
fig, ax = plt.subplots(figsize=(7.6, 4.8))
x = np.arange(4)
for tr in ORDRE:
    ax.plot(x, DEN.loc[tr].values, color=COLOR[tr], lw=2.0, marker=MARCA[tr],
            ms=7, mec="white", mew=1.3, label=NOM[tr], zorder=3)
ax.set_xticks(x)
ax.set_xticklabels([d + "/2026" for d in DATES])
ax.set_ylabel("LAI / fc  (m² de fulla per m² de sòl cobert)", fontsize=10.5, labelpad=8)
ax.set_ylim(5, 11)
ax.set_yticks(range(5, 12))
net(ax)
ax.legend(frameon=False, fontsize=9, ncol=3, loc="upper center",
          bbox_to_anchor=(0.5, -0.11))
desa(fig, "fig_densitat_paret.png")

# --- 2. LAI contra fc
fig, ax = plt.subplots(figsize=(7.0, 5.0))
for tr in ORDRE:
    ax.plot(P["fc"].loc[tr].values, P["LAI"].loc[tr].values, ls="none",
            marker=MARCA[tr], ms=9, color=COLOR[tr], mec="white", mew=1.3,
            label=NOM[tr], zorder=3)
xx = np.linspace(0.15, 0.45, 50)
a, b = np.polyfit(g.fc, g.LAI, 1)
r2 = np.corrcoef(g.fc, g.LAI)[0, 1] ** 2
ax.plot(xx, a * xx + b, color=MUTED, lw=1.1, ls=(0, (5, 3)), zorder=2)
ax.text(0.435, a * 0.435 + b - 0.07,
        f"LAI = {a:.2f}·fc + {b:.2f}".replace(".", ",") + f"\nR² = {r2:.2f}".replace(".", ","),
        fontsize=8.6, color=MUTED, ha="right", va="top")
ax.set_xlabel("Fracció de coberta, fc (–)", fontsize=10.5, labelpad=8)
ax.set_ylabel("Índex d'àrea foliar, LAI (m² m⁻²)", fontsize=10.5, labelpad=8)
ax.set_xlim(0.15, 0.45)
ax.set_ylim(1.6, 2.6)
ax.grid(True, color=RULE, lw=0.55, alpha=0.8)
ax.set_axisbelow(True)
for s in ("top", "right"):
    ax.spines[s].set_visible(False)
ax.legend(frameon=False, fontsize=9, loc="upper left")
desa(fig, "fig_lai_vs_fc.png")

# --- 3. reg, estructura i consum, relatius al control
fig, ax = plt.subplots(figsize=(8.4, 4.6))
SERIES = [("Dosi de reg (Kr)", [KR[t] / KR["C"] for t in ORDRE], "#8c96a3"),
          ("Fracció de coberta", (camp.fc / camp.loc["C", "fc"]).values, "#2a78d6"),
          ("Àrea foliar", (camp.LAI / camp.loc["C", "LAI"]).values, "#1baf7a"),
          ("Consum d'aigua", (camp.ETa / camp.loc["C", "ETa"]).values, "#eb6834")]
x = np.arange(len(ORDRE))
w = 0.2
for k, (lab, v, col) in enumerate(SERIES):
    ax.bar(x + (k - 1.5) * w, v, w, color=col, label=lab, zorder=3)
ax.set_xticks(x)
ax.set_xticklabels([NOM[t].replace(" + ", "\n+ ") for t in ORDRE], fontsize=9)
ax.set_ylabel("Valor relatiu al tractament sense esporga", fontsize=10.5, labelpad=8)
ax.set_ylim(0, 1.15)
ax.set_yticks(np.arange(0, 1.2, 0.2))
net(ax)
ax.legend(frameon=False, fontsize=9, ncol=4, loc="upper center",
          bbox_to_anchor=(0.5, 1.12))
desa(fig, "fig_reg_vs_consum.png")
