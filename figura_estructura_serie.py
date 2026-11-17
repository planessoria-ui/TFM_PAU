# -*- coding: utf-8 -*-
"""
Figura 16: evolució de la fracció de coberta i de l'alçada del dosser al llarg
dels quatre vols de 2026, per tractament de maneig de capçada.

Forma: canvi al llarg del temps amb identitat de sèrie -> línies amb marcadors,
dos panells (no un doble eix, que són magnituds d'escala diferent).
Color: paleta categòrica validada (slots 1-5), amb marcadors de forma diferent
com a codificació secundària i etiquetes directes; la Taula 14 fa de vista de
taula per a la regla de relleu del contrast.
"""
import csv, json, os, statistics as st, datetime as dt
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from PIL import Image

U = "/root/.claude/uploads/446d5a3f-407d-5d51-8cea-50e8348ad898/"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "figures")
os.makedirs(OUT, exist_ok=True)

F = [("03/06", dt.date(2026, 6, 3), "a290cdc2-biophysical_traits_YIELD4CAST2026_covercrop_P25_v2.csv"),
     ("30/06", dt.date(2026, 6, 30), "57255596-biophysical_traits_YIELD4CAST2026_covercrop_P25.csv"),
     ("22/07", dt.date(2026, 7, 22), "7e8f3dc6-biophysical_traits_YIELD4CAST2026_covercrop_P25_22072026.csv"),
     ("12/08", dt.date(2026, 8, 12), "9b33ef42-biophysical_traits_YIELD4CAST2026_covercrop_P25_12082026.csv")]
SHORT = {"Esporga severa i desfullat a verolat": "ES+DV", "Esporga severa + desfullat": "ES+DV",
         "Esporga severa": "ES", "Esporga mitjana": "EM",
         "Esporga lleugera": "EL", "Sense esporga": "SE"}
ORDER = ["SE", "EL", "EM", "ES", "ES+DV"]
LABEL = {"SE": "Sense esporga", "EL": "Esporga lleugera", "EM": "Esporga mitjana",
         "ES": "Esporga severa", "ES+DV": "Esporga severa + desfullat"}

# paleta categòrica validada (slots 1-5) + marcadors com a codificació secundària
SERIES = {"SE": ("#2a78d6", "o"), "EL": ("#eb6834", "s"), "EM": ("#1baf7a", "^"),
          "ES": ("#eda100", "D"), "ES+DV": ("#e87ba4", "v")}
INK, MUTED, RULE = "#1a2230", "#52514e", "#d4d7dc"


def f2(s):
    s = (s or "").strip().replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


rows = []
for lab, day, fn in F:
    raw = open(U + fn, encoding="utf-8-sig").read()
    d = ";" if raw.count(";") > raw.count(",") else ","
    for r in csv.DictReader(raw.splitlines(), delimiter=d):
        if not (r.get("tree") or "").strip():
            continue
        rows.append({"day": day, "trt": SHORT[r["Treatment"].strip()],
                     "fc": f2(r.get("fc")), "hc": f2(r.get("hc"))})

plt.rcParams.update({"font.family": "DejaVu Sans", "text.color": INK,
                     "axes.labelcolor": INK, "xtick.color": MUTED,
                     "ytick.color": MUTED, "axes.edgecolor": RULE})

fig, axes = plt.subplots(1, 2, figsize=(9.4, 4.3))
days = [d for _, d, _ in F]

for ax, var, ylab, ylim in [(axes[0], "fc", "Fracció de coberta (–)", (0.09, 0.48)),
                            (axes[1], "hc", "Alçada del dosser (m)", (1.58, 2.60))]:
    for trt in ORDER:
        col, mk = SERIES[trt]
        vals = [[r[var] for r in rows if r["day"] == d and r["trt"] == trt
                 and r[var] is not None] for d in days]
        ys = [st.mean(v) for v in vals]
        sd = [st.pstdev(v) for v in vals]
        ax.errorbar(days, ys, yerr=sd, fmt="none", ecolor=col, elinewidth=1.0,
                    capsize=2.6, capthick=1.0, alpha=0.55, zorder=2)
        ax.plot(days, ys, color=col, lw=2.0, marker=mk, ms=7,
                mec="white", mew=1.2, zorder=3, label=LABEL[trt], clip_on=False)
        ax.annotate(trt, xy=(days[-1], ys[-1]), xytext=(7, 0),
                    textcoords="offset points", va="center", ha="left",
                    fontsize=8.2, color=INK, zorder=5)
    ax.set_ylabel(ylab, fontsize=10, labelpad=8)
    ax.set_ylim(*ylim)
    ax.set_xticks(days)
    ax.set_xticklabels([l for l, _, _ in F], fontsize=9)
    ax.set_xlim(days[0] - dt.timedelta(days=4), days[-1] + dt.timedelta(days=15))
    ax.grid(True, axis="y", color=RULE, lw=0.6)
    ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.tick_params(length=3)

axes[0].text(0.0, 1.06, "(A)", transform=axes[0].transAxes, fontsize=10,
             weight="bold", color=INK)
axes[1].text(0.0, 1.06, "(B)", transform=axes[1].transAxes, fontsize=10,
             weight="bold", color=INK)

handles, labels = axes[0].get_legend_handles_labels()
fig.legend(handles, labels, loc="lower center", ncol=5, frameon=False,
           fontsize=8.6, handlelength=2.2, columnspacing=1.6,
           bbox_to_anchor=(0.5, -0.035))
fig.tight_layout(rect=[0, 0.07, 1, 1])
p = os.path.join(OUT, "fig16_estructura_serie.png")
fig.savefig(p, dpi=220, bbox_inches="tight", facecolor="white")
plt.close(fig)
print("fig16_estructura_serie.png", Image.open(p).size)
