# -*- coding: utf-8 -*-
"""
Genera el full de càlcul del calendari de maneig de la campanya 2026 i la
figura del calendari per al TFM, a partir de la infografia de l'assaig.

Sortides:
  calendari_maneig_2026.xlsx   full editable amb les dades i un gràfic natiu
  figures/fig17_calendari.png  figura sòbria per al document
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
from PIL import Image

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
FIGDIR = os.path.join(BASE, "figures")
os.makedirs(FIGDIR, exist_ok=True)

MESOS = ["Març", "Abril", "Maig", "Juny", "Juliol", "Agost", "Setembre", "Octubre"]
NSET = 32  # 8 mesos x 4 setmanes

# índex de setmana (0-31) dels esdeveniments
BROTACIO, PODA, VEROLAT, DESFULLAT = 3, 12, 18, 18
VEREMA = (23, 26)
VOLS = [("1r vol\n03/06", 12), ("2n vol\n30/06", 15),
        ("3r vol\n22/07", 19), ("4t vol\n12/08", 21)]
KC_SPAN, KC = (8, 27), 0.45
REG = [((8, 18), "100 %"), ((19, 26), "25 %"), ((27, 27), "100 %")]

TRACT = [("Sense esporga (C)", 1.0, 1.4, 0.365),
         ("Esporga lleugera (ELL)", 1.0, 1.1, 0.319),
         ("Esporga mitjana (EM)", 1.0, 1.0, 0.243),
         ("Esporga severa (ES)", 1.0, 0.7, 0.178),
         ("Esporga severa + desfullat (ES+DV)", 1.0, 0.7, 0.171)]

# ---------------------------------------------------------------------------
# Full de càlcul
# ---------------------------------------------------------------------------
INK = "1A2230"
HEAD = PatternFill("solid", fgColor="1F5C8B")
SUB = PatternFill("solid", fgColor="E8EDF2")
THIN = Side(style="thin", color="C8CDD4")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FILLS = {"fen": "1BAF7A", "tasca": "EB6834", "vol": "2A78D6",
         "kc": "1F5C8B", "reg1": "6250D6", "reg2": "BCB6EE", "kr": "C98E7A"}

wb = openpyxl.Workbook()

# -- full 1: calendari setmanal ---------------------------------------------
ws = wb.active
ws.title = "Calendari"
ws["A1"] = "Calendari de maneig de la vinya · campanya 2026"
ws["A1"].font = Font(name="Calibri", size=14, bold=True, color=INK)
ws["A2"] = ("Assaig ADAPTEX de maneig de capçada · parcel·la El polvorín (Raimat). "
            "Setmanes de març a octubre.")
ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="5B6675")

R0 = 4                      # fila de la capçalera de mesos
C0 = 3                      # primera columna de setmana
ws.column_dimensions["A"].width = 34
ws.column_dimensions["B"].width = 10
for i in range(NSET):
    ws.column_dimensions[get_column_letter(C0 + i)].width = 4.2

for m, mes in enumerate(MESOS):
    c = C0 + m * 4
    ws.merge_cells(start_row=R0, start_column=c, end_row=R0, end_column=c + 3)
    cell = ws.cell(R0, c, mes)
    cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    cell.fill = HEAD
    cell.alignment = Alignment(horizontal="center")
    for s in range(4):
        x = ws.cell(R0 + 1, c + s, s + 1)
        x.font = Font(name="Calibri", size=9, color="5B6675")
        x.alignment = Alignment(horizontal="center")
        x.fill = SUB
ws.cell(R0, 1, "Concepte").font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
ws.cell(R0, 1).fill = HEAD
ws.cell(R0, 2, "Valor").font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
ws.cell(R0, 2).fill = HEAD


def banda(row, label, valor, spans, color, text=None):
    c = ws.cell(row, 1, label)
    c.font = Font(name="Calibri", size=10, bold=True, color=INK)
    v = ws.cell(row, 2, valor)
    v.font = Font(name="Calibri", size=10, color=INK)
    v.alignment = Alignment(horizontal="center")
    for i in range(NSET):
        ws.cell(row, C0 + i).border = BOX
    for a, b in spans:
        for i in range(a, b + 1):
            ws.cell(row, C0 + i).fill = PatternFill("solid", fgColor=color)
        if text:
            t = ws.cell(row, C0 + a, text)
            t.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")


r = R0 + 2
banda(r, "Fenologia · brotació", "Març, s4", [(BROTACIO, BROTACIO)], FILLS["fen"]); r += 1
banda(r, "Fenologia · verolat", "Juliol, s3", [(VEROLAT, VEROLAT)], FILLS["fen"]); r += 1
banda(r, "Tasca · esporga (poda en verd)", "Juny, s1", [(PODA, PODA)], FILLS["tasca"]); r += 1
banda(r, "Tasca · desfullat (només ES+DV)", "Juliol, s3", [(DESFULLAT, DESFULLAT)], FILLS["tasca"]); r += 1
banda(r, "Tasca · verema", "Ago. s4 – set. s3", [VEREMA], FILLS["tasca"]); r += 1
for lab, i in VOLS:
    banda(r, "Vol de dron · " + lab.replace("\n", " "), lab.split("\n")[1],
          [(i, i)], FILLS["vol"]); r += 1
banda(r, "Kc · coeficient de cultiu", KC, [KC_SPAN], FILLS["kc"], str(KC).replace(".", ",")); r += 1
for lab, kr_maig, kr_resta, _ in TRACT:
    banda(r, "Kr · " + lab, kr_resta, [(8, 11), (12, 27)], FILLS["kr"])
    ws.cell(r, C0 + 8, str(kr_maig).replace(".", ",")).font = Font(size=9, bold=True, color="FFFFFF")
    ws.cell(r, C0 + 12, str(kr_resta).replace(".", ",")).font = Font(size=9, bold=True, color="FFFFFF")
    r += 1
for (a, b), txt in REG:
    banda(r, "Reg aplicat", txt, [(a, b)],
          FILLS["reg1"] if txt == "100 %" else FILLS["reg2"], txt)
    r += 1

r += 1
ws.cell(r, 1, "Kc = coeficient de cultiu · Kr = coeficient de reducció per capçada · "
              "el 25 % correspon al reg deficitari controlat des del verolat.").font = \
    Font(name="Calibri", size=9, italic=True, color="5B6675")

# -- full 2: Kc, Kr i reg ----------------------------------------------------
ws2 = wb.create_sheet("Kc_Kr_Reg")
ws2["A1"] = "Coeficients de reg per tractament · campanya 2026"
ws2["A1"].font = Font(name="Calibri", size=13, bold=True, color=INK)
hdr = ["Tractament", "Kc", "Kr (maig)", "Kr (juny–setembre)",
       "Reg fins al verolat", "Reg des del verolat",
       "fc mitjana mesurada", "fc relativa a EM"]
for j, h in enumerate(hdr, start=1):
    c = ws2.cell(3, j, h)
    c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    c.fill = HEAD
    c.alignment = Alignment(horizontal="center", wrap_text=True)
ws2.column_dimensions["A"].width = 34
for j in range(2, 9):
    ws2.column_dimensions[get_column_letter(j)].width = 15
ref_fc = TRACT[2][3]
for i, (lab, kr_m, kr_r, fc) in enumerate(TRACT):
    row = 4 + i
    for j, v in enumerate([lab, KC, kr_m, kr_r, "100 %", "25 %",
                           round(fc, 3), round(fc / ref_fc, 2)], start=1):
        c = ws2.cell(row, j, v)
        c.font = Font(name="Calibri", size=10, color=INK)
        c.alignment = Alignment(horizontal="left" if j == 1 else "center")
        c.border = BOX

ch = BarChart()
ch.type = "col"
ch.style = 2
ch.title = "Kr aplicat i fracció de coberta relativa (EM = 1)"
ch.y_axis.title = "Valor"
ch.height, ch.width = 8.5, 18
ch.add_data(Reference(ws2, min_col=4, min_row=3, max_row=8), titles_from_data=True)
ch.add_data(Reference(ws2, min_col=8, min_row=3, max_row=8), titles_from_data=True)
ch.set_categories(Reference(ws2, min_col=1, min_row=4, max_row=8))
ws2.add_chart(ch, "A11")

ws2.cell(32, 1, "La fracció de coberta prové de la mitjana dels quatre vols de 2026 "
                "(Taula 14 del TFM). La correlació entre Kr i fc relativa és r² = 0,96.").font = \
    Font(name="Calibri", size=9, italic=True, color="5B6675")

OUTX = os.path.join(BASE, "calendari_maneig_2026.xlsx")
wb.save(OUTX)
print("calendari_maneig_2026.xlsx")

# ---------------------------------------------------------------------------
# Figura per al document
# ---------------------------------------------------------------------------
C_FEN, C_TAS, C_VOL, C_REG, C_RDC = "#1baf7a", "#eb6834", "#2a78d6", "#6250d6", "#bcb6ee"
TINK, MUTED, RULE, FILL = "#1a2230", "#5b6675", "#d4d7dc", "#f4f6f8"
plt.rcParams.update({"font.family": "DejaVu Sans", "text.color": TINK})

fig, ax = plt.subplots(figsize=(9.6, 3.9))
ax.set_xlim(-0.4, NSET + 0.4)
ax.set_ylim(-0.80, 5.4)
ax.axis("off")

ROWS = [("Reg aplicat", 0.55), ("Vols de dron", 1.55), ("Tasques", 2.55), ("Fenologia", 3.55)]
H = 0.62
for lab, y in ROWS:
    ax.add_patch(FancyBboxPatch((0, y), NSET, H, boxstyle="square,pad=0",
                                fc=FILL, ec=RULE, lw=0.7, zorder=1))
    ax.text(-0.6, y + H / 2, lab, ha="right", va="center", fontsize=9.2,
            weight="bold", color=TINK)

# capçalera de mesos
for m, mes in enumerate(MESOS):
    x = m * 4
    ax.add_patch(FancyBboxPatch((x, 4.45), 4, 0.40, boxstyle="square,pad=0",
                                fc="#eceff3" if m % 2 == 0 else "#f6f8fa",
                                ec=RULE, lw=0.7, zorder=1))
    ax.text(x + 2, 4.65, mes, ha="center", va="center", fontsize=8.8,
            weight="bold", color=TINK)
    for s in range(4):
        ax.text(x + s + 0.5, 4.26, str(s + 1), ha="center", va="center",
                fontsize=6.8, color=MUTED)
for m in range(1, 8):
    ax.plot([m * 4, m * 4], [0.4, 4.45], color=RULE, lw=0.7, zorder=0)


def marca(y, a, b, color, text, fs=8.0, tc="white"):
    ax.add_patch(FancyBboxPatch((a, y + 0.06), b - a + 1, H - 0.12,
                                boxstyle="round,pad=0,rounding_size=0.12",
                                fc=color, ec="white", lw=1.0, zorder=3))
    ax.text((a + b + 1) / 2, y + H / 2, text, ha="center", va="center",
            fontsize=fs, weight="bold", color=tc, zorder=4)


marca(3.55, BROTACIO, BROTACIO, C_FEN, "")
ax.text(BROTACIO + 1.3, 3.55 + H / 2, "Brotació", ha="left", va="center",
        fontsize=8.4, color=TINK)
marca(3.55, VEROLAT, VEROLAT, C_FEN, "")
ax.text(VEROLAT + 1.3, 3.55 + H / 2, "Verolat", ha="left", va="center",
        fontsize=8.4, color=TINK)

marca(2.55, PODA, PODA, C_TAS, "")
ax.text(PODA + 1.3, 2.55 + H / 2, "Desbrot i tall en altura", ha="left",
        va="center", fontsize=8.4, color=TINK)
marca(2.55, DESFULLAT, DESFULLAT, C_TAS, "")
ax.text(DESFULLAT + 1.3, 2.55 + H / 2, "Desfullat", ha="left", va="center",
        fontsize=8.4, color=TINK)
marca(2.55, VEREMA[0], VEREMA[1], C_TAS, "Verema")

for n, (lab, i) in enumerate(VOLS, start=1):
    marca(1.55, i, i, C_VOL, str(n), fs=8.6)

for (a, b), txt in REG:
    wide = (b - a) >= 3
    marca(0.55, a, b, C_REG if txt == "100 %" else C_RDC,
          txt if wide else "", tc="white" if txt == "100 %" else TINK)
    if not wide:
        ax.text((a + b + 1) / 2, 0.55 + H + 0.1, txt, ha="center", va="bottom",
                fontsize=7.6, color=TINK, weight="bold")

ax.text(0, 0.10, "Vols de dron:  1 · 03/06    2 · 30/06    3 · 22/07    4 · 12/08",
        ha="left", va="center", fontsize=8.0, color=TINK)
ax.text(0, -0.22, "Desbrot i tall en altura: 02/06 · desfullat del tractament ES+DV: "
                  "20/07, coincidint amb el verolat.",
        ha="left", va="center", fontsize=7.8, color=MUTED)
ax.text(0, -0.52, "Kc = 0,45 constant de maig a setembre. El 25 % correspon al reg "
                  "deficitari controlat aplicat des del verolat.",
        ha="left", va="center", fontsize=7.8, color=MUTED, style="italic")

fig.savefig(os.path.join(FIGDIR, "fig17_calendari.png"), dpi=220,
            bbox_inches="tight", facecolor="white")
plt.close(fig)
print("fig17_calendari.png",
      Image.open(os.path.join(FIGDIR, "fig17_calendari.png")).size)
