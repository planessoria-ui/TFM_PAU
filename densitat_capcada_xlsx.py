# -*- coding: utf-8 -*-
"""
Full de càlcul amb la relació entre la fracció de coberta i l'índex d'àrea
foliar dels quatre vols de 2026, i els gràfics com a objectes natius d'Excel.

  Gràfic 1  densitat de la paret vegetal (LAI / fc) per tractament i vol
  Gràfic 2  LAI contra fc, un núvol per tractament
  Gràfic 3  petjada, fulla, consum i dosi de reg, relatius al control

Sortida: densitat_capcada_2026.xlsx
"""
import os
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, ScatterChart, BarChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.series import SeriesLabel
from openpyxl.drawing.line import LineProperties

BASE = os.path.dirname(os.path.abspath(__file__))
X = "/tmp/claude-0/-home-user-TFM-PAU/446d5a3f-407d-5d51-8cea-50e8348ad898/xls/"
D = "/tmp/claude-0/-home-user-TFM-PAU/446d5a3f-407d-5d51-8cea-50e8348ad898/csvdata/"
OUT = os.path.join(BASE, "densitat_capcada_2026.xlsx")

FS = {"03/06/2026": "790d17a6-inputs_03062026.csv",
      "30/06/2026": "c64943ce-inputs_30062026.csv",
      "23/07/2026": "4a8253d5-inputs_23072026.csv",
      "12/08/2026": "244b6521-inputs_12082026.csv"}
DATES = list(FS)
ORDRE = ["ES+DV", "ES", "EM", "ELL", "C"]
NOM = {"ES+DV": "Esporga severa + desfullat", "ES": "Esporga severa",
       "EM": "Esporga mitjana", "ELL": "Esporga lleugera", "C": "Sense esporga"}
# mateix color per tractament que a la figura dels paràmetres estructurals
COLOR = {"C": "2A78D6", "ELL": "EB6834", "EM": "1BAF7A",
         "ES": "EDA100", "ES+DV": "E87BA4"}
MARCA = {"C": "circle", "ELL": "square", "EM": "triangle",
         "ES": "diamond", "ES+DV": "x"}
KR = {"ES+DV": 0.70, "ES": 0.70, "EM": 1.00, "ELL": 1.10, "C": 1.40}

INK, GRIS = "1A2230", "5B6675"
HEAD = PatternFill("solid", fgColor="1F5C8B")
SUB = PatternFill("solid", fgColor="E8EDF2")
THIN = Side(style="thin", color="C8CDD4")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# ---------------------------------------------------------------------------
# Dades
# ---------------------------------------------------------------------------
m = pd.read_csv(X + "merged.csv")
inp = pd.concat([pd.read_csv(D + f, index_col=0).assign(data=d[:5], id=range(1, 119))
                 for d, f in FS.items()])
m = m.merge(inp[["data", "id", "fc"]], on=["data", "id"])
CURT = {d[:5]: d for d in DATES}
g = (m.groupby(["t", "data"])
       .agg(fc=("fc", "mean"), LAI=("lai_bo", "mean"), ETa=("ETa_DAILY_2T", "mean"))
       .reset_index())
g["data"] = g.data.map(CURT)
P = {v: g.pivot_table(index="t", columns="data", values=v).loc[ORDRE, DATES]
     for v in ("fc", "LAI", "ETa")}
camp = m.groupby("t").agg(fc=("fc", "mean"), LAI=("lai_bo", "mean"),
                          ETa=("ETa_DAILY_2T", "mean")).loc[ORDRE]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Dades"
ws["A1"] = "Fracció de coberta, àrea foliar i densitat de la paret vegetal · campanya 2026"
ws["A1"].font = Font(name="Calibri", size=14, bold=True, color=INK)
ws["A2"] = ("Mitjana dels ceps de cada tractament. La densitat és el quocient entre "
            "l'índex d'àrea foliar i la fracció de coberta, i expressa els metres "
            "quadrats de fulla per metre quadrat de sòl cobert.")
ws["A2"].font = Font(name="Calibri", size=10, italic=True, color=GRIS)
ws.column_dimensions["A"].width = 32
for j in range(2, 10):
    ws.column_dimensions[get_column_letter(j)].width = 13
ws.freeze_panes = "B5"


def bloc(row, titol, dades, numfmt, formula=None):
    ws.cell(row, 1, titol).font = Font(name="Calibri", size=11, bold=True, color=INK)
    row += 1
    h = ws.cell(row, 1, "Tractament")
    h.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    h.fill = HEAD
    for j, d in enumerate(DATES):
        c = ws.cell(row, 2 + j, d)
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = HEAD
        c.alignment = Alignment(horizontal="center")
    cap = row
    row += 1
    primera = row
    for tr in ORDRE:
        ws.cell(row, 1, NOM[tr]).font = Font(name="Calibri", size=10, color=INK)
        ws.cell(row, 1).border = BOX
        for j, d in enumerate(DATES):
            c = ws.cell(row, 2 + j,
                        formula(row, j) if formula else round(dades.loc[tr, d], 4))
            c.number_format = numfmt
            c.font = Font(name="Calibri", size=10, color=INK)
            c.alignment = Alignment(horizontal="center")
            c.border = BOX
        row += 1
    return cap, primera, row + 1


cap_fc, fila_fc, row = bloc(4, "Fracció de coberta, fc (–)", P["fc"], "0.000")
cap_lai, fila_lai, row = bloc(row, "Índex d'àrea foliar, LAI (m² m⁻²)", P["LAI"], "0.000")
prev_den = row + 2          # primera fila de dades del bloc de densitat
cap_den, fila_den, row = bloc(
    row, "Densitat de la paret vegetal, LAI / fc (m² m⁻²)", None, "0.00",
    formula=lambda r, j: f"={get_column_letter(2+j)}{fila_lai + (r - prev_den)}"
                         f"/{get_column_letter(2+j)}{fila_fc + (r - prev_den)}")
assert fila_den == prev_den, (fila_den, prev_den)

# bloc relatiu al control
ws.cell(row, 1, "Valors relatius al tractament sense esporga (mitjana de campanya)").font = \
    Font(name="Calibri", size=11, bold=True, color=INK)
row += 1
cap_rel = row
h = ws.cell(row, 1, "Tractament")
h.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
h.fill = HEAD
for j, t in enumerate(["Dosi de reg (Kr)", "Fracció de coberta", "Àrea foliar",
                       "Consum d'aigua"]):
    c = ws.cell(row, 2 + j, t)
    c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    c.fill = HEAD
    c.alignment = Alignment(horizontal="center", wrap_text=True)
row += 1
fila_rel = row
for tr in ORDRE:
    ws.cell(row, 1, NOM[tr]).font = Font(name="Calibri", size=10, color=INK)
    ws.cell(row, 1).border = BOX
    vals = [KR[tr] / KR["C"], camp.loc[tr, "fc"] / camp.loc["C", "fc"],
            camp.loc[tr, "LAI"] / camp.loc["C", "LAI"],
            camp.loc[tr, "ETa"] / camp.loc["C", "ETa"]]
    for j, v in enumerate(vals):
        c = ws.cell(row, 2 + j, round(v, 3))
        c.number_format = "0.00"
        c.font = Font(name="Calibri", size=10, color=INK)
        c.alignment = Alignment(horizontal="center")
        c.border = BOX
    row += 1
ws.cell(row + 1, 1,
        "La dosi de reg es va fixar amb un coeficient proporcional a la fracció de "
        "coberta. El consum d'aigua se situa per sobre, més a prop de l'àrea foliar."
        ).font = Font(name="Calibri", size=9, italic=True, color=GRIS)

# ---------------------------------------------------------------------------
# Gràfics
# ---------------------------------------------------------------------------
wg = wb.create_sheet("Gràfics")
wg["A1"] = "Relació entre la fracció de coberta i l'àrea foliar"
wg["A1"].font = Font(name="Calibri", size=14, bold=True, color=INK)
wg["A2"] = "Els gràfics estan vinculats al full Dades i es refan sols si en canvien els valors."
wg["A2"].font = Font(name="Calibri", size=10, italic=True, color=GRIS)

GRID = GraphicalProperties(ln=LineProperties(solidFill="E3E6EA", w=9525))
SENSE_MARC = GraphicalProperties(ln=LineProperties(noFill=True))

# --- 1. densitat de la paret vegetal
ch = LineChart()
ch.title = "Densitat de la paret vegetal per tractament"
ch.style = None
ch.y_axis.title = "LAI / fc (m² de fulla per m² de sòl cobert)"
ch.x_axis.title = "Data del vol"
ch.height, ch.width = 10.5, 17.5
ch.y_axis.majorGridlines.spPr = GRID
ch.x_axis.majorGridlines = None
ch.y_axis.scaling.min, ch.y_axis.scaling.max = 5.0, 11.0
ch.y_axis.majorUnit = 1.0
ch.y_axis.numFmt = "0"
ch.graphical_properties = SENSE_MARC
for i, tr in enumerate(ORDRE):
    r = fila_den + i
    ch.add_data(Reference(ws, min_col=2, max_col=5, min_row=r, max_row=r),
                from_rows=True, titles_from_data=False)
    s = ch.series[-1]
    s.tx = SeriesLabel(v=NOM[tr])
    s.graphicalProperties = GraphicalProperties(
        ln=LineProperties(solidFill=COLOR[tr], w=28575))
    s.marker = Marker(symbol=MARCA[tr], size=7, spPr=GraphicalProperties(
        solidFill=COLOR[tr], ln=LineProperties(solidFill="FFFFFF", w=9525)))
    s.smooth = False
ch.set_categories(Reference(ws, min_col=2, max_col=5, min_row=cap_den))
ch.legend.position = "b"
ch.legend.overlay = False
wg.add_chart(ch, "A4")

# --- 2. LAI contra fc
sc = ScatterChart()
sc.title = "Índex d'àrea foliar contra fracció de coberta"
sc.style = None
sc.x_axis.title = "fc (–)"
sc.y_axis.title = "LAI (m² m⁻²)"
sc.height, sc.width = 10.5, 17.5
sc.y_axis.majorGridlines.spPr = GRID
sc.x_axis.majorGridlines.spPr = GRID
sc.x_axis.scaling.min, sc.x_axis.scaling.max = 0.15, 0.45
sc.x_axis.majorUnit = 0.05
sc.y_axis.scaling.min, sc.y_axis.scaling.max = 1.6, 2.6
sc.y_axis.majorUnit = 0.2
sc.x_axis.numFmt, sc.y_axis.numFmt = "0.00", "0.0"
sc.graphical_properties = SENSE_MARC
for i, tr in enumerate(ORDRE):
    xr = Reference(ws, min_col=2, max_col=5, min_row=fila_fc + i)
    yr = Reference(ws, min_col=2, max_col=5, min_row=fila_lai + i)
    s = Series(yr, xr, title=NOM[tr])
    s.marker = Marker(symbol=MARCA[tr], size=8, spPr=GraphicalProperties(
        solidFill=COLOR[tr], ln=LineProperties(solidFill="FFFFFF", w=9525)))
    s.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    sc.series.append(s)
sc.legend.position = "b"
sc.legend.overlay = False
wg.add_chart(sc, "L4")

# --- 3. petjada, fulla i consum, relatius al control
bc = BarChart()
bc.type = "col"
bc.title = "Dosi de reg, estructura i consum, relatius al tractament sense esporga"
bc.style = None
bc.y_axis.title = "Valor relatiu al tractament sense esporga"
bc.x_axis.title = "Tractament"
bc.height, bc.width = 10.5, 20.0
bc.y_axis.majorGridlines.spPr = GRID
bc.x_axis.majorGridlines = None
bc.y_axis.scaling.min, bc.y_axis.scaling.max = 0.0, 1.1
bc.y_axis.majorUnit = 0.1
bc.y_axis.numFmt = "0.0"
bc.gapWidth = 60
bc.overlap = -10
bc.graphical_properties = SENSE_MARC
bc.add_data(Reference(ws, min_col=2, max_col=5, min_row=cap_rel,
                      max_row=fila_rel + len(ORDRE) - 1), titles_from_data=True)
bc.set_categories(Reference(ws, min_col=1, min_row=fila_rel,
                            max_row=fila_rel + len(ORDRE) - 1))
for s, col in zip(bc.series, ["8C96A3", "2A78D6", "1BAF7A", "EB6834"]):
    s.graphicalProperties = GraphicalProperties(
        solidFill=col, ln=LineProperties(noFill=True))
bc.legend.position = "b"
bc.legend.overlay = False
wg.add_chart(bc, "A26")

wb.save(OUT)
print(f"{os.path.basename(OUT)} · {len(wg._charts)} gràfics natius")
print()
print("Densitat LAI/fc:")
print((P["LAI"] / P["fc"]).round(2).to_string())
