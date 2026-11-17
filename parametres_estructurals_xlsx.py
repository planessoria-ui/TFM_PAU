# -*- coding: utf-8 -*-
"""
Full de càlcul amb els paràmetres estructurals del dosser dels quatre vols de
2026 i els gràfics corresponents com a objectes natius d'Excel, editables.

Sortida: parametres_estructurals_2026.xlsx
"""
import csv, json, os, statistics as st

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.data_source import NumDataSource, NumRef
from openpyxl.chart.error_bar import ErrorBars
from openpyxl.drawing.line import LineProperties

BASE = os.path.dirname(os.path.abspath(__file__))
U = "/root/.claude/uploads/446d5a3f-407d-5d51-8cea-50e8348ad898/"
OUT = os.path.join(BASE, "parametres_estructurals_2026.xlsx")

F = [("03/06/2026", "a290cdc2-biophysical_traits_YIELD4CAST2026_covercrop_P25_v2.csv"),
     ("30/06/2026", "57255596-biophysical_traits_YIELD4CAST2026_covercrop_P25.csv"),
     ("22/07/2026", "7e8f3dc6-biophysical_traits_YIELD4CAST2026_covercrop_P25_22072026.csv"),
     ("12/08/2026", "9b33ef42-biophysical_traits_YIELD4CAST2026_covercrop_P25_12082026.csv")]
SHORT = {"Esporga severa i desfullat a verolat": "ES+DV", "Esporga severa + desfullat": "ES+DV",
         "Esporga severa": "ES", "Esporga mitjana": "EM",
         "Esporga lleugera": "EL", "Sense esporga": "SE"}
ORDER = ["SE", "EL", "EM", "ES", "ES+DV"]
NOM = {"SE": "Sense esporga", "EL": "Esporga lleugera", "EM": "Esporga mitjana",
       "ES": "Esporga severa", "ES+DV": "Esporga severa + desfullat"}
# paleta categòrica validada, slots 1-5
COLOR = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4"]
MARKER = ["circle", "square", "triangle", "diamond", "x"]

VARS = [("fc", "Fracció de coberta", "fc (–)", "0.000", (0.10, 0.45, 0.05)),
        ("hc", "Alçada del dosser", "hc (m)", "0.00", (1.6, 2.5, 0.1)),
        ("fVolc", "Volum de copa per unitat de superfície de sòl", "fVolc (m)",
         "0.000", (0.15, 0.70, 0.05)),
        ("wc_hc", "Relació entre amplada i alçada de la capçada", "wc/hc (–)",
         "0.000", (0.40, 0.70, 0.05))]

INK = "1A2230"
HEAD = PatternFill("solid", fgColor="1F5C8B")
SUB = PatternFill("solid", fgColor="E8EDF2")
THIN = Side(style="thin", color="C8CDD4")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def f2(s):
    s = (s or "").strip().replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


DATA = []
for date, fn in F:
    raw = open(U + fn, encoding="utf-8-sig").read()
    d = ";" if raw.count(";") > raw.count(",") else ","
    for r in csv.DictReader(raw.splitlines(), delimiter=d):
        if not (r.get("tree") or "").strip():
            continue
        rec = {"date": date, "trt": SHORT[r["Treatment"].strip()]}
        for k, *_ in VARS:
            rec[k] = f2(r.get(k))
        DATA.append(rec)


def stats(var, date, trt):
    v = [x[var] for x in DATA if x["date"] == date and x["trt"] == trt and x[var] is not None]
    return (st.mean(v), st.pstdev(v), len(v)) if v else (None, None, 0)


wb = openpyxl.Workbook()

# ---------------------------------------------------------------------------
# Full 1: dades
# ---------------------------------------------------------------------------
ws = wb.active
ws.title = "Dades"
ws["A1"] = "Paràmetres estructurals del dosser · campanya 2026"
ws["A1"].font = Font(name="Calibri", size=14, bold=True, color=INK)
ws["A2"] = ("Mitjana i desviació estàndard dels ceps de cada tractament, "
            "obtingudes del processament dels quatre vols de dron.")
ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="5B6675")
ws.column_dimensions["A"].width = 32
for j in range(2, 10):
    ws.column_dimensions[get_column_letter(j)].width = 13
ws.freeze_panes = "B5"

BLOCKS = {}
row = 4
for key, titol, _, numfmt, _scale in VARS:
    ws.cell(row, 1, titol).font = Font(name="Calibri", size=11, bold=True, color=INK)
    row += 1
    hdr = ws.cell(row, 1, "Tractament")
    hdr.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    hdr.fill = HEAD
    hdr.alignment = Alignment(horizontal="left")
    for j, (date, _) in enumerate(F):
        c = ws.cell(row, 2 + j, date)
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = HEAD
        c.alignment = Alignment(horizontal="center")
    for j, (date, _) in enumerate(F):
        c = ws.cell(row, 6 + j, "σ " + date[:5])
        c.font = Font(name="Calibri", size=9, bold=True, color=INK)
        c.fill = SUB
        c.alignment = Alignment(horizontal="center")
    head_row = row
    row += 1
    first = row
    for i, trt in enumerate(ORDER):
        ws.cell(row, 1, NOM[trt]).font = Font(name="Calibri", size=10, color=INK)
        ws.cell(row, 1).border = BOX
        for j, (date, _) in enumerate(F):
            m, sd, _n = stats(key, date, trt)
            cm = ws.cell(row, 2 + j, round(m, 4))
            cm.number_format = numfmt
            cm.font = Font(name="Calibri", size=10, color=INK)
            cm.alignment = Alignment(horizontal="center")
            cm.border = BOX
            cs = ws.cell(row, 6 + j, round(sd, 4))
            cs.number_format = numfmt
            cs.font = Font(name="Calibri", size=9, color="5B6675")
            cs.alignment = Alignment(horizontal="center")
            cs.border = BOX
        row += 1
    BLOCKS[key] = (head_row, first, row - 1)
    row += 2

ws.cell(row, 1, "n = 24 ceps per tractament, llevat de l'esporga severa i del "
                "tractament sense esporga, que en tenen 23.").font = \
    Font(name="Calibri", size=9, italic=True, color="5B6675")

# ---------------------------------------------------------------------------
# Full 2: gràfics natius
# ---------------------------------------------------------------------------
wg = wb.create_sheet("Gràfics")
wg["A1"] = "Evolució dels paràmetres estructurals per tractament"
wg["A1"].font = Font(name="Calibri", size=14, bold=True, color=INK)
wg["A2"] = ("Gràfics vinculats al full Dades: si en canvien els valors, "
            "els gràfics s'actualitzen sols.")
wg["A2"].font = Font(name="Calibri", size=10, italic=True, color="5B6675")

ANCHORS = ["A4", "L4", "A26", "L26"]
for (key, titol, ytit, _nf, scale), anchor in zip(VARS, ANCHORS):
    head_row, first, last = BLOCKS[key]
    ch = LineChart()
    ch.title = titol
    ch.style = None
    ch.y_axis.title = ytit
    ch.x_axis.title = "Data del vol"
    ch.height, ch.width = 10.5, 17.5
    ch.y_axis.majorGridlines.spPr = GraphicalProperties(
        ln=LineProperties(solidFill="E3E6EA", w=9525))
    ch.x_axis.majorGridlines = None
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = scale[0], scale[1]
    ch.y_axis.majorUnit = scale[2]
    ch.y_axis.numFmt = _nf
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))

    for i, trt in enumerate(ORDER):
        r = first + i
        data = Reference(ws, min_col=2, max_col=5, min_row=r, max_row=r)
        ch.add_data(data, from_rows=True, titles_from_data=False)
        s = ch.series[-1]
        s.tx = None
        s.graphicalProperties = GraphicalProperties(
            ln=LineProperties(solidFill=COLOR[i], w=28575))
        s.marker = Marker(symbol=MARKER[i], size=7,
                          spPr=GraphicalProperties(solidFill=COLOR[i],
                                                   ln=LineProperties(solidFill="FFFFFF",
                                                                     w=9525)))
        s.smooth = False
        # desviació estàndard com a barres d'error
        err = Reference(ws, min_col=6, max_col=9, min_row=r, max_row=r)
        s.errBars = ErrorBars(errBarType="both", errValType="cust",
                              plus=NumDataSource(numRef=NumRef(f=err.__str__())),
                              minus=NumDataSource(numRef=NumRef(f=err.__str__())),
                              noEndCap=False)

    ch.set_categories(Reference(ws, min_col=2, max_col=5, min_row=head_row))
    ch.legend.position = "b"
    ch.legend.overlay = False
    wg.add_chart(ch, anchor)

# noms de sèrie a la llegenda
for ci, ch in enumerate(wg._charts):
    for i, s in enumerate(ch.series):
        from openpyxl.chart.series import SeriesLabel
        from openpyxl.chart.data_source import StrRef
        s.tx = SeriesLabel(v=NOM[ORDER[i]])

wb.save(OUT)
print("parametres_estructurals_2026.xlsx ·", len(wg._charts), "gràfics natius ·",
      len(DATA), "registres")
