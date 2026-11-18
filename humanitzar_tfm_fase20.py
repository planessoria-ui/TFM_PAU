# -*- coding: utf-8 -*-
"""
Fase 20: substitueix la imatge del flux de treball metodològic (Figura 4) per
una taula amb format natiu de Word, editable des del mateix document.
"""
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PATH = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v8.docx"
doc = docx.Document(PATH)
FONT = "Arial"
ACCENT = RGBColor(0x1F, 0x5C, 0x8B)
INK = RGBColor(0x1A, 0x22, 0x30)
MUTED = RGBColor(0x5B, 0x66, 0x75)
RULE_HEX = "C8CDD4"
ACCENT_HEX = "1F5C8B"
FILL_HEX = "F2F5F8"

PASSOS = [
    ("01", "Adquisició de dades",
     "Vols de dron amb MicaSense Altum-PT (cinc bandes espectrals i banda tèrmica "
     "LWIR) i mesures de camp simultànies: ceptòmetre, potencial hídric de tija i "
     "meteorologia."),
    ("02", "Processament fotogramètric",
     "Ortomosaics, model digital de superfície i del terreny, i model digital de "
     "vegetació amb l'alçada del dosser."),
    ("03", "Calibratge radiomètric",
     "Correcció de les cinc bandes espectrals amb els panells de referència i de la "
     "banda tèrmica amb el termòmetre d'infraroig."),
    ("04", "Extracció de variables per cep",
     "Màscares d'alçada i de SAVI, índexs de vegetació i paràmetres estructurals "
     "(fracció de coberta, alçada i volum de copa) sobre la capa de polígons."),
    ("05", "Estimació del LAI",
     "Ceptòmetre i formulació de Campbell i Norman, validació amb ceps defoliats "
     "(R², RMSE) i extensió a tots els ceps mitjançant Random Forest."),
    ("06", "Temperatures de component",
     "Temperatura de dosser (Tc) i de sòl (Ts) a partir de les màscares, amb retall "
     "de percentils per descartar els píxels barrejats."),
    ("07", "Model TSEB-2T",
     "Formulació de Campbell i Norman per a cultius en filera; evapotranspiració "
     "real i partició entre transpiració i evaporació."),
    ("08", "Validació",
     "Flux de saba, aigua de reg aplicada i potencial hídric de tija."),
]


def _font(run, size=12, bold=False, color=None, name=FONT):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts'); rpr.append(rf)
    for a in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rf.set(qn(a), name)


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:shd')):
        tcPr.remove(old)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill)
    tcPr.append(shd)


def borders(cell, spec):
    """spec: {'left': (sz_eighths, color), ...}; absents queden sense vora."""
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:tcBorders')):
        tcPr.remove(old)
    tcB = OxmlElement('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement('w:' + side)
        if side in spec:
            sz, color = spec[side]
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), str(sz))
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), color)
        else:
            el.set(qn('w:val'), 'nil')
        tcB.append(el)
    tcPr.append(tcB)


def margins(cell, top=80, bottom=80, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement('w:tcMar')
    for side, val in (('top', top), ('left', left), ('bottom', bottom), ('right', right)):
        el = OxmlElement('w:' + side)
        el.set(qn('w:w'), str(val))
        el.set(qn('w:type'), 'dxa')
        mar.append(el)
    tcPr.append(mar)


def no_split(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement('w:cantSplit')
    trPr.append(el)


# ---------------------------------------------------------------------------
# Localitza la imatge de la Figura 4 i la substitueix per la taula
# ---------------------------------------------------------------------------
cap = [p for p in doc.paragraphs if "flux de treball metodològic" in p.text
       and p.text.strip().startswith(("Figura", "Figura 4"))][0]
img_p = cap._element.getprevious()
holder = docx.text.paragraph.Paragraph(img_p, doc.paragraphs[0]._parent)
for r in holder.runs:
    r._element.getparent().remove(r._element)
holder.paragraph_format.space_after = Pt(0)

nrows = len(PASSOS) * 2 - 1
tbl = doc.add_table(rows=nrows, cols=2)
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl.autofit = False
W0, W1 = Cm(1.5), Cm(14.3)

# disseny fix: sense això el Word reparteix les columnes pel seu compte
tblPr = tbl._tbl.tblPr
layout = OxmlElement('w:tblLayout')
layout.set(qn('w:type'), 'fixed')
tblPr.append(layout)
grid = tbl._tbl.find(qn('w:tblGrid'))
for gc, w in zip(grid.findall(qn('w:gridCol')), (W0, W1)):
    gc.set(qn('w:w'), str(int(w.twips)))

for i, (num, titol, cos) in enumerate(PASSOS):
    r = i * 2
    row = tbl.rows[r]
    no_split(row)

    c0 = tbl.cell(r, 0)
    c0.width = W0
    c0.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    shade(c0, FILL_HEX)
    borders(c0, {'left': (24, ACCENT_HEX), 'top': (4, RULE_HEX),
                 'bottom': (4, RULE_HEX)})
    margins(c0, left=100, right=60)
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_after = Pt(0)
    p0.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    _font(p0.add_run(num), 11, bold=True, color=ACCENT)

    c1 = tbl.cell(r, 1)
    c1.width = W1
    c1.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    shade(c1, FILL_HEX)
    borders(c1, {'top': (4, RULE_HEX), 'bottom': (4, RULE_HEX),
                 'right': (4, RULE_HEX)})
    margins(c1, left=60, right=160)
    pt = c1.paragraphs[0]
    pt.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pt.paragraph_format.space_after = Pt(2)
    pt.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    _font(pt.add_run(titol), 11, bold=True, color=INK)
    pb = c1.add_paragraph()
    pb.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pb.paragraph_format.space_after = Pt(0)
    pb.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    _font(pb.add_run(cos), 9.5, color=MUTED)

    if r + 1 < nrows:
        arow = tbl.rows[r + 1]
        no_split(arow)
        merged = tbl.cell(r + 1, 0).merge(tbl.cell(r + 1, 1))
        borders(merged, {})
        margins(merged, top=20, bottom=20)
        pa = merged.paragraphs[0]
        pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pa.paragraph_format.space_before = Pt(0)
        pa.paragraph_format.space_after = Pt(0)
        pa.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        _font(pa.add_run("▼"), 9, color=RGBColor(0x9A, 0xA3, 0xAE))

img_p.addnext(tbl._element)

# peu actualitzat
runs = cap.runs
endi = None
for i, r in enumerate(runs):
    f = r._element.find(qn('w:fldChar'))
    if f is not None and f.get(qn('w:fldCharType')) == 'end':
        endi = i
if endi is not None:
    for r in runs[endi + 1:]:
        r._element.getparent().remove(r._element)
    _font(cap.add_run(
        ". Esquema general del flux de treball metodològic del TFM, des de "
        "l'adquisició de dades fins a l'estimació de l'evapotranspiració real amb "
        "el model TSEB-2T i la seva validació. Font: elaboració pròpia."), 10)
else:
    for r in runs[1:]:
        r._element.getparent().remove(r._element)
    runs[0].text = "Figura 4."
    _font(runs[0], 10, bold=True)
    _font(cap.add_run(
        " Esquema general del flux de treball metodològic del TFM, des de "
        "l'adquisició de dades fins a l'estimació de l'evapotranspiració real amb "
        "el model TSEB-2T i la seva validació. Font: elaboració pròpia."), 10)

doc.save(PATH)
print(f"OK · {len(PASSOS)} passos · taula de {nrows} files · "
      f"taules al document: {len(doc.tables)}")
