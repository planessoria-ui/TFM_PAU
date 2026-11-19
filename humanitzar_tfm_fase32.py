# -*- coding: utf-8 -*-
"""
Fase 32: d'on surt cada columna del fitxer d'entrada.

L'hora és la del vol en UTC, no l'hora solar. La geometria i les temperatures
vénen del processament de les imatges, el LAI de la relació empírica amb la
fIPAR i la meteorologia de l'estació instal·lada a la mateixa parcel·la, no de
la de la xarxa agrometeorològica.
"""
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PATH = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v8.docx"
doc = docx.Document(PATH)
FONT = "Arial"
_PARENT = doc.paragraphs[0]._parent


def _font(run, size=12, bold=False, color=None):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts'); rpr.append(rf)
    for a in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rf.set(qn(a), FONT)


def P(sub):
    for p in doc.paragraphs:
        if sub in p.text:
            return p
    raise LookupError(sub[:70])


def set_text(p, t, size=12, color=None):
    for r in p.runs[1:]:
        r._element.getparent().remove(r._element)
    if p.runs:
        p.runs[0].text = t
        _font(p.runs[0], size, color=color)
    else:
        _font(p.add_run(t), size, color=color)


def cell_set(cl, t, size, bold=False):
    pp = cl.paragraphs[0]
    for r in pp.runs:
        r._element.getparent().remove(r._element)
    _font(pp.add_run(t), size, bold=bold)


def after(el):
    new = OxmlElement('w:p')
    el.addnext(new)
    return docx.text.paragraph.Paragraph(new, _PARENT)


def body(el, text):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    p.paragraph_format.space_after = Pt(6)
    _font(p.add_run(text), 12)
    return p._element


# ===========================================================================
# 1. Taula 17: l'hora i la procedència de cada família de variables
# ===========================================================================
t17 = [t for t in doc.tables if t.cell(0, 0).text.strip() == "Columna"][0]
CANVIS = {
    "hour": ("Hora del vol, en temps universal coordinat (UTC)", None),
    "Identificació i moment": ("", None),
    "Estructura de la capçada, per cep":
        (None, "Estructura de la capçada, del processament de les imatges"),
    "Temperatures": (None, "Temperatures, de la capa tèrmica calibrada"),
    "Meteorologia, comuna a tota l'escena":
        (None, "Meteorologia, de l'estació de la parcel·la"),
}
for row in t17.rows:
    clau = row.cells[0].text.strip()
    if clau not in CANVIS:
        continue
    sig, nou_grup = CANVIS[clau]
    if sig:
        cell_set(row.cells[1], sig, 9)
    if nou_grup:
        cell_set(row.cells[0], nou_grup, 9, bold=True)

# ===========================================================================
# 2. Taula 18: l'hora, en UTC
# ===========================================================================
t18 = [t for t in doc.tables if t.cell(0, 1).text.strip() == "03/06/2026"][0]
HORES = {"hour (h)": ("hour (UTC)", ["12:30", "12:05", "11:51", "12:00"]),
         "Dia de l'any i hora solar": ("Dia de l'any i hora del vol", None)}
for row in t18.rows:
    clau = row.cells[0].text.strip()
    if clau not in HORES:
        continue
    nou, vals = HORES[clau]
    cell_set(row.cells[0], nou, 8.5, bold=(vals is None))
    if vals:
        for cl, v in zip(row.cells[1:], vals):
            cell_set(cl, v, 8.5)

# ===========================================================================
# 3. D'on surt cada columna
# ===========================================================================
el = P("El model no es crida sobre la imatge sencera sinó sobre una taula")._element
body(el,
     "Les columnes vénen de tres llocs. La geometria de la capçada i les tres "
     "temperatures surten del processament de les imatges, tal com es descriu als "
     "apartats 4.6.2 i 4.6.4. El LAI, de la relació empírica amb la fIPAR diària de "
     "l'apartat 4.7.3. I les variables meteorològiques, de l'estació instal·lada a "
     "la mateixa parcel·la d'El polvorín. S'hi recorre en comptes de fer-ho a "
     "l'estació de Raimat de la Xarxa Agrometeorològica de Catalunya, que és la més "
     "propera, perquè les seves lectures descriuen millor les condicions en què es "
     "va volar. L'hora que consta a cada fila és la del vol en temps universal "
     "coordinat, de manera que els quatre vols es van fer entre les 11:51 i les "
     "12:30 UTC.")

# ===========================================================================
# 4. La llista de variables del 4.7 deia que la meteo venia de la XAC
# ===========================================================================
set_text(P("obtingudes de l'estació agrometeorològica més propera de la Xarxa"),
    "Variables meteorològiques: temperatura de l'aire, humitat relativa, dèficit de "
    "pressió de vapor (VPD), velocitat del vent, radiació solar i evapotranspiració "
    "de referència, registrades per l'estació instal·lada a la mateixa parcel·la "
    "d'El polvorín, que són les entrades meteorològiques dels models de balanç "
    "energètic.")

# ===========================================================================
# 5. Les notes que preguntaven per l'estació, ja resoltes en aquest punt
# ===========================================================================
set_text(P("identificador i distància de l'estació XAC utilitzada"),
    "[A CONCRETAR] Nombre de ceps mostrejats per tractament i bloc en cada mesura; "
    "hores exactes de les corbes diürnes; i model de cambra de pressió.",
    color=RGBColor(0xC0, 0x00, 0x00))

set_text(P("Indiqueu també si les dades meteorològiques provenen d'una estació pròpia"),
    "[A COMPLETAR] Ompliu les columnes de model i fabricant, nombre d'unitats, "
    "ubicació per tractament i bloc, i freqüència i sistema de registre.",
    color=RGBColor(0xC0, 0x00, 0x00))

# ===========================================================================
# 6. Taula 2: on és l'estació
# ===========================================================================
t2 = [t for t in doc.tables
      if t.cell(0, 0).text.strip().startswith("Sensor o equip")][0]
for row in t2.rows:
    if row.cells[0].text.strip() == "Estació meteorològica":
        cell_set(row.cells[3], "1", 10)
        cell_set(row.cells[4], "A la mateixa parcel·la (El polvorín)", 10)

doc.save(PATH)
print(f"OK · paràgrafs: {len(doc.paragraphs)}")
