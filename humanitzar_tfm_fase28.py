# -*- coding: utf-8 -*-
"""
Fase 28: el fitxer d'entrada del model d'evapotranspiració.

Documenta la taula amb què es crida el TSEB-2T, a partir del fitxer real del
vol del 3 de juny de 2026 (118 files, una per cep), i corregeix l'explicació de
per què la inversió de la fIPAR subestimava el LAI.
"""
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
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


def set_text(p, t, size=12):
    for r in p.runs[1:]:
        r._element.getparent().remove(r._element)
    if p.runs:
        p.runs[0].text = t
        _font(p.runs[0], size)
    else:
        _font(p.add_run(t), size)


def before(el):
    new = OxmlElement('w:p')
    el.addprevious(new)
    return docx.text.paragraph.Paragraph(new, _PARENT)


def after(el):
    new = OxmlElement('w:p')
    el.addnext(new)
    return docx.text.paragraph.Paragraph(new, _PARENT)


def body(el, text, color=None):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    p.paragraph_format.space_after = Pt(6)
    _font(p.add_run(text), 12, color=color)
    return p._element


def caption(el, head, text):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    _font(p.add_run(head), 10, bold=True)
    _font(p.add_run(" " + text), 10)
    return p._element


def table(el, rows, widths, grups=()):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = doc.styles['Table Grid']
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = t._tbl.tblPr
    lay = OxmlElement('w:tblLayout'); lay.set(qn('w:type'), 'fixed'); tblPr.append(lay)
    grid = t._tbl.find(qn('w:tblGrid'))
    for gc, w in zip(grid.findall(qn('w:gridCol')), widths):
        gc.set(qn('w:w'), str(int(w.twips)))
    for r, data in enumerate(rows):
        for c, val in enumerate(data):
            cl = t.cell(r, c)
            cl.width = widths[c]
            pp = cl.paragraphs[0]
            pp.alignment = (WD_ALIGN_PARAGRAPH.LEFT if c == 1 and r
                            else WD_ALIGN_PARAGRAPH.CENTER)
            pp.paragraph_format.space_before = Pt(2)
            pp.paragraph_format.space_after = Pt(2)
            pp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            _font(pp.add_run(val), 9, bold=(r == 0 or r in grups))
    el.addnext(t._element)
    return t._element


# ===========================================================================
# 1. Per què la inversió subestimava el LAI: els motius reals
# ===========================================================================
set_text(P("La subestimació té una explicació coneguda en vinya"),
    "El ceptòmetre no mesura el LAI: el dedueix de la fracció de radiació que el "
    "dosser intercepta. Aquesta inversió es pot fer de dues maneres, amb el model "
    "que incorpora el mateix LP-80 i amb la formulació de Campbell i Norman, i totes "
    "dues es van provar amb els ceps defoliats, dels quals es coneix el LAI real. El "
    "resultat va ser el mateix en els dos casos: els valors invertits quedaven per "
    "sota dels mesurats. La inversió arrossega factors que no es controlen del tot, "
    "com l'hora a què es pren la lectura i, per tant, l'angle del sol, la inclinació "
    "amb què es manté la barra i l'alçada a què se situa sota el dosser. Qualsevol "
    "d'aquestes desviacions fa que el sensor rebi més radiació de la que li tocaria "
    "i que el LAI que se'n dedueix surti massa baix.")

# ===========================================================================
# 2. El fitxer d'entrada del model
# ===========================================================================
anchor = P("L'evapotranspiració potencial (ETp) s'estima amb el model de Shuttleworth")
el = before(anchor._element)._element

el = body(el,
    "El model no es crida sobre la imatge sencera sinó sobre una taula. El fitxer "
    "d'entrada té una fila per cep i per data, i cada fila porta tot el que el "
    "balanç d'energia necessita per a aquell cep en el moment del vol: la geometria "
    "de la capçada, les tres temperatures i les condicions meteorològiques. La "
    "Taula 17 en recull les columnes, amb el seu significat i les seves unitats, i "
    "hi afegeix el que donen els 118 ceps de l'assaig en el vol del 3 de juny de "
    "2026.")

el = caption(el, "Taula 17.",
             "Columnes del fitxer d'entrada del model TSEB-2T i valors del vol del "
             "3 de juny de 2026 (118 ceps). On la variable canvia de cep a cep "
             "s'indica el rang i, entre parèntesis, la mitjana; on és comuna a tota "
             "l'escena, el valor únic. Font: elaboració pròpia.")

FILES = [
    ["Columna", "Significat", "Unitats", "Vol del 03/06/2026"],
    ["Identificació i moment", "", "", ""],
    ["tree", "Identificador del cep dins l'assaig", "–", "1 a 118"],
    ["date, doy", "Data del vol i dia de l'any", "–", "03/06/2026, 154"],
    ["hour", "Hora solar de l'adquisició", "h", "12,5"],
    ["Estructura de la capçada, per cep", "", "", ""],
    ["fc", "Fracció de coberta", "–", "0,157 – 0,386 (0,233)"],
    ["hc", "Alçada del dosser", "m", "1,72 – 2,65 (2,14)"],
    ["wc_hc", "Relació entre l'amplada i l'alçada de la copa", "–",
     "0,413 – 0,658 (0,526)"],
    ["LAI", "Índex d'àrea foliar, de l'equació de l'apartat 4.7.3", "m² m⁻²",
     "1,61 – 2,33 (1,91)"],
    ["Temperatures", "", "", ""],
    ["T_R", "Temperatura radiomètrica del conjunt", "K", "311,8 – 322,2 (318,2)"],
    ["T_C", "Temperatura del dosser", "K", "298,20"],
    ["T_S", "Temperatura del sòl", "K", "317,4 – 326,4 (322,2)"],
    ["Meteorologia, comuna a tota l'escena", "", "", ""],
    ["Ta", "Temperatura de l'aire", "K", "301,02"],
    ["u10", "Velocitat del vent a 10 m", "m s⁻¹", "3,1"],
    ["RH", "Humitat relativa, en tant per u", "–", "0,395"],
    ["Sp", "Radiació solar incident en el moment del vol", "W m⁻²", "965"],
    ["Eto", "Evapotranspiració de referència horària", "mm h⁻¹", "0,65"],
    ["Sp_daily", "Radiació solar acumulada del dia", "MJ m⁻² d⁻¹", "31"],
    ["Eto_daily", "Evapotranspiració de referència diària", "mm d⁻¹", "5,59"],
]
W = [Cm(3.0), Cm(7.0), Cm(2.0), Cm(4.0)]
el = table(el, FILES, W, grups=(1, 5, 10, 14))

el = body(el,
    "De totes aquestes columnes, les que donen identitat a cada cep són la fracció "
    "de coberta, l'alçada, la relació entre amplada i alçada, el LAI i les "
    "temperatures radiomètrica i de sòl; la resta són les condicions de l'escena i "
    "es repeteixen a les 118 files. El LAI hi entra com a valor de parcel·la "
    "elemental i no de cep: els 118 ceps comparteixen catorze valors de LAI. "
    "Substituint-los a l'equació de l'apartat 4.7.3 se'n recuperen fIPAR diàries "
    "d'entre 0,37 i 0,65, totes dins del rang amb què es va ajustar la relació, cosa "
    "que confirma que els valors del fitxer surten efectivament d'aquesta via i no "
    "de la inversió del ceptòmetre.")

el = body(el,
    "Amb aquesta taula el model resol el balanç d'energia de cada cep i en retorna "
    "el flux de calor latent, que és l'evapotranspiració real, juntament amb la "
    "partició entre la transpiració del dosser i l'evaporació del sòl.")

# ---------------------------------------------------------------------------
# Punts a verificar del fitxer
# ---------------------------------------------------------------------------
el = body(el,
    "[A VERIFICAR] Al fitxer del 3 de juny la temperatura del dosser és la mateixa "
    "per als 118 ceps (298,20 K, és a dir 25,05 °C), mentre que la del sòl i la "
    "radiomètrica varien de cep a cep. Amb una sola Tc, les diferències d'ETa entre "
    "tractaments només poden venir de la geometria de la capçada i del LAI, i no de "
    "la temperatura, que és justament la via que la tesi vol posar a prova. Cal "
    "confirmar si això és deliberat o si falta assignar a cada cep la seva Tc.",
    color=RGBColor(0xC0, 0x00, 0x00))

el = body(el,
    "[A VERIFICAR] Dotze dels catorze valors de LAI cobreixen entre cinc i sis ceps, "
    "és a dir una parcel·la elemental, però dos se'n reparteixen 23 i 24, i sumen 47 "
    "dels 118 ceps. Convé comprovar si aquestes parcel·les no tenien mesura de fIPAR "
    "i se'ls va assignar un valor comú.",
    color=RGBColor(0xC0, 0x00, 0x00))

body(el,
    "[A VERIFICAR] Les unitats de Sp i Sp_daily s'han deduït de la magnitud dels "
    "valors (965 i 31) i de la data; convé confirmar-les amb l'estació de "
    "referència.",
    color=RGBColor(0xC0, 0x00, 0x00))

doc.save(PATH)
print(f"OK · paràgrafs: {len(doc.paragraphs)} · taules: {len(doc.tables)}")
