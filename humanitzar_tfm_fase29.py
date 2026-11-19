# -*- coding: utf-8 -*-
"""
Fase 29: el fitxer d'entrada del model, per als quatre vols.

La Taula 17 passa a descriure només l'estructura del fitxer i la resolució de
cada variable, i la nova Taula 18 recull els valors dels quatre vols. Les notes
de verificació s'actualitzen amb el que mostren els quatre fitxers.
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


def set_text(p, t, size=12, color=None):
    for r in p.runs[1:]:
        r._element.getparent().remove(r._element)
    if p.runs:
        p.runs[0].text = t
        _font(p.runs[0], size, color=color)
    else:
        _font(p.add_run(t), size, color=color)


def cell_text(cl, t, size=9, bold=False):
    pp = cl.paragraphs[0]
    for r in pp.runs:
        r._element.getparent().remove(r._element)
    _font(pp.add_run(t), size, bold=bold)


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
            pp.alignment = (WD_ALIGN_PARAGRAPH.LEFT if c == 0 and r
                            else WD_ALIGN_PARAGRAPH.CENTER)
            pp.paragraph_format.space_before = Pt(2)
            pp.paragraph_format.space_after = Pt(2)
            pp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            parts = val.split("\n")
            run = pp.add_run(parts[0]); _font(run, 8.5, bold=(r == 0 or r in grups))
            for extra in parts[1:]:
                run.add_break()
                run2 = pp.add_run(extra); _font(run2, 8.5, bold=(r == 0 or r in grups))
    el.addnext(t._element)
    return t._element


# ===========================================================================
# 1. Taula 17: l'última columna passa a ser la resolució de cada variable
# ===========================================================================
t17 = [t for t in doc.tables if t.cell(0, 0).text.strip() == "Columna"][0]
RES = {
    "Columna": "Resolució", "tree": "Cep", "date, doy": "Vol", "hour": "Vol",
    "fc": "Cep", "hc": "Cep", "wc_hc": "Cep", "LAI": "Parcel·la elemental",
    "T_R": "Cep", "T_C": "Vol", "T_S": "Cep", "Ta": "Vol", "u10": "Vol",
    "RH": "Vol", "Sp": "Vol", "Eto": "Vol", "Sp_daily": "Vol",
    "Eto_daily": "Vol",
}
for row in t17.rows:
    clau = row.cells[0].text.strip()
    cell_text(row.cells[3], RES.get(clau, ""), bold=(clau == "Columna"))
# una mica més d'espai per al significat, menys per a l'última columna
for row in t17.rows:
    for c, w in zip(row.cells, (Cm(3.0), Cm(7.6), Cm(2.0), Cm(3.4))):
        c.width = w
grid = t17._tbl.find(qn('w:tblGrid'))
for gc, w in zip(grid.findall(qn('w:gridCol')),
                 (Cm(3.0), Cm(7.6), Cm(2.0), Cm(3.4))):
    gc.set(qn('w:w'), str(int(w.twips)))

p17 = P("Columnes del fitxer d'entrada del model TSEB-2T i valors del vol")
set_text(p17, "Taula 17.", size=10)
p17.runs[0].font.bold = True
_font(p17.add_run(
    " Columnes del fitxer d'entrada del model TSEB-2T, amb el seu significat, les "
    "seves unitats i la resolució a què es dona cada variable: un valor per cep, "
    "un per parcel·la elemental o un de sol per a tot el vol. "
    "Font: elaboració pròpia."), 10)

# ===========================================================================
# 2. El text que presenta la taula
# ===========================================================================
set_text(P("El model no es crida sobre la imatge sencera sinó sobre una taula"),
    "El model no es crida sobre la imatge sencera sinó sobre una taula. El fitxer "
    "d'entrada té una fila per cep i per data, i cada fila porta tot el que el "
    "balanç d'energia necessita per a aquell cep en el moment del vol: la geometria "
    "de la capçada, les tres temperatures i les condicions meteorològiques. La "
    "Taula 17 en recull les columnes i la Taula 18, els valors dels quatre vols de "
    "la campanya.")

# ===========================================================================
# 3. Taula 18: els valors dels quatre vols
# ===========================================================================
anc = P("De totes aquestes columnes, les que donen identitat a cada cep")
set_text(anc,
    "No totes les columnes tenen el mateix gra. La fracció de coberta, l'alçada, la "
    "relació entre amplada i alçada i les temperatures radiomètrica i de sòl surten "
    "de la imatge cep a cep i prenen 118 valors diferents. El LAI entra com a valor "
    "de parcel·la elemental, perquè és la unitat en què es mesura la fIPAR: als vols "
    "del 30 de juny, del 23 de juliol i del 12 d'agost hi ha vint valors de LAI, un "
    "per cadascuna de les vint parcel·les de l'assaig. La resta de columnes són les "
    "condicions de l'escena i es repeteixen a les 118 files.")
el = anc._element

el = body(el,
    "Tornant enrere els valors de LAI per l'equació de l'apartat 4.7.3 se'n "
    "recuperen les fIPAR diàries de partida, que van de 0,37 al juny a 0,77 a "
    "l'agost. Totes queden dins del rang amb què es va ajustar la relació, excepte "
    "la més alta del vol d'agost, que el supera lleugerament. La progressió al llarg "
    "de la campanya és la que s'espera: a mesura que la capçada creix, la fracció de "
    "radiació interceptada puja i el LAI que se'n deriva també.")

el = caption(el, "Taula 18.",
             "Valors del fitxer d'entrada del model TSEB-2T per als quatre vols de "
             "la campanya de 2026. En les variables que es donen cep a cep s'indica "
             "el mínim, el màxim i, entre parèntesis, la mitjana dels 118 ceps; en "
             "les que són comunes al vol, el valor únic. Les temperatures s'hi "
             "recullen en graus Celsius, tot i que al fitxer són en kelvin. "
             "Font: elaboració pròpia.")

T18 = [
    ["Variable", "03/06/2026", "30/06/2026", "23/07/2026", "12/08/2026"],
    ["Dia de l'any i hora solar", "", "", "", ""],
    ["doy", "154", "181", "204", "225"],
    ["hour (h)", "12,50", "12,08", "11,85", "12,00"],
    ["Estructura de la capçada", "", "", "", ""],
    ["fc (–)", "0,16–0,39\n(0,23)", "0,15–0,46\n(0,26)", "0,12–0,57\n(0,29)",
     "0,15–0,57\n(0,29)"],
    ["hc (m)", "1,72–2,65\n(2,14)", "1,55–2,44\n(1,99)", "1,67–2,62\n(2,05)",
     "1,70–2,63\n(2,09)"],
    ["wc_hc (–)", "0,41–0,66\n(0,53)", "0,44–0,73\n(0,58)", "0,33–0,80\n(0,58)",
     "0,33–0,78\n(0,55)"],
    ["LAI (m² m⁻²)", "1,61–2,33\n(1,91)", "1,84–2,47\n(2,15)", "1,83–2,39\n(2,10)",
     "2,00–2,53\n(2,25)"],
    ["Parcel·les amb LAI propi", "14", "20", "20", "20"],
    ["Temperatures (°C)", "", "", "", ""],
    ["T_R", "38,7–49,1\n(45,1)", "40,8–51,8\n(47,8)", "41,9–52,2\n(48,2)",
     "39,9–50,4\n(46,7)"],
    ["T_C", "25,1", "30,3", "34,0", "32,3"],
    ["T_S", "44,2–53,3\n(49,0)", "45,1–55,0\n(51,6)", "46,1–54,0\n(51,0)",
     "43,0–51,9\n(49,2)"],
    ["Meteorologia del vol", "", "", "", ""],
    ["Ta (°C)", "27,9", "33,3", "35,6", "37,1"],
    ["T_C − Ta (°C)", "−2,8", "−3,0", "−1,6", "−4,8"],
    ["u10 (m s⁻¹)", "3,10", "1,20", "1,76", "1,20"],
    ["RH (–)", "0,395", "0,311", "0,327", "0,203"],
    ["Sp (W m⁻²)", "965", "886", "857", "747"],
    ["Eto (mm h⁻¹)", "0,650", "0,691", "0,681", "0,690"],
    ["Sp_daily (MJ m⁻² d⁻¹)", "31,0", "30,0", "27,9", "25,4"],
    ["Eto_daily (mm d⁻¹)", "5,59", "6,66", "6,51", "5,98"],
]
W18 = [Cm(4.4), Cm(2.9), Cm(2.9), Cm(2.9), Cm(2.9)]
el = table(el, T18, W18, grups=(1, 4, 10, 14))

el = body(el,
    "Les condicions meteorològiques acompanyen l'estació: la temperatura de l'aire "
    "puja de 27,9 °C al juny a 37,1 °C a l'agost, la humitat relativa baixa fins al "
    "20 % i la radiació incident en el moment del vol disminueix de 965 a 747 W m⁻², "
    "cosa que és coherent amb la pèrdua d'alçada del sol a mesura que avança "
    "l'estiu. L'evapotranspiració de referència diària té el màxim al vol del 30 de "
    "juny, amb 6,66 mm.")

el = body(el,
    "Amb aquesta informació el model resol el balanç d'energia de cada cep i en "
    "retorna el flux de calor latent, que és l'evapotranspiració real, juntament amb "
    "la partició entre la transpiració del dosser i l'evaporació del sòl.")

# ===========================================================================
# 4. Les notes de verificació, actualitzades amb els quatre fitxers
# ===========================================================================
set_text(P("[A VERIFICAR] Al fitxer del 3 de juny la temperatura del dosser"),
    "[A VERIFICAR] Als quatre fitxers la temperatura del dosser és un valor únic per "
    "a tot el vol, el mateix per als 118 ceps, mentre que la del sòl i la "
    "radiomètrica varien de cep a cep. Les diferències respecte de l'aire són de "
    "−2,8 °C el 3 de juny, −3,0 el 30 de juny, −1,6 el 23 de juliol i −4,8 el 12 "
    "d'agost, de manera que es comporten com la temperatura d'un dosser ben regat i "
    "no com la de cada tractament. Amb una sola Tc, les diferències d'ETa entre "
    "tractaments només poden venir de la geometria de la capçada i del LAI, i no de "
    "la temperatura, que és justament la via que la tesi vol posar a prova. Cal "
    "aclarir si és deliberat, i en aquest cas explicar-ho al text, o si falta "
    "assignar a cada cep la seva Tc.",
    color=RGBColor(0xC0, 0x00, 0x00))

set_text(P("[A VERIFICAR] Dotze dels catorze valors de LAI cobreixen"),
    "[A VERIFICAR] L'assaig té vint parcel·les elementals, de sis ceps cadascuna "
    "excepte dues que en tenen cinc, i els fitxers del 30 de juny, del 23 de juliol "
    "i del 12 d'agost donen efectivament vint valors de LAI. El del 3 de juny, en "
    "canvi, només en té catorze: vuit parcel·les hi comparteixen dos valors, un per "
    "a 23 ceps i un altre per a 24. Convé comprovar si aquelles parcel·les no tenien "
    "mesura de fIPAR en aquella data i se'ls va assignar un valor comú.",
    color=RGBColor(0xC0, 0x00, 0x00))

set_text(P("[A VERIFICAR] Les unitats de Sp i Sp_daily"),
    "[A VERIFICAR] El dia de l'any dels fitxers confirma la data en tres dels quatre "
    "vols. El del 23 de juliol porta doy 204, que és efectivament el 23 de juliol, "
    "de manera que caldria corregir la data del 22 de juliol que consta a la Taula 3 "
    "i a la Taula 14. El del 12 d'agost porta doy 225, que correspon al 13 d'agost, "
    "la mateixa data que apareix al full de calibratge tèrmic; aquí cal decidir si "
    "el vol va ser el 12 o el 13.",
    color=RGBColor(0xC0, 0x00, 0x00))

doc.save(PATH)
print(f"OK · paràgrafs: {len(doc.paragraphs)} · taules: {len(doc.tables)}")
