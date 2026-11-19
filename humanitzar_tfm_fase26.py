# -*- coding: utf-8 -*-
"""
Fase 26: els blocs de processament de la parcel·la general.

Torna al terme "bloc de processament", que és el que apareix a la figura
original, i hi afegeix la Figura 23, amb el repartiment dels 208 ceps de
referència entre els nou blocs de la parcel·la ZV_CV.
"""
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PATH = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v8.docx"
FIG = "/home/user/TFM_PAU/figures/"
doc = docx.Document(PATH)
FONT = "Arial"
_PARENT = doc.paragraphs[0]._parent


def _font(run, size=12, bold=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
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


# ---------------------------------------------------------------------------
# 1. Els blocs de processament de la parcel·la general
# ---------------------------------------------------------------------------
p = P("aquí sí que es treballa per tessel")
set_text(p,
    "El mateix processament s'aplica després a la resta de ceps de la parcel·la "
    "general. Com que la superfície és molt més gran, aquí es treballa per blocs: "
    "cada bloc de processament té el seu polígon d'àrea d'interès i la seva capa de "
    "ceps, un script auxiliar els recorre un darrere l'altre i de cadascun en surt "
    "un fitxer amb les mateixes variables. Es processen per separat per no esgotar "
    "la memòria disponible i, en acabar, tots els fitxers s'ajunten en un de sol.")
el = p._element

el = body(el,
    "Els blocs són una graella de fins a 100 m de costat que cobreix la parcel·la i "
    "no tenen cap significat agronòmic. Els ceps de referència, en canvi, es "
    "reparteixen a l'atzar per tota la superfície, de manera que el nombre de ceps "
    "que toca a cada bloc varia molt. A la parcel·la ZV_CV hi ha 208 ceps distribuïts "
    "en nou blocs: el T2_3 en conté 88 i el T1_0, 75, mentre que els set restants es "
    "mouen entre tres i quinze ceps (Figura 23). Aquest desequilibri no afecta el "
    "càlcul, perquè cada cep es tracta per separat, però sí que fa que el temps de "
    "processament es concentri en dos blocs.")

pic = after(el)
pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
pic.paragraph_format.space_before = Pt(10)
pic.paragraph_format.space_after = Pt(2)
pic.add_run().add_picture(FIG + "fig23_blocs_processament.png", width=Cm(15.5))
el = pic._element

cap = after(el)
cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
cap.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
cap.paragraph_format.space_before = Pt(6)
cap.paragraph_format.space_after = Pt(6)
_font(cap.add_run("Figura 23."), 10, bold=True)
_font(cap.add_run(
    " Blocs de processament de la parcel·la ZV_CV, sobreposats a la màscara de la "
    "parcel·la, amb els 208 ceps de referència repartits a l'atzar i el nombre de "
    "ceps que conté cada bloc. Coordenades en UTM 31N (EPSG:32631), vol del 3 de "
    "juny de 2026. Font: elaboració pròpia."), 10)
el = cap._element

body(el,
    "Els fitxers de tots els blocs, junts, formen el conjunt de dades amb què "
    "s'entrena el model de Random Forest descrit a l'apartat 4.8. El model aprèn la "
    "relació entre els paràmetres estructurals i espectrals i el LAI en els ceps on "
    "el LAI es coneix, i és el que permet estimar-lo als ceps on no s'ha mesurat la "
    "fIPAR.")

# el paràgraf original acabava amb la part del Random Forest, que ara va a part
# ---------------------------------------------------------------------------
# 2. Introducció de l'Annex B
# ---------------------------------------------------------------------------
p = P("un cop per tessel·la, amb el codi de la tessel")
set_text(p,
    "Es reprodueix l'script de processament corresponent al vol del 3 de juny de "
    "2026, tal com es descriu als apartats 4.6.1 a 4.6.4. Per a la zona de l'assaig "
    "s'executa un sol cop i, per a la resta de la parcel·la, un cop per bloc de "
    "processament, amb el codi del bloc com a única variable que canvia entre "
    "execucions. Es recolza en la llibreria interna raster_ET4DROUGHT, que no es "
    "reprodueix aquí per extensió i que és la que implementa les classes de "
    "col·lecció d'imatges, el remostreig, l'emmascarament i els estadístics zonals. "
    "Les úniques modificacions respecte de l'original són el tall d'algunes línies "
    "massa llargues per a l'amplada de la pàgina; els comentaris i els noms de "
    "variable es mantenen tal com són al codi.")

doc.save(PATH)
print(f"OK · paràgrafs: {len(doc.paragraphs)}")
