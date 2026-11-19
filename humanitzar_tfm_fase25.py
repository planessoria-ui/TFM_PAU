# -*- coding: utf-8 -*-
"""
Fase 25: corregeix com es processa la parcel·la.

La zona de l'assaig ADAPTEX es processa d'una sola vegada, amb un únic polígon
d'àrea d'interès i una única capa de ceps. Les tessel·les s'utilitzen després,
per a la resta de ceps de la parcel·la general, amb l'objectiu de construir el
conjunt de dades amb què s'entrena el model de Random Forest.
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
# 1. L'assaig es processa d'una sola vegada
# ---------------------------------------------------------------------------
p = P("no es processa d’un sol cop")
set_text(p,
    "La zona de l'assaig ADAPTEX es processa d'una sola vegada. L'script rep el "
    "polígon que en delimita l'àrea d'interès i la capa que conté un element per "
    "cep, i recorre els 118 ceps en una única execució. La Figura 22 mostra aquesta "
    "capa de polígons sobreposada a l'ortomosaic del vol: cada rectangle marca els "
    "4,25 m² que corresponen a un cep i és la unitat sobre la qual es calculen tots "
    "els estadístics.")
el = p._element

pic = after(el)
pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
pic.paragraph_format.space_before = Pt(10)
pic.paragraph_format.space_after = Pt(2)
pic.add_run().add_picture(FIG + "fig22_poligons_ortomosaic.png", width=Cm(15.0))
el = pic._element

cap = after(el)
cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
cap.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
cap.paragraph_format.space_before = Pt(6)
cap.paragraph_format.space_after = Pt(6)
_font(cap.add_run("Figura 22."), 10, bold=True)
_font(cap.add_run(
    " Capa de polígons de l'assaig ADAPTEX sobreposada a l'ortomosaic. Cada "
    "rectangle delimita els 4,25 m² que corresponen a un cep, agrupats de sis en "
    "sis per parcel·la experimental, i és la unitat sobre la qual es calculen els "
    "estadístics zonals. Font: elaboració pròpia."), 10)
el = cap._element

# ---------------------------------------------------------------------------
# 2. Les tessel·les: la resta de la parcel·la i el Random Forest
# ---------------------------------------------------------------------------
el = body(el,
    "El mateix processament s'aplica després a la resta de ceps de la parcel·la "
    "general. Com que la superfície és molt més gran, aquí sí que es treballa per "
    "tessel·les: cada tessel·la té el seu polígon d'àrea d'interès i la seva capa de "
    "ceps, un script auxiliar les recorre una darrere l'altra i de cadascuna en surt "
    "un fitxer amb les mateixes variables. Tots aquests fitxers, junts, formen el "
    "conjunt de dades amb què s'entrena el model de Random Forest descrit a "
    "l'apartat 4.8. El model aprèn la relació entre els paràmetres estructurals i "
    "espectrals i el LAI en els ceps on el LAI es coneix, i és el que permet "
    "estimar-lo als ceps on no s'ha mesurat la fIPAR.")

# ---------------------------------------------------------------------------
# 3. Fitxers de sortida: deixar de parlar de blocs
# ---------------------------------------------------------------------------
p = P("El resultat del processament de cada bloc és un fitxer de text")
set_text(p,
    "El processament produeix un fitxer de text delimitat amb una fila per cep i per "
    "data i quaranta-tres columnes, que la Taula 13 agrupa per famílies. Hi ha la "
    "identificació del cep amb el seu tractament, el bloc experimental i el marc de "
    "plantació; els estadístics dels cinc índexs espectrals, amb mitjana i desviació "
    "estàndard, tant per als píxels de dosser com per als de l'entrefilera; les tres "
    "temperatures; i els paràmetres estructurals de la Taula 11. Els fitxers dels "
    "quatre vols, concatenats, formen la taula de treball sobre la qual s'ajusten el "
    "model de Random Forest per al LAI i el model TSEB-2T. Se'n desen també les "
    "capes ràster intermèdies, és a dir la temperatura de superfície i de sòl, el "
    "NDVI, l'alçada i les màscares de dosser i d'entrefilera, que permeten comprovar "
    "visualment que la classificació ha funcionat abans de donar les dades per "
    "bones.")

# ---------------------------------------------------------------------------
# 4. Introducció de l'Annex B
# ---------------------------------------------------------------------------
p = P("S'executa un cop per bloc de la parcel·la")
set_text(p,
    "Es reprodueix l'script de processament corresponent al vol del 3 de juny de "
    "2026, tal com es descriu als apartats 4.6.1 a 4.6.4. Per a la zona de l'assaig "
    "s'executa un sol cop i, per a la resta de la parcel·la, un cop per tessel·la, "
    "amb el codi de la tessel·la com a única variable que canvia entre execucions. "
    "Es recolza en la llibreria interna raster_ET4DROUGHT, que no es reprodueix aquí "
    "per extensió i que és la que implementa les classes de col·lecció d'imatges, el "
    "remostreig, l'emmascarament i els estadístics zonals. Les úniques modificacions "
    "respecte de l'original són el tall d'algunes línies massa llargues per a "
    "l'amplada de la pàgina; els comentaris i els noms de variable es mantenen tal "
    "com són al codi.")

doc.save(PATH)
print(f"OK · paràgrafs: {len(doc.paragraphs)}")
