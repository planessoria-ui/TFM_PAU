# -*- coding: utf-8 -*-
"""
Fase 31: les tres capes ràster sobre les quals es generen les màscares.

Obre l'apartat 4.6.1 amb l'ortomosaic multiespectral, la capa tèrmica i el
model d'alçades (Figura 25), explica per què la temperatura de dosser es pren
sobre el quartil més fred, i corregeix el peu de l'esquema del TSEB-2T, que
descrivia un criteri d'escena que no és el que fa l'script.
"""
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PATH = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v8.docx"
FIG = "/home/user/TFM_PAU/figures/"
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
# 1. Les tres capes, a l'obertura del 4.6.1
# ===========================================================================
h461 = [p for p in doc.paragraphs
        if p.style.name == 'Heading 3' and p.text.strip().startswith('4.6.1')][0]
el = h461._element

el = body(el,
    "El processament parteix de tres capes ràster del mateix vol, totes "
    "remostrejades sobre la mateixa graella perquè un píxel ocupi la mateixa "
    "posició a totes tres (Figura 25). La primera és l'ortomosaic multiespectral, "
    "amb les cinc bandes de reflectància i la sisena ja convertida a graus Celsius "
    "amb les rectes de l'apartat anterior. La segona és la capa tèrmica que se'n "
    "deriva, que és d'on surten les temperatures de component. I la tercera és el "
    "model d'alçades del dosser, en metres sobre el terreny.")

el = body(el,
    "Les màscares es construeixen sobre la capa d'alçades i les bandes de "
    "reflectància, i és amb aquestes màscares que després es mostreja la tèrmica: "
    "dins de cada polígon, els píxels classificats com a dosser donen la temperatura "
    "del dosser i els d'entrefilera, la del sòl. La temperatura del dosser es pren "
    "sobre el quartil més fred dels seus píxels, amb el criteri que detalla "
    "l'apartat 4.6.4, perquè n'hi ha prou amb un píxel de sòl colat dins la màscara "
    "per fer-la pujar uns quants graus.")

pic = after(el)
pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
pic.paragraph_format.space_before = Pt(10)
pic.paragraph_format.space_after = Pt(2)
pic.add_run().add_picture(FIG + "fig25_termica_alcades.png", width=Cm(16.0))
el = pic._element

cap = after(el)
cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
cap.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
cap.paragraph_format.space_before = Pt(6)
cap.paragraph_format.space_after = Pt(6)
_font(cap.add_run("Figura 25."), 10, bold=True)
_font(cap.add_run(
    " Les dues capes sobre les quals es classifiquen els p\u00edxels i se'n prenen les "
    "temperatures, al vol del 3 de juny de 2026, amb la capa de pol\u00edgons dels ceps "
    "sobreposada: (A) capa t\u00e8rmica calibrada, en graus Celsius, i (B) model "
    "d'al\u00e7ades del dosser, en metres sobre el terreny. El quadre de valors de cada "
    "panell mostra la lectura del p\u00edxel consultat: 39,19 \u00b0C i 0,657 m, "
    "respectivament. Font: elaboraci\u00f3 pr\u00f2pia."), 10)

# ===========================================================================
# 1b. L'ortomosaic multiespectral, al 4.6.3
# ===========================================================================
h463 = [p for p in doc.paragraphs
        if p.style.name == 'Heading 3' and p.text.strip().startswith('4.6.3')][0]
el = h463._element

el = body(el,
    "Els \u00edndexs es calculen sobre les cinc bandes de reflect\u00e0ncia de l'ortomosaic, "
    "p\u00edxel a p\u00edxel i sobre la mateixa graella que les altres capes (Figura 26).")

pic = after(el)
pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
pic.paragraph_format.space_before = Pt(10)
pic.paragraph_format.space_after = Pt(2)
pic.add_run().add_picture(FIG + "fig26_ortomosaic.png", width=Cm(16.0))
el = pic._element

cap = after(el)
cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
cap.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
cap.paragraph_format.space_before = Pt(6)
cap.paragraph_format.space_after = Pt(6)
_font(cap.add_run("Figura 26."), 10, bold=True)
_font(cap.add_run(
    " Ortomosaic multiespectral del vol del 3 de juny de 2026, amb la capa de "
    "pol\u00edgons dels ceps sobreposada. El quadre de valors mostra, per al p\u00edxel "
    "consultat, la reflect\u00e0ncia de les cinc bandes (0,102; 0,141; 0,235; 0,266 i "
    "0,368) i el valor de la sisena, ja convertit a graus Celsius (41,78 \u00b0C). "
    "Font: elaboraci\u00f3 pr\u00f2pia."), 10)

# ===========================================================================
# 2. El peu de l'esquema del TSEB-2T deia un criteri que no és el de l'script
# ===========================================================================
p10 = P("Esquema del model TSEB-2T")
set_text(p10, "Figura 10.", size=10)
p10.runs[0].font.bold = True
_font(p10.add_run(
    " Esquema del model TSEB-2T. La temperatura del dosser (Tc) i la del sòl (Ts) "
    "s'extreuen de la capa tèrmica calibrada dins del polígon de cada cep, "
    "separant-ne els píxels amb les màscares de l'apartat 4.6.1 i aplicant-hi el "
    "criteri de percentils de l'apartat 4.6.4, i alimenten la xarxa de resistències "
    "del balanç d'energia. Font: elaboració pròpia."), 10)

# ===========================================================================
# 3. La nota de la Tc única, amb el que diu l'script
# ===========================================================================
set_text(P("[A VERIFICAR] Als quatre fitxers la temperatura del dosser"),
    "[A VERIFICAR] L'script calcula la temperatura del dosser polígon a polígon, com "
    "la mitjana dels píxels de dosser compresos entre els percentils 3 i 25, de "
    "manera que el fitxer d'entrada hauria de portar 118 valors diferents. En canvi, "
    "als quatre fitxers hi ha un sol valor per vol, el mateix per a tots els ceps, "
    "mentre que la temperatura del sòl i la radiomètrica sí que varien de cep a cep. "
    "Les diferències respecte de l'aire són de −2,8 °C el 3 de juny, −3,0 el 30 de "
    "juny, −1,6 el 23 de juliol i −4,8 el 12 d'agost, de manera que es comporten com "
    "la temperatura d'un dosser ben regat i no com la de cada tractament. Amb una "
    "sola Tc, les diferències d'ETa entre tractaments només poden venir de la "
    "geometria de la capçada i del LAI, i no de la temperatura, que és justament la "
    "via que la tesi vol posar a prova. Cal aclarir en quin punt es perd la Tc de "
    "cada cep entre la sortida del processament i l'entrada del model.",
    color=RGBColor(0xC0, 0x00, 0x00))

doc.save(PATH)
print(f"OK · paràgrafs: {len(doc.paragraphs)}")
