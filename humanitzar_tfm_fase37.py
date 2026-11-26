# -*- coding: utf-8 -*-
"""
Fase 37: el fitxer d'entrada del TSEB-2T, columna a columna.

1. La meteorologia no surt sempre de la mateixa estació: el primer vol es va
   fer abans de muntar la de la parcel·la i utilitza la de Raimat.
2. La temperatura de dosser del fitxer és un sol valor per vol, el més baix
   dels ceps del tractament de control, i no la Tc de cada cep.
3. Queda la constatació que al vol del 23 de juliol el valor adoptat no és
   d'un cep de control sinó del mínim absolut del vol.
"""
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PATH = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v8.docx"
doc = docx.Document(PATH)
FONT = "Arial"
VERMELL = RGBColor(0xC0, 0x00, 0x00)
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


def body(el, text, color=None):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    p.paragraph_format.space_after = Pt(6)
    _font(p.add_run(text), 12, color=color)
    return p._element


def cell_set(cl, t, size, bold=False):
    pp = cl.paragraphs[0]
    for r in pp.runs:
        r._element.getparent().remove(r._element)
    _font(pp.add_run(t), size, bold=bold)


# ===========================================================================
# 1. La procedència de la meteorologia, vol a vol
# ===========================================================================
p = P("Les columnes vénen de tres llocs")
set_text(p, (
    "Les columnes vénen de tres llocs. La geometria de la capçada i les "
    "temperatures de sòl i radiomètrica surten del processament de les "
    "imatges, tal com es descriu als apartats 4.6.2 i 4.6.4; el LAI, de la "
    "relació empírica amb la fIPAR diària de l'apartat 4.7.3; i les variables "
    "meteorològiques, de l'estació. Aquí la font canvia amb la data: el vol "
    "del 3 de juny es va fer abans que l'estació de la parcel·la estigués "
    "muntada i les seves dades provenen de l'estació de Raimat de la Xarxa "
    "Agrometeorològica de Catalunya, que és la més propera, mentre que els "
    "tres vols següents ja utilitzen l'estació instal·lada dins de la mateixa "
    "parcel·la d'El polvorín, que descriu millor les condicions en què es va "
    "volar. L'hora que consta a cada fila és la del vol en temps universal "
    "coordinat, de manera que els quatre vols es van fer entre les 11:51 i "
    "les 12:30 UTC."
))
print("1. procedència de la meteorologia  OK")

p = P("registrades per l'estació instal·lada a la mateixa parcel·la d'El polvorín")
mida = 12
for rn in p.runs:
    if rn.font.size is not None:
        mida = rn.font.size.pt
        break
set_text(p, (
    "Variables meteorològiques: temperatura de l'aire, humitat relativa, "
    "dèficit de pressió de vapor (VPD), velocitat del vent, radiació solar i "
    "evapotranspiració de referència. Del segon vol endavant es registren amb "
    "l'estació instal·lada a la mateixa parcel·la d'El polvorín; el primer vol "
    "es va fer abans de muntar-la i utilitza l'estació de Raimat de la Xarxa "
    "Agrometeorològica de Catalunya. Són les entrades meteorològiques dels "
    "models de balanç energètic."
), size=mida)
print("   bullet de 4.7 actualitzat  OK")

# ===========================================================================
# 2. La temperatura de dosser del fitxer d'entrada
# ===========================================================================
p = P("L'script calcula la temperatura del dosser polígon a polígon")
set_text(p, (
    "La temperatura de dosser del fitxer d'entrada mereix una explicació a "
    "part, perquè és l'única columna que no es pren cep a cep. L'script la "
    "calcula polígon a polígon, com la mitjana dels píxels de dosser "
    "compresos entre els percentils 3 i 25, i en dona 118 valors per vol; el "
    "fitxer d'entrada, en canvi, en porta un de sol, el mateix per a tots els "
    "ceps, mentre que la temperatura del sòl i la radiomètrica sí que varien "
    "de cep a cep. El motiu és el biaix tèrmic descrit a l'apartat 4.6: fins "
    "i tot amb la recta de calibratge d'aquesta campanya, la Tc individual "
    "surt massa alta i no és coherent amb la temperatura de l'aire del mateix "
    "dia. El valor que s'adopta és el més baix dels ceps del tractament de "
    "control, que és el de capçada més gran i el que més s'acosta a un dosser "
    "sense limitació de transpiració, i per tant el menys contaminat per sòl "
    "calent."
))
el = body(p._element, (
    "Amb aquest criteri la temperatura de dosser queda per sota de la de "
    "l'aire als quatre vols: −2,8 °C el 3 de juny, −3,0 el 30 de juny, −1,6 "
    "el 23 de juliol i −4,8 el 12 d'agost, que és el comportament que "
    "s'espera d'un dosser que transpira. El preu és que la variabilitat "
    "tèrmica entre ceps i entre tractaments no entra al model: amb una sola "
    "Tc, les diferències d'ETa entre tractaments només poden venir de la "
    "geometria de la capçada, del LAI i de les temperatures de sòl i "
    "radiomètrica, que sí que es donen cep a cep. Cal tenir-ho present en "
    "llegir els apartats 5.3 i 5.4."
))
body(el, (
    "[A VERIFICAR] En tres dels quatre vols el valor adoptat correspon "
    "efectivament a un cep del tractament de control: el cep 117 (p20) el 3 "
    "de juny, el 98 (p17) el 30 de juny i el 116 (p20) el 12 d'agost. Al vol "
    "del 23 de juliol, en canvi, els 33,99 °C del fitxer corresponen al cep "
    "74 (p16), que és d'esporga lleugera; el mínim dels ceps de control "
    "d'aquell dia és de 34,92 °C, 0,93 °C més alt. Convé confirmar si aquell "
    "dia es va prendre el mínim absolut del vol en lloc del mínim del "
    "control, perquè una Tc més baixa es tradueix en més transpiració per a "
    "tots els ceps alhora."
), color=VERMELL)
print("2. temperatura de dosser del fitxer d'entrada  OK")

# ===========================================================================
# 3. Taula 17: la font de la meteorologia i el criteri de la T_C
# ===========================================================================
t17 = [t for t in doc.tables if t.cell(0, 0).text.strip() == "Columna"][0]
mida_t = 10
for rn in t17.rows[1].cells[0].paragraphs[0].runs:
    if rn.font.size is not None:
        mida_t = rn.font.size.pt
        break
for r in t17.rows:
    c0 = r.cells[0].text.strip()
    if c0.startswith("Meteorologia"):
        cell_set(r.cells[0],
                 "Meteorologia (estació de Raimat al primer vol, estació de "
                 "la parcel·la als altres tres)", mida_t, bold=True)
    elif c0 == "T_C":
        cell_set(r.cells[1],
                 "Temperatura del dosser: valor més baix dels ceps del "
                 "tractament de control", mida_t)
print("3. Taula 17 actualitzada  OK")

doc.save(PATH)
print("\nDesat:", PATH)
