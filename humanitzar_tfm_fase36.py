# -*- coding: utf-8 -*-
"""
Fase 36: el protocol complet de la fracció de coberta i la cadena del LAI.

1. La fc final de cada vol és el màxim entre la del NDVI > 0,5 i la de la
   màscara d'alçades, cep a cep.
2. La fIPAR instantània del ceptòmetre es passa a fIPAR diària amb el model
   d'Oyarzún et al. (2007) abans d'entrar a la relació empírica del LAI.
3. Es mesura un cep per parcel·la elemental, de manera que hi ha vint fIPAR
   per vol i els sis ceps de la parcel·la en comparteixen el valor.
"""
import copy

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
# 1. La fc final: el màxim de les dues mesures
# ===========================================================================
p = P("El llindar i la seqüència són els mateixos a totes les dates")
set_text(p, (
    "El llindar i la seqüència són els mateixos a totes les dates, de manera "
    "que les diferències de fracció de coberta entre vols i entre tractaments "
    "no arrosseguen cap canvi de criteri. La fracció de coberta que "
    "s'utilitza a partir d'aquí, però, no és directament la del NDVI: per a "
    "cada cep i cada vol es pren el més gran dels dos valors, el de la capa "
    "de NDVI i el del recompte de píxels de la màscara d'alçades. Les dues "
    "mesures acoten el mateix dosser per vies independents i cadascuna falla "
    "per defecte en situacions diferents, de manera que quedar-se amb el "
    "màxim corregeix el defecte sense afegir cobertura que no s'hagi vist en "
    "cap de les dues capes."
))
el = body(p._element, (
    "A la pràctica el criteri només té efecte on hi ha poca fulla. Al vol del "
    "3 de juny, que és el que té el núvol de punts més dens, la màscara "
    "d'alçades guanya en els 118 ceps i la fracció de coberta combinada "
    "coincideix exactament amb la geomètrica. Als tres vols següents la capa "
    "de NDVI aporta un valor més alt en 32, 44 i 78 ceps respectivament, amb "
    "increments de fins a 0,06, 0,11 i 0,14, i aquests ceps es concentren en "
    "els tractaments de capçada més petita: dels 32 del 30 de juny, 31 són "
    "d'esporga severa amb desfullat, esporga severa o esporga mitjana, i dels "
    "44 del 22 de juliol, 40. Al 12 d'agost el criteri ja afecta tots cinc "
    "tractaments, perquè a final de campanya el núvol de punts tampoc no "
    "resol bé la vora de les capçades grans."
))
el = body(el, (
    "El preu d'aquesta decisió és que la fracció de coberta deixa de ser "
    "estrictament la projecció dels píxels classificats per alçada i passa a "
    "ser la dels píxels fotosintèticament actius vistos des del nadir, que "
    "d'altra banda és el que el model de balanç energètic necessita per "
    "repartir la radiació interceptada entre dosser i sòl. Els dos conjunts "
    "conviuen: als fitxers de sortida del processament hi consta la fracció "
    "geomètrica, i als fitxers d'entrada del model, la combinada."
))
print("1. protocol de la fc (màxim de les dues)  OK")

# ===========================================================================
# 2. Taula 11: afegir la fc combinada
# ===========================================================================
t11 = [t for t in doc.tables
       if t.cell(0, 0).text.strip() == "Variable"
       and "Àrea de copa" in t._tbl.xml][0]
fila_ndvi = None
for r in t11.rows:
    if r.cells[0].text.strip().startswith("fc (NDVI"):
        fila_ndvi = r
        break
if fila_ndvi is None:
    raise LookupError("fila fc (NDVI) de la Taula 11")

mida = 11
for rn in fila_ndvi.cells[0].paragraphs[0].runs:
    if rn.font.size is not None:
        mida = rn.font.size.pt
        break

nova = copy.deepcopy(fila_ndvi._tr)
fila_ndvi._tr.addnext(nova)
r_nova = [r for r in t11.rows if r._tr is nova][0]
cell_set(r_nova.cells[0], "fc (combinada)", mida)
cell_set(r_nova.cells[1],
         "Màxim de les dues anteriors, cep a cep i vol a vol. És la que entra "
         "al model", mida)
cell_set(r_nova.cells[2], "–", mida)
print("2. Taula 11: fila fc combinada  OK")

# ===========================================================================
# 3. El disseny de mostreig de la fIPAR
# ===========================================================================
p = P("El ceptòmetre AccuPAR LP-80 de METER Group mesura la radiació PAR")
el = body(p._element, (
    "El mostreig es fa per parcel·la elemental i no per cep. A cada parcel·la "
    "es tria un cep representatiu i s'hi pren la lectura, de manera que cada "
    "vol dona vint fIPAR, una per parcel·la, i els sis ceps que la componen "
    "hereten el valor de la parcel·la. Això fixa la unitat experimental de "
    "totes les variables derivades de la fIPAR, el LAI inclòs: hi ha vint "
    "observacions independents per data, quatre per tractament, i no 118. Al "
    "vol del 3 de juny només se'n van poder prendre catorze: vuit parcel·les "
    "van quedar sense lectura pròpia i se'ls va assignar el valor de les "
    "parcel·les del mateix tractament que sí que en tenien, de manera que en "
    "aquella data dos valors cobreixen 23 i 24 ceps. És la data amb menys "
    "informació independent i convé tenir-ho present en qualsevol comparació "
    "que la faci servir."
))
print("3. disseny de mostreig de la fIPAR  OK")

p = P("Nombre de ceps mostrejats per tractament i bloc en cada mesura")
set_text(p, (
    "[A CONCRETAR] Hores exactes de les corbes diürnes i model de cambra de "
    "pressió."
), color=VERMELL)
print("   flag de 4.7.1 escurçat  OK")

# ===========================================================================
# 4. De la fIPAR instantània a la fIPAR diària
# ===========================================================================
p = P("El ceptòmetre no mesura el LAI: el dedueix de la fracció de radiació")
el = body(p._element, (
    "Abans d'entrar en qualsevol relació, la lectura del ceptòmetre s'ha de "
    "convertir. El LP-80 dona la fracció interceptada en el moment de la "
    "mesura, que es pren al migdia solar i que per tant correspon a l'angle "
    "solar més alt del dia, que en vinya en filera és precisament el pitjor "
    "moment per a la intercepció (apartat 2.6). La variable que descriu el "
    "dosser és la fracció interceptada al llarg de tota la jornada, perquè "
    "integra el recorregut del sol sobre la filera i no depèn de l'hora a què "
    "s'hagi passat pel camp. El pas d'una a l'altra es fa amb el model "
    "geomètric d'Oyarzún et al. (2007), plantejat per a cultius llenyosos en "
    "filera, que reconstrueix la corba diürna d'intercepció a partir de la "
    "geometria del dosser i de la posició del sol i en dona el valor diari "
    "ponderat per la radiació incident. Totes les fIPAR d'aquest apartat i de "
    "la resta del treball són diàries."
))
print("4. pas a fIPAR diària amb el model d'Oyarzún  OK")

# ===========================================================================
# 5. L'equació: variable, rang i nombre de ceps
# ===========================================================================
p = P("LAI = 1,2615")
set_text(p, "LAI = 1,2615 · ln(fIPARₑ) + 2,8671            R² = 0,6943")

p = P("L'equació explica prop del 70 % de la variació del LAI mesurat")
set_text(p, (
    "La fIPARₑ de l'equació és la fIPAR diària expressada com a fracció "
    "entre 0 i 1, no en percentatge. L'ajust explica prop del 70 % de la "
    "variació del LAI mesurat, raonable tenint en compte que es treballa amb "
    "quinze ceps i que la mesura destructiva i la de radiació no es prenen "
    "exactament al mateix instant. Substituint-hi la fIPAR diària de cada "
    "parcel·la de l'assaig ADAPTEX se n'obté el LAI dels seus 118 ceps, i "
    "aquests valors sí que resulten coherents amb l'amplada, l'alçada i la "
    "fracció de coberta que les imatges del dron donen per als mateixos ceps. "
    "Són també els valors amb què es calibra el model de Random Forest de "
    "l'apartat 4.8, que estén l'estimació als ceps de la resta de la "
    "parcel·la, on no hi ha mesures de fIPAR."
))
print("5. equació: variable, rang i n = 15  OK")

# ===========================================================================
# 6. El flag dels catorze valors de LAI del 3 de juny, resolt
# ===========================================================================
p = P("L'assaig té vint parcel·les elementals, de sis ceps cadascuna")
set_text(p, (
    "El LAI té resolució de parcel·la elemental i no de cep, perquè surt "
    "d'aplicar l'equació de l'apartat 4.7.3 a la fIPAR diària, que es mesura "
    "en un cep per parcel·la. Els fitxers del 30 de juny, del 23 de juliol i "
    "del 12 d'agost porten per tant vint valors, un per parcel·la, i el del "
    "3 de juny, catorze, perquè aquell dia vuit parcel·les van quedar sense "
    "lectura pròpia i comparteixen dos valors que cobreixen 23 i 24 ceps."
))
print("6. flag dels catorze valors de LAI reescrit  OK")

# ===========================================================================
# 7. Oyarzún: accent, citació i referència
# ===========================================================================
p = P("Per això la mesura al camp s'integra amb el model geomètric")
t = p.text.replace("model geomètric d'Oyarzun",
                   "model geomètric d'Oyarzún et al. (2007)")
set_text(p, t)

cap = P("Variació horària de la fiPAR estimada pel model geomètric d'Oyarzun")
mida_cap = 10
for rn in cap.runs:
    if rn.font.size is not None:
        mida_cap = rn.font.size.pt
        break
set_text(cap, cap.text.replace("model geomètric d'Oyarzun",
                               "model geomètric d'Oyarzún et al. (2007)"),
         size=mida_cap)

ref_ant = P("Norman, J. M., Kustas, W. P., & Humes, K. S. (1995)")
mida_ref = 12
for rn in ref_ant.runs:
    if rn.font.size is not None:
        mida_ref = rn.font.size.pt
        break
nova_ref = after(ref_ant._element)
nova_ref.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
nova_ref.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
nova_ref.paragraph_format.space_after = Pt(6)
_font(nova_ref.add_run(
    "Oyarzún, R. A., Stöckle, C. O., & Whiting, M. D. (2007). A simple "
    "approach to modeling radiation interception by fruit-tree orchards. "
    "Agricultural and Forest Meteorology, 142(1), 12–24. "
    "https://doi.org/10.1016/j.agrformet.2006.10.004"), mida_ref)
print("7. Oyarzún: accent, citació i referència  OK")

doc.save(PATH)
print("\nDesat:", PATH)
