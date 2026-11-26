# -*- coding: utf-8 -*-
"""
Fase 35: decisions de mètode del processament biofísic que faltaven documentar.

1. Per què la temperatura de dosser es pren al quartil més fred i no al
   percentil 97 com la del sòl i la radiomètrica.
2. Per què s'ajusta una recta tèrmica per vol i no s'utilitza la recta
   històrica: amb la històrica la Tc sortia incoherent amb la temperatura de
   l'aire als quatre vols.
3. Com es calcula realment la fracció de coberta: capa de NDVI > 0,5 sobre
   l'ortomosaic i estadístiques de zona a QGIS, perquè el núvol de punts dels
   tres últims vols no reprodueix l'amplada del dosser.
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
# 1. El criteri del percentil de la temperatura de dosser
# ===========================================================================
p = P("La temperatura de superfície del conjunt de l'escena i la del sòl")
set_text(p, (
    "La temperatura de superfície del conjunt de l'escena i la del sòl es "
    "calculen com la mitjana dels píxels compresos entre els percentils 3 i 97 "
    "de cada polígon. La temperatura de dosser, en canvi, es calcula entre els "
    "percentils 3 i 25, és a dir sobre el quartil més fred dels píxels de "
    "dosser. El motiu de no aplicar-hi el mateix límit superior que a les "
    "altres dues és que, amb el percentil 97, hi entren els píxels de la vora "
    "de la capçada, que es solapen amb l'entrefilera i estan barrejats amb "
    "sòl: amb aquest criteri la temperatura de dosser puja i s'acosta a la del "
    "sòl, en lloc de descriure la fulla. El retall al quartil més fred busca "
    "precisament els píxels de fulla pura, que són els més freds perquè "
    "transpiren. La temperatura de dosser i la de sòl són les dues entrades "
    "tèrmiques del model TSEB-2T descrit a l'apartat 4.8."
))
print("1. criteri del percentil de la Tc  OK")

# ===========================================================================
# 2. Per què una recta tèrmica per vol i no la recta històrica
# ===========================================================================
cap9 = P("Taula 9. Rectes de calibratge de la banda tèrmica")
body(cap9._element, (
    "Que la recta s'ajusti vol a vol, i no s'utilitzi la recta històrica del "
    "sensor, respon a un problema detectat en el primer processament. Amb la "
    "recta històrica la temperatura de dosser sortia incoherent amb la "
    "temperatura de l'aire registrada el mateix dia, bastant més alta, i això "
    "passava als quatre vols de la campanya. Es van provar diversos "
    "calibratges abans de tancar les capes tèrmiques i l'única cosa que la va "
    "fer baixar va ser ajustar la recta amb els elements de referència dels "
    "vols d'aquesta mateixa campanya, cosa que apunta a una descalibració del "
    "sensor tèrmic de l'Altum-PT. Les quatre rectes de la Taula 9 són les que "
    "s'han utilitzat, i les temperatures de tot el treball en surten."
))
print("2. recta per vol enfront de recta històrica  OK")

# ===========================================================================
# 3. El que queda obert del calibratge tèrmic
# ===========================================================================
p = P("La calibració de la banda tèrmica presenta dos problemes")
set_text(p, (
    "[A VERIFICAR] Del calibratge tèrmic queden dues coses per tancar. La "
    "primera és d'ajust: cap de les quatre rectes de la Taula 9 es reprodueix "
    "ajustant per mínims quadrats els cinc elements de referència del mateix "
    "full, que donen pendents d'entre 0,0164 i 0,0183 enfront dels "
    "0,0130–0,0132 registrats. La segona és que, un cop aplicada la recta de "
    "la campanya, la temperatura de dosser encara no concorda del tot amb la "
    "que es va mesurar a camp amb el termòmetre d'infraroig: comparant plot a "
    "plot la Tc mitjana dels sis ceps amb la mesura de l'IRT del full "
    "TEMP_v5, el dron dona de mitjana 3,5 °C més, amb un error quadràtic "
    "mitjà de 4,2 °C. La diferència tampoc no és constant, sinó que creix a "
    "mesura que disminueix la fracció de coberta del tractament, des de "
    "+0,7 °C de mitjana al tractament sense esporga fins a +5,7 °C a "
    "l'esporga severa amb desfullat; la correlació entre la fracció de "
    "coberta del tractament i el biaix és de −0,94. Això és compatible amb "
    "una contaminació residual dels píxels de dosser per sòl calent a les "
    "capçades més obertes, que el retall al percentil 25 no acaba "
    "d'eliminar. Com que el biaix va en la mateixa direcció que l'efecte dels "
    "tractaments que es vol mesurar, convé comprovar amb l'equip l'hora de "
    "les mesures de l'IRT respecte de la del vol i decidir si cal endurir el "
    "llindar de la màscara de dosser o abaixar encara més el percentil."
), color=VERMELL)
print("3. flag del calibratge tèrmic reescrit  OK")

# ===========================================================================
# 4. Taula 11: distingir la fc geomètrica de la que entra al model
# ===========================================================================
t11 = [t for t in doc.tables
       if t.cell(0, 0).text.strip() == "Variable"
       and "Àrea de copa" in t._tbl.xml][0]
fila_fc = None
for r in t11.rows:
    if r.cells[0].text.strip() == "fc":
        fila_fc = r
        break
if fila_fc is None:
    raise LookupError("fila fc de la Taula 11")

mida = 11
for rn in fila_fc.cells[0].paragraphs[0].runs:
    if rn.font.size is not None:
        mida = rn.font.size.pt
        break

cell_set(fila_fc.cells[0], "fc (geomètrica)", mida)
cell_set(fila_fc.cells[1], "Àrea de copa / marc de plantació", mida)

nova = copy.deepcopy(fila_fc._tr)
fila_fc._tr.addnext(nova)
r_nova = [r for r in t11.rows if r._tr is nova][0]
cell_set(r_nova.cells[0], "fc (NDVI > 0,5)", mida)
cell_set(r_nova.cells[1],
         "Píxels amb NDVI > 0,5 dins del polígon / píxels totals del polígon",
         mida)
cell_set(r_nova.cells[2], "–", mida)
print("4. Taula 11: fila fc desdoblada  OK")

# ===========================================================================
# 5. Com es calcula la fracció de coberta que entra al model
# ===========================================================================
cap11 = P("Taula 11. Paràmetres estructurals del dosser calculats per a cada cep")
el = cap11._element

el = body(el, (
    "La fracció de coberta, però, no es calcula d'aquesta manera. El recompte "
    "de píxels de dosser depèn de la màscara d'alçades, i aquesta depèn al seu "
    "torn de la densitat del núvol de punts amb què s'ha reconstruït el model "
    "digital de vegetació. El vol del 3 de juny es va reconstruir amb 343,6 "
    "milions de punts i els blocs de l'assaig dels altres tres, amb entre "
    "120,5 i 125,5 milions (Taula 15). Amb un núvol tres vegades menys dens "
    "l'amplada del dosser es queda curta: els píxels de la vora de la paret "
    "vegetal no arriben a superar el llindar de 0,5 m i queden fora de la "
    "màscara. L'efecte és més gran com més petita és la capçada, de manera que "
    "afecta sobretot l'esporga severa i l'esporga severa amb desfullat, que "
    "són justament els tractaments la fracció de coberta dels quals s'ha de "
    "poder comparar amb la del control."
))

el = body(el, (
    "Per això la fracció de coberta es recalcula sobre l'ortomosaic, que "
    "recull tota l'amplada del dosser, i no sobre la capa d'alçades. El "
    "procediment té tres passos i s'aplica igual als quatre vols. Primer es "
    "genera la capa de NDVI de l'apartat 4.6.3. Després, amb una operació "
    "ràster, es posa a NoData tot el que no arriba a un NDVI de 0,5, que és el "
    "llindar amb què queden fora el sòl i la coberta de l'entrefilera i es "
    "conserva només el dosser, i la capa resultant es retalla pel marc de "
    "plantació de cada cep. Finalment, amb les estadístiques de zona de QGIS "
    "es compta quants píxels de dosser queden dins de cada polígon i es "
    "divideix pel nombre total de píxels del marc; aquest quocient és la "
    "fracció de coberta del cep."
))

el = body(el, (
    "El llindar i la seqüència són els mateixos a totes les dates, de manera "
    "que les diferències de fracció de coberta entre vols i entre tractaments "
    "no arrosseguen cap canvi de criteri. A canvi, la fracció de coberta deixa "
    "de ser la projecció dels píxels classificats per alçada i passa a ser la "
    "dels píxels fotosintèticament actius vistos des del nadir, que és el que "
    "el model de balanç energètic necessita per repartir la radiació "
    "interceptada entre dosser i sòl. Els dos conjunts conviuen als fitxers de "
    "sortida: la fc geomètrica de la Taula 11 i la fc del NDVI, que és la que "
    "s'utilitza a partir d'aquí i la que entra als fitxers d'entrada del "
    "model."
))
print("5. mètode de la fc sobre NDVI > 0,5  OK")

doc.save(PATH)
print("\nDesat:", PATH)
