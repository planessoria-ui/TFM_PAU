# -*- coding: utf-8 -*-
"""
Fase 30: la capa de polígons passa al 4.6.2.

La Figura 22 i el text que la presenta van a l'apartat de la quadrícula i els
paràmetres estructurals, just després de l'aixecament GPS. El bloc de la resta
de la parcel·la (Figura 23 i Random Forest) els segueix, al final del mateix
apartat, de manera que la seqüència de figures hi queda ordenada.
"""
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PATH = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v8.docx"
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


# ---------------------------------------------------------------------------
# Elements que es mouen
# ---------------------------------------------------------------------------
p_adaptex = P("La zona de l'assaig ADAPTEX es processa d'una sola vegada")
el_foto22 = p_adaptex._element.getnext()          # paràgraf amb la imatge
el_peu22 = el_foto22.getnext()                    # peu de la Figura 22

p_blocs1 = P("El mateix processament s'aplica després a la resta de ceps")
p_blocs2 = P("Cada bloc és un rectangle de fins a 100 m")
el_foto23 = p_blocs2._element.getnext()
el_peu23 = el_foto23.getnext()
p_rf = P("Els fitxers de tots els blocs, junts, formen el conjunt de dades")

# comprovacions, per no moure el que no toca
assert el_foto22.findall('.//' + qn('w:drawing')), "no hi ha imatge a la Figura 22"
assert "Figura 22." in "".join(t.text or "" for t in el_peu22.iter(qn('w:t')))
assert el_foto23.findall('.//' + qn('w:drawing')), "no hi ha imatge a la Figura 23"
assert "Figura 23." in "".join(t.text or "" for t in el_peu23.iter(qn('w:t')))

# ---------------------------------------------------------------------------
# 1. El bloc de la resta de la parcel·la, al final del 4.6.2
# ---------------------------------------------------------------------------
h463 = [p for p in doc.paragraphs
        if p.style.name == 'Heading 3' and p.text.strip().startswith('4.6.3')][0]
dest = h463._element.getprevious()
for el in (p_blocs1._element, p_blocs2._element, el_foto23, el_peu23,
           p_rf._element):
    dest.addnext(el)
    dest = el

# ---------------------------------------------------------------------------
# 2. La capa de polígons sobre l'ortomosaic, després de l'aixecament GPS
# ---------------------------------------------------------------------------
peu20 = P("Figura 20. Aixecament de la coordenada GPS")
p = after(peu20._element)
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
p.paragraph_format.space_after = Pt(6)
_font(p.add_run(
    "La Figura 22 mostra aquesta capa sobreposada a l'ortomosaic del vol. Cada "
    "rectangle marca els 4,25 m² del marc de plantació que corresponen a un cep i "
    "és la unitat sobre la qual es calculen tots els estadístics dels quals surten "
    "els paràmetres estructurals."), 12)
dest = p._element
for el in (el_foto22, el_peu22):
    dest.addnext(el)
    dest = el

# ---------------------------------------------------------------------------
# 3. El que queda al pòrtic del 4.6: només com s'executa l'script
# ---------------------------------------------------------------------------
set_text(p_adaptex,
    "La zona de l'assaig ADAPTEX es processa d'una sola vegada. L'script rep el "
    "polígon que en delimita l'àrea d'interès i la capa que conté un element per "
    "cep, i recorre els 118 ceps en una única execució. Com es construeix aquesta "
    "capa i quins paràmetres se'n deriven es descriu a l'apartat 4.6.2.")

doc.save(PATH)
print(f"OK · paràgrafs: {len(doc.paragraphs)}")
