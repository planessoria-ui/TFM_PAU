# -*- coding: utf-8 -*-
"""
Fase 27: com s'estima realment el LAI.

Les dues inversions de la fIPAR subestimaven el LAI, de manera que s'ajusta una
relació empírica entre el LAI mesurat destructivament i la fIPAR diària i és
aquesta equació la que s'aplica als ceps de l'assaig. S'obre l'apartat 4.7.3,
s'hi incorpora la Figura 24 i s'ajusten el pas 3 del 4.8 i la nota del 4.6.
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


def _font(run, size=12, bold=False, italic=False, color=None):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
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


def after(el, style=None):
    new = OxmlElement('w:p')
    el.addnext(new)
    p = docx.text.paragraph.Paragraph(new, _PARENT)
    if style:
        p.style = doc.styles[style]
    return p


def body(el, text):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    p.paragraph_format.space_after = Pt(6)
    _font(p.add_run(text), 12)
    return p._element


# ===========================================================================
# 1. Nou apartat 4.7.3, just abans del 4.8
# ===========================================================================
h48 = [p for p in doc.paragraphs
       if p.style.name == 'Heading 2' and p.text.strip().startswith('4.8.')][0]
el = h48._element.getprevious()

h = after(el, 'Heading 3')
_font(h.add_run("4.7.3. Estimació del LAI a partir de la fIPAR diària"), 13, bold=True)
el = h._element

el = body(el,
    "El ceptòmetre no mesura el LAI: el dedueix de la fracció de radiació que el "
    "dosser intercepta. Aquesta inversió es pot fer de dues maneres, amb el model "
    "que incorpora el mateix LP-80 i amb la formulació de Campbell i Norman, i totes "
    "dues es van provar amb els ceps defoliats, dels quals es coneix el LAI real. El "
    "resultat va ser el mateix en els dos casos: els valors invertits quedaven per "
    "sota dels mesurats. La subestimació té una explicació coneguda en vinya en "
    "espatllera, perquè les dues formulacions tracten el dosser com una capa "
    "horitzontal contínua i aquí la vegetació es concentra en una paret estreta, de "
    "manera que una part de la radiació que arriba al sensor no ha travessat fulles "
    "sinó l'entrefilera.")

el = body(el,
    "Davant d'això es va deixar de banda la inversió i es va ajustar una relació "
    "empírica directa. S'aparella el LAI de cada cep defoliat, mesurat fulla a fulla "
    "amb el procediment de l'apartat 4.7.2, amb la seva fIPAR diària, i s'hi ajusta "
    "la corba que millor descriu el núvol de punts, que és de tipus logarítmic "
    "(Figura 24):")

p = after(el)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
p.paragraph_format.space_before = Pt(6)
p.paragraph_format.space_after = Pt(6)
_font(p.add_run("LAI = 1,2615 · ln(fIPAR) + 2,8671            R² = 0,6943"), 12)
el = p._element

el = body(el,
    "L'equació explica prop del 70 % de la variació del LAI mesurat, un ajust "
    "raonable tenint en compte que es treballa amb un nombre reduït de ceps i que la "
    "mesura destructiva i la de radiació no es prenen exactament al mateix instant. "
    "Substituint a la x la fIPAR diària de cadascun dels 118 ceps de l'assaig "
    "ADAPTEX se n'obté el LAI, i aquests valors sí que resulten coherents amb "
    "l'amplada, l'alçada i la fracció de coberta que les imatges del dron donen per "
    "als mateixos ceps. Són també els valors amb què es calibra el model de Random "
    "Forest de l'apartat 4.8, que estén l'estimació als ceps de la resta de la "
    "parcel·la, on no hi ha mesures de fIPAR.")

el = body(el,
    "La relació s'ha ajustat amb fIPAR diàries d'entre 0,30 i 0,74, que és el rang "
    "que cobreixen els ceps defoliats, i no s'aplica fora d'aquest interval. La forma "
    "logarítmica hi té sentit: en capçades ja atapeïdes, afegir fulles augmenta cada "
    "cop menys la radiació interceptada, de manera que la corba s'aplana.")

pic = after(el)
pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
pic.paragraph_format.space_before = Pt(10)
pic.paragraph_format.space_after = Pt(2)
pic.add_run().add_picture(FIG + "fig24_lai_fipar.png", width=Cm(14.5))
el = pic._element

cap = after(el)
cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
cap.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
cap.paragraph_format.space_before = Pt(6)
cap.paragraph_format.space_after = Pt(6)
_font(cap.add_run("Figura 24."), 10, bold=True)
_font(cap.add_run(
    " Relació entre el LAI mesurat destructivament en els ceps defoliats i la seva "
    "fIPAR diària, amb la corba logarítmica ajustada i el seu coeficient de "
    "determinació. És l'equació que s'aplica a la fIPAR diària dels ceps de l'assaig "
    "per obtenir-ne el LAI. Font: elaboració pròpia."), 10)

# ===========================================================================
# 2. El pas 3 del procediment del 4.8
# ===========================================================================
set_text(P("Obtenir el LAI de la resta de ceps mesurats i corregir-lo"),
    "Com que totes dues inversions subestimen el LAI, estimar-lo a partir de la "
    "relació empírica entre el LAI mesurat destructivament i la fIPAR diària, "
    "descrita a l'apartat 4.7.3.")

# ===========================================================================
# 3. La nota del 4.6, ara reduïda al que encara queda obert
# ===========================================================================
p = P("[A CONCRETAR] Mètode i índexs concrets per estimar el LAI")
for r in p.runs[1:]:
    r._element.getparent().remove(r._element)
set_text(p,
    "El LAI no surt directament de cap índex espectral. Per als ceps de l'assaig es "
    "deriva de la fIPAR diària amb la relació empírica de l'apartat 4.7.3, i per als "
    "de la resta de la parcel·la, del model de Random Forest de l'apartat 4.8, que "
    "s'alimenta dels paràmetres estructurals i dels índexs espectrals descrits tot "
    "seguit.")
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
p.paragraph_format.space_after = Pt(6)

nota = after(p._element)
nota.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
nota.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
nota.paragraph_format.space_after = Pt(6)
_font(nota.add_run(
    "[A CONCRETAR] Com es deriva la fIPAR a partir de les imatges del dron, que és "
    "la que els apartats 4.7.1 i 5.2 preveuen validar contra el ceptòmetre."),
    12, color=RGBColor(0xC0, 0x00, 0x00))

doc.save(PATH)
print(f"OK · paràgrafs: {len(doc.paragraphs)}")
