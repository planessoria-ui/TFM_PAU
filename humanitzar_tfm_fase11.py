# -*- coding: utf-8 -*-
"""
Fase 11: substitueix les figures redissenyades dins del document, hi afegeix
la figura del pla de vol i completa el tram final del protocol de processat
(màscares, quadrícula i paràmetres estructurals, índexs espectrals).
"""
import copy
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PATH = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v8.docx"
FIG = "/home/user/TFM_PAU/figures/"
doc = docx.Document(PATH)
FONT = "Arial"


def _font(run, size=12, bold=False, italic=False, red=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if red:
        run.font.color.rgb = RGBColor(0xB0, 0x00, 0x00)
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts'); rpr.append(rf)
    for a in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rf.set(qn(a), FONT)


def set_text(p, new):
    runs = p.runs
    if not runs:
        _font(p.add_run(new)); return
    runs[0].text = new
    for r in runs[1:]:
        r._element.getparent().remove(r._element)


def find(prefix, style=None):
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith(prefix) and (style is None or p.style.name == style):
            return i
    raise LookupError(prefix[:70])


def P(prefix, style=None):
    return doc.paragraphs[find(prefix, style)]


def new_par_after(anchor_el, parent, text, style=None, size=12, red=False,
                  align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    el = OxmlElement('w:p')
    anchor_el.addnext(el)
    p = docx.text.paragraph.Paragraph(el, parent)
    if style:
        p.style = doc.styles[style]
    p.alignment = align
    r = p.add_run(text)
    _font(r, size, bold=style and style.startswith('Heading'), italic=red, red=red)
    return p


# ---------------------------------------------------------------------------
# 1. Substitució de les figures redissenyades
# ---------------------------------------------------------------------------
repl = {"/word/media/image15.png": "fig3_cwsi.png",
        "/word/media/image16.png": "fig5_workflow.png"}
done = []
for rel in doc.part.rels.values():
    if 'image' in rel.reltype:
        name = str(rel.target_part.partname)
        if name in repl:
            rel.target_part._blob = open(FIG + repl[name], 'rb').read()
            done.append(repl[name])
print("figures substituïdes:", done)

# ---------------------------------------------------------------------------
# 2. Figura del pla de vol, just després de la Taula 7
# ---------------------------------------------------------------------------
cap7 = P("Taula 7. Paràmetres de planificació")
tbl7 = cap7._element.getnext()          # la taula va immediatament després del peu
parent = cap7._parent

p_img = new_par_after(tbl7, parent, "", align=WD_ALIGN_PARAGRAPH.CENTER)
p_img.add_run().add_picture(FIG + "fig8_pla_vol.png", width=Cm(11.5))

p_cap = new_par_after(p_img._element, parent, "", style="Caption",
                      align=WD_ALIGN_PARAGRAPH.LEFT)
for r in p_cap.runs:
    r._element.getparent().remove(r._element)
_font(p_cap.add_run("Figura 8. Planificació de la missió de vol sobre la parcel·la "
                    "experimental: (A) model de càmera, mode i interval de captura, "
                    "velocitat, alçada de vol i resolució a terra; (B) solapaments "
                    "frontal i lateral, angle de passada, orientació de la càmera i "
                    "acció en acabar la missió. "), 10)
_font(p_cap.add_run("Font: captures del programari de planificació de vol."), 10, italic=True)

# ---------------------------------------------------------------------------
# 3. Tram final del protocol
# ---------------------------------------------------------------------------
p_end = P("[A CONCRETAR] Mètode i índexs concrets")
anchor = p_end._element
parent = p_end._parent

blocs = [
    ("4.6.1. Generació de màscares", "Heading 3", None),
    ("Abans d'extreure les variables de cada cep cal separar, dins de cada capa, els "
     "píxels que corresponen al cultiu dels que no hi corresponen. Les màscares es "
     "generen amb un script en Python que aplica un llindar sobre la capa escollida; "
     "quan no se n'indica cap, el llindar es determina de manera automàtica pel mètode "
     "d'Otsu. Se'n fan servir quatre tipus.", None, None),
    ("La màscara d'alçada separa el dosser del sòl a partir del model digital de "
     "vegetació, amb un llindar fixat manualment segons la mida dels ceps. De la "
     "màscara multiespectral se'n generen dues: una per a la fracció de coberta, amb "
     "llindar automàtic i la quadrícula com a referència, i una altra més restrictiva "
     "per extreure els índexs de vegetació amb píxels purs; en aquest cas el llindar "
     "s'aplica sobre l'índex SAVI. La màscara tèrmica incorpora la d'alçada per "
     "descartar ombres i vegetació del carrer, i és la que permet obtenir la "
     "temperatura de dosser (Tc). Finalment, la màscara de sòl aïlla el sòl del carrer "
     "per obtenir la temperatura de sòl (Ts).", None, None),

    ("4.6.2. Quadrícula i paràmetres estructurals del dosser", "Heading 3", None),
    ("Amb les capes ja emmascarades es genera una quadrícula de polígons, un per cep, "
     "mitjançant un script que rep les coordenades d'origen, les dimensions del "
     "polígon, la distància entre polígons —que es correspon amb el marc de plantació "
     "d'1,7 × 2,5 m—, el nombre de files i de columnes i l'angle d'orientació de les "
     "fileres. Aquesta capa vectorial és la que permet assignar a cada cep el "
     "tractament i el bloc als quals pertany, i és la mateixa que s'utilitza al pas 4 "
     "del procediment descrit a l'apartat 4.8.", None, None),
    ("Creuant la quadrícula amb les capes d'alçada i tèrmica s'obtenen, per a cada cep, "
     "els paràmetres estructurals del dosser: l'alçada, l'àrea de copa, la fracció de "
     "coberta i el volum de copa. Aquests paràmetres, juntament amb els índexs de "
     "vegetació, són les variables que alimenten el model de Random Forest per estimar "
     "el LAI de tots els ceps de l'assaig.", None, None),

    ("4.6.3. Índexs espectrals", "Heading 3", None),
    ("A partir de les bandes calibrades es calculen els índexs de vegetació que "
     "descriuen l'estat del dosser. El NDVI relaciona les bandes de l'infraroig proper "
     "i del vermell i pren valors entre −1 i 1: els valors alts indiquen vegetació "
     "densa i activa, els propers a zero poca vegetació o sòl descobert, i els negatius "
     "absència de vegetació. El SAVI n'és una variant que incorpora un factor de "
     "correcció del sòl (L), habitualment de 0,5, i dona una estimació més fiable quan "
     "la coberta vegetal és escassa o el sòl queda a la vista, cosa que el fa "
     "especialment adequat en un cultiu en files com la vinya.", None, None),
    ("En el cas de l'Altum-PT, l'infraroig proper correspon a la banda 5 i el vermell a "
     "la banda 3, de manera que els dos índexs es calculen amb la calculadora ràster "
     "com NDVI = (B5 − B3) / (B5 + B3) i SAVI = [(B5 − B3) / (B5 + B3 + 0,5)] × 1,5.",
     None, None),
]
for text, style, _ in reversed(blocs):
    new_par_after(anchor, parent, text, style=style,
                  size=13 if style else 12,
                  align=WD_ALIGN_PARAGRAPH.LEFT if style else WD_ALIGN_PARAGRAPH.JUSTIFY)

doc.save(PATH)
print("Fase 11 completada.")
