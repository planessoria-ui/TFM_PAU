# -*- coding: utf-8 -*-
"""
Fase 15: incorpora les dades de la capa de marc de plantació, de la màscara de
l'assaig, del fitxer info_vines i de la sortida real del processament.

  (a) 4.1: georeferenciació real de la zona d'assaig.
  (b) 4.2: els 118 ceps monitoritzats i el polígon de 4,25 m2.
  (c) 4.6.2: descripció real de la capa de polígons i mida de píxel del vol.
  (d) 4.6.4: Taula 13 amb l'estructura del fitxer de sortida.
  (e) Refà la nota de verificació del calibratge tèrmic amb la comparació
      correcta entre la Tc del dron i les mesures del termòmetre d'infraroig.
  (f) Completa la nota de la màscara de l'entrefilera amb els valors reals.
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
RED = RGBColor(0xC0, 0x00, 0x00)
_PARENT = doc.paragraphs[0]._parent


def _font(run, size=12, bold=False, italic=False, name=FONT, color=None):
    run.font.name = name
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
        rf.set(qn(a), name)


def set_text(p, new, size=12):
    runs = p.runs
    if not runs:
        _font(p.add_run(new), size); return
    runs[0].text = new
    for r in runs[1:]:
        r._element.getparent().remove(r._element)


def para(sub):
    for p in doc.paragraphs:
        if sub in p.text:
            return p
    raise LookupError(sub[:70])


def after(el, style=None):
    new = OxmlElement('w:p')
    el.addnext(new)
    p = docx.text.paragraph.Paragraph(new, _PARENT)
    if style:
        p.style = doc.styles[style]
    return p


def body(el, text, size=12):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    p.paragraph_format.space_after = Pt(6)
    _font(p.add_run(text), size)
    return p._element


def caption(el, head, text, size=10):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    _font(p.add_run(head), size, bold=True)
    _font(p.add_run(" " + text), size)
    return p._element


def table(el, rows, widths, left=()):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = doc.styles['Table Grid']
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r, data in enumerate(rows):
        for c, val in enumerate(data):
            cell = t.cell(r, c)
            cell.width = widths[c]
            pp = cell.paragraphs[0]
            pp.alignment = (WD_ALIGN_PARAGRAPH.LEFT if (c in left and r)
                            else WD_ALIGN_PARAGRAPH.CENTER)
            pp.paragraph_format.space_before = Pt(2)
            pp.paragraph_format.space_after = Pt(2)
            pp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            _font(pp.add_run(val), 9.5, bold=(r == 0))
    el.addnext(t._element)
    return t._element


# ===========================================================================
# (a) 4.1 Georeferenciació de la zona d'assaig
# ===========================================================================
p = para("La vinya, plantada l'any 2013, es condueix en doble cordó")
body(p._element,
     "La zona on es concentra l'assaig ocupa 2.318 m² dins d'aquesta parcel·la, en "
     "un polígon de 55 × 42 m amb les fileres orientades 33,5° respecte del nord. "
     "El seu centre es troba a 41,6503° N i 0,5193° E, a 377 m d'altitud, i totes "
     "les capes es treballen en coordenades UTM del fus 31N. La posició de cada cep "
     "de l'assaig es va aixecar amb GPS amb una precisió horitzontal declarada d'un "
     "centímetre, de manera que les mesures de camp i les imatges de dron queden "
     "referenciades sobre la mateixa base geomètrica.")

p = para("Coordenades exactes de la parcel·la, tipus de sòl")
for r in p.runs:
    r.text = r.text.replace("Coordenades exactes de la parcel·la, tipus de sòl",
                            "Tipus de sòl")

# ===========================================================================
# (b) 4.2 Nombre de ceps monitoritzats
# ===========================================================================
p = para("L'assaig segueix un disseny en blocs amb cinc tractaments")
set_text(p, p.text.rstrip() + (
    " El seguiment amb dron cobreix, doncs, 118 ceps individualitzats: 24 per "
    "tractament, llevat de l'esporga severa i del tractament sense esporga, que en "
    "tenen 23. A cada cep se li assigna un polígon de 4,25 m², que és la superfície "
    "de sòl que li correspon pel marc de plantació."))

# ===========================================================================
# (c) 4.6.2 Capa de polígons real
# ===========================================================================
p = para("Amb les capes ja emmascarades es genera una quadrícula de polígons")
set_text(p, (
    "Amb les capes ja emmascarades, els estadístics es calculen sobre una capa "
    "vectorial que conté un polígon per cep. Cada polígon cobreix 4,25 m², que és "
    "la superfície de sòl que correspon a un cep segons el marc de plantació d'1,7 "
    "× 2,5 m, i s'orienta seguint la direcció de les fileres. La capa es construeix "
    "a partir de la posició de cada cep aixecada amb GPS de precisió centimètrica, "
    "i cada element porta com a atributs l'identificador del cep, el bloc i el "
    "tractament, de manera que l'assignació experimental viatja amb la geometria i "
    "no cal creuar-la en un pas posterior. Amb la mida de píxel d'aquest vol, d'1,34 "
    "cm, cada polígon conté unes 23.800 cel·les. És la mateixa capa que s'utilitza "
    "al pas 4 del procediment descrit a l'apartat 4.8."))

# ===========================================================================
# (d) 4.6.4 Taula 13 amb l'estructura del fitxer de sortida
# ===========================================================================
p = para("El resultat del processament de cada bloc és un fitxer de text")
set_text(p, (
    "El resultat del processament de cada bloc és un fitxer de text delimitat amb "
    "una fila per cep i per data i quaranta-tres columnes, que la Taula 13 agrupa "
    "per famílies. Hi ha la identificació del cep amb el seu tractament, el bloc i "
    "el marc de plantació; els estadístics dels cinc índexs espectrals, amb mitjana "
    "i desviació estàndard, tant per als píxels de dosser com per als de "
    "l'entrefilera; les tres temperatures; i els paràmetres estructurals de la "
    "Taula 11. Concatenats, aquests fitxers formen la taula de treball sobre la qual "
    "s'ajusten el model de Random Forest per al LAI i el model TSEB-2T. De cada bloc "
    "se'n desen també les capes ràster intermèdies —temperatura de superfície i de "
    "sòl, NDVI, alçada i les màscares de dosser i d'entrefilera—, que permeten "
    "comprovar visualment que la classificació ha funcionat abans de donar les dades "
    "per bones."))
el = caption(p._element, "Taula 13.",
             "Estructura del fitxer de sortida del processament, amb una fila per "
             "cep i per data. El fitxer corresponent al vol del 3 de juny de 2026 "
             "conté 118 files, una per cada cep de l'assaig.")
el = table(el, [
    ("Grup de camps", "Contingut", "Nre."),
    ("Identificació",
     "Identificador i nom del cep, tractament, marc de plantació i data del vol", "5"),
    ("Recompte de píxels",
     "Píxels classificats com a dosser i com a entrefilera dins del polígon", "2"),
    ("Índexs del dosser",
     "Mitjana i desviació estàndard de NDVI, SAVI, MTVI, NDWI i CIG", "10"),
    ("Índexs de l'entrefilera",
     "Mitjana i desviació estàndard dels mateixos cinc índexs", "10"),
    ("Temperatures",
     "Temperatura de superfície, de dosser i de sòl", "3"),
    ("Geometria del dosser",
     "Alçada, superfície del píxel, àrea i volum de copa, fracció de coberta, "
     "volum per unitat de superfície, radi, amplada i relació amplada/alçada", "9"),
    ("Entrefilera",
     "Àrea i fracció de coberta de l'entrefilera", "2"),
], [Cm(4.2), Cm(9.4), Cm(1.8)], left=(1,))

# ===========================================================================
# (e) Nota de verificació del calibratge tèrmic, refeta
# ===========================================================================
p = para("Cap de les quatre rectes de la Taula 9 es reprodueix")
set_text(p, "", 11)
p.runs[0].text = "[A VERIFICAR] "
_font(p.runs[0], 11, bold=True, color=RED)
_font(p.add_run(
    "La calibració de la banda tèrmica presenta dos problemes que convé resoldre "
    "abans d'utilitzar cap temperatura al balanç energètic. El primer és d'ajust: "
    "cap de les quatre rectes de la Taula 9 es reprodueix ajustant per mínims "
    "quadrats els cinc elements de referència del mateix full, que donen pendents "
    "d'entre 0,0164 i 0,0183 enfront dels 0,0130–0,0132 registrats. El segon, i més "
    "seriós, és que la temperatura de dosser que surt del processament no concorda "
    "amb la que es va mesurar a camp amb el termòmetre d'infraroig. Comparant plot a "
    "plot la Tc mitjana dels sis ceps amb la mesura de l'IRT del full TEMP_v5, el "
    "dron dona de mitjana 3,5 °C més, amb un error quadràtic mitjà de 4,2 °C. I la "
    "diferència no és constant: creix a mesura que disminueix la fracció de coberta "
    "del tractament, des de +0,7 °C de mitjana al tractament sense esporga, que és "
    "el de capçada més gran, fins a +5,7 °C a l'esporga severa amb desfullat, que és "
    "el de capçada més petita; la correlació entre la fracció de coberta del "
    "tractament i el biaix és de −0,94. El patró és compatible amb una contaminació "
    "dels píxels de dosser per sòl calent a les capçades més obertes, que el retall "
    "al percentil 25 no acaba d'eliminar. Com que el biaix va en la mateixa direcció "
    "que l'efecte dels tractaments que es vol mesurar, podria generar diferències "
    "d'ETa entre tractaments que no fossin reals. Cal comprovar amb l'equip l'hora de "
    "les mesures de l'IRT respecte de la del vol, si convé endurir el llindar de la "
    "màscara de dosser o abaixar el percentil, i revisar la recta de calibratge."),
    11, color=RED)

# ===========================================================================
# (f) Nota de la màscara de l'entrefilera, completada
# ===========================================================================
p = para("la màscara de l'entrefilera s'anomena SOIL_GREEN")
set_text(p, "", 11)
p.runs[0].text = "[A VERIFICAR] "
_font(p.runs[0], 11, bold=True, color=RED)
_font(p.add_run(
    "A l'script, la màscara de l'entrefilera s'anomena SOIL_GREEN i les variables "
    "que se'n deriven porten l'etiqueta covercrop, però el llindar que l'estableix "
    "conserva els píxels amb un SAVI inferior a 0,1, que són precisament els de sòl "
    "sense vegetació activa. La sortida del vol del 3 de juny ho confirma: la "
    "fracció de coberta de l'entrefilera que en resulta val 0,0003 de mitjana, 37 "
    "dels 118 ceps no tenen cap píxel en aquesta classe i 91 en tenen menys de deu, "
    "sobre les prop de 23.800 cel·les que conté cada polígon. Tal com està "
    "calculada, doncs, la variable no mesura la coberta del carrer, perquè al juny "
    "l'entrefilera era verda i el llindar selecciona el sòl nu. Cal decidir si es "
    "gira la comparació, i aleshores la variable passa a ser la coberta vegetal del "
    "carrer, o si es manté el llindar i es canvia el nom perquè descrigui el que "
    "realment selecciona."), 11, color=RED)

doc.save(PATH)
print(f"OK · paràgrafs: {len(doc.paragraphs)} · taules: {len(doc.tables)}")
