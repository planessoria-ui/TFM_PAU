# -*- coding: utf-8 -*-
"""
Fase 17: incorpora els paràmetres reals del processament fotogramètric dels
quatre vols (Taula 15) i substitueix l'explicació provisional de l'anomalia de
l'alçada del dosser per la causa real.
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


def _font(run, size=12, bold=False, name=FONT, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
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


def after(el):
    new = OxmlElement('w:p')
    el.addnext(new)
    return docx.text.paragraph.Paragraph(new, _PARENT)


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
            cl = t.cell(r, c)
            cl.width = widths[c]
            pp = cl.paragraphs[0]
            pp.alignment = (WD_ALIGN_PARAGRAPH.LEFT if (c in left and r)
                            else WD_ALIGN_PARAGRAPH.CENTER)
            pp.paragraph_format.space_before = Pt(2)
            pp.paragraph_format.space_after = Pt(2)
            pp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            _font(pp.add_run(val), 9, bold=(r == 0))
    el.addnext(t._element)
    return t._element


# ===========================================================================
# Taula 15: paràmetres del processament fotogramètric
# ===========================================================================
anchor = para("totes les capes georreferenciades")
prev = anchor._element.getprevious()

el = body(prev,
    "Els quatre vols s'han processat seguint aquest mateix procediment, però no amb "
    "els mateixos paràmetres, i les diferències tenen conseqüències sobre les "
    "variables que se'n deriven. Els vols del 30 de juny, del 22 de juliol i del 12 "
    "d'agost es van tractar en dos blocs: un de general, amb núvol de punts dens de "
    "qualitat alta i filtratge moderat, del qual surt l'ortomosaic; i un de "
    "restringit a la zona de l'assaig, amb núvol dens de qualitat extra alta i "
    "filtratge agressiu, del qual surt el model digital d'elevacions que després "
    "alimenta el model digital de vegetació. El vol del 3 de juny, en canvi, només "
    "té el bloc general, de manera que tant l'ortomosaic com el model d'elevacions "
    "en provenen. La Taula 15 recull els valors de cada vol.")
el = body(el,
    "La graella de treball del processament posterior és la de l'ortomosaic, perquè "
    "totes les capes es remostregen sobre la banda de l'infraroig proper. D'aquí ve "
    "que la mida de píxel que consta als fitxers de sortida coincideixi, amb menys "
    "d'un 2 % de diferència, amb la resolució de l'ortomosaic de cada vol, i no amb "
    "la del model d'elevacions.")
el = caption(el, "Taula 15.",
             "Paràmetres del processament fotogramètric de cada vol a Agisoft "
             "Metashape. El bloc general cobreix tota la superfície volada i és el "
             "que genera l'ortomosaic; el bloc de l'assaig es restringeix a la zona "
             "experimental i es processa amb més qualitat per obtenir el model "
             "d'elevacions. El vol del 3 de juny només té bloc general.")
el = table(el, [
    ("Vol", "Bloc", "Imatges orientades", "Núvol dens (milions de punts)",
     "Qualitat i filtratge", "MDE (cm/píxel)", "Ortomosaic (cm/píxel)"),
    ("03/06/2026", "General", "1.946", "343,6", "Alta, moderat", "2,66", "1,33"),
    ("30/06/2026", "General", "1.986", "69,8", "Alta, moderat", "5,77", "2,89"),
    ("", "Assaig", "881", "125,5", "Extra alta, agressiu", "2,84", "—"),
    ("22/07/2026", "General", "1.994", "77,0", "Alta, moderat", "5,81", "2,90"),
    ("", "Assaig", "755", "120,5", "Extra alta, agressiu", "2,87", "—"),
    ("12/08/2026", "General", "1.844", "81,7", "Alta, moderat", "5,29", "2,65"),
    ("", "Assaig", "662", "122,1", "Extra alta, agressiu", "2,60", "—"),
], [Cm(2.4), Cm(1.8), Cm(2.3), Cm(2.6), Cm(2.8), Cm(2.0), Cm(2.3)])

# ===========================================================================
# Explicació real de l'anomalia de l'alçada, a 5.1
# ===========================================================================
p = para("L'evolució de l'alçada s'ha de llegir amb una precaució")
set_text(p, (
    "L'evolució de l'alçada s'ha de llegir amb precaució, perquè el vol del 3 de "
    "juny no es va processar com els altres tres (Taula 15). D'una banda, el seu "
    "ortomosaic té una mida de píxel d'1,33 cm i la dels altres volta els 2,9 cm, de "
    "manera que la graella sobre la qual es calculen els estadístics és el doble de "
    "fina. De l'altra, el seu model d'elevacions prové d'un núvol dens de qualitat "
    "alta amb filtratge moderat, mentre que el dels altres tres surt d'un núvol de "
    "qualitat extra alta amb filtratge agressiu, que elimina més punts aïllats. Com "
    "que l'alçada del dosser es pren com el valor màxim dins del polígon de cada "
    "cep, els dos factors empenyen en la mateixa direcció: una graella més fina "
    "mostreja més punts i un filtratge més suau en conserva més d'extrems, i totes "
    "dues coses apugen el màxim. Això explica per què el vol de juny apareix per "
    "damunt dels altres a tots els tractaments de la Figura 16B, i per què la "
    "davallada aparent entre el juny i el juliol no s'ha d'interpretar com una "
    "pèrdua real d'alçada del dosser. La fracció de coberta, que és un recompte "
    "relatiu, hi és molt menys sensible, i de fet la seva evolució sí que segueix un "
    "patró coherent amb l'estació."))

p = para("Convé decidir amb l'equip si es tornen a processar")
set_text(p, "", 11)
p.runs[0].text = "[A VERIFICAR] "
_font(p.runs[0], 11, bold=True, color=RED)
_font(p.add_run(
    "Per poder comparar l'alçada entre dates caldria reprocessar el vol del 3 de "
    "juny amb els mateixos paràmetres que els altres tres, o bé substituir el màxim "
    "per un percentil alt, per exemple el 95, que és molt menys sensible al "
    "filtratge i a la mida de píxel. Mentre no es faci, la comparació d'alçades "
    "entre tractaments dins d'una mateixa data és vàlida, però la comparació entre "
    "dates no ho és."), 11, color=RED)

# ===========================================================================
# Conflicte de data del tercer vol
# ===========================================================================
p = para("Els quatre vols processats són els del 3 i el 30 de juny")
_font(p.add_run(
    " Hi ha encara un tercer joc de dates: el projecte de Metashape del tercer vol "
    "està etiquetat com a 23 de juliol, mentre que tant el nom del fitxer de sortida "
    "com la seva columna de data diuen 22 de juliol. A la Taula 3 s'ha fet constar "
    "el 22, que és el que diu la cadena de processament, però cal confirmar quina "
    "és la data bona."), 11, color=RED)

# ===========================================================================
# Recàlcul dels camps SEQ
# ===========================================================================
counters = {"Figura": 0, "Taula": 0}
for p in doc.paragraphs:
    runs = p._element.findall(qn('w:r'))
    for i, r in enumerate(runs):
        instr = r.find(qn('w:instrText'))
        if instr is None or not instr.text or 'SEQ' not in instr.text:
            continue
        name = ('Figura' if 'Figura' in instr.text
                else 'Taula' if 'Taula' in instr.text else None)
        if name is None:
            continue
        counters[name] += 1
        seen_sep = False
        for r2 in runs[i + 1:]:
            fch = r2.find(qn('w:fldChar'))
            if fch is not None:
                ft = fch.get(qn('w:fldCharType'))
                if ft == 'separate':
                    seen_sep = True; continue
                if ft == 'end':
                    break
            if seen_sep:
                tnode = r2.find(qn('w:t'))
                if tnode is not None:
                    tnode.text = str(counters[name])

doc.save(PATH)
print(f"OK · paràgrafs: {len(doc.paragraphs)} · taules: {len(doc.tables)}")
