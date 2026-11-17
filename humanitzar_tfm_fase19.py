# -*- coding: utf-8 -*-
"""
Fase 19: defineix amb precisió els cinc tractaments de maneig de capçada a
partir de les operacions realment aplicades (desbrot, tall en altura i
desfullat) i les seves dates.
"""
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
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


# ===========================================================================
# 1. Prosa de 4.2: les tres operacions de maneig
# ===========================================================================
el = para("L'assaig segueix un disseny en blocs amb cinc tractaments")._element
el = body(el,
    "Els cinc tractaments es construeixen combinant tres operacions de maneig en "
    "verd, aplicades el mateix dia a tots els blocs. La primera és el desbrot, que "
    "es va fer el 2 de juny de 2026, el dia abans del primer vol. Consisteix a "
    "eliminar brots de cada pitó, o punt de càrrega, per limitar els que es deixen "
    "créixer: al tractament d'esporga lleugera se'n van deixar tres per pitó i als "
    "d'esporga mitjana, severa i severa amb desfullat, dos. El tractament sense "
    "esporga no es desbrota, de manera que conserva tots els brots i és el de "
    "capçada màxima.")
el = body(el,
    "La segona operació és el tall en altura, que és pròpiament l'esporga i que es "
    "va aplicar el mateix 2 de juny només als dos tractaments severs. Retalla la "
    "part superior de la paret vegetal i, combinat amb el desbrot a dos brots per "
    "pitó, és el que separa aquests dos tractaments de la resta pel que fa a la mida "
    "del dosser.")
el = body(el,
    "La tercera és el desfullat, que només rep el tractament més sever i que es va "
    "fer el 20 de juliol de 2026, coincidint amb el verolat. Es van treure dues o "
    "tres fulles de cada brot a la zona del raïm i només per la cara est de la "
    "filera, que és la que rep el sol amb menys intensitat; limitar el desfullat a "
    "aquest costat airejà i exposa el raïm sense sotmetre'l a la insolació de la "
    "tarda, que el cremaria. La Taula 1 resumeix quines operacions rep cada "
    "tractament.")

# ===========================================================================
# 2. Taula 1 refeta com a matriu d'operacions
# ===========================================================================
t1 = [t for t in doc.tables if t.cell(0, 0).text.strip() == "Codi"][0]
ROWS = [
    ("Codi", "Tractament", "Desbrot (02/06/2026)",
     "Tall en altura (02/06/2026)", "Desfullat (20/07/2026)"),
    ("SE", "Sense esporga", "No se'n fa; es conserven tots els brots", "No", "No"),
    ("EL", "Esporga lleugera", "3 brots per pitó", "No", "No"),
    ("EM", "Esporga mitjana (referència comercial)", "2 brots per pitó", "No", "No"),
    ("ES", "Esporga severa", "2 brots per pitó", "Sí", "No"),
    ("ES+DV", "Esporga severa amb desfullat a verolat", "2 brots per pitó", "Sí",
     "Sí: 2–3 fulles per brot a la zona del raïm, només a la cara est"),
]
while len(t1.columns) < 5:
    t1.add_column(Cm(3.0))
WID = [Cm(1.6), Cm(4.4), Cm(3.4), Cm(2.6), Cm(4.0)]
for r, data in enumerate(ROWS):
    for c, val in enumerate(data):
        cell = t1.cell(r, c)
        cell.width = WID[c]
        p = cell.paragraphs[0]
        for extra in cell.paragraphs[1:]:
            extra._element.getparent().remove(extra._element)
        for run in p.runs:
            run._element.getparent().remove(run._element)
        p.alignment = (WD_ALIGN_PARAGRAPH.LEFT if (c == 1 and r) else
                       WD_ALIGN_PARAGRAPH.CENTER)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        _font(p.add_run(val), 9, bold=(r == 0 or c == 0))

cap = para("Tractaments de maneig de capçada (esporga) de l'assaig ADAPTEX")
txt = cap.text.split(".", 1)[1]
runs = cap.runs
tail = None
for r in runs:
    if "Tractaments de maneig" in r.text:
        tail = r
        break
if tail is not None:
    tail.text = (". Tractaments de maneig de capçada de l'assaig ADAPTEX, definits "
                 "per les operacions en verd que rep cadascun i les dates en què es "
                 "van aplicar la campanya de 2026. El tractament EM correspon a la "
                 "pràctica comercial de referència. Font: elaboració pròpia.")
    for r in runs[runs.index(tail) + 1:]:
        r._element.getparent().remove(r._element)

# ===========================================================================
# 3. 4.4: la data del verolat
# ===========================================================================
p = para("Des del verolat, que el 2026 es va produir")
set_text(p, p.text.replace("que el 2026 es va produir cap al 22 de juliol",
                           "que el 2026 es va produir el 20 de juliol"))

# ===========================================================================
# 4. La nota de dates: la primera meitat ja està resolta
# ===========================================================================
p = para("Dues dates que convé precisar")
set_text(p, "", 11)
p.runs[0].text = "[A VERIFICAR] "
_font(p.runs[0], 11, bold=True, color=RED)
_font(p.add_run(
    "Queda per precisar la finestra de verema: l'apartat diu que comença a principis "
    "de setembre, mentre que l'esquema de maneig la situa entre la quarta setmana "
    "d'agost i la tercera de setembre."), 11, color=RED)

# ===========================================================================
# 5. 5.1: l'ordre mesurat reprodueix el de les operacions
# ===========================================================================
p = para("L'evolució estacional també és diferent segons el tractament")
set_text(p, p.text.rstrip() + (
    " L'ordre que marquen aquestes fraccions reprodueix el de les operacions de "
    "maneig de la Taula 1: el tractament que no es desbrota té la capçada més gran, "
    "el desbrot a tres brots per pitó en redueix la coberta, el desbrot a dos la "
    "redueix més, i el tall en altura dels dos tractaments severs la deixa per sota "
    "de 0,20."))

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
