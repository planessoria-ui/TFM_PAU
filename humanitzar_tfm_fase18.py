# -*- coding: utf-8 -*-
"""
Fase 18: incorpora el calendari de maneig de la campanya 2026 (Taula 16 i
Figura 17) i la comparació entre els coeficients Kr i la fracció de coberta
mesurada.
"""
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PATH = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v8.docx"
FIG = "/home/user/TFM_PAU/figures/"
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


def note(el, text, tag="[A VERIFICAR] "):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(6)
    _font(p.add_run(tag), 11, bold=True, color=RED)
    _font(p.add_run(text), 11, color=RED)
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
            _font(pp.add_run(val), 9.5, bold=(r == 0))
    el.addnext(t._element)
    return t._element


def picture(el, name, width_cm):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(FIG + name, width=Cm(width_cm))
    return p._element


# ===========================================================================
# 4.4: Kc, Kr i Taula 16
# ===========================================================================
el = para("La dosi de reg es calcula a partir de l'evapotranspiraci")._element
el = body(el,
    "En aquesta campanya el coeficient de cultiu s'ha mantingut constant en 0,45 des "
    "de la primera setmana de maig fins a l'última de setembre. A més, la dosi de "
    "cada tractament es modula amb un coeficient de reducció per capçada (Kr), que "
    "recull la diferència de mida del dosser que provoca l'esporga, de manera que la "
    "dosi teòrica de cada tractament és ETc = ET0 × Kc × Kr. Fins a finals de maig "
    "tots els tractaments comparteixen un Kr d'1,0 i, a partir de l'aplicació dels "
    "tractaments al juny, cadascun pren el seu valor. El tractament d'esporga "
    "mitjana, que és la pràctica comercial de referència, manté el Kr a 1,0; els dos "
    "tractaments severs baixen a 0,7 i els de capçada més gran pugen a 1,1 i 1,4 "
    "(Taula 16).")
el = caption(el, "Taula 16.",
             "Coeficient de cultiu, coeficient de reducció per capçada i fracció de "
             "reg aplicada a cada tractament durant la campanya de 2026. La dosi "
             "teòrica de cada tractament és ETc = ET0 × Kc × Kr; el 25 % correspon "
             "al reg deficitari controlat que s'aplica des del verolat fins a la "
             "verema.")
el = table(el, [
    ("Tractament", "Kc", "Kr (maig)", "Kr (juny–setembre)",
     "Reg fins al verolat", "Reg des del verolat"),
    ("Sense esporga (C)", "0,45", "1,0", "1,4", "100 %", "25 %"),
    ("Esporga lleugera (ELL)", "0,45", "1,0", "1,1", "100 %", "25 %"),
    ("Esporga mitjana (EM)", "0,45", "1,0", "1,0", "100 %", "25 %"),
    ("Esporga severa (ES)", "0,45", "1,0", "0,7", "100 %", "25 %"),
    ("Esporga severa + desfullat (ES+DV)", "0,45", "1,0", "0,7", "100 %", "25 %"),
], [Cm(5.4), Cm(1.6), Cm(2.2), Cm(3.0), Cm(2.6), Cm(2.6)], left=(0,))
el = note(el,
    "A l'esquema de maneig de l'assaig els cinc tractaments apareixen identificats "
    "com a T1 a T5 i no pel nom. L'assignació de la Taula 16 s'ha deduït dels valors "
    "de Kr, que ordenen els tractaments de menys a més capçada, i és coherent amb "
    "les fraccions de coberta mesurades amb dron (apartat 5.1), però convé "
    "confirmar-la amb l'equip abans de donar-la per bona.")

# ===========================================================================
# 4.4: Figura 17, calendari de maneig
# ===========================================================================
el = para("El seguiment del reg i del consum es fa amb la plataforma Treetoscope")._element
el = body(el,
    "La Figura 17 situa en el calendari de la campanya les fases fenològiques, les "
    "operacions de maneig, els quatre vols de dron i el règim de reg. L'esporga es "
    "va aplicar la primera setmana de juny i el desfullat del tractament ES+DV "
    "coincideix amb el verolat, a la tercera setmana de juliol, que és també el "
    "moment en què comença el reg deficitari. Els vols queden repartits al llarg "
    "d'aquest període: el primer just després de l'esporga, el segon i el tercer "
    "flanquejant el verolat i el quart ja dins del reg deficitari, poques setmanes "
    "abans de l'inici de la verema.")
el = picture(el, "fig17_calendari.png", 16.0)
el = caption(el, "Figura 17.",
             "Calendari de maneig de la campanya de 2026: fases fenològiques, "
             "operacions de maneig, dates dels quatre vols de dron i fracció de reg "
             "aplicada, per setmanes de març a octubre. Font: elaboració pròpia a "
             "partir de l'esquema de maneig de l'assaig.")
el = note(el,
    "Dues dates que convé precisar. La primera, si el vol del 3 de juny es va fer "
    "abans o després de l'esporga, perquè totes dues coses cauen a la primera "
    "setmana de juny i d'això depèn que el primer vol reculli ja la capçada de la "
    "campanya o encara la de l'any anterior. La segona, la finestra de verema: "
    "l'apartat diu que comença a principis de setembre, mentre que l'esquema de "
    "maneig la situa entre la quarta setmana d'agost i la tercera de setembre.")

# ===========================================================================
# 4.4: nota de dosi, actualitzada
# ===========================================================================
p = para("Indiqueu la dosi de reg aplicada abans del verolat")
set_text(p, "", 11)
p.runs[0].text = "[A COMPLETAR] "
_font(p.runs[0], 11, bold=True, color=RED)
_font(p.add_run(
    "La fracció de reg de cada fase ja consta a la Taula 16, però falta el volum "
    "real: els mil·límetres aplicats per tractament abans del verolat i el total de "
    "la campanya. Comproveu també quins blocs es monitoren a Treetoscope, perquè a "
    "la captura de la Figura 8 només n'hi apareixen tres i l'assaig en té cinc."),
    11, color=RED)

# ===========================================================================
# 5.1: els Kr contrastats amb la fracció de coberta mesurada
# ===========================================================================
el = para("La relació entre amplada i alçada, en canvi, es manté entre")._element
el = body(el,
    "Val la pena contrastar aquestes fraccions de coberta amb els coeficients de "
    "reducció per capçada que s'utilitzen per calcular la dosi de reg de cada "
    "tractament (Taula 16), perquè es van establir a partir de campanyes anteriors i "
    "no a partir d'aquestes imatges. Prenent l'esporga mitjana com a referència, les "
    "fraccions de coberta mitjanes dels quatre vols donen una relació d'1,50 al "
    "tractament sense esporga, d'1,31 a l'esporga lleugera, de 0,73 a l'esporga "
    "severa i de 0,70 a l'esporga severa amb desfullat, mentre que els Kr aplicats "
    "són 1,4, 1,1, 0,7 i 0,7. Les dues sèries es corresponen gairebé punt per punt "
    "(r² = 0,96), de manera que la mesura directa de la capçada amb dron confirma "
    "els coeficients amb què es modula el reg.")

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
