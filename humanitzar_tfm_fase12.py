# -*- coding: utf-8 -*-
"""
Fase 12: insereix a l'apartat 4.2 la fotografia dels cinc tractaments de
maneig de capçada, entre la taula que els descriu i el croquis del disseny,
i recalcula la numeració de figures i taules.
"""
import copy
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PATH = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v8.docx"
FIG = "/home/user/TFM_PAU/figures/"
doc = docx.Document(PATH)
FONT = "Arial"


def _font(run, size=12, bold=False, italic=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
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


# ---------------------------------------------------------------------------
# 1. Referència a la nova figura dins del text de 4.2
# ---------------------------------------------------------------------------
p_txt = doc.paragraphs[find("L'assaig segueix un disseny en blocs")]
t = p_txt.text
old = "Els cinc tractaments es descriuen a la Taula 1 i la seva distribució a la parcel·la es representa a la Figura 5."
new = ("Els cinc tractaments es descriuen a la Taula 1, la Figura 5 en mostra "
       "l'aspecte al camp i la Figura 6 representa la seva distribució a la parcel·la.")
if old in t:
    set_text(p_txt, t.replace(old, new))
    print("referència actualitzada")
else:
    print("AVÍS: no s'ha trobat la frase exacta; text actual:", t[-120:])

# ---------------------------------------------------------------------------
# 2. Inserció de la figura després de la taula de tractaments
# ---------------------------------------------------------------------------
cap_t1 = doc.paragraphs[find("Taula 1. Tractaments de maneig")]
anchor = cap_t1._element.getnext()            # la taula
while anchor is not None and anchor.tag != qn('w:tbl'):
    anchor = anchor.getnext()
if anchor is None:
    raise SystemExit("no s'ha trobat la taula de tractaments")

parent = cap_t1._parent

el_img = OxmlElement('w:p')
anchor.addnext(el_img)
p_img = docx.text.paragraph.Paragraph(el_img, parent)
p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_img.paragraph_format.space_before = Pt(10)
p_img.paragraph_format.space_after = Pt(2)
p_img.add_run().add_picture(FIG + "fig9_tractaments.png", width=Cm(15.5))

el_cap = OxmlElement('w:p')
el_img.addnext(el_cap)
p_cap = docx.text.paragraph.Paragraph(el_cap, parent)
p_cap.style = doc.styles['Caption']
p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT

_font(p_cap.add_run("Figura "), 10, bold=True)
rb = p_cap.add_run(); _font(rb, 10, bold=True)
fc = OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'), 'begin'); rb._element.append(fc)
ri = p_cap.add_run(); _font(ri, 10, bold=True)
it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve')
it.text = ' SEQ Figura \\* ARABIC '; ri._element.append(it)
rs = p_cap.add_run(); _font(rs, 10, bold=True)
fs = OxmlElement('w:fldChar'); fs.set(qn('w:fldCharType'), 'separate'); rs._element.append(fs)
rv = p_cap.add_run("5"); _font(rv, 10, bold=True)
re_ = p_cap.add_run(); _font(re_, 10, bold=True)
fe = OxmlElement('w:fldChar'); fe.set(qn('w:fldCharType'), 'end'); re_._element.append(fe)
_font(p_cap.add_run(
    ". Aspecte del dosser de cada tractament de maneig de capçada a la parcel·la "
    "experimental, d'esquerra a dreta: sense esporga (SE), esporga lleugera (EL), "
    "esporga mitjana (EM), esporga severa (ES) i esporga severa amb desfullat a "
    "verolat (ES+DV). "), 10)
_font(p_cap.add_run("Font: Blanco et al. (2025)."), 10, italic=True)

# ---------------------------------------------------------------------------
# 3. Recàlcul de la numeració desada dels camps SEQ
# ---------------------------------------------------------------------------
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
print(f"Figures numerades: {counters['Figura']} | Taules: {counters['Taula']}")
