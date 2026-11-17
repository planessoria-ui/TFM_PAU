# -*- coding: utf-8 -*-
"""
Versió editable del calendari de maneig en una diapositiva de PowerPoint.
Tot són formes i quadres de text natius, de manera que es poden moure, canviar
de color o reescriure sense tornar a generar res.
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "calendari_maneig_2026.pptx")

MESOS = ["MARÇ", "ABRIL", "MAIG", "JUNY", "JULIOL", "AGOST", "SETEMBRE", "OCTUBRE"]
NSET = 32
BROTACIO, DESBROT, VEROLAT, DESFULLAT = 3, 12, 18, 18
VEREMA = (23, 26)
VOLS = [("03/06", 12), ("30/06", 15), ("22/07", 19), ("12/08", 21)]
KC_SPAN = (8, 27)
REG = [((8, 18), "100 %"), ((19, 26), "25 %"), ((27, 27), "100 %")]
TRACT = [("SE", "Sense esporga", "1,4"), ("EL", "Esporga lleugera", "1,1"),
         ("EM", "Esporga mitjana", "1,0"), ("ES", "Esporga severa", "0,7"),
         ("ES+DV", "Esporga severa + desfullat", "0,7")]

INK = RGBColor(0x1A, 0x22, 0x30)
MUTED = RGBColor(0x52, 0x51, 0x4E)
RULE = RGBColor(0xD4, 0xD7, 0xDC)
FILL = RGBColor(0xF4, 0xF6, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_FEN = RGBColor(0x1B, 0xAF, 0x7A)
C_TAS = RGBColor(0x62, 0x50, 0xD6)
C_KC = RGBColor(0x3D, 0x44, 0x51)
HEADA = RGBColor(0xE9, 0xED, 0xF1)
HEADB = RGBColor(0xF5, 0xF7, 0xF9)
KR = {"0,7": (RGBColor(0xE7, 0xA3, 0x84), INK), "1,0": (RGBColor(0xD9, 0x82, 0x5A), INK),
      "1,1": (RGBColor(0xC5, 0x64, 0x37), WHITE), "1,4": (RGBColor(0x9E, 0x44, 0x1B), WHITE)}
RG = {"25 %": (RGBColor(0x86, 0xB6, 0xEF), INK), "100 %": (RGBColor(0x1C, 0x5C, 0xAB), WHITE)}

LX, RX = 3.62, 12.92
W = (RX - LX) / NSET

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])


def text(x, y, w, h, s, size=9, bold=False, color=INK, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = s
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = "Calibri"
    return tb


def box(x, y, w, h, fill, line=None, rounded=False):
    sh = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(0.75)
    sh.shadow.inherit = False
    if rounded:
        sh.adjustments[0] = 0.18
    return sh


def chip(a, b, y, h, fill, label="", tc=WHITE, size=9):
    sh = box(LX + a * W, y, (b - a + 1) * W, h, fill, line=WHITE, rounded=True)
    tf = sh.text_frame
    tf.margin_left = tf.margin_right = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.color.rgb = tc
    r.font.name = "Calibri"
    return sh


def band(y, h):
    box(LX, y, RX - LX, h, FILL, line=RULE)


text(0.42, 0.28, 9.0, 0.34, "Calendari de maneig de la vinya · campanya 2026",
     size=18, bold=True)
text(0.42, 0.63, 11.0, 0.24,
     "Assaig ADAPTEX de maneig de capçada · parcel·la El polvorín, Raimat",
     size=10.5, color=MUTED)

# capçalera
for m, mes in enumerate(MESOS):
    box(LX + m * 4 * W, 0.98, 4 * W, 0.26, HEADA if m % 2 == 0 else HEADB, line=RULE)
    text(LX + m * 4 * W, 0.98, 4 * W, 0.26, mes, size=9, bold=True,
         align=PP_ALIGN.CENTER)
    for s in range(4):
        text(LX + (m * 4 + s) * W, 1.26, W, 0.18, str(s + 1), size=7,
             color=MUTED, align=PP_ALIGN.CENTER)

RH, KH = 0.40, 0.30
y = 1.50
band(y, RH)
text(0.42, y, 3.05, RH, "Fenologia", size=11, bold=True, align=PP_ALIGN.RIGHT)
for idx, lab in ((BROTACIO, "Brotació"), (VEROLAT, "Verolat")):
    chip(idx, idx, y + 0.04, RH - 0.08, C_FEN)
    text(LX + (idx + 1.25) * W, y, 2.0, RH, lab, size=9.5)

y += 0.48
band(y, RH)
text(0.42, y, 3.05, RH, "Operacions", size=11, bold=True, align=PP_ALIGN.RIGHT)
chip(DESBROT, DESBROT, y + 0.04, RH - 0.08, C_TAS)
text(LX + (DESBROT - 7.6) * W, y, 7.3 * W, RH, "Desbrot i tall en altura",
     size=9.5, align=PP_ALIGN.RIGHT)
chip(DESFULLAT, DESFULLAT, y + 0.04, RH - 0.08, C_TAS)
text(LX + (DESFULLAT + 1.25) * W, y, 2.0, RH, "Desfullat", size=9.5)
chip(VEREMA[0], VEREMA[1], y + 0.04, RH - 0.08, C_TAS, "Verema", WHITE, 9.5)

y += 0.48
band(y, RH)
text(0.42, y, 3.05, RH, "Vols de dron", size=11, bold=True, align=PP_ALIGN.RIGHT)
for n, (dat, idx) in enumerate(VOLS, start=1):
    chip(idx, idx, y + 0.04, RH - 0.08, INK, str(n), WHITE, 9.5)
    text(LX + (idx - 0.6) * W, y + RH + 0.01, 2.2 * W, 0.18, dat, size=7.5,
         color=MUTED, align=PP_ALIGN.CENTER)

y += 0.68
band(y, RH)
text(0.42, y - 0.01, 3.05, 0.22, "Kc", size=11, bold=True, align=PP_ALIGN.RIGHT)
text(0.42, y + 0.19, 3.05, 0.18, "coeficient de cultiu", size=8, color=MUTED,
     align=PP_ALIGN.RIGHT)
chip(KC_SPAN[0], KC_SPAN[1], y + 0.04, RH - 0.08, C_KC, "0,45", WHITE, 10)

y += 0.60
text(0.42, y, 3.05, 0.22, "Kr", size=11, bold=True, align=PP_ALIGN.RIGHT)
text(0.42, y + 0.20, 3.05, 0.18, "coeficient de reducció per capçada", size=8,
     color=MUTED, align=PP_ALIGN.RIGHT)
y += 0.44
for code, nom, kr in TRACT:
    band(y, KH)
    text(0.42, y, 3.05, KH, f"{nom}  ({code})", size=9, align=PP_ALIGN.RIGHT)
    f1, t1 = KR["1,0"]
    chip(8, 11, y + 0.03, KH - 0.06, f1, "1,0", t1, 8)
    f2, t2 = KR[kr]
    chip(12, 27, y + 0.03, KH - 0.06, f2, kr, t2, 8)
    y += 0.345

y += 0.22
band(y, RH)
text(0.42, y - 0.01, 3.05, 0.22, "Reg aplicat", size=11, bold=True,
     align=PP_ALIGN.RIGHT)
text(0.42, y + 0.19, 3.05, 0.18, "fracció de la dosi", size=8, color=MUTED,
     align=PP_ALIGN.RIGHT)
for (a, b), txt in REG:
    fc, tc = RG[txt]
    chip(a, b, y + 0.04, RH - 0.08, fc, txt if b - a >= 3 else "", tc, 9.5)
    if b - a < 3:
        text(LX + (b + 1.25) * W, y, 1.2, RH, txt, size=8.5, bold=True)

# llegenda i notes
ly = y + 0.56
text(LX, ly, 0.4, 0.20, "Kr:", size=9, bold=True)
x = LX + 0.42
for kr in ("0,7", "1,0", "1,1", "1,4"):
    box(x, ly + 0.02, 0.22, 0.16, KR[kr][0], line=WHITE, rounded=True)
    text(x + 0.27, ly, 0.35, 0.20, kr, size=8.5)
    x += 0.72
x += 0.25
text(x, ly, 0.45, 0.20, "Reg:", size=9, bold=True)
x += 0.48
for txt, desc in (("100 %", "sense restricció"), ("25 %", "deficitari controlat")):
    box(x, ly + 0.02, 0.22, 0.16, RG[txt][0], line=WHITE, rounded=True)
    text(x + 0.27, ly, 2.2, 0.20, f"{txt}, {desc}", size=8.5)
    x += 2.6

text(LX, ly + 0.34, 9.4, 0.20,
     "Desbrot i tall en altura el 02/06; desfullat del tractament ES+DV el 20/07, "
     "coincidint amb el verolat.", size=8.5, color=MUTED)
text(LX, ly + 0.58, 9.4, 0.20,
     "El Kc es manté constant de la primera setmana de maig a l'última de setembre. "
     "El reg deficitari controlat s'aplica des del verolat fins a la verema.",
     size=8.5, color=MUTED)

prs.save(OUT)
print("calendari_maneig_2026.pptx", len(slide.shapes), "formes")
