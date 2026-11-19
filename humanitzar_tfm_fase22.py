# -*- coding: utf-8 -*-
"""
Fase 22: incorpora les fotografies de camp i de laboratori als apartats que
els corresponen.

  Figura 18, al 4.6   presa de temperatures de referència amb el termòmetre
                      d'infraroig (panell negre i vegetació)
  Figura 19, al 4.7.2 mesura destructiva de l'àrea foliar dels ceps defoliats
  Figura 20, al 4.6.2 aixecament de la coordenada GPS de cada cep
"""
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PATH = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v8.docx"
FIG = "/home/user/TFM_PAU/figures/"
doc = docx.Document(PATH)
FONT = "Arial"
_PARENT = doc.paragraphs[0]._parent


def _font(run, size=12, bold=False, name=FONT):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts'); rpr.append(rf)
    for a in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rf.set(qn(a), name)


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


def picture(el, name, width_cm):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(FIG + name, width=Cm(width_cm))
    return p._element


def caption(el, head, text):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    _font(p.add_run(head), 10, bold=True)
    _font(p.add_run(" " + text), 10)
    return p._element


# ===========================================================================
# Figura 18: termòmetre d'infraroig, al 4.6
# ===========================================================================
el = para("Les temperatures provenen del fitxer enregistrat")._element
el = body(el,
    "Les lectures es prenen el mateix dia del vol i a la mateixa franja horària, "
    "apuntant el termòmetre d'infraroig directament a cada element i a poca "
    "distància de la superfície. Els panells de referència es mesuren estesos a "
    "l'entrefilera, al lloc on s'han col·locat per al vol (Figura 18A), i la "
    "vegetació es mesura a la part alta del dosser, que és la que la càmera veu des "
    "del dron (Figura 18B). De cada element se'n prenen unes deu lectures i la seva "
    "mitjana és la que s'aparella amb el valor de la banda 6 del mateix element.")
el = picture(el, "fig_calibratge_irt.png", 16.0)
el = caption(el, "Figura 18.",
             "Presa de les temperatures de referència amb el termòmetre d'infraroig "
             "el dia del vol: (A) panell negre, estès a l'entrefilera al costat del "
             "panell blanc, i (B) vegetació, mesurada a la part alta del dosser. "
             "Font: elaboració pròpia.")

# ===========================================================================
# Figura 20: aixecament GPS, al 4.6.2
# ===========================================================================
el = para("Amb les capes ja emmascarades, els estadístics es calculen")._element
el = body(el,
    "L'aixecament es fa cep a cep: s'hi col·loca l'antena del GPS a la base de la "
    "soca, se'n registra la coordenada i es passa al següent (Figura 20). Amb les "
    "coordenades de tots els ceps de l'assaig es munta la capa de polígons, que "
    "després s'ajusta sobre l'ortomosaic i el model digital de vegetació ja "
    "georeferenciats. És aquesta capa la que determina de quins píxels surt cada "
    "paràmetre estructural de cada cep.")
el = picture(el, "fig_gps_ceps.png", 8.2)
el = caption(el, "Figura 20.",
             "Aixecament de la coordenada GPS de cada cep de la parcel·la "
             "experimental, a partir del qual es construeix la capa de polígons "
             "d'on s'extreuen els paràmetres estructurals del dosser. "
             "Font: elaboració pròpia.")

# ===========================================================================
# Figura 19: nou apartat 4.7.2, abans del 4.8
# ===========================================================================
h48 = [p for p in doc.paragraphs
       if p.style.name == 'Heading 2' and p.text.strip().startswith('4.8.')][0]
anchor = h48._element.getprevious()

h = after(anchor, 'Heading 3')
_font(h.add_run("4.7.2. Mesura destructiva de l'àrea foliar"), 13, bold=True)
el = h._element
el = body(el,
    "El ceptòmetre i les imatges de dron donen estimacions del LAI, de manera que "
    "calen valors mesurats directament per comprovar si s'hi acosten. Amb aquesta "
    "finalitat es defolien completament alguns ceps i se'n mesura al laboratori la "
    "superfície de totes les fulles. Les fulles es numeren una a una (Figura 19A) i "
    "es passen pel mesurador d'àrea foliar LI-3100C, que les fa avançar sobre una "
    "cinta transparent i en retorna la superfície en centímetres quadrats "
    "(Figura 19B).")
el = body(el,
    "Sumant les fulles d'un cep s'obté la seva àrea foliar total i, dividint-la pel "
    "marc de plantació de 4,25 m², el seu índex d'àrea foliar. A banda dels ceps "
    "defoliats dins d'aquest assaig, durant la campanya se n'han defoliat quinze "
    "més en altres assajos del mateix programa. Aquests ceps amplien el nombre de "
    "parelles disponibles i el rang de mides de capçada, i són els que permeten "
    "ajustar les relacions empíriques entre el LAI mesurat, la fIPAR diària i el "
    "LAI estimat a partir de les imatges.")
el = picture(el, "fig_lai_destructiu.png", 16.0)
el = caption(el, "Figura 19.",
             "Mesura destructiva de l'àrea foliar dels ceps defoliats: (A) fulles "
             "d'un cep numerades abans de passar-les pel mesurador i (B) mesurador "
             "d'àrea foliar LI-3100C. Font: elaboració pròpia.")

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
print(f"OK · paràgrafs: {len(doc.paragraphs)}")
