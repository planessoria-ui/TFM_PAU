# -*- coding: utf-8 -*-
"""
Fase 33: omple els apartats 5.3 i 5.4 dels resultats, el capítol 6 de discussió
i el capítol 7 de conclusions, amb l'evapotranspiració, el coeficient de cultiu
i la seva relació amb la fIPAR, i la comparació amb Picón-Toro et al. (2012).
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
VERMELL = RGBColor(0xC0, 0x00, 0x00)
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


def H(nivell, inici):
    for p in doc.paragraphs:
        if p.style.name == f'Heading {nivell}' and p.text.strip().startswith(inici):
            return p
    raise LookupError(inici)


def after(el, style=None):
    new = OxmlElement('w:p')
    el.addnext(new)
    p = docx.text.paragraph.Paragraph(new, _PARENT)
    if style:
        p.style = doc.styles[style]
    return p


def body(el, text, color=None, size=12):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    p.paragraph_format.space_after = Pt(6)
    _font(p.add_run(text), size, color=color)
    return p._element


def head(el, nivell, text):
    p = after(el, f'Heading {nivell}')
    _font(p.add_run(text), 13 if nivell == 2 else 12, bold=True)
    return p._element


def caption(el, cap, text):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    _font(p.add_run(cap), 10, bold=True)
    _font(p.add_run(" " + text), 10)
    return p._element


def picture(el, nom, cm):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(FIG + nom, width=Cm(cm))
    return p._element


def table(el, rows, widths, grups=()):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = doc.styles['Table Grid']
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    lay = OxmlElement('w:tblLayout'); lay.set(qn('w:type'), 'fixed')
    t._tbl.tblPr.append(lay)
    grid = t._tbl.find(qn('w:tblGrid'))
    for gc, w in zip(grid.findall(qn('w:gridCol')), widths):
        gc.set(qn('w:w'), str(int(w.twips)))
    for r, data in enumerate(rows):
        for c, val in enumerate(data):
            cl = t.cell(r, c)
            cl.width = widths[c]
            pp = cl.paragraphs[0]
            pp.alignment = (WD_ALIGN_PARAGRAPH.LEFT if c == 0 and r
                            else WD_ALIGN_PARAGRAPH.CENTER)
            pp.paragraph_format.space_before = Pt(2)
            pp.paragraph_format.space_after = Pt(2)
            pp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            _font(pp.add_run(val), 9, bold=(r == 0 or r in grups))
    el.addnext(t._element)
    return t._element


def llista(el, text):
    p = after(el, 'List Number')
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    p.paragraph_format.space_after = Pt(6)
    _font(p.add_run(text), 12)
    return p._element


NOMS = ["Esporga severa + desfullat", "Esporga severa", "Esporga mitjana",
        "Esporga lleugera", "Sense esporga"]
D4 = ["03/06/2026", "30/06/2026", "23/07/2026", "12/08/2026"]

# ===========================================================================
# 5.3 Evapotranspiració
# ===========================================================================
el = H(2, '5.3.')._element

el = body(el,
    "El model TSEB-2T dona, per a cada cep i cada vol, l'evapotranspiració real "
    "diària. La Taula 19 recull la mitjana per tractament, el coeficient de cultiu "
    "que en resulta, calculat com el quocient entre l'ETa i l'evapotranspiració de "
    "referència del dia, i la fIPAR diària que hi correspon.")

el = caption(el, "Taula 19.",
             "Evapotranspiració real, coeficient de cultiu i fIPAR diària per "
             "tractament de maneig de capçada i data de vol. Mitjana dels ceps de "
             "cada tractament. El coeficient de cultiu és el quocient entre l'ETa "
             "i l'ETo del dia, que val 5,59, 6,66, 6,51 i 5,98 mm d⁻¹ "
             "respectivament. Font: elaboració pròpia.")

T19 = [["Variable i tractament"] + D4,
       ["Evapotranspiració real, ETa (mm d⁻¹)", "", "", "", ""],
       [NOMS[0], "3,216", "3,659", "3,170", "3,450"],
       [NOMS[1], "3,139", "3,665", "3,295", "3,479"],
       [NOMS[2], "3,567", "4,269", "3,947", "3,960"],
       [NOMS[3], "3,660", "4,714", "4,512", "4,339"],
       [NOMS[4], "4,266", "4,947", "4,726", "4,516"],
       ["Coeficient de cultiu, Kc (–)", "", "", "", ""],
       [NOMS[0], "0,575", "0,549", "0,487", "0,577"],
       [NOMS[1], "0,562", "0,550", "0,506", "0,582"],
       [NOMS[2], "0,638", "0,641", "0,606", "0,662"],
       [NOMS[3], "0,655", "0,708", "0,693", "0,726"],
       [NOMS[4], "0,763", "0,743", "0,726", "0,755"],
       ["Fracció de radiació interceptada, fIPAR (%)", "", "", "", ""],
       [NOMS[0], "22,7", "28,3", "25,9", "32,9"],
       [NOMS[1], "24,9", "29,9", "30,7", "35,5"],
       [NOMS[2], "28,1", "41,0", "38,0", "42,7"],
       [NOMS[3], "30,0", "48,1", "47,5", "51,5"],
       [NOMS[4], "39,3", "51,9", "45,5", "54,7"]]
el = table(el, T19, [Cm(6.4), Cm(2.4), Cm(2.4), Cm(2.4), Cm(2.4)], grups=(1, 7, 13))

el = body(el,
    "L'ordenació dels tractaments es manté als quatre vols. Com a mitjana de la "
    "campanya, el tractament sense esporga consumeix 4,61 mm d⁻¹, l'esporga "
    "lleugera 4,31 (un 6,7 % menys), l'esporga mitjana 3,94 (un 14,7 % menys), "
    "l'esporga severa 3,39 (un 26,4 % menys) i l'esporga severa amb desfullat 3,37 "
    "(un 26,9 % menys). Els dos tractaments d'esporga severa són indistingibles "
    "entre ells, de manera que el desfullat del 20 de juliol no redueix el consum "
    "per damunt del que ja fa l'esporga.")

el = picture(el, "fig27_eta_kc.png", 16.0)
el = caption(el, "Figura 27.",
             "Evolució de l'evapotranspiració real i del coeficient de cultiu per "
             "tractament al llarg de la campanya de 2026. Les barres del panell "
             "esquerre són la desviació estàndard dels ceps de cada tractament. "
             "Font: elaboració pròpia.")

el = body(el,
    "Tots cinc tractaments tenen el consum màxim el 30 de juny. El 23 de juliol el "
    "coeficient de cultiu és el mínim de l'any a quatre dels cinc tractaments, i "
    "això passa amb la capçada més gran de la campanya. Entre les dues dates l'ETa "
    "baixa un 7,6 % mentre que l'ETo només baixa un 2,3 %, i el quocient entre "
    "l'ETa i l'ETp passa de 0,875 a 0,728, la davallada més gran de la campanya. "
    "La reducció del coeficient ve, doncs, de la planta i no de la demanda "
    "atmosfèrica.")

el = body(el,
    "El coeficient de cultiu segueix la fracció de radiació interceptada (Figura "
    "28). Dins de cada vol la relació és molt estreta, amb coeficients de "
    "determinació d'entre 0,95 i 0,996, però la recta es desplaça d'una data a "
    "l'altra. Ajustant les vint mitjanes de tractament i data alhora, la relació és "
    "Kc = 0,708 · fIPAR + 0,370, amb un coeficient de determinació de 0,688. El vol "
    "del 3 de juny queda 0,063 per damunt d'aquesta recta i el del 23 de juliol, "
    "0,032 per sota: per a una mateixa fracció de radiació interceptada, el consum "
    "relatiu és més alt a principi de campanya i més baix al moment de més dèficit "
    "hídric.")

el = picture(el, "fig28_kc_vs_fipar.png", 14.0)
el = caption(el, "Figura 28.",
             "Coeficient de cultiu contra fracció de radiació interceptada, per "
             "tractament i data de vol. Cada punt és la mitjana dels ceps d'un "
             "tractament. La recta discontínua és l'ajust conjunt de les vint "
             "observacions. Font: elaboració pròpia.")

el = body(el,
    "[A VERIFICAR] L'ETa d'aquest apartat prové d'executar el model amb el fitxer "
    "d'entrada descrit a l'apartat 4.8, la fracció de coberta del qual no coincideix "
    "amb la que recull la Taula 14 del 5.1. La diferència creix al llarg de la "
    "campanya i amb la severitat de l'esporga, fins a un 36 % al tractament ES+DV "
    "del 12 d'agost. Cal determinar quin dels dos conjunts prové del model "
    "d'elevacions del bloc de l'assaig i refer el que calgui, perquè si el bo és el "
    "de la Taula 14 s'ha de tornar a executar el model.", color=VERMELL)

# ===========================================================================
# 5.4 Estat hídric
# ===========================================================================
el = H(2, '5.4.')._element

el = body(el,
    "El potencial hídric de tija al migdia és la variable amb què es decideix el reg "
    "i la que verifica que el disseny de l'assaig funciona. La Taula 20 en recull la "
    "mitjana per tractament, juntament amb l'índex d'estrès hídric derivat del "
    "balanç d'energia.")

el = caption(el, "Taula 20.",
             "Potencial hídric de tija al migdia i índex d'estrès hídric del cultiu "
             "per tractament i data de vol. El CWSI es calcula com 1 − ETa/ETp amb "
             "l'ETp de les hores de llum. Font: elaboració pròpia.")

T20 = [["Variable i tractament"] + D4,
       ["Potencial hídric de tija (MPa)", "", "", "", ""],
       [NOMS[0], "−0,43", "−0,71", "−0,91", "−0,96"],
       [NOMS[1], "−0,47", "−0,70", "−0,99", "−1,05"],
       [NOMS[2], "−0,47", "−0,73", "−0,93", "−1,10"],
       [NOMS[3], "−0,47", "−0,72", "−0,83", "−1,04"],
       [NOMS[4], "−0,51", "−0,76", "−0,99", "−0,96"],
       ["Índex d'estrès hídric, CWSI (–)", "", "", "", ""],
       [NOMS[0], "0,179", "0,180", "0,370", "0,378"],
       [NOMS[1], "0,219", "0,189", "0,368", "0,382"],
       [NOMS[2], "0,136", "0,130", "0,279", "0,325"],
       [NOMS[3], "0,129", "0,075", "0,190", "0,273"],
       [NOMS[4], "0,054", "0,048", "0,153", "0,244"]]
el = table(el, T20, [Cm(6.4), Cm(2.4), Cm(2.4), Cm(2.4), Cm(2.4)], grups=(1, 7))

el = body(el,
    "El potencial de tija baixa de manera continuada al llarg de la campanya, de "
    "−0,47 MPa de mitjana el 3 de juny a −1,02 MPa el 12 d'agost. La diferència "
    "entre tractaments dins d'una mateixa data, en canvi, es manté entre 0,06 i "
    "0,16 MPa, molt per sota dels 0,55 MPa que separen el primer vol de l'últim. "
    "Regar cada tractament amb una dosi proporcional a la seva capçada ha "
    "aconseguit el que es buscava: els cinc tractaments arriben a cada data amb "
    "pràcticament el mateix estat hídric, i les diferències de consum de l'apartat "
    "anterior són, per tant, atribuïbles a la capçada i no a l'aigua disponible.")

el = picture(el, "fig29_psi_cwsi.png", 16.0)
el = caption(el, "Figura 29.",
             "Potencial hídric de tija al migdia i índex d'estrès hídric del cultiu "
             "per tractament al llarg de la campanya de 2026. La línia vertical "
             "marca l'inici del reg deficitari controlat. Font: elaboració pròpia.")

el = body(el,
    "L'índex d'estrès hídric segueix la mateixa evolució estacional: puja de 0,14 "
    "de mitjana el 3 de juny a 0,32 el 12 d'agost, amb el salt més gran entre el 30 "
    "de juny i el 23 de juliol, coincidint amb el reg deficitari. Entre tractaments, "
    "en canvi, l'índex reprodueix l'ordre de la mida de la capçada a totes les "
    "dates, amb coeficients de correlació amb la fIPAR d'entre −0,92 i −0,99.")

el = body(el,
    "[A VERIFICAR] Les diferències de CWSI entre tractaments no es poden interpretar "
    "com a diferències d'estrès hídric. El potencial de tija mostra que tots cinc "
    "estan al mateix estat hídric, i dins de les dues primeres dates el CWSI "
    "correlaciona positivament amb el potencial, és a dir en sentit contrari al que "
    "li tocaria. L'origen és aritmètic: l'ETa varia entre un 31 i un 49 % entre "
    "tractaments mentre que l'ETp només varia entre un 8 i un 16 %, de manera que el "
    "quocient penalitza les capçades petites. Perquè l'índex pugui comparar "
    "tractaments, l'ETp ha de respondre a la capçada tant com hi respon l'ETa, cosa "
    "que passa per assignar a cada cep la seva pròpia temperatura de dosser.",
    color=VERMELL)

doc.save(PATH)
print(f"OK · paràgrafs: {len(doc.paragraphs)} · taules: {len(doc.tables)}")
