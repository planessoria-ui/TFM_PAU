# -*- coding: utf-8 -*-
"""
Fase 16: incorpora els quatre vols de 2026.

  (a) Taula 3: corregeix la data del tercer vol i hi afegeix la resolució real.
  (b) Resol la nota de les dates de calibratge.
  (c) Taula 13: quatre vols, 472 registres.
  (d) Reforça la nota del calibratge tèrmic amb la inversió de l'ordre.
  (e) Reforça la nota de la màscara de l'entrefilera amb la sèrie estacional.
  (f) 5.1 amb la Taula 14 i la Figura 16.
"""
import csv, statistics as st
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PATH = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v8.docx"
FIG = "/home/user/TFM_PAU/figures/"
U = "/root/.claude/uploads/446d5a3f-407d-5d51-8cea-50e8348ad898/"
doc = docx.Document(PATH)
FONT = "Arial"
RED = RGBColor(0xC0, 0x00, 0x00)
_PARENT = doc.paragraphs[0]._parent

F = [("03/06", "a290cdc2-biophysical_traits_YIELD4CAST2026_covercrop_P25_v2.csv"),
     ("30/06", "57255596-biophysical_traits_YIELD4CAST2026_covercrop_P25.csv"),
     ("22/07", "7e8f3dc6-biophysical_traits_YIELD4CAST2026_covercrop_P25_22072026.csv"),
     ("12/08", "9b33ef42-biophysical_traits_YIELD4CAST2026_covercrop_P25_12082026.csv")]
SHORT = {"Esporga severa i desfullat a verolat": "ES+DV", "Esporga severa + desfullat": "ES+DV",
         "Esporga severa": "ES", "Esporga mitjana": "EM",
         "Esporga lleugera": "ELL", "Sense esporga": "C"}
ORDER = ["C", "ELL", "EM", "ES", "ES+DV"]
NAME = {"C": "Sense esporga", "ELL": "Esporga lleugera", "EM": "Esporga mitjana",
        "ES": "Esporga severa", "ES+DV": "Esporga severa + desfullat"}


def f2(s):
    s = (s or "").strip().replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


DATA = []
for lab, fn in F:
    raw = open(U + fn, encoding="utf-8-sig").read()
    d = ";" if raw.count(";") > raw.count(",") else ","
    for r in csv.DictReader(raw.splitlines(), delimiter=d):
        if not (r.get("tree") or "").strip():
            continue
        DATA.append({"d": lab, "t": SHORT[r["Treatment"].strip()],
                     "fc": f2(r.get("fc")), "hc": f2(r.get("hc")),
                     "fVolc": f2(r.get("fVolc")), "wc_hc": f2(r.get("wc_hc"))})


def cell(var, date, trt, dec):
    v = [x[var] for x in DATA if x["d"] == date and x["t"] == trt and x[var] is not None]
    if not v:
        return "—"
    return f"{st.mean(v):.{dec}f} ± {st.pstdev(v):.{dec}f}".replace(".", ",")


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


def note(el, text, tag="[A VERIFICAR] "):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(6)
    _font(p.add_run(tag), 11, bold=True, color=RED)
    _font(p.add_run(text), 11, color=RED)
    return p._element


def table(el, rows, widths, left=(), bold_rows=()):
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
            _font(pp.add_run(val), 9, bold=(r == 0 or r in bold_rows))
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
# (a) Taula 3: data del tercer vol i resolució real
# ===========================================================================
cap3 = para("Taula 3. Calendari previst de vols")
set_text(cap3, "", 10)
cap3.runs[0].text = "Taula 3."
_font(cap3.runs[0], 10, bold=True)
_font(cap3.add_run(
    " Vols de dron de la campanya de 2026 i mesures de camp simultànies, amb la "
    "fase fenològica de la vinya i la mida de píxel amb què s'ha processat cada "
    "vol. Font: elaboració pròpia."), 10)

tb = cap3._element.getnext()
while tb is not None and tb.tag != qn('w:tbl'):
    tb = tb.getnext()
t3 = [t for t in doc.tables if t._element is tb][0]
t3.cell(3, 2).paragraphs[0].runs[0].text = "22/07/2026"
t3.add_column(Cm(2.6))
for i, v in enumerate(["Mida de píxel (cm)", "1,34", "2,94", "2,92", "2,66"]):
    c = t3.cell(i, len(t3.columns) - 1)
    pp = c.paragraphs[0]
    pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    _font(pp.add_run(v), 9.5, bold=(i == 0))

# ===========================================================================
# (b) Nota de les dates de calibratge, resolta
# ===========================================================================
p = para("Les dates de calibratge registrades al full Temp_CAL")
set_text(p, "", 11)
p.runs[0].text = "[A VERIFICAR] "
_font(p.runs[0], 11, bold=True, color=RED)
_font(p.add_run(
    "Els quatre vols processats són els del 3 i el 30 de juny, el 22 de juliol i "
    "el 12 d'agost de 2026, segons la columna de data dels fitxers de sortida. Al "
    "full Temp_CAL, en canvi, hi consten el 3 i el 22 de juny, el 30 de juny i el "
    "13 d'agost. Les files del full estan ordenades cronològicament excepte la del "
    "22 de juny, que hi apareix després de la del 30 de juny, de manera que sembla "
    "que hauria de dir 22 de juliol; i la del 13 d'agost queda a un sol dia del vol "
    "del 12. Si es confirma, les quatre rectes de la Taula 9 es corresponen una a "
    "una amb els quatre vols. Val la pena verificar-ho i corregir el full, perquè "
    "tal com està ara dos dels quatre vols no tenen cap recta amb data coincident."),
    11, color=RED)

# ===========================================================================
# (c) Taula 13: quatre vols
# ===========================================================================
p = para("El fitxer corresponent al vol del 3 de juny de 2026 conté 118 files")
for r in p.runs:
    r.text = r.text.replace(
        "El fitxer corresponent al vol del 3 de juny de 2026 conté 118 files, una "
        "per cada cep de l'assaig.",
        "Cada vol en genera 118, una per cep, de manera que la campanya de 2026 "
        "aplega 472 registres.")

# ===========================================================================
# (d) Nota del calibratge tèrmic: inversió de l'ordre entre tractaments
# ===========================================================================
p = para("Cal comprovar amb l'equip l'hora de les mesures de l'IRT")
el = note(p._element,
    "Hi ha encara un tercer indici, i és el més difícil d'explicar. Al vol del 3 de "
    "juny, l'ordre dels tractaments segons la temperatura de dosser surt invertit "
    "entre els dos mètodes: el termòmetre d'infraroig dona el tractament sense "
    "esporga com el més calent (27,1 °C de mitjana) i l'esporga severa amb desfullat "
    "com el més fred (25,7 °C), mentre que el dron dona exactament el contrari, 27,8 "
    "°C i 31,4 °C respectivament. A més, el dron obre una forquilla de 3,6 °C entre "
    "tractaments allà on l'IRT només en marca 1,4. La sèrie completa reprodueix "
    "aquest mateix ordre als quatre vols, sempre amb els tractaments de capçada més "
    "petita com els més calents. Si la Tc del dron està inflada justament als "
    "tractaments amb menys coberta, per barreja amb el sòl de l'entrefilera, el "
    "model TSEB els atribuirà menys evapotranspiració i la diferència entre "
    "tractaments podria ser en bona part un artefacte. És el punt que caldria "
    "resoldre abans de res, perquè afecta la pregunta central del treball.")

# ===========================================================================
# (e) Nota de la màscara de l'entrefilera: sèrie estacional
# ===========================================================================
p = para("Cal decidir si es gira la comparació")
for r in p.runs:
    r.text = r.text.replace(
        "Cal decidir si es gira la comparació",
        "Els altres tres vols hi afegeixen una comprovació: la fracció creix al "
        "llarg de l'estiu, de 0,0003 el 3 de juny a 0,019 el 12 d'agost, que és "
        "just el que s'espera del sòl nu a mesura que la coberta del carrer "
        "s'asseca, i no pas de la coberta vegetal, que hauria de seguir el camí "
        "contrari. Fins i tot al vol d'agost, però, només un 1,9 % de la superfície "
        "del polígon queda classificada en aquesta classe, de manera que el gruix "
        "de l'entrefilera no el recull cap de les dues màscares. Cal decidir si es "
        "gira la comparació")

# ===========================================================================
# (f) 5.1 amb dades reals
# ===========================================================================
p = para("CAPÍTOL A OMPLIR AMB LES DADES EXPERIMENTALS REALS")
set_text(p, "", 11)
p.runs[0].text = "[CAPÍTOL PARCIALMENT OMPLERT] "
_font(p.runs[0], 11, bold=True, color=RED)
_font(p.add_run(
    "L'apartat 5.1 recull ja els paràmetres estructurals dels quatre vols de 2026. "
    "Els apartats 5.2 a 5.4 queden pendents de les mesures de ceptòmetre, de la "
    "sortida dels models d'evapotranspiració i de les lectures de potencial hídric "
    "de tija."), 11, color=RED)

p = para("Taules i figures de LAI, fIPAR, alçada i volum del dosser")
set_text(p, (
    "Els quatre vols de la campanya de 2026 han donat, per a cadascun dels 118 ceps "
    "de l'assaig, el joc complet de paràmetres estructurals descrit a l'apartat "
    "4.6.2. La Taula 14 en recull la mitjana i la desviació estàndard per tractament "
    "i data, i la Figura 16 mostra com evolucionen la fracció de coberta i l'alçada "
    "del dosser al llarg de l'estiu."))
el = p._element

el = body(el,
    "La fracció de coberta separa els tractaments amb claredat i en manté l'ordre a "
    "les quatre dates. El tractament sense esporga és sempre el que té més coberta, "
    "entre 0,31 i 0,41 segons la data, seguit de l'esporga lleugera; l'esporga "
    "mitjana se situa a la franja intermèdia, al voltant de 0,25; i els dos "
    "tractaments severs queden per sota de 0,20. Al vol del 22 de juliol la "
    "diferència entre els extrems supera el doble. L'evolució estacional també és "
    "diferent segons el tractament: els de capçada més gran creixen fins al juliol i "
    "després es mantenen o baixen lleugerament, mentre que l'esporga severa amb "
    "desfullat disminueix de manera sostinguda i és l'únic tractament que acaba "
    "l'estiu amb menys coberta que al juny.")
el = body(el,
    "L'alçada del dosser separa els tractaments en dos grups nets. Els tres menys "
    "severs es mouen entre 2,15 i 2,37 m i els dos severs, entre 1,69 i 1,88 m, amb "
    "una diferència d'uns 40 cm que es manté a totes les dates. Dins de cada grup "
    "les diferències són petites i les desviacions estàndard se solapen, de manera "
    "que l'alçada distingeix la intensitat de l'esporga però no els matisos entre "
    "tractaments contigus.")
el = body(el,
    "El volum de copa per unitat de superfície de sòl amplifica les diferències de "
    "la fracció de coberta, perquè hi combina l'alçada: va de 0,22 m a l'esporga "
    "severa amb desfullat fins a 0,64 m al tractament sense esporga. La relació "
    "entre amplada i alçada, en canvi, es manté entre 0,48 i 0,66 a tots els "
    "tractaments i dates, cosa que indica que les capçades conserven una forma "
    "relativament constant encara que canviïn de mida.")

el = caption(el, "Taula 14.",
             "Paràmetres estructurals del dosser per tractament de maneig de capçada "
             "i data de vol, expressats com a mitjana ± desviació estàndard dels "
             "ceps de cada tractament (24 ceps, o 23 a l'esporga severa i al "
             "tractament sense esporga). fc és la fracció de coberta, hc l'alçada "
             "del dosser, fVolc el volum de copa per unitat de superfície de sòl i "
             "wc/hc la relació entre l'amplada i l'alçada de la capçada.")

rows = [("Variable i tractament", "03/06/2026", "30/06/2026", "22/07/2026", "12/08/2026")]
bold = set()
for var, lab, dec in [("fc", "Fracció de coberta, fc (–)", 3),
                      ("hc", "Alçada del dosser, hc (m)", 2),
                      ("fVolc", "Volum per unitat de sòl, fVolc (m)", 3),
                      ("wc_hc", "Relació amplada/alçada, wc/hc (–)", 3)]:
    bold.add(len(rows))
    rows.append((lab, "", "", "", ""))
    for t in ORDER:
        rows.append((NAME[t],) + tuple(cell(var, d, t, dec) for d, _ in F))
el = table(el, rows, [Cm(5.0), Cm(2.8), Cm(2.8), Cm(2.8), Cm(2.8)],
           left=(0,), bold_rows=bold)

el = picture(el, "fig16_estructura_serie.png", 16.0)
el = caption(el, "Figura 16.",
             "Evolució dels paràmetres estructurals del dosser al llarg dels quatre "
             "vols de 2026, per tractament de maneig de capçada: (A) fracció de "
             "coberta i (B) alçada del dosser. Cada punt és la mitjana dels ceps del "
             "tractament; la dispersió es recull a la Taula 14. Font: elaboració "
             "pròpia a partir de la sortida del processament.")

el = body(el,
    "L'evolució de l'alçada s'ha de llegir amb una precaució. Els quatre vols no "
    "s'han processat a la mateixa resolució: la mida de píxel és d'1,34 cm al vol "
    "del 3 de juny i d'entre 2,66 i 2,94 cm als altres tres (Taula 3). Com que "
    "l'alçada del dosser es pren com el valor màxim dins del polígon de cada cep, un "
    "píxel més fi mostreja més punts i tendeix a donar un màxim més alt, i això "
    "explica probablement per què el vol de juny apareix per damunt dels altres a "
    "tots els tractaments de la Figura 16B. La fracció de coberta, que és un "
    "recompte relatiu, hi és molt menys sensible, i de fet la seva evolució sí que "
    "segueix un patró de creixement coherent amb l'estació.")
el = note(el,
    "Convé decidir amb l'equip si es tornen a processar els quatre vols a una "
    "resolució comuna o si l'alçada del dosser es calcula amb un percentil alt, per "
    "exemple el 95, en lloc del màxim. Mentre l'alçada depengui de la mida de píxel, "
    "la comparació entre dates no és neta, encara que la comparació entre "
    "tractaments dins d'una mateixa data sí que ho sigui.")
el = body(el,
    "Les temperatures de dosser i de sòl que surten d'aquest mateix processament no "
    "es presenten en aquest apartat, perquè depenen de la calibració de la banda "
    "tèrmica, que té pendent la verificació descrita a l'apartat 4.6. Un cop "
    "resolta, s'incorporaran aquí i alimentaran els apartats 5.3 i 5.4.")

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
print(f"OK · paràgrafs: {len(doc.paragraphs)} · taules: {len(doc.tables)} "
      f"· registres processats: {len(DATA)}")
