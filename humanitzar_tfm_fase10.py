# -*- coding: utf-8 -*-
"""
Fase 10: complementa el TFM amb el protocol de processat de dades.

  - 4.5: corregeix la descripció de les càmeres (la tèrmica és la banda 6 LWIR
    de l'Altum-PT) i hi afegeix els paràmetres de planificació dels vols.
  - 4.6: completa la calibració de la banda tèrmica amb les dades reals dels
    elements de referència i les equacions dels quatre vols.
  - Condensa el detall operatiu de QGIS i treu els marcadors ja resolts.
"""
import copy
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SRC = "/root/.claude/uploads/446d5a3f-407d-5d51-8cea-50e8348ad898/baefef89-TFM_Pau_Planes_IRTA_v7_1_1_1_1_1.docx"
DST = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v8.docx"
doc = docx.Document(SRC)
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


def find(prefix, style=None, start=0):
    for i in range(start, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        if p.text.strip().startswith(prefix) and (style is None or p.style.name == style):
            return i
    raise LookupError(prefix[:70])


def P(prefix, style=None, start=0):
    return doc.paragraphs[find(prefix, style, start)]


def insert_after(ref, text, template, red=False):
    el = copy.deepcopy(template._element)
    ref._element.addnext(el)
    p = docx.text.paragraph.Paragraph(el, ref._parent)
    set_text(p, text)
    if red:
        for r in p.runs:
            _font(r, 11, italic=True, red=True)
    return p


def add_table_after(ref, headers, rows, widths, caption, tmpl_caption):
    cap_el = copy.deepcopy(tmpl_caption._element)
    ref._element.addnext(cap_el)
    cap = docx.text.paragraph.Paragraph(cap_el, ref._parent)
    for r in cap.runs:
        r._element.getparent().remove(r._element)
    cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _font(cap.add_run(caption), 10)

    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _font(c.paragraphs[0].add_run(h), 10, bold=True)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            _font(cells[i].paragraphs[0].add_run(v), 10)
    for i, w in enumerate(widths):
        for row in t.rows:
            row.cells[i].width = Cm(w)
    cap._element.addnext(t._element)
    return t


tmpl_body = P("Es faran quatre vols de dron")
tmpl_cap = P("Taula 3. Calendari previst", style="Caption")

# ---------------------------------------------------------------------------
# 1. Apartat 4.5: càmeres i pla de vol
# ---------------------------------------------------------------------------
p416 = P("Es faran quatre vols de dron")
set_text(p416,
         "Es faran quatre vols de dron al voltant del migdia solar, repartits al llarg "
         "del cicle vegetatiu, per seguir l'evolució estacional del dosser i de l'estat "
         "hídric. La càmera emprada és una MicaSense Altum-PT, que integra cinc bandes "
         "espectrals (blau 475 nm, verd 560 nm, vermell 668 nm, red-edge 717 nm i "
         "infraroig proper 842 nm), una banda tèrmica LWIR (banda 6) i un sensor de llum "
         "incident. La temperatura de la superfície s'obté d'aquesta banda tèrmica "
         "integrada, que és la que se sotmet al calibratge descrit a l'apartat 4.6. El "
         "calendari de vols i de mesures de camp associades es resumeix a la Taula 3.")

p_todo45 = P("[A CONCRETAR] Model de dron")
set_text(p_todo45,
         "Les missions es planifiquen amb el programari de control del dron. La parcel·la "
         "es troba dins d'una zona CTR, de manera que l'alçada de vol queda limitada a "
         "60 m. Els solapaments frontal i lateral es fixen en valors alts perquè cada "
         "punt del terreny quedi vist des de posicions diferents, cosa que millora el "
         "núvol de punts i, amb ell, la reconstrucció de l'estructura del cep. La càmera "
         "vola orientada al nadir i captura a interval constant. La Taula 7 recull els "
         "paràmetres de la missió.")

add_table_after(
    p_todo45,
    ["Paràmetre", "Valor"],
    [["Alçada de vol", "60 m (límit imposat per la zona CTR)"],
     ["Velocitat de vol", "4,9 m/s"],
     ["Mode de captura", "interval constant d'1,5 s"],
     ["Solapament frontal", "80 %"],
     ["Solapament lateral", "70 %"],
     ["Angle de passada", "329°"],
     ["Orientació de la càmera", "−90° (nadir)"],
     ["Nombre de passades", "31"],
     ["Punts de pas", "64"],
     ["Longitud del recorregut", "6.620 m"],
     ["Superfície coberta", "9,17 ha"],
     ["Resolució a terra", "7,8 cm/píxel"],
     ["Acció en acabar", "retorn al punt d'enlairament"]],
    [6.0, 9.5],
    "Taula 7. Paràmetres de planificació de les missions de vol sobre la parcel·la "
    "experimental.",
    tmpl_cap)

# ---------------------------------------------------------------------------
# 2. Apartat 4.6: condensar el detall operatiu de QGIS
# ---------------------------------------------------------------------------
i0 = find("Un cop creada la capa de polígons")
set_text(doc.paragraphs[i0],
         "Amb la capa de polígons creada, s'extreu per a cada element de referència la "
         "mitjana de cada banda de l'ortomosaic mitjançant les estadístiques de zona de "
         "QGIS. El càlcul s'executa com un procés per lots que recorre les bandes b1 a b6 "
         "i en desa el resultat en un fitxer CSV per banda (Figura 10).")

# esborra la llista de paràmetres de menú (sense tocar paràgrafs amb imatges)
to_drop = []
for j in range(i0 + 1, i0 + 8):
    p = doc.paragraphs[j]
    if p._element.findall('.//' + qn('a:blip')):
        break
    t = p.text.strip()
    if t.startswith(("Capa ràster:", "Banda ràster:", "Capa vectorial",
                     "Prefix de la columna", "Estadístiques a calcular",
                     "Això, es fa mitjançant")):
        to_drop.append(p)
for p in to_drop:
    p._element.getparent().remove(p._element)
print("paràgrafs de menú condensats:", len(to_drop))

# ---------------------------------------------------------------------------
# 3. Calibració de la banda tèrmica
# ---------------------------------------------------------------------------
set_text(P("La sisena banda de l"),
         "La sisena banda de l'Altum-PT és la tèrmica i el seu calibratge relaciona el DN "
         "amb la temperatura de superfície. La recta es construeix en un full a part "
         "(Temp_CAL) amb les dades de calibratge de cada vol. Les temperatures provenen "
         "del fitxer enregistrat a camp amb el termòmetre d'infraroig, que conté unes deu "
         "mesures per element; se'n calcula la mitjana per separat per als elements de "
         "referència i per al cos del sensor, i l'ordre dels elements el dona un fitxer "
         "que acompanya el registre.")

p_b6 = P("mesures per element")
set_text(p_b6,
         "A cada element s'hi assigna el DN de la banda 6 obtingut amb les estadístiques "
         "de zona; per al blanc i el negre s'utilitzen els panells t_white i t_black, i "
         "s'hi afegeixen el sòl nu (bs), la coberta vegetal (cc) i la vegetació (vg) per "
         "cobrir tot el rang de temperatures de l'escena. Si un punt s'allunya molt de la "
         "resta, s'elimina abans d'ajustar la recta. La Taula 8 recull, a tall d'exemple, "
         "els valors del vol del 3 de juny de 2026.")

add_table_after(
    p_b6,
    ["Element de referència", "DN (banda 6)", "Temperatura mesurada (°C)"],
    [["Panell negre (t_black)", "32.444,7", "55,4"],
     ["Panell blanc (t_white)", "31.172,4", "32,9"],
     ["Sòl nu (bs)", "31.687,9", "43,1"],
     ["Coberta vegetal (cc)", "31.554,6", "37,2"],
     ["Vegetació (vg)", "30.633,6", "26,1"]],
    [6.0, 4.5, 5.0],
    "Taula 8. Valors digitals de la banda tèrmica i temperatures mesurades amb el "
    "termòmetre d'infraroig sobre els elements de referència del vol del 3 de juny de "
    "2026, que són les parelles de dades amb què s'ajusta la recta de calibratge.",
    tmpl_cap)

ref = doc.paragraphs[find("Taula 8. Valors digitals")]
# el paràgraf següent a la taula serveix de referència per continuar el text
p_eq = insert_after(ref,
    "L'ajust es repeteix per a cada vol, perquè la resposta del sensor depèn de les "
    "condicions d'adquisició. Les equacions obtingudes en la campanya de 2026 són "
    "T = 0,0130 · DN − 370,18 (3 de juny), T = 0,0131 · DN − 372,82 (22 de juny), "
    "T = 0,0131 · DN − 370,90 (30 de juny) i T = 0,0132 · DN − 374,24 (13 d'agost). Els "
    "pendents es mantenen pràcticament constants entre dates, cosa que indica una "
    "resposta estable del sensor, mentre que l'ordenada a l'origen recull les diferències "
    "de condicions de cada vol. Cada equació s'aplica a la banda 6 del mosaic "
    "corresponent amb la calculadora ràster i en resulta una capa de temperatura de "
    "superfície en graus Celsius, que és la que alimenta l'extracció de la temperatura "
    "de dosser i de sòl.", tmpl_body)

p_n1 = insert_after(p_eq,
    "[A VERIFICAR] Les dates de calibratge registrades al full Temp_CAL (3 i 22 de juny, "
    "30 de juny i 13 d'agost de 2026) no coincideixen del tot amb les de la Taula 3 "
    "(3 i 30 de juny, 23 de juliol i 12 d'agost). Cal unificar-les.", tmpl_body, red=True)

insert_after(p_n1,
    "[A VERIFICAR] En recalcular la recta del 3 de juny per mínims quadrats amb els cinc "
    "elements de la Taula 8 s'obté un pendent de 0,0164 en lloc del 0,0130 registrat al "
    "full. Confirmeu amb l'equip quins punts i quin criteri d'ajust s'utilitzen.",
    tmpl_body, red=True)

# ---------------------------------------------------------------------------
# 4. Marcador ja resolt
# ---------------------------------------------------------------------------
p_old = P("[A CONCRETAR] Programari fotogramètric")
set_text(p_old,
         "[A CONCRETAR] Mètode i índexs concrets per estimar el LAI i la fIPAR a partir "
         "de les imatges, i la seva calibració amb les mesures de camp.")
for r in p_old.runs:
    _font(r, 11, italic=True, red=True)

doc.save(DST)
print("Fase 10 completada ->", DST)
