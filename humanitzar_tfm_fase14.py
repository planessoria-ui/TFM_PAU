# -*- coding: utf-8 -*-
"""
Fase 14: reestructura i completa l'apartat 4.6 a partir del protocol, del full
Temp_CAL i dels scripts de processament de l'IRTA.

  (a)  Taula 9 amb la recta de calibratge tèrmic de cada vol de 2026.
  (b)  Renumera la taula de coeficients radiomètrics (9 -> 10).
  (c)  Refà la nota de verificació del calibratge tèrmic amb els números reals.
  (d)  Mou a 4.5 els dos paràgrafs de planificació de vol que eren a 4.6.
  (e)  Afegeix la descripció del flux d'extracció de variables per cep.
  (f)  Refà 4.6.1 amb els llindars reals (0,5 m i SAVI 0,1) + Figura 15.
  (g)  Refà 4.6.2 amb els paràmetres estructurals reals + Taula 11.
  (h)  Refà 4.6.3 amb els cinc índexs + Taula 12.
  (i)  Nou 4.6.4 sobre temperatures de component i fitxers de sortida.
  (j)  Cinc referències noves a la bibliografia.
  (k)  Annex B amb l'script de paràmetres biofísics.
  (l)  Recalcula la numeració dels camps SEQ.
"""
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PATH = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v8.docx"
FIG = "/home/user/TFM_PAU/figures/"
SRC_BIO = ("/root/.claude/uploads/446d5a3f-407d-5d51-8cea-50e8348ad898/"
           "3be5aba7-biophysical_traits_UAVs_YIELD4CAST_2026.py")
doc = docx.Document(PATH)
FONT, MONO = "Arial", "Courier New"
RED = RGBColor(0xC0, 0x00, 0x00)


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


def find(sub, style=None):
    for i, p in enumerate(doc.paragraphs):
        if sub in p.text and (style is None or p.style.name == style):
            return i
    raise LookupError(sub[:70])


def para(p):
    return doc.paragraphs[find(p)]


_PARENT = None


def after(el, style=None):
    """Insereix un paràgraf buit just després de l'element XML i el retorna."""
    global _PARENT
    if _PARENT is None:
        _PARENT = doc.paragraphs[0]._parent
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


def note(el, text):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(6)
    _font(p.add_run("[A VERIFICAR] "), 11, bold=True, color=RED)
    _font(p.add_run(text), 11, color=RED)
    return p._element


def table(el, rows, widths, align_left_col0=False):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = doc.styles['Table Grid']
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r, data in enumerate(rows):
        for c, val in enumerate(data):
            cell = t.cell(r, c)
            cell.width = widths[c]
            pp = cell.paragraphs[0]
            pp.alignment = (WD_ALIGN_PARAGRAPH.LEFT
                            if (align_left_col0 and c and r) else
                            WD_ALIGN_PARAGRAPH.CENTER)
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
# (b) La taula de coeficients radiomètrics passa a ser la Taula 10
# ===========================================================================
p = para("Coeficients de calibratge radiomètric introduïts")
set_text(p, "")
p.runs[0].text = "Taula 10."
_font(p.runs[0], 10, bold=True)
_font(p.add_run(
    " Coeficients de calibratge radiomètric introduïts a l'script per a cada "
    "banda de la càmera MicaSense Altum-PT. Les cinc primeres bandes converteixen "
    "el valor digital en reflectància (ρ = m · DN + n) i la sisena converteix el "
    "valor de la banda tèrmica en temperatura de superfície en graus Celsius "
    "(T = m · B6 + n)."), 10)

p = para("coeficients emprats es recullen a la Taula")
set_text(p, p.text.replace("es recullen a la Taula 9", "es recullen a la Taula 10"))

p = para("provenen de la versi")
for r in p.runs:
    r.text = r.text.replace("Els valors de la Taula 9", "Els valors de la Taula 10")

# ===========================================================================
# (a) Taula 9: rectes de calibratge tèrmic de cada vol
# ===========================================================================
p = para("ajust es repeteix per a cada vol")
set_text(p, (
    "L'ajust es repeteix per a cada vol, perquè la resposta del sensor depèn de "
    "les condicions d'adquisició. En les equacions resultants la variable "
    "independent és la mateixa banda 6 del mosaic (B6), de manera que la recta no "
    "es fa servir element a element sinó que s'aplica a tota la capa: en aplicar-la "
    "sobre la banda 6 s'obté directament la capa tèrmica corregida, en graus "
    "Celsius. A la pràctica aquesta operació no es fa amb la calculadora ràster de "
    "QGIS sinó com la sisena entrada de l'script de calibratge radiomètric "
    "(Taula 10 i Annex A), de manera que les cinc bandes espectrals i la tèrmica es "
    "corregeixen en una sola passada. La Taula 9 recull les quatre rectes de la "
    "campanya de 2026. Els pendents es mantenen pràcticament constants entre dates, "
    "entre 0,0130 i 0,0132, cosa que indica una resposta estable del sensor, mentre "
    "que l'ordenada a l'origen recull les diferències de condicions de cada vol. La "
    "capa tèrmica corregida és la que alimenta l'extracció de la temperatura de "
    "dosser i de sòl."))

el = caption(p._element, "Taula 9.",
             "Rectes de calibratge de la banda tèrmica (banda 6) de l'Altum-PT per "
             "a cada vol de la campanya de 2026, obtingudes al full Temp_CAL. T és "
             "la temperatura de superfície en graus Celsius i B6, el valor de la "
             "banda 6 de l'ortomosaic. Els elements de referència són el panell "
             "negre (bk), el panell blanc (wh), el sòl nu (bs), la coberta de "
             "l'entrefilera (cc) i la vegetació (vg).")
el = table(el, [
    ("Vol", "Recta de calibratge", "Elements de referència",
     "Rang de temperatura cobert (°C)"),
    ("03/06/2026", "T = 0,0130 · B6 − 370,18", "bk, wh, bs, cc, vg", "26,1 – 55,4"),
    ("22/06/2026", "T = 0,0131 · B6 − 372,82", "bk, wh, bs, cc, vg", "30,4 – 64,9"),
    ("30/06/2026", "T = 0,0131 · B6 − 370,90", "bk, wh, bs, cc, vg", "30,0 – 63,1"),
    ("13/08/2026", "T = 0,0132 · B6 − 374,24", "bk, wh, bs, cc, vg", "34,7 – 65,7"),
], [Cm(2.6), Cm(5.2), Cm(4.2), Cm(4.0)])

# ===========================================================================
# (c) Nota de verificació del calibratge tèrmic
# ===========================================================================
p = para("En recalcular la recta del 3 de juny")
set_text(p, "", 11)
p.runs[0].text = "[A VERIFICAR] "
_font(p.runs[0], 11, bold=True, color=RED)
_font(p.add_run(
    "Cap de les quatre rectes de la Taula 9 es reprodueix ajustant per mínims "
    "quadrats els cinc elements de referència del mateix full: l'ajust directe dona "
    "pendents d'entre 0,0164 i 0,0183, molt per sobre dels 0,0130–0,0132 "
    "registrats, i amb les rectes registrades l'error quadràtic mitjà sobre aquests "
    "mateixos elements va de 2,6 a 5,1 °C. Com que la discrepància es repeteix als "
    "quatre vols, no sembla un error puntual sinó una diferència de criteri. A més, "
    "el full TEMP_v5 conté la temperatura de dosser mesurada amb el termòmetre "
    "d'infraroig als vint PE del vol del 3 de juny, entre 24,1 i 28,6 °C; aplicant-hi "
    "la recta registrada s'obtenen valors d'entre 30,1 i 35,0 °C, és a dir un biaix "
    "proper a +7 °C. Cal confirmar amb l'equip quins punts entren a l'ajust, si el "
    "pendent es fixa a partir de la resposta nominal del sensor en lloc d'ajustar-lo "
    "vol a vol i a quina hora es van prendre les mesures de contrast, abans "
    "d'utilitzar cap temperatura en el balanç energètic."), 11, color=RED)

# ===========================================================================
# (d) Els dos paràgrafs de planificació de vol passen a 4.5
# ===========================================================================
anchor = para("Figura 8. Planificació de la missió de vol")._element
for frag in ("La parcel·la (El polvorín, Raimat) es troba dins",
             "les càmeres del dron (en aquest cas"):
    el = para(frag)._element
    el.getparent().remove(el)
    anchor.addnext(el)
    anchor = el

# ===========================================================================
# (e) Flux d'extracció de variables per cep
# ===========================================================================
el = para("Figura 14. Script de calibratge radiomètric obert")._element
el = body(el,
    "Amb les capes ja calibrades, l'extracció de les variables de cada cep es fa "
    "amb una llibreria interna de l'IRTA escrita en Python, que organitza els "
    "ortomosaics d'una campanya com una col·lecció d'imatges i hi encadena les "
    "operacions de retall, remostreig, emmascarament i estadística zonal. Damunt "
    "d'aquesta llibreria hi ha un script de processament que fixa, per a cada vol, "
    "les carpetes d'entrada, els fitxers vectorials de treball i els llindars de "
    "classificació. L'script complet es reprodueix a l'Annex B.")
el = body(el,
    "Les entrades són tres capes del mateix vol: l'ortomosaic multiespectral "
    "calibrat, la capa de temperatura de superfície derivada de la banda 6 i el "
    "model digital de vegetació amb l'alçada del dosser. Com que no comparteixen "
    "necessàriament la mateixa mida de píxel, les capes d'alçada i de temperatura "
    "es retallen a l'àrea d'interès i es remostregen per interpolació bilineal "
    "sobre la graella de la banda de l'infraroig proper, que és la que marca la "
    "resolució de referència de tot el processament.")
el = body(el,
    "La parcel·la no es processa sencera sinó per blocs. Cada bloc té associat un "
    "polígon que en delimita l'àrea d'interès i una capa de polígons amb un element "
    "per cep, i un script auxiliar els recorre tots de manera seqüencial. Treballar "
    "per blocs evita haver de carregar tota la parcel·la a memòria i permet repetir "
    "només la part afectada quan un bloc falla.")

# ===========================================================================
# (f) 4.6.1 Generació de màscares
# ===========================================================================
p = para("extreure les variables de cada cep")
set_text(p, (
    "Abans d'extreure les variables de cada cep cal separar, dins de cada capa, els "
    "píxels que corresponen al dosser dels que corresponen a l'entrefilera. La "
    "classificació es fa en dos passos encadenats: el primer és geomètric i el "
    "segon, espectral (Figura 15)."))
el = p._element

p = para("separa el dosser del sòl a partir del model digital")
set_text(p, (
    "El primer pas utilitza el model digital de vegetació. Abans d'aplicar-hi cap "
    "llindar se'n descarten els valors atípics amb el criteri del rang "
    "interquartílic, que elimina els píxels situats per damunt i per sota de les "
    "tanques calculades a partir del primer i del tercer quartil; així es netegen "
    "els pics espuris que arrossega el núvol de punts. Tot seguit la capa es parteix "
    "amb un llindar d'alçada de 0,5 m: els píxels que arriben o superen aquesta "
    "alçada formen la classe de dosser i la resta, la d'entrefilera."))
el = p._element
el = body(el,
    "El segon pas hi afegeix la informació espectral, perquè no tot el que és alt és "
    "vegetació activa ni tot el que és baix és sòl nu. Dins de la classe de dosser "
    "només es conserven els píxels amb un SAVI igual o superior a 0,1, que són els "
    "que corresponen a fulla fotosintèticament activa i deixen fora la fusta, les "
    "ombres i els píxels de vora barrejats amb el carrer. Dins de la classe "
    "d'entrefilera s'aplica el llindar complementari sobre el mateix índex. Les dues "
    "màscares resultants són les que s'apliquen a la resta de capes —índexs de "
    "vegetació, alçada i temperatura— abans de calcular cap estadístic.")
el = picture(el, "fig15_classificacio.png", 13.5)
el = caption(el, "Figura 15.",
             "Classificació dels píxels de cada vol en dosser i entrefilera. El "
             "llindar d'alçada de 0,5 m sobre el model digital de vegetació separa "
             "les dues classes i el llindar de SAVI de 0,1 descarta els píxels "
             "barrejats; sobre les màscares resultants es calculen els estadístics "
             "zonals de cada cep. Font: elaboració pròpia a partir de l'script de "
             "processament de l'IRTA.")
el = note(el,
    "A l'script, la màscara de l'entrefilera s'anomena SOIL_GREEN i les variables "
    "que se'n deriven porten l'etiqueta covercrop, però el llindar que l'estableix "
    "conserva els píxels amb un SAVI inferior a 0,1, que són precisament els de sòl "
    "sense vegetació activa. O bé el nom no descriu el que selecciona el llindar, o "
    "bé la comparació hauria de ser la contrària si el que es vol quantificar és la "
    "coberta vegetal del carrer. Cal aclarir-ho abans d'interpretar la fracció de "
    "coberta de l'entrefilera.")

# ===========================================================================
# (g) 4.6.2 Paràmetres estructurals + Taula 11
# ===========================================================================
p = para("Creuant la quadrícula amb les capes")
set_text(p, (
    "Creuant aquesta capa amb les màscares s'obtenen, cep a cep, els estadístics "
    "zonals que després es combinen en els paràmetres estructurals del dosser. De "
    "la capa d'alçada se'n pren el valor màxim dins del polígon, que és l'alçada del "
    "dosser, i també la suma de les alçades de tots els píxels de dosser, que "
    "multiplicada per la superfície del píxel dona el volum de copa. El recompte de "
    "píxels de dosser, multiplicat per aquesta mateixa superfície, dona l'àrea de "
    "copa. Dividint l'àrea i el volum pel marc de plantació s'obtenen la fracció de "
    "coberta i el volum per unitat de superfície de sòl, que són les magnituds que "
    "entren al model de balanç energètic. De l'àrea de copa encara se'n dedueixen el "
    "radi i l'amplada equivalents del dosser, suposant-lo circular en planta, i la "
    "relació entre amplada i alçada, que descriu la forma de la paret vegetal. La "
    "Taula 11 en recull la definició."))
el = p._element
el = caption(el, "Taula 11.",
             "Paràmetres estructurals del dosser calculats per a cada cep a partir "
             "de les màscares i de la capa de polígons. El marc de plantació "
             "s'introdueix al processament com un atribut de cada cep.")
el = table(el, [
    ("Variable", "Definició", "Unitats"),
    ("hc", "Valor màxim de l'alçada del dosser dins del polígon del cep", "m"),
    ("Àrea de copa", "Nombre de píxels de dosser × superfície del píxel", "m²"),
    ("Volum de copa", "Suma de les alçades dels píxels de dosser × superfície del píxel", "m³"),
    ("fc", "Àrea de copa / marc de plantació", "–"),
    ("fVolc", "Volum de copa / marc de plantació", "m"),
    ("rc", "Radi equivalent de la copa, √(àrea de copa / π)", "m"),
    ("wc", "Amplada equivalent de la copa, 2 · rc", "m"),
    ("wc / hc", "Relació entre amplada i alçada del dosser", "–"),
    ("Àrea de l'entrefilera", "Nombre de píxels d'entrefilera × superfície del píxel", "m²"),
    ("fc de l'entrefilera", "Àrea de l'entrefilera / marc de plantació", "–"),
], [Cm(3.6), Cm(9.6), Cm(2.2)], align_left_col0=True)

# ===========================================================================
# (h) 4.6.3 Índexs espectrals + Taula 12
# ===========================================================================
p = para("A partir de les bandes calibrades es calculen els")
set_text(p, (
    "A partir de les bandes calibrades es calculen cinc índexs de vegetació que "
    "descriuen aspectes complementaris del dosser (Taula 12). El NDVI relaciona les "
    "bandes de l'infraroig proper i del vermell i pren valors entre −1 i 1: els "
    "valors alts indiquen vegetació densa i activa, els propers a zero poca "
    "vegetació o sòl descobert i els negatius, absència de vegetació (Tucker, 1979). "
    "El SAVI n'és una variant que incorpora un factor de correcció del sòl, amb "
    "L = 0,5, i dona una estimació més fiable quan la coberta vegetal és escassa o "
    "el sòl queda a la vista, cosa que el fa especialment adequat en un cultiu en "
    "files com la vinya (Huete, 1988); és també l'índex amb què es construeixen les "
    "màscares de l'apartat 4.6.1."))
el = p._element

p = para("infraroig proper correspon a la banda 5")
set_text(p, (
    "Els altres tres completen la descripció del dosser. El MTVI és poc sensible a "
    "la variació de clorofil·la i respon sobretot a la quantitat de fulla, de manera "
    "que s'utilitza com a predictor del LAI (Haboudane et al., 2004). El NDWI, "
    "calculat aquí amb les bandes del verd i de l'infraroig proper, respon al "
    "contingut d'aigua de la coberta (McFeeters, 1996). I el CIG és un índex de "
    "clorofil·la basat en la relació entre l'infraroig proper i el verd (Gitelson "
    "et al., 2005). Cap d'ells no es calcula amb la calculadora ràster sinó dins de "
    "la mateixa llibreria de processament, i de cadascun se'n desa la mitjana i la "
    "desviació estàndard dins del polígon de cada cep, tant per als píxels de dosser "
    "com per als d'entrefilera."))
el = p._element
el = caption(el, "Taula 12.",
             "Índexs de vegetació calculats per la llibreria de processament. V, G, "
             "VR i IRP són la reflectància de les bandes del vermell (668 nm), del "
             "verd (560 nm), de la vora del vermell (717 nm) i de l'infraroig proper "
             "(840 nm) de la càmera Altum-PT.")
el = table(el, [
    ("Índex", "Expressió", "Bandes (nm)"),
    ("NDVI", "(IRP − V) / (IRP + V)", "840, 668"),
    ("SAVI", "1,5 · (IRP − V) / (IRP + V + 0,5)", "840, 668"),
    ("MTVI", "1,5 · [1,2 · (IRP − G) − 2,5 · (VR − G)] / "
             "√[(2 · IRP + 1)² − (6 · IRP − 5 · √VR) − 0,5]", "840, 560, 717"),
    ("NDWI", "(G − IRP) / (G + IRP)", "560, 840"),
    ("CIG", "(IRP / G) − 1", "840, 560"),
], [Cm(2.2), Cm(10.4), Cm(2.8)], align_left_col0=True)
el = note(el,
    "L'expressió del MTVI implementada a la llibreria utilitza la banda de la vora "
    "del vermell allà on la formulació original del MTVI2 de Haboudane et al. (2004) "
    "utilitza la del vermell. Cal decidir si es documenta com una variant adaptada a "
    "l'Altum-PT, i aleshores no se li hauria de dir MTVI2, o si es recupera la "
    "formulació original abans d'ajustar el model de LAI.")

# ===========================================================================
# (i) 4.6.4 Temperatures de component i fitxers de sortida
# ===========================================================================
h = after(el, 'Heading 3')
_font(h.add_run("4.6.4. Temperatures de component i fitxers de sortida"), 13, bold=True)
el = h._element
el = body(el,
    "El model de balanç energètic necessita la temperatura del dosser i la del sòl "
    "per separat, i totes dues surten de la mateixa capa tèrmica calibrada aplicant-hi "
    "les màscares de l'apartat 4.6.1. L'estadístic, però, no és una mitjana simple: "
    "abans de calcular-la es retallen els extrems de la distribució de cada polígon, "
    "perquè els píxels més freds i més calents solen ser barreges de dosser i sòl o "
    "artefactes de vora.")
el = body(el,
    "La temperatura de superfície del conjunt de l'escena i la del sòl es calculen "
    "com la mitjana dels píxels compresos entre els percentils 3 i 97 de cada "
    "polígon. La temperatura de dosser, en canvi, es calcula entre els percentils 3 "
    "i 25, és a dir sobre el quartil més fred dels píxels de dosser. El criteri busca "
    "quedar-se amb els píxels de fulla pura, que són els més freds perquè transpiren, "
    "i evitar la contaminació dels píxels barrejats amb el sòl de l'entrefilera, molt "
    "més calent. La temperatura de dosser i la de sòl són les dues entrades tèrmiques "
    "del model TSEB-2T descrit a l'apartat 4.8.")
el = body(el,
    "El resultat del processament de cada bloc és un fitxer de text delimitat amb una "
    "fila per cep i per data, que recull l'identificador del cep amb el tractament i "
    "el bloc als quals pertany, els estadístics dels cinc índexs per a les dues "
    "classes de píxel, els paràmetres estructurals de la Taula 11 i les tres "
    "temperatures. Concatenats, aquests fitxers formen la taula de treball sobre la "
    "qual s'ajusten el model de Random Forest per al LAI i el model TSEB-2T. De cada "
    "bloc se'n desen també les capes ràster intermèdies —temperatura de superfície i "
    "de sòl, NDVI, alçada i les màscares de dosser i d'entrefilera—, que permeten "
    "comprovar visualment que la classificació ha funcionat abans de donar les dades "
    "per bones.")

# ===========================================================================
# (j) Referències noves
# ===========================================================================
REFS = [
    ("Fereres", "Gitelson, A. A., Viña, A., Ciganda, V., Rundquist, D. C., & "
     "Arkebauer, T. J. (2005). Remote estimation of canopy chlorophyll content in "
     "crops. Geophysical Research Letters, 32(8), L08403. "
     "https://doi.org/10.1029/2005GL022688"),
    ("Gitelson, A. A.", "Haboudane, D., Miller, J. R., Pattey, E., Zarco-Tejada, "
     "P. J., & Strachan, I. B. (2004). Hyperspectral vegetation indices and novel "
     "algorithms for predicting green LAI of crop canopies: Modeling and validation "
     "in the context of precision agriculture. Remote Sensing of Environment, 90(3), "
     "337–352. https://doi.org/10.1016/j.rse.2003.12.013"),
    ("Haboudane, D.", "Huete, A. R. (1988). A soil-adjusted vegetation index (SAVI). "
     "Remote Sensing of Environment, 25(3), 295–309. "
     "https://doi.org/10.1016/0034-4257(88)90106-X"),
    ("Kustas, W. P., & Norman", "McFeeters, S. K. (1996). The use of the Normalized "
     "Difference Water Index (NDWI) in the delineation of open water features. "
     "International Journal of Remote Sensing, 17(7), 1425–1432. "
     "https://doi.org/10.1080/01431169608948714"),
    ("Shuttleworth, W. J.", "Tucker, C. J. (1979). Red and photographic infrared "
     "linear combinations for monitoring vegetation. Remote Sensing of Environment, "
     "8(2), 127–150. https://doi.org/10.1016/0034-4257(79)90013-0"),
]
for anchor_txt, text in REFS:
    src = para(anchor_txt)
    p = after(src._element)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.line_spacing_rule = src.paragraph_format.line_spacing_rule
    pf.left_indent = src.paragraph_format.left_indent
    pf.first_line_indent = src.paragraph_format.first_line_indent
    pf.space_after = src.paragraph_format.space_after
    _font(p.add_run(text), 12)

# ===========================================================================
# (k) Annex B amb l'script de paràmetres biofísics
# ===========================================================================
WRAP = {
"                      resampling=[True, {'band_resampling': ['HEIGHT'], 'band_base': 'NIR'}])":
["                      resampling=[True, {'band_resampling': ['HEIGHT'],",
 "                                         'band_base': 'NIR'}])"],
"                      resampling=[True, {'band_resampling': ['LST'], 'band_base': 'NIR'}])":
["                      resampling=[True, {'band_resampling': ['LST'],",
 "                                         'band_base': 'NIR'}])"],
"# Bloc a processar. executar_blocs.py l'envia des de fora; si l'executes sol, fa servir T0_0":
["# Bloc a processar. executar_blocs.py l'envia des de fora; si l'executes",
 "# sol, fa servir T0_0"],
"max_height_df = (ImageCollection.reduce_regions(band_name=['CANOPY_HEIGHT'], geometry=trees, reducer='max')":
["max_height_df = (ImageCollection.reduce_regions(band_name=['CANOPY_HEIGHT'],",
 "                                                geometry=trees, reducer='max')"],
"                 .rename(columns={'CANOPY_HEIGHT_max': 'hc'}))  # zonal stats by removing outlier pixels":
["                 .rename(columns={'CANOPY_HEIGHT_max': 'hc'}))  # zonal stats"],
"sum_height_df = (ImageCollection.reduce_regions(band_name=['CANOPY_HEIGHT'], geometry=trees, reducer='sum')":
["sum_height_df = (ImageCollection.reduce_regions(band_name=['CANOPY_HEIGHT'],",
 "                                                geometry=trees, reducer='sum')"],
"    ImageCollection.reduce_regions(['CANOPY_NDVI', 'CANOPY_SAVI', 'CANOPY_MTVI', 'CANOPY_NDWI', 'CANOPY_CIG'], trees,":
["    ImageCollection.reduce_regions(['CANOPY_NDVI', 'CANOPY_SAVI',",
 "                                    'CANOPY_MTVI', 'CANOPY_NDWI',",
 "                                    'CANOPY_CIG'], trees,"],
"    ['CANOPY_NDVI', 'CANOPY_SAVI', 'CANOPY_MTVI', 'CANOPY_NDWI', 'CANOPY_CIG'], trees, ['std'])":
["    ['CANOPY_NDVI', 'CANOPY_SAVI', 'CANOPY_MTVI', 'CANOPY_NDWI',",
 "     'CANOPY_CIG'], trees, ['std'])"],
"    ImageCollection.reduce_regions(['SOIL_NDVI', 'SOIL_SAVI', 'SOIL_MTVI', 'SOIL_NDWI', 'SOIL_CIG'], trees,":
["    ImageCollection.reduce_regions(['SOIL_NDVI', 'SOIL_SAVI', 'SOIL_MTVI',",
 "                                    'SOIL_NDWI', 'SOIL_CIG'], trees,"],
"# surface temperature is the temperature of the scene taking into account the canopy and soil":
["# surface temperature is the temperature of the scene taking into account",
 "# the canopy and soil"],
"surface_temperature_df = (ImageCollection.reduce_regions_percentile('LST', 'tree', trees, 'mean', pDown=3, pUp=97)":
["surface_temperature_df = (ImageCollection.reduce_regions_percentile(",
 "                              'LST', 'tree', trees, 'mean', pDown=3, pUp=97)"],
"    ImageCollection.reduce_regions_percentile('CANOPY_LST', 'tree', trees, 'mean', pDown=3, pUp=25)":
["    ImageCollection.reduce_regions_percentile('CANOPY_LST', 'tree', trees,",
 "                                              'mean', pDown=3, pUp=25)"],
"soil_temperature_df = (ImageCollection.reduce_regions_percentile('SOIL_LST', 'tree', trees, 'mean', pDown=3, pUp=97)":
["soil_temperature_df = (ImageCollection.reduce_regions_percentile(",
 "                           'SOIL_LST', 'tree', trees, 'mean', pDown=3, pUp=97)"],
"# = df_merge([max_height_df, sum_height_df, canopy_count_df, index_mean_df, index_std_df, surface_temperature_df,":
["# = df_merge([max_height_df, sum_height_df, canopy_count_df, index_mean_df,",
 "#             index_std_df, surface_temperature_df,"],
"info = pd.read_csv('20260603_YIELD4CAST_ZV_CV/info_vinesADAPTEX2026_ZV_CV.csv').astype({'tree': str})  # Add contextual info trees":
["# Add contextual info trees",
 "info = pd.read_csv('20260603_YIELD4CAST_ZV_CV/'",
 "                   'info_vinesADAPTEX2026_ZV_CV.csv').astype({'tree': str})"],
"dataset = dataset.drop(['sizex', 'sizey', 'sum_hc', 'index'], axis=1)      #drop del dataset anteriorment hi havia també 'area_pixel'":
["# drop del dataset; anteriorment hi havia també 'area_pixel'",
 "dataset = dataset.drop(['sizex', 'sizey', 'sum_hc', 'index'], axis=1)"],
"dataset.to_csv(f'20260603_YIELD4CAST_ZV_CV/outputs/biophysical_traits_YIELD4CAST2026_covercrop_ZV_CV_03062026_{TILE}.csv')":
["dataset.to_csv(f'20260603_YIELD4CAST_ZV_CV/outputs/'",
 "               f'biophysical_traits_YIELD4CAST2026_covercrop_ZV_CV_'",
 "               f'03062026_{TILE}.csv')"],
"ImageCollection.save_raster(output_folder=out_rasters, band_name=['LST','SOIL_LST', 'NDVI','HEIGHT','CANOPY','SOIL', 'CANOPY_NDVI' ])":
["ImageCollection.save_raster(output_folder=out_rasters,",
 "                            band_name=['LST', 'SOIL_LST', 'NDVI', 'HEIGHT',",
 "                                       'CANOPY', 'SOIL', 'CANOPY_NDVI'])"],
"ImageCollection.rename(['HEIGHT_MASK'], ['CANOPY'])  # rename previously generated band":
["ImageCollection.rename(['HEIGHT_MASK'], ['CANOPY'])  # rename the new band"],
"                   .rename(columns={'CANOPY_GREEN_count': 'CANOPY_COUNT'}))  # zonal stats":
["                   .rename(columns={'CANOPY_GREEN_count': 'CANOPY_COUNT'}))"],
"                   .rename(columns={'SOIL_GREEN_count': 'SOIL_COUNT'}))  # zonal stats":
["                   .rename(columns={'SOIL_GREEN_count': 'SOIL_COUNT'}))"],
"pixel_info = ImageCollection.resolution('HEIGHT')  # Get the pixel size of each images":
["pixel_info = ImageCollection.resolution('HEIGHT')  # pixel size of images"],
}

with open(SRC_BIO, encoding='utf-8') as fh:
    raw = [l.rstrip() for l in fh.read().split('\n')]
code = []
for ln in raw:
    code.extend(WRAP[ln] if ln in WRAP else [ln])
while code and not code[-1].strip():
    code.pop()
too_long = [l for l in code if len(l) > 92]
if too_long:
    print('AVÍS, línies massa llargues a l\'Annex B:')
    for l in too_long:
        print(f'   {len(l):3d} | {l}')

h = doc.add_paragraph(style='Heading 1')
_font(h.add_run("Annex B. Script d'extracció dels paràmetres biofísics del dosser"),
      14, bold=True)
h.paragraph_format.page_break_before = True

intro = doc.add_paragraph()
intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
_font(intro.add_run(
    "Es reprodueix l'script de processament corresponent al vol del 3 de juny de "
    "2026, tal com es descriu als apartats 4.6.1 a 4.6.4. S'executa un cop per bloc "
    "de la parcel·la, amb el codi del bloc com a única variable que canvia entre "
    "execucions, i s'recolza en la llibreria interna raster_ET4DROUGHT, que no es "
    "reprodueix aquí per extensió i que és la que implementa les classes de "
    "col·lecció d'imatges, el remostreig, l'emmascarament i els estadístics zonals. "
    "Les úniques modificacions respecte de l'original són el tall d'algunes línies "
    "massa llargues per a l'amplada de la pàgina; els comentaris i els noms de "
    "variable es mantenen tal com són al codi."), 12)

for line in code:
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = Pt(0)
    pf.left_indent = Cm(0.3)
    r = p.add_run(line if line else " ")
    _font(r, 8, name=MONO)
    r._element.find(qn('w:t')).set(qn('xml:space'), 'preserve')

# ===========================================================================
# (l) Recàlcul dels camps SEQ
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
      f"· Annex B: {len(code)} línies de codi")
