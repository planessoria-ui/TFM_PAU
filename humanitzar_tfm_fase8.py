# -*- coding: utf-8 -*-
"""
Fase 8: actualitza la metodologia segons l'esquema de treball del TSEB de
l'assaig de capçades.

  - Reescriu l'apartat 4.8 amb el flux de set passos i el model TSEB-2T amb la
    formulació de Campbell i Norman per a cultius en filera (hedgerow).
  - Treu l'apartat d'anàlisi estadística i les seves crides.
  - Limita el CWSI a un paper secundari de comprovació.
  - Afegeix les referències de Campbell i Norman (1998) i Nieto et al. (2019).
"""
import copy
import docx
from docx.oxml.ns import qn

SRC = "/root/.claude/uploads/446d5a3f-407d-5d51-8cea-50e8348ad898/a3823eaf-TFM_Pau_Planes_IRTA_v6.docx"
DST = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v7.docx"
doc = docx.Document(SRC)


def find(prefix, style=None):
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith(prefix) and (style is None or p.style.name == style):
            return i
    raise LookupError(prefix[:70])


def P(prefix, style=None):
    return doc.paragraphs[find(prefix, style)]


def set_text(p, new):
    runs = p.runs
    if not runs:
        p.add_run(new); return
    runs[0].text = new
    for r in runs[1:]:
        r._element.getparent().remove(r._element)


def insert_after(ref, text, template, style=None):
    el = copy.deepcopy(template._element)
    ref._element.addnext(el)
    p = docx.text.paragraph.Paragraph(el, ref._parent)
    set_text(p, text)
    if style:
        p.style = doc.styles[style]
    return p


def drop(p):
    p._element.getparent().remove(p._element)


# ---------------------------------------------------------------------------
# 1. Apartat 4.8: TSEB-2T i flux de treball
# ---------------------------------------------------------------------------
set_text(P("4.8. Models d'evapotranspiració", style="Heading 2"),
         "4.8. Estimació de l'evapotranspiració amb el model TSEB-2T")

p_intro = P("L'evapotranspiració real (ETa) de cada tractament es calcularà")
set_text(p_intro,
         "L'evapotranspiració real (ETa) de cada tractament s'estima amb el model de "
         "balanç energètic de dues fonts en la variant de dues temperatures (TSEB-2T), "
         "que pren la temperatura de dosser (Tc) i la de sòl (Ts) per separat en lloc de "
         "derivar-les d'una única temperatura radiomètrica (Norman et al., 1995; Nieto et "
         "al., 2019). La partició de la radiació dins del dosser es resol amb la "
         "formulació de Campbell i Norman adaptada a cultius en filera o hedgerow "
         "(Campbell i Norman, 1998), que és la que correspon a l'estructura en doble "
         "cordó de la parcel·la. La implementació es fa amb el paquet obert pyTSEB (Nieto "
         "i Kustas).")

p_lead = insert_after(p_intro,
                      "El procediment de treball consta de set passos:", p_intro)

# plantilla de llista numerada (dels objectius específics)
tmpl_num = P("Estimar les variables biofísiques del dosser")

passos = [
    "Obtenir la fIPAR diària a partir de les mesures de camp de fIPAR en els ceps "
    "defoliats i derivar-ne el LAI per dues vies: el model d'inversió del mateix "
    "ceptòmetre i la formulació de Campbell i Norman.",
    "Validar aquestes estimacions amb els ceps defoliats i quantificar-ne l'ajust amb "
    "el coeficient de determinació (R²) i l'error quadràtic mitjà (RMSE).",
    "Obtenir el LAI de la resta de ceps mesurats i corregir-lo amb l'error obtingut en "
    "la validació anterior.",
    "Crear una capa vectorial amb els cinc tractaments i les quatre repeticions per "
    "extreure la informació individual de cada cep.",
    "Estimar el LAI de tots els ceps de l'assaig amb un model de Random Forest alimentat "
    "amb la informació de les imatges (índexs de vegetació, alçada, fracció de coberta i "
    "volum de copa) i calibrat amb els ceps validats al pas anterior.",
    "Obtenir la temperatura de dosser (Tc) i la de sòl (Ts) discriminant de manera "
    "automàtica els píxels de vegetació i els de sòl a partir de l'alçada i dels índexs "
    "de vegetació.",
    "Executar el model TSEB-2T amb la formulació de Campbell i Norman per a cultius en "
    "filera i obtenir l'ETa i la partició entre transpiració i evaporació.",
]
ref = p_lead
for s in passos:
    ref = insert_after(ref, s, tmpl_num)

p_last = insert_after(ref,
    "[A VERIFICAR] A l'esquema de treball, la primera via d'inversió del LAI apareix "
    "anomenada «Jarvis». Confirmeu-ne el nom i la referència exactes per citar-la "
    "correctament.", p_intro)
for r in p_last.runs:
    from docx.shared import RGBColor, Pt
    r.font.color.rgb = RGBColor(0xB0, 0x00, 0x00); r.font.italic = True; r.font.size = Pt(11)

# CWSI limitat a un paper secundari
set_text(P("A partir d'aquestes dues estimacions"),
         "L'evapotranspiració potencial (ETp) s'estima amb el model de "
         "Shuttleworth–Wallace (Shuttleworth i Wallace, 1985) a partir de l'estructura "
         "real del dosser. De manera complementària es pot calcular l'índex d'estrès "
         "hídric del cultiu com CWSI = 1 − ETa/ETp (Idso et al., 1981; Jackson et al., "
         "1981). Com que el reg individualitzat manté tots els tractaments en un mateix "
         "estat hídric, aquest índex hi té un paper secundari i serveix només com a "
         "comprovació que l'estat hídric és equivalent; el seguiment de l'estat hídric "
         "recau en les mesures de potencial hídric de tija. La Taula 4 resumeix les "
         "entrades i les sortides principals de cada model.")

# ---------------------------------------------------------------------------
# 2. Treure l'apartat d'anàlisi estadística
# ---------------------------------------------------------------------------
drop(P("[A CONCRETAR] Programari estadístic"))
drop(P("L'efecte dels tractaments de maneig de capçada sobre les variables biofísiques"))
drop(P("4.9. Anàlisi estadística", style="Heading 2"))

set_text(P("El treball s"),
         P("El treball s").text
         .replace(", les mesures de camp, els models d'ET i l’anàlisi estadística",
                  ", les mesures de camp i els models d'ET")
         .replace(", les mesures de camp, els models d'ET i l'anàlisi estadística",
                  ", les mesures de camp i els models d'ET"))

set_text(P("Taules i figures de LAI, fIPAR, alçada"),
         "Taules i figures de LAI, fIPAR, alçada i volum del dosser per als cinc "
         "tractaments (SE, EL, EM, ES i ES+DV) i per a cada data de vol.")

# ---------------------------------------------------------------------------
# 3. Limitar el CWSI als objectius i als resultats
# ---------------------------------------------------------------------------
set_text(P("Comprovar, amb l'índex d'estrès hídric del cultiu"),
         "Comprovar, amb les mesures de potencial hídric de tija (Ψstem), que els cinc "
         "tractaments es mantenen en un mateix estat hídric, de manera que les "
         "diferències d'ETa es puguin atribuir al maneig de capçada.")

set_text(P("5.4.", style="Heading 2"), "5.4. Estat hídric dels tractaments")

set_text(P("CWSI per tractament i regressió CWSI"),
         "Evolució del potencial hídric de tija per tractament i comprovació que els "
         "valors són equivalents entre tractaments, cosa que confirmaria que l'estat "
         "hídric s'ha igualat. Si es calcula el CWSI, presenteu-lo com a comprovació "
         "complementària.")

# ---------------------------------------------------------------------------
# 4. Abreviatures i referències noves
# ---------------------------------------------------------------------------
p_tseb = P("TSEB: Balanç energètic de dues fonts")
insert_after(p_tseb,
             "TSEB-2T: Variant del TSEB que utilitza les temperatures de dosser i de sòl "
             "per separat (two-temperature)", p_tseb)
p_fipar = P("fIPAR:")
insert_after(p_fipar, "fc: Fracció de coberta vegetal", p_fipar)

p_chone = P("Choné, X., Van Leeuwen")
el = copy.deepcopy(p_chone._element)
p_chone._element.addprevious(el)
set_text(docx.text.paragraph.Paragraph(el, p_chone._parent),
         "Campbell, G. S., & Norman, J. M. (1998). An introduction to environmental "
         "biophysics (2a ed.). Springer.")

p_norman = P("Norman, J. M., Kustas, W. P., & Humes")
el = copy.deepcopy(p_norman._element)
p_norman._element.addprevious(el)
set_text(docx.text.paragraph.Paragraph(el, p_norman._parent),
         "Nieto, H., Kustas, W. P., Torres-Rúa, A., Alfieri, J. G., Gao, F., Anderson, M. "
         "C., White, W. A., Song, L., Alsina, M. del M., Prueger, J. H., McKee, M., "
         "Elarab, M., & McKee, L. G. (2019). Evaluation of TSEB turbulent fluxes using "
         "different methods for the retrieval of soil and canopy component temperatures "
         "from UAV thermal and multispectral imagery. Irrigation Science, 37(3), 389–406. "
         "https://doi.org/10.1007/s00271-018-0585-9")

doc.save(DST)
print("Fase 8 completada ->", DST)
