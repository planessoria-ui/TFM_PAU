# -*- coding: utf-8 -*-
"""
Fase 34: omple el capítol 6 de discussió i el capítol 7 de conclusions, i
incorpora la referència de Picón-Toro et al. (2012) a la bibliografia.
"""
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
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


def H(n, ini):
    for p in doc.paragraphs:
        if p.style.name == f'Heading {n}' and p.text.strip().startswith(ini):
            return p
    raise LookupError(ini)


def after(el, style=None):
    new = OxmlElement('w:p')
    el.addnext(new)
    p = docx.text.paragraph.Paragraph(new, _PARENT)
    if style:
        p.style = doc.styles[style]
    return p


def body(el, text, color=None):
    p = after(el)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    p.paragraph_format.space_after = Pt(6)
    _font(p.add_run(text), 12, color=color)
    return p._element


def head(el, n, text):
    p = after(el, f'Heading {n}')
    _font(p.add_run(text), 13 if n == 2 else 12, bold=True)
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


def num(el, text):
    p = after(el, 'List Number')
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    p.paragraph_format.space_after = Pt(6)
    _font(p.add_run(text), 12)
    return p._element


# ===========================================================================
# 6. Discussió
# ===========================================================================
el = H(1, '6.')._element

el = head(el, 2, "6.1. El coeficient de cultiu en relació amb la bibliografia")
el = body(el,
    "Picón-Toro et al. (2012) van determinar el coeficient de cultiu basal d'una "
    "vinya de Tempranillo amb un lisímetre de pesada durant cinc anys a Badajoz i "
    "el van relacionar amb la mida de la capçada. Per a condicions de sòl sec i "
    "sense estrès hídric van obtenir Kcb = 0,01 · fIPAR(%) + 0,23, amb un "
    "coeficient de determinació de 0,68 sobre 54 observacions, i Kcb = 0,22 · LAI + "
    "0,13, amb 0,81 sobre 36. És la relació que permet traslladar el seu coeficient "
    "a vinyes amb una capçada diferent de la seva.")

el = body(el,
    "Aplicant la primera equació a la fIPAR dels 472 registres d'aquest assaig "
    "s'obté un Kcb mitjà de 0,604, davant del 0,635 que donen les mesures. El biaix "
    "és de +0,031 i l'error quadràtic mitjà, de 0,075. Per vol, el biaix val +0,010 "
    "el 30 de juny, −0,002 el 23 de juliol i −0,004 el 12 d'agost, i només se'n "
    "surt el vol del 3 de juny, amb +0,119. Que una relació ajustada amb lisímetre "
    "a l'oest d'Extremadura reprodueixi el consum mesurat amb balanç d'energia a "
    "Raimat amb aquest marge és una coincidència notable, tenint en compte que els "
    "marcs de plantació són diferents, 3,0 m² per cep allà i 4,25 aquí.")

el = body(el,
    "La coincidència, però, no s'ha de llegir com una equivalència directa. El Kcb "
    "de l'article és basal i sense estrès, mentre que el coeficient d'aquest treball "
    "és actual i inclou tant la reducció per dèficit hídric com l'evaporació del "
    "sòl. En la descomposició de la FAO-56, Kc = Ks · Kcb + Ke, amb un Ks mitjà de "
    "0,785 la reducció per estrès val −0,130 i el terme d'evaporació del sòl, "
    "+0,161. Els dos termes són de mida semblant i de signe contrari, i el +0,031 "
    "que s'observa n'és el residu. Això explica també el comportament del 3 de "
    "juny, que és el vol amb menys estrès i, per tant, aquell en què la compensació "
    "encara no s'ha produït.")

el = body(el,
    "La comparació que sí que és directa és amb la component de transpiració. "
    "Ajustant el quocient entre la transpiració potencial de Shuttleworth i Wallace "
    "i l'ETo contra la fIPAR s'obté un pendent de 0,0098 per unitat de percentatge, "
    "davant del 0,0100 de Picón-Toro et al. (2012), amb una correlació de 0,902 "
    "entre les dues sèries. El pendent coincideix amb un 2 % de diferència i el que "
    "canvia és l'ordenada a l'origen, 0,11 superior en aquest treball, que és la "
    "part atribuïble a l'evaporació de l'entrefilera. La sensibilitat del consum "
    "d'aigua a la radiació interceptada és, doncs, la mateixa a les dues vinyes.")

el = body(el,
    "Cal situar l'assaig dins del rang de l'article. La vinya de Badajoz arriba a un "
    "LAI de 4,3 i a una fIPAR del 71 %, amb el Kcb estabilitzat a la unitat per "
    "damunt d'un LAI de 3. Aquest assaig es mou entre un LAI d'1,61 i 2,53 i un Kcb "
    "estimat d'entre 0,43 i 0,89, de manera que ocupa el tram ascendent de la seva "
    "corba i només el tractament sense esporga del 12 d'agost s'acosta a la zona de "
    "plateau. És també el tram on la regressió està més ben determinada. Convé "
    "remarcar que l'experiment de Picón-Toro et al. (2012), com aquest, compara "
    "capçades de mida diferent sense estrès hídric, i és aquesta coincidència de "
    "condicions la que fa legítima la comparació.")

el = head(el, 2, "6.2. Estructura i funció del dosser")
el = body(el,
    "La fracció de coberta i l'índex d'àrea foliar ordenen els tractaments de la "
    "mateixa manera als quatre vols, amb coeficients de correlació d'entre 0,93 i "
    "0,99 a escala de tractament, tot i provenir de dues mesures independents: la "
    "classificació de les imatges de dron i les lectures de ceptòmetre sota el "
    "dosser. La relació entre totes dues, però, no és proporcional.")

el = body(el,
    "El quocient entre l'índex d'àrea foliar i la fracció de coberta expressa els "
    "metres quadrats de fulla per metre quadrat de sòl cobert, és a dir la densitat "
    "de la paret vegetal. Val entre 9,9 i 10,2 als dos tractaments d'esporga severa "
    "i entre 6,3 i 6,8 al tractament sense esporga (Figura 30). L'esporga severa no "
    "elimina fullatge en proporció a la petjada que retalla, sinó que el concentra "
    "en una paret més estreta: la petjada del tractament ES+DV és el 52 % de la del "
    "control mentre que la seva àrea foliar n'és el 82 %.")

el = picture(el, "fig_densitat_paret.png", 14.5)
el = caption(el, "Figura 30.",
             "Densitat de la paret vegetal, expressada com el quocient entre "
             "l'índex d'àrea foliar i la fracció de coberta, per tractament i data "
             "de vol. Font: elaboració pròpia.")

el = body(el,
    "La poda també canvia la manera com creix la capçada. Entre el primer vol i "
    "l'últim, el tractament ES+DV guanya un 17,6 % d'àrea foliar amb la fracció de "
    "coberta pràcticament igual, de manera que la seva densitat passa de 8,67 a "
    "10,12. L'esporga lleugera, en canvi, guanya un 50,4 % de fracció de coberta i "
    "només un 23,1 % d'àrea foliar, i la seva densitat baixa de 8,02 a 6,57. A "
    "l'esporga severa la fulla s'acumula dins d'un volum fix i la paret s'espesseix; "
    "a l'esporga lleugera la capçada s'escampa i es dilueix. Com més densa és la "
    "paret, més autoombrejament hi ha i menys aporta cada fulla afegida, cosa que "
    "ajuda a entendre per què el consum d'aigua no escala amb la petjada.")

el = body(el,
    "La densitat toca el mínim de la campanya el 23 de juliol als tractaments EM, "
    "ELL i C, que és exactament la data en què el coeficient de cultiu també és "
    "mínim. La capçada és aleshores la més estesa i la menys densa, en coincidència "
    "amb el potencial de tija més baix dels tres primers vols.")

el = head(el, 2, "6.3. El coeficient de reducció del reg")
el = body(el,
    "El reg de cada tractament es va escalar amb un coeficient proporcional a la "
    "fracció de coberta, de 0,70 als dos tractaments d'esporga severa fins a 1,40 "
    "al tractament sense esporga. Les dades d'evapotranspiració mostren que el "
    "consum no escala linealment amb la fracció de coberta: la regressió de l'ETa "
    "relativa sobre el coeficient de reg té un pendent de 0,47, amb una component "
    "independent de la capçada propera al 53 % del consum del tractament de "
    "referència.")

el = body(el,
    "En conseqüència, i en termes relatius, els dos tractaments d'esporga severa "
    "van rebre entre un 22 i un 23 % menys aigua de la que consumien, i el "
    "tractament sense esporga un 16 % més (Figura 31). El desajust es repeteix als "
    "quatre vols. Si l'objectiu és igualar l'estat hídric entre capçades de mida "
    "diferent, el coeficient de reducció no hauria de ser la fracció de coberta "
    "sinó la relació de consum, que amb aquestes dades val 0,86, 0,86, 1,00, 1,09 i "
    "1,17 en lloc de 0,70, 0,70, 1,00, 1,10 i 1,40.")

el = picture(el, "fig_reg_vs_consum.png", 16.0)
el = caption(el, "Figura 31.",
             "Dosi de reg, fracció de coberta, àrea foliar i consum d'aigua de cada "
             "tractament, expressats en relació amb el tractament sense esporga. "
             "Mitjanes de la campanya de 2026. Font: elaboració pròpia.")

el = body(el,
    "Val la pena comprovar si l'aigua addicional del tractament sense esporga "
    "n'infla l'ETa per evaporació del sòl. Si fos així, el biaix respecte del Kcb "
    "de sòl sec de Picón-Toro et al. (2012) hauria de créixer amb la dosi, i no ho "
    "fa: val +0,043, +0,018, +0,033, +0,023 i +0,038 de l'esporga més severa a la "
    "més lleugera, sense cap tendència. Les diferències de consum entre tractaments "
    "són, doncs, de capçada i no de sòl humit.")

el = head(el, 2, "6.4. Limitacions i fonts d'incertesa")
el = body(el,
    "La principal limitació afecta la via tèrmica. Als quatre fitxers d'entrada del "
    "model la temperatura de dosser és un valor únic per vol, el mateix per als 118 "
    "ceps, mentre que l'script de processament la calcula polígon a polígon. Amb una "
    "sola temperatura de dosser, les diferències d'evapotranspiració entre "
    "tractaments només poden venir de la geometria de la capçada i de l'àrea foliar. "
    "L'acord amb Picón-Toro et al. (2012), que és una relació purament geomètrica, "
    "valida la part estructural del treball, però no en valida la part tèrmica, que "
    "queda pendent de resoldre.", color=VERMELL)

el = body(el,
    "El vol del 3 de juny no és comparable amb els altres tres quant als paràmetres "
    "estructurals, pels motius que exposa l'apartat 5.1. En segon lloc, la fracció "
    "de coberta dels fitxers de sortida del processament i la dels fitxers d'entrada "
    "del model no coincideixen, amb diferències que creixen fins al 36 % en el "
    "tractament ES+DV del 12 d'agost. En tercer lloc, el vol del 3 de juny no té "
    "fIPAR pròpia en set de les vint parcel·les, que comparteixen només dos valors, "
    "i els vols del 23 de juliol i del 12 d'agost repeteixen el mateix valor de "
    "fIPAR en dues parcel·les de tractaments diferents. Finalment, l'índex d'estrès "
    "hídric derivat del balanç d'energia no es pot fer servir per comparar "
    "tractaments mentre l'ETp no respongui a la capçada.", color=VERMELL)

el = body(el,
    "Tots els resultats corresponen a una sola campanya i a una sola parcel·la, amb "
    "quatre repeticions per tractament i quatre dates de vol. Les diferències entre "
    "tractaments que s'hi descriuen s'han de confirmar amb més anys abans de "
    "generalitzar-les.")

# ===========================================================================
# 7. Conclusions
# ===========================================================================
el = H(1, '7.')._element
el = body(el,
    "De l'anàlisi dels quatre vols de la campanya de 2026 se'n deriven les "
    "conclusions següents.")

for t in [
    "El maneig de capçada modifica el consum d'aigua de la vinya de manera "
    "substancial. Respecte del tractament sense esporga, l'esporga lleugera "
    "redueix l'evapotranspiració real un 6,7 %, l'esporga mitjana un 14,7 % i "
    "l'esporga severa un 26,4 %.",

    "El desfullat aplicat a l'envero no redueix el consum per damunt del que ja fa "
    "l'esporga severa. Els tractaments ES i ES+DV consumeixen 3,39 i 3,37 mm d⁻¹ "
    "de mitjana, una diferència que no és distingible.",

    "Les diferències de consum són atribuïbles a la capçada i no a l'estat hídric. "
    "El potencial de tija separa els tractaments entre 0,06 i 0,16 MPa dins de cada "
    "data, davant dels 0,55 MPa que separen el primer vol de l'últim, de manera que "
    "el reg diferenciat va aconseguir mantenir-los al mateix estat hídric.",

    "El coeficient de cultiu es pot estimar a partir de la fracció de radiació "
    "interceptada. La relació és molt estreta dins de cada vol, amb coeficients de "
    "determinació d'entre 0,95 i 0,996, i l'ajust conjunt dels quatre vols dona "
    "Kc = 0,708 · fIPAR + 0,370 amb 0,688.",

    "L'estimació del coeficient de cultiu basal a partir de la relació de "
    "Picón-Toro et al. (2012) reprodueix el consum mesurat en tres dels quatre "
    "vols, amb un biaix inferior a 0,01, i el pendent de la relació amb la fIPAR "
    "coincideix amb el d'aquell treball dins d'un 2 %. La transferència de "
    "coeficients de cultiu entre vinyes de Tempranillo amb capçades diferents és, "
    "per tant, viable sempre que es faci sobre la component de transpiració.",

    "L'esporga severa concentra el fullatge en lloc d'eliminar-lo en proporció. La "
    "densitat de la paret vegetal passa de 6,3 m² de fulla per m² de sòl cobert al "
    "tractament sense esporga a més de 10 als tractaments severs, i és aquesta "
    "concentració la que explica que el consum no escali amb la fracció de coberta.",

    "Com a conseqüència pràctica, el coeficient amb què es redueix la dosi de reg "
    "en vinyes amb capçades de mida diferent no hauria de ser proporcional a la "
    "fracció de coberta. Escalar el reg amb la fracció de coberta va deixar els "
    "tractaments d'esporga severa amb un 22 % menys aigua de la que consumien. La "
    "relació de consum, que en aquest assaig val 0,86, 0,86, 1,00, 1,09 i 1,17, és "
    "un criteri més ajustat.",
]:
    el = num(el, t)

# ===========================================================================
# 8. Referència
# ===========================================================================
anc = P("Norman, J. M., Kustas, W. P., & Humes, K. S. (1995)")
p = after(anc._element)
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
p.paragraph_format.space_after = Pt(6)
pf = p.paragraph_format
pf.first_line_indent = Cm(-1.25)
pf.left_indent = Cm(1.25)
_font(p.add_run(
    "Picón-Toro, J., González-Dugo, V., Uriarte, D., Mancha, L. A., & Testi, L. "
    "(2012). Effects of canopy size and water stress over the crop coefficient of "
    "a “Tempranillo” vineyard in south-western Spain. "), 12)
_font(p.add_run("Irrigation Science, 30"), 12, italic=True)
_font(p.add_run(", 419–432. https://doi.org/10.1007/s00271-012-0351-3"), 12)

doc.save(PATH)
print(f"OK · paràgrafs: {len(doc.paragraphs)} · taules: {len(doc.tables)}")
