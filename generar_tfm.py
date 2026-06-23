# -*- coding: utf-8 -*-
"""
Generador del document Word (.docx) del Treball Final de Màster:
"Avaluació de l'evapotranspiració en vinya sota diferents tractaments de poda
mitjançant teledetecció d'alta precisió amb drons i mesures de radiació PAR"

Autor: Pau Planes Soria — MENAG-i (ETSEA-UdL) en col·laboració amb l'IRTA.

Format: Arial 12 pt, interlineat 1,5, text justificat, marges 2,5 cm,
títols jeràrquics, numeració de pàgina i bibliografia APA 7a.

Aquest script construeix el document de forma reproduïble. El contingut es va
ampliant pas a pas (PAS 2: front matter + Introducció + Marc teòric).
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = "Arial"

# ---------------------------------------------------------------------------
# Utilitats de format
# ---------------------------------------------------------------------------

def set_cell_font(doc):
    pass


def _set_run_font(run, size=12, bold=False, italic=False, color=None):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    # Assegurar font també per a scripts complexos
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:ascii'), FONT)
    rfonts.set(qn('w:hAnsi'), FONT)
    rfonts.set(qn('w:cs'), FONT)


def setup_styles(doc):
    # Estil Normal: Arial 12, 1.5, justificat
    normal = doc.styles['Normal']
    normal.font.name = FONT
    normal.font.size = Pt(12)
    pf = normal.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_after = Pt(6)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Encapçalaments
    specs = {'Heading 1': (14, True), 'Heading 2': (13, True), 'Heading 3': (12, True)}
    for name, (sz, bold) in specs.items():
        st = doc.styles[name]
        st.font.name = FONT
        st.font.size = Pt(sz)
        st.font.bold = bold
        st.font.color.rgb = RGBColor(0, 0, 0)
        st.paragraph_format.space_before = Pt(14)
        st.paragraph_format.space_after = Pt(8)
        st.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

    # Marges 2,5 cm
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)


def add_page_number_footer(doc):
    section = doc.sections[-1]
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = 'PAGE'
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
    run._element.append(fldChar1); run._element.append(instr); run._element.append(fldChar2)
    _set_run_font(run, size=10)


def para(doc, text, style=None, align=None, size=12, bold=False, italic=False,
         space_after=6, space_before=0):
    p = doc.add_paragraph(style=style)
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    if text:
        run = p.add_run(text)
        _set_run_font(run, size=size, bold=bold, italic=italic)
    return p


def rich(doc, segments, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=12, space_after=6):
    """segments: llista de (text, {bold,italic}) per a citacions en cursiva, etc."""
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    for text, opts in segments:
        run = p.add_run(text)
        _set_run_font(run, size=size, bold=opts.get('bold', False),
                      italic=opts.get('italic', False))
    return p


def heading(doc, text, level):
    h = doc.add_heading(level=level)
    run = h.add_run(text)
    sz = {1: 14, 2: 13, 3: 12}[level]
    _set_run_font(run, size=sz, bold=True)
    h.paragraph_format.space_before = Pt(16 if level == 1 else 12)
    h.paragraph_format.space_after = Pt(8)
    return h


def add_toc(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    fldChar = OxmlElement('w:fldChar'); fldChar.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve')
    instr.text = 'TOC \\o "1-3" \\h \\z \\u'
    fldSep = OxmlElement('w:fldChar'); fldSep.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t'); t.text = "Feu clic dret aquí i trieu «Actualitza els camps» per generar l'índex."
    fldEnd = OxmlElement('w:fldChar'); fldEnd.set(qn('w:fldCharType'), 'end')
    run._element.append(fldChar); run._element.append(instr)
    run._element.append(fldSep); run._element.append(t); run._element.append(fldEnd)


def placeholder(doc, text):
    p = para(doc, text, italic=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT)
    for run in p.runs:
        run.font.color.rgb = RGBColor(0xB0, 0x00, 0x00)
    return p


# ---------------------------------------------------------------------------
# Construcció del document
# ---------------------------------------------------------------------------

def build():
    doc = Document()
    setup_styles(doc)

    # ---------------- PORTADA ----------------
    para(doc, "Escola Tècnica Superior d'Enginyeria Agroalimentària i Forestal i de Veterinària",
         align=WD_ALIGN_PARAGRAPH.CENTER, size=13, bold=True, space_before=24)
    para(doc, "Universitat de Lleida (UdL)", align=WD_ALIGN_PARAGRAPH.CENTER, size=12)
    para(doc, "", space_after=24)
    para(doc, "TREBALL FINAL DE MÀSTER", align=WD_ALIGN_PARAGRAPH.CENTER, size=16, bold=True)
    para(doc, "Màster en Enginyeria Agronòmica interuniversitari (MENAG-i)",
         align=WD_ALIGN_PARAGRAPH.CENTER, size=12)
    para(doc, "", space_after=36)
    para(doc, "Avaluació de l'evapotranspiració en vinya sota diferents tractaments "
              "de poda mitjançant teledetecció d'alta precisió amb drons i mesures "
              "de radiació PAR",
         align=WD_ALIGN_PARAGRAPH.CENTER, size=16, bold=True, space_after=48)
    para(doc, "Autor: Pau Planes Soria", align=WD_ALIGN_PARAGRAPH.CENTER, size=12)
    para(doc, "Tutor: José Antonio Martínez Casasnovas (UdL)",
         align=WD_ALIGN_PARAGRAPH.CENTER, size=12)
    para(doc, "Co-tutor: Joaquim Bellvert (IRTA)", align=WD_ALIGN_PARAGRAPH.CENTER, size=12)
    para(doc, "", space_after=24)
    para(doc, "Treball desenvolupat en col·laboració amb l'IRTA (Institut de Recerca "
              "i Tecnologia Agroalimentàries) — Programa d'Ús Eficient de l'Aigua",
         align=WD_ALIGN_PARAGRAPH.CENTER, size=11, italic=True)
    para(doc, "Curs 2025-26", align=WD_ALIGN_PARAGRAPH.CENTER, size=12, bold=True, space_before=24)
    doc.add_page_break()

    # ---------------- DADES DEL TFM ----------------
    heading(doc, "Dades del Treball Final de Màster", 1)
    dades = [
        ("Autor", "Pau Planes Soria"),
        ("Títol", "Avaluació de l'evapotranspiració en vinya sota diferents tractaments de "
                  "poda mitjançant teledetecció d'alta precisió amb drons i mesures de radiació PAR"),
        ("Any d'elaboració", "2025-2026"),
        ("Tutor", "Dr. José Antonio Martínez Casasnovas (Departament de Medi Ambient i "
                  "Ciències del Sòl, Universitat de Lleida)"),
        ("Co-tutor", "Dr. Joaquim Bellvert (Programa d'Ús Eficient de l'Aigua, IRTA)"),
        ("Tipus de treball", "Treball Final de Màster (experimental)"),
        ("Titulació", "Màster en Enginyeria Agronòmica interuniversitari (MENAG-i)"),
        ("Centre", "Escola Tècnica Superior d'Enginyeria Agroalimentària i Forestal i de "
                   "Veterinària (ETSEAFiV), Universitat de Lleida"),
        ("Paraules clau", "vinya; evapotranspiració; teledetecció; dron; radiació PAR; "
                          "poda; estrès hídric; reg de precisió"),
        ("Palabras clave", "viña; evapotranspiración; teledetección; dron; radiación PAR; "
                           "poda; estrés hídrico; riego de precisión"),
        ("Key words", "vineyard; evapotranspiration; remote sensing; UAV; PAR radiation; "
                      "pruning; water stress; precision irrigation"),
    ]
    for k, v in dades:
        rich(doc, [(k + ": ", {'bold': True}), (v, {})], align=WD_ALIGN_PARAGRAPH.LEFT)
    doc.add_page_break()

    # ---------------- AGRAÏMENTS ----------------
    heading(doc, "Agraïments", 1)
    placeholder(doc, "[A PERSONALITZAR PER L'AUTOR] Espai per agrair als tutors "
                "(Dr. José A. Martínez Casasnovas i Dr. Joaquim Bellvert), a l'equip del "
                "Programa d'Ús Eficient de l'Aigua de l'IRTA, al personal de la finca "
                "experimental i a la família i amics. Vegeu l'estil dels TFM d'exemple "
                "(Alsina, 2020; Cuñé, 2019).")
    doc.add_page_break()

    # ---------------- RESUM (CAT) ----------------
    heading(doc, "Resum", 1)
    para(doc,
         "La gestió eficient de l'aigua de reg és un dels principals reptes de la "
         "viticultura mediterrània en un context d'escassetat hídrica creixent i de "
         "canvi climàtic. La caracterització precisa de l'evapotranspiració real (ETa) "
         "a escala de parcel·la és clau per ajustar el reg a la demanda real del cultiu. "
         "La poda és una de les pràctiques agronòmiques que més modifica l'arquitectura "
         "del dosser vegetal —el seu volum, l'índex d'àrea foliar (LAI) i la fracció de "
         "radiació fotosintèticament activa interceptada (fIPAR)— i, per tant, condiciona "
         "directament la transpiració i el balanç energètic de la vinya.")
    para(doc,
         "Aquest Treball Final de Màster, desenvolupat en col·laboració amb l'IRTA, té "
         "com a objectiu avaluar l'efecte de diferents tractaments de poda sobre "
         "l'evapotranspiració de la vinya combinant teledetecció d'alta precisió amb dron "
         "i mesures de radiació PAR en camp. En una parcel·la experimental amb un disseny "
         "en blocs a l'atzar i tres tractaments de poda, es realitzaran tres vols de dron "
         "al llarg del cicle vegetatiu amb càmeres tèrmica i multiespectral. A partir de "
         "les imatges es derivaran les variables biofísiques del dosser (LAI, fIPAR, "
         "alçada i volum) i les temperatures de dosser i de sòl, que alimentaran els "
         "models de balanç energètic de dues fonts (TSEB) i de Shuttleworth–Wallace per "
         "estimar l'ETa i l'ET potencial. L'índex d'estrès hídric del cultiu "
         "(CWSI = 1 − ETa/ETp) es relacionarà amb mesures de potencial hídric de tija "
         "(Ψstem), i les estimacions de fIPAR es validaran amb ceptòmetre i amb imatges "
         "hemisfèriques de baix cost.")
    placeholder(doc, "[A COMPLETAR AMB ELS RESULTATS REALS] Resum quantitatiu dels "
                "principals resultats (diferències entre tractaments en LAI, fIPAR, ETa i "
                "CWSI; bondat d'ajust de les validacions) i conclusió principal.")
    rich(doc, [("Paraules clau: ", {'bold': True}),
               ("vinya; evapotranspiració; teledetecció; dron; radiació PAR; poda; "
                "estrès hídric; reg de precisió.", {})], align=WD_ALIGN_PARAGRAPH.LEFT)

    # ---------------- RESUMEN (ES) ----------------
    heading(doc, "Resumen", 1)
    para(doc,
         "La gestión eficiente del agua de riego es uno de los principales retos de la "
         "viticultura mediterránea en un contexto de creciente escasez hídrica y de cambio "
         "climático. La caracterización precisa de la evapotranspiración real (ETa) a "
         "escala de parcela es clave para ajustar el riego a la demanda real del cultivo. "
         "La poda es una de las prácticas agronómicas que más modifica la arquitectura del "
         "dosel vegetal —su volumen, el índice de área foliar (LAI) y la fracción de "
         "radiación fotosintéticamente activa interceptada (fIPAR)— y, por tanto, "
         "condiciona directamente la transpiración y el balance energético de la viña.")
    para(doc,
         "Este Trabajo Final de Máster, desarrollado en colaboración con el IRTA, tiene "
         "como objetivo evaluar el efecto de diferentes tratamientos de poda sobre la "
         "evapotranspiración de la viña combinando teledetección de alta precisión con dron "
         "y medidas de radiación PAR en campo. En una parcela experimental con un diseño en "
         "bloques al azar y tres tratamientos de poda, se realizarán tres vuelos de dron a "
         "lo largo del ciclo vegetativo con cámaras térmica y multiespectral. A partir de "
         "las imágenes se derivarán las variables biofísicas del dosel (LAI, fIPAR, altura "
         "y volumen) y las temperaturas de dosel y de suelo, que alimentarán los modelos de "
         "balance energético de dos fuentes (TSEB) y de Shuttleworth–Wallace para estimar la "
         "ETa y la ET potencial. El índice de estrés hídrico del cultivo "
         "(CWSI = 1 − ETa/ETp) se relacionará con medidas de potencial hídrico de tallo "
         "(Ψstem), y las estimaciones de fIPAR se validarán con ceptómetro y con imágenes "
         "hemisféricas de bajo coste.")
    placeholder(doc, "[A COMPLETAR CON LOS RESULTADOS REALES] Resumen cuantitativo de los "
                "principales resultados y conclusión principal.")
    rich(doc, [("Palabras clave: ", {'bold': True}),
               ("viña; evapotranspiración; teledetección; dron; radiación PAR; poda; "
                "estrés hídrico; riego de precisión.", {})], align=WD_ALIGN_PARAGRAPH.LEFT)

    # ---------------- ABSTRACT (EN) ----------------
    heading(doc, "Abstract", 1)
    para(doc,
         "Efficient irrigation water management is one of the main challenges of "
         "Mediterranean viticulture under increasing water scarcity and climate change. "
         "An accurate characterisation of actual evapotranspiration (ETa) at the plot scale "
         "is key to matching irrigation to the real crop demand. Pruning is one of the "
         "agronomic practices that most strongly modifies canopy architecture —its volume, "
         "leaf area index (LAI) and fraction of intercepted photosynthetically active "
         "radiation (fIPAR)— and therefore directly conditions transpiration and the energy "
         "balance of the vineyard.")
    para(doc,
         "This Master's Thesis, carried out in collaboration with IRTA, aims to assess the "
         "effect of different pruning treatments on vineyard evapotranspiration by combining "
         "high-resolution UAV remote sensing and field PAR measurements. In an experimental "
         "plot with a randomised block design and three pruning treatments, three UAV "
         "flights will be performed over the growing season with thermal and multispectral "
         "cameras. Canopy biophysical variables (LAI, fIPAR, height and volume) and canopy "
         "and soil temperatures will be derived from the imagery and used to drive the "
         "two-source energy balance (TSEB) and Shuttleworth–Wallace models to estimate ETa "
         "and potential ET. The crop water stress index (CWSI = 1 − ETa/ETp) will be related "
         "to stem water potential (Ψstem) measurements, and fIPAR estimates will be "
         "validated against a ceptometer and low-cost hemispherical imaging.")
    placeholder(doc, "[TO BE COMPLETED WITH ACTUAL RESULTS] Quantitative summary of the main "
                "findings and main conclusion.")
    rich(doc, [("Key words: ", {'bold': True}),
               ("vineyard; evapotranspiration; remote sensing; UAV; PAR radiation; pruning; "
                "water stress; precision irrigation.", {})], align=WD_ALIGN_PARAGRAPH.LEFT)
    doc.add_page_break()

    # ---------------- ÍNDEX ----------------
    heading(doc, "Índex", 1)
    add_toc(doc)
    doc.add_page_break()

    # ---------------- LLISTA D'ABREVIATURES ----------------
    heading(doc, "Llista d'abreviatures i símbols", 1)
    abrev = [
        ("CWSI", "Índex d'estrès hídric del cultiu (Crop Water Stress Index)"),
        ("DSM / DTM", "Model digital de superfície / Model digital del terreny"),
        ("ET0", "Evapotranspiració de referència"),
        ("ETa", "Evapotranspiració real (actual)"),
        ("ETc", "Evapotranspiració del cultiu"),
        ("ETp", "Evapotranspiració potencial"),
        ("fIPAR", "Fracció de radiació fotosintèticament activa interceptada"),
        ("LAI", "Índex d'àrea foliar (Leaf Area Index)"),
        ("NDVI", "Índex de vegetació de diferència normalitzada"),
        ("PAR", "Radiació fotosintèticament activa (Photosynthetically Active Radiation)"),
        ("S–W", "Model de Shuttleworth–Wallace"),
        ("Tc / Ts", "Temperatura de dosser / Temperatura de sòl"),
        ("TSEB", "Balanç energètic de dues fonts (Two-Source Energy Balance)"),
        ("UAV", "Vehicle aeri no tripulat (dron)"),
        ("VPD", "Dèficit de pressió de vapor"),
        ("XAC", "Xarxa Agrometeorològica de Catalunya"),
        ("Ψstem", "Potencial hídric de tija (stem water potential)"),
    ]
    for k, v in abrev:
        rich(doc, [(k + " — ", {'bold': True}), (v, {})], align=WD_ALIGN_PARAGRAPH.LEFT)
    doc.add_page_break()

    # =====================================================================
    # 1. INTRODUCCIÓ
    # =====================================================================
    heading(doc, "1. Introducció", 1)

    heading(doc, "1.1. Context i justificació", 2)
    para(doc,
         "L'aigua és el principal factor limitant de la producció agrària a la conca "
         "mediterrània. Els escenaris de canvi climàtic projecten per al sud d'Europa un "
         "augment de la temperatura, una major freqüència d'episodis de sequera i una "
         "reducció de la disponibilitat d'aigua per al reg, fet que situa la millora de "
         "l'eficiència en l'ús de l'aigua al centre de l'agenda agronòmica. En aquest "
         "context, la viticultura mediterrània —un sector de gran rellevància econòmica i "
         "paisatgística a Catalunya— ha d'evolucionar cap a estratègies de reg de precisió "
         "que ajustin els aports hídrics a la demanda real de cada parcel·la.")
    para(doc,
         "L'evapotranspiració (ET) és la variable que integra la pèrdua d'aigua del sistema "
         "sòl–planta–atmosfera i, per tant, quantifica el consum hídric del cultiu. La seva "
         "estimació precisa a escala de parcel·la és la base de qualsevol programació de reg "
         "eficient (Allen et al., 1998). Tradicionalment, l'ET del cultiu s'ha estimat "
         "multiplicant l'evapotranspiració de referència per un coeficient de cultiu "
         "tabulat; tanmateix, aquest enfocament no captura la variabilitat espacial ni "
         "l'efecte de pràctiques agronòmiques concretes sobre el dosser vegetal. La "
         "teledetecció basada en models de balanç energètic ha permès superar aquesta "
         "limitació estimant l'ET real (ETa) píxel a píxel a partir de la temperatura de la "
         "superfície (Norman et al., 1995; Bellvert et al., 2021).")
    para(doc,
         "La poda és una de les pràctiques de maneig que més condiciona l'arquitectura del "
         "dosser de la vinya. En determinar el nombre de gemmes i de pàmpols, la poda regula "
         "el desenvolupament de l'àrea foliar, la mida i el volum de la capçada i, per tant, "
         "la quantitat de radiació interceptada i la superfície transpirant (Choné et al., "
         "2001). Diferents intensitats de poda donen lloc a dossers amb diferent capacitat "
         "d'interceptar radiació i d'evapotranspirar, però la quantificació d'aquest efecte "
         "amb mètodes operacionals i d'alta resolució encara és escassa en vinya.")
    para(doc,
         "Estudis recents han demostrat la viabilitat d'estimar l'ETa en cultius llenyosos "
         "combinant imatges tèrmiques i multiespectrals d'alta resolució obtingudes amb dron "
         "amb el model de balanç energètic de dues fonts (TSEB), que separa la transpiració "
         "del dosser de l'evaporació del sòl (Bellvert et al., 2021). Paral·lelament, s'han "
         "proposat metodologies de baix cost basades en imatges hemisfèriques per mesurar la "
         "fracció de radiació fotosintèticament activa interceptada (fIPAR) al llarg del dia, "
         "com a alternativa pràctica al ceptòmetre tradicional (Belaid et al., 2025). Aquest "
         "treball, desenvolupat en col·laboració amb l'IRTA, integra ambdós enfocaments per "
         "estudiar com els tractaments de poda modifiquen l'ET de la vinya.")

    heading(doc, "1.2. Estructura del treball", 2)
    para(doc,
         "El treball s'organitza en vuit capítols. Després d'aquesta introducció, el capítol "
         "2 desenvolupa el marc teòric, que revisa els fonaments de l'evapotranspiració, els "
         "models de balanç energètic, les variables biofísiques del dosser, la teledetecció "
         "amb dron, els mètodes de mesura de la radiació PAR, l'efecte de la poda i els "
         "indicadors d'estrès hídric. El capítol 3 presenta els objectius i les hipòtesis. "
         "El capítol 4 detalla els materials i mètodes (àrea d'estudi, disseny experimental, "
         "vols de dron, processament d'imatges, mesures de camp, models d'ET i anàlisi "
         "estadística). Els capítols 5 i 6 recullen, respectivament, els resultats i la seva "
         "discussió. El capítol 7 sintetitza les conclusions i, finalment, el capítol 8 "
         "llista les referències bibliogràfiques.")

    # =====================================================================
    # 2. MARC TEÒRIC
    # =====================================================================
    heading(doc, "2. Marc teòric", 1)

    heading(doc, "2.1. La viticultura mediterrània i el repte de l'escassetat hídrica", 2)
    para(doc,
         "La vinya (Vitis vinifera L.) és un cultiu adaptat al clima mediterrani, "
         "caracteritzat per estius càlids i secs que coincideixen amb el període de màxima "
         "demanda evaporativa. Tradicionalment conreada en secà, en les darreres dècades "
         "s'ha estès el reg de suport (reg deficitari) per estabilitzar la producció i la "
         "qualitat del raïm. Tanmateix, la creixent competència per l'aigua i la reducció "
         "dels recursos disponibles obliguen a optimitzar cada metre cúbic aplicat. El reg "
         "deficitari controlat, que aplica menys aigua que l'ET màxima del cultiu en moments "
         "fenològics determinats, s'ha consolidat com a estratègia per millorar l'eficiència "
         "de l'ús de l'aigua i la qualitat del most, però requereix conèixer amb precisió "
         "tant la demanda hídrica com l'estat hídric de la planta (Choné et al., 2001; "
         "Bellvert et al., 2014).")

    heading(doc, "2.2. L'evapotranspiració: ET0, ETc, ETa i ETp i el balanç d'energia", 2)
    para(doc,
         "L'evapotranspiració engloba dos processos simultanis: l'evaporació de l'aigua del "
         "sòl i la transpiració de les plantes. La metodologia FAO-56 (Allen et al., 1998) "
         "estableix el marc de referència estàndard, definint l'evapotranspiració de "
         "referència (ET0) com la d'una superfície hipotètica de gespa ben regada, calculada "
         "amb l'equació de Penman–Monteith a partir de dades meteorològiques. "
         "L'evapotranspiració del cultiu en condicions estàndard (ETc) s'obté multiplicant "
         "ET0 per un coeficient de cultiu (Kc), mentre que l'evapotranspiració real (ETa) "
         "incorpora les limitacions reals (estrès hídric, coberta parcial del sòl) i sol ser "
         "inferior a l'ETc. L'evapotranspiració potencial (ETp), en aquest treball, es pren "
         "com l'ET que tindria el cultiu sense restriccions d'aigua, de manera que la relació "
         "ETa/ETp informa del grau d'estrès.")
    para(doc,
         "Físicament, l'ET és un terme del balanç d'energia de la superfície: la radiació "
         "neta (Rn) es reparteix entre el flux de calor al sòl (G), el flux de calor sensible "
         "(H) i el flux de calor latent (LE), aquest últim equivalent a l'energia consumida "
         "en evapotranspirar (Rn = G + H + LE). Els models de teledetecció basats en el "
         "balanç energètic estimen LE com a residu, després de calcular Rn, G i H a partir de "
         "la temperatura de la superfície i de dades meteorològiques (Norman et al., 1995).")

    heading(doc, "2.3. Models de balanç energètic per teledetecció: TSEB i Shuttleworth–Wallace", 2)
    para(doc,
         "El model de balanç energètic de dues fonts (Two-Source Energy Balance, TSEB), "
         "formulat per Norman et al. (1995) i desenvolupat posteriorment per Kustas i Norman, "
         "tracta el sòl i la vegetació com dues fonts diferenciades d'energia i de "
         "temperatura. A partir de la temperatura radiomètrica de la superfície, l'índex "
         "d'àrea foliar i les dades meteorològiques, el TSEB separa la transpiració del "
         "dosser de l'evaporació del sòl, fet especialment rellevant en cultius llenyosos en "
         "files com la vinya, on el sòl descobert ocupa una fracció important de la "
         "superfície. Aquesta capacitat de partició dels fluxos fa del TSEB l'eina de "
         "referència per estimar l'ETa en fruiters i vinya mitjançant imatges d'alta "
         "resolució de dron (Bellvert et al., 2021). La seva implementació es troba "
         "disponible de forma oberta en el paquet pyTSEB (Nieto i Kustas).")
    para(doc,
         "El model de Shuttleworth–Wallace (S–W) (Shuttleworth i Wallace, 1985) estén "
         "l'equació combinada de Penman–Monteith a cobertes vegetals esparses mitjançant una "
         "xarxa de resistències que acobla els fluxos del sòl i del dosser. En aquest treball "
         "s'utilitza per estimar l'ET potencial (ETp) de cada tractament, és a dir, l'ET que "
         "tindria el cultiu en absència de restriccions hídriques, prenent com a referència "
         "l'estructura real del dosser. La comparació entre l'ETa derivada del TSEB i l'ETp "
         "del S–W permet quantificar el grau d'estrès hídric (vegeu §2.8).")

    heading(doc, "2.4. Variables biofísiques del dosser: LAI, fIPAR i fracció de coberta", 2)
    para(doc,
         "El funcionament del dosser i el seu consum d'aigua depenen de variables biofísiques "
         "clau. L'índex d'àrea foliar (LAI) és la superfície foliar per unitat de superfície "
         "de sòl i determina la capacitat fotosintètica i transpirant del cultiu. La fracció "
         "de radiació fotosintèticament activa interceptada (fIPAR) quantifica la proporció "
         "de radiació PAR que el dosser intercepta, i està directament relacionada amb el LAI "
         "i amb l'estructura de la coberta. Aquestes variables són paràmetres d'entrada "
         "essencials dels models de balanç energètic i, alhora, descriptors directes de "
         "l'efecte de la poda sobre el dosser (Bellvert et al., 2021; Belaid et al., 2025).")
    para(doc,
         "En cultius en files com la vinya, la fracció de coberta vegetal i el volum de la "
         "capçada condicionen tant la radiació interceptada com la partició entre "
         "transpiració i evaporació del sòl. Per això, la seva caracterització precisa —ja "
         "sigui amb mètodes òptics de camp o amb teledetecció— és un pas previ indispensable "
         "per a una estimació fiable de l'ET.")

    heading(doc, "2.5. Teledetecció amb drons: sensors tèrmics i multiespectrals", 2)
    para(doc,
         "Els vehicles aeris no tripulats (UAV o drons) han revolucionat el monitoratge dels "
         "cultius en permetre adquirir imatges amb resolucions espacials centimètriques i amb "
         "una gran flexibilitat temporal. Equipats amb càmeres multiespectrals (que "
         "registren bandes del visible, el red-edge i l'infraroig proper) permeten calcular "
         "índexs de vegetació com el NDVI i estimar variables com el LAI i la fIPAR. Les "
         "càmeres tèrmiques, per la seva banda, mesuren la temperatura de la superfície, "
         "necessària per als models de balanç energètic i per als indicadors d'estrès hídric "
         "(Bellvert et al., 2014). El processament fotogramètric de les imatges "
         "(structure-from-motion) genera, a més, models digitals de superfície i del terreny "
         "a partir dels quals es deriven l'alçada i el volum de la capçada de cada cep.")
    para(doc,
         "L'alta resolució de les imatges de dron és especialment valuosa en vinya, on "
         "permet discriminar els píxels purs de vegetació dels de sòl, i així obtenir "
         "temperatures de dosser (Tc) i de sòl (Ts) diferenciades que milloren la partició de "
         "fluxos del model TSEB respecte de sensors satel·litaris de menor resolució "
         "(Bellvert et al., 2021).")

    heading(doc, "2.6. Mesura de la radiació PAR i la fIPAR: ceptòmetre vs. imatge hemisfèrica", 2)
    para(doc,
         "La mesura de referència de la fIPAR en camp es realitza amb ceptòmetres lineals "
         "(p. ex., AccuPAR LP-80), que comparen la radiació PAR per sobre i per sota del "
         "dosser. No obstant això, la mesura puntual al migdia solar no captura la variació "
         "diürna de la interceptació de llum, que en cultius en files depèn fortament de "
         "l'orientació de les fileres, de la geometria del dosser i de l'angle solar. Per "
         "superar aquesta limitació, Belaid et al. (2025) han proposat i validat un mètode "
         "basat en imatges hemisfèriques obtingudes amb una càmera d'acció de baix cost "
         "(GoPro), que reconstrueix la corba diürna de fIPAR i el seu valor integrat diari "
         "amb un elevat grau d'acord respecte del ceptòmetre. Aquest mètode operacional i "
         "econòmic és especialment atractiu per a la seva aplicació en vinya i fruiters.")

    heading(doc, "2.7. La poda en vinya: efectes sobre el dosser i el consum d'aigua", 2)
    para(doc,
         "La poda regula el balanç entre vigor vegetatiu i càrrega productiva del cep. En "
         "determinar el nombre de gemmes que brotaran, condiciona el desenvolupament de "
         "l'àrea foliar, la mida de la capçada i la quantitat de radiació interceptada. "
         "Diferents sistemes i intensitats de poda (per exemple, poda llarga tipus Guyot, "
         "poda curta en cordó Royat o poda mínima) generen dossers amb estructures, volums i "
         "superfícies transpirants diferents. Com que la transpiració és proporcional a la "
         "superfície foliar activa i a la radiació interceptada, és previsible que els "
         "tractaments amb major vigor i volum de dosser presentin valors més elevats d'ETa, "
         "mentre que les podes més restrictives redueixin el consum d'aigua però puguin "
         "augmentar l'estrès relatiu del cep (Choné et al., 2001). Quantificar aquesta "
         "relació de manera precisa és l'objecte central d'aquest treball.")

    heading(doc, "2.8. Estrès hídric i indicadors: CWSI i potencial hídric de tija", 2)
    para(doc,
         "L'índex d'estrès hídric del cultiu (Crop Water Stress Index, CWSI) va ser proposat "
         "per Idso et al. (1981) i Jackson et al. (1981) a partir de la diferència entre la "
         "temperatura del dosser i la de l'aire, normalitzada pel dèficit de pressió de "
         "vapor. Conceptualment, el CWSI es pot expressar com 1 − ETa/ETp, de manera que "
         "pren valors propers a 0 quan el cultiu transpira sense restriccions i propers a 1 "
         "en condicions d'estrès màxim. Quan s'estima a partir d'imatges tèrmiques d'alta "
         "resolució, el CWSI s'ha correlacionat estretament amb l'estat hídric de la planta "
         "en vinya (Bellvert et al., 2014).")
    para(doc,
         "La mesura de referència de l'estat hídric de la planta és el potencial hídric de "
         "tija al migdia (Ψstem), determinat amb cambra de pressió en fulles prèviament "
         "embolcallades. Choné et al. (2001) van demostrar que el Ψstem és l'indicador més "
         "sensible i discriminant de l'estat hídric de la vinya, motiu pel qual s'utilitza en "
         "aquest treball com a referència per validar el CWSI derivat de la teledetecció.")

    # =====================================================================
    # 8. REFERÈNCIES (sembra; s'ampliarà)
    # =====================================================================
    heading(doc, "8. Referències bibliogràfiques", 1)
    refs = [
        "Allen, R. G., Pereira, L. S., Raes, D., & Smith, M. (1998). Crop "
        "evapotranspiration: Guidelines for computing crop water requirements (FAO "
        "Irrigation and Drainage Paper No. 56). Food and Agriculture Organization of the "
        "United Nations.",
        "Belaid, M. I., Escolà, A., & Casadesús, J. (2025). Hemispherical imaging of "
        "canopy light interception: A ceptometer alternative for precision irrigation in "
        "orchards and vineyards. Agricultural and Forest Meteorology, 377, 110958. "
        "https://doi.org/10.1016/j.agrformet.2025.110958",
        "Bellvert, J., Nieto, H., Pelechá, A., Jofre-Čekalović, C., Zazurca, L., & "
        "Miarnau, X. (2021). Remote sensing energy balance model for the assessment of crop "
        "evapotranspiration and water status in an almond rootstock collection. Frontiers in "
        "Plant Science, 12, 608967. https://doi.org/10.3389/fpls.2021.608967",
        "Bellvert, J., Zarco-Tejada, P. J., Girona, J., & Fereres, E. (2014). Mapping crop "
        "water stress index in a 'Pinot-noir' vineyard: Comparing ground measurements with "
        "thermal remote sensing imagery from an unmanned aerial vehicle. Precision "
        "Agriculture, 15(4), 361–376. https://doi.org/10.1007/s11119-013-9334-5",
        "Choné, X., Van Leeuwen, C., Dubourdieu, D., & Gaudillère, J. P. (2001). Stem water "
        "potential is a sensitive indicator of grapevine water status. Annals of Botany, "
        "87(4), 477–483. https://doi.org/10.1006/anbo.2000.1361",
        "Idso, S. B., Jackson, R. D., Pinter, P. J., Reginato, R. J., & Hatfield, J. L. "
        "(1981). Normalizing the stress-degree-day parameter for environmental variability. "
        "Agricultural Meteorology, 24, 45–55. https://doi.org/10.1016/0002-1571(81)90032-7",
        "Jackson, R. D., Idso, S. B., Reginato, R. J., & Pinter, P. J. (1981). Canopy "
        "temperature as a crop water stress indicator. Water Resources Research, 17(4), "
        "1133–1138. https://doi.org/10.1029/WR017i004p01133",
        "Norman, J. M., Kustas, W. P., & Humes, K. S. (1995). Source approach for "
        "estimating soil and vegetation energy fluxes in observations of directional "
        "radiometric surface temperature. Agricultural and Forest Meteorology, 77(3–4), "
        "263–293. https://doi.org/10.1016/0168-1923(95)02265-Y",
        "Shuttleworth, W. J., & Wallace, J. S. (1985). Evaporation from sparse crops: An "
        "energy combination theory. Quarterly Journal of the Royal Meteorological Society, "
        "111(469), 839–855. https://doi.org/10.1002/qj.49711146910",
    ]
    for r in sorted(refs):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.left_indent = Cm(1.25)
        p.paragraph_format.first_line_indent = Cm(-1.25)  # sagnat francès APA
        run = p.add_run(r)
        _set_run_font(run, size=12)

    add_page_number_footer(doc)
    doc.save("/home/user/TFM_PAU/TFM_Pau_Planes.docx")
    print("Document desat: TFM_Pau_Planes.docx")


if __name__ == "__main__":
    build()
