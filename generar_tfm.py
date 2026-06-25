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
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FIGDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "figures")
CENTER = WD_ALIGN_PARAGRAPH.CENTER

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

    _ensure_caption_style(doc)


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


def _ensure_caption_style(doc):
    try:
        st = doc.styles['Caption']
    except KeyError:
        st = doc.styles.add_style('Caption', WD_STYLE_TYPE.PARAGRAPH)
    st.font.name = FONT
    st.font.size = Pt(10)
    st.font.italic = False
    st.font.bold = False
    st.font.color.rgb = RGBColor(0, 0, 0)
    st.paragraph_format.space_before = Pt(2)
    st.paragraph_format.space_after = Pt(10)
    st.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE


def _seq_field(paragraph, seq_name):
    run = paragraph.add_run()
    f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve')
    instr.text = f' SEQ {seq_name} \\* ARABIC '
    f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'end')
    run._element.append(f1); run._element.append(instr); run._element.append(f2)
    _set_run_font(run, size=10, bold=True)


def add_figure(doc, filename, caption, source, width_cm=14.0):
    p = doc.add_paragraph(); p.alignment = CENTER
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(2)
    run = p.add_run()
    run.add_picture(os.path.join(FIGDIR, filename), width=Cm(width_cm))
    cap = doc.add_paragraph(style='Caption'); cap.alignment = CENTER
    r1 = cap.add_run('Figura '); _set_run_font(r1, size=10, bold=True)
    _seq_field(cap, 'Figura')
    r2 = cap.add_run('. ' + caption + ' '); _set_run_font(r2, size=10)
    r3 = cap.add_run('Font: ' + source); _set_run_font(r3, size=10, italic=True)


def add_table_caption(doc, caption, source):
    cap = doc.add_paragraph(style='Caption'); cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r1 = cap.add_run('Taula '); _set_run_font(r1, size=10, bold=True)
    _seq_field(cap, 'Taula')
    r2 = cap.add_run('. ' + caption + ' '); _set_run_font(r2, size=10)
    if source:
        r3 = cap.add_run('Font: ' + source); _set_run_font(r3, size=10, italic=True)


def add_table(doc, headers, rows, widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].paragraphs[0].alignment = CENTER
        run = hdr[i].paragraphs[0].add_run(h)
        _set_run_font(run, size=10.5, bold=True)
    for row in rows:
        cells = table.add_row().cells
        for i, val in enumerate(row):
            para_c = cells[i].paragraphs[0]
            para_c.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = para_c.add_run(val)
            _set_run_font(run, size=10.5)
    if widths:
        for i, w in enumerate(widths):
            for row in table.rows:
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return table


def add_list_of(doc, seq_name):
    p = doc.add_paragraph()
    run = p.add_run()
    f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve')
    instr.text = f' TOC \\h \\z \\c "{seq_name}" '
    sep = OxmlElement('w:fldChar'); sep.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t')
    t.text = "Actualitzeu els camps (clic dret → «Actualitza els camps») per generar la llista."
    end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
    for el in (f1, instr, sep, t, end):
        run._element.append(el)


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
        ("Paraules clau", "vinya; Tempranillo; evapotranspiració; teledetecció; dron; radiació PAR; "
                          "maneig de capçada; estrès hídric; reg de precisió"),
        ("Palabras clave", "viña; Tempranillo; evapotranspiración; teledetección; dron; radiación PAR; "
                           "manejo de copa; estrés hídrico; riego de precisión"),
        ("Key words", "vineyard; Tempranillo; evapotranspiration; remote sensing; UAV; PAR radiation; "
                      "canopy management; water stress; precision irrigation"),
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
         "i mesures de radiació PAR en camp. En una parcel·la experimental de la varietat "
         "'Ull de Llebre' (Tempranillo) a Raimat, amb un disseny en blocs de cinc "
         "tractaments de maneig de capçada i quatre repeticions, es realitzaran tres vols de "
         "dron al llarg del cicle vegetatiu amb càmeres tèrmica i multiespectral. A partir de "
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
         "y medidas de radiación PAR en campo. En una parcela experimental de la variedad "
         "'Tempranillo' en Raimat, con un diseño en bloques de cinco tratamientos de manejo "
         "de copa y cuatro repeticiones, se realizarán tres vuelos de dron a lo largo del "
         "ciclo vegetativo con cámaras térmica y multiespectral. A partir de "
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
         "'Tempranillo' plot in Raimat, with a block design of five canopy-management "
         "treatments and four replicates, three UAV "
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

    # ---------------- ÍNDEX DE FIGURES I TAULES ----------------
    heading(doc, "Índex de figures", 1)
    add_list_of(doc, "Figura")
    para(doc, "", space_after=12)
    heading(doc, "Índex de taules", 1)
    add_list_of(doc, "Taula")
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
    para(doc,
         "Concretament, el treball s'emmarca en l'assaig de maneig de capçada que el Programa "
         "d'Ús Eficient de l'Aigua de l'IRTA desenvolupa en una parcel·la de la varietat "
         "'Ull de Llebre' (Tempranillo) a Raimat, dins dels projectes VITIMPACT i ADAPTEX, "
         "amb l'objectiu d'augmentar la resiliència de la vinya davant del canvi climàtic. En "
         "aquest assaig, cinc nivells d'esporga generen capçades de mides contrastades i, en "
         "campanyes anteriors, s'han observat diferències notables en el consum d'aigua de "
         "reg entre tractaments —les capçades més grans han arribat a consumir fins a un 14 % "
         "més d'aigua que el maneig comercial, mentre que l'esporga severa l'ha reduït prop "
         "d'un 20 %. El present TFM aborda precisament una de les línies de treball futures "
         "previstes en aquest assaig: la quantificació de la transpiració i de "
         "l'evapotranspiració mitjançant teledetecció, com a complement a les mesures "
         "directes de flux de saba i de potencial hídric.")

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
    para(doc,
         "El reg deficitari es defineix com l'aplicació d'aigua per sota dels requeriments "
         "complets d'evapotranspiració del cultiu, amb l'objectiu de reduir el consum "
         "d'aigua amb una penalització mínima —o fins i tot un benefici— sobre la qualitat "
         "de la collita (Fereres i Soriano, 2007). En vinya, aquesta estratègia és "
         "especialment efectiva perquè un cert nivell de dèficit hídric controlat afavoreix "
         "l'acumulació de compostos fenòlics i la concentració del most. Tanmateix, "
         "l'aplicació de reg deficitari de manera precisa exigeix conèixer, parcel·la a "
         "parcel·la, tant l'evapotranspiració real del cultiu com el grau d'estrès al qual "
         "està sotmesa la planta; és precisament aquesta necessitat la que justifica "
         "l'aproximació metodològica d'aquest treball.")

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
         "la temperatura de la superfície i de dades meteorològiques (Norman et al., 1995). "
         "La Figura 1 resumeix esquemàticament aquest balanç i la partició dels fluxos entre "
         "el sòl i la vegetació en què es basa el model emprat en aquest treball.")
    add_figure(doc, "fig1_tseb.png",
               "Esquema del balanç d'energia de la superfície i de la partició de fluxos "
               "entre el sòl i el dosser en què es fonamenta el model de balanç energètic de "
               "dues fonts (TSEB). Rn: radiació neta; G: flux de calor al sòl; H: calor "
               "sensible; LE: calor latent (LEc del dosser i LEs del sòl); Tc i Ts: "
               "temperatures de dosser i de sòl.",
               "elaboració pròpia a partir de Norman et al. (1995) i Kustas i Norman (1999).",
               width_cm=13.5)

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
         "disponible de forma oberta en el paquet pyTSEB (Nieto i Kustas). Diverses "
         "avaluacions del model han mostrat que les estimacions combinades dels fluxos de "
         "calor del sòl i de la vegetació s'ajusten a les observacions amb errors de l'ordre "
         "del 20 %, fet que valida la seva aplicació en cobertes amb coberta parcial com els "
         "cultius en files (Kustas i Norman, 1999).")
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
         "econòmic és especialment atractiu per a la seva aplicació en vinya i fruiters. La "
         "Figura 2 il·lustra per què una única mesura al migdia pot subestimar la "
         "interceptació diària de llum en un cultiu en files.")
    add_figure(doc, "fig2_fipar_diurnal.png",
               "Patró diürn de la fracció de radiació interceptada (fIPAR) en un cultiu en "
               "files. La mesura puntual del ceptòmetre al migdia solar pot diferir del valor "
               "diari integrat que recupera la imatge hemisfèrica al llarg del dia.",
               "elaboració pròpia a partir de Belaid et al. (2025).",
               width_cm=12.5)

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
    para(doc,
         "A més de l'efecte directe sobre la mida del dosser, el sistema de poda determina "
         "la distribució espacial de la fusta i del fullatge i, per tant, la geometria de la "
         "intercepció de la llum al llarg del dia. Aquesta dimensió temporal —difícil de "
         "capturar amb mesures puntuals— reforça l'interès de combinar la teledetecció "
         "instantània amb dron amb la caracterització de la corba diürna de fIPAR descrita a "
         "l'apartat anterior, de manera que es pugui relacionar de forma robusta "
         "l'arquitectura del dosser de cada tractament amb el seu consum real d'aigua.")

    heading(doc, "2.8. Estrès hídric i indicadors: CWSI i potencial hídric de tija", 2)
    para(doc,
         "L'índex d'estrès hídric del cultiu (Crop Water Stress Index, CWSI) va ser proposat "
         "per Idso et al. (1981) i Jackson et al. (1981) a partir de la diferència entre la "
         "temperatura del dosser i la de l'aire, normalitzada pel dèficit de pressió de "
         "vapor. Conceptualment, el CWSI es pot expressar com 1 − ETa/ETp, de manera que "
         "pren valors propers a 0 quan el cultiu transpira sense restriccions i propers a 1 "
         "en condicions d'estrès màxim. Quan s'estima a partir d'imatges tèrmiques d'alta "
         "resolució, el CWSI s'ha correlacionat estretament amb l'estat hídric de la planta "
         "en vinya (Bellvert et al., 2014). El càlcul del CWSI es recolza en dos límits de "
         "referència de la diferència entre la temperatura del dosser i la de l'aire en "
         "funció del dèficit de pressió de vapor, tal com es mostra a la Figura 3.")
    add_figure(doc, "fig3_cwsi.png",
               "Fonament del càlcul de l'índex d'estrès hídric del cultiu (CWSI) a partir de "
               "la diferència entre la temperatura del dosser i de l'aire (Tc − Ta) i el "
               "dèficit de pressió de vapor (VPD): límit inferior (cultiu ben regat, sense "
               "estrès) i límit superior (dosser que no transpira).",
               "elaboració pròpia a partir d'Idso et al. (1981) i Jackson et al. (1981).",
               width_cm=12.5)
    para(doc,
         "La mesura de referència de l'estat hídric de la planta és el potencial hídric de "
         "tija al migdia (Ψstem), determinat amb cambra de pressió en fulles prèviament "
         "embolcallades. Choné et al. (2001) van demostrar que el Ψstem és l'indicador més "
         "sensible i discriminant de l'estat hídric de la vinya, motiu pel qual s'utilitza en "
         "aquest treball com a referència per validar el CWSI derivat de la teledetecció.")

    # =====================================================================
    # 3. OBJECTIUS I HIPÒTESIS
    # =====================================================================
    heading(doc, "3. Objectius i hipòtesis", 1)

    heading(doc, "3.1. Objectiu general", 2)
    para(doc,
         "L'objectiu general d'aquest treball és avaluar l'efecte de diferents tractaments "
         "de poda sobre l'evapotranspiració de la vinya mitjançant teledetecció d'alta "
         "precisió amb dron i mesures de radiació PAR en camp.")

    heading(doc, "3.2. Objectius específics", 2)
    objectius = [
        "Estimar les variables biofísiques del dosser (LAI, fIPAR diari, alçada i volum "
        "del dosser) per a cada tractament de poda a partir d'imatges multiespectrals i del "
        "núvol de punts fotogramètric obtinguts en tres vols de dron al llarg del cicle "
        "vegetatiu.",
        "Calcular l'ETa i l'ETp de cada tractament de poda aplicant els models TSEB i "
        "Shuttleworth–Wallace a partir de les temperatures de dosser (Tc) i de sòl (Ts) "
        "obtingudes per imatge tèrmica d'alta resolució.",
        "Validar les estimacions de fIPAR derivades de les imatges de dron amb mesures "
        "simultànies de ceptòmetre i d'imatges hemisfèriques, seguint el protocol de Belaid "
        "et al. (2025).",
        "Quantificar l'índex d'estrès hídric del cultiu (CWSI = 1 − ETa/ETp) per a cada "
        "tractament de poda i relacionar-lo amb mesures de potencial hídric de tija "
        "(Ψstem).",
    ]
    for i, o in enumerate(objectius, 1):
        p = doc.add_paragraph(style='List Number')
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(o)
        _set_run_font(run, size=12)

    heading(doc, "3.3. Hipòtesis", 2)
    para(doc,
         "A partir dels objectius anteriors i del marc teòric, es plantegen les hipòtesis "
         "de treball següents:")
    hipotesis = [
        "H1. Els cinc tractaments de maneig de capçada generen diferències significatives en "
        "les variables biofísiques del dosser (LAI, fIPAR, alçada i volum), detectables amb "
        "teledetecció d'alta resolució amb dron.",
        "H2. Els tractaments amb major mida de capçada (SE, sense esporga) presenten valors "
        "més elevats d'ETa, mentre que els d'esporga severa (ES i ES+DV) en presenten de més "
        "baixos, en consonància amb les diferències de consum d'aigua de reg observades en "
        "campanyes anteriors de l'assaig.",
        "H3. En mantenir-se un mateix estat hídric entre tractaments mitjançant reg "
        "individualitzat, les diferències d'ETa i de CWSI s'expliquen principalment per "
        "l'arquitectura del dosser i no per diferències de Ψstem imposades.",
        "H4. La fIPAR estimada amb imatges hemisfèriques de baix cost (GoPro) presenta un "
        "grau d'acord elevat amb la mesura de ceptòmetre i amb la derivada de les imatges de "
        "dron, fet que en confirma l'aplicabilitat operacional en vinya.",
    ]
    for h in hipotesis:
        rich(doc, [(h[:3], {'bold': True}), (h[3:], {})], align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # =====================================================================
    # 4. MATERIALS I MÈTODES
    # =====================================================================
    heading(doc, "4. Materials i mètodes", 1)
    para(doc,
         "La Figura 4 resumeix el flux de treball metodològic seguit en aquest treball, des "
         "de l'adquisició de dades amb dron i en camp fins a l'anàlisi estadística de "
         "l'efecte dels tractaments de poda, passant pel processament d'imatges, l'aplicació "
         "dels models d'evapotranspiració i la validació. A continuació es descriu cada "
         "etapa en detall.")
    add_figure(doc, "fig5_workflow.png",
               "Esquema general del flux de treball metodològic del TFM, des de l'adquisició "
               "de dades fins a l'anàlisi estadística.",
               "elaboració pròpia.", width_cm=10.5)

    heading(doc, "4.1. Àrea d'estudi i material vegetal", 2)
    para(doc,
         "L'estudi s'emmarca en l'assaig de maneig de capçada de la vinya que el Programa "
         "d'Ús Eficient de l'Aigua de l'IRTA desenvolupa en el marc dels projectes VITIMPACT "
         "i ADAPTEX, orientat a augmentar la resiliència del cultiu davant del canvi "
         "climàtic. La parcel·la experimental se situa a Raimat (comarca del Segrià, Lleida) "
         "i està plantada amb la varietat 'Ull de Llebre' (Tempranillo) empeltada sobre el "
         "portaempelt Richter 110 (R-110).")
    para(doc,
         "La vinya, plantada l'any 2013, es condueix en doble cordó amb un marc de plantació "
         "d'1,6 m entre ceps i 2,5 m entre fileres. El reg s'aplica per degoteig seguint una "
         "estratègia de reg deficitari controlat (RDC), amb imposició del dèficit en el "
         "període post-verolat fins a la verema. Un aspecte clau del disseny és que cada "
         "tractament de maneig de capçada es rega de manera individualitzada per mantenir "
         "tots els tractaments en un mateix estat hídric, independentment de la mida de la "
         "capçada; d'aquesta manera, les diferències observades es poden atribuir a "
         "l'arquitectura del dosser i no a diferències d'estat hídric imposades. El potencial "
         "hídric de tija al migdia s'utilitza com a indicador de referència per controlar "
         "aquest estat hídric (Choné et al., 2001).")
    placeholder(doc, "[VERIFICAR/COMPLETAR] Coordenades exactes de la parcel·la, tipus de sòl "
                "i dades climàtiques de la campanya (p. ex., precipitació acumulada; el 2024-25 "
                "va ser de ~330 mm), i dotacions de reg reals d'aquesta temporada.")

    heading(doc, "4.2. Disseny experimental", 2)
    para(doc,
         "L'assaig segueix un disseny en blocs amb cinc tractaments de maneig de capçada "
         "(esporga) i quatre repeticions (blocs) per tractament, fet que dona lloc a 20 "
         "parcel·les experimentals. Cada parcel·la elemental està formada per 24 ceps "
         "distribuïts en tres fileres de vuit ceps, dels quals es prenen com a ceps de "
         "mostreig els centrals per minimitzar els efectes de vora. Els cinc tractaments es "
         "descriuen a la Taula 1 i la distribució dels blocs a la parcel·la es representa a "
         "la Figura 5.")
    add_table_caption(doc, "Tractaments de maneig de capçada (esporga) de l'assaig ADAPTEX. "
                      "El tractament EM (esporga mitjana) correspon a la pràctica comercial "
                      "de referència.", "")
    add_table(doc,
              ["Codi", "Tractament", "Descripció"],
              [["SE", "Sense esporga",
                "No s'elimina vegetació; capçada de mida màxima"],
               ["EL", "Esporga lleugera",
                "Reducció lleugera de la capçada"],
               ["EM", "Esporga mitjana (comercial)",
                "Maneig comercial de referència; capçada intermèdia"],
               ["ES", "Esporga severa",
                "Reducció intensa de la capçada; mida mínima"],
               ["ES+DV", "Esporga severa + desfullat a verolat",
                "Esporga severa amb desfullat addicional en verolat"]],
              widths=[1.6, 4.6, 6.8])
    add_figure(doc, "fig4_disseny.png",
               "Croquis del disseny experimental de l'assaig ADAPTEX: cinc tractaments de "
               "maneig de capçada (SE, EL, EM, ES, ES+DV) × quatre blocs = 20 parcel·les "
               "experimentals de 24 ceps cadascuna.",
               "elaboració pròpia a partir del croquis de l'assaig ADAPTEX (IRTA).",
               width_cm=13.5)

    heading(doc, "4.3. Vols de dron i sensors", 2)
    para(doc,
         "Es realitzaran tres vols de dron al voltant del migdia solar en tres moments del "
         "cicle vegetatiu —brotació, tancament del raïm (proximitat a floració/quallat) i "
         "verol— per capturar la dinàmica estacional del dosser i de l'estat hídric. El dron "
         "anirà equipat amb dues càmeres: (i) una càmera multiespectral de sis bandes, "
         "incloent-hi la regió del red-edge, per al càlcul d'índexs de vegetació i "
         "l'estimació de LAI i fIPAR; i (ii) una càmera tèrmica per a l'obtenció de la "
         "temperatura de la superfície. Els vols es planificaran amb un solapament frontal i "
         "lateral elevat per garantir una bona reconstrucció fotogramètrica, i s'inclouran "
         "panells de calibratge radiomètric i punts de control terrestre georeferenciats. "
         "El calendari previst de vols i de mesures de camp associades es resumeix a la "
         "Taula 2.")
    add_table_caption(doc, "Calendari previst de vols de dron i mesures de camp simultànies "
                      "segons la fase fenològica de la vinya.", "")
    add_table(doc,
              ["Vol", "Fase fenològica", "Període aproximat", "Mesures de camp simultànies"],
              [["1", "Brotació", "[a concretar]",
                "fIPAR (ceptòmetre + GoPro), Ψstem, meteo"],
               ["2", "Tancament del raïm", "[a concretar]",
                "fIPAR (ceptòmetre + GoPro), Ψstem, meteo"],
               ["3", "Verol", "[a concretar]",
                "fIPAR (ceptòmetre + GoPro), Ψstem, meteo"]],
              widths=[1.3, 4.0, 4.0, 5.7])
    placeholder(doc, "[A CONCRETAR] Model de dron i de cada càmera (fabricant, bandes "
                "espectrals, resolució tèrmica, GSD a l'altura de vol); altura de vol i "
                "solapaments; dates exactes dels tres vols; nombre i distribució dels punts "
                "de control (GCP); programari de planificació de vol.")

    heading(doc, "4.4. Processament fotogramètric i d'imatges", 2)
    para(doc,
         "Les imatges adquirides es processaran fotogramètricament mitjançant tècniques de "
         "structure-from-motion per generar ortomosaics multiespectrals i tèrmics "
         "georeferenciats, així com el model digital de superfície (DSM) i el model digital "
         "del terreny (DTM). La diferència entre el DSM i el DTM permetrà derivar l'alçada "
         "del dosser i, per integració, el volum de capçada de cada cep. A partir dels "
         "ortomosaics multiespectrals es calcularan índexs de vegetació (p. ex., NDVI) i "
         "s'estimaran el LAI i la fIPAR mitjançant relacions empíriques calibrades amb les "
         "mesures de camp.")
    para(doc,
         "Per separar les fonts del balanç energètic, es realitzarà una segmentació "
         "supervisada de la coberta que distingirà els píxels purs de vegetació dels de sòl, "
         "obtenint així la temperatura de dosser (Tc) i la temperatura de sòl (Ts) "
         "necessàries per als models TSEB i Shuttleworth–Wallace. Tota la cadena de "
         "processament seguirà l'esquema descrit per Bellvert et al. (2021) per a cultius "
         "llenyosos amb imatges de dron.")
    placeholder(doc, "[A CONCRETAR] Programari fotogramètric utilitzat (p. ex., Agisoft "
                "Metashape o Pix4D); resolució dels ortomosaics; mètode i índexs concrets "
                "per estimar LAI i fIPAR i la seva calibració; criteri de segmentació "
                "vegetació/sòl.")

    heading(doc, "4.5. Mesures de camp i validació", 2)
    para(doc,
         "Simultàniament a cada vol es prendran mesures de camp per parametritzar i validar "
         "els models:")
    camp = [
        "fIPAR: mesura al migdia solar i corba diürna amb ceptòmetre lineal (AccuPAR LP-80) "
        "i, en paral·lel, amb imatges hemisfèriques obtingudes amb una càmera d'acció "
        "(GoPro), seguint el protocol de Belaid et al. (2025). Aquestes dades serviran per "
        "validar la fIPAR derivada del dron.",
        "Potencial hídric de tija (Ψstem): determinat al migdia amb cambra de pressió "
        "(Scholander) en fulles prèviament embolcallades amb bossa opaca i reflectora, com a "
        "mesura de referència de l'estat hídric (Choné et al., 2001).",
        "Variables meteorològiques: temperatura de l'aire, humitat relativa, dèficit de "
        "pressió de vapor (VPD), velocitat del vent i radiació solar, obtingudes de "
        "l'estació agrometeorològica més propera de la Xarxa Agrometeorològica de Catalunya "
        "(XAC), necessàries com a entrades dels models de balanç energètic.",
    ]
    for c in camp:
        p = doc.add_paragraph(style='List Bullet')
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(c)
        _set_run_font(run, size=12)
    placeholder(doc, "[A CONCRETAR] Nombre de ceps mostrejats per tractament i bloc en cada "
                "mesura; hores exactes de les corbes diürnes; model de cambra de pressió; "
                "identificador i distància de l'estació XAC utilitzada.")

    heading(doc, "4.6. Models d'evapotranspiració i càlcul del CWSI", 2)
    para(doc,
         "L'evapotranspiració real (ETa) de cada tractament es calcularà amb el model de "
         "balanç energètic de dues fonts (TSEB; Norman et al., 1995), que estima els fluxos "
         "de calor latent del dosser i del sòl per separat a partir de Tc, Ts, el LAI i les "
         "dades meteorològiques. La implementació es farà amb el paquet obert pyTSEB (Nieto "
         "i Kustas), tal com s'ha aplicat en fruiters amb imatges de dron (Bellvert et al., "
         "2021). L'evapotranspiració potencial (ETp) s'estimarà amb el model de "
         "Shuttleworth–Wallace (Shuttleworth i Wallace, 1985), que descriu l'ET d'una "
         "coberta esparsa en condicions sense restricció hídrica a partir de l'estructura "
         "real del dosser.")
    para(doc,
         "A partir d'ambdues estimacions, l'índex d'estrès hídric del cultiu es definirà com "
         "CWSI = 1 − ETa/ETp (Idso et al., 1981; Jackson et al., 1981), que pren valors "
         "propers a 0 en absència d'estrès i propers a 1 en estrès màxim. El CWSI obtingut "
         "per teledetecció es contrastarà amb les mesures de Ψstem per avaluar-ne la "
         "capacitat de diagnòstic de l'estat hídric (Bellvert et al., 2014). La Taula 3 "
         "resumeix les variables d'entrada i les sortides principals de cada model.")
    add_table_caption(doc, "Variables d'entrada i sortides principals dels models "
                      "d'evapotranspiració emprats.", "")
    add_table(doc,
              ["Model", "Entrades principals", "Sortida principal"],
              [["TSEB", "Tc, Ts, LAI, Rn, dades meteorològiques (Ta, VPD, vent, radiació)",
                "ETa (transpiració + evaporació)"],
               ["Shuttleworth–Wallace", "Estructura del dosser (LAI, alçada), resistències, "
                "dades meteorològiques", "ETp (sense restricció hídrica)"],
               ["CWSI", "ETa i ETp", "Índex d'estrès hídric = 1 − ETa/ETp"]],
              widths=[3.8, 7.7, 3.5])

    heading(doc, "4.7. Anàlisi estadística", 2)
    para(doc,
         "L'efecte dels tractaments de maneig de capçada sobre les variables biofísiques del "
         "dosser (LAI, fIPAR, alçada, volum), sobre l'ETa, l'ETp i el CWSI, i sobre el Ψstem "
         "s'analitzarà comparant els cinc tractaments en cada data de mesura. Quan es "
         "compleixin els supòsits de normalitat dels residus i d'homogeneïtat de variàncies, "
         "s'aplicarà una anàlisi de la variància (ANOVA) amb separació de mitjanes mitjançant "
         "el test HSD de Tukey (p < 0,05), tal com s'ha fet en estudis equivalents en "
         "fruiters (Bellvert et al., 2021). En cas que aquests supòsits no es compleixin, "
         "s'utilitzarà l'alternativa no paramètrica de Kruskal–Wallis, criteri adoptat de "
         "manera consistent amb l'anàlisi de l'assaig ADAPTEX a l'IRTA. La concordança entre "
         "els mètodes de mesura de la fIPAR (dron, ceptòmetre i imatge hemisfèrica) i entre "
         "el CWSI i el Ψstem s'avaluarà amb regressió lineal i estadístics de bondat d'ajust "
         "(coeficient de determinació R², error quadràtic mitjà RMSE i biaix).")
    placeholder(doc, "[A CONCRETAR] Programari estadístic utilitzat (p. ex., R o JMP) i "
                "nivell de significació adoptat si difereix de 0,05.")

    # =====================================================================
    # 5–7. RESULTATS, DISCUSSIÓ I CONCLUSIONS (esquelet per omplir)
    # =====================================================================
    heading(doc, "5. Resultats", 1)
    placeholder(doc, "[CAPÍTOL A OMPLIR AMB LES DADES EXPERIMENTALS REALS] Estructura "
                "proposada (afegiu taules i figures a cada apartat):")
    heading(doc, "5.1. Variables biofísiques del dosser per tractament de capçada", 2)
    placeholder(doc, "Taules/figures de LAI, fIPAR, alçada i volum del dosser per als cinc "
                "tractaments (SE, EL, EM, ES, ES+DV) i les tres dates; resultats del test "
                "(ANOVA + Tukey o Kruskal–Wallis) amb lletres de separació de mitjanes.")
    heading(doc, "5.2. Validació de la fIPAR (dron vs. ceptòmetre vs. imatge hemisfèrica)", 2)
    placeholder(doc, "Regressions i estadístics (R², RMSE, biaix) entre els tres mètodes; "
                "corbes diürnes de fIPAR per tractament.")
    heading(doc, "5.3. Evapotranspiració (ETa i ETp) per tractament de capçada", 2)
    placeholder(doc, "Valors d'ETa i ETp estimats amb TSEB i S–W per tractament i data; "
                "partició transpiració/evaporació; comparació amb l'aigua de reg aplicada i, "
                "si escau, amb les mesures de flux de saba (cabalímetres) de l'assaig.")
    heading(doc, "5.4. Índex d'estrès hídric (CWSI) i relació amb el Ψstem", 2)
    placeholder(doc, "CWSI per tractament; regressió CWSI–Ψstem; mapes d'estrès hídric. "
                "Atès que el reg manté un mateix estat hídric entre tractaments, valoreu fins "
                "a quin punt el CWSI reflecteix l'arquitectura del dosser més que no pas "
                "l'estrès imposat.")

    heading(doc, "6. Discussió", 1)
    placeholder(doc, "[CAPÍTOL A OMPLIR] Interpreteu els resultats a la llum de la literatura "
                "(Bellvert et al., 2014, 2021; Belaid et al., 2025; Choné et al., 2001) i de "
                "les dades de l'assaig ADAPTEX (Blanco et al., 2025). Discutiu: (i) com el "
                "maneig de capçada modifica el dosser i l'ET; (ii) la coherència de l'ETa "
                "estimada amb el consum d'aigua de reg observat i amb el flux de saba; "
                "(iii) la fiabilitat de la fIPAR per imatge hemisfèrica respecte del "
                "ceptòmetre; (iv) la relació CWSI–Ψstem; (v) limitacions i implicacions per "
                "al reg de precisió.")

    heading(doc, "7. Conclusions", 1)
    placeholder(doc, "[CAPÍTOL A OMPLIR] Conclusions numerades que responguin directament als "
                "objectius i contrastin les hipòtesis H1–H4, més una conclusió pràctica sobre "
                "quina estratègia de poda optimitza l'ús de l'aigua en vinya.")

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
        "Blanco, V., Gutiérrez, N., Mata, M., Paris, C., Biru, A., & Girona, J. (2025). "
        "Maneig de capçada com a estratègia per augmentar la resiliència de la vinya enfront "
        "del canvi climàtic: Resultats temporada 2025 [Comunicació interna]. IRTA – Programa "
        "d'Ús Eficient de l'Aigua en Agricultura.",
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
        "Fereres, E., & Soriano, M. A. (2007). Deficit irrigation for reducing agricultural "
        "water use. Journal of Experimental Botany, 58(2), 147–159. "
        "https://doi.org/10.1093/jxb/erl165",
        "Kustas, W. P., & Norman, J. M. (1999). Evaluation of soil and vegetation heat flux "
        "predictions using a simple two-source model with radiometric temperatures for "
        "partial canopy cover. Agricultural and Forest Meteorology, 94(1), 13–29. "
        "https://doi.org/10.1016/S0168-1923(99)00005-2",
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
