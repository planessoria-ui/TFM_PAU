# -*- coding: utf-8 -*-
"""
Fase 13: incorporació de l'script de calibratge radiomètric al TFM.

  (a) Amplia el paràgraf que presenta l'script i hi insereix la Taula 9 amb
      els coeficients de cada banda.
  (b) Completa el peu de la Figura 14 (captura de l'script a Spyder).
  (c) Corregeix l'aplicació de la recta tèrmica: no es fa amb la calculadora
      ràster sinó com a coeficients de la banda 6 del mateix script.
  (d) Afegeix l'Annex A amb el codi complet en text seleccionable (Courier New).
  (e) Recalcula la numeració desada dels camps SEQ.
"""
import docx
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PATH = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v8.docx"
doc = docx.Document(PATH)
FONT = "Arial"
MONO = "Courier New"


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


def set_text(p, new):
    runs = p.runs
    if not runs:
        _font(p.add_run(new)); return
    runs[0].text = new
    for r in runs[1:]:
        r._element.getparent().remove(r._element)


def find(sub, style=None):
    for i, p in enumerate(doc.paragraphs):
        if sub in p.text and (style is None or p.style.name == style):
            return i
    raise LookupError(sub[:70])


def new_para_after(el, style=None):
    """Crea un paràgraf buit just després de l'element XML donat."""
    new = OxmlElement('w:p')
    el.addnext(new)
    p = docx.text.paragraph.Paragraph(new, doc.paragraphs[0]._parent)
    if style:
        p.style = doc.styles[style]
    return p


# ---------------------------------------------------------------------------
# (a) Paràgraf de presentació de l'script + Taula 9 de coeficients
# ---------------------------------------------------------------------------
i_script = find("s’utilitza un script de python mitjançant el software Spyder")
p_script = doc.paragraphs[i_script]
set_text(p_script, (
    "Conseqüentment, els coeficients d'aquestes rectes s'introdueixen en un "
    "script de Python, executat des de l'entorn Spyder, que recorre les imatges "
    "d'una carpeta i hi aplica la transformació banda per banda. L'script llegeix "
    "cada ortomosaic amb la biblioteca GDAL, multiplica els valors digitals de "
    "cada banda pel pendent corresponent, hi suma l'ordenada a l'origen i desa el "
    "resultat com un GeoTIFF de sis bandes en coma flotant que conserva la "
    "projecció i la geotransformació de l'original, identificat amb el sufix "
    "_Raf.tif. Els coeficients emprats es recullen a la Taula 9 i el codi complet "
    "es reprodueix a l'Annex A."
))

# -- peu de la taula (manual, Taula 9: el document arriba fins a la Taula 8) --
cap_t = new_para_after(p_script._element)
cap_t.alignment = WD_ALIGN_PARAGRAPH.LEFT
cap_t.paragraph_format.space_before = Pt(10)
cap_t.paragraph_format.space_after = Pt(4)
cap_t.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
_font(cap_t.add_run("Taula 9."), 10, bold=True)
_font(cap_t.add_run(
    " Coeficients de calibratge radiomètric introduïts a l'script per a cada "
    "banda de la càmera MicaSense Altum-PT. Les cinc primeres bandes converteixen "
    "el valor digital en reflectància (ρ = m · DN + n) i la sisena converteix el "
    "valor de la banda tèrmica en temperatura de superfície en graus Celsius "
    "(T = m · B6 + n)."), 10)

# -- taula --
rows = [
    ("Banda", "Longitud d'ona (nm)", "Pendent (m)", "Ordenada (n)", "Variable de sortida"),
    ("1", "475", "1,0 · 10⁻⁵", "−0,0007", "Reflectància (–)"),
    ("2", "560", "7,0 · 10⁻⁶", "+0,0077", "Reflectància (–)"),
    ("3", "668", "2,0 · 10⁻⁵", "+0,0130", "Reflectància (–)"),
    ("4", "717", "3,0 · 10⁻⁵", "+0,0013", "Reflectància (–)"),
    ("5", "840", "2,0 · 10⁻⁵", "−0,0017", "Reflectància (–)"),
    ("6", "LWIR (tèrmica)", "0,0132", "−374,37", "Temperatura (°C)"),
]
tbl = doc.add_table(rows=len(rows), cols=5)
tbl.style = doc.styles['Table Grid']
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
widths = [Cm(1.8), Cm(3.6), Cm(3.0), Cm(3.0), Cm(4.0)]
for r, data in enumerate(rows):
    for c, val in enumerate(data):
        cell = tbl.cell(r, c)
        cell.width = widths[c]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        _font(p.add_run(val), 10, bold=(r == 0))
cap_t._element.addnext(tbl._element)      # mou la taula just sota el peu

# -- nota de verificació --
note = new_para_after(tbl._element)
note.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
note.paragraph_format.space_before = Pt(6)
_font(note.add_run("[A SUBSTITUIR] "), 11, bold=True, color=RGBColor(0xC0, 0x00, 0x00))
_font(note.add_run(
    "Els valors de la Taula 9 provenen de la versió de l'script facilitada per "
    "l'IRTA, ajustada sobre el vol del 30 de juliol de 2026 d'una altra parcel·la "
    "(Manso Guilamany). Cal substituir-los pels coeficients obtinguts a les fulles "
    "CALIBRATION i Temp_CAL de cada vol d'aquest assaig abans de tancar la memòria.",
), 11, color=RGBColor(0xC0, 0x00, 0x00))

# ---------------------------------------------------------------------------
# (b) Peu de la Figura 14 (captura de l'script a Spyder)
# ---------------------------------------------------------------------------
i_f14 = find("Figura 14.")
p_f14 = doc.paragraphs[i_f14]
set_text(p_f14, "")
_font(p_f14.runs[0], 10, bold=True)
p_f14.runs[0].text = "Figura 14."
_font(p_f14.add_run(
    " Script de calibratge radiomètric obert a l'entorn Spyder. La llista "
    "corrections conté, per a cada banda de l'Altum-PT, el pendent i l'ordenada "
    "de la recta corresponent; el bucle les aplica a tots els píxels de cada "
    "ortomosaic de la carpeta de treball i escriu la imatge corregida amb el "
    "sufix _Raf.tif. El codi complet es reprodueix a l'Annex A."), 10)

# ---------------------------------------------------------------------------
# (c) Correcció: la recta tèrmica s'aplica dins del mateix script
# ---------------------------------------------------------------------------
i_term = find("L'ajust es repeteix per a cada vol")
p_term = doc.paragraphs[i_term]
t = p_term.text
old = ("introduint-la a la calculadora ràster sobre la banda 6 s'obté directament "
       "la capa tèrmica corregida, en graus Celsius")
new = ("en aplicar-la sobre la banda 6 s'obté directament la capa tèrmica "
       "corregida, en graus Celsius. A la pràctica aquesta operació no es fa amb "
       "la calculadora ràster de QGIS sinó com la sisena entrada de l'script de "
       "calibratge radiomètric (Taula 9 i Annex A), de manera que les cinc bandes "
       "espectrals i la tèrmica es corregeixen en una sola passada")
if old in t:
    set_text(p_term, t.replace(old, new))
    print("recta tèrmica: aplicació corregida")
else:
    print("AVÍS: no s'ha trobat la frase de la calculadora ràster")

# -- referència a l'annex allà on es presenta l'script --
i_asd = find("Els coeficients d’aquestes rectes es passen a l’script")
p_asd = doc.paragraphs[i_asd]
t = p_asd.text
old = "una altra per a sis (Altum-PT)."
new = "una altra per a sis (Altum-PT), que és la que s'utilitza aquí (Annex A)."
if old in t:
    set_text(p_asd, t.replace(old, new))
    print("referència a l'annex afegida al paràgraf de l'ASD")

# ---------------------------------------------------------------------------
# (d) Annex A amb el codi complet
# ---------------------------------------------------------------------------
h = doc.add_paragraph(style='Heading 1')
_font(h.add_run("Annex A. Script de calibratge radiomètric de les imatges de "
                "l'Altum-PT"), 14, bold=True)
h.paragraph_format.page_break_before = True

intro = doc.add_paragraph()
intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
_font(intro.add_run(
    "Es reprodueix l'script radiometric_calibration6Bandes.py, de l'equip de "
    "Teledetecció i Reg de l'IRTA, tal com s'utilitza en el processament descrit a "
    "l'apartat 4.6. L'únic element que cal editar entre vols són la ruta de la "
    "carpeta de treball (workdir), que ha d'apuntar a la subcarpeta Corrected dels "
    "ortomosaics del vol, i la llista corrections, que recull el pendent i "
    "l'ordenada de les sis rectes de calibratge d'aquell vol (Taula 9). "
    "L'script requereix els paquets numpy i GDAL i es llança des de l'entorn "
    "Spyder; per a cada fitxer .tif de la carpeta escriu un GeoTIFF de sis bandes "
    "en coma flotant amb el sufix _Raf.tif, que conserva la projecció i la "
    "geotransformació de l'original. Els comentaris i els missatges de consola es "
    "mantenen en l'idioma original del codi."), 12)

CODE = r'''import numpy as np
from osgeo import gdal
import os

# Directorio de trabajo / imagen a corregir
workdir = r'...\<AAAAMMDD>_<parcela>_AltumPT\Processed\Corrected'
os.chdir(workdir)

# Numero de bandas
bands_number = 6

# Coeficientes de correccion para las bandas
corrections = [
    [0.00001,  -0.0007],   # Banda 1  475 nm
    [0.000007, +0.0077],   # Banda 2  560 nm
    [0.00002,  +0.013],    # Banda 3  668 nm
    [0.00003,  +0.0013],   # Banda 4  717 nm
    [0.00002,  -0.0017],   # Banda 5  840 nm
    [0.0132,   -374.37]    # Banda 6  LWIR (termica)
]

# ========================================

for f in os.listdir():
    if not f.endswith(".tif"):
        continue

    # Abrir el archivo
    fid = gdal.Open(f)
    if fid is None:
        print(f"No se pudo abrir el archivo {f}")
        continue

    print(f"Procesando el archivo {f}")

    prj = fid.GetProjection()
    geo = fid.GetGeoTransform()

    band = fid.GetRasterBand(1).ReadAsArray()

    driver = gdal.GetDriverByName('GTiff')
    image = driver.Create(f[:-4] + '_Raf.tif', band.shape[1],
                          band.shape[0], bands_number, gdal.GDT_Float32)

    image.SetProjection(prj)
    image.SetGeoTransform(geo)

    # Aplicar correccion a cada banda
    for i in range(1, bands_number + 1):
        band = fid.GetRasterBand(i).ReadAsArray()
        print(f"Valores originales banda {i}: {band.min()}, {band.max()}")

        coef, offset = corrections[i - 1]
        nband = band * coef + offset

        # Verificar si hay NaN o valores incorrectos
        if np.isnan(nband).any():
            print(f"Se encontraron valores NaN en la banda {i}")

        print(f"Valores corregidos banda {i}: {nband.min()}, {nband.max()}")

        image.GetRasterBand(i).WriteArray(nband)

    image = None
    fid = None

    print(f"Imagen {f} corregida y guardada.")'''

for line in CODE.split("\n"):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = Pt(0)
    pf.left_indent = Cm(0.5)
    r = p.add_run(line if line else " ")
    _font(r, 9, name=MONO)
    r._element.find(qn('w:t')).set(qn('xml:space'), 'preserve')

# ---------------------------------------------------------------------------
# (e) Recàlcul de la numeració desada dels camps SEQ
# ---------------------------------------------------------------------------
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
      f"· SEQ figures: {counters['Figura']} taules: {counters['Taula']}")
