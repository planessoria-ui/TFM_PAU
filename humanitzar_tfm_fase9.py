# -*- coding: utf-8 -*-
"""
Fase 9: refà la figura del flux de treball metodològic perquè reflecteixi la
metodologia TSEB-2T (sense l'apartat d'anàlisi estadística) i la substitueix
dins del document conservant-ne les dimensions exactes.
També ajusta els textos que encara parlaven d'anàlisi estadística.
"""
import io
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch
from PIL import Image
import docx

PATH = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v7.docx"
TARGET = "/word/media/image16.png"
DIMS = (1254, 1453)          # mides originals, per no deformar la imatge
plt.rcParams.update({'font.family': 'DejaVu Sans'})

# ---------------------------------------------------------------------------
# 1. Nova figura
# ---------------------------------------------------------------------------
fig, ax = plt.subplots(figsize=(6.4, 7.4))
ax.set_xlim(0, 10); ax.set_ylim(0, 15); ax.axis('off')

boxes = [
    (12.8, '#34495e', "VOLS DE DRON\ncàmera tèrmica (FLIR) + multiespectral (Altum-PT)\n+ MESURES DE CAMP (ceptòmetre, Ψstem, meteo)"),
    (10.4, '#2980b9', "PROCESSAMENT FOTOGRAMÈTRIC I D'IMATGES\nDSM/DTM → alçada, fracció de coberta i volum de copa\níndexs de vegetació"),
    (8.0, '#16a085', "ESTIMACIÓ DEL LAI\nceptòmetre i Campbell & Norman → validació amb ceps\ndefoliats (R², RMSE) → Random Forest per a tots els ceps"),
    (5.6, '#8e44ad', "TEMPERATURES DE COMPONENT\nTc i Ts per discriminació automàtica\nvegetació/sòl (alçada i índexs)"),
    (3.2, '#c0392b', "MODEL TSEB-2T\nformulació de Campbell & Norman per a cultius\nen filera (hedgerow) → ETa i partició T/E"),
    (0.9, '#7f8c8d', "VALIDACIÓ\nflux de saba · aigua de reg aplicada · Ψstem"),
]
for y, c, txt in boxes:
    ax.add_patch(FancyBboxPatch((0.6, y), 8.8, 1.7, boxstyle="round,pad=0.1",
                 fc=c, ec='black', alpha=0.9))
    ax.text(5.0, y + 0.85, txt, ha='center', va='center', color='white',
            fontsize=8.6, weight='bold')
for i in range(len(boxes) - 1):
    ax.add_patch(FancyArrowPatch((5, boxes[i][0]), (5, boxes[i + 1][0] + 1.7),
                 arrowstyle='-|>', mutation_scale=20, color='black', lw=1.6))

buf = io.BytesIO()
fig.savefig(buf, format='png', dpi=200, bbox_inches='tight')
plt.close(fig)
buf.seek(0)
img = Image.open(buf).convert('RGB').resize(DIMS, Image.LANCZOS)
out = io.BytesIO(); img.save(out, format='PNG'); blob = out.getvalue()
open("/home/user/TFM_PAU/figures/fig5_workflow.png", "wb").write(blob)
print("figura refeta:", img.size, len(blob), "bytes")

# ---------------------------------------------------------------------------
# 2. Substitució dins del document
# ---------------------------------------------------------------------------
doc = docx.Document(PATH)
done = False
for rel in doc.part.rels.values():
    if 'image' in rel.reltype and str(rel.target_part.partname) == TARGET:
        rel.target_part._blob = blob
        done = True
print("imatge substituïda:", done)


def set_text(p, new):
    runs = p.runs
    if not runs:
        p.add_run(new); return
    runs[0].text = new
    for r in runs[1:]:
        r._element.getparent().remove(r._element)


# ---------------------------------------------------------------------------
# 3. Textos que encara parlaven d'anàlisi estadística
# ---------------------------------------------------------------------------
for p in doc.paragraphs:
    t = p.text
    if t.strip().startswith("El treball s") and "capítol 2" in t:
        t = (t.replace(", els models d'ET i l’anàlisi estadística", " i els models d'ET")
               .replace(", els models d'ET i l'anàlisi estadística", " i els models d'ET")
               .replace("els indicadors d'estrès hídric", "els indicadors de l'estat hídric"))
        set_text(p, t)
    elif t.strip().startswith("La Figura 4 resumeix el flux de treball"):
        set_text(p,
                 "La Figura 4 resumeix el flux de treball seguit en aquest TFM, des de "
                 "l'adquisició de dades amb dron i en camp fins a l'estimació de "
                 "l'evapotranspiració amb el model TSEB-2T i la seva validació. A "
                 "continuació es descriu cada etapa.")
    elif p.style.name == 'Caption' and t.strip().startswith("Figura 4"):
        set_text(p, "")
        from docx.shared import Pt
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        r1 = p.add_run("Figura "); r1.font.size = Pt(10); r1.font.bold = True; r1.font.name = 'Arial'
        rb = p.add_run(); rb.font.size = Pt(10); rb.font.bold = True; rb.font.name = 'Arial'
        fc = OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'), 'begin'); rb._element.append(fc)
        ri = p.add_run(); ri.font.size = Pt(10); ri.font.bold = True; ri.font.name = 'Arial'
        it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve')
        it.text = ' SEQ Figura \\* ARABIC '; ri._element.append(it)
        rs = p.add_run(); fs = OxmlElement('w:fldChar'); fs.set(qn('w:fldCharType'), 'separate'); rs._element.append(fs)
        rv = p.add_run("4"); rv.font.size = Pt(10); rv.font.bold = True; rv.font.name = 'Arial'
        re_ = p.add_run(); fe = OxmlElement('w:fldChar'); fe.set(qn('w:fldCharType'), 'end'); re_._element.append(fe)
        r2 = p.add_run(". Esquema general del flux de treball metodològic del TFM, des de "
                       "l'adquisició de dades fins a l'estimació de l'ETa amb el model "
                       "TSEB-2T i la seva validació. ")
        r2.font.size = Pt(10); r2.font.name = 'Arial'
        r3 = p.add_run("Font: elaboració pròpia.")
        r3.font.size = Pt(10); r3.font.italic = True; r3.font.name = 'Arial'

doc.save(PATH)
print("Fase 9 completada.")
