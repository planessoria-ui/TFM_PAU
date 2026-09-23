# -*- coding: utf-8 -*-
"""
Fase 7: reenfoca el treball sobre la premissa que tots els tractaments es
mantenen en un mateix estat hídric.

Idea que es trasllada al document:
  - El factor que varia entre tractaments és el maneig de capçada, no l'estrès.
  - El potencial hídric de tija és el criteri de decisió del reg: segons el
    valor de cada tractament s'aporta més o menys aigua (reg normal o RDC).
  - El CWSI passa a servir per verificar que l'estat hídric és equivalent i
    per copsar l'efecte de l'estructura de la capçada.

Es parteix de la versió de l'autor i només s'afegeix o es reformula el que
canvia aquest enfocament. La resta (imatges, taules, format) es conserva.
"""
import copy
import docx
from docx.oxml.ns import qn

SRC = "/root/.claude/uploads/446d5a3f-407d-5d51-8cea-50e8348ad898/5f8ef775-TFM_Pau_Planes_IRTA_v5_1_1.docx"
DST = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v6.docx"
doc = docx.Document(SRC)


def find(prefix, style=None):
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith(prefix) and (style is None or p.style.name == style):
            return i
    raise LookupError(prefix[:70])


def set_text(p, new):
    runs = p.runs
    if not runs:
        p.add_run(new); return
    runs[0].text = new
    for r in runs[1:]:
        r._element.getparent().remove(r._element)


def append(prefix, extra, style=None):
    """Afegeix text al final d'un paràgraf conservant-ne el format."""
    p = doc.paragraphs[find(prefix, style)]
    last = p.runs[-1]
    new_r = copy.deepcopy(last._element)
    last._element.addnext(new_r)
    run = docx.text.run.Run(new_r, p)
    run.text = extra
    return p


def swap(prefix, old, new, style=None):
    """Substitueix una cadena dins d'un paràgraf."""
    p = doc.paragraphs[find(prefix, style)]
    if old in p.text:
        set_text(p, p.text.replace(old, new))
    return p


# ---------------------------------------------------------------------------
# 1. Resums: fer explícita la premissa
# ---------------------------------------------------------------------------
append("Aquest Treball Final de Màster, desenvolupat",
       " Com que cada tractament es rega de manera individualitzada per mantenir-los "
       "tots en un mateix estat hídric, les diferències d'evapotranspiració entre "
       "tractaments s'atribueixen a la mida i a l'arquitectura de la capçada i no a un "
       "estrès hídric desigual. El potencial hídric de tija serveix per comprovar "
       "aquesta igualtat i per decidir si cal aportar més o menys aigua a cada "
       "tractament.")

append("Este Trabajo Final de Máster, desarrollado",
       " Como cada tratamiento se riega de forma individualizada para mantenerlos todos "
       "en un mismo estado hídrico, las diferencias de evapotranspiración entre "
       "tratamientos se atribuyen al tamaño y a la arquitectura de la copa y no a un "
       "estrés hídrico desigual. El potencial hídrico de tallo sirve para comprobar esa "
       "igualdad y para decidir si hay que aportar más o menos agua a cada tratamiento.")

append("This Master's Thesis, carried out",
       " Because each treatment is irrigated individually so that all of them stay at "
       "the same water status, differences in evapotranspiration between treatments are "
       "attributed to canopy size and architecture rather than to unequal water stress. "
       "Stem water potential is used to check this and to decide whether each treatment "
       "needs more or less water.")

# Paraules clau: estrès -> estat hídric
for p in doc.paragraphs:
    t = p.text.strip()
    if t.startswith("Paraules clau") and "estrès hídric" in t:
        set_text(p, t.replace("estrès hídric", "estat hídric"))
    elif t.startswith("Palabras clave") and "estrés hídrico" in t:
        set_text(p, t.replace("estrés hídrico", "estado hídrico"))
    elif t.startswith("Key words") and "water stress" in t:
        set_text(p, t.replace("water stress", "water status"))

# ---------------------------------------------------------------------------
# 2. Introducció: enunciar la premissa quan es presenta l'assaig
# ---------------------------------------------------------------------------
append("El treball s'emmarca en l'assaig de maneig de capçada",
       " Un tret central del disseny és que cada tractament es rega per separat per "
       "mantenir-los tots en un mateix estat hídric, de manera que el factor que varia "
       "entre tractaments és la mida de la capçada i no l'estrès hídric.")

swap("El treball s", "els indicadors d'estrès hídric", "els indicadors de l'estat hídric")

# ---------------------------------------------------------------------------
# 3. Marc teòric 2.8: el Ψstem com a criteri de decisió del reg
# ---------------------------------------------------------------------------
set_text(doc.paragraphs[find("2.8.", style="Heading 2")],
         "2.8. Indicadors de l'estat hídric: potencial hídric de tija i CWSI")

append("La mesura de referència de l'estat hídric de la planta",
       " A més de servir de referència per validar el CWSI, el Ψstem és l'indicador amb "
       "què es decideix la dosi de reg: segons el valor que marca cada tractament "
       "s'aporta més o menys aigua, amb reg normal o amb reg deficitari controlat. Com "
       "que en aquest treball tots els tractaments es condueixen al mateix Ψstem, el "
       "CWSI hi té un paper diferent de l'habitual: no serveix per classificar els "
       "tractaments per grau d'estrès, sinó per comprovar que l'estat hídric és "
       "equivalent i per copsar l'efecte que l'estructura de la capçada té sobre el "
       "balanç energètic.")

# ---------------------------------------------------------------------------
# 4. Objectius
# ---------------------------------------------------------------------------
set_text(doc.paragraphs[find("Quantificar l'índex d'estrès hídric del cultiu")],
         "Comprovar, amb l'índex d'estrès hídric del cultiu (CWSI = 1 − ETa/ETp) i amb "
         "les mesures de potencial hídric de tija (Ψstem), que els cinc tractaments es "
         "mantenen en un mateix estat hídric, de manera que les diferències d'ETa es "
         "puguin atribuir al maneig de capçada.")

# ---------------------------------------------------------------------------
# 5. Materials i mètodes
# ---------------------------------------------------------------------------
append("La vinya, plantada l'any 2013",
       " A partir d'aquestes lectures es decideix si cada tractament ha de rebre més o "
       "menys aigua, amb reg normal o amb reg deficitari controlat segons el moment del "
       "cicle.")

append("Cada tractament es rega de manera independent",
       " El Ψstem és, doncs, el criteri de decisió del reg: si un tractament se "
       "n'allunya, se li ajusta la dosi fins a retornar-lo al valor comú.")

append("Potencial hídric de tija (Ψstem): determinat al migdia",
       " Aquestes lectures són també el criteri per ajustar la dosi de reg de cada "
       "tractament i mantenir-los tots en el mateix estat hídric.")

append("A partir d'aquestes dues estimacions",
       " Cal tenir present que, com que el reg individualitzat manté tots els "
       "tractaments en un mateix estat hídric, no s'esperen diferències de CWSI "
       "atribuïbles a un estrès desigual. El CWSI s'utilitza aquí per verificar aquesta "
       "equivalència i per copsar l'efecte de l'estructura de la capçada sobre la "
       "relació ETa/ETp.")

# ---------------------------------------------------------------------------
# 6. Resultats i discussió
# ---------------------------------------------------------------------------
set_text(doc.paragraphs[find("5.4.", style="Heading 2")],
         "5.4. Estat hídric dels tractaments: potencial hídric de tija i CWSI")

set_text(doc.paragraphs[find("CWSI per tractament; regressió CWSI")],
         "CWSI per tractament i regressió CWSI–Ψstem. Comproveu primer que el Ψstem és "
         "equivalent entre tractaments, cosa que confirmaria que l'estat hídric s'ha "
         "igualat, i valoreu després quina part del CWSI s'explica per la mida i "
         "l'estructura de la capçada.")

swap("[CAPÍTOL A OMPLIR] Interpreteu els resultats",
     "(iv) la relació CWSI–Ψstem;",
     "(iv) si el Ψstem i el CWSI confirmen que l'estat hídric s'ha mantingut equivalent "
     "entre tractaments i què implica això per interpretar les diferències d'ETa;")

doc.save(DST)
print("Fase 7 completada ->", DST)
