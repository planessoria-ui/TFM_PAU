# -*- coding: utf-8 -*-
"""
Fase 3: desa el resultat calculat dels camps SEQ dels peus de figura i taula,
perquè la numeració es vegi correctament encara que no s'actualitzin els camps
en obrir el document. El camp continua sent un camp de Word, de manera que
l'Índex de figures i l'Índex de taules el continuen recollint.
"""
import copy
import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

PATH = "/home/user/TFM_PAU/TFM_Pau_Planes_IRTA_v3.docx"
doc = docx.Document(PATH)

counters = {"Figura": 0, "Taula": 0}
fixed = 0

for p in doc.paragraphs:
    runs = p._element.findall(qn('w:r'))
    for i, r in enumerate(runs):
        instr = r.find(qn('w:instrText'))
        if instr is None or not instr.text or 'SEQ' not in instr.text:
            continue
        name = 'Figura' if 'Figura' in instr.text else ('Taula' if 'Taula' in instr.text else None)
        if name is None:
            continue
        # localitza el run amb fldChar 'end' posterior
        end_run = None
        for r2 in runs[i + 1:]:
            fc = r2.find(qn('w:fldChar'))
            if fc is not None and fc.get(qn('w:fldCharType')) == 'end':
                end_run = r2
                break
            if fc is not None and fc.get(qn('w:fldCharType')) == 'separate':
                end_run = None
                break
        if end_run is None:
            continue
        counters[name] += 1

        rpr = r.find(qn('w:rPr'))
        # run amb fldChar 'separate'
        r_sep = OxmlElement('w:r')
        if rpr is not None:
            r_sep.append(copy.deepcopy(rpr))
        fc_sep = OxmlElement('w:fldChar'); fc_sep.set(qn('w:fldCharType'), 'separate')
        r_sep.append(fc_sep)
        # run amb el número calculat
        r_val = OxmlElement('w:r')
        if rpr is not None:
            r_val.append(copy.deepcopy(rpr))
        t = OxmlElement('w:t'); t.text = str(counters[name])
        r_val.append(t)

        end_run.addprevious(r_sep)
        end_run.addprevious(r_val)
        fixed += 1

doc.save(PATH)
print(f"Numeració desada: {counters['Figura']} figures, {counters['Taula']} taules "
      f"({fixed} camps).")
