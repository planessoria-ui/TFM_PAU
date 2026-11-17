# -*- coding: utf-8 -*-
"""Consolida els quatre vols de 2026 i en resumeix els paràmetres estructurals."""
import csv, json, os, statistics as st

U = "/root/.claude/uploads/446d5a3f-407d-5d51-8cea-50e8348ad898/"
F = [("03/06", "a290cdc2-biophysical_traits_YIELD4CAST2026_covercrop_P25_v2.csv"),
     ("30/06", "57255596-biophysical_traits_YIELD4CAST2026_covercrop_P25.csv"),
     ("22/07", "7e8f3dc6-biophysical_traits_YIELD4CAST2026_covercrop_P25_22072026.csv"),
     ("12/08", "9b33ef42-biophysical_traits_YIELD4CAST2026_covercrop_P25_12082026.csv")]
SHORT = {"Esporga severa i desfullat a verolat": "ES+DV", "Esporga severa + desfullat": "ES+DV",
         "Esporga severa": "ES", "Esporga mitjana": "EM",
         "Esporga lleugera": "ELL", "Sense esporga": "C"}
ORDER = ["C", "ELL", "EM", "ES", "ES+DV"]

g = json.load(open(U + "7d51909c-MARC_PLANTACIO_ADAPTEX_v2.geojson", encoding="utf-8"))
BLOC = {f["properties"]["tree"]: f["properties"]["BLOC"] for f in g["features"]}


def f2(s):
    s = (s or "").strip().replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


data = []
for date, fn in F:
    raw = open(U + fn, encoding="utf-8-sig").read()
    d = ";" if raw.count(";") > raw.count(",") else ","
    for r in csv.DictReader(raw.splitlines(), delimiter=d):
        if not (r.get("tree") or "").strip():
            continue
        t = int(float(r["tree"]))
        rec = {"date": date, "tree": t, "trt": SHORT[r["Treatment"].strip()],
               "bloc": BLOC.get(t)}
        for k in ("hc", "fc", "fVolc", "wc_hc", "CANOPY_AREA", "CANOPY_VOLUME",
                  "Tr", "Tc", "Ts", "area_pixel", "CANOPY_NDVI", "CANOPY_SAVI",
                  "fc_covercrop", "SOIL_COUNT"):
            rec[k] = f2(r.get(k))
        data.append(rec)

print(f"registres: {len(data)}  ({len({(x['date'],x['tree']) for x in data})} combinacions cep-data)\n")

print("mida de píxel de cada vol")
for date, _ in F:
    px = [x["area_pixel"] for x in data if x["date"] == date and x["area_pixel"]]
    print(f"   {date}: {st.mean(px)**0.5*100:5.2f} cm   ({len(px)} ceps)")

for var, lab, fmt in [("hc", "alçada del dosser (m)", "6.2f"),
                      ("fc", "fracció de coberta", "6.3f"),
                      ("fVolc", "volum per unitat de sòl (m)", "6.3f"),
                      ("wc_hc", "relació amplada/alçada", "6.3f"),
                      ("Tc", "temperatura de dosser (°C)", "6.2f"),
                      ("Ts", "temperatura de sòl (°C)", "6.2f")]:
    print(f"\n{lab}")
    print("   trt     " + "".join(f"{d:>14s}" for d, _ in F))
    for t in ORDER:
        line = f"   {t:7s} "
        for date, _ in F:
            v = [x[var] for x in data if x["date"] == date and x["trt"] == t and x[var] is not None]
            line += f"{st.mean(v):{fmt}} ± {st.pstdev(v):4.2f}" if v else f"{'—':>14s}"
        print(line)
    allv = [x[var] for x in data if x[var] is not None]
    print(f"   global  min {min(allv):.3f}  max {max(allv):.3f}")

print("\nmàscara de l'entrefilera (fc_covercrop)")
for date, _ in F:
    v = [x["fc_covercrop"] for x in data if x["date"] == date and x["fc_covercrop"] is not None]
    z = sum(1 for x in data if x["date"] == date and x["SOIL_COUNT"] == 0)
    print(f"   {date}: mitjana {st.mean(v):.5f}  màx {max(v):.4f}  | ceps amb 0 píxels de sòl: {z}")
