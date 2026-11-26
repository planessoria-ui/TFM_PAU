# -*- coding: utf-8 -*-
"""
Base de treball compactada: els quatre fitxers biophysical_traits en un sol
full, amb els noms de tractament normalitzats, la parcel·la elemental extreta
del camp Nom i un full de mitjanes per tractament i vol.

Sortida: base_biophysical_2026.xlsx i base_biophysical_2026.csv
"""
import os
import numpy as np
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
B = "/tmp/claude-0/-home-user-TFM-PAU/446d5a3f-407d-5d51-8cea-50e8348ad898/bio2/"

F = [("03/06/2026", "a4e318f7-biophysical_traits_YIELD4CAST2026_covercrop_P25_v2.csv"),
     ("30/06/2026", "38aaf1c1-biophysical_traits_YIELD4CAST2026_covercrop_P25.csv"),
     ("22/07/2026", "93507a78-biophysical_traits_YIELD4CAST2026_covercrop_P25_22072026.csv"),
     ("13/08/2026", "272820cb-biophysical_traits_YIELD4CAST2026_covercrop_P25_12082026.csv")]
TR = {"Esporga severa + desfullat": "ES+DV", "Esporga severa i desfullat a verolat": "ES+DV",
      "Esporga severa": "ES", "Esporga mitjana": "EM",
      "Esporga lleugera": "ELL", "Sense esporga": "C"}
ORDRE = ["ES+DV", "ES", "EM", "ELL", "C"]
TEXT = ("Nom", "Treatment", "date", "Unnamed: 4")

# variables que alimenten el model TSEB-2T
INPUT_TSEB = ["fc", "hc", "wc_hc", "Tr", "Tc", "Ts"]
# la resta de variables estructurals i espectrals que es conserven
ESTRUCT = ["wc", "rc", "fVolc", "CANOPY_AREA", "CANOPY_VOLUME", "area_pixel",
           "CANOPY_COUNT", "SOIL_COUNT", "fc_covercrop", "SOIL_AREA_covercrop"]
ESPECTRE = [f"CANOPY_{i}" for i in ("NDVI", "SAVI", "MTVI", "NDWI", "CIG")]


def num(s):
    """Converteix a nombre acceptant coma o punt com a separador decimal."""
    if isinstance(s, float):
        return s
    s = str(s).strip()
    if not s or s.lower() in ("nan", "none"):
        return np.nan
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return np.nan


taules = []
for data, fn in F:
    df = pd.read_csv(B + fn, delimiter=";", dtype=str)
    df = df[df["tree"].notna()].copy()
    for c in df.columns:
        if c not in TEXT:
            df[c] = df[c].map(num)
    df["vol"] = data
    df["data_fitxer"] = df["date"]
    df["t"] = df["Treatment"].str.strip().map(TR)
    df["parcela"] = df["Nom"].str.split(".").str[0]
    df["tree"] = df["tree"].astype(int)
    taules.append(df)

big = pd.concat(taules, ignore_index=True)
COLS = (["vol", "data_fitxer", "tree", "Nom", "parcela", "t", "Treatment",
         "marco_plantacion"] + INPUT_TSEB + ESTRUCT + ESPECTRE
        + [f"{c}_std" for c in ESPECTRE])
big = big[[c for c in COLS if c in big.columns]]
big.to_csv(os.path.join(BASE, "base_biophysical_2026.csv"), index=False)

VOLS = [d for d, _ in F]
resum = (big.groupby(["t", "vol"])[INPUT_TSEB + ["wc", "fVolc"] + ESPECTRE]
            .mean().reset_index())

with pd.ExcelWriter(os.path.join(BASE, "base_biophysical_2026.xlsx"),
                    engine="openpyxl") as w:
    big.to_excel(w, sheet_name="Ceps", index=False)
    for v in INPUT_TSEB + ["wc", "fVolc", "CANOPY_NDVI"]:
        (resum.pivot(index="t", columns="vol", values=v)
              .loc[ORDRE, VOLS].round(4)
              .to_excel(w, sheet_name="Mitjanes", startrow=w.sheets.get("Mitjanes").max_row + 2
                        if "Mitjanes" in w.sheets else 0))

print(f"base_biophysical_2026: {len(big)} files, {big.vol.nunique()} vols, "
      f"{big.tree.nunique()} ceps, {big.parcela.nunique()} parcel·les")
print()
print("Mitjanes per vol de les entrades del TSEB-2T:")
print(big.groupby("vol")[INPUT_TSEB].mean().loc[VOLS].round(4).to_string())
print()
print("Valors diferents sobre 118 ceps:")
print(big.groupby("vol")[INPUT_TSEB].nunique().loc[VOLS].to_string())
print()
print("Data que consta dins de cada fitxer:")
for v in VOLS:
    print(f"   {v} -> {sorted(set(big[big.vol == v].data_fitxer))[0]}")
