# -*- coding: utf-8 -*-
"""
Composició de les fotografies de camp i de laboratori en figures de dos
panells, amb l'etiqueta (A) i (B) a sobre de cada imatge.

  fig_calibratge_irt.png   (A) panell negre   (B) vegetació
  fig_lai_destructiu.png   (A) fulles numerades  (B) mesurador d'àrea foliar
  fig_gps_ceps.png         aixecament GPS, una sola imatge
"""
import os
from PIL import Image, ImageDraw, ImageFont

SRC = "/tmp/claude-0/-home-user-TFM-PAU/446d5a3f-407d-5d51-8cea-50e8348ad898/images"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "figures")
os.makedirs(OUT, exist_ok=True)

INK = (26, 34, 48)
RULE = (190, 196, 204)
PAD, LAB, GAP = 16, 54, 20

try:
    FONT = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 34)
except OSError:
    FONT = ImageFont.load_default()


def ample_fix(im, w):
    """Escala la imatge a una amplada donada conservant la proporció."""
    h = int(round(im.height * w / im.width))
    return im.resize((w, h), Image.LANCZOS)


def alt_fixa(im, h):
    w = int(round(im.width * h / im.height))
    return im.resize((w, h), Image.LANCZOS)


def dos_panells(a, b, nom, amplada=2200, alinea="alt"):
    """Dues imatges de costat. alinea='alt' iguala l'alçada; 'ample', l'amplada."""
    if alinea == "alt":
        h = min(a.height, b.height)
        a, b = alt_fixa(a, h), alt_fixa(b, h)
        esc = (amplada - 2 * PAD - GAP) / (a.width + b.width)
        a = ample_fix(a, int(a.width * esc))
        b = ample_fix(b, int(b.width * esc))
    W = PAD * 2 + a.width + b.width + GAP
    H = PAD * 2 + LAB + max(a.height, b.height)
    cv = Image.new("RGB", (W, H), "white")
    dr = ImageDraw.Draw(cv)
    for im, x, txt in ((a, PAD, "(A)"), (b, PAD + a.width + GAP, "(B)")):
        dr.text((x, PAD + 4), txt, fill=INK, font=FONT)
        cv.paste(im, (x, PAD + LAB))
        dr.rectangle([x, PAD + LAB, x + im.width - 1, PAD + LAB + im.height - 1],
                     outline=RULE, width=2)
    p = os.path.join(OUT, nom)
    cv.save(p, quality=92)
    print(nom, cv.size)


def un_panell(a, nom, amplada=1500):
    a = ample_fix(a, amplada)
    cv = Image.new("RGB", (a.width + PAD * 2, a.height + PAD * 2), "white")
    cv.paste(a, (PAD, PAD))
    ImageDraw.Draw(cv).rectangle(
        [PAD, PAD, PAD + a.width - 1, PAD + a.height - 1], outline=RULE, width=2)
    p = os.path.join(OUT, nom)
    cv.save(p, quality=92)
    print(nom, cv.size)


op = lambda f: Image.open(os.path.join(SRC, f)).convert("RGB")

dos_panells(op("22.jpg"), op("23.jpg"), "fig_calibratge_irt.png")
dos_panells(op("24.jpg"), op("25.jpg"), "fig_lai_destructiu.png")
un_panell(op("26.jpg"), "fig_gps_ceps.png")

# abans i després del desfullat del tractament ES+DV
dos_panells(op("27.jpg"), op("28.jpg"), "fig_desfullat_abans_despres.png")
