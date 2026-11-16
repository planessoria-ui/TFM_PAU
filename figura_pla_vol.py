# -*- coding: utf-8 -*-
"""
Figura del pla de vol en dos panells (A i B), cadascun amb la captura sencera
del planificador: A, pestanya de càmera i paràmetres de vol; B, pestanya de
solapaments i geometria de passada.
"""
import os
from PIL import Image, ImageDraw, ImageFont

SHOTS = "/tmp/claude-0/-home-user-TFM-PAU/446d5a3f-407d-5d51-8cea-50e8348ad898/images"
OUT = "/home/user/TFM_PAU/figures/fig8_pla_vol.png"

INK = (26, 34, 48)
RULE = (170, 178, 188)

a = Image.open(os.path.join(SHOTS, '5.jpg')).convert('RGB')   # Basic
b = Image.open(os.path.join(SHOTS, '4.jpg')).convert('RGB')   # Advanced

w, h = a.size
pad, lab, gap = 18, 40, 26
W = w + pad * 2
H = pad * 2 + (lab + h) * 2 + gap

canvas = Image.new('RGB', (W, H), 'white')
draw = ImageDraw.Draw(canvas)

try:
    font = ImageFont.truetype(
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 26)
except OSError:
    font = ImageFont.load_default()

for i, (img, label) in enumerate(((a, '(A)'), (b, '(B)'))):
    y = pad + i * (lab + h + gap)
    draw.text((pad, y + 4), label, fill=INK, font=font)
    canvas.paste(img, (pad, y + lab))
    draw.rectangle([pad, y + lab, pad + w - 1, y + lab + h - 1], outline=RULE, width=2)

canvas.save(OUT)
print("fig8_pla_vol.png", canvas.size)
