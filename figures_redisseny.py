# -*- coding: utf-8 -*-
"""
Redisseny de les figures del TFM amb un estil sobri de publicació:
paleta neutra, traços fins, sense blocs de color saturat ni text blanc sobre
fons cridaners. Inclou la correcció de contingut del diagrama de flux (la
càmera tèrmica és la banda 6 LWIR integrada a l'Altum-PT).

Sortides: figures/fig3_cwsi.png, figures/fig5_workflow.png, figures/fig8_pla_vol.png
"""
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, Rectangle
import numpy as np
from PIL import Image

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "figures")
SHOTS = "/tmp/claude-0/-home-user-TFM-PAU/446d5a3f-407d-5d51-8cea-50e8348ad898/images"
os.makedirs(OUT, exist_ok=True)

# paleta sòbria
INK = '#1a2230'
MUTED = '#5b6675'
RULE = '#b9c0c9'
FILL = '#f4f6f8'
ACCENT = '#1f5c8b'
WARM = '#9a6b1f'

plt.rcParams.update({
    'font.family': 'DejaVu Sans',
    'text.color': INK, 'axes.labelcolor': INK,
    'xtick.color': MUTED, 'ytick.color': MUTED,
    'axes.edgecolor': RULE, 'axes.linewidth': 0.9,
})


# ---------------------------------------------------------------------------
# Figura 3 — fonament del CWSI
# ---------------------------------------------------------------------------
def fig_cwsi(dims):
    fig, ax = plt.subplots(figsize=(7.2, 4.75))
    vpd = np.linspace(0.5, 5, 200)
    lower = 3 - 1.6 * vpd
    upper = np.full_like(vpd, 3.0)

    ax.fill_between(vpd, lower, upper, color=ACCENT, alpha=0.055, zorder=1)
    ax.plot(vpd, lower, color=ACCENT, lw=2.0, zorder=3,
            label='Límit inferior · cultiu ben regat')
    ax.plot(vpd, upper, color=WARM, lw=1.8, ls=(0, (6, 3)), zorder=3,
            label='Límit superior · dosser que no transpira')

    x0 = 3.0
    y_lo = 3 - 1.6 * x0
    y_up = 3.0
    y_m = y_lo + 0.55 * (y_up - y_lo)

    # segments a i b
    ax.annotate('', xy=(x0, y_m), xytext=(x0, y_lo),
                arrowprops=dict(arrowstyle='<->', color=MUTED, lw=1.0))
    ax.annotate('', xy=(x0, y_up), xytext=(x0, y_m),
                arrowprops=dict(arrowstyle='<->', color=MUTED, lw=1.0))
    ax.text(x0 - 0.12, (y_lo + y_m) / 2, 'a', ha='right', va='center',
            fontsize=11, style='italic', color=INK)
    ax.text(x0 - 0.12, (y_m + y_up) / 2, 'b', ha='right', va='center',
            fontsize=11, style='italic', color=INK)

    ax.plot([x0], [y_m], 'o', ms=7, color=INK, zorder=5)
    ax.annotate('(Tc − Ta) mesurat',
                xy=(x0 - 0.06, y_m), xytext=(2.70, y_m + 1.05),
                fontsize=9.5, color=INK, ha='right',
                arrowprops=dict(arrowstyle='-', color=MUTED, lw=0.8,
                                shrinkA=2, shrinkB=4))
    ax.text(4.92, 2.55, 'CWSI = a / (a + b)', ha='right', va='top',
            fontsize=10.5, color=INK, zorder=6,
            bbox=dict(boxstyle='round,pad=0.45', fc='white', ec=RULE, lw=0.8))

    ax.set_xlabel('Dèficit de pressió de vapor, VPD (kPa)', fontsize=10.5, labelpad=7)
    ax.set_ylabel('Tc − Ta (°C)', fontsize=10.5, labelpad=7)
    ax.set_xlim(0.5, 5); ax.set_ylim(-5.4, 3.6)
    ax.grid(True, color=RULE, lw=0.5, alpha=0.55)
    ax.set_axisbelow(True)
    for s in ('top', 'right'):
        ax.spines[s].set_visible(False)
    leg = ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.155), ncol=2,
                    fontsize=9, frameon=False, handlelength=2.4,
                    columnspacing=2.4)
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, 'fig3_cwsi.png'), dpi=220,
                bbox_inches='tight', facecolor='white')
    plt.close(fig)
    _resize('fig3_cwsi.png', dims)


# ---------------------------------------------------------------------------
# Figura 4 — flux de treball metodològic
# ---------------------------------------------------------------------------
def fig_workflow(dims):
    steps = [
        ("Adquisició de dades",
         "Vols de dron amb MicaSense Altum-PT (cinc bandes espectrals\n"
         "i banda tèrmica LWIR) i mesures de camp simultànies:\n"
         "ceptòmetre, potencial hídric de tija i meteorologia"),
        ("Processament fotogramètric",
         "Ortomosaics, model digital de superfície i del terreny;\n"
         "alçada del dosser, fracció de coberta i volum de copa"),
        ("Calibratge radiomètric",
         "Correcció de les cinc bandes espectrals amb els panells de\n"
         "referència i de la banda tèrmica amb el termòmetre d'infraroig"),
        ("Estimació del LAI",
         "Ceptòmetre i formulació de Campbell i Norman, validació amb\n"
         "ceps defoliats (R², RMSE) i extensió a tots els ceps\n"
         "mitjançant Random Forest"),
        ("Temperatures de component",
         "Temperatura de dosser (Tc) i de sòl (Ts) per discriminació\n"
         "automàtica de vegetació i sòl a partir de l'alçada i dels índexs"),
        ("Model TSEB-2T",
         "Formulació de Campbell i Norman per a cultius en filera;\n"
         "evapotranspiració real i partició transpiració / evaporació"),
        ("Validació",
         "Flux de saba, aigua de reg aplicada i potencial hídric de tija"),
    ]
    n = len(steps)
    gap = 0.34
    # alçada de cada caixa segons el nombre de línies del text
    heights = [0.70 + 0.30 * (b.count('\n') + 1) for _, b in steps]
    total = sum(heights) + gap * (n - 1) + 0.5

    fig, ax = plt.subplots(figsize=(7.0, 8.1))
    ax.set_xlim(0, 10); ax.set_ylim(0, total); ax.axis('off')

    y = total - 0.25
    for i, (title, body) in enumerate(steps):
        bh = heights[i]
        y -= bh
        ax.add_patch(Rectangle((0.75, y), 8.9, bh, fc=FILL, ec=RULE, lw=0.9, zorder=2))
        ax.add_patch(Rectangle((0.75, y), 0.085, bh, fc=ACCENT, ec='none', zorder=3))
        ax.text(1.12, y + bh - 0.26, f'{i+1:02d}', fontsize=8.5, color=ACCENT,
                weight='bold', va='top', zorder=4)
        ax.text(1.72, y + bh - 0.24, title, fontsize=10.2, color=INK,
                weight='bold', va='top', zorder=4)
        ax.text(1.72, y + bh - 0.56, body, fontsize=8.3, color=MUTED,
                va='top', linespacing=1.5, zorder=4)
        if i < n - 1:
            ax.add_patch(FancyArrowPatch((5.2, y), (5.2, y - gap + 0.03),
                         arrowstyle='-|>', mutation_scale=11,
                         color=RULE, lw=1.1, zorder=1))
            y -= gap
    fig.savefig(os.path.join(OUT, 'fig5_workflow.png'), dpi=220,
                bbox_inches='tight', facecolor='white')
    plt.close(fig)
    _resize('fig5_workflow.png', dims)


# ---------------------------------------------------------------------------
# Figura nova — pla de vol (dos panells de la captura del planificador)
# ---------------------------------------------------------------------------
def fig_pla_vol():
    adv = Image.open(os.path.join(SHOTS, '4.jpg')).convert('RGB')   # pestanya Advanced
    bas = Image.open(os.path.join(SHOTS, '5.jpg')).convert('RGB')   # pestanya Basic
    w, h = bas.size
    # retalla el panell de paràmetres (dreta) de cada captura
    left = int(w * 0.663)
    pa = bas.crop((left, 0, w, h))
    pb = adv.crop((left, 0, w, h))

    pad, lab, gap = 26, 34, 22
    pw, ph = pa.size
    canvas = Image.new('RGB', (pad * 2 + pw * 2 + gap, pad * 2 + lab + ph), 'white')
    canvas.paste(pa, (pad, pad + lab))
    canvas.paste(pb, (pad + pw + gap, pad + lab))

    fig = plt.figure(figsize=(canvas.width / 200, canvas.height / 200), dpi=200)
    ax = fig.add_axes([0, 0, 1, 1]); ax.imshow(canvas); ax.axis('off')
    W, H = canvas.size
    for x0, txt in ((pad, '(A)'),
                    (pad + pw + gap, '(B)')):
        ax.text(x0, pad + lab - 12, txt, fontsize=10.5, color=INK,
                weight='bold', va='bottom', ha='left')
    for x0 in (pad, pad + pw + gap):
        ax.add_patch(Rectangle((x0, pad + lab), pw - 1, ph - 1,
                     fill=False, ec=RULE, lw=1.0))
    ax.set_xlim(0, W); ax.set_ylim(H, 0)
    fig.savefig(os.path.join(OUT, 'fig8_pla_vol.png'), dpi=200,
                bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print('fig8_pla_vol.png', Image.open(os.path.join(OUT, 'fig8_pla_vol.png')).size)


def _resize(name, dims):
    p = os.path.join(OUT, name)
    im = Image.open(p).convert('RGB').resize(dims, Image.LANCZOS)
    im.save(p)
    print(name, im.size)


if __name__ == "__main__":
    fig_cwsi((1294, 855))
    fig_workflow((1254, 1453))
    fig_pla_vol()
