# -*- coding: utf-8 -*-
"""
Genera figures ORIGINALS (esquemes/diagrames) per al TFM, basades en els
conceptes descrits a la literatura citada. Totes són elaboració pròpia i
s'han d'atribuir com a tal, indicant la font conceptual a peu de figura.

Sortida: carpeta figures/*.png
"""
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch
import numpy as np

OUT = os.path.join(os.path.dirname(__file__), "figures")
os.makedirs(OUT, exist_ok=True)
plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 11})
DPI = 200


def fig1_tseb():
    fig, ax = plt.subplots(figsize=(7.2, 4.6))
    ax.set_xlim(0, 10); ax.set_ylim(0, 7); ax.axis('off')
    # Sòl
    ax.add_patch(mpatches.Rectangle((0, 0), 10, 1.4, color='#caa472'))
    ax.text(0.2, 0.55, 'SÒL', fontsize=10, weight='bold', color='#5c3d1a')
    # Dosser (vegetació)
    ax.add_patch(mpatches.Ellipse((5, 3.0), 3.6, 2.4, color='#3a7d34', alpha=0.85))
    ax.text(5, 3.0, 'DOSSER', ha='center', va='center', color='white', weight='bold')
    # Sol
    ax.add_patch(mpatches.Circle((1.2, 6.2), 0.5, color='#f2b21a'))
    ax.text(1.2, 6.2, '☼', ha='center', va='center', fontsize=14)
    # Rn (radiació neta) cap avall
    ax.add_patch(FancyArrowPatch((1.6, 5.9), (2.6, 4.4), arrowstyle='-|>',
                 mutation_scale=18, color='#d9a300', lw=2))
    ax.text(1.5, 5.2, 'Rn', color='#a37b00', weight='bold')
    # Fluxos del dosser: H i LE (transpiració)
    ax.add_patch(FancyArrowPatch((5.6, 4.0), (6.6, 5.6), arrowstyle='-|>',
                 mutation_scale=16, color='#c0392b', lw=2))
    ax.text(6.7, 5.4, 'Hc', color='#c0392b', weight='bold')
    ax.add_patch(FancyArrowPatch((4.4, 4.0), (3.4, 5.6), arrowstyle='-|>',
                 mutation_scale=16, color='#1f6fb2', lw=2))
    ax.text(2.9, 5.4, 'LEc\n(transpiració)', color='#1f6fb2', weight='bold',
            ha='center', fontsize=9)
    # Fluxos del sòl: Hs, LEs (evaporació), G
    ax.add_patch(FancyArrowPatch((8.2, 1.5), (8.8, 3.0), arrowstyle='-|>',
                 mutation_scale=14, color='#e67e22', lw=2))
    ax.text(8.9, 2.6, 'Hs', color='#e67e22', weight='bold')
    ax.add_patch(FancyArrowPatch((7.4, 1.5), (7.0, 3.0), arrowstyle='-|>',
                 mutation_scale=14, color='#2980b9', lw=2))
    ax.text(6.3, 2.6, 'LEs\n(evaporació)', color='#2980b9', weight='bold',
            ha='center', fontsize=9)
    ax.add_patch(FancyArrowPatch((9.0, 1.2), (9.0, 0.3), arrowstyle='-|>',
                 mutation_scale=14, color='#7f5539', lw=2))
    ax.text(9.15, 0.6, 'G', color='#7f5539', weight='bold')
    # Temperatures
    ax.text(5, 1.9, 'Tc', ha='center', color='white', fontsize=10, weight='bold')
    ax.text(2.0, 1.0, 'Ts', ha='center', color='#5c3d1a', fontsize=10, weight='bold')
    # Equació
    ax.text(5, 6.6, 'Rn = G + H + LE        (LE = LEc + LEs)',
            ha='center', fontsize=12, weight='bold')
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, 'fig1_tseb.png'), dpi=DPI, bbox_inches='tight')
    plt.close(fig)


def fig2_fipar_diurnal():
    fig, ax = plt.subplots(figsize=(6.8, 4.2))
    h = np.linspace(6, 18, 200)
    # patró diürn en cultiu en files: més interceptació al matí/tarda, mínim al migdia
    fipar = 0.55 - 0.18 * np.exp(-((h - 12) ** 2) / 6.0)
    ax.plot(h, fipar, color='#1f6fb2', lw=2.4,
            label='Corba diürna contínua (imatge hemisfèrica / GoPro)')
    # punt del ceptòmetre al migdia
    ax.plot(12, 0.55 - 0.18, 'o', color='#c0392b', ms=11,
            label='Mesura puntual al migdia (ceptòmetre)')
    # valor diari integrat
    fday = np.trapezoid(fipar, h) / (18 - 6)
    ax.axhline(fday, color='#2e7d32', ls='--', lw=1.8,
               label='fIPAR diari integrat')
    ax.annotate('Subestimació\nde la mesura puntual',
                xy=(12, 0.37), xytext=(13.6, 0.30),
                arrowprops=dict(arrowstyle='->', color='#c0392b'),
                fontsize=9, color='#c0392b')
    ax.set_xlabel('Hora del dia (h solar)'); ax.set_ylabel('fIPAR')
    ax.set_xlim(6, 18); ax.set_ylim(0.25, 0.65)
    ax.legend(fontsize=8.5, loc='upper center')
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, 'fig2_fipar_diurnal.png'), dpi=DPI, bbox_inches='tight')
    plt.close(fig)


def fig3_cwsi():
    fig, ax = plt.subplots(figsize=(6.6, 4.4))
    vpd = np.linspace(0.5, 5, 100)
    lower = 3 - 1.6 * vpd           # baseline inferior: sense estrès (ben regat)
    upper = np.full_like(vpd, 3.0)  # baseline superior: sense transpiració
    ax.plot(vpd, lower, color='#1f6fb2', lw=2.2, label='Límit inferior (sense estrès)')
    ax.plot(vpd, upper, color='#c0392b', lw=2.2, ls='--', label='Límit superior (no transpira)')
    ax.fill_between(vpd, lower, upper, color='#f1c40f', alpha=0.15)
    # punt mesurat
    x0 = 3.0; y_low = 3 - 1.6 * x0; y_up = 3.0; y_meas = y_low + 0.55 * (y_up - y_low)
    ax.plot(x0, y_meas, 'ko', ms=9)
    ax.annotate('(Tc − Ta) mesurat', xy=(x0, y_meas), xytext=(x0 + 0.3, y_meas + 0.6),
                fontsize=9)
    ax.annotate('', xy=(x0, y_low), xytext=(x0, y_up),
                arrowprops=dict(arrowstyle='<->', color='gray'))
    ax.text(x0 + 0.05, (y_low + y_up) / 2,
            'CWSI = (mesurat − inferior) /\n(superior − inferior)',
            fontsize=8.5, color='dimgray')
    ax.set_xlabel('Dèficit de pressió de vapor, VPD (kPa)')
    ax.set_ylabel('Tc − Ta (°C)')
    ax.legend(fontsize=9, loc='lower left'); ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, 'fig3_cwsi.png'), dpi=DPI, bbox_inches='tight')
    plt.close(fig)


def fig4_disseny():
    fig, ax = plt.subplots(figsize=(7.6, 4.4))
    ax.set_xlim(0, 11.5); ax.set_ylim(0, 6.6); ax.axis('off')
    # 5 tractaments (franges horitzontals) x 4 blocs (columnes), segons croquis ADAPTEX
    tracts = [
        ('SE',    'Sense esporga',                       '#6666ff'),
        ('EL',    'Esporga lleugera',                    '#ffc000'),
        ('EM',    'Esporga mitjana (comercial)',         '#a40079'),
        ('ES',    'Esporga severa',                      '#5b9bd5'),
        ('ES+DV', 'Esporga severa + desfullat a verolat','#92d050'),
    ]
    ax.text(5.75, 6.3, "Assaig ADAPTEX – Ull de Llebre (Tempranillo), Raimat",
            ha='center', fontsize=10.5, weight='bold')
    ax.text(5.75, 5.95, "5 tractaments × 4 blocs = 20 parcel·les (24 ceps/parcel·la)",
            ha='center', fontsize=9, style='italic', color='dimgray')
    # capçaleres de bloc
    for b in range(4):
        ax.text(3.0 + b * 2.0, 5.55, f'Bloc {b+1}', ha='center', fontsize=9, weight='bold')
    for ti, (code, name, color) in enumerate(tracts):
        y = 4.6 - ti * 0.95
        ax.text(2.05, y + 0.3, f'{code}', ha='right', va='center', fontsize=9.5,
                weight='bold')
        for b in range(4):
            x = 2.1 + b * 2.0
            ax.add_patch(FancyBboxPatch((x, y), 1.8, 0.6, boxstyle="round,pad=0.02",
                         fc=color, ec='black', alpha=0.85))
            ax.text(x + 0.9, y + 0.3, code, ha='center', va='center',
                    color='white', fontsize=8, weight='bold')
    handles = [mpatches.Patch(color=c, label=f'{code} – {name}')
               for code, name, c in tracts]
    ax.legend(handles=handles, loc='lower center', ncol=2, fontsize=8,
              bbox_to_anchor=(0.5, -0.16), frameon=False)
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, 'fig4_disseny.png'), dpi=DPI, bbox_inches='tight')
    plt.close(fig)


def fig5_workflow():
    fig, ax = plt.subplots(figsize=(6.4, 7.4))
    ax.set_xlim(0, 10); ax.set_ylim(0, 15); ax.axis('off')
    boxes = [
        (12.8, '#34495e', "VOLS DE DRON\n(càmera tèrmica + multiespectral)  +  MESURES DE CAMP\n(ceptòmetre, GoPro, Ψstem, meteo XAC)"),
        (10.4, '#2980b9', "PROCESSAMENT FOTOGRAMÈTRIC I D'IMATGES\nDSM/DTM → alçada i volum · NDVI → LAI, fIPAR\nsegmentació → Tc i Ts"),
        (8.0, '#16a085', "MODELS DE BALANÇ ENERGÈTIC\nTSEB → ETa        Shuttleworth–Wallace → ETp"),
        (5.6, '#8e44ad', "ÍNDEX D'ESTRÈS HÍDRIC\nCWSI = 1 − ETa/ETp"),
        (3.2, '#c0392b', "VALIDACIÓ\nfIPAR (dron vs. ceptòmetre vs. hemisfèrica)\nCWSI vs. Ψstem"),
        (0.9, '#7f8c8d', "ANÀLISI ESTADÍSTICA\nANOVA (tractament · data · bloc) + Tukey HSD"),
    ]
    for y, c, txt in boxes:
        ax.add_patch(FancyBboxPatch((0.6, y), 8.8, 1.7, boxstyle="round,pad=0.1",
                     fc=c, ec='black', alpha=0.9))
        ax.text(5.0, y + 0.85, txt, ha='center', va='center', color='white',
                fontsize=9.5, weight='bold')
    for i in range(len(boxes) - 1):
        y_top = boxes[i][0]; y_bot = boxes[i + 1][0] + 1.7
        ax.add_patch(FancyArrowPatch((5, y_top), (5, y_bot), arrowstyle='-|>',
                     mutation_scale=20, color='black', lw=1.6))
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, 'fig5_workflow.png'), dpi=DPI, bbox_inches='tight')
    plt.close(fig)


if __name__ == "__main__":
    fig1_tseb(); fig2_fipar_diurnal(); fig3_cwsi(); fig4_disseny(); fig5_workflow()
    print("Figures generades a", OUT)
