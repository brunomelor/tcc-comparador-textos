# -*- coding: utf-8 -*-
"""
Gera diagramas para o artigo:
- arquitetura.png: diagrama de componentes (camadas Cliente / API / NLP).
- sequencia.png: diagrama de sequencia do POST /api/compare.
"""
from __future__ import annotations

import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'img')
os.makedirs(OUT_DIR, exist_ok=True)


def caixa(ax, x: float, y: float, w: float, h: float, label: str,
          fill: str = '#dbeafe', stroke: str = '#1d4ed8', fontsize: int = 10) -> None:
    box = FancyBboxPatch((x, y), w, h,
                         boxstyle='round,pad=0.02,rounding_size=0.08',
                         linewidth=1.2, facecolor=fill, edgecolor=stroke)
    ax.add_patch(box)
    ax.text(x + w / 2, y + h / 2, label,
            ha='center', va='center', fontsize=fontsize, color='#0f172a')


def seta(ax, x1: float, y1: float, x2: float, y2: float, rotulo: str = '') -> None:
    arrow = FancyArrowPatch((x1, y1), (x2, y2),
                            arrowstyle='->', mutation_scale=16,
                            color='#334155', linewidth=1.2)
    ax.add_patch(arrow)
    if rotulo:
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2 + 0.10
        ax.text(mx, my, rotulo, ha='center', va='bottom',
                fontsize=8, color='#334155', style='italic')


# ===== ARQUITETURA =====
def arquitetura() -> None:
    fig, ax = plt.subplots(figsize=(9.5, 5), dpi=200)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.axis('off')

    # Camada Cliente
    ax.add_patch(FancyBboxPatch((0.2, 4.4), 9.6, 1.3, boxstyle='round,pad=0.02',
                                facecolor='#f0f9ff', edgecolor='#0369a1', linewidth=1, linestyle='--'))
    ax.text(0.4, 5.55, 'Camada de Apresentacao (Cliente)', fontsize=9.5,
            color='#0369a1', style='italic')
    caixa(ax, 0.8, 4.65, 2.2, 0.9, 'TextInput\n(React)', fill='#bae6fd', stroke='#0284c7')
    caixa(ax, 3.4, 4.65, 2.2, 0.9, 'Results\n(React)', fill='#bae6fd', stroke='#0284c7')
    caixa(ax, 6.0, 4.65, 2.2, 0.9, 'SharedTerms\n(React)', fill='#bae6fd', stroke='#0284c7')
    caixa(ax, 8.4, 4.85, 1.3, 0.55, 'Spinner', fill='#bae6fd', stroke='#0284c7', fontsize=9)

    # Camada API
    ax.add_patch(FancyBboxPatch((0.2, 2.7), 9.6, 1.3, boxstyle='round,pad=0.02',
                                facecolor='#fffbeb', edgecolor='#b45309', linewidth=1, linestyle='--'))
    ax.text(0.4, 3.85, 'Camada de API (FastAPI)', fontsize=9.5, color='#b45309', style='italic')
    caixa(ax, 1.2, 2.95, 3.0, 0.9, 'POST /api/compare\n(CompareRequest)',
          fill='#fde68a', stroke='#b45309')
    caixa(ax, 5.0, 2.95, 3.0, 0.9, 'GET /api/health',
          fill='#fde68a', stroke='#b45309')
    caixa(ax, 8.4, 3.15, 1.4, 0.55, 'CORS', fill='#fde68a', stroke='#b45309', fontsize=9)

    # Camada NLP
    ax.add_patch(FancyBboxPatch((0.2, 0.5), 9.6, 1.7, boxstyle='round,pad=0.02',
                                facecolor='#ecfdf5', edgecolor='#047857', linewidth=1, linestyle='--'))
    ax.text(0.4, 2.05, 'Camada de Processamento (PLN)', fontsize=9.5,
            color='#047857', style='italic')
    caixa(ax, 0.6, 0.75, 1.9, 1.05, 'Pre-processamento\nNLTK\n(tokens + RSLP)',
          fill='#a7f3d0', stroke='#047857', fontsize=9)
    caixa(ax, 2.7, 0.75, 1.9, 1.05, 'Vetorizacao\nscikit-learn\n(TF-IDF)',
          fill='#a7f3d0', stroke='#047857', fontsize=9)
    caixa(ax, 4.8, 0.75, 1.9, 1.05, 'Similaridade\nCosseno + Jaccard',
          fill='#a7f3d0', stroke='#047857', fontsize=9)
    caixa(ax, 6.9, 0.75, 2.7, 1.05, 'Mapeamento reverso\nstem -> palavra\n(interpretabilidade)',
          fill='#86efac', stroke='#047857', fontsize=9)

    # Setas entre camadas
    seta(ax, 5.0, 4.55, 5.0, 3.95, 'HTTPS / JSON')
    seta(ax, 2.7, 2.85, 2.7, 1.85, 'preprocess()')
    seta(ax, 5.7, 2.85, 5.7, 1.85, 'compare_texts()')
    plt.tight_layout()
    out = os.path.join(OUT_DIR, 'arquitetura.png')
    plt.savefig(out, bbox_inches='tight')
    plt.close(fig)
    print(f'[OK] {out}')


# ===== SEQUENCIA =====
def sequencia() -> None:
    fig, ax = plt.subplots(figsize=(10, 6), dpi=200)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.axis('off')

    atores = [
        ('Usuario', 0.7, '#7c3aed'),
        ('React\n(App.tsx)', 2.6, '#0284c7'),
        ('FastAPI\n(main.py)', 4.7, '#b45309'),
        ('TextComparer\n(nlp.py)', 7.0, '#047857'),
        ('NLTK / sklearn', 9.0, '#475569'),
    ]
    for nome, x, cor in atores:
        ax.add_patch(FancyBboxPatch((x - 0.55, 9.0), 1.1, 0.7,
                                    boxstyle='round,pad=0.02', facecolor='white', edgecolor=cor))
        ax.text(x, 9.35, nome, ha='center', va='center', fontsize=9, color=cor)
        ax.plot([x, x], [0.4, 9.0], color=cor, linewidth=0.6, linestyle=':', alpha=0.6)

    def msg(y, x1, x2, label, dashed=False):
        style = '-' if not dashed else '--'
        ax.annotate('', xy=(x2, y), xytext=(x1, y),
                    arrowprops={'arrowstyle': '->', 'linestyle': style, 'color': '#334155', 'linewidth': 1.0})
        ax.text((x1 + x2) / 2, y + 0.12, label, ha='center', va='bottom',
                fontsize=8.5, color='#0f172a')

    # Fluxo
    msg(8.5, 0.7, 2.6, '1: clica "Comparar" (text1, text2)')
    msg(7.9, 2.6, 4.7, '2: POST /api/compare {text1, text2}')
    msg(7.3, 4.7, 7.0, '3: compare_texts(text1, text2)')
    msg(6.7, 7.0, 9.0, '4: preprocess() — tokens, stopwords, RSLP')
    msg(6.1, 9.0, 7.0, '5: stems + stem_to_word', dashed=True)
    msg(5.5, 7.0, 9.0, '6: TfidfVectorizer.fit_transform([d1, d2])')
    msg(4.9, 9.0, 7.0, '7: matriz TF-IDF', dashed=True)
    msg(4.3, 7.0, 9.0, '8: cosine_similarity(v1, v2)')
    msg(3.7, 9.0, 7.0, '9: cos, jaccard', dashed=True)
    msg(3.1, 7.0, 4.7, '10: CompareResultDict', dashed=True)
    msg(2.5, 4.7, 2.6, '11: 200 OK { cosine, jaccard, shared_terms, ... }', dashed=True)
    msg(1.9, 2.6, 0.7, '12: render Results + SharedTerms', dashed=True)

    plt.tight_layout()
    out = os.path.join(OUT_DIR, 'sequencia.png')
    plt.savefig(out, bbox_inches='tight')
    plt.close(fig)
    print(f'[OK] {out}')


if __name__ == '__main__':
    arquitetura()
    sequencia()
