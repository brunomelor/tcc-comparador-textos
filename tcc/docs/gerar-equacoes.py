# -*- coding: utf-8 -*-
"""
Gera as equacoes (1), (2) e (3) como PNG (LaTeX-quality) para inclusao no artigo.
Saida: tcc/docs/img/eq-tfidf.png, eq-cosseno.png, eq-jaccard.png
"""
from __future__ import annotations

import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'img')
os.makedirs(OUT_DIR, exist_ok=True)

EQUACOES = [
    ('eq-tfidf.png',
     r'$\mathrm{tf\text{-}idf}(t, d) = \mathrm{tf}(t, d) \cdot \log\dfrac{N}{\mathrm{df}(t)}$'),
    ('eq-cosseno.png',
     r'$\cos(\theta) = \dfrac{\vec{A} \cdot \vec{B}}{\|\vec{A}\| \, \|\vec{B}\|} = \dfrac{\sum_{i=1}^{n} A_i B_i}{\sqrt{\sum_{i=1}^{n} A_i^2} \, \sqrt{\sum_{i=1}^{n} B_i^2}}$'),
    ('eq-jaccard.png',
     r'$J(A, B) = \dfrac{|A \cap B|}{|A \cup B|}$'),
]


def render(name: str, latex: str) -> None:
    fig = plt.figure(figsize=(7, 1.2), dpi=300)
    fig.text(0.5, 0.5, latex, ha='center', va='center', fontsize=18)
    out = os.path.join(OUT_DIR, name)
    plt.savefig(out, bbox_inches='tight', pad_inches=0.15, transparent=False)
    plt.close(fig)
    print(f'[OK] {out}')


def main() -> None:
    for name, latex in EQUACOES:
        render(name, latex)


if __name__ == '__main__':
    main()
