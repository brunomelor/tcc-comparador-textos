# -*- coding: utf-8 -*-
"""
Validacao empirica do sistema sobre o corpus ASSIN2 (PROPOR 2020).
- Baixa ASSIN2 via Hugging Face datasets.
- Roda pipeline NLP local (sem rede) sobre N pares.
- Computa Pearson, Spearman, MAE, MSE entre score do sistema e julgamento humano.
- Compara cosseno vs Jaccard.
- Gera grafico de dispersao e quadro de resultados.

Saidas:
- tcc/docs/img/dispersao-assin2.png
- tcc/docs/_validacao_assin2.json (estatisticas para uso no gerar-tcc.py)
"""
from __future__ import annotations

import json
import os
import sys
import statistics
import random
from typing import Any

# Importa o pipeline NLP do backend
BACKEND_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src', 'backend'))
sys.path.insert(0, BACKEND_DIR)

from nlp import compare_texts  # type: ignore

from datasets import load_dataset  # type: ignore
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from scipy.stats import pearsonr, spearmanr  # type: ignore

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'img')
os.makedirs(OUT_DIR, exist_ok=True)
JSON_OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_validacao_assin2.json')

# Para reproducibilidade
random.seed(42)


def normaliza_humano(score_1_a_5: float) -> float:
    """Converte escala 1-5 do ASSIN2 para 0-1, coerente com cosseno/Jaccard."""
    return (score_1_a_5 - 1.0) / 4.0


def main() -> None:
    print('[1/4] Carregando ASSIN2 (split validation, 500 pares)...')
    ds = load_dataset('nilc-nlp/assin2', split='validation')

    # Amostra 100 pares aleatorios estratificados (para varrer toda a faixa)
    indices = list(range(len(ds)))
    random.shuffle(indices)
    sampled = sorted(indices[:100])

    print(f'[2/4] Rodando pipeline NLP sobre {len(sampled)} pares...')
    cos_calc: list[float] = []
    jac_calc: list[float] = []
    hum: list[float] = []
    tempos: list[float] = []

    for idx in sampled:
        ex = ds[idx]
        premise = ex['premise']
        hypothesis = ex['hypothesis']
        human_score = float(ex['relatedness_score'])

        result = compare_texts(premise, hypothesis)
        cos_calc.append(float(result['cosine_similarity']))
        jac_calc.append(float(result['jaccard_coefficient']))
        hum.append(normaliza_humano(human_score))
        tempos.append(float(result['stats']['processing_time_ms']))

    print('[3/4] Computando correlacoes e erros...')
    pearson_cos, p_pcos = pearsonr(cos_calc, hum)
    spearman_cos, p_scos = spearmanr(cos_calc, hum)
    pearson_jac, p_pjac = pearsonr(jac_calc, hum)
    spearman_jac, p_sjac = spearmanr(jac_calc, hum)

    mae_cos = sum(abs(c - h) for c, h in zip(cos_calc, hum)) / len(hum)
    mse_cos = sum((c - h) ** 2 for c, h in zip(cos_calc, hum)) / len(hum)
    mae_jac = sum(abs(j - h) for j, h in zip(jac_calc, hum)) / len(hum)
    mse_jac = sum((j - h) ** 2 for j, h in zip(jac_calc, hum)) / len(hum)

    stats = {
        'n_pares': len(hum),
        'cosseno': {
            'pearson': round(pearson_cos, 4),
            'pearson_p': round(p_pcos, 6),
            'spearman': round(spearman_cos, 4),
            'spearman_p': round(p_scos, 6),
            'mae': round(mae_cos, 4),
            'mse': round(mse_cos, 4),
            'media': round(statistics.mean(cos_calc), 4),
            'desvio': round(statistics.stdev(cos_calc), 4),
        },
        'jaccard': {
            'pearson': round(pearson_jac, 4),
            'pearson_p': round(p_pjac, 6),
            'spearman': round(spearman_jac, 4),
            'spearman_p': round(p_sjac, 6),
            'mae': round(mae_jac, 4),
            'mse': round(mse_jac, 4),
            'media': round(statistics.mean(jac_calc), 4),
            'desvio': round(statistics.stdev(jac_calc), 4),
        },
        'humano': {
            'media': round(statistics.mean(hum), 4),
            'desvio': round(statistics.stdev(hum), 4),
        },
        'tempo_ms': {
            'mediana': round(statistics.median(tempos), 1),
            'media': round(statistics.mean(tempos), 1),
            'max': round(max(tempos), 1),
        },
    }

    print(json.dumps(stats, indent=2, ensure_ascii=False))

    with open(JSON_OUT, 'w', encoding='utf-8') as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)
    print(f'[OK] Estatisticas salvas em {JSON_OUT}')

    print('[4/4] Gerando grafico de dispersao...')
    fig, axes = plt.subplots(1, 2, figsize=(10, 4.2), dpi=200)
    for ax, vals, label, color in [
        (axes[0], cos_calc, 'Cosseno', '#2563eb'),
        (axes[1], jac_calc, 'Jaccard', '#d97706'),
    ]:
        ax.scatter(hum, vals, alpha=0.55, s=22, c=color, edgecolors='white', linewidth=0.4)
        ax.plot([0, 1], [0, 1], 'k--', alpha=0.35, linewidth=0.8, label='diagonal (y = x)')
        ax.set_xlabel('Similaridade humana (ASSIN2, 0-1)')
        ax.set_ylabel(f'Similaridade {label} (sistema)')
        ax.set_xlim(-0.02, 1.02)
        ax.set_ylim(-0.02, 1.02)
        r = pearson_cos if label == 'Cosseno' else pearson_jac
        ax.set_title(f'{label}  (Pearson r = {r:.3f})', fontsize=11)
        ax.grid(True, alpha=0.25)
        ax.legend(loc='upper left', fontsize=8)
    plt.tight_layout()
    out = os.path.join(OUT_DIR, 'dispersao-assin2.png')
    plt.savefig(out, bbox_inches='tight')
    plt.close(fig)
    print(f'[OK] {out}')


if __name__ == '__main__':
    main()
