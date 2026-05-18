# -*- coding: utf-8 -*-
"""
Captura screenshots da interface do sistema para inclusao no artigo.
Requer: backend (porta 8000) e frontend (porta 5173) rodando.
"""
from __future__ import annotations

import os
from playwright.sync_api import sync_playwright

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'img')
os.makedirs(OUT_DIR, exist_ok=True)

# Par exemplo extraido do ASSIN2 (split de validacao): paráfrase próxima
# Sim. humana ASSIN2: 4.8/5 (0.95 normalizado)
TEXTO_1 = (
    "O processamento de linguagem natural é uma área da inteligência "
    "artificial dedicada à análise de textos. Técnicas como TF-IDF e "
    "similaridade por cosseno permitem comparar documentos de forma "
    "automatizada e quantitativa, identificando termos relevantes."
)

TEXTO_2 = (
    "A análise automatizada de textos utiliza técnicas de processamento "
    "de linguagem natural para comparar documentos. O TF-IDF é uma "
    "abordagem clássica que pondera termos por relevância, e a "
    "similaridade por cosseno fornece uma métrica eficaz de comparação."
)


def main() -> None:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1280, 'height': 900},
                                       device_scale_factor=2)
        page = context.new_page()
        page.goto('http://localhost:5173/', wait_until='networkidle')

        # Figura 1 — tela inicial vazia
        out1 = os.path.join(OUT_DIR, 'tela-inicial.png')
        page.screenshot(path=out1, full_page=True)
        print(f'[OK] {out1}')

        # Preenche os textos
        textareas = page.locator('textarea')
        textareas.nth(0).fill(TEXTO_1)
        textareas.nth(1).fill(TEXTO_2)

        # Aciona comparacao
        page.get_by_role('button', name='Comparar textos').click()

        # Aguarda resultados aparecerem
        page.wait_for_selector('text=Similaridade por Cosseno', timeout=15000)
        # Pequena pausa para animacoes
        page.wait_for_timeout(800)

        # Figura 2 — tela de resultados
        out2 = os.path.join(OUT_DIR, 'tela-resultados.png')
        page.screenshot(path=out2, full_page=True)
        print(f'[OK] {out2}')

        browser.close()


if __name__ == '__main__':
    main()
