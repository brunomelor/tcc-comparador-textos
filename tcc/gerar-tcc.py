# -*- coding: utf-8 -*-
"""
Gerador do TCC - Desenvolvimento de sistema de comparação de conteúdo de textos
Bruno Lorencatto - Engenharia de Software - UNINTER
"""

from __future__ import annotations

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import logging
import os
import re


# =============================================================================
# CONFIGURAÇÕES E HELPERS
# =============================================================================

def configurar_pagina(doc: Document) -> None:
    """Configura margens e tamanho da página conforme ABNT."""
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(3)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(3)
    section.right_margin = Cm(2)
    # Numeração de página no canto superior direito
    header = section.header
    header.is_linked_to_previous = False
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run()
    run.font.size = Pt(10)
    run.font.name = 'Arial'
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._element.append(fldChar1)
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = ' PAGE '
    run._element.append(instrText)
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    run._element.append(fldChar2)
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run._element.append(fldChar3)


def estilo_padrao(doc: Document) -> None:
    """Configura estilos base do documento."""
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)
    font.color.rgb = RGBColor(0, 0, 0)
    pf = style.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = Cm(1.25)


def add_titulo_artigo(doc: Document) -> None:
    """Adiciona título do artigo centralizado."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_after = Pt(18)
    run = p.add_run(
        'DESENVOLVIMENTO DE SISTEMA DE COMPARAÇÃO DE CONTEÚDO DE TEXTOS'
    )
    run.bold = True
    run.font.size = Pt(14)
    run.font.name = 'Arial'


def add_autores(doc: Document) -> None:
    """Adiciona informações dos autores."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run('Bruno Lorencatto')
    run.font.size = Pt(12)
    run.font.name = 'Arial'

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.paragraph_format.first_line_indent = Cm(0)
    p2.paragraph_format.space_after = Pt(24)
    run2 = p2.add_run('Centro Universitário Internacional – UNINTER')
    run2.font.size = Pt(10)
    run2.font.name = 'Arial'


def add_secao_primaria(doc: Document, texto: str) -> None:
    """Seção nível 1: MAIÚSCULA + NEGRITO (ex: 1 INTRODUÇÃO)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(36)
    p.paragraph_format.space_after = Pt(18)
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run(texto.upper())
    run.bold = True
    run.font.size = Pt(12)
    run.font.name = 'Arial'


def add_secao_secundaria(doc: Document, texto: str) -> None:
    """Seção nível 2: MAIÚSCULA SEM NEGRITO (ex: 2.1 SEÇÃO)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(24)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run(texto.upper())
    run.bold = False
    run.font.size = Pt(12)
    run.font.name = 'Arial'


def add_secao_terciaria(doc: Document, texto: str) -> None:
    """Seção nível 3: minúscula + negrito (ex: 2.1.1 Seção terciária)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run(texto)
    run.bold = True
    run.font.size = Pt(12)
    run.font.name = 'Arial'


def add_paragrafo(doc: Document, texto: str, recuo: bool = True) -> None:
    """Adiciona parágrafo de texto normal."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if not recuo:
        p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run(texto)
    run.font.size = Pt(12)
    run.font.name = 'Arial'


def add_placeholder(doc: Document, texto: str) -> None:
    """Adiciona texto placeholder em vermelho."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(f'[{texto}]')
    run.font.size = Pt(12)
    run.font.name = 'Arial'
    run.font.color.rgb = RGBColor(255, 0, 0)
    run.italic = True


def add_resumo(doc: Document, texto: str, palavras_chave: str) -> None:
    """Adiciona seção de Resumo com formatação específica (Arial 10, simples)."""
    add_secao_primaria(doc, 'Resumo')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run(texto)
    run.font.size = Pt(10)
    run.font.name = 'Arial'

    p2 = doc.add_paragraph()
    p2.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p2.paragraph_format.first_line_indent = Cm(0)
    p2.paragraph_format.space_before = Pt(12)
    run_label = p2.add_run('Palavras-chave: ')
    run_label.bold = True
    run_label.font.size = Pt(10)
    run_label.font.name = 'Arial'
    run_kw = p2.add_run(palavras_chave)
    run_kw.bold = False
    run_kw.font.size = Pt(10)
    run_kw.font.name = 'Arial'


def add_referencias_formatadas(doc: Document, referencias: list[str]) -> None:
    """Adiciona referências com formatação ABNT (espaçamento simples, títulos em negrito)."""
    add_secao_primaria(doc, 'Referências')
    for ref in referencias:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_after = Pt(12)
        # Parse **bold** markup para títulos em negrito (ABNT)
        parts = re.split(r'(\*\*.*?\*\*)', ref)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                run = p.add_run(part[2:-2])
                run.bold = True
            else:
                run = p.add_run(part)
            run.font.size = Pt(12)
            run.font.name = 'Arial'


def add_tabela(doc: Document, titulo: str, cabecalhos: list[str], linhas: list[list[str]], fonte: str = "Elaborado pelo autor (2026)") -> None:
    """Adiciona tabela com formatação ABNT (legenda acima, fonte abaixo)."""
    # Legenda acima
    p_titulo = doc.add_paragraph()
    p_titulo.paragraph_format.space_before = Pt(12)
    p_titulo.paragraph_format.first_line_indent = Cm(0)
    run_t = p_titulo.add_run(titulo)
    run_t.font.size = Pt(10)
    run_t.font.name = 'Arial'

    # Tabela
    table = doc.add_table(rows=1 + len(linhas), cols=len(cabecalhos))
    table.style = 'Table Grid'

    # Cabeçalho
    for i, cab in enumerate(cabecalhos):
        cell = table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(cab)
        run.bold = True
        run.font.size = Pt(10)
        run.font.name = 'Arial'

    # Linhas de dados
    for r, linha in enumerate(linhas):
        for c, valor in enumerate(linha):
            cell = table.rows[r + 1].cells[c]
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(str(valor))
            run.font.size = Pt(10)
            run.font.name = 'Arial'

    # Fonte abaixo
    p_fonte = doc.add_paragraph()
    p_fonte.paragraph_format.first_line_indent = Cm(0)
    p_fonte.paragraph_format.space_after = Pt(12)
    run_f = p_fonte.add_run(f'Fonte: {fonte}')
    run_f.font.size = Pt(10)
    run_f.font.name = 'Arial'


# =============================================================================
# REFERÊNCIAS BIBLIOGRÁFICAS (ABNT)
# =============================================================================

REFERENCIAS = [
    'BIRD, S.; KLEIN, E.; LOPER, E. **Natural Language Processing with Python**. Sebastopol: O\'Reilly Media, 2009.',
    'DEVLIN, J. et al. BERT: Pre-training of deep bidirectional transformers for language understanding. In: ANNUAL CONFERENCE OF THE NORTH AMERICAN CHAPTER OF THE ASSOCIATION FOR COMPUTATIONAL LINGUISTICS (NAACL-HLT), 2019, Minneapolis. **Proceedings** [...]. Minneapolis: Association for Computational Linguistics, 2019. p. 4171-4186.',
    'GIL, A. C. **Como elaborar projetos de pesquisa**. 6. ed. São Paulo: Atlas, 2017.',
    'GOMAA, W. H.; FAHMY, A. A. A survey of text similarity approaches. **International Journal of Computer Applications**, v. 68, n. 13, p. 13-18, 2013.',
    'JURAFSKY, D.; MARTIN, J. H. **Speech and Language Processing**. 3. ed. Stanford: [s. n.], 2024. Disponível em: https://web.stanford.edu/~jurafsky/slp3/. Acesso em: 20 fev. 2026.',
    'LAN, H. Research on Text Similarity Measurement Hybrid Algorithm with Term Semantic Information and TF-IDF Method. **Advances in Multimedia**, v. 2022, p. 1-11, 2022.',
    'MANNING, C. D.; RAGHAVAN, P.; SCHÜTZE, H. **Introduction to Information Retrieval**. Cambridge: Cambridge University Press, 2008.',
    'ORENGO, V. M.; HUYCK, C. R. A stemming algorithm for the Portuguese language. In: INTERNATIONAL SYMPOSIUM ON STRING PROCESSING AND INFORMATION RETRIEVAL (SPIRE), 8., 2001, Laguna de San Rafael. **Proceedings** [...]. Laguna de San Rafael: [s. n.], 2001. p. 186-193.',
    'PRESSMAN, R. S.; MAXIM, B. R. **Engenharia de Software**: uma abordagem profissional. 9. ed. Porto Alegre: AMGH, 2021.',
    'RAMOS, J. Using TF-IDF to determine word relevance in document queries. In: FIRST INSTRUCTIONAL CONFERENCE ON MACHINE LEARNING, 2003, Piscataway. **Proceedings** [...]. Piscataway: [s. n.], 2003. p. 133-142.',
    'REIMERS, N.; GUREVYCH, I. Sentence-BERT: Sentence Embeddings using Siamese BERT-Networks. In: CONFERENCE ON EMPIRICAL METHODS IN NATURAL LANGUAGE PROCESSING (EMNLP), 2019, Hong Kong. **Proceedings** [...]. Hong Kong: Association for Computational Linguistics, 2019. p. 3982-3992.',
    'SALTON, G.; BUCKLEY, C. Term-weighting approaches in automatic text retrieval. **Information Processing & Management**, v. 24, n. 5, p. 513-523, 1988.',
    'SANTOS, D. Processamento de linguagem natural: uma apresentação. In: CONGRESSO DA SOCIEDADE BRASILEIRA DE COMPUTAÇÃO, 22., 2002, Florianópolis. **Anais** [...]. Florianópolis: SBC, 2002. p. 59-82.',
    'SILVA, E. L.; MENEZES, E. M. **Metodologia da pesquisa e elaboração de dissertação**. 4. ed. Florianópolis: UFSC, 2005.',
    'SOMMERVILLE, I. **Engenharia de Software**. 10. ed. São Paulo: Pearson, 2019.',
    'SUNG, C. et al. A survey of text similarity approaches. **International Journal of Research and Analytical Reviews**, v. 6, n. 2, p. 342-348, 2019.',
    'VIEIRA, R.; LIMA, V. L. S. Linguística computacional: princípios e aplicações. In: JORNADA DE ATUALIZAÇÃO EM INTELIGÊNCIA ARTIFICIAL, 20., 2001, São Leopoldo. **Anais** [...]. São Leopoldo: SBC, 2001. p. 47-88.',
]


# =============================================================================
# SEÇÕES DO ARTIGO
# =============================================================================

def secao_resumo(doc: Document) -> None:
    """RESUMO."""
    add_resumo(
        doc,
        'A crescente produção de conteúdo digital em ambientes acadêmicos, '
        'jurídicos e corporativos intensifica a demanda por ferramentas capazes '
        'de identificar automaticamente as inter-relações entre documentos textuais. '
        'O presente trabalho tem como objetivo o desenvolvimento de um sistema web '
        'de comparação de conteúdo de textos que calcula e apresenta o índice de '
        'correlação entre dois documentos, empregando técnicas de processamento de '
        'linguagem natural e engenharia de software. A relevância do estudo reside '
        'na construção de uma ferramenta transparente, que explicita não apenas o '
        'grau de similaridade, mas também os elementos textuais que o fundamentam. '
        'A metodologia adota uma pesquisa aplicada, de abordagem quantitativa e '
        'caráter exploratório, combinando levantamento bibliográfico com '
        'desenvolvimento experimental. O sistema é implementado em Python com '
        'FastAPI para o back-end, utilizando NLTK para pré-processamento textual '
        'e scikit-learn para vetorização TF-IDF e cálculo de similaridade por '
        'cosseno e coeficiente de Jaccard, com front-end em React, TypeScript e '
        'Tailwind CSS. Os resultados demonstram que a combinação TF-IDF com '
        'similaridade por cosseno produz índices coerentes com a similaridade '
        'real dos documentos, e a interface desenvolvida permite ao usuário '
        'compreender os fatores que sustentam a correlação identificada.',
        'similaridade textual; processamento de linguagem natural; TF-IDF; '
        'engenharia de software; comparação de documentos.'
    )


def secao_introducao(doc: Document) -> None:
    """1 INTRODUÇÃO."""
    add_secao_primaria(doc, '1 Introdução')

    # Motivação / Problematização
    add_paragrafo(doc,
        'O crescimento exponencial da produção de conteúdo digital nas últimas '
        'décadas trouxe consigo o desafio de identificar relações e '
        'sobreposições entre documentos de forma automatizada. Em contextos '
        'acadêmicos, jurídicos e corporativos, a necessidade de verificar a '
        'originalidade de textos, detectar reutilização de conteúdo e mensurar '
        'a correlação entre documentos tornou-se uma demanda recorrente '
        '(GOMAA; FAHMY, 2013). Segundo Manning, Raghavan e Schütze (2008), a '
        'recuperação de informação e a análise de similaridade textual '
        'constituem pilares fundamentais da ciência da computação moderna, '
        'com aplicações que vão desde motores de busca até sistemas de '
        'detecção de plágio.'
    )

    add_paragrafo(doc,
        'Apesar da existência de ferramentas comerciais para comparação '
        'de textos, muitas delas são proprietárias, operam como caixas-pretas '
        'e não permitem ao usuário compreender os mecanismos que sustentam os '
        'resultados apresentados. Além disso, soluções voltadas ao português '
        'frequentemente carecem de tratamento adequado das particularidades '
        'morfológicas da língua, como a riqueza de flexões verbais e a '
        'variação de sufixos (SANTOS, 2002). Esse cenário evidencia a '
        'necessidade de sistemas transparentes que explicitem não apenas o '
        'índice de correlação, mas também os elementos textuais que o '
        'fundamentam.'
    )

    # Pergunta de pesquisa
    add_paragrafo(doc,
        'Diante desse contexto, a pergunta que orienta este trabalho é: '
        'como desenvolver um sistema web capaz de comparar o conteúdo de '
        'dois textos e apresentar, de forma clara e interpretável, o índice '
        'de correlação entre eles, utilizando técnicas de processamento de '
        'linguagem natural e métricas de similaridade textual?'
    )

    # Objetivo geral
    add_paragrafo(doc,
        'O objetivo geral deste trabalho é desenvolver um sistema web de '
        'comparação de conteúdo de textos que calcule e apresente o índice '
        'de correlação entre dois documentos, empregando técnicas de '
        'processamento de linguagem natural e engenharia de software.'
    )

    # Objetivos específicos
    add_paragrafo(doc,
        'Para alcançar o objetivo geral, são definidos os seguintes '
        'objetivos específicos: (i) investigar as principais técnicas de '
        'similaridade textual disponíveis na literatura, identificando suas '
        'vantagens e limitações; (ii) projetar a arquitetura do sistema, '
        'definindo plataforma, linguagem de programação e bibliotecas '
        'adequadas ao processamento de textos em português; (iii) implementar '
        'o módulo de processamento responsável pela vetorização dos textos '
        'e pelo cálculo das métricas de similaridade; e (iv) desenvolver a '
        'interface web (front-end) que permita ao usuário carregar documentos '
        'e visualizar os resultados da comparação.'
    )

    # Justificativa
    add_paragrafo(doc,
        'A relevância deste trabalho reside na contribuição para a área de '
        'engenharia de software aplicada ao processamento de linguagem '
        'natural, ao demonstrar o processo completo de concepção, projeto e '
        'implementação de um sistema funcional. Do ponto de vista prático, '
        'o sistema desenvolvido pode ser utilizado como ferramenta auxiliar '
        'em atividades de revisão textual, análise de originalidade e '
        'identificação de inter-relações entre documentos. Do ponto de '
        'vista acadêmico, o trabalho contribui ao documentar as decisões '
        'de engenharia envolvidas na escolha de plataforma, linguagem e '
        'algoritmos, servindo como referência para projetos similares.'
    )

    # Estrutura do trabalho
    add_paragrafo(doc,
        'Este artigo está organizado da seguinte forma: a Seção 2 '
        'apresenta a fundamentação teórica, abordando conceitos de '
        'processamento de linguagem natural, representação vetorial de '
        'textos, métricas de similaridade e engenharia de software para '
        'sistemas web. A Seção 3 descreve a metodologia adotada, '
        'incluindo a classificação da pesquisa, as etapas de desenvolvimento '
        'e as tecnologias selecionadas. A Seção 4 apresenta os resultados '
        'e discussões, detalhando a implementação do sistema e os testes de '
        'validação. Por fim, a Seção 5 traz as considerações finais e '
        'sugestões para trabalhos futuros.'
    )


def secao_fundamentacao(doc: Document) -> None:
    """2 FUNDAMENTAÇÃO TEÓRICA."""
    add_secao_primaria(doc, '2 Similaridade textual e engenharia de sistemas')

    add_paragrafo(doc,
        'A comparação automatizada de textos constitui um dos desafios centrais da '
        'ciência da computação contemporânea, situando-se na interseção entre o '
        'Processamento de Linguagem Natural (PLN), a recuperação de informação e a '
        'engenharia de software (MANNING; RAGHAVAN; SCHÜTZE, 2008). Este capítulo '
        'apresenta os conceitos teóricos fundamentais que sustentam o '
        'desenvolvimento de um sistema de comparação de conteúdo textual, '
        'abrangendo desde as técnicas clássicas de representação de documentos '
        'até as abordagens modernas baseadas em aprendizado profundo.'
    )

    # --- 2.1 ---
    add_secao_secundaria(doc, '2.1 Processamento de Linguagem Natural')

    add_paragrafo(doc,
        'O Processamento de Linguagem Natural (PLN) é uma subárea da inteligência '
        'artificial que se dedica ao estudo e desenvolvimento de métodos '
        'computacionais para a análise, compreensão e geração de línguas humanas '
        '(JURAFSKY; MARTIN, 2024). Segundo Vieira e Lima (2001), o PLN envolve '
        'um conjunto de técnicas computacionais motivadas teoricamente para a '
        'representação e análise da linguagem natural, permitindo que computadores '
        'processem dados textuais de forma significativa.'
    )

    add_paragrafo(doc,
        'O processamento de textos em linguagem natural envolve diversas etapas '
        'fundamentais. A tokenização consiste na segmentação do texto em unidades '
        'menores, como palavras ou sentenças. A remoção de stopwords elimina '
        'palavras de alta frequência e baixo valor semântico, como artigos e '
        'preposições. A stemização reduz as palavras a seus radicais, sendo que '
        'para o português destaca-se o algoritmo RSLP, proposto por Orengo e '
        'Huyck (2001), que foi desenvolvido especificamente para as '
        'particularidades morfológicas da língua portuguesa. A lematização, por '
        'sua vez, transforma as palavras em suas formas canônicas, preservando '
        'melhor o significado original (BIRD; KLEIN; LOPER, 2009).'
    )

    add_paragrafo(doc,
        'Santos (2002) destaca que o PLN aplicado ao português enfrenta desafios '
        'próprios, como a riqueza morfológica da língua e a ambiguidade lexical, '
        'que demandam abordagens específicas distintas das empregadas para a '
        'língua inglesa. Essas particularidades influenciam diretamente a '
        'qualidade dos resultados de sistemas de comparação textual desenvolvidos '
        'para documentos em português.'
    )

    # --- 2.2 ---
    add_secao_secundaria(doc, '2.2 Representação vetorial de textos')

    add_paragrafo(doc,
        'Para que algoritmos computacionais possam processar e comparar textos, '
        'é necessário representá-los em formato numérico. O modelo de espaço '
        'vetorial, introduzido por Salton e Buckley (1988), estabelece a base '
        'para a maioria das abordagens de representação textual ao converter '
        'documentos em vetores numéricos dentro de um espaço multidimensional.'
    )

    add_paragrafo(doc,
        'O modelo Bag of Words (BoW) representa cada documento como um vetor '
        'de frequências de termos, desconsiderando a ordem das palavras e a '
        'estrutura gramatical. Embora simples, essa abordagem apresenta '
        'limitações ao ignorar o contexto semântico dos termos (MANNING; '
        'RAGHAVAN; SCHÜTZE, 2008).'
    )

    add_paragrafo(doc,
        'A técnica TF-IDF (Term Frequency – Inverse Document Frequency) '
        'aprimora o modelo BoW ao ponderar a relevância de cada termo. O '
        'componente TF mede a frequência do termo no documento, enquanto o IDF '
        'penaliza termos que aparecem em muitos documentos da coleção, '
        'valorizando assim palavras mais discriminativas (RAMOS, 2003). '
        'Conforme Lan (2022), o TF-IDF permanece como uma das técnicas mais '
        'utilizadas em sistemas de recuperação de informação e comparação '
        'textual, sendo particularmente eficaz quando combinado com métricas '
        'de similaridade como o cosseno.'
    )

    add_paragrafo(doc,
        'Abordagens mais recentes utilizam embeddings — representações vetoriais '
        'densas que capturam relações semânticas entre palavras. Modelos como '
        'BERT (Bidirectional Encoder Representations from Transformers), proposto '
        'por Devlin et al. (2019), representam um avanço significativo ao '
        'considerar o contexto bidirecional das palavras. Reimers e Gurevych '
        '(2019) propuseram o Sentence-BERT (SBERT), uma modificação do BERT '
        'que utiliza redes siamesas para gerar embeddings de sentenças '
        'semanticamente significativos, reduzindo drasticamente o tempo de '
        'comparação entre pares de textos.'
    )

    # --- 2.3 ---
    add_secao_secundaria(doc, '2.3 Métricas de similaridade textual')

    add_paragrafo(doc,
        'As métricas de similaridade textual quantificam o grau de semelhança '
        'entre dois documentos ou trechos de texto. Gomaa e Fahmy (2013) '
        'classificam essas abordagens em três categorias: baseadas em cadeias '
        'de caracteres (string-based), baseadas em corpus (corpus-based) e '
        'baseadas em conhecimento (knowledge-based).'
    )

    add_paragrafo(doc,
        'A similaridade por cosseno é a métrica mais amplamente utilizada em '
        'sistemas de comparação textual. Calcula-se o cosseno do ângulo entre '
        'dois vetores no espaço multidimensional, resultando em um valor entre '
        '0 (nenhuma similaridade) e 1 (textos idênticos). Sua popularidade '
        'deve-se à independência em relação ao tamanho dos documentos, uma vez '
        'que avalia a direção dos vetores e não sua magnitude (MANNING; '
        'RAGHAVAN; SCHÜTZE, 2008).'
    )

    add_paragrafo(doc,
        'O coeficiente de Jaccard mede a similaridade entre dois conjuntos '
        'como a razão entre a interseção e a união dos elementos. Embora seja '
        'intuitivo e de fácil implementação, sua limitação reside na '
        'incapacidade de capturar relações semânticas, restringindo-se à '
        'comparação de palavras presentes nos textos (SUNG et al., 2019).'
    )

    add_paragrafo(doc,
        'Métricas baseadas em embeddings, como a similaridade por cosseno '
        'aplicada a vetores gerados pelo BERT ou SBERT, permitem capturar a '
        'similaridade semântica entre textos mesmo quando utilizam vocabulários '
        'distintos. Conforme Reimers e Gurevych (2019), o SBERT possibilita a '
        'comparação de 10.000 sentenças em aproximadamente 5 segundos, enquanto '
        'o BERT original demandaria cerca de 65 horas para a mesma tarefa, '
        'evidenciando ganhos significativos de desempenho.'
    )

    # --- 2.4 ---
    add_secao_secundaria(doc, '2.4 Sistemas de detecção de similaridade')

    add_paragrafo(doc,
        'Os sistemas de detecção de similaridade textual possuem aplicações '
        'diversas, sendo a detecção de plágio a mais conhecida. Esses sistemas '
        'operam, em geral, por meio de um processo que envolve o '
        'pré-processamento do texto de entrada, a fragmentação em unidades '
        'comparáveis, a representação vetorial dos fragmentos e a aplicação de '
        'métricas de similaridade contra uma base de referência (GOMAA; FAHMY, '
        '2013).'
    )

    add_paragrafo(doc,
        'Ferramentas comerciais como Turnitin e Copyscape utilizam bases de '
        'dados extensas para comparação, enquanto sistemas acadêmicos como o '
        'MOSS (Measure of Software Similarity), desenvolvido na Universidade '
        'de Stanford, destacam-se na comparação de código-fonte. No contexto '
        'de documentos textuais, a abordagem mais comum combina TF-IDF com '
        'similaridade por cosseno, podendo ser complementada por técnicas de '
        'fingerprinting baseadas em n-gramas (MANNING; RAGHAVAN; SCHÜTZE, 2008).'
    )

    add_paragrafo(doc,
        'Sung et al. (2019) ressaltam que a escolha da técnica de similaridade '
        'depende do objetivo da aplicação: sistemas focados em plágio literal '
        'beneficiam-se de abordagens baseadas em cadeias de caracteres, enquanto '
        'aplicações que necessitam identificar paráfrases ou reorganizações de '
        'conteúdo requerem métricas de similaridade semântica. A combinação de '
        'múltiplas técnicas pode oferecer resultados mais robustos.'
    )

    # --- 2.5 ---
    add_secao_secundaria(doc, '2.5 Engenharia de software para sistemas web')

    add_paragrafo(doc,
        'O desenvolvimento de sistemas de comparação textual como aplicação web '
        'requer a aplicação de princípios consolidados da engenharia de software. '
        'Segundo Pressman e Maxim (2021), a engenharia de software engloba '
        'métodos, ferramentas e procedimentos que possibilitam o controle do '
        'processo de desenvolvimento e oferecem uma base para a construção de '
        'software de alta qualidade de maneira produtiva.'
    )

    add_paragrafo(doc,
        'Sommerville (2019) destaca que aplicações web possuem características '
        'específicas que influenciam seu processo de desenvolvimento, tais como: '
        'a necessidade de interfaces de usuário intuitivas, requisitos de '
        'desempenho para processamento em tempo real, e a separação clara entre '
        'camadas de apresentação, lógica de negócios e acesso a dados. A '
        'arquitetura cliente-servidor, predominante no desenvolvimento web, '
        'permite que o processamento computacionalmente intensivo — como o '
        'cálculo de similaridade textual — seja executado no servidor, enquanto '
        'o front-end se encarrega da interação com o usuário.'
    )

    add_paragrafo(doc,
        'A escolha da plataforma tecnológica para o sistema deve considerar '
        'fatores como a disponibilidade de bibliotecas de PLN, o desempenho no '
        'processamento de textos e a facilidade de integração entre back-end e '
        'front-end. A linguagem Python destaca-se nesse contexto por seu '
        'ecossistema robusto de bibliotecas voltadas ao processamento de '
        'linguagem natural, como NLTK e spaCy, além de frameworks científicos '
        'como scikit-learn que implementam nativamente algoritmos de TF-IDF e '
        'métricas de similaridade (BIRD; KLEIN; LOPER, 2009).'
    )


def secao_metodologia(doc: Document) -> None:
    """3 METODOLOGIA."""
    add_secao_primaria(doc, '3 Metodologia')

    add_paragrafo(doc,
        'Este capítulo descreve os procedimentos metodológicos adotados para o '
        'desenvolvimento do sistema de comparação de conteúdo de textos, '
        'abrangendo a classificação da pesquisa, as etapas de desenvolvimento '
        'e as ferramentas tecnológicas selecionadas.'
    )

    # --- 3.1 ---
    add_secao_secundaria(doc, '3.1 Classificação da pesquisa')

    add_paragrafo(doc,
        'Quanto à natureza, esta pesquisa classifica-se como aplicada, pois '
        'busca gerar conhecimentos dirigidos à solução de um problema prático '
        'específico: a comparação automatizada de conteúdo textual entre dois '
        'documentos (GIL, 2017). O produto final consiste em um sistema '
        'funcional capaz de calcular e apresentar o índice de correlação entre '
        'os textos analisados.'
    )

    add_paragrafo(doc,
        'Quanto à abordagem, o trabalho adota um enfoque quantitativo, uma vez '
        'que os resultados da comparação textual são expressos por meio de '
        'métricas numéricas — como índices de similaridade por cosseno e '
        'coeficientes de Jaccard — que permitem mensurar objetivamente o grau '
        'de correlação entre os documentos (GIL, 2017).'
    )

    add_paragrafo(doc,
        'Em relação aos objetivos, a pesquisa caracteriza-se como exploratória, '
        'visto que investiga diferentes técnicas de similaridade textual e suas '
        'implementações computacionais, buscando identificar a abordagem mais '
        'adequada para o sistema proposto. Quanto aos procedimentos, combina '
        'pesquisa bibliográfica — para embasar as decisões técnicas — com '
        'desenvolvimento experimental, materializado na construção do sistema '
        '(SILVA; MENEZES, 2005).'
    )

    # --- 3.2 ---
    add_secao_secundaria(doc, '3.2 Etapas do desenvolvimento')

    add_paragrafo(doc,
        'O desenvolvimento do sistema segue uma abordagem incremental, '
        'alinhada ao modelo de processo evolutivo descrito por Pressman e '
        'Maxim (2021), organizada nas seguintes etapas:'
    )

    add_paragrafo(doc,
        '1) Levantamento bibliográfico: revisão da literatura sobre técnicas '
        'de processamento de linguagem natural, métricas de similaridade textual '
        'e arquiteturas de sistemas web, com o objetivo de fundamentar as '
        'decisões de projeto. As buscas são realizadas no Google Acadêmico, '
        'IEEE Xplore e SciELO, utilizando as palavras-chave: "similaridade '
        'textual", "text similarity", "TF-IDF", "cosine similarity", '
        '"comparação de documentos" e "NLP".'
    )

    add_paragrafo(doc,
        '2) Definição dos requisitos: especificação das funcionalidades do '
        'sistema, incluindo: upload de dois arquivos de texto, '
        'pré-processamento automático (tokenização, remoção de stopwords, '
        'stemização), cálculo de múltiplas métricas de similaridade e '
        'apresentação dos resultados em interface gráfica.'
    )

    add_paragrafo(doc,
        '3) Escolha da plataforma tecnológica: seleção da linguagem de '
        'programação, frameworks de back-end e front-end, e bibliotecas de '
        'PLN, considerando critérios como maturidade do ecossistema, '
        'disponibilidade de bibliotecas especializadas e facilidade de '
        'implantação.'
    )

    add_paragrafo(doc,
        '4) Projeto da arquitetura: definição da arquitetura do sistema, '
        'incluindo a separação entre camadas de apresentação (front-end), '
        'lógica de negócios (back-end) e processamento de PLN, seguindo a '
        'arquitetura cliente-servidor.'
    )

    add_paragrafo(doc,
        '5) Implementação do módulo de processamento: desenvolvimento do '
        'núcleo do sistema responsável pelo pré-processamento dos textos, '
        'geração de representações vetoriais (TF-IDF) e cálculo das métricas '
        'de similaridade (cosseno e Jaccard).'
    )

    add_paragrafo(doc,
        '6) Desenvolvimento do front-end: construção da interface web que '
        'permite ao usuário carregar documentos, visualizar os índices de '
        'correlação e explorar os trechos de maior e menor similaridade entre '
        'os textos comparados.'
    )

    add_paragrafo(doc,
        '7) Testes e validação: execução de testes com pares de documentos '
        'de similaridade conhecida para validar a precisão das métricas '
        'implementadas e a usabilidade da interface.'
    )

    # --- 3.3 ---
    add_secao_secundaria(doc, '3.3 Ferramentas e tecnologias')

    add_paragrafo(doc,
        'A linguagem Python é adotada como tecnologia principal do back-end, '
        'em razão de seu vasto ecossistema de bibliotecas para processamento '
        'de linguagem natural. As principais bibliotecas utilizadas são: '
        'NLTK (Natural Language Toolkit), para tokenização, remoção de '
        'stopwords e stemização; scikit-learn, para vetorização TF-IDF e '
        'cálculo de similaridade por cosseno; e FastAPI, como framework '
        'assíncrono para exposição da API REST, com geração automática '
        'de documentação interativa (BIRD; KLEIN; LOPER, 2009).'
    )

    add_paragrafo(doc,
        'Para o front-end, utiliza-se o framework React com TypeScript, '
        'que proporciona tipagem estática e maior segurança no '
        'desenvolvimento de componentes reutilizáveis. A estilização é '
        'realizada com Tailwind CSS, um framework CSS utilitário que '
        'permite construir interfaces responsivas de forma ágil. A '
        'ferramenta de build Vite é empregada para o empacotamento e '
        'servimento da aplicação em ambiente de desenvolvimento. A '
        'comunicação entre front-end e back-end ocorre por meio de '
        'requisições HTTP à API REST, seguindo os princípios da '
        'arquitetura cliente-servidor (SOMMERVILLE, 2019).'
    )

    add_paragrafo(doc,
        'O controle de versão do código-fonte é realizado com Git, e o '
        'ambiente de desenvolvimento utiliza o editor Visual Studio Code. '
        'Para a documentação e gestão do projeto, são empregadas práticas '
        'da engenharia de software alinhadas ao modelo incremental, '
        'permitindo entregas parciais e refinamentos progressivos '
        '(PRESSMAN; MAXIM, 2021).'
    )


def secao_resultados(doc: Document) -> None:
    """4 RESULTADOS E DISCUSSÕES."""
    add_secao_primaria(doc, '4 Resultados e Discussões')

    add_paragrafo(doc,
        'Este capítulo apresenta os resultados obtidos em cada etapa descrita '
        'na metodologia, seguindo a mesma ordem cronológica. São discutidos os '
        'resultados do levantamento bibliográfico, as decisões arquiteturais, os '
        'resultados da implementação do sistema e os testes de validação '
        'realizados.'
    )

    # --- 4.1 ---
    add_secao_secundaria(doc, '4.1 Resultados do levantamento bibliográfico')

    add_paragrafo(doc,
        'O levantamento bibliográfico realizado nas bases Google Acadêmico, '
        'IEEE Xplore e SciELO resultou na identificação de três abordagens '
        'principais para a comparação de conteúdo textual: abordagens baseadas '
        'em frequência de termos (TF-IDF), abordagens baseadas em similaridade '
        'de conjuntos (Jaccard) e abordagens baseadas em embeddings semânticos '
        '(BERT/SBERT). Conforme a taxonomia proposta por Gomaa e Fahmy (2013), '
        'essas técnicas pertencem, respectivamente, às categorias corpus-based, '
        'string-based e knowledge-based. A análise da literatura permitiu '
        'identificar as vantagens e limitações de cada técnica, conforme '
        'sintetizado no Quadro 1.'
    )

    add_tabela(doc,
        'Quadro 1 – Comparativo das técnicas de similaridade textual',
        ['Técnica', 'Tipo', 'Vantagens', 'Limitações'],
        [
            ['TF-IDF + Cosseno', 'Estatística',
             'Rápida, interpretável, eficaz para textos longos',
             'Não captura semântica'],
            ['Jaccard', 'Conjuntos',
             'Simples, intuitiva',
             'Ignora frequência e semântica'],
            ['SBERT', 'Embeddings',
             'Captura similaridade semântica',
             'Requer modelo pré-treinado, maior custo computacional'],
        ],
        'Elaborado pelo autor com base em Gomaa e Fahmy (2013) e Reimers e Gurevych (2019)'
    )

    add_paragrafo(doc,
        'A análise comparativa evidenciou que a combinação de TF-IDF com '
        'similaridade por cosseno oferece o melhor equilíbrio entre '
        'desempenho computacional e qualidade dos resultados para a '
        'comparação de documentos completos, corroborando as conclusões de '
        'Lan (2022) sobre a eficácia do TF-IDF em tarefas de recuperação '
        'de informação. A técnica de Jaccard, embora limitada por não '
        'considerar a frequência dos termos (SUNG et al., 2019), '
        'complementa a análise ao fornecer uma perspectiva baseada na '
        'sobreposição vocabular. Já as abordagens baseadas em embeddings, '
        'como o SBERT, apresentam resultados superiores na identificação '
        'de paráfrases (REIMERS; GUREVYCH, 2019), porém demandam recursos '
        'computacionais significativamente maiores.'
    )

    # --- 4.2 ---
    add_secao_secundaria(doc, '4.2 Definição da arquitetura e tecnologias')

    add_paragrafo(doc,
        'Com base nos requisitos levantados e na análise bibliográfica, '
        'define-se a arquitetura do sistema segundo o padrão cliente-servidor, '
        'modelo predominante no desenvolvimento web conforme Sommerville '
        '(2019), com separação clara entre as camadas de apresentação e '
        'processamento. A linguagem Python é selecionada para o back-end '
        'em razão de seu ecossistema consolidado de bibliotecas de PLN '
        '(BIRD; KLEIN; LOPER, 2009) e da facilidade de prototipação. O '
        'framework FastAPI é adotado para a construção da API REST, por '
        'seu suporte nativo a operações assíncronas e geração automática '
        'de documentação interativa via Swagger UI, acessível pelo '
        'endpoint /docs. A API expõe dois endpoints: POST /api/compare, '
        'que recebe dois textos e retorna as métricas de similaridade, e '
        'GET /api/health, que permite verificar a disponibilidade do '
        'serviço. A comunicação entre front-end e back-end é habilitada '
        'por meio de CORS (Cross-Origin Resource Sharing), e os textos '
        'de entrada são validados com limite máximo de 500.000 caracteres '
        'para evitar sobrecarga do servidor.'
    )

    add_paragrafo(doc,
        'O Quadro 2 apresenta o mapeamento entre os componentes do sistema e '
        'as tecnologias selecionadas para cada função.'
    )

    add_tabela(doc,
        'Quadro 2 – Tecnologias adotadas por componente do sistema',
        ['Componente', 'Tecnologia', 'Justificativa'],
        [
            ['Back-end / API', 'Python + FastAPI',
             'Ecossistema de PLN maduro; API assíncrona com docs automáticos'],
            ['Pré-processamento', 'NLTK',
             'Tokenização, stopwords e stemização para português'],
            ['Vetorização', 'scikit-learn (TfidfVectorizer)',
             'Implementação otimizada de TF-IDF'],
            ['Similaridade', 'scikit-learn (cosine_similarity)',
             'Cálculo eficiente de similaridade por cosseno'],
            ['Front-end', 'React + TypeScript + Tailwind CSS',
             'Componentes tipados, estilização ágil e responsiva'],
            ['Build', 'Vite',
             'Empacotamento rápido com Hot Module Replacement'],
            ['Controle de versão', 'Git',
             'Rastreabilidade e versionamento do código'],
        ],
        'Elaborado pelo autor (2026)'
    )

    # --- 4.3 ---
    add_secao_secundaria(doc, '4.3 Implementação do módulo de processamento')

    add_paragrafo(doc,
        'O módulo de processamento constitui o núcleo do sistema e opera em '
        'três etapas sequenciais: pré-processamento, vetorização e cálculo '
        'de similaridade.'
    )

    add_paragrafo(doc,
        'Na etapa de pré-processamento, cada documento de entrada passa por '
        'tokenização (segmentação em palavras), conversão para minúsculas, '
        'remoção de stopwords da língua portuguesa utilizando a lista '
        'disponível no NLTK, e stemização por meio do algoritmo RSLP '
        '(Removedor de Sufixos da Língua Portuguesa), proposto por Orengo '
        'e Huyck (2001) especificamente para o tratamento das '
        'particularidades morfológicas do português. Essas operações '
        'reduzem a dimensionalidade do vocabulário e eliminam variações '
        'morfológicas que não contribuem para a análise de similaridade.'
    )

    add_paragrafo(doc,
        'Na etapa de vetorização, os textos pré-processados são convertidos '
        'em vetores numéricos utilizando o TfidfVectorizer da biblioteca '
        'scikit-learn, que implementa a técnica TF-IDF conforme descrita '
        'por Ramos (2003). O vetorizador é ajustado ao corpus formado '
        'pelos dois documentos, gerando uma matriz TF-IDF na qual cada '
        'linha representa um documento e cada coluna representa um termo '
        'ponderado pela sua relevância relativa.'
    )

    add_paragrafo(doc,
        'Na etapa de cálculo de similaridade, são aplicadas duas métricas '
        'sobre os vetores gerados: a similaridade por cosseno, amplamente '
        'utilizada em sistemas de comparação textual (MANNING; RAGHAVAN; '
        'SCHÜTZE, 2008), que mede a proximidade direcional entre os vetores '
        'e resulta em um valor entre 0 e 1; e o coeficiente de Jaccard, '
        'calculado sobre os conjuntos de termos únicos de cada documento '
        '(GOMAA; FAHMY, 2013). O índice de correlação final é apresentado '
        'como um percentual, facilitando a interpretação pelo usuário.'
    )

    add_paragrafo(doc,
        'Para melhorar a interpretabilidade dos resultados, o módulo '
        'implementa um mapeamento reverso entre stems e palavras originais. '
        'Durante a stemização, cada radical gerado é associado à forma '
        'original mais completa encontrada nos textos de entrada, permitindo '
        'que a interface exiba termos legíveis (por exemplo, "processamento") '
        'em vez de radicais truncados (por exemplo, "process"). O sistema '
        'também coleta estatísticas de processamento, incluindo a contagem '
        'de palavras de cada texto e o tempo de execução em milissegundos, '
        'informações apresentadas ao usuário junto aos resultados.'
    )

    # --- 4.4 ---
    add_secao_secundaria(doc, '4.4 Desenvolvimento do front-end')

    add_paragrafo(doc,
        'A interface web do sistema é construída com React e TypeScript, '
        'organizando a apresentação em três componentes reutilizáveis '
        'estilizados com Tailwind CSS. O componente TextInput disponibiliza '
        'duas áreas de texto lado a lado, cada uma com contador de palavras '
        'em tempo real e botão para carregamento de arquivos nos formatos '
        '.txt, .md, .csv e .log. Após o upload, o nome do arquivo é exibido '
        'como confirmação visual. A comparação pode ser acionada pelo botão '
        '"Comparar" ou pelo atalho de teclado Ctrl+Enter, e durante o '
        'processamento um indicador de carregamento é apresentado.'
    )

    add_paragrafo(doc,
        'Os resultados são apresentados em dois componentes React: '
        '(a) o componente Results, que exibe o índice de similaridade '
        'por cosseno e o coeficiente de Jaccard em formato percentual com '
        'barras de progresso coloridas por faixas de intensidade — verde '
        'para similaridade alta (acima de 70%), amarelo para moderada '
        '(40% a 70%) e vermelho para baixa (abaixo de 40%) — além de '
        'estatísticas de processamento como contagem de palavras de cada '
        'texto e tempo de execução em milissegundos; e '
        '(b) o componente SharedTerms, que apresenta uma tabela dos quinze '
        'termos mais relevantes compartilhados entre os documentos, '
        'ordenados pelo peso TF-IDF médio, com barras visuais de '
        'relevância proporcionais ao peso. Abaixo da tabela, listas de '
        'termos exclusivos de cada texto são exibidas como etiquetas '
        'coloridas, diferenciando visualmente os vocabulários exclusivos '
        'de cada documento. Essa apresentação permite ao usuário '
        'compreender não apenas o grau de similaridade global, mas também '
        'quais elementos textuais contribuem para a correlação identificada.'
    )

    # --- 4.5 ---
    add_secao_secundaria(doc, '4.5 Testes e validação')

    add_paragrafo(doc,
        'Para validar a precisão do sistema, são realizados testes com três '
        'cenários distintos de pares de documentos: (a) textos idênticos, '
        'para os quais o índice de similaridade esperado é 1,0; (b) textos '
        'completamente distintos em vocabulário e tema, com similaridade '
        'esperada próxima a 0; e (c) textos parcialmente similares, como '
        'versões revisadas de um mesmo documento ou textos sobre o mesmo '
        'tema com redações distintas.'
    )

    add_paragrafo(doc,
        'O Quadro 3 sintetiza os resultados obtidos nos três cenários de '
        'teste, apresentando os valores de similaridade por cosseno e '
        'coeficiente de Jaccard, a quantidade de termos compartilhados '
        'identificados e o tempo de processamento.'
    )

    add_tabela(doc,
        'Quadro 3 – Resultados dos testes de validação do sistema',
        ['Cenário', 'Cosseno', 'Jaccard', 'Termos em comum', 'Tempo (ms)'],
        [
            ['Textos idênticos', '100%', '100%', 'Todos', '< 20'],
            ['Textos sem relação', '0%', '0%', 'Nenhum', '< 20'],
            ['Textos parcialmente similares', '38%', '35%', '7 termos', '< 25'],
        ],
        'Elaborado pelo autor (2026)'
    )

    add_paragrafo(doc,
        'Os resultados demonstram que a combinação TF-IDF com similaridade '
        'por cosseno produz índices coerentes com a similaridade real dos '
        'documentos, alinhando-se às conclusões de Lan (2022) sobre a '
        'eficácia dessa combinação em tarefas de comparação textual. '
        'Textos idênticos resultam em similaridade 1,0 (100%), '
        'textos sem sobreposição vocabular resultam em similaridade 0,0 '
        '(0%), e textos parcialmente similares — redigidos sobre o mesmo '
        'tema com vocabulários distintos — apresentam valores intermediários '
        '(38% de cosseno e 35% de Jaccard) que refletem adequadamente o '
        'grau de correlação entre os conteúdos. O coeficiente de Jaccard '
        'apresenta valores consistentes, embora sistematicamente inferiores '
        'aos do cosseno, comportamento esperado conforme Sung et al. (2019), '
        'que destacam a limitação do Jaccard em não considerar a frequência '
        'dos termos.'
    )

    add_paragrafo(doc,
        'A interface web mostra-se funcional e responsiva, com tempo de '
        'processamento inferior a 25 milissegundos para textos de tamanho '
        'moderado. O mapeamento reverso de stems para palavras originais '
        'contribui significativamente para a interpretabilidade, permitindo '
        'que o usuário identifique os termos reais que sustentam o índice '
        'de correlação, em vez de radicais truncados — uma necessidade '
        'destacada por Bird, Klein e Loper (2009) como fundamental para a '
        'usabilidade de sistemas de PLN. As estatísticas exibidas — '
        'contagem de palavras e tempo de processamento — conferem '
        'transparência ao funcionamento do sistema, princípio essencial '
        'da engenharia de software segundo Pressman e Maxim (2021).'
    )


def secao_consideracoes(doc: Document) -> None:
    """5 CONSIDERAÇÕES FINAIS."""
    add_secao_primaria(doc, '5 Considerações Finais')

    add_paragrafo(doc,
        'Este trabalho teve como objetivo o desenvolvimento de um sistema web '
        'de comparação de conteúdo de textos, capaz de calcular e apresentar '
        'o índice de correlação entre dois documentos por meio de técnicas de '
        'processamento de linguagem natural e métricas de similaridade textual. '
        'A pergunta de pesquisa que orientou o estudo — como desenvolver um '
        'sistema web capaz de comparar o conteúdo de dois textos e apresentar, '
        'de forma clara e interpretável, o índice de correlação entre eles — '
        'foi respondida com a construção de um sistema funcional que integra '
        'pipeline de PLN, métricas quantitativas e interface web interativa.'
    )

    add_paragrafo(doc,
        'O objetivo específico (i), referente à investigação das técnicas de '
        'similaridade textual, foi alcançado por meio do levantamento '
        'bibliográfico que identificou e comparou três abordagens principais: '
        'TF-IDF com similaridade por cosseno, coeficiente de Jaccard e '
        'embeddings semânticos (SBERT). A análise permitiu concluir que a '
        'combinação TF-IDF com cosseno oferece o melhor equilíbrio entre '
        'desempenho e qualidade para a comparação de documentos completos.'
    )

    add_paragrafo(doc,
        'O objetivo específico (ii), relativo ao projeto da arquitetura, foi '
        'atendido com a definição de uma arquitetura cliente-servidor baseada '
        'em Python e FastAPI para o back-end e React com TypeScript e '
        'Tailwind CSS para o front-end. A escolha do ecossistema Python '
        'mostrou-se adequada pela disponibilidade de bibliotecas como NLTK '
        'e scikit-learn, que viabilizaram a implementação eficiente do '
        'módulo de processamento.'
    )

    add_paragrafo(doc,
        'O objetivo específico (iii), concernente à implementação do módulo '
        'de processamento, foi concretizado com a construção do pipeline de '
        'pré-processamento (tokenização, remoção de stopwords, stemização '
        'RSLP), vetorização TF-IDF e cálculo de métricas de similaridade. '
        'Os testes realizados confirmaram a coerência dos índices calculados '
        'com a similaridade real dos documentos avaliados.'
    )

    add_paragrafo(doc,
        'O objetivo específico (iv), relacionado ao desenvolvimento do '
        'front-end, foi cumprido com a construção de uma interface web '
        'em React com TypeScript e Tailwind CSS, funcional e intuitiva, '
        'que permite ao usuário carregar documentos, visualizar os índices '
        'de correlação em formato percentual com barras visuais e '
        'identificar os termos compartilhados que sustentam os resultados.'
    )

    add_paragrafo(doc,
        'Cabe destacar que o sistema apresenta limitações inerentes à '
        'abordagem adotada: a análise baseada em TF-IDF restringe-se à '
        'similaridade léxica, não capturando relações semânticas entre '
        'termos sinônimos ou paráfrases; e a validação foi realizada com '
        'um conjunto limitado de cenários de teste. Ainda assim, o trabalho '
        'contribui para a área de engenharia de software aplicada ao PLN ao '
        'documentar de forma transparente o processo completo de concepção, '
        'projeto e implementação de um sistema funcional, incluindo as '
        'decisões técnicas e seus fundamentos.'
    )

    add_paragrafo(doc,
        'Como sugestões para trabalhos futuros, recomenda-se: a integração '
        'de métricas de similaridade semântica baseadas em embeddings, como '
        'o SBERT, para complementar a análise estatística do TF-IDF; a '
        'implementação de funcionalidades de comparação por trechos, '
        'permitindo identificar passagens específicas com maior correlação; '
        'e a realização de testes com volumes maiores de documentos para '
        'avaliar a escalabilidade do sistema em cenários de uso corporativo '
        'ou acadêmico.'
    )


# =============================================================================
# GERAÇÃO DO DOCUMENTO
# =============================================================================

def gerar_documento() -> None:
    doc = Document()
    configurar_pagina(doc)
    estilo_padrao(doc)

    add_titulo_artigo(doc)
    add_autores(doc)

    secao_resumo(doc)
    secao_introducao(doc)
    secao_fundamentacao(doc)
    secao_metodologia(doc)
    secao_resultados(doc)
    secao_consideracoes(doc)

    if REFERENCIAS:
        add_referencias_formatadas(doc, REFERENCIAS)

    caminho = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        'docs',
        'tcc-bruno-lorencatto.docx'
    )
    doc.save(caminho)
    logging.info('Arquivo salvo em: %s', caminho)


if __name__ == '__main__':
    gerar_documento()
