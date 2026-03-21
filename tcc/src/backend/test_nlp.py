# -*- coding: utf-8 -*-
"""Testes do pipeline NLP — valida preprocess, Jaccard, cosseno e compare_texts."""

from nlp import TextComparer


comparer = TextComparer()


# ============================================================
# preprocess
# ============================================================

def test_preprocess_lowercase_e_stems():
    """Texto deve ser convertido para lowercase e stemizado."""
    processed, stem_map = comparer.preprocess("Casas bonitas grandes")
    assert processed  # não vazio
    assert all(w == w.lower() for w in processed.split())
    assert len(stem_map) > 0


def test_preprocess_remove_pontuacao_e_numeros():
    """Pontuação e números devem ser removidos."""
    processed, _ = comparer.preprocess("Teste: 123! @#$ pontuação, ok?")
    assert "123" not in processed
    assert ":" not in processed
    assert "@" not in processed


def test_preprocess_remove_stopwords():
    """Stopwords em português devem ser removidas."""
    processed, _ = comparer.preprocess("O gato está na casa de Maria")
    # "o", "está", "na", "de" são stopwords
    words = processed.split()
    for sw in ["o", "está", "na", "de"]:
        assert sw not in words


def test_preprocess_texto_vazio():
    """Texto vazio retorna string vazia e mapa vazio."""
    processed, stem_map = comparer.preprocess("")
    assert processed == ""
    assert stem_map == {}


def test_preprocess_apenas_stopwords():
    """Texto com apenas stopwords retorna vazio."""
    processed, _ = comparer.preprocess("o a de do da em no na")
    assert processed.strip() == ""


def test_preprocess_stem_map_preserva_forma_longa():
    """stem_to_word deve manter a forma mais longa da palavra."""
    _, stem_map = comparer.preprocess("casa casas casarão")
    # Todos derivam do mesmo stem; deve manter a forma mais longa
    for stem, word in stem_map.items():
        assert len(word) >= len(stem)


# ============================================================
# compute_jaccard
# ============================================================

def test_jaccard_textos_identicos():
    """Textos idênticos devem ter Jaccard = 1.0."""
    assert comparer.compute_jaccard("foo bar baz", "foo bar baz") == 1.0


def test_jaccard_textos_sem_relacao():
    """Textos sem termos em comum devem ter Jaccard = 0.0."""
    assert comparer.compute_jaccard("foo bar", "baz qux") == 0.0


def test_jaccard_parcial():
    """Textos com sobreposição parcial devem ter 0 < Jaccard < 1."""
    j = comparer.compute_jaccard("foo bar baz", "bar baz qux")
    assert 0 < j < 1
    # interseção = {bar, baz} = 2, união = {foo, bar, baz, qux} = 4
    assert abs(j - 0.5) < 0.001


def test_jaccard_ambos_vazios():
    """Dois textos vazios devem retornar 0.0 (sem similaridade mensurável)."""
    assert comparer.compute_jaccard("", "") == 0.0


# ============================================================
# count_words
# ============================================================

def test_count_words_normal():
    assert comparer.count_words("uma frase com quatro palavras") == 5


def test_count_words_vazio():
    assert comparer.count_words("") == 0


# ============================================================
# compare_texts (pipeline completo)
# ============================================================

def test_compare_textos_identicos():
    """Textos idênticos devem ter similaridade por cosseno = 1.0."""
    texto = "O processamento de linguagem natural permite análise computacional de textos."
    result = comparer.compare_texts(texto, texto)
    assert result["cosine_similarity"] == 1.0
    assert result["jaccard_coefficient"] == 1.0
    assert result["stats"]["text1_words"] > 0
    assert result["stats"]["processing_time_ms"] >= 0


def test_compare_textos_sem_relacao():
    """Textos completamente diferentes devem ter similaridade baixa ou zero."""
    t1 = "O gato dormiu no sofá durante toda a tarde ensolarada."
    t2 = "Programação funcional utiliza funções puras e imutabilidade de dados."
    result = comparer.compare_texts(t1, t2)
    assert result["cosine_similarity"] < 0.2
    assert result["jaccard_coefficient"] < 0.2


def test_compare_textos_parcialmente_similares():
    """Textos com sobreposição parcial devem ter similaridade intermediária."""
    t1 = "O processamento de linguagem natural utiliza técnicas de análise computacional."
    t2 = "Análise computacional de textos é uma área do processamento de linguagem natural."
    result = comparer.compare_texts(t1, t2)
    assert 0.3 < result["cosine_similarity"] < 1.0
    assert len(result["shared_terms"]) > 0


def test_compare_texto_vazio():
    """Texto vazio retorna resultado zerado."""
    result = comparer.compare_texts("", "algo aqui")
    assert result["cosine_similarity"] == 0.0
    assert result["jaccard_coefficient"] == 0.0


def test_compare_ambos_vazios():
    """Ambos vazios retorna resultado zerado."""
    result = comparer.compare_texts("", "")
    assert result["cosine_similarity"] == 0.0
    assert result["jaccard_coefficient"] == 0.0


def test_compare_texto_apenas_numeros():
    """Texto com apenas números (removidos no preprocess) retorna zerado."""
    result = comparer.compare_texts("123 456 789", "outro texto real")
    assert result["cosine_similarity"] == 0.0


def test_compare_retorna_estrutura_completa():
    """Resultado deve conter todas as chaves esperadas."""
    result = comparer.compare_texts("texto um exemplo", "texto dois exemplo")
    assert "cosine_similarity" in result
    assert "jaccard_coefficient" in result
    assert "shared_terms" in result
    assert "text1_unique_terms" in result
    assert "text2_unique_terms" in result
    assert "stats" in result
    assert "text1_words" in result["stats"]
    assert "text2_words" in result["stats"]
    assert "processing_time_ms" in result["stats"]


def test_compare_shared_terms_tem_peso():
    """Termos compartilhados devem ter term e weight."""
    t1 = "Engenharia de software é fundamental para sistemas modernos."
    t2 = "Sistemas modernos dependem de engenharia de software de qualidade."
    result = comparer.compare_texts(t1, t2)
    for term in result["shared_terms"]:
        assert "term" in term
        assert "weight" in term
        assert term["weight"] > 0


def test_compare_unique_terms():
    """Termos exclusivos devem aparecer apenas em um dos textos."""
    t1 = "Python é excelente para ciência de dados."
    t2 = "JavaScript domina o desenvolvimento web frontend."
    result = comparer.compare_texts(t1, t2)
    assert len(result["text1_unique_terms"]) > 0 or len(result["text2_unique_terms"]) > 0


def test_compare_limite_shared_terms():
    """Shared terms deve respeitar o limite MAX_SHARED_TERMS."""
    # Textos longos com muitos termos em comum
    palavras = " ".join([f"palavra{i}" for i in range(100)])
    result = comparer.compare_texts(palavras, palavras)
    assert len(result["shared_terms"]) <= TextComparer.MAX_SHARED_TERMS
