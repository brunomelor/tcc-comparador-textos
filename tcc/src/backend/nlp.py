# -*- coding: utf-8 -*-
"""
Módulo de Processamento de Linguagem Natural (PLN)
Pipeline: tokenização → lowercase → remoção de stopwords → stemização RSLP → TF-IDF → métricas
"""

from __future__ import annotations

import logging
import time
import re
from typing import TypedDict

import numpy as np
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.stem import RSLPStemmer
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

logger = logging.getLogger(__name__)


class SharedTermDict(TypedDict):
    term: str
    weight: float


class StatsDict(TypedDict):
    text1_words: int
    text2_words: int
    processing_time_ms: int


class CompareResultDict(TypedDict):
    cosine_similarity: float
    jaccard_coefficient: float
    shared_terms: list[SharedTermDict]
    text1_unique_terms: list[str]
    text2_unique_terms: list[str]
    stats: StatsDict


_NLTK_RESOURCES = {
    "punkt": "tokenizers/punkt",
    "punkt_tab": "tokenizers/punkt_tab",
    "stopwords": "corpora/stopwords",
    "rslp": "stemmers/rslp",
}


class TextComparer:
    """Encapsula o pipeline de comparação de textos."""

    MAX_TEXT_LENGTH = 500_000
    MAX_SHARED_TERMS = 15
    MAX_UNIQUE_TERMS = 10
    MIN_TOKEN_LENGTH = 1

    def __init__(self) -> None:
        self._initialize_nltk()
        self._stemmer = RSLPStemmer()
        self._stopwords_pt = frozenset(stopwords.words("portuguese"))

    @staticmethod
    def _initialize_nltk() -> None:
        """Baixa recursos NLTK apenas se necessário."""
        for name, path in _NLTK_RESOURCES.items():
            try:
                nltk.data.find(path)
            except LookupError:
                try:
                    nltk.download(name, quiet=True)
                except (OSError, RuntimeError) as e:
                    logger.warning(f"Falha ao baixar recurso NLTK '{name}' — verifique conexao de rede: {e}")

    def preprocess(self, text: str) -> tuple[str, dict[str, str]]:
        """
        Pré-processa um texto: lowercase, tokenização, remoção de stopwords, stemização.

        Returns:
            (texto_processado, stem_to_word) — texto com stems e mapa stem→palavra original
        """
        text = text.lower()
        text = re.sub(r"[^\w\s]", " ", text)
        text = re.sub(r"\d+", "", text)
        tokens = word_tokenize(text, language="portuguese")
        tokens = [t for t in tokens if t not in self._stopwords_pt and len(t) > self.MIN_TOKEN_LENGTH]

        stem_to_word: dict[str, str] = {}
        stems: list[str] = []
        for token in tokens:
            stem = self._stemmer.stem(token)
            stems.append(stem)
            if stem not in stem_to_word or len(token) > len(stem_to_word[stem]):
                stem_to_word[stem] = token

        return " ".join(stems), stem_to_word

    def compute_jaccard(self, text1: str, text2: str) -> float:
        """Calcula o coeficiente de Jaccard entre dois textos pré-processados."""
        set1 = set(text1.split())
        set2 = set(text2.split())
        if not set1 and not set2:
            return 0.0  # Conjuntos vazios => sem similaridade mensuravel
        intersection = set1 & set2
        union = set1 | set2
        return len(intersection) / len(union)

    @staticmethod
    def count_words(text: str) -> int:
        """Conta palavras de um texto bruto."""
        return len(text.split())

    def _build_empty_result(self, text1: str, text2: str, elapsed_ms: int = 0) -> CompareResultDict:
        """Constrói resultado vazio padrão."""
        return {
            "cosine_similarity": 0.0,
            "jaccard_coefficient": 0.0,
            "shared_terms": [],
            "text1_unique_terms": [],
            "text2_unique_terms": [],
            "stats": {
                "text1_words": self.count_words(text1),
                "text2_words": self.count_words(text2),
                "processing_time_ms": elapsed_ms,
            },
        }

    def _extract_terms(
        self,
        feature_names: np.ndarray,
        vec1: np.ndarray,
        vec2: np.ndarray,
        combined_map: dict[str, str],
        stem_map1: dict[str, str],
        stem_map2: dict[str, str],
    ) -> tuple[list[SharedTermDict], list[str], list[str]]:
        """Extrai termos compartilhados e únicos dos vetores TF-IDF."""
        shared: list[SharedTermDict] = []
        set1_terms: set[str] = set()
        set2_terms: set[str] = set()

        for i, stem in enumerate(feature_names):
            w1, w2 = vec1[i], vec2[i]
            original = combined_map.get(stem, stem)
            if w1 > 0 and w2 > 0:
                shared.append({"term": original, "weight": round(float((w1 + w2) / 2), 4)})
            elif w1 > 0:
                set1_terms.add(stem_map1.get(stem, stem))
            elif w2 > 0:
                set2_terms.add(stem_map2.get(stem, stem))

        shared.sort(key=lambda x: x["weight"], reverse=True)
        return (
            shared[:self.MAX_SHARED_TERMS],
            sorted(set1_terms)[:self.MAX_UNIQUE_TERMS],
            sorted(set2_terms)[:self.MAX_UNIQUE_TERMS],
        )

    def compare_texts(self, text1: str, text2: str) -> CompareResultDict:
        """Compara dois textos e retorna métricas de similaridade."""
        start = time.perf_counter()

        if not text1.strip() or not text2.strip():
            return self._build_empty_result(text1, text2)

        text1 = text1[:self.MAX_TEXT_LENGTH]
        text2 = text2[:self.MAX_TEXT_LENGTH]

        processed1, stem_map1 = self.preprocess(text1)
        processed2, stem_map2 = self.preprocess(text2)

        if not processed1.strip() or not processed2.strip():
            elapsed_ms = round((time.perf_counter() - start) * 1000)
            return self._build_empty_result(text1, text2, elapsed_ms)

        # Mapa combinado stem→palavra original (prioriza forma mais completa)
        combined_map = {**stem_map1, **stem_map2}
        for stem, word in stem_map1.items():
            if stem in stem_map2 and len(stem_map2[stem]) > len(word):
                combined_map[stem] = stem_map2[stem]

        # Vetorização TF-IDF (max_features limita vocabulario para prevenir uso excessivo de memoria)
        vectorizer = TfidfVectorizer(max_features=10_000)
        tfidf_matrix = vectorizer.fit_transform([processed1, processed2])
        feature_names = vectorizer.get_feature_names_out()

        cosine_sim = float(cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:2])[0][0])
        jaccard = self.compute_jaccard(processed1, processed2)

        vec1 = tfidf_matrix[0].toarray()[0]
        vec2 = tfidf_matrix[1].toarray()[0]

        shared, unique1, unique2 = self._extract_terms(
            feature_names, vec1, vec2, combined_map, stem_map1, stem_map2
        )

        elapsed_ms = round((time.perf_counter() - start) * 1000)

        return {
            "cosine_similarity": round(cosine_sim, 4),
            "jaccard_coefficient": round(jaccard, 4),
            "shared_terms": shared,
            "text1_unique_terms": unique1,
            "text2_unique_terms": unique2,
            "stats": {
                "text1_words": self.count_words(text1),
                "text2_words": self.count_words(text2),
                "processing_time_ms": elapsed_ms,
            },
        }


# Instância padrão para manter compatibilidade com main.py
_default_comparer = TextComparer()
compare_texts = _default_comparer.compare_texts

# Constantes expostas no nível do módulo para compatibilidade
MAX_TEXT_LENGTH = TextComparer.MAX_TEXT_LENGTH
MAX_SHARED_TERMS = TextComparer.MAX_SHARED_TERMS
MAX_UNIQUE_TERMS = TextComparer.MAX_UNIQUE_TERMS
MIN_TOKEN_LENGTH = TextComparer.MIN_TOKEN_LENGTH
