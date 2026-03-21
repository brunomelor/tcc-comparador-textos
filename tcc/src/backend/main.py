# -*- coding: utf-8 -*-
"""
API REST — Sistema de Comparação de Conteúdo de Textos
FastAPI + Uvicorn
"""

from __future__ import annotations

import logging
import os

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field

from nlp import compare_texts, MAX_TEXT_LENGTH

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Comparador de Textos",
    description="API para comparação de conteúdo textual usando TF-IDF e métricas de similaridade",
    version="1.0.0",
)

_default_origins = ["http://localhost:5173", "http://127.0.0.1:5173"]
_cors_origins = os.environ.get("CORS_ORIGINS", "").split(",") if os.environ.get("CORS_ORIGINS") else _default_origins

app.add_middleware(
    CORSMiddleware,
    allow_origins=_cors_origins,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["Content-Type"],
)


class CompareRequest(BaseModel):
    text1: str = Field(..., max_length=MAX_TEXT_LENGTH)
    text2: str = Field(..., max_length=MAX_TEXT_LENGTH)


class SharedTerm(BaseModel):
    term: str
    weight: float


class Stats(BaseModel):
    text1_words: int
    text2_words: int
    processing_time_ms: int


class CompareResponse(BaseModel):
    cosine_similarity: float
    jaccard_coefficient: float
    shared_terms: list[SharedTerm]
    text1_unique_terms: list[str]
    text2_unique_terms: list[str]
    stats: Stats


@app.get("/api/health")
async def health() -> dict[str, str]:
    """Verifica se a API está ativa."""
    return {"status": "ok"}


@app.post("/api/compare", response_model=CompareResponse)
async def compare(req: CompareRequest) -> CompareResponse:
    """Compara dois textos e retorna métricas de similaridade."""
    if not req.text1.strip() or not req.text2.strip():
        raise HTTPException(status_code=400, detail="Ambos os textos devem conter conteúdo.")
    try:
        result = compare_texts(req.text1, req.text2)
    except (ValueError, RuntimeError, TypeError):
        logger.exception("Erro ao processar comparação de textos")
        raise HTTPException(status_code=500, detail="Erro ao processar os textos.")
    return result


if __name__ == "__main__":
    import uvicorn

    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8000"))
    reload = os.environ.get("RELOAD", "false").lower() == "true"
    uvicorn.run("main:app", host=host, port=port, reload=reload)
