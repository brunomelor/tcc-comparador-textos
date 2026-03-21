# Backend — Comparador de Textos

API REST para comparação de similaridade textual.

## Endpoints

| Método | Rota | Descrição |
|---|---|---|
| GET | `/api/health` | Verificação de disponibilidade |
| POST | `/api/compare` | Compara dois textos |

### POST /api/compare

**Request:**
```json
{
  "text1": "Primeiro texto...",
  "text2": "Segundo texto..."
}
```

**Response:**
```json
{
  "cosine_similarity": 0.85,
  "jaccard_coefficient": 0.62,
  "shared_terms": [{"term": "exemplo", "weight": 0.42}],
  "text1_unique_terms": ["termo1"],
  "text2_unique_terms": ["termo2"],
  "stats": {"text1_words": 150, "text2_words": 120, "processing_time_ms": 45}
}
```

## Módulos

- `main.py` — API FastAPI, endpoints, modelos Pydantic
- `nlp.py` — Pipeline NLP: tokenização, stopwords, stemização RSLP, TF-IDF, métricas

## Execução

```bash
pip install -r requirements.txt
python main.py
```
