# TCC — Sistema de Comparação de Conteúdo de Textos

Trabalho de Conclusão de Curso — Engenharia de Software — UNINTER
Bruno Meinertz Lorençatto — RU: 3802493

---

## Pré-requisitos

- **Python 3.10+** — [python.org/downloads](https://www.python.org/downloads/)
- **Node.js 18+** — [nodejs.org](https://nodejs.org/)

Para verificar se já estão instalados, abra o terminal e execute:

```bash
python --version
node --version
```

---

## Como executar

São dois terminais: um para o backend (API) e outro para o frontend (interface).

### 1. Backend (Terminal 1)

```bash
cd tcc/src/backend
pip install -r requirements.txt
python main.py
```

Aguarde a mensagem `Uvicorn running on http://127.0.0.1:8000`. O backend está pronto.

### 2. Frontend (Terminal 2)

```bash
cd tcc/src/frontend
npm install
npm run dev
```

Aguarde a mensagem `Local: http://localhost:5173/`.

### 3. Acessar o sistema

Abra o navegador em <http://localhost:5173>

---

## Como usar

1. Cole ou carregue dois textos nos campos "Texto 1" e "Texto 2"
2. Clique em **Comparar** (ou pressione Ctrl+Enter)
3. O sistema exibe:
   - **Similaridade por cosseno** — proximidade vetorial (TF-IDF)
   - **Coeficiente de Jaccard** — sobreposição de termos únicos
   - **Termos compartilhados** — palavras em comum com peso TF-IDF
   - **Termos exclusivos** — palavras que aparecem em apenas um dos textos

---

## Estrutura do projeto

```text
tcc/
├── src/
│   ├── backend/           # API REST em Python (FastAPI)
│   │   ├── main.py        # Servidor e endpoints
│   │   ├── nlp.py         # Pipeline de processamento NLP
│   │   └── test_nlp.py    # Testes automatizados
│   └── frontend/          # Interface web (React + TypeScript)
│       └── src/
│           ├── App.tsx    # Componente principal
│           └── ...
├── docs/
│   └── tcc-bruno-lorencatto.docx   # Artigo do TCC
└── gerar-tcc.py           # Script que gera o artigo (.docx)
```

## Autor

Bruno Meinertz Lorençatto — RU: 3802493 — Centro Universitário Internacional (UNINTER)
