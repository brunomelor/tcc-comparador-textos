import { useState, useEffect, useCallback } from "react";
import TextInput from "./components/TextInput";
import Results from "./components/Results";
import SharedTerms from "./components/SharedTerms";
import Spinner from "./components/Spinner";
import { compareTexts } from "./services/api";
import type { CompareResult } from "./types/api";

/** Componente principal — orquestra comparação de textos e exibe resultados. */
function App(): React.JSX.Element {
  const [text1, setText1] = useState<string>("");
  const [text2, setText2] = useState<string>("");
  const [result, setResult] = useState<CompareResult | null>(null);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>("");

  const handleCompare = useCallback(async (): Promise<void> => {
    if (!text1.trim() || !text2.trim()) {
      setError("Insira os dois textos para comparar.");
      return;
    }

    setError("");
    setLoading(true);

    try {
      const data = await compareTexts({ text1, text2 });
      setResult(data);
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : "Erro desconhecido");
    } finally {
      setLoading(false);
    }
  }, [text1, text2]);

  const handleClear = useCallback((): void => {
    setText1("");
    setText2("");
    setResult(null);
    setError("");
  }, []);

  useEffect(() => {
    const handler = (e: KeyboardEvent): void => {
      if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
        e.preventDefault();
        void handleCompare();
      }
    };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, [handleCompare]);

  return (
    <div className="min-h-screen bg-gray-50">
      <header className="bg-white border-b border-gray-200 shadow-sm">
        <div className="max-w-6xl mx-auto px-4 py-4">
          <h1 className="text-xl font-bold text-gray-800">
            Comparador de Conteúdo de Textos
          </h1>
          <p className="text-sm text-gray-500">
            Análise de similaridade textual com TF-IDF, cosseno e Jaccard
          </p>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8 space-y-6">
        <div className="flex flex-col md:flex-row gap-4">
          <TextInput
            label="Texto 1"
            value={text1}
            onChange={setText1}
            placeholder="Cole ou carregue o primeiro texto..."
          />
          <TextInput
            label="Texto 2"
            value={text2}
            onChange={setText2}
            placeholder="Cole ou carregue o segundo texto..."
          />
        </div>

        <div className="flex gap-3 justify-center items-center">
          <button
            onClick={() => void handleCompare()}
            disabled={loading}
            className="px-8 py-2.5 bg-blue-600 text-white font-semibold rounded-lg hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors cursor-pointer focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2"
            aria-busy={loading}
            aria-label={
              loading ? "Processando comparação" : "Comparar textos"
            }
          >
            {loading ? <Spinner label="Processando..." /> : "Comparar"}
          </button>
          <button
            onClick={handleClear}
            className="px-6 py-2.5 bg-gray-200 text-gray-700 font-semibold rounded-lg hover:bg-gray-300 transition-colors cursor-pointer"
            aria-label="Limpar textos e resultados"
          >
            Limpar
          </button>
          <span className="text-xs text-gray-400 hidden sm:inline">
            Ctrl+Enter
          </span>
        </div>

        {error && (
          <div
            className="p-4 bg-red-50 border border-red-200 rounded-lg text-red-700 text-sm"
            role="alert"
            aria-live="assertive"
          >
            {error}
          </div>
        )}

        {result && (
          <div className="space-y-6">
            <Results
              cosineSimilarity={result.cosine_similarity}
              jaccardCoefficient={result.jaccard_coefficient}
              stats={result.stats}
            />
            <SharedTerms
              sharedTerms={result.shared_terms}
              text1UniqueTerms={result.text1_unique_terms}
              text2UniqueTerms={result.text2_unique_terms}
            />
          </div>
        )}
      </main>

      <footer className="text-center py-4 text-xs text-gray-400">
        TCC — Engenharia de Software — UNINTER — Bruno Lorencatto
      </footer>
    </div>
  );
}

export default App;
