import type { CompareResult, CompareRequest } from "../types/api";

const API_URL = import.meta.env.VITE_API_URL ?? "http://localhost:8000";

function isCompareResult(data: unknown): data is CompareResult {
  if (typeof data !== "object" || data === null) return false;
  const obj = data as Record<string, unknown>;
  return (
    typeof obj.cosine_similarity === "number" &&
    typeof obj.jaccard_coefficient === "number" &&
    Array.isArray(obj.shared_terms) &&
    Array.isArray(obj.text1_unique_terms) &&
    Array.isArray(obj.text2_unique_terms) &&
    typeof obj.stats === "object" &&
    obj.stats !== null
  );
}

export async function compareTexts(req: CompareRequest): Promise<CompareResult> {
  const res = await fetch(`${API_URL}/api/compare`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(req),
  });

  if (!res.ok) {
    const body: { detail?: string } | null = await res.json().catch(() => null);
    throw new Error(body?.detail ?? `Erro ${res.status}`);
  }

  const data: unknown = await res.json();
  if (!isCompareResult(data)) {
    throw new Error("Resposta da API em formato inválido");
  }
  return data;
}
