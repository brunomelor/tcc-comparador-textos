import type { Stats } from "../types/api";

interface ResultsProps {
  cosineSimilarity: number;
  jaccardCoefficient: number;
  stats: Stats;
}

interface MetricBarProps {
  label: string;
  value: number;
  description: string;
}

const SIMILARITY_THRESHOLDS: { min: number; color: string }[] = [
  { min: 0.7, color: "bg-green-500" },
  { min: 0.4, color: "bg-yellow-500" },
  { min: 0.0, color: "bg-red-500" },
];

function getBarColor(value: number): string {
  const threshold = SIMILARITY_THRESHOLDS.find((t) => value >= t.min);
  return threshold?.color ?? "bg-gray-500";
}

function MetricBar({ label, value, description }: MetricBarProps): React.JSX.Element {
  const pct = Math.round(value * 100);
  return (
    <div className="space-y-1">
      <div className="flex justify-between items-baseline">
        <span className="text-sm font-semibold text-gray-700">{label}</span>
        <span className="text-2xl font-bold text-gray-900">{pct}%</span>
      </div>
      <div
        className="w-full h-4 bg-gray-200 rounded-full overflow-hidden"
        role="progressbar"
        aria-valuenow={pct}
        aria-valuemin={0}
        aria-valuemax={100}
        aria-label={`${label}: ${pct.toString()}%`}
      >
        <div
          className={`h-full rounded-full transition-all duration-700 ease-out ${getBarColor(value)}`}
          style={{ width: `${pct.toString()}%` }}
        />
      </div>
      <p className="text-xs text-gray-500">{description}</p>
    </div>
  );
}

/** Exibe métricas de similaridade (cosseno e Jaccard) com barras de progresso coloridas. */
export default function Results({
  cosineSimilarity,
  jaccardCoefficient,
  stats,
}: ResultsProps): React.JSX.Element {
  return (
    <div className="space-y-6 p-6 bg-white rounded-xl border border-gray-200 shadow-sm">
      <div className="flex justify-between items-center">
        <h2 className="text-lg font-bold text-gray-800">
          Resultados da Comparação
        </h2>
        <div className="flex gap-4 text-xs text-gray-500">
          <span>Texto 1: {stats.text1_words} palavras</span>
          <span>Texto 2: {stats.text2_words} palavras</span>
          <span>Processado em {stats.processing_time_ms}ms</span>
        </div>
      </div>
      <MetricBar
        label="Similaridade por Cosseno"
        value={cosineSimilarity}
        description="Mede a proximidade direcional dos vetores TF-IDF (0 = nenhuma, 1 = idênticos)"
      />
      <MetricBar
        label="Coeficiente de Jaccard"
        value={jaccardCoefficient}
        description="Razão entre termos compartilhados e total de termos únicos"
      />
    </div>
  );
}
