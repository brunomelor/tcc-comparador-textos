import type { SharedTerm } from "../types/api";

interface SharedTermsProps {
  sharedTerms: SharedTerm[];
  text1UniqueTerms: string[];
  text2UniqueTerms: string[];
}

/** Exibe termos compartilhados em tabela e termos exclusivos de cada texto em badges. */
export default function SharedTerms({
  sharedTerms,
  text1UniqueTerms,
  text2UniqueTerms,
}: SharedTermsProps): React.JSX.Element {
  return (
    <div className="space-y-6 p-6 bg-white rounded-xl border border-gray-200 shadow-sm">
      <div>
        <h3 className="text-lg font-bold text-gray-800 mb-3">
          Termos Compartilhados
        </h3>
        {sharedTerms.length === 0 ? (
          <p className="text-sm text-gray-500">
            Nenhum termo em comum encontrado.
          </p>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-gray-200">
                  <th className="text-left py-2 px-3 text-gray-600 font-semibold">
                    Termo
                  </th>
                  <th className="text-right py-2 px-3 text-gray-600 font-semibold">
                    Peso TF-IDF
                  </th>
                  <th className="text-left py-2 px-3 text-gray-600 font-semibold w-40">
                    Relevância
                  </th>
                </tr>
              </thead>
              <tbody>
                {sharedTerms.map((t) => (
                  <tr
                    key={t.term}
                    className="border-b border-gray-100 hover:bg-gray-50"
                  >
                    <td className="py-2 px-3 font-mono text-gray-800">
                      {t.term}
                    </td>
                    <td className="py-2 px-3 text-right text-gray-700">
                      {t.weight.toFixed(4)}
                    </td>
                    <td className="py-2 px-3">
                      <div
                        className="w-full h-2 bg-gray-200 rounded-full overflow-hidden"
                        role="progressbar"
                        aria-valuenow={Math.round(t.weight * 100)}
                        aria-valuemin={0}
                        aria-valuemax={100}
                        aria-label={`Relevância de ${t.term}: ${Math.round(t.weight * 100).toString()}%`}
                      >
                        <div
                          className="h-full bg-blue-500 rounded-full"
                          style={{
                            width: `${Math.min(t.weight * 100, 100).toString()}%`,
                          }}
                        />
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      <div className="grid grid-cols-2 gap-4">
        <div>
          <h4 className="text-sm font-semibold text-gray-600 mb-2">
            Exclusivos do Texto 1
          </h4>
          <div className="flex flex-wrap gap-1">
            {text1UniqueTerms.map((t) => (
              <span
                key={t}
                className="px-2 py-0.5 text-xs bg-blue-100 text-blue-700 rounded-full"
                aria-label={`Termo exclusivo do Texto 1: ${t}`}
              >
                {t}
              </span>
            ))}
          </div>
        </div>
        <div>
          <h4 className="text-sm font-semibold text-gray-600 mb-2">
            Exclusivos do Texto 2
          </h4>
          <div className="flex flex-wrap gap-1">
            {text2UniqueTerms.map((t) => (
              <span
                key={t}
                className="px-2 py-0.5 text-xs bg-purple-100 text-purple-700 rounded-full"
                aria-label={`Termo exclusivo do Texto 2: ${t}`}
              >
                {t}
              </span>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}
