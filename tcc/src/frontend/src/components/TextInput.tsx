import { useRef, useState, useMemo } from "react";

const ACCEPTED_FILE_TYPES = ".txt,.md,.csv,.log";

interface TextInputProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
}

/** Campo de entrada de texto com suporte a upload de arquivo (.txt, .md, .csv, .log). */
export default function TextInput({
  label,
  value,
  onChange,
  placeholder,
}: TextInputProps): React.JSX.Element {
  const fileRef = useRef<HTMLInputElement>(null);
  const [fileName, setFileName] = useState<string | null>(null);

  const handleFile = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (): void => {
      if (typeof reader.result === "string") {
        onChange(reader.result);
      }
    };
    reader.onerror = (): void => {
      setFileName(null);
    };
    reader.readAsText(file, "utf-8");
    e.target.value = "";
  };

  const handleTextChange = (val: string): void => {
    setFileName(null);
    onChange(val);
  };

  const wordCount = useMemo(
    () => value.split(/\s+/).filter(Boolean).length,
    [value],
  );

  return (
    <div className="flex flex-col gap-2 flex-1 min-w-0">
      <div className="flex items-center justify-between">
        <label className="text-sm font-semibold text-gray-700">{label}</label>
        <div className="flex items-center gap-2">
          {fileName && (
            <span className="text-xs text-green-600 truncate max-w-32">
              {fileName}
            </span>
          )}
          <button
            type="button"
            onClick={() => fileRef.current?.click()}
            className="text-xs text-blue-600 hover:text-blue-800 cursor-pointer"
          >
            Carregar arquivo
          </button>
          <input
            ref={fileRef}
            type="file"
            accept={ACCEPTED_FILE_TYPES}
            onChange={handleFile}
            className="hidden"
            aria-label={`Carregar arquivo para ${label}`}
          />
        </div>
      </div>
      <textarea
        value={value}
        onChange={(e) => handleTextChange(e.target.value)}
        placeholder={placeholder}
        rows={12}
        className="w-full rounded-lg border border-gray-300 p-3 text-sm text-gray-800 focus:border-blue-500 focus:ring-2 focus:ring-blue-200 outline-none resize-y"
        aria-label={label}
      />
      <span className="text-xs text-gray-400 text-right">
        {wordCount} {wordCount === 1 ? "palavra" : "palavras"}
      </span>
    </div>
  );
}
