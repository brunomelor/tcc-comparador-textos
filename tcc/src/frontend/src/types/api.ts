export interface SharedTerm {
  term: string;
  weight: number;
}

export interface Stats {
  text1_words: number;
  text2_words: number;
  processing_time_ms: number;
}

export interface CompareResult {
  cosine_similarity: number;
  jaccard_coefficient: number;
  shared_terms: SharedTerm[];
  text1_unique_terms: string[];
  text2_unique_terms: string[];
  stats: Stats;
}

export interface CompareRequest {
  text1: string;
  text2: string;
}
