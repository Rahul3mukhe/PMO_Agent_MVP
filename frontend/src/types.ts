export interface Project {
  project_id: string;
  project_name: string;
  project_type: string;
  sponsor?: string;
  estimated_budget?: number;
  actual_budget_consumed?: number;
  total_time_taken_days?: number;
  labour_cost?: number;
  development_cost?: number;
  test_cost?: number;
  software_cost?: number;
  infrastructure_cost?: number;
  overhead_cost?: number;
  timeline_summary?: string;
  scope_summary?: string;
  key_deliverables: string[];
  known_risks: string[];
}

export type DocStatus = "NOT_AVAILABLE" | "NOT_SUFFICIENT" | "SUFFICIENT" | "SUFFICIENT_WITH_FLAGS";

export interface DocumentArtifact {
  doc_type: string;
  title: string;
  content_markdown: string;
  status: DocStatus;
  reasons: string[];
  file_path?: string;
}

export interface GateResult {
  gate: string;
  passed: boolean;
  findings: string[];
}

export interface GenerationLogEntry {
  doc: string;
  provider: string;
  model: string;
  status: 'ok' | 'fallback' | 'failed';
  note: string;
}

export interface PMOState {
  project: Project;
  docs: Record<string, DocumentArtifact>;
  gates: GateResult[];
  decision?: string;
  summary?: string;
  audit: Record<string, any> & {
    generation_log?: GenerationLogEntry[];
  };
}
