import json
from nodes.base import BaseNode
from schemas import PMOState
from llm_providers import generate_text

class ProjectExtractor(BaseNode):
    def __call__(self, state: PMOState) -> PMOState:
        raw_text = state.audit.get("raw_upload_text", "")
        if not raw_text.strip():
            # Fallback if no documents were uploaded yet
            state.project.project_id = state.project.project_id or "PRJ-001"
            state.project.project_name = state.project.project_name or "New Uploaded Project"
            return state

        prompt = f"""
You are an AI extracting project context from unstructured documents.
Read the following text and extract details to match this JSON schema precisely:
{{
  "project_id": "string (e.g. PRJ-123)",
  "project_name": "string",
  "project_type": "string",
  "sponsor": "string",
  "estimated_budget": "number or null",
  "actual_budget_consumed": "number or null",
  "total_time_taken_days": "number or null",
  "timeline_summary": "string",
  "scope_summary": "string",
  "key_deliverables": ["string"],
  "known_risks": ["string"]
}}

CRITICAL INSTRUCTIONS:
1. For budget fields (estimated_budget, actual_budget_consumed), look for any currency values.
2. If a value is NOT found, set it to null in the JSON. NEVER default to 0.0 unless the text explicitly mentions zero.
3. If multiple values are found, use the most plausible "total" or "summary" figure.
4. Only output valid JSON. Output NOTHING else.

TEXT to analyze:
{raw_text[:12000]}
"""
        try:
            response = generate_text(
                provider=state.provider,
                model=state.model,
                prompt=prompt,
                api_key=state.audit.get("api_key"),
                temperature=0.0,
                max_tokens=2048,
                standards=state.standards,
                project=state.project,
                doc_type="extraction"
            )
            
            start = response.find("{")
            end = response.rfind("}") + 1
            if start != -1 and end != -1:
                json_str = response[start:end]
                data = json.loads(json_str)
                for k, v in data.items():
                    if hasattr(state.project, k):
                        # Only set if not already provided by user manually? 
                        # Actually the current logic overwrites. 
                        # We'll stick to current logic for now.
                        setattr(state.project, k, v)
        except Exception as e:
            state.audit["extraction_error"] = str(e)
        
        return state
