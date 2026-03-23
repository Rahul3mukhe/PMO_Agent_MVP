# nodes/generator.py
#
# GeneratorNode  — generates all missing docs for the current PMO run.
# RepairNode     — re-generates docs flagged NOT_SUFFICIENT by ValidatorNode.
#
# Risk Registry is routed through risk_registry_generator.py which calls the
# LLM for structured JSON, builds guardrail-valid markdown, AND produces a
# full professional Word document in one pipeline call.

from nodes.base import BaseNode
from schemas import PMOState
from llm_providers import generate_text


def build_generation_prompt(state: PMOState, doc_type: str) -> str:
    org     = state.standards["org"]["name"]
    doc_std = state.standards["docs"][doc_type]
    proj    = state.project

    sections = "\n".join([f"- {s}" for s in doc_std["required_sections"]])

    return f"""
You are a PMO governance documentation assistant for {org}.
Generate a professional internal document in Markdown.

Document type: {doc_std["title"]}

STRICT RULES:
- Use headings EXACTLY as: "## <Section Name>" for each required section.
- Include ALL required sections in this order:
{sections}
- Use a formal, audit-ready tone.
- Be specific; include numbers/ranges where appropriate.
- Avoid placeholders like TBD, N/A, lorem, etc.
- IMPORTANT: For sections that imply a list (like "Risk List", "Key Deliverables", "Roles"), use Markdown bullet points ("- "). Generate at least 5 detailed items for these lists to satisfy PMO standards.
- Include an Approvals section with placeholders for role-based approvals (Name/Role/Date).

Project context:
- Project ID: {proj.project_id}
- Project Name: {proj.project_name}
- Project Type: {proj.project_type}
- Sponsor: {proj.sponsor}
- Estimated Budget: {proj.estimated_budget}
- Actual Budget Consumed: {proj.actual_budget_consumed}
- Total Time Taken (days): {proj.total_time_taken_days}
- Timeline Summary: {proj.timeline_summary}
- Scope Summary: {proj.scope_summary}
- Key Deliverables: {proj.key_deliverables}
- Known Risks: {proj.known_risks}

Now output ONLY the full Markdown document.
""".strip()


def _needed_docs(state: PMOState) -> set[str]:
    return set(state.standards["docs"].keys())


class GeneratorNode(BaseNode):
    def __call__(self, state: PMOState) -> PMOState:
        # Import here so the module loads even if the file isn't present
        # (e.g. during unit tests that mock out the graph).
        from risk_registry_generator import generate_risk_registry_artifact

        needed    = _needed_docs(state)
        generated = []

        for d in needed:
            art = state.docs[d]

            # Skip docs that already have content (uploaded by user)
            if art.content_markdown and art.content_markdown.strip():
                continue

            # ── Risk Registry: dedicated structured pipeline ──────────────────
            if d == "risk_registry":
                state.docs[d] = generate_risk_registry_artifact(state)
                generated.append(d)
                continue

            # ── All other doc types: generic LLM → markdown ──────────────────
            prompt = build_generation_prompt(state, d)

            try:
                md = generate_text(
                    provider=state.provider,
                    model=state.model,
                    prompt=prompt,
                    project=state.project,
                    doc_type=d,
                    standards=state.standards,
                    api_key=state.audit.get("api_key"),
                    temperature=state.audit.get("temperature", 0.0),
                    max_tokens=state.audit.get("max_tokens", 2048),
                )
            except Exception as gen_err:
                from llm_providers import _local_template_generate
                md = _local_template_generate(state.project, d, state.standards)
                state.audit[f"gen_fallback_{d}"] = str(gen_err)

            if md and "<!-- FALLBACK" in md:
                md = md[:md.index("<!-- FALLBACK")].rstrip()

            art.content_markdown = md or ""
            art.status  = "NOT_SUFFICIENT"
            art.reasons = ["Generated; pending validation"]
            generated.append(d)

        state.audit["generated_docs"] = generated
        return state


class RepairNode(BaseNode):
    def __call__(self, state: PMOState) -> PMOState:
        from risk_registry_generator import repair_risk_registry_artifact

        needed  = _needed_docs(state)
        repaired = []

        for d in needed:
            art = state.docs[d]
            if art.status != "NOT_SUFFICIENT":
                continue

            # ── Risk Registry: dedicated repair pipeline ──────────────────────
            if d == "risk_registry":
                state.docs[d] = repair_risk_registry_artifact(state)
                repaired.append(d)
                continue

            # ── All other doc types: generic repair prompt ────────────────────
            issues = "\n".join([f"- {r}" for r in art.reasons])
            prompt = build_generation_prompt(state, d) + (
                f"\n\nREPAIR REQUIRED. Fix these issues:\n{issues}\n"
            )

            try:
                md = generate_text(
                    provider=state.provider,
                    model=state.model,
                    prompt=prompt,
                    project=state.project,
                    doc_type=d,
                    standards=state.standards,
                    api_key=state.audit.get("api_key"),
                    temperature=state.audit.get("temperature", 0.0),
                    max_tokens=state.audit.get("max_tokens", 2048),
                )
            except Exception as rep_err:
                from llm_providers import _local_template_generate
                md = _local_template_generate(state.project, d, state.standards)
                state.audit[f"repair_fallback_{d}"] = str(rep_err)

            if md and "<!-- FALLBACK" in md:
                md = md[:md.index("<!-- FALLBACK")].rstrip()

            art.content_markdown = md
            art.status  = "NOT_SUFFICIENT"
            art.reasons = ["Regenerated; pending validation"]
            repaired.append(d)

        state.audit["repaired_docs"] = repaired
        return state