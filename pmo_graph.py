import os
import yaml
import json
from typing import Dict, Set

from langgraph.graph import StateGraph, START, END

from schemas import PMOState, DocumentArtifact
from decisioning import compute_requirements, evaluate_gates
from guardrails import validate_doc
from llm_providers import generate_text
from storage import make_run_dir
from doc_templates import create_official_docx, create_decision_report_docx

from nodes.extractor import ProjectExtractor
from nodes.requirements import RequirementsNode, InitDocsNode
from nodes.uploader import LoadUploadedDocsNode
from nodes.generator import GeneratorNode, RepairNode
from nodes.validator import ValidatorNode
from nodes.decision import DecisionNode

def load_standards(path: str) -> Dict:
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def _needed_docs(state: PMOState) -> Set[str]:
    return set(state.standards["docs"].keys())

def should_repair(state: PMOState) -> str:
    needed = _needed_docs(state)
    if any(state.docs[d].status == "NOT_SUFFICIENT" for d in needed):
        # Only repair once
        if not state.audit.get("repaired_docs"):
            return "repair"
    return "decide"

def build_graph():
    g = StateGraph(PMOState)

    # Instantiate node classes
    extractor = ProjectExtractor()
    requirements = RequirementsNode()
    init_docs = InitDocsNode()
    uploader = LoadUploadedDocsNode()
    generator = GeneratorNode()
    validator = ValidatorNode()
    repair = RepairNode()
    decision = DecisionNode()

    g.add_node("extract_context", extractor)
    g.add_node("requirements", requirements)
    g.add_node("init_docs", init_docs)
    g.add_node("load_uploaded_docs", uploader)
    g.add_node("generate_missing_docs", generator)
    g.add_node("validate_docs", validator)
    g.add_node("repair_once", repair)
    g.add_node("validate_again", validator) # Validator can be reused
    g.add_node("decide", decision)

    g.add_edge(START, "extract_context")
    g.add_edge("extract_context", "requirements")
    g.add_edge("requirements", "init_docs")
    g.add_edge("init_docs", "load_uploaded_docs")
    g.add_edge("load_uploaded_docs", "generate_missing_docs")
    g.add_edge("generate_missing_docs", "validate_docs")

    g.add_conditional_edges("validate_docs", should_repair, {
        "repair": "repair_once",
        "decide": "decide"
    })

    g.add_edge("repair_once", "validate_again")
    g.add_edge("validate_again", "decide")
    g.add_edge("decide", END)

    return g.compile()
()