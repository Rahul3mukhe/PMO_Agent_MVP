from nodes.base import BaseNode
from schemas import PMOState, DocumentArtifact
from decisioning import compute_requirements

class RequirementsNode(BaseNode):
    def __call__(self, state: PMOState) -> PMOState:
        state.required_docs = compute_requirements(state.standards, state.project.project_type)
        state.audit["requirements"] = state.required_docs
        return state

class InitDocsNode(BaseNode):
    def __call__(self, state: PMOState) -> PMOState:
        for doc_type, meta in state.standards["docs"].items():
            if doc_type not in state.docs:
                state.docs[doc_type] = DocumentArtifact(
                    doc_type=doc_type,
                    title=meta["title"]
                )
        return state
