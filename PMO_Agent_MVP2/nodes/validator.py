from nodes.base import BaseNode
from schemas import PMOState
from guardrails import validate_doc

def _needed_docs(state: PMOState) -> set[str]:
    return set(state.standards["docs"].keys())

class ValidatorNode(BaseNode):
    def __call__(self, state: PMOState) -> PMOState:
        needed = _needed_docs(state)
        for d in needed:
            art = state.docs[d]
            status, reasons = validate_doc(d, art.content_markdown, state.standards)
            art.status = status
            art.reasons = reasons if reasons else art.reasons
        return state
