from nodes.base import BaseNode
from schemas import PMOState
from decisioning import evaluate_gates

class DecisionNode(BaseNode):
    def __call__(self, state: PMOState) -> PMOState:
        return evaluate_gates(state)
