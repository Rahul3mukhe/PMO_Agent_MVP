from abc import ABC, abstractmethod
from schemas import PMOState

class BaseNode(ABC):
    @abstractmethod
    def __call__(self, state: PMOState) -> PMOState:
        pass
