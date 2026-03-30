from nodes.base import BaseNode
from schemas import PMOState

class LoadUploadedDocsNode(BaseNode):
    def __call__(self, state: PMOState) -> PMOState:
        mapping = state.audit.get("uploaded_mapping", {})
        loaded = []
        
        for doc_type, txt in mapping.items():
            if doc_type in state.docs and txt and txt.strip():
                state.docs[doc_type].content_markdown = txt
                state.docs[doc_type].status = "NOT_SUFFICIENT"
                state.docs[doc_type].reasons = ["Uploaded by user; pending validation"]
                loaded.append(doc_type)

        state.audit["loaded_docs"] = loaded
        return state
