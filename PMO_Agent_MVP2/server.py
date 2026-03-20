import os
import io
import uvicorn
import json
import logging
from dotenv import load_dotenv
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from typing import List, Optional, Dict

from pmo_graph import build_graph, load_standards
from schemas import Project, PMOState, DocumentArtifact
from doc_templates import create_official_docx, create_decision_report_docx
import fastapi
from fastapi.responses import Response

load_dotenv()

app = FastAPI()

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

STANDARDS_PATH = "config/standards.yml"
standards = load_standards(STANDARDS_PATH)

def extract_text(filename: str, content: bytes) -> str:
    name = filename.lower()
    if name.endswith((".txt", ".md", ".json")):
        return content.decode("utf-8", errors="ignore")
    if name.endswith(".docx"):
        try:
            from docx import Document
            doc = Document(io.BytesIO(content))
            return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        except Exception: return ""
    if name.endswith(".pdf"):
        try:
            import PyPDF2
            reader = PyPDF2.PdfReader(io.BytesIO(content))
            return "\n".join(pg.extract_text() for pg in reader.pages if pg.extract_text())
        except Exception: return ""
    return ""

@app.post("/analyze")
async def analyze_project(
    project_data: str = Form(...),
    uploaded_mapping: str = Form("{}"),
    files: List[UploadFile] = File(None)
):
    try:
        logger.info(f"Starting analysis for project data: {project_data[:200]}...")
        data = json.loads(project_data)
        mapping = json.loads(uploaded_mapping)
        
        project = Project(**data)
        state = PMOState(
            project=project,
            standards=standards,
            provider="groq",
            model="llama-3.3-70b-versatile"
        )
        
        doc_mapping = {}
        raw_text_for_extraction = ""
        
        if files:
            for file in files:
                content = await file.read()
                text = extract_text(file.filename, content)
                
                # Check if this file is mapped to a specific doc type
                if file.filename in mapping and mapping[file.filename]:
                    doc_type = mapping[file.filename]
                    doc_mapping[doc_type] = (doc_mapping.get(doc_type, "") + "\n\n" + text).strip()
                else:
                    # Otherwise use it for general project extraction
                    raw_text_for_extraction += "\n\n" + text

        state.audit["uploaded_mapping"] = doc_mapping
        state.audit["raw_upload_text"] = raw_text_for_extraction
        
        graph = build_graph()
        result = graph.invoke(state)
        
        # LangGraph returns the state as a dict
        return result
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        logger.error(f"Analysis failed: {str(e)}\n{error_trace}")
        raise HTTPException(status_code=500, detail=f"Analysis failed: {str(e)}")

@app.post("/export/docx")
async def export_docx(state: PMOState, doc_type: str):
    if doc_type not in state.docs:
        raise HTTPException(status_code=404, detail=f"Document {doc_type} not found")
    
    doc = state.docs[doc_type]
    org = state.standards.get("org", {}).get("name", "PMO Agent")
    header = state.standards.get("org", {}).get("doc_header", "Internal")
    footer = state.standards.get("org", {}).get("doc_footer", "")
    
    docx_bytes = create_official_docx(
        org_name=org,
        header=header,
        footer=footer,
        title=doc.title or doc_type,
        md_content=doc.content_markdown
    )
    
    filename = f"{state.project.project_id}_{doc_type}.docx"
    return Response(
        content=docx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.post("/export/report")
async def export_report(state: PMOState):
    org = state.standards.get("org", {}).get("name", "PMO Agent")
    header = state.standards.get("org", {}).get("doc_header", "Internal")
    footer = state.standards.get("org", {}).get("doc_footer", "")
    
    docx_bytes = create_decision_report_docx(
        org_name=org,
        header=header,
        footer=footer,
        project_id=state.project.project_id,
        project_name=state.project.project_name,
        decision=state.decision or "PENDING",
        summary=state.summary or "",
        gates=state.gates,
        docs=state.docs
    )
    
    filename = f"{state.project.project_id}_PMO_Decision_Report.docx"
    return Response(
        content=docx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/config")
async def get_config():
    return {
        "doc_info": standards["docs"],
        "project_schema": Project.model_json_schema()
    }

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
