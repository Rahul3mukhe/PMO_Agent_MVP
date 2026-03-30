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
from fastapi.responses import Response, FileResponse
from fastapi.staticfiles import StaticFiles
 
load_dotenv()
 
app = FastAPI()

from fastapi.exceptions import RequestValidationError
from fastapi.requests import Request
from fastapi.responses import JSONResponse
import sys

@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    print(f"\n--- VALIDATION ERROR ON {request.url} ---", file=sys.stderr)
    for err in exc.errors():
        print(f"Location: {err.get('loc')} | msg: {err.get('msg')} | type: {err.get('type')}", file=sys.stderr)
    print("-------------------------------------------\n", file=sys.stderr)
    return JSONResponse(status_code=422, content={"detail": exc.errors()}) 
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
        except Exception:
            return ""
    if name.endswith(".pdf"):
        try:
            import PyPDF2
            reader = PyPDF2.PdfReader(io.BytesIO(content))
            return "\n".join(pg.extract_text() for pg in reader.pages if pg.extract_text())
        except Exception:
            return ""
    return ""
 
 
@app.post("/analyze")
async def analyze_project(
    project_data: str = Form(...),
    uploaded_mapping: str = Form("{}"),
    files: List[UploadFile] = File(None)
):
    try:
        logger.info(f"Starting analysis for project data: {project_data[:200]}...")
        data    = json.loads(project_data)
        mapping = json.loads(uploaded_mapping)
 
        project = Project(**data)
        state   = PMOState(
            project=project,
            standards=standards,
            provider="groq",
            model="llama-3.3-70b-versatile"
        )
 
        doc_mapping             = {}
        raw_text_for_extraction = ""
 
        if files:
            for file in files:
                content = await file.read()
                text    = extract_text(file.filename, content)
 
                if file.filename in mapping and mapping[file.filename]:
                    doc_type = mapping[file.filename]
                    doc_mapping[doc_type] = (
                        doc_mapping.get(doc_type, "") + "\n\n" + text
                    ).strip()
                else:
                    raw_text_for_extraction += "\n\n" + text
 
        state.audit["uploaded_mapping"] = doc_mapping
        state.audit["raw_upload_text"]  = raw_text_for_extraction
 
        graph  = build_graph()
        result = graph.invoke(state)

        # Attach LLM generation log so the frontend can show generation status
        from llm_providers import get_status_log, clear_status_log
        gen_log = get_status_log()
        if isinstance(result.get("audit"), dict):
            result["audit"]["generation_log"] = gen_log
        clear_status_log()

        # Remove un-serializable raw docx bytes before returning JSON
        # frontend will send back structured data and /export/docx will rebuild it
        if "audit" in result and "risk_registry_docx" in result["audit"]:
            del result["audit"]["risk_registry_docx"]

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
 
    doc  = state.docs[doc_type]
    proj = state.project
    org  = state.standards.get("org", {}).get("name", "PMO Agent")
 
    # ── Risk Registry: return the dedicated professional Word document ─────────
    # The risk_registry_generator pipeline stores the pre-built docx bytes in
    # state.audit["risk_registry_docx"] during the graph run.  We return those
    # bytes directly so the fully structured document (cover page, scoring
    # matrix, detail cards, heatmap, etc.) is delivered unchanged.
    if doc_type == "risk_registry":
        docx_bytes: Optional[bytes] = None
 
        # 1. Use pre-built bytes from the graph run (fastest path)
        if isinstance(state.audit.get("risk_registry_docx"), bytes):
            docx_bytes = state.audit["risk_registry_docx"]
            logger.info("Returning pre-built risk_registry docx from audit cache.")
 
        # 2. Rebuild from structured data if cache is missing
        #    (e.g. the client sent a hand-crafted PMOState without running the graph)
        if not docx_bytes:
            try:
                from risk_registry_generator import (
                    build_risk_registry_docx,
                    _fill_defaults,
                )
                raw_data = state.audit.get("risk_registry_structured_data", {})
                if not raw_data:
                    raw_data = _fill_defaults({}, proj)
                docx_bytes = build_risk_registry_docx(raw_data, state)
                logger.info("Rebuilt risk_registry docx from structured data.")
            except Exception as e:
                logger.warning(
                    f"risk_registry dedicated rebuild failed ({e}); "
                    "falling back to generic template."
                )
 
        # 3. Ultimate fallback — generic markdown → docx template
        if not docx_bytes:
            header = state.standards.get("org", {}).get("doc_header", "Internal")
            footer = state.standards.get("org", {}).get("doc_footer", "")
            docx_bytes = create_official_docx(
                org_name=org,
                header=header,
                footer=footer,
                title=doc.title or doc_type,
                md_content=doc.content_markdown,
            )
            logger.info("Using generic template fallback for risk_registry export.")
 
        filename = f"{proj.project_id}_{doc_type}.docx"
        return Response(
            content=docx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
 
    # ── All other document types: existing generic renderer ───────────────────
    header = state.standards.get("org", {}).get("doc_header", "Internal")
    footer = state.standards.get("org", {}).get("doc_footer", "")
 
    docx_bytes = create_official_docx(
        org_name=org,
        header=header,
        footer=footer,
        title=doc.title or doc_type,
        md_content=doc.content_markdown,
    )
 
    filename = f"{proj.project_id}_{doc_type}.docx"
    return Response(
        content=docx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )
 
 
@app.post("/export/pptx")
async def export_pptx(state: PMOState):
    from presentation_generator import generate_client_pptx
    try:
        pptx_bytes = generate_client_pptx(state)
        filename = f"{state.project.project_id}_Client_Status.pptx"
        return Response(
            content=pptx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        logger.error(f"Failed to generate PPTX: {e}")
        raise HTTPException(status_code=500, detail=f"PPTX Export Failed: {str(e)}")
@app.post("/export/report")
async def export_report(state: PMOState):
    try:
        org    = state.standards.get("org", {}).get("name", "PMO Agent")
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
    except Exception as e:
        logger.error(f"Failed to generate Decision Report: {e}")
        raise HTTPException(status_code=500, detail=f"Report Export Failed: {str(e)}")

 
 
@app.get("/config")
async def get_config():
    return {
        "doc_info":       standards["docs"],
        "project_schema": Project.model_json_schema()
    }
 
# ─────────────────────────────────────────────────────────────────────────────
# FRONTEND SERVING
# ─────────────────────────────────────────────────────────────────────────────
FRONTEND_PATH = os.path.join(os.path.dirname(__file__), "frontend", "dist")
 
if os.path.exists(FRONTEND_PATH):
    app.mount("/assets", StaticFiles(directory=os.path.join(FRONTEND_PATH, "assets")), name="assets")
 
    @app.get("/{full_path:path}")
    async def serve_react_app(full_path: str):
        # Serve the API usually takes precedence, but for anything else, serve index.html
        # Check if the path is an actual file in dist (like favicon, etc.)
        file_path = os.path.join(FRONTEND_PATH, full_path)
        if os.path.isfile(file_path):
            return FileResponse(file_path)
        # Otherwise, serve the SPA entry point
        return FileResponse(os.path.join(FRONTEND_PATH, "index.html"))
else:
    logger.warning(f"Frontend dist not found at {FRONTEND_PATH}. UI will not be served.")
 
if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
