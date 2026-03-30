import requests
import json
import os

API_BASE = "http://localhost:8000"

# Mock PMOState (Branded and Merged)
mock_state = {
    "project": {
        "project_id": "NTT-2026-003",
        "project_name": "Tesseract Forge Commercialization & Hardening",
        "project_type": "Standard",
        "sponsor": "Executive Director",
        "estimated_budget": 1000000,
        "actual_budget_consumed": 770000
    },
    "standards": {
        "org": {
            "name": "NTT DATA",
            "doc_header": "NTT DATA - Internal Funding Approval",
            "doc_footer": "Commercial-in-Confidence | Page {PAGE}"
        },
        "docs": {
            "risk_registry": {
                "title": "Risk Registry",
                "required_sections": [
                    "Overview", "Risk Summary", "Detailed Risks", "Mitigations", 
                    "Owners", "Registry Overview", "Risk List", "Review Cadence", "Approvals"
                ]
            }
        }
    },
    "provider": "test_provider",
    "model": "test_model",
    "docs": {
        "risk_registry": {
            "doc_type": "risk_registry",
            "title": "Risk Registry",
            "content_markdown": "# Risk Registry & Assessment\nThis is a consolidated risk document for NTT DATA.",
            "status": "SUFFICIENT"
        }
    },
    "required_docs": {"START": ["risk_registry"]},
    "gates": [
        {"gate": "Start Gate", "passed": True, "findings": ["Risk registry is sufficient."]}
    ],
    "decision": "APPROVED",
    "summary": "Project approved for internal funding.",
    "audit": {}
}

def test_export_branded():
    print("Testing /export/docx for Branded Risk Registry...")
    response = requests.post(
        f"{API_BASE}/export/docx?doc_type=risk_registry",
        json=mock_state
    )
    if response.status_code == 200:
        print("✅ Branded DOCX Export SUCCESS")
        with open("test_ntt_registry.docx", "wb") as f:
            f.write(response.content)
        print("Saved test_ntt_registry.docx")
    else:
        print(f"❌ Branded DOCX Export FAILED: {response.status_code} - {response.text}")

if __name__ == "__main__":
    try:
        test_export_branded()
    except Exception as e:
        print(f"An error occurred: {e}")
