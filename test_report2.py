import urllib.request
import urllib.error
import json

data = json.dumps({
    "project": {"project_id": "P1", "project_name": "T", "project_type": "T"},
    "standards": {"org": {"name": "Test"}, "docs": {"risk_registry": {}}},
    "provider": "groq",
    "model": "llama",
    "docs": {"risk_registry": {"doc_type": "risk_registry", "title": "Test", "status": "NOT_SUFFICIENT", "reasons": ["bad"]}},
    "required_docs": {},
    "gates": [{"gate": "A", "passed": False, "findings": ["nope"]}],
    "decision": "INVALIDATE",
    "summary": "Tested bad",
    "audit": {}
}).encode('utf-8')

req = urllib.request.Request(
    'http://127.0.0.1:8000/export/report',
    data=data,
    headers={'Content-Type': 'application/json'}
)

try:
    response = urllib.request.urlopen(req)
    print('SUCCESS', response.status)
except urllib.error.HTTPError as e:
    print(f'ERROR: {e.code}')
    print(e.read().decode('utf-8'))
