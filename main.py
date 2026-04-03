"""
AuditX Report API
Accepts audit JSON → returns KPMG-grade Excel file
Deploy on Render.com (free tier)
"""
from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Any, Dict
import io, json, sys, os

sys.path.insert(0, os.path.dirname(__file__))
from generate_report import generate_report

app = FastAPI(title="AuditX Report API", version="1.0.0")

# Allow requests from your Netlify domain (and localhost for testing)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],   # tighten to your Netlify URL after testing
    allow_methods=["POST", "GET", "OPTIONS"],
    allow_headers=["*"],
)

@app.get("/")
def root():
    return {"status": "AuditX Report API is running", "version": "1.0.0"}

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/generate-report")
async def generate_excel_report(payload: Dict[str, Any]):
    """
    Accepts audit JSON, returns a formatted KPMG-grade Excel file.
    """
    try:
        company = str(payload.get("companyName", "Audit")).replace(" ", "_")[:30]
        filename = f"{company}_KPMG_Report.xlsx"

        # Write to a temp in-memory buffer
        import tempfile, os
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name

        generate_report(payload, tmp_path)

        with open(tmp_path, "rb") as f:
            excel_bytes = f.read()
        os.unlink(tmp_path)

        return StreamingResponse(
            io.BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f'attachment; filename="{filename}"',
                "Content-Length": str(len(excel_bytes)),
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
