import os
import sys
import tempfile

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOCAL_TEMP_DIR = os.path.join(BASE_DIR, "uploads", "temp")
os.makedirs(LOCAL_TEMP_DIR, exist_ok=True)
os.environ["TEMP"] = LOCAL_TEMP_DIR
os.environ["TMP"] = LOCAL_TEMP_DIR
tempfile.tempdir = LOCAL_TEMP_DIR

import shutil
import uuid
from typing import Optional, List
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from .models import ReportData
from .extractor import CaseExtractor
from .excel_generator import generate_excel_report

UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
STATIC_DIR = os.path.join(BASE_DIR, "static")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

app = FastAPI(title="Banking Technical Valuation & OCR Extraction System", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"]
)

SERVER_CONFIG = {
    "gemini_api_key": os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY") or "",
    "model_name": "gemini-2.5-flash"
}

class FolderProcessRequest(BaseModel):
    folder_path: str
    api_key: Optional[str] = None
    model_name: Optional[str] = None

class SettingsRequest(BaseModel):
    gemini_api_key: str
    model_name: Optional[str] = "gemini-2.5-flash"

@app.get("/api/health")
def get_health():
    return {
        "status": "healthy",
        "has_api_key": bool(SERVER_CONFIG["gemini_api_key"]),
        "model_name": SERVER_CONFIG["model_name"]
    }

@app.get("/api/sample-cases")
def list_sample_cases():
    cases = []
    for item in os.listdir(BASE_DIR):
        item_path = os.path.join(BASE_DIR, item)
        if os.path.isdir(item_path) and item not in ["backend", "frontend", "static", "templates", "uploads", "outputs", ".git", ".idea", "__pycache__"]:
            files = [f for f in os.listdir(item_path) if os.path.isfile(os.path.join(item_path, f))]
            cases.append({
                "name": item,
                "path": item_path,
                "file_count": len(files)
            })
    return {"cases": cases}

@app.post("/api/settings")
def update_settings(req: SettingsRequest):
    SERVER_CONFIG["gemini_api_key"] = req.gemini_api_key.strip()
    if req.model_name:
        SERVER_CONFIG["model_name"] = req.model_name.strip()
    return {"message": "Settings updated successfully", "has_api_key": bool(SERVER_CONFIG["gemini_api_key"])}

@app.post("/api/process-folder")
def process_folder(req: FolderProcessRequest):
    folder_path = req.folder_path.strip()
    if not os.path.isabs(folder_path):
        folder_path = os.path.join(BASE_DIR, folder_path)

    if not os.path.exists(folder_path):
        raise HTTPException(status_code=404, detail=f"Folder path not found: {folder_path}")

    api_key = req.api_key or SERVER_CONFIG["gemini_api_key"]
    model_name = req.model_name or SERVER_CONFIG["model_name"]

    extractor = CaseExtractor(api_key=api_key, model_name=model_name)
    try:
        report_data = extractor.process_case_folder(folder_path)
        return {
            "success": True,
            "data": report_data.dict(),
            "source_folder": folder_path
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/upload")
async def upload_files(
    files: List[UploadFile] = File(...),
    case_name: Optional[str] = Form(None),
    api_key: Optional[str] = Form(None)
):
    session_id = str(uuid.uuid4())[:8]
    case_title = case_name.strip() if case_name else f"Case_{session_id}"
    target_dir = os.path.join(UPLOAD_DIR, f"{case_title}_{session_id}")
    os.makedirs(target_dir, exist_ok=True)

    for f in files:
        file_path = os.path.join(target_dir, f.filename)
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(f.file, buffer)

    active_api_key = api_key or SERVER_CONFIG["gemini_api_key"]
    extractor = CaseExtractor(api_key=active_api_key, model_name=SERVER_CONFIG["model_name"])
    
    try:
        report_data = extractor.process_case_folder(target_dir)
        report_data.case_name = case_title
        return {
            "success": True,
            "data": report_data.dict(),
            "source_folder": target_dir
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/generate-excel")
def generate_excel(report: ReportData):
    """Directly returns the binary Excel file with explicit attachment headers."""
    applicant = report.header.applicant_name or report.case_name or "Report"
    app_id = report.header.application_id or ""
    safe_name = "".join(c for c in f"{applicant}_{app_id}" if c.isalnum() or c in ('_', '-')).strip('_')
    if not safe_name:
        safe_name = "Valuation_Report"
    
    filename = f"{safe_name}.xlsx"
    output_path = os.path.join(OUTPUT_DIR, filename)

    try:
        generate_excel_report(report, output_path)
        headers = {
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Access-Control-Expose-Headers": "Content-Disposition"
        }
        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename,
            headers=headers
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Excel Generation Failed: {str(e)}")

@app.get("/api/download-excel/{filename}")
def download_excel(filename: str):
    file_path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Requested report file not found.")
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"',
        "Access-Control-Expose-Headers": "Content-Disposition"
    }
    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
        headers=headers
    )

@app.get("/api/view-file")
def view_case_file(filepath: str):
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail="File not found")
    
    ext = os.path.splitext(filepath)[1].lower()
    media_types = {
        ".pdf": "application/pdf",
        ".jpeg": "image/jpeg",
        ".jpg": "image/jpeg",
        ".png": "image/png",
        ".webp": "image/webp",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }
    media_type = media_types.get(ext, "application/octet-stream")
    return FileResponse(filepath, media_type=media_type)

app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
