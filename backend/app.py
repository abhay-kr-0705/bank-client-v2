import os
import sys
import tempfile
import shutil
import uuid
import threading
from typing import Optional, List, Any, Union
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Query
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOCAL_TEMP_DIR = os.path.join(BASE_DIR, "uploads", "temp")
os.makedirs(LOCAL_TEMP_DIR, exist_ok=True)
os.environ["TEMP"] = LOCAL_TEMP_DIR
os.environ["TMP"] = LOCAL_TEMP_DIR
tempfile.tempdir = LOCAL_TEMP_DIR

from .models import ReportData
from .extractor import CaseExtractor
from .excel_generator import generate_excel_report
from .session_engine import get_session, SessionManager

UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
STATIC_DIR = os.path.join(BASE_DIR, "static")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

app = FastAPI(title="BankTech Valuation OCR & Dynamic Excel Engine", version="2.5.0")

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

# Global Lock to serialize document processing and prevent OOM on Render Free Tier (512MB RAM cap)
PROCESSING_LOCK = threading.Lock()

class FolderProcessRequest(BaseModel):
    folder_path: str
    session_id: Optional[str] = "default_session"
    api_key: Optional[str] = None
    model_name: Optional[str] = None

class SingleFileProcessRequest(BaseModel):
    file_path: str
    session_id: Optional[str] = "default_session"
    api_key: Optional[str] = None
    model_name: Optional[str] = None

class RemoveFileRequest(BaseModel):
    file_path: str
    session_id: Optional[str] = "default_session"

class SettingsRequest(BaseModel):
    gemini_api_key: str
    model_name: Optional[str] = "gemini-2.5-flash"

class CellUpdateRequest(BaseModel):
    session_id: Optional[str] = "default_session"
    coordinate: str
    value: Any

class ReportUpdateRequest(BaseModel):
    session_id: Optional[str] = "default_session"
    report_data: ReportData

@app.get("/api/health")
def get_health():
    return {
        "status": "healthy",
        "engine": "Offline RapidOCR & Browser Accelerated Engine",
        "render_tier": "Free Tier Safe (<512MB RAM)",
        "has_api_key": bool(SERVER_CONFIG["gemini_api_key"]),
        "model_name": SERVER_CONFIG["model_name"]
    }

@app.post("/api/settings")
def update_settings(req: SettingsRequest):
    SERVER_CONFIG["gemini_api_key"] = req.gemini_api_key.strip()
    if req.model_name:
        SERVER_CONFIG["model_name"] = req.model_name.strip()
    return {"message": "Settings updated successfully", "has_api_key": bool(SERVER_CONFIG["gemini_api_key"])}

# ----------------- Dynamic Template Endpoints ----------------- #

@app.post("/api/template/upload")
async def upload_custom_template(
    file: UploadFile = File(...),
    session_id: Optional[str] = Form("default_session")
):
    """Uploads a custom .xlsx template (blank or filled) and sanitizes it."""
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx Excel templates are supported.")

    session = get_session(session_id)
    temp_template = os.path.join(session.template_dir, f"uploaded_{file.filename}")
    
    with open(temp_template, "wb") as f:
        shutil.copyfileobj(file.file, f)

    res = session.set_custom_template(temp_template)
    return res

@app.post("/api/template/reset")
def reset_template(session_id: Optional[str] = Query("default_session")):
    """Resets the active template back to base India Shelter format."""
    session = get_session(session_id)
    return session.reset_to_default_template()

@app.get("/api/template/clean-preview")
def get_template_clean_preview(session_id: Optional[str] = Query("default_session")):
    """Returns the clean blank template grid JSON for direct in-browser inspection."""
    session = get_session(session_id)
    return session.get_template_clean_preview()

import json

# ----------------- Incremental File & OCR Ingestion ----------------- #

@app.post("/api/session/upload-files")
async def upload_session_files(
    files: List[UploadFile] = File(...),
    session_id: Optional[str] = Form("default_session"),
    api_key: Optional[str] = Form(None),
    client_text: Optional[str] = Form(None)
):
    """
    Ingests 1 or more files (PDF, DOCX, JPG, PNG, CSV, ZIP, XLSX) incrementally into session.
    Extracts text, performs OCR, updates report data, and returns live filled spreadsheet grid.
    Supports browser pre-extracted text to dramatically reduce Render Free Tier CPU/RAM consumption.
    """
    session = get_session(session_id)
    saved_paths = []

    for f in files:
        target_path = os.path.join(session.docs_dir, f.filename)
        os.makedirs(os.path.dirname(target_path), exist_ok=True)
        with open(target_path, "wb") as buffer:
            shutil.copyfileobj(f.file, buffer)
        saved_paths.append(target_path)

    # Save any client pre-extracted text as sidecars
    if client_text:
        try:
            client_dict = json.loads(client_text)
            if isinstance(client_dict, dict):
                for fname, txt in client_dict.items():
                    if txt and str(txt).strip():
                        clean_fname = os.path.basename(fname)
                        txt_path = os.path.join(session.docs_dir, f"{clean_fname}.client.txt")
                        with open(txt_path, "w", encoding="utf-8", errors="ignore") as tf:
                            tf.write(str(txt).strip())
        except Exception as e:
            print(f"[Client Text Warning: {e}]")

    active_api_key = api_key or SERVER_CONFIG["gemini_api_key"]
    model_name = SERVER_CONFIG["model_name"]

    def bg_worker():
        with PROCESSING_LOCK:
            try:
                session.ingest_documents_incrementally(
                    saved_paths,
                    api_key=active_api_key,
                    model_name=model_name
                )
            except Exception as e:
                session.update_progress(step=1, percent=0, message=f"Extraction Error: {str(e)}", state="error")
            finally:
                import gc
                gc.collect()

    threading.Thread(target=bg_worker, daemon=True).start()

    return {
        "success": True,
        "status": "processing",
        "message": f"Uploaded and queued {len(saved_paths)} file(s) for background extraction.",
        "session_id": session_id,
        "files_count": len(saved_paths)
    }

@app.post("/api/session/process-file")
def process_single_file_endpoint(req: SingleFileProcessRequest):
    """
    Processes an individual file in the session, extracts data, and merges into live spreadsheet.
    """
    session = get_session(req.session_id)
    active_api_key = req.api_key or SERVER_CONFIG["gemini_api_key"]
    return session.process_single_file(
        req.file_path,
        api_key=active_api_key,
        model_name=req.model_name or SERVER_CONFIG["model_name"]
    )

@app.post("/api/session/remove-file")
def remove_session_file(req: RemoveFileRequest):
    """Removes a file from the session."""
    session = get_session(req.session_id)
    return session.remove_file(req.file_path)

@app.get("/api/session/progress")
def get_session_progress(session_id: Optional[str] = Query("default_session")):
    """Returns real-time background processing metrics and current file for live progress tracking."""
    session = get_session(session_id)
    return getattr(session, "progress", {
        "state": "idle",
        "step": 0,
        "percent": 0,
        "message": "Ready",
        "current_file": "",
        "processed_files": 0,
        "total_files": 0,
        "elapsed_seconds": 0.0
    })

@app.get("/api/session/live-grid")
def get_live_grid(
    session_id: Optional[str] = Query("default_session"),
    sheet: Optional[str] = Query("0")
):
    """Returns the live populated spreadsheet grid JSON for the interactive web viewer."""
    session = get_session(session_id)
    sheet_val: Union[int, str] = int(sheet) if sheet.isdigit() else sheet
    return {
        "success": True,
        "grid": session.get_live_grid_preview(sheet=sheet_val),
        "report_data": session.report_data.dict(),
        "files": session.uploaded_files
    }

@app.post("/api/session/update-cell")
def update_grid_cell(req: CellUpdateRequest):
    """Updates a cell value manually edited in the web spreadsheet grid."""
    session = get_session(req.session_id)
    session.update_cell_value(req.coordinate, req.value)
    return {"success": True, "updated": req.coordinate, "value": req.value}

@app.post("/api/session/update-report")
def update_session_report(req: ReportUpdateRequest):
    """Updates session report data from frontend form sync and returns updated live grid."""
    session = get_session(req.session_id)
    grid = session.update_report_data(req.report_data)
    return {
        "success": True,
        "grid": grid,
        "report_data": session.report_data.dict()
    }

@app.post("/api/session/clear-data")
def clear_session_data(session_id: Optional[str] = Query("default_session")):
    """Clears data values from the live preview grid while preserving template headers and formulas."""
    session = get_session(session_id)
    grid = session.clear_all_data()
    return {"success": True, "grid": grid, "report_data": session.report_data.dict()}

@app.post("/api/session/add-row")
def add_grid_row(session_id: Optional[str] = Query("default_session")):
    """Adds a new row to the active spreadsheet grid."""
    session = get_session(session_id)
    grid = session.add_row()
    return {"success": True, "grid": grid}

@app.post("/api/session/add-column")
def add_grid_column(session_id: Optional[str] = Query("default_session")):
    """Adds a new column to the active spreadsheet grid."""
    session = get_session(session_id)
    grid = session.add_column()
    return {"success": True, "grid": grid}

@app.post("/api/session/download-excel")
def download_session_excel(session_id: Optional[str] = Query("default_session")):
    """Generates and downloads the final populated Excel file with 100% format preservation."""
    session = get_session(session_id)
    applicant = session.report_data.header.applicant_name or "Valuation_Report"
    app_id = session.report_data.header.application_id or ""
    safe_name = "".join(c for c in f"{applicant}_{app_id}" if c.isalnum() or c in ('_', '-')).strip('_')
    if not safe_name:
        safe_name = "Valuation_Report"
    
    filename = f"{safe_name}.xlsx"
    output_path = os.path.join(OUTPUT_DIR, filename)

    try:
        session.export_final_excel(output_path)
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

# ----------------- Direct Folder Ingestion ----------------- #

@app.post("/api/process-folder")
def process_folder(req: FolderProcessRequest):
    folder_path = req.folder_path.strip()
    if not os.path.isabs(folder_path):
        folder_path = os.path.join(BASE_DIR, folder_path)

    if not os.path.exists(folder_path):
        raise HTTPException(
            status_code=404,
            detail=f"Folder not found on host machine: '{folder_path}'. If you are accessing this app remotely from another PC or phone, please use the 'Add Files / ZIP' button to upload your documents directly!"
        )

    session = get_session(req.session_id)
    saved_files = []
    for root, _, files in os.walk(folder_path):
        for f in sorted(files):
            if not f.startswith("~$") and not f.startswith(".") and not "__MACOSX" in root:
                src = os.path.join(root, f)
                dst = os.path.join(session.docs_dir, f)
                shutil.copyfile(src, dst)
                saved_files.append(dst)

    active_api_key = req.api_key or SERVER_CONFIG["gemini_api_key"]
    model_name = req.model_name or SERVER_CONFIG["model_name"]

    def bg_worker():
        try:
            session.ingest_documents_incrementally(
                saved_files,
                api_key=active_api_key,
                model_name=model_name
            )
        except Exception as e:
            session.update_progress(step=1, percent=0, message=f"Extraction Error: {str(e)}", state="error")

    threading.Thread(target=bg_worker, daemon=True).start()

    return {
        "success": True,
        "status": "processing",
        "message": f"Queued {len(saved_files)} file(s) for background extraction.",
        "session_id": req.session_id,
        "source_folder": folder_path,
        "files_count": len(saved_files)
    }

@app.post("/api/generate-excel")
def generate_excel(report: ReportData):
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
