import os
import shutil
import zipfile
import uuid
import openpyxl
from typing import Dict, Any, List, Optional
from .models import ReportData
from .template_engine import TemplateEngine
from .smart_mapper import SmartFieldMapper
from .extractor import CaseExtractor

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "base_template.xlsx")
SESSIONS_DIR = os.path.join(BASE_DIR, "uploads", "sessions")
os.makedirs(SESSIONS_DIR, exist_ok=True)

class SessionManager:
    """
    Manages interactive case sessions, custom templates, incremental multi-file ingestion,
    ZIP unpacking, real-time spreadsheet grid state, and Excel export.
    """

    def __init__(self, session_id: Optional[str] = None):
        self.session_id = session_id or str(uuid.uuid4())[:8]
        self.session_dir = os.path.join(SESSIONS_DIR, self.session_id)
        self.docs_dir = os.path.join(self.session_dir, "documents")
        self.template_dir = os.path.join(self.session_dir, "template")
        
        os.makedirs(self.docs_dir, exist_ok=True)
        os.makedirs(self.template_dir, exist_ok=True)

        self.template_path = os.path.join(self.template_dir, "active_template.xlsx")
        self.sanitized_template_path = os.path.join(self.template_dir, "sanitized_template.xlsx")

        # Initialize with default template if not set
        if not os.path.exists(self.template_path):
            shutil.copyfile(DEFAULT_TEMPLATE_PATH, self.template_path)
            self._sanitize_active_template()

        self.report_data = ReportData()
        self.uploaded_files = []
        self.cell_overrides = {}

    def _sanitize_active_template(self):
        engine = TemplateEngine(self.template_path)
        engine.sanitize_template(self.sanitized_template_path)

    def set_custom_template(self, uploaded_file_path: str) -> Dict[str, Any]:
        """Sets a new custom template uploaded by the user, sanitizes it, and returns grid metadata."""
        shutil.copyfile(uploaded_file_path, self.template_path)
        self._sanitize_active_template()
        self.cell_overrides.clear()
        
        engine = TemplateEngine(self.template_path)
        clean_grid = engine.export_grid_json()
        return {
            "success": True,
            "template_name": os.path.basename(uploaded_file_path),
            "grid": clean_grid
        }

    def reset_to_default_template(self) -> Dict[str, Any]:
        """Resets active template back to base India Shelter template."""
        shutil.copyfile(DEFAULT_TEMPLATE_PATH, self.template_path)
        self._sanitize_active_template()
        self.cell_overrides.clear()
        
        engine = TemplateEngine(self.template_path)
        return {
            "success": True,
            "template_name": "base_template.xlsx (India Shelter Standard)",
            "grid": engine.export_grid_json()
        }

    def get_template_clean_preview(self) -> Dict[str, Any]:
        """Returns the blank sanitized template grid for in-browser viewing."""
        engine = TemplateEngine(self.sanitized_template_path)
        return engine.export_grid_json()

    def unpack_zip_if_needed(self, file_path: str) -> List[str]:
        """If file is a ZIP, extracts all contained documents and returns list of paths."""
        extracted_paths = []
        if file_path.lower().endswith(".zip"):
            extract_target = os.path.join(self.docs_dir, "unzipped_" + str(uuid.uuid4())[:6])
            os.makedirs(extract_target, exist_ok=True)
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(extract_target)
            
            for root, _, files in os.walk(extract_target):
                for f in files:
                    if not f.startswith("~$") and not f.startswith("."):
                        extracted_paths.append(os.path.join(root, f))
        else:
            extracted_paths.append(file_path)
        return extracted_paths

    def ingest_documents_incrementally(
        self,
        new_file_paths: List[str],
        api_key: Optional[str] = None,
        model_name: str = "gemini-2.5-flash"
    ) -> Dict[str, Any]:
        """
        Incrementally processes newly added documents (PDF, DOCX, Images, CSV, ZIP),
        extracts data via OCR/NLP, merges with session state, and maps to template.
        """
        all_new_files = []
        for p in new_file_paths:
            all_new_files.extend(self.unpack_zip_if_needed(p))

        extractor = CaseExtractor(api_key=api_key, model_name=model_name)
        new_report = extractor.process_case_folder(self.docs_dir)

        # Merge new extracted entities into active session report_data
        self._merge_report_data(new_report)

        # Keep file list updated
        for f in all_new_files:
            fname = os.path.basename(f)
            sz = os.path.getsize(f)
            if not any(item["path"] == f for item in self.uploaded_files):
                self.uploaded_files.append({
                    "filename": fname,
                    "path": f,
                    "size_display": f"{round(sz / 1024, 1)} KB",
                    "type": os.path.splitext(fname)[1].lower()
                })

        return {
            "success": True,
            "report_data": self.report_data.dict(),
            "files": self.uploaded_files,
            "grid": self.get_live_grid_preview()
        }

    def _merge_report_data(self, incoming: ReportData):
        """Merges incoming extracted data without overwriting already populated fields with blanks."""
        inc_dict = incoming.dict()
        curr_dict = self.report_data.dict()

        for section, sec_val in inc_dict.items():
            if isinstance(sec_val, dict):
                for k, v in sec_val.items():
                    if v is not None and v != "" and v != 0 and v != "NA":
                        curr_dict[section][k] = v
            elif isinstance(sec_val, str) and sec_val.strip():
                curr_dict[section] = sec_val

        self.report_data = ReportData(**curr_dict)

    def update_cell_value(self, coord: str, value: Any):
        """Updates a cell value manually edited in the web spreadsheet grid."""
        self.cell_overrides[coord] = value

    def get_live_grid_preview(self) -> Dict[str, Any]:
        """Returns the populated spreadsheet grid with current session values and formulas."""
        cell_values = SmartFieldMapper.map_to_cells(self.report_data)
        # Apply manual web overrides
        cell_values.update(self.cell_overrides)

        engine = TemplateEngine(self.sanitized_template_path)
        return engine.export_grid_json(filled_values=cell_values)

    def export_final_excel(self, output_path: str) -> str:
        """Generates the final Excel file preserving all original styling and formulas."""
        wb = openpyxl.load_workbook(self.sanitized_template_path, data_only=False)
        ws = wb.active

        cell_values = SmartFieldMapper.map_to_cells(self.report_data)
        cell_values.update(self.cell_overrides)

        for coord, val in cell_values.items():
            try:
                cell = ws[coord]
                # If template already had a formula and we have a non-formula override, preserve formula unless numeric
                if cell.value and str(cell.value).startswith("=") and not str(val).startswith("="):
                    continue  # Keep original Excel formula
                cell.value = val
            except Exception as e:
                print(f"[Export Cell Warning: {coord} -> {e}]")

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        wb.save(output_path)
        wb.close()
        return output_path

# Global Session Store (singleton or per-session)
GLOBAL_SESSIONS: Dict[str, SessionManager] = {}

def get_session(session_id: Optional[str] = None) -> SessionManager:
    s_id = session_id or "default_session"
    if s_id not in GLOBAL_SESSIONS:
        GLOBAL_SESSIONS[s_id] = SessionManager(session_id=s_id)
    return GLOBAL_SESSIONS[s_id]
