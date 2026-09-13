import os
import shutil
import zipfile
import uuid
import openpyxl
from openpyxl.utils import get_column_letter
from typing import Dict, Any, List, Optional, Union
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
    Manages interactive case sessions with dynamic row/col additions,
    clear-data actions, multi-sheet switching, single-file & batch extraction,
    and Excel formula generation.
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

        # Initialize with default template if active template doesn't exist
        if not os.path.exists(self.template_path):
            shutil.copyfile(DEFAULT_TEMPLATE_PATH, self.template_path)
            self._sanitize_active_template()

        self.report_data = ReportData()
        self.uploaded_files: List[Dict[str, Any]] = []
        self.cell_overrides: Dict[str, Any] = {}
        self.extra_rows = 0
        self.extra_cols = 0
        self.active_sheet = 0

    def _sanitize_active_template(self):
        engine = TemplateEngine(self.template_path)
        engine.sanitize_template(self.sanitized_template_path)

    def set_custom_template(self, uploaded_file_path: str) -> Dict[str, Any]:
        """Sets a new custom template uploaded by user, sanitizes it, and returns grid metadata."""
        shutil.copyfile(uploaded_file_path, self.template_path)
        self._sanitize_active_template()
        self.cell_overrides.clear()
        self.extra_rows = 0
        self.extra_cols = 0
        
        engine = TemplateEngine(self.sanitized_template_path)
        clean_grid = engine.export_grid_json(sheet=self.active_sheet)
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
        self.extra_rows = 0
        self.extra_cols = 0
        
        engine = TemplateEngine(self.sanitized_template_path)
        return {
            "success": True,
            "template_name": "Base Template (India Shelter)",
            "grid": engine.export_grid_json(sheet=self.active_sheet)
        }

    def clear_all_data(self) -> Dict[str, Any]:
        """Clears all extracted/entered data while keeping template structure, formulas, and headers."""
        self.report_data = ReportData()
        self.cell_overrides.clear()
        return self.get_live_grid_preview()

    def add_row(self) -> Dict[str, Any]:
        """Adds a new row to the active spreadsheet grid."""
        self.extra_rows += 1
        return self.get_live_grid_preview()

    def add_column(self) -> Dict[str, Any]:
        """Adds a new column to the active spreadsheet grid."""
        self.extra_cols += 1
        return self.get_live_grid_preview()

    def get_template_clean_preview(self) -> Dict[str, Any]:
        """Returns the blank sanitized template grid for in-browser viewing."""
        engine = TemplateEngine(self.sanitized_template_path)
        return engine.export_grid_json(sheet=self.active_sheet)

    def unpack_zip_if_needed(self, file_path: str) -> List[str]:
        """If file is a ZIP, extracts all contained documents and returns list of paths."""
        extracted_paths = []
        if file_path.lower().endswith(".zip"):
            extract_target = os.path.join(self.docs_dir, "unzipped_" + str(uuid.uuid4())[:6])
            os.makedirs(extract_target, exist_ok=True)
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(extract_target)
            
            for root, _, files in os.walk(extract_target):
                for f in sorted(files):
                    if not f.startswith("~$") and not f.startswith(".") and not "__MACOSX" in root:
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
        Incrementally processes newly added documents (PDF, DOCX, Images, CSV, ZIP, Folder),
        extracts data via OCR/NLP, merges with session state, and maps to template.
        If any .xlsx template is present, auto-activates and sanitizes it as the target template.
        """
        all_new_files = []
        for p in new_file_paths:
            all_new_files.extend(self.unpack_zip_if_needed(p))

        # Copy any incoming external files to self.docs_dir
        saved_paths = []
        for src in all_new_files:
            if not os.path.exists(src):
                continue
            fname = os.path.basename(src)
            dst = os.path.join(self.docs_dir, fname)
            if os.path.abspath(src) != os.path.abspath(dst):
                shutil.copyfile(src, dst)
                saved_paths.append(dst)
            else:
                saved_paths.append(src)

        # Check if an Excel template was uploaded in the batch
        for f in saved_paths:
            if f.lower().endswith(".xlsx") and not os.path.basename(f).startswith("~$"):
                self.set_custom_template(f)

        # Collect all active non-temp files in docs_dir
        current_doc_paths = []
        for root, _, files in os.walk(self.docs_dir):
            for f in sorted(files):
                if not f.startswith("~$") and not f.startswith(".") and not f.lower().endswith(".xlsx"):
                    current_doc_paths.append(os.path.join(root, f))

        extractor = CaseExtractor(api_key=api_key, model_name=model_name)
        new_report = extractor.process_file_list(current_doc_paths)

        # Merge new extracted entities into active session report_data
        self._merge_report_data(new_report)

        # Update uploaded files metadata
        self.uploaded_files.clear()
        for root, _, files in os.walk(self.docs_dir):
            for f in sorted(files):
                if not f.startswith("~$") and not f.startswith("."):
                    full_p = os.path.join(root, f)
                    sz = os.path.getsize(full_p)
                    self.uploaded_files.append({
                        "filename": f,
                        "path": full_p,
                        "size_display": f"{round(sz / 1024, 1)} KB",
                        "type": os.path.splitext(f)[1].lower(),
                        "status": "Processed"
                    })

        return {
            "success": True,
            "report_data": self.report_data.dict(),
            "files": self.uploaded_files,
            "grid": self.get_live_grid_preview()
        }

    def process_single_file(
        self,
        file_path: str,
        api_key: Optional[str] = None,
        model_name: str = "gemini-2.5-flash"
    ) -> Dict[str, Any]:
        """
        Extracts entities from an individual document and merges them into active session state.
        Allows users to force-extract or update data per file.
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        # If it's an Excel template, activate it
        if file_path.lower().endswith(".xlsx"):
            self.set_custom_template(file_path)
            return {
                "success": True,
                "message": f"Template {os.path.basename(file_path)} set and sanitized.",
                "report_data": self.report_data.dict(),
                "files": self.uploaded_files,
                "grid": self.get_live_grid_preview()
            }

        extractor = CaseExtractor(api_key=api_key, model_name=model_name)
        file_report = extractor.extract_from_single_file(file_path)
        self._merge_report_data(file_report)

        # Update status in file list
        for item in self.uploaded_files:
            if item.get("path") == file_path:
                item["status"] = "Updated"

        return {
            "success": True,
            "message": f"Extracted data from {os.path.basename(file_path)} and merged into report.",
            "report_data": self.report_data.dict(),
            "files": self.uploaded_files,
            "grid": self.get_live_grid_preview()
        }

    def remove_file(self, file_path: str) -> Dict[str, Any]:
        """Removes a file from the session."""
        self.uploaded_files = [f for f in self.uploaded_files if f.get("path") != file_path and f.get("filename") != os.path.basename(file_path)]
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception:
                pass
        return {
            "success": True,
            "files": self.uploaded_files
        }

    def _merge_report_data(self, incoming: ReportData):
        """Merges incoming extracted data without overwriting already populated fields with blanks."""
        inc_dict = incoming.dict()
        curr_dict = self.report_data.dict()

        for section, sec_val in inc_dict.items():
            if isinstance(sec_val, dict):
                for k, v in sec_val.items():
                    if v is not None and v != "" and v != 0 and v != "NA":
                        # If incoming is list of floor items
                        if k == "floors" and isinstance(v, list):
                            for i, f_obj in enumerate(v):
                                if i < len(curr_dict[section]["floors"]):
                                    if f_obj.get("actual_area"):
                                        curr_dict[section]["floors"][i]["actual_area"] = f_obj["actual_area"]
                                    if f_obj.get("permissible_area"):
                                        curr_dict[section]["floors"][i]["permissible_area"] = f_obj["permissible_area"]
                                    if f_obj.get("adopted_area"):
                                        curr_dict[section]["floors"][i]["adopted_area"] = f_obj["adopted_area"]
                        else:
                            curr_dict[section][k] = v
            elif isinstance(sec_val, str) and sec_val.strip():
                curr_dict[section] = sec_val

        curr_dict["field_confidences"] = incoming.field_confidences
        curr_dict["overall_confidence"] = incoming.overall_confidence
        curr_dict["review_needed_count"] = incoming.review_needed_count

        self.report_data = ReportData(**curr_dict)

    def update_report_data(self, new_data: Union[ReportData, dict]) -> Dict[str, Any]:
        """Updates session report data from frontend form sync and re-validates."""
        if isinstance(new_data, dict):
            self.report_data = ReportData(**new_data)
        else:
            self.report_data = new_data
        from .validator import ValuationValidator
        self.report_data = ValuationValidator.validate_and_score(self.report_data)
        return self.get_live_grid_preview()

    def update_cell_value(self, coord: str, value: Any):
        """Updates a cell value manually edited in the web spreadsheet grid."""
        self.cell_overrides[coord] = value

    def get_live_grid_preview(self, sheet: Union[str, int] = 0) -> Dict[str, Any]:
        """Returns the populated spreadsheet grid with current session values, formulas, review flags, extra rows/cols."""
        self.active_sheet = sheet
        cell_values, cell_metadata = SmartFieldMapper.map_to_cells_with_metadata(self.report_data)
        # Apply manual web overrides
        cell_values.update(self.cell_overrides)

        engine = TemplateEngine(self.sanitized_template_path)
        base_grid = engine.export_grid_json(sheet=sheet, filled_values=cell_values if (sheet == 0 or sheet == "Sheet1") else None)
        
        # Decorate data cells with confidence and review status
        for row_obj in base_grid["rows"]:
            for cell_obj in row_obj["cells"]:
                coord = cell_obj["coord"]
                meta = cell_metadata.get(coord)
                if meta:
                    cell_obj["confidence"] = meta.get("confidence", 1.0)
                    cell_obj["needs_review"] = meta.get("needs_review", False)
                    cell_obj["review_reason"] = meta.get("review_reason")
                    if meta.get("needs_review"):
                        cell_obj["fill_color"] = "FFF3CD"
                else:
                    cell_obj["confidence"] = 1.0
                    cell_obj["needs_review"] = False
                    cell_obj["review_reason"] = None

        # Expand extra rows/cols if added by user
        current_rows = base_grid["rows"]
        max_c = base_grid["max_col"] + self.extra_cols
        
        if self.extra_cols > 0:
            for r_idx, row_obj in enumerate(current_rows):
                start_c = len(row_obj["cells"]) + 1
                for c in range(start_c, max_c + 1):
                    c_letter = get_column_letter(c)
                    coord = f"{c_letter}{row_obj['row']}"
                    val = self.cell_overrides.get(coord, "")
                    row_obj["cells"].append({
                        "col": c,
                        "col_letter": c_letter,
                        "row": row_obj["row"],
                        "coord": coord,
                        "value": str(val),
                        "is_formula": str(val).startswith("="),
                        "formula": str(val) if str(val).startswith("=") else None,
                        "is_header": False,
                        "is_bold": False,
                        "fill_color": None,
                        "needs_review": False,
                        "confidence": 1.0
                    })

        start_r = len(current_rows) + 1
        for r in range(start_r, start_r + self.extra_rows):
            cols_data = []
            for c in range(1, max_c + 1):
                c_letter = get_column_letter(c)
                coord = f"{c_letter}{r}"
                val = self.cell_overrides.get(coord, "")
                cols_data.append({
                    "col": c,
                    "col_letter": c_letter,
                    "row": r,
                    "coord": coord,
                    "value": str(val),
                    "is_formula": str(val).startswith("="),
                    "formula": str(val) if str(val).startswith("=") else None,
                    "is_header": False,
                    "is_bold": False,
                    "fill_color": None,
                    "needs_review": False,
                    "confidence": 1.0
                })
            current_rows.append({"row": r, "cells": cols_data})

        base_grid["max_row"] = len(current_rows)
        base_grid["max_col"] = max_c
        base_grid["overall_confidence"] = getattr(self.report_data, "overall_confidence", 1.0)
        base_grid["review_needed_count"] = getattr(self.report_data, "review_needed_count", 0)
        return base_grid

    def export_final_excel(self, output_path: str) -> str:
        """Generates the final Excel file preserving all original styling, formulas, and highlighting review items."""
        from openpyxl.styles import PatternFill
        from openpyxl.comments import Comment

        wb = openpyxl.load_workbook(self.sanitized_template_path, data_only=False)
        ws = wb.active

        review_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
        cell_values, cell_metadata = SmartFieldMapper.map_to_cells_with_metadata(self.report_data)
        cell_values.update(self.cell_overrides)

        for coord, val in cell_values.items():
            try:
                cell = ws[coord]
                # If template already had a formula and we have a non-formula override, preserve formula
                if cell.value and str(cell.value).startswith("=") and not str(val).startswith("="):
                    continue
                
                # Convert numbers if numeric string
                if isinstance(val, str) and not val.startswith("="):
                    if val.isdigit():
                        cell.value = int(val)
                    else:
                        try:
                            cell.value = float(val)
                        except ValueError:
                            cell.value = val
                else:
                    cell.value = val

                # Apply soft amber highlight and comment if field needs review
                meta = cell_metadata.get(coord)
                if meta and meta.get("needs_review"):
                    cell.fill = review_fill
                    conf_pct = int(meta.get("confidence", 0.5) * 100)
                    reason = meta.get("review_reason") or "Please verify this field value"
                    cell.comment = Comment(f"[REVIEW NEEDED] (Confidence: {conf_pct}%)\n{reason}", "Valuation Engine")
            except Exception as e:
                print(f"[Export Cell Warning: {coord} -> {e}]")

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        wb.save(output_path)
        wb.close()
        return output_path

# Global Session Store
GLOBAL_SESSIONS: Dict[str, SessionManager] = {}

def get_session(session_id: Optional[str] = None) -> SessionManager:
    s_id = session_id or "default_session"
    if s_id not in GLOBAL_SESSIONS:
        GLOBAL_SESSIONS[s_id] = SessionManager(session_id=s_id)
    return GLOBAL_SESSIONS[s_id]
