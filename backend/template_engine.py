import os
import re
import openpyxl
from openpyxl.utils import get_column_letter
from typing import Dict, Any, List, Optional, Union

class TemplateEngine:
    """
    Intelligently analyzes, sanitizes, and maps arbitrary Excel templates.
    Preserves all headers, labels, formulas, merged cells, borders, fonts, and Sheet2 dropdowns.
    Supports multi-sheet grid export.
    """

    KNOWN_LABELS = {
        "report title", "application id", "applicant name", "borrower", "customer name",
        "type of property", "percentage of completion", "structure type", "age of property",
        "dwelling units", "geo tag", "latitude", "longitude", "street name", "nearest landmark",
        "village", "city", "plot", "house", "flat", "khasra", "floor", "colony",
        "address as per site", "property address as per documents", "pincode", "district",
        "east", "west", "north", "south", "boundary matching", "mismatch remarks", "occupancy",
        "length", "breadth", "land area", "adopted land area", "per unit land rate", "total land value",
        "solar", "roof", "cracks", "parapet", "no. of floor", "toilet", "lift", "electricity",
        "person meet", "relation", "sanction plan", "public road", "opinion", "remarks",
        "floors", "basement", "stilt", "ground floor", "first floor", "second floor",
        "third floor", "fourth floor", "fifth floor", "six floor", "seven floor"
    }

    KNOWN_DATA_COORDS = {
        "B2", "D2", "B3", "B4", "B5", "B6", "B7", "B8", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "B17", "B18", "B19", "D19",
        "A22", "B22", "C22", "D22", "A25", "B25", "C25", "D25", "B26", "B27", "B28",
        "B30", "B31", "B32", "B33", "B34", "B35", "B37", "D37", "B38", "B39", "B40", "B41", "B42",
        "B44", "D44", "B45", "B46", "B47", "B64", "C66",
        "B76", "D76", "B77", "D77", "B78", "D78",
        "B88", "B90", "D90", "B91", "D91", "B92", "D92", "B93", "D93", "B94", "D94", "B95", "D95", "B96", "D96", "B97", "D97", "B98", "D98", "B99", "D99",
        "B102", "B103", "B104", "A106"
    }

    def __init__(self, template_path: str):
        self.template_path = template_path
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found at: {template_path}")

    @staticmethod
    def is_formula(val: Any) -> bool:
        if isinstance(val, str) and val.strip().startswith("="):
            return True
        return False

    def is_header_or_label(self, cell_val: Any, col_idx: int, row_idx: int, coord: Optional[str] = None) -> bool:
        """Determines if a cell is a static label/header or a data cell."""
        if cell_val is None:
            return False
        
        str_val = str(cell_val).strip()
        if not str_val:
            return False

        if self.is_formula(str_val):
            return False

        # If it is a known data coordinate or data area, it is NOT a header/label
        if coord and coord in self.KNOWN_DATA_COORDS:
            return False
        if row_idx in (22, 25) or row_idx >= 106:
            return False
        if 52 <= row_idx <= 61 and col_idx in (2, 3, 4):
            return False
        if col_idx in (2, 4) and row_idx not in (1, 9, 20, 21, 23, 24, 29, 36, 43, 49, 50, 51, 67, 81, 84, 87, 100, 101, 105):
            return False

        # Row 1 header or title
        if row_idx == 1:
            return True

        val_lower = str_val.lower()
        for kw in self.KNOWN_LABELS:
            if kw in val_lower and len(str_val) < 60:
                return True

        # Columns A and C in standard valuation templates are almost always label columns
        if col_idx in [1, 3] and not str_val.replace('.', '', 1).isdigit():
            if len(str_val) < 85:
                return True

        return False

    def get_sheet_names(self) -> List[str]:
        wb = openpyxl.load_workbook(self.template_path, read_only=True)
        names = wb.sheetnames
        wb.close()
        return names

    def sanitize_template(self, output_sanitized_path: str) -> Dict[str, Any]:
        """
        Creates a clean copy of the template where data cells are blanked out,
        while strictly preserving all headers, formulas, styles, and validations across all sheets.
        """
        wb = openpyxl.load_workbook(self.template_path, data_only=False)
        sheet_names = wb.sheetnames
        main_sheet_name = sheet_names[0]
        ws = wb[main_sheet_name]

        metadata_cells = {}
        cleared_count = 0
        preserved_formulas = 0

        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(row=r, column=c)
                val = cell.value
                coord = cell.coordinate

                if val is None:
                    continue

                if self.is_formula(val):
                    preserved_formulas += 1
                    metadata_cells[coord] = {
                        "type": "formula",
                        "formula": str(val),
                        "row": r,
                        "col": c
                    }
                elif self.is_header_or_label(val, c, r, coord):
                    metadata_cells[coord] = {
                        "type": "label",
                        "value": str(val),
                        "row": r,
                        "col": c
                    }
                else:
                    # It is an existing data value - clear it for clean template
                    metadata_cells[coord] = {
                        "type": "data_placeholder",
                        "original_value": str(val),
                        "row": r,
                        "col": c
                    }
                    cell.value = ""
                    cleared_count += 1

        os.makedirs(os.path.dirname(os.path.abspath(output_sanitized_path)), exist_ok=True)
        wb.save(output_sanitized_path)
        wb.close()

        return {
            "main_sheet": main_sheet_name,
            "all_sheets": sheet_names,
            "max_row": ws.max_row,
            "max_col": ws.max_column,
            "cleared_cells": cleared_count,
            "preserved_formulas": preserved_formulas,
            "metadata_cells": metadata_cells,
            "sanitized_file": output_sanitized_path
        }

    def export_grid_json(
        self,
        sheet: Union[str, int] = 0,
        filled_values: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        Generates full grid representation for in-browser spreadsheet rendering for a given sheet.
        Includes cell values, formulas, merged cells, headers, and colors.
        """
        wb = openpyxl.load_workbook(self.template_path, data_only=False)
        sheet_names = wb.sheetnames

        if isinstance(sheet, int):
            target_idx = max(0, min(sheet, len(sheet_names) - 1))
            ws = wb[sheet_names[target_idx]]
        else:
            ws = wb[sheet] if sheet in wb else wb.active

        # Extract merged ranges
        merged_ranges = [str(rng) for rng in ws.merged_cells.ranges]

        # Determine effective row/col counts
        max_r = max(ws.max_row, 1)
        max_c = max(ws.max_column, 4)

        rows_data = []
        for r in range(1, max_r + 1):
            cols_data = []
            for c in range(1, max_c + 1):
                cell = ws.cell(row=r, column=c)
                coord = cell.coordinate
                raw_val = cell.value

                display_val = ""
                is_formula_cell = False
                formula_str = None

                if filled_values and coord in filled_values:
                    display_val = filled_values[coord]
                elif raw_val is not None:
                    if self.is_formula(raw_val):
                        is_formula_cell = True
                        formula_str = str(raw_val)
                        display_val = formula_str
                    else:
                        display_val = str(raw_val)

                is_lbl = self.is_header_or_label(raw_val, c, r)

                cols_data.append({
                    "col": c,
                    "col_letter": get_column_letter(c),
                    "row": r,
                    "coord": coord,
                    "value": str(display_val) if display_val is not None else "",
                    "is_formula": is_formula_cell,
                    "formula": formula_str,
                    "is_header": is_lbl or (r == 1),
                    "is_bold": bool(cell.font and cell.font.bold),
                    "fill_color": cell.fill.start_color.rgb if (cell.fill and cell.fill.start_color and hasattr(cell.fill.start_color, 'rgb') and isinstance(cell.fill.start_color.rgb, str) and not cell.fill.start_color.rgb.startswith("00000000")) else None
                })
            rows_data.append({"row": r, "cells": cols_data})

        sheet_title = ws.title
        wb.close()

        return {
            "sheet_name": sheet_title,
            "all_sheets": sheet_names,
            "max_row": max_r,
            "max_col": max_c,
            "merged_ranges": merged_ranges,
            "rows": rows_data
        }
