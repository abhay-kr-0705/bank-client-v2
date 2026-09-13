"""
Document Classifier Layer
Classifies input documents into semantic banking & valuation roles:
- Title Deeds & Chains (GPA, ATS, Sale Deed, Will, Possession)
- Field Engineer Notes (DOCX observations, 13-point summaries)
- Valuation Drafts (PDF reports)
- Site & Meter Photos (JPEG/PNG with or without GPS EXIF)
- Target Excel Templates (.xlsx)
- Other Supporting Documents (Bills, Tax Receipts)
"""

import os
import re
from typing import Dict, Any

class DocumentClassifier:
    CATEGORY_TITLE_DEED = "TITLE_DEED_OR_CHAIN"
    CATEGORY_FIELD_NOTES = "FIELD_NOTES_DOCX"
    CATEGORY_VALUATION_DRAFT = "VALUATION_DRAFT_PDF"
    CATEGORY_SITE_PHOTO = "SITE_PHOTO_IMAGE"
    CATEGORY_EXCEL_TEMPLATE = "EXCEL_TEMPLATE"
    CATEGORY_SUPPORTING_DOC = "SUPPORTING_DOC"

    @classmethod
    def classify_document(cls, filename: str, preview_text: str = "", has_exif_gps: bool = False) -> Dict[str, Any]:
        """
        Classifies an individual document based on filename patterns, file extension, and preview content.
        """
        lower_name = filename.lower()
        ext = os.path.splitext(lower_name)[1]
        lower_text = (preview_text or "").lower()[:2000]

        # 1. Target Excel Template
        if ext in [".xlsx", ".xls", ".xlsm"]:
            return {
                "category": cls.CATEGORY_EXCEL_TEMPLATE,
                "label": "Excel Valuation Template",
                "priority": 10,
                "is_template": True
            }

        # 2. Field Notes DOCX
        if ext == ".docx" or "field notes" in lower_name or "visit" in lower_name:
            return {
                "category": cls.CATEGORY_FIELD_NOTES,
                "label": "Field Inspection Notes (DOCX)",
                "priority": 9,
                "is_narrative_source": True
            }

        # 3. Site Photos / Meter Photos
        if ext in [".jpeg", ".jpg", ".png", ".webp", ".bmp", ".tiff"]:
            photo_type = "Site Photo"
            if "meter" in lower_name or "meter" in lower_text:
                photo_type = "Meter Photo"
            elif "road" in lower_name or "street" in lower_name:
                photo_type = "Road / Approach Photo"
            elif has_exif_gps:
                photo_type = "Geo-Tagged Site Photo"

            return {
                "category": cls.CATEGORY_SITE_PHOTO,
                "label": photo_type,
                "priority": 8,
                "has_gps": has_exif_gps
            }

        # 4. Title Deeds & Legal Ownership Chains
        deed_keywords = ["gpa", "ats", "deed", "sale deed", "will", "possession", "khasra", "agreement to sell", "power of attorney", "registry"]
        if any(kw in lower_name for kw in deed_keywords) or any(kw in lower_text for kw in deed_keywords):
            return {
                "category": cls.CATEGORY_TITLE_DEED,
                "label": "Title Deed / Ownership Chain",
                "priority": 7,
                "is_legal": True
            }

        # 5. Valuation Draft or Technical Report PDF
        if ext == ".pdf":
            if "valuation" in lower_name or "report" in lower_name or "technical" in lower_name or "application id" in lower_text:
                return {
                    "category": cls.CATEGORY_VALUATION_DRAFT,
                    "label": "Valuation Technical Sheet",
                    "priority": 6,
                    "is_report": True
                }
            return {
                "category": cls.CATEGORY_TITLE_DEED,
                "label": "Legal / Property PDF",
                "priority": 5,
                "is_legal": True
            }

        # 6. Supporting Documents
        return {
            "category": cls.CATEGORY_SUPPORTING_DOC,
            "label": "Supporting Document",
            "priority": 1,
            "is_supporting": True
        }
