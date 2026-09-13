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

        # 2. Valuation Draft or Technical Report PDF (Highest content authority)
        report_keywords = [
            "report", "valuation", "technical", "application id", "applicant name",
            "type of property", "india shelter", "shelter report", "survey",
            "inspection report", "adopted land", "built up area", "property limit"
        ]
        if ext == ".pdf" and (any(kw in lower_name for kw in ["valuation", "report", "technical", "sheet"]) or any(kw in lower_text for kw in report_keywords)):
            return {
                "category": cls.CATEGORY_VALUATION_DRAFT,
                "label": "Valuation Technical Sheet",
                "priority": 10,
                "is_report": True
            }

        # 3. Field Notes DOCX (Inspection remarks & site observations)
        if ext in [".docx", ".doc"] or "field notes" in lower_name or "visit" in lower_name:
            return {
                "category": cls.CATEGORY_FIELD_NOTES,
                "label": "Field Inspection Notes (DOCX)",
                "priority": 9,
                "is_narrative_source": True
            }

        # 4. Title Deeds & Legal Ownership Chains (GPA, ATS, Sale Deed, Will)
        deed_keywords = ["gpa", "ats", "deed", "sale deed", "will", "possession", "khasra", "agreement to sell", "power of attorney", "registry", "paper"]
        if any(kw in lower_name for kw in deed_keywords) or any(kw in lower_text for kw in deed_keywords):
            return {
                "category": cls.CATEGORY_TITLE_DEED,
                "label": "Title Deed / Ownership Chain",
                "priority": 6,
                "is_legal": True
            }

        # 5. Generic PDF (Check if draft report vs legal deed fallback)
        if ext == ".pdf":
            # If PDF has digital text mentioning applicant/application/property, treat as valuation report
            if any(kw in lower_text for kw in ["applicant", "application", "property no", "khasra"]):
                return {
                    "category": cls.CATEGORY_VALUATION_DRAFT,
                    "label": "Valuation Report PDF",
                    "priority": 10,
                    "is_report": True
                }
            return {
                "category": cls.CATEGORY_TITLE_DEED,
                "label": "Legal / Property PDF",
                "priority": 6,
                "is_legal": True
            }

        # 6. Supporting Documents (Tax receipts, Electricity Bills)
        if "bill" in lower_name or "tax" in lower_name or "receipt" in lower_name or "electricity" in lower_name:
            return {
                "category": cls.CATEGORY_SUPPORTING_DOC,
                "label": "Supporting Document",
                "priority": 4,
                "is_supporting": True
            }

        # 7. Site Photos / Meter Photos (Visual evidence & GPS coordinates)
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
                "priority": 2,
                "has_gps": has_exif_gps
            }

        # Fallback
        return {
            "category": cls.CATEGORY_SUPPORTING_DOC,
            "label": "Supporting Document",
            "priority": 3,
            "is_supporting": True
        }
