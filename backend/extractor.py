"""
Robust Offline Valuation Document Extractor
Implements 100% local, offline extraction without any cloud AI API key:
INPUT DOCUMENTS -> Document Classifier -> Document Processor -> OCR + Vision Layer -> Document Understanding + Semantic Extraction -> Structured JSON + Confidence Scores
"""

import os
import re
import json
import io
from typing import List, Dict, Any, Optional, Tuple
from PIL import Image
import fitz  # PyMuPDF
from docx import Document

from .models import (
    ReportData, HeaderInfo, AddressInfo, BoundariesInfo,
    SolarRoofVicinity, LandMeasurements, ConstructionFloors,
    AccommodationInfo, LegalStatutoryChecks, ReferenceEnquiry, FloorArea
)
from .classifier import DocumentClassifier
from .ocr_engine import OfflineOCREngine
from .validator import ValuationValidator

def extract_exif_gps(image_path: str) -> Optional[Tuple[float, float]]:
    """Extracts GPS coordinates from image EXIF metadata if present."""
    try:
        img = Image.open(image_path)
        exif = img._getexif()
        if not exif:
            return None
        gps_info = exif.get(34853)
        if not gps_info:
            return None
        
        def _convert_to_degrees(value):
            d = float(value[0])
            m = float(value[1])
            s = float(value[2])
            return d + (m / 60.0) + (s / 3600.0)

        lat_ref = gps_info.get(1, 'N')
        lat = _convert_to_degrees(gps_info[2])
        if lat_ref != 'N':
            lat = -lat

        lon_ref = gps_info.get(3, 'E')
        lon = _convert_to_degrees(gps_info[4])
        if lon_ref != 'E':
            lon = -lon

        return (round(lat, 6), round(lon, 6))
    except Exception:
        return None

def parse_docx_file(filepath: str) -> str:
    """Reads digital text and tables from Word docx."""
    try:
        doc = Document(filepath)
        texts = []
        for p in doc.paragraphs:
            if p.text.strip():
                texts.append(p.text.strip())
        for t in doc.tables:
            for row in t.rows:
                row_txt = " | ".join(c.text.strip() for c in row.cells if c.text.strip())
                if row_txt:
                    texts.append(row_txt)
        return "\n".join(texts)
    except Exception as e:
        return f"[DOCX Error: {e}]"

def parse_pdf_text_and_images(filepath: str, max_pages: int = 25) -> Tuple[str, List[bytes]]:
    """
    Extracts digital text and extracts page images.
    If a page has zero digital text (scanned PDF), runs local RapidOCR on the page pixmap.
    """
    text_content = []
    page_images = []
    try:
        doc = fitz.open(filepath)
        for page_num in range(min(len(doc), max_pages)):
            page = doc[page_num]
            txt = page.get_text()

            # If digital text exists, use it
            if txt.strip():
                text_content.append(f"--- [Page {page_num + 1}] ---\n{txt}")
            else:
                # Scanned page: render pixmap and run local OfflineOCREngine
                pix = page.get_pixmap(dpi=150)
                img_bytes = pix.tobytes("jpeg")
                ocr_text, _ = OfflineOCREngine.extract_text_from_image_bytes(img_bytes)
                if ocr_text.strip():
                    text_content.append(f"--- [Page {page_num + 1} (OCR)] ---\n{ocr_text}")
                page_images.append(img_bytes)
        doc.close()
    except Exception as e:
        text_content.append(f"[PDF Error: {e}]")
    return "\n\n".join(text_content), page_images

def parse_text_or_csv_file(filepath: str) -> str:
    """Reads plain text or CSV file content."""
    try:
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    except Exception as e:
        return f"[Text File Error: {e}]"


class CaseExtractor:
    def __init__(self, api_key: Optional[str] = None, model_name: str = "offline-engine"):
        self.api_key = api_key
        self.model_name = model_name

    def inspect_file(self, file_path: str) -> Dict[str, Any]:
        """
        Step 1 & 2: Classify document and process content via digital parsing or local OCR.
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        filename = os.path.basename(file_path)
        ext = os.path.splitext(filename)[1].lower()
        size = os.path.getsize(file_path)
        text = ""
        images = []
        gps = None

        if ext in [".jpeg", ".jpg", ".png", ".webp", ".bmp", ".tiff"]:
            gps_coords = extract_exif_gps(file_path)
            if gps_coords:
                gps = f"{gps_coords[0]}, {gps_coords[1]}"

        # Document Classifier
        classification = DocumentClassifier.classify_document(
            filename=filename,
            preview_text="",
            has_exif_gps=bool(gps)
        )

        # Document Processor
        if ext == ".xlsx":
            pass
        elif ext == ".docx":
            text = parse_docx_file(file_path)
        elif ext == ".pdf":
            text, images = parse_pdf_text_and_images(file_path, max_pages=20)
        elif ext in [".jpeg", ".jpg", ".png", ".webp", ".bmp", ".tiff"]:
            with open(file_path, "rb") as img_f:
                img_data = img_f.read()
                images.append(img_data)
                # Run local OCR on image only if it is not purely a site photo
                if classification.get("category") != "SITE_PHOTO_IMAGE":
                    ocr_t, _ = OfflineOCREngine.extract_text_from_image_bytes(img_data)
                    if ocr_t.strip():
                        text = f"[OCR from {filename}]:\n{ocr_t}"
        elif ext in [".txt", ".csv", ".log"]:
            text = parse_text_or_csv_file(file_path)

        return {
            "filename": filename,
            "path": file_path,
            "type": classification["category"],
            "classification_label": classification["label"],
            "priority": classification["priority"],
            "size_bytes": size,
            "size_display": f"{round(size / 1024, 1)} KB",
            "text": text,
            "images": images,
            "gps": gps
        }

    def inspect_folder(self, folder_path: str) -> Dict[str, Any]:
        """Scans folder, classifies all raw files, and aggregates text and metadata."""
        if not os.path.exists(folder_path):
            raise FileNotFoundError(f"Folder not found: {folder_path}")

        files_info = []
        aggregated_text = []
        images_to_process = []
        found_gps = None

        for filename in sorted(os.listdir(folder_path)):
            if filename.startswith("~$") or filename.startswith("."):
                continue
            
            file_path = os.path.join(folder_path, filename)
            if not os.path.isfile(file_path):
                continue

            f_res = self.inspect_file(file_path)
            files_info.append({
                "filename": f_res["filename"],
                "path": f_res["path"],
                "type": f_res["type"],
                "classification_label": f_res["classification_label"],
                "priority": f_res["priority"],
                "size_bytes": f_res["size_bytes"],
                "size_display": f_res["size_display"]
            })

            if f_res["text"]:
                aggregated_text.append(f"=== DOCUMENT [{f_res['classification_label']}]: {filename} ===\n{f_res['text']}")
            if f_res["images"]:
                images_to_process.extend(f_res["images"][:2])
            if f_res["gps"] and not found_gps:
                found_gps = f_res["gps"]

        return {
            "files": files_info,
            "all_text": "\n\n".join(aggregated_text),
            "images": images_to_process,
            "gps_from_exif": found_gps
        }

    def dynamic_entity_extractor(self, all_text: str, gps_exif: Optional[str] = None) -> ReportData:
        """
        Document Understanding & Semantic Extraction Engine (100% Offline & Pure Regex/NLP).
        Extracts exact field values from raw document text.
        CRITICAL RULE: If no match is found, keeps the field completely empty ("" or 0.0).
        NO HARDCODED FALLBACKS.
        """
        report = ReportData()
        if not all_text or not all_text.strip():
            if gps_exif:
                report.header.geo_tag = gps_exif
            return report

        clean_text = " ".join(all_text.split())

        # 1. Application ID (e.g. AP-10524478, DEL-29182, Loan No: 129381)
        app_id_m = re.search(r'(?:Application\s*ID|App\s*No\.?|Loan\s*No\.?|Ref\s*No\.?|Case\s*No\.?)\s*[:\-]?\s*([A-Za-z0-9\-_/]{4,25})', all_text, re.IGNORECASE)
        if app_id_m:
            report.header.application_id = app_id_m.group(1).strip()
        else:
            id_fallback = re.search(r'\b(AP[-_]\d{5,12})\b', all_text, re.IGNORECASE)
            if id_fallback:
                report.header.application_id = id_fallback.group(1).strip()

        # 2. Applicant Name
        name_m = re.search(r'(?:Applicant\s*Name|Borrower\s*Name|Customer\s*Name)\s*[:\-]?\s*([A-Za-z\.\s]{3,40})', all_text, re.IGNORECASE)
        if name_m:
            candidate = name_m.group(1).strip()
            candidate = re.split(r'(\n|\r|Type|Address|Age|Plot|and|Son|Wife|Phone)', candidate, flags=re.IGNORECASE)[0].strip()
            if len(candidate) >= 3 and not candidate.lower().startswith("type"):
                report.header.applicant_name = candidate
        else:
            purchaser_m = re.search(r'(?:Purchaser|Second\s*Party)\s*[:\-]?\s*(?:Mrs?\.?|Smt\.?|Sh\.?|Shri)?\s*([A-Za-z\s]{3,35})', all_text, re.IGNORECASE)
            if purchaser_m:
                cand = purchaser_m.group(1).strip()
                cand = re.split(r'(\n|\r|W/o|S/o|D/o|and|for)', cand, flags=re.IGNORECASE)[0].strip()
                if len(cand) >= 3:
                    report.header.applicant_name = cand

        # 3. Geo Tag / GPS Coordinates
        geo_m = re.search(r'(?:Geo[\s\-]*Tag|Coordinates?|Geo-ordinates?)\s*(?:of\s*Subject\s*Property\s*are)?\s*[:\-]?\s*(\d{1,2}\.\d{4,8})\s*,\s*(\d{1,3}\.\d{4,8})', all_text, re.IGNORECASE)
        if geo_m:
            report.header.geo_tag = f"{geo_m.group(1)}, {geo_m.group(2)}"
        elif gps_exif:
            report.header.geo_tag = gps_exif
        else:
            raw_coords = re.search(r'\b(\d{2}\.\d{4,8})\s*,\s*(\d{2}\.\d{4,8})\b', all_text)
            if raw_coords:
                report.header.geo_tag = f"{raw_coords.group(1)}, {raw_coords.group(2)}"

        # 4. Property Type
        if re.search(r'\b(Row\s*House)\b', all_text, re.IGNORECASE):
            report.header.property_type = "Row House"
        elif re.search(r'\b(Independent\s*Floor)\b', all_text, re.IGNORECASE):
            report.header.property_type = "Independent Floor"
        elif re.search(r'\b(Apartment|Flat)\b', all_text, re.IGNORECASE):
            report.header.property_type = "Apartment"
        elif re.search(r'\b(Commercial\s*Building|Commercial\s*Property)\b', all_text, re.IGNORECASE):
            report.header.property_type = "Commercial Building"

        # Percentage of completion
        comp_m = re.search(r'(?:Percentage\s*of\s*Completion|Completion\s*Percent(?:age)?)\s*[:\-]?\s*(\d{1,3})\s*%', all_text, re.IGNORECASE)
        if comp_m:
            report.header.completion_percent = round(float(comp_m.group(1)) / 100.0, 2)
        elif "100%" in all_text or "fully constructed" in all_text.lower():
            report.header.completion_percent = 1.0

        # Age of Property
        age_m = re.search(r'(\d{1,2}\s*(?:Years?|Yrs?))\s*(?:old|age)?', all_text, re.IGNORECASE)
        if age_m:
            num = re.search(r'\d+', age_m.group(1)).group(0)
            report.header.age_of_property = f"{int(num):02d} Years"

        # Structure Type
        if "load bearing" in all_text.lower():
            report.header.structure_type = "Load Bearing"
        elif "rcc" in all_text.lower():
            report.header.structure_type = "RCC"

        # Dwelling Units
        units_m = re.search(r'(?:Dwelling\s*Units?\s*Owned)\s*[:\-]?\s*(\d+)', all_text, re.IGNORECASE)
        if units_m:
            report.header.dwelling_units_owned = int(units_m.group(1))

        # 5. Property / Plot / Khasra Identification
        plot_m = re.search(r'(?:Plot\s*No\.?|Property\s*(?:Bearing\s*Plot\s*)?No\.?)\s*([A-Za-z0-9\-_/]+)', all_text, re.IGNORECASE)
        if plot_m:
            plot_val = plot_m.group(1).strip()
            report.address.plot_house_khasra = f"Property No. {plot_val}" if not plot_val.lower().startswith("property") else plot_val

        # 6. Dimensions (Length & Breadth)
        dim_all = re.findall(r'(?:dimension|size|area)?[\s\:]*\(?\b(\d{1,3}(?:\.\d{1,2})?)\s*(?:x|X|\*|by)\s*(\d{1,3}(?:\.\d{1,2})?)\b\)?', all_text, re.IGNORECASE)
        valid_dims = []
        for d1_s, d2_s in dim_all:
            try:
                v1, v2 = float(d1_s), float(d2_s)
                if 5.0 <= v1 <= 200.0 and 5.0 <= v2 <= 200.0:
                    valid_dims.append((max(v1, v2), min(v1, v2)))
            except Exception:
                pass

        if valid_dims:
            from collections import Counter
            common_dim = Counter(valid_dims).most_common(1)[0][0]
            length, breadth = common_dim
            report.land_measurements.land_length = length
            report.land_measurements.land_breadth = breadth
            report.solar_roof_vicinity.roof_length_sqft = breadth
            report.solar_roof_vicinity.roof_breadth_sqft = length

        # 7. Land Area (Sq Yds / Sq Ft)
        area_sqyds_m = re.search(r'(\d+(?:\.\d+)?)\s*(?:Sq\.?\s*Yds?|Sq\.?\s*Yards?|Sq\s*Yrd)', all_text, re.IGNORECASE)
        area_sqft_m = re.search(r'(\d+(?:\.\d+)?)\s*(?:Sq\.?\s*Ft|Sqft|Square\s*Feet)', all_text, re.IGNORECASE)
        
        if area_sqft_m:
            sqft_val = float(area_sqft_m.group(1))
            report.land_measurements.land_area_site_sqft = f"{sqft_val} Sqft"
            report.land_measurements.adopted_land_area_sqft = sqft_val
        elif area_sqyds_m:
            sqyds_val = float(area_sqyds_m.group(1))
            sqft_calc = round(sqyds_val * 9.0, 1)
            report.land_measurements.land_area_site_sqft = f"{sqft_calc} Sqft"
            report.land_measurements.adopted_land_area_sqft = sqft_calc
        elif report.land_measurements.land_length > 0 and report.land_measurements.land_breadth > 0:
            calc_area = round(report.land_measurements.land_length * report.land_measurements.land_breadth, 1)
            report.land_measurements.land_area_site_sqft = f"{calc_area} Sqft"
            report.land_measurements.adopted_land_area_sqft = calc_area

        # 8. Pincode & City
        pin_m = re.search(r'(?:Delhi|Pincode|Pin|PIN\s*Code)[\s\:\-]*([1-9]\d{5})\b', all_text, re.IGNORECASE)
        if not pin_m and report.address.address_docs:
            pin_m = re.search(r'\b([1-9]\d{5})\b', report.address.address_docs)
        if not pin_m and report.address.address_site:
            pin_m = re.search(r'\b([1-9]\d{5})\b', report.address.address_site)
        if not pin_m:
            pin_m = re.search(r'\b([1-9]\d{5})\b', all_text)
        if pin_m:
            report.address.pincode = pin_m.group(1)

        city_m = re.search(r'(?:City|District)\s*[:\-]?\s*([A-Za-z\s]{3,20})', all_text, re.IGNORECASE)
        if city_m:
            c_val = city_m.group(1).strip()
            if not c_val.lower().startswith("pincode") and not c_val.lower().startswith("depart"):
                report.address.city = c_val
                report.address.district = c_val
        if not report.address.city or report.address.city.lower() in ["departnment", "department"]:
            if "new delhi" in all_text.lower():
                report.address.city = "New Delhi"
                report.address.district = "New Delhi"
            elif "delhi" in all_text.lower():
                report.address.city = "Delhi"
                report.address.district = "Delhi"

        # 9. Street / Landmark / Village / Colony
        street_m = re.search(r'(Gali\s*No\.?\s*[0-9A-Za-z\-_]+|Road\s*No\.?\s*[0-9A-Za-z\-_]+|Street\s*[0-9A-Za-z\-_]+)', all_text, re.IGNORECASE)
        if street_m:
            report.address.street_name = street_m.group(1).strip()

        landmark_m = re.search(r'(?:Nearest\s*Landmark|Landmark)\s*[:\-]?\s*([A-Za-z0-9\.\s]{3,35})', all_text, re.IGNORECASE)
        if landmark_m:
            report.address.nearest_landmark = landmark_m.group(1).strip()

        colony_m = re.search(r'(?:Abadi\s*Known\s*as|Colony|Enclave|Garden|Vihar|Nagar)\s*[:\-]?\s*([A-Za-z\s]+(?:Extn\.?|Extension|Vihar|Nagar|Garden|Enclave))', all_text, re.IGNORECASE)
        if colony_m:
            report.address.colony_name = colony_m.group(1).strip()

        village_m = re.search(r'(?:Revenue\s*Estate\s*of\s*Village[\- ]*|Village[\- ]+)([A-Za-z]+)', all_text, re.IGNORECASE)
        if village_m:
            report.address.village_name = village_m.group(1).strip()

        # 10. Document Address & Site Address
        doc_match = re.search(r'(Property\s*(?:Bearing\s*)?Plot\s*No\.[^\n\r]{10,180}?(?:110\d{3}|\d{6}))', clean_text, re.IGNORECASE)
        if doc_match:
            report.address.address_docs = doc_match.group(1).strip()

        site_match = re.search(r'(Property\s*No\.\s*[0-9A-Za-z\-_]+,\s*Situated\s*in[^\n\r]{10,180}?(?:110\d{3}|\d{6}))', clean_text, re.IGNORECASE)
        if site_match:
            report.address.address_site = site_match.group(1).strip()

        # 11. Boundaries (Site vs Deed)
        # Site Boundaries
        site_east_m = re.search(r'(?:Site\s*East|East\s*as\s*per\s*site)\s*[:\-]?\s*([^\n\r,]{3,40})', all_text, re.IGNORECASE)
        if site_east_m:
            report.boundaries.site_east = site_east_m.group(1).strip()

        site_west_m = re.search(r'(?:Site\s*West|West\s*as\s*per\s*site)\s*[:\-]?\s*([^\n\r,]{3,40})', all_text, re.IGNORECASE)
        if site_west_m:
            report.boundaries.site_west = site_west_m.group(1).strip()

        site_north_m = re.search(r'(?:Site\s*North|North\s*as\s*per\s*site)\s*[:\-]?\s*([^\n\r,]{3,40})', all_text, re.IGNORECASE)
        if site_north_m:
            report.boundaries.site_north = site_north_m.group(1).strip()

        site_south_m = re.search(r'(?:Site\s*South|South\s*as\s*per\s*site)\s*[:\-]?\s*([^\n\r,]{3,40})', all_text, re.IGNORECASE)
        if site_south_m:
            report.boundaries.site_south = site_south_m.group(1).strip()

        # Deed Boundaries
        deed_east_m = re.search(r'(?:Deed\s*East|East\s*as\s*per\s*deed)\s*[:\-]?\s*([^\n\r,]{3,40})', all_text, re.IGNORECASE)
        if deed_east_m:
            report.boundaries.deed_east = deed_east_m.group(1).strip()

        deed_west_m = re.search(r'(?:Deed\s*West|West\s*as\s*per\s*deed)\s*[:\-]?\s*([^\n\r,]{3,40})', all_text, re.IGNORECASE)
        if deed_west_m:
            report.boundaries.deed_west = deed_west_m.group(1).strip()

        deed_north_m = re.search(r'(?:Deed\s*North|North\s*as\s*per\s*deed)\s*[:\-]?\s*([^\n\r,]{3,40})', all_text, re.IGNORECASE)
        if deed_north_m:
            report.boundaries.deed_north = deed_north_m.group(1).strip()

        deed_south_m = re.search(r'(?:Deed\s*South|South\s*as\s*per\s*deed)\s*[:\-]?\s*([^\n\r,]{3,40})', all_text, re.IGNORECASE)
        if deed_south_m:
            report.boundaries.deed_south = deed_south_m.group(1).strip()

        # Boundary Matching & Occupancy
        if "boundary matching" in all_text.lower():
            bm_m = re.search(r'Boundary\s*Matching\s*[:\-]?\s*(Yes|No)', all_text, re.IGNORECASE)
            if bm_m:
                report.boundaries.boundary_matching = bm_m.group(1).capitalize()
        if re.search(r'\b(Seller|Borrower|Tenant|Vacant)\b', all_text, re.IGNORECASE):
            occ_m = re.search(r'(?:Occupancy\s*Status)\s*[:\-]?\s*([A-Za-z]+)', all_text, re.IGNORECASE)
            if occ_m:
                report.boundaries.occupancy_status = occ_m.group(1).capitalize()

        # 12. Floors & Accommodation
        if re.search(r'(?:S\+UG\+3|S\s*\+\s*UG\s*\+\s*3|5\s*storied|5\s*floor)', all_text, re.IGNORECASE):
            report.accommodation.no_of_floors = 5
        elif re.search(r'(\d+)\s*(?:storied|floors?)', all_text, re.IGNORECASE):
            fl_m = re.search(r'(\d+)\s*(?:storied|floors?)', all_text, re.IGNORECASE)
            report.accommodation.no_of_floors = int(fl_m.group(1))

        # 13. Legal Checks & Person Met
        person_m = re.search(r'(?:Person\s*Meet|Met\s*at\s*site|Contact\s*Person)\s*[:\-]?\s*(?:Mr\.?|Mrs\.?|Sh\.?)?\s*([A-Za-z\s]{3,25})', all_text, re.IGNORECASE)
        if person_m:
            cand = person_m.group(1).strip()
            cand = re.split(r'(\n|\r|Relation|Applicant|Son|Phone)', cand, flags=re.IGNORECASE)[0].strip()
            if len(cand) >= 3 and not cand.lower().startswith("relation"):
                report.legal_checks.person_met = f"Mr. {cand}" if not cand.lower().startswith("mr") else cand

        rel_m = re.search(r'(?:Reation|Relation)\s*(?:with\s*(?:the\s*)?(?:Property\s*)?Owner)?\s*[:\-]?\s*([A-Za-z\'\s]{3,25})', all_text, re.IGNORECASE)
        if rel_m:
            cand = rel_m.group(1).strip()
            cand = re.split(r'(\n|\r|Property|MC|Phone)', cand, flags=re.IGNORECASE)[0].strip()
            if len(cand) >= 3:
                report.legal_checks.relation_with_owner = cand

        if "road" in all_text.lower():
            road_w_m = re.search(r'(\d+\s*(?:Ft|Feet|Meter|Mtr)\s*Wide)', all_text, re.IGNORECASE)
            if road_w_m:
                report.legal_checks.width_of_public_road = road_w_m.group(1).strip()

        # 14. Reference & Feedback
        phone_m = re.search(r'\b([6-9]\d{9})\b', all_text)
        if phone_m:
            report.reference.reference_mobile = phone_m.group(1)

        fb_m = re.search(r'(?:Feedback)\s*[:\-]?\s*([0-9\sA-Za-z\.\,\-toperSqyds]{3,40})(?:\n|\r|$)', all_text, re.IGNORECASE)
        if fb_m and not fb_m.group(1).strip().startswith("1.") and "subject property" not in fb_m.group(1).lower():
            report.reference.feedback = fb_m.group(1).strip()

        # 15. Narrative Remarks
        remarks_block_m = re.search(r'(1\.\s*Subject Property[\s\S]*?(?:\d+\.\s*[^\n\r]+(?:\n|\r|$))+)', all_text)
        if remarks_block_m:
            report.remarks = remarks_block_m.group(1).strip() + "\n"
        else:
            numbered_m = re.findall(r'(\d+\.\s*[^\n\r]+)', all_text)
            if len(numbered_m) >= 3:
                report.remarks = "\n".join(numbered_m) + "\n"

        return report

    def extract_from_single_file(self, file_path: str) -> ReportData:
        """Processes a single file and extracts entities."""
        f_res = self.inspect_file(file_path)
        text = f_res["text"]
        gps = f_res["gps"]

        report = self.dynamic_entity_extractor(text, gps)
        report = ValuationValidator.validate_and_score(report)
        report.raw_files_summary = [{
            "filename": f_res["filename"],
            "path": f_res["path"],
            "type": f_res["type"],
            "classification_label": f_res["classification_label"],
            "priority": f_res["priority"],
            "size_bytes": f_res["size_bytes"],
            "size_display": f_res["size_display"]
        }]
        return report

    def process_file_list(self, file_paths: List[str], progress_callback: Optional[Any] = None) -> ReportData:
        """Processes a list of document file paths, aggregates texts, runs extraction & validation."""
        files_info = []
        aggregated_text = []
        found_gps = None

        valid_paths = [
            p for p in file_paths
            if os.path.exists(p) and os.path.isfile(p) and not os.path.basename(p).startswith("~$") and not os.path.basename(p).startswith(".") and not p.lower().endswith(".xlsx")
        ]
        total_files = len(valid_paths)

        for idx, fpath in enumerate(valid_paths):
            fname = os.path.basename(fpath)
            if progress_callback:
                pct = 20 + int(((idx + 0.5) / max(total_files, 1)) * 60)
                try:
                    progress_callback(
                        step=2,
                        percent=min(pct, 80),
                        message=f"OCR & Vision analysis ({idx + 1}/{total_files}): {fname}",
                        current_file=fname,
                        processed=idx + 1,
                        total=total_files
                    )
                except Exception:
                    pass

            try:
                f_res = self.inspect_file(fpath)
                files_info.append({
                    "filename": f_res["filename"],
                    "path": f_res["path"],
                    "type": f_res["type"],
                    "classification_label": f_res["classification_label"],
                    "priority": f_res["priority"],
                    "size_bytes": f_res["size_bytes"],
                    "size_display": f_res["size_display"]
                })
                if f_res.get("text"):
                    aggregated_text.append(f"=== DOCUMENT [{f_res['classification_label']}]: {fname} ===\n{f_res['text']}")
                if f_res.get("gps") and not found_gps:
                    found_gps = f_res["gps"]
            except Exception as e:
                print(f"[Warning] Error inspecting file {fpath}: {e}")

        if progress_callback:
            try:
                progress_callback(
                    step=3,
                    percent=85,
                    message="Matching cross-document entities & resolving property details...",
                    current_file="",
                    processed=total_files,
                    total=total_files
                )
            except Exception:
                pass

        all_text = "\n\n".join(aggregated_text)
        report = self.dynamic_entity_extractor(all_text, found_gps)

        if progress_callback:
            try:
                progress_callback(
                    step=3,
                    percent=92,
                    message="Validating compliance, boundaries & calculating field confidences...",
                    current_file="",
                    processed=total_files,
                    total=total_files
                )
            except Exception:
                pass

        report = ValuationValidator.validate_and_score(report)
        report.raw_files_summary = files_info
        return report

    def process_case_folder(self, folder_path: str, progress_callback: Optional[Any] = None) -> ReportData:
        """Full offline pipeline: Ingest files -> Classify -> Local OCR/Parsing -> Semantic Extractor -> Validation."""
        file_paths = []
        for root, _, files in os.walk(folder_path):
            for f in sorted(files):
                if not f.startswith("~$") and not f.startswith(".") and not f.lower().endswith(".xlsx"):
                    file_paths.append(os.path.join(root, f))

        report = self.process_file_list(file_paths, progress_callback=progress_callback)
        report.case_name = os.path.basename(folder_path.rstrip("/\\"))
        return report

    def ingest_incremental_files(self, existing_report: ReportData, new_file_paths: List[str]) -> ReportData:
        """Incrementally parses new files and updates report data."""
        files_info = list(existing_report.raw_files_summary)
        aggregated_text = []
        found_gps = existing_report.header.geo_tag

        for fpath in new_file_paths:
            f_res = self.inspect_file(fpath)
            fname = os.path.basename(fpath)
            files_info.append({
                "filename": f_res["filename"],
                "path": f_res["path"],
                "type": f_res["type"],
                "classification_label": f_res["classification_label"],
                "priority": f_res["priority"],
                "size_bytes": f_res["size_bytes"],
                "size_display": f_res["size_display"]
            })

            if f_res["text"]:
                aggregated_text.append(f"=== DOCUMENT [{f_res['classification_label']}]: {fname} ===\n{f_res['text']}")
            if f_res["gps"] and not found_gps:
                found_gps = f_res["gps"]

        all_text = "\n\n".join(aggregated_text)
        report = self.dynamic_entity_extractor(all_text, found_gps)
        report = ValuationValidator.validate_and_score(report)
        report.raw_files_summary = files_info
        return report
