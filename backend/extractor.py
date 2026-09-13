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
    Uses memory-efficient 110 DPI rendering to prevent Render Free Tier OOM crashes.
    """
    # Check for browser client pre-extracted text first
    sidecar_client = filepath + ".client.txt"
    if os.path.exists(sidecar_client):
        try:
            with open(sidecar_client, "r", encoding="utf-8", errors="ignore") as sc:
                client_text = sc.read().strip()
            if client_text:
                return client_text, []
        except Exception:
            pass

    text_content = []
    page_images = []
    try:
        doc = fitz.open(filepath)
        for page_num in range(min(len(doc), max_pages)):
            page = doc[page_num]
            txt = page.get_text()

            # If digital text exists, use it
            if txt.strip():
                # Extract structured tabular key-values if present in digital PDF
                try:
                    tabs = page.find_tables()
                    for tab in getattr(tabs, 'tables', []):
                        rows = tab.extract()
                        prev_is_boundary = False
                        boundary_count = 0
                        for r in rows:
                            clean = [str(c).strip().replace('\n', ' ') if c else '' for c in r]
                            if len(clean) >= 4 and [c.lower() for c in clean[:4]] == ['east', 'west', 'north', 'south']:
                                prev_is_boundary = True
                                boundary_count += 1
                                continue
                            if prev_is_boundary and len(clean) >= 4:
                                prefix = 'Site' if boundary_count == 1 else 'Deed'
                                text_content.append(f"{prefix} East: {clean[0]}")
                                text_content.append(f"{prefix} West: {clean[1]}")
                                text_content.append(f"{prefix} North: {clean[2]}")
                                text_content.append(f"{prefix} South: {clean[3]}")
                                prev_is_boundary = False
                                continue
                            prev_is_boundary = False
                            if len(clean) == 4 and clean[0] and clean[1] and clean[2] and clean[3]:
                                text_content.append(f"{clean[0]}: {clean[1]}")
                                text_content.append(f"{clean[2]}: {clean[3]}")
                            elif len(clean) >= 2 and clean[0] and clean[1]:
                                text_content.append(f"{clean[0]}: {clean[1]}")
                except Exception:
                    pass
                text_content.append(f"--- [Page {page_num + 1}] ---\n{txt}")
            else:
                # Scanned page: render pixmap at memory-efficient 110 DPI
                pix = page.get_pixmap(dpi=110)
                img_bytes = pix.tobytes("jpeg")
                ocr_text, _ = OfflineOCREngine.extract_text_from_image_bytes(img_bytes)
                if ocr_text.strip():
                    text_content.append(f"--- [Page {page_num + 1} (OCR)] ---\n{ocr_text}")
                if len(page_images) < 2:
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

        # Check for browser client pre-extracted text
        sidecar_client = file_path + ".client.txt"
        if os.path.exists(sidecar_client):
            try:
                with open(sidecar_client, "r", encoding="utf-8", errors="ignore") as sc:
                    client_t = sc.read().strip()
                if client_t:
                    text = f"[{classification.get('label', 'Document')} - Client Accelerated]:\n{client_t}"
            except Exception:
                pass

        # If text was not provided by client, run standard local extraction
        if not text:
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

        # 9. Dwelling Units & Floor Number
        dw_m = re.search(r'Dwelling\s*Units\s*Owned\s*[:\-]?\s*(\d+)', all_text, re.IGNORECASE)
        if dw_m:
            report.header.dwelling_units_owned = int(dw_m.group(1))

        fl_no_m = re.search(r'Floor\s*Number\s*[:\-]?\s*([A-Za-z0-9\s]{3,30})', all_text, re.IGNORECASE)
        if fl_no_m:
            cand = fl_no_m.group(1).strip()
            if not cand.lower().startswith("project") and not cand.lower().startswith("colony") and not cand.lower().startswith("address"):
                report.address.floor_number = cand

        plot_m = re.search(r'Plot/House/Flat\s*Number/Khasra\s*No\.?\s*[:\-]?\s*([A-Za-z0-9\s\.\-_]{3,35})', all_text, re.IGNORECASE)
        if plot_m:
            cand = plot_m.group(1).strip()
            if not cand.lower().startswith("floor"):
                report.address.plot_house_khasra = cand

        # 10. Street / Landmark / Village / Colony
        street_m = re.search(r'(?:Street\s*Name/Number\s*[:\-]?\s*([A-Za-z0-9\s\.\-_]{3,35})|Gali\s*No\.?\s*[0-9A-Za-z\-_]+|Road\s*No\.?\s*[0-9A-Za-z\-_]+)', all_text, re.IGNORECASE)
        if street_m:
            val = street_m.group(1) or street_m.group(0)
            if not val.lower().startswith("nearest"):
                val = val.strip()
                if val.isupper():
                    val = val.title()
                report.address.street_name = val

        landmark_m = re.search(r'(?:Nearest\s*Landmark|Landmark)\s*[:\-]?\s*([A-Za-z0-9\.\s]{3,35})', all_text, re.IGNORECASE)
        if landmark_m:
            cand = landmark_m.group(1).strip()
            if not cand.lower().startswith("village") and not cand.lower().startswith("city"):
                report.address.nearest_landmark = cand

        colony_m = re.search(r'(?:Project/Society/Colony\s*Name|Abadi\s*Known\s*as)\s*[:\-]?\s*([A-Za-z\s\.\-_]+?)(?=\n|Address|$)', all_text, re.IGNORECASE)
        if not colony_m:
            colony_m = re.search(r'(?:Colony|Enclave|Garden|Vihar|Nagar)\s*[:\-]?\s*([A-Za-z\s]+(?:Extn\.?|Extension|Vihar|Nagar|Garden|Enclave))', all_text, re.IGNORECASE)
        if colony_m:
            cand = colony_m.group(1).strip()
            if not cand.lower().startswith("address"):
                report.address.colony_name = cand

        village_m = re.search(r'(?:Village\s*Name|Revenue\s*Estate\s*of\s*Village[\- ]*|Village[\- ]+)\s*[:\-]?\s*([A-Za-z]+)', all_text, re.IGNORECASE)
        if village_m:
            cand = village_m.group(1).strip()
            if not cand.lower().startswith("city") and not cand.lower().startswith("plot"):
                report.address.village_name = cand

        # 11. Document Address & Site Address
        doc_match = re.search(r'(Property\s*(?:Bearing\s*)?Plot\s*No\.[^\n\r]{10,180}?(?:110\d{3}|\d{6}))', clean_text, re.IGNORECASE)
        if doc_match:
            report.address.address_docs = doc_match.group(1).strip()

        site_match = re.search(r'(Property\s*No\.\s*[0-9A-Za-z\-_]+,\s*Situated\s*in[^\n\r]{10,180}?(?:110\d{3}|\d{6}))', clean_text, re.IGNORECASE)
        if site_match:
            report.address.address_site = site_match.group(1).strip()

        # 11. Boundaries (Site vs Deed)
        # Check boundary block pattern first (e.g. East\nWest\nNorth\nSouth...)
        bound_block_m = re.search(r'East\s*\n\s*West\s*\n\s*North\s*\n\s*South\s*\n\s*([^\n\r]+)\s*\n\s*([^\n\r]+(?:\n\s*[0-9A-Za-z]+)?)\s*\n\s*([^\n\r]+)\s*\n\s*([^\n\r]+)\s*\n\s*East\s*\n\s*West\s*\n\s*North\s*\n\s*South\s*\n\s*([^\n\r]+)\s*\n\s*([^\n\r]+)\s*\n\s*([^\n\r]+)\s*\n\s*([^\n\r]+)', all_text, re.IGNORECASE)
        if bound_block_m:
            report.boundaries.site_east = bound_block_m.group(1).strip()
            report.boundaries.site_west = bound_block_m.group(2).replace('\n', ' ').strip()
            report.boundaries.site_north = bound_block_m.group(3).strip()
            report.boundaries.site_south = bound_block_m.group(4).strip()
            report.boundaries.deed_east = bound_block_m.group(5).strip()
            report.boundaries.deed_west = bound_block_m.group(6).strip()
            report.boundaries.deed_north = bound_block_m.group(7).strip()
            report.boundaries.deed_south = bound_block_m.group(8).strip()

        # Site Boundaries key-value fallback
        if not report.boundaries.site_east:
            site_east_m = re.search(r'(?:Site\s*East|East\s*as\s*per\s*site)\s*[:\-]?\s*([^\n\r,]{3,50})', all_text, re.IGNORECASE)
            if site_east_m:
                report.boundaries.site_east = site_east_m.group(1).strip()

        if not report.boundaries.site_west:
            site_west_m = re.search(r'(?:Site\s*West|West\s*as\s*per\s*site)\s*[:\-]?\s*([^\n\r,]{3,50})', all_text, re.IGNORECASE)
            if site_west_m:
                report.boundaries.site_west = site_west_m.group(1).strip()

        if not report.boundaries.site_north:
            site_north_m = re.search(r'(?:Site\s*North|North\s*as\s*per\s*site)\s*[:\-]?\s*([^\n\r,]{3,50})', all_text, re.IGNORECASE)
            if site_north_m:
                report.boundaries.site_north = site_north_m.group(1).strip()

        if not report.boundaries.site_south:
            site_south_m = re.search(r'(?:Site\s*South|South\s*as\s*per\s*site)\s*[:\-]?\s*([^\n\r,]{3,50})', all_text, re.IGNORECASE)
            if site_south_m:
                report.boundaries.site_south = site_south_m.group(1).strip()

        # Deed Boundaries key-value fallback
        if not report.boundaries.deed_east:
            deed_east_m = re.search(r'(?:Deed\s*East|East\s*as\s*per\s*(?:title\s*)?deed)\s*[:\-]?\s*([^\n\r,]{3,50})', all_text, re.IGNORECASE)
            if deed_east_m:
                report.boundaries.deed_east = deed_east_m.group(1).strip()

        if not report.boundaries.deed_west:
            deed_west_m = re.search(r'(?:Deed\s*West|West\s*as\s*per\s*(?:title\s*)?deed)\s*[:\-]?\s*([^\n\r,]{3,50})', all_text, re.IGNORECASE)
            if deed_west_m:
                report.boundaries.deed_west = deed_west_m.group(1).strip()

        if not report.boundaries.deed_north:
            deed_north_m = re.search(r'(?:Deed\s*North|North\s*as\s*per\s*(?:title\s*)?deed)\s*[:\-]?\s*([^\n\r,]{3,50})', all_text, re.IGNORECASE)
            if deed_north_m:
                report.boundaries.deed_north = deed_north_m.group(1).strip()

        if not report.boundaries.deed_south:
            deed_south_m = re.search(r'(?:Deed\s*South|South\s*as\s*per\s*(?:title\s*)?deed)\s*[:\-]?\s*([^\n\r,]{3,50})', all_text, re.IGNORECASE)
            if deed_south_m:
                report.boundaries.deed_south = deed_south_m.group(1).strip()

        # Boundary Matching & Occupancy
        if "boundary matching" in all_text.lower():
            bm_m = re.search(r'Boundary\s*Matching\s*[:\-]?\s*(Yes|No)', all_text, re.IGNORECASE)
            if bm_m:
                report.boundaries.boundary_matching = bm_m.group(1).capitalize()
        if "mismatch remarks" in all_text.lower():
            mm_m = re.search(r'Mismatch\s*Remarks\s*[:\-]?\s*([A-Za-z0-9\s]+)', all_text, re.IGNORECASE)
            if mm_m:
                report.boundaries.mismatch_remarks = mm_m.group(1).strip()

        occ_m = re.search(r'(?:Occupancy\s*Status)\s*[:\-]?\s*([A-Za-z]+)', all_text, re.IGNORECASE)
        if occ_m:
            cand = occ_m.group(1).strip().capitalize()
            if cand in ["Seller", "Borrower", "Tenant", "Vacant"]:
                report.boundaries.occupancy_status = cand
        elif re.search(r'\b(seller[\- ]occupied|occupied\s*by\s*seller)\b', all_text, re.IGNORECASE):
            report.boundaries.occupancy_status = "Seller"
        elif re.search(r'\b(borrower[\- ]occupied|occupied\s*by\s*borrower)\b', all_text, re.IGNORECASE):
            report.boundaries.occupancy_status = "Borrower"
        elif re.search(r'\b(tenant[\- ]occupied|occupied\s*by\s*tenant)\b', all_text, re.IGNORECASE):
            report.boundaries.occupancy_status = "Tenant"
        elif re.search(r'\b(vacant\s*property|property\s*is\s*vacant)\b', all_text, re.IGNORECASE):
            report.boundaries.occupancy_status = "Vacant"

        # 12. Solar & Roof Vicinity
        sol_loc_m = re.search(r'Solar\s*Panel\s*Can\s*be\s*installed\s*at\s*[:\-]?\s*([A-Za-z\s]{3,20})', all_text, re.IGNORECASE)
        if sol_loc_m:
            report.solar_roof_vicinity.solar_install_location = sol_loc_m.group(1).strip()

        roof_len_m = re.search(r'Length\s*in\s*Sq\s*ft\s*[:\-]?\s*(\d+(?:\.\d+)?)', all_text, re.IGNORECASE)
        if roof_len_m:
            report.solar_roof_vicinity.roof_length_sqft = float(roof_len_m.group(1))

        roof_br_m = re.search(r'Breadth\s*in\s*Sq\s*ft\s*[:\-]?\s*(\d+(?:\.\d+)?)', all_text, re.IGNORECASE)
        if roof_br_m:
            report.solar_roof_vicinity.roof_breadth_sqft = float(roof_br_m.group(1))

        outreach_m = re.search(r'is\s*Property\s*Situated\s*in\s*Outreach\s*[:\-]?\s*(Yes|No)', all_text, re.IGNORECASE)
        if outreach_m:
            report.solar_roof_vicinity.is_outreach = outreach_m.group(1).capitalize()

        pop_m = re.search(r'Population\s*within\s*1\s*KM\s*Radius\s*[:\-]?\s*([A-Za-z0-9\s]{3,30})', all_text, re.IGNORECASE)
        if pop_m:
            cand = pop_m.group(1).strip()
            if not cand.lower().startswith("no"):
                report.solar_roof_vicinity.population_1km = cand

        pri_m = re.search(r'No\.?\s*of\s*Primary\s*Schools[^\n\r]*?[:\-]?\s*(\d+)', all_text, re.IGNORECASE)
        if pri_m:
            report.solar_roof_vicinity.primary_schools_1km = int(pri_m.group(1))

        sec_m = re.search(r'No\.?\s*of\s*Secondary\s*Schools[^\n\r]*?[:\-]?\s*(\d+)', all_text, re.IGNORECASE)
        if sec_m:
            report.solar_roof_vicinity.secondary_schools_1km = int(sec_m.group(1))

        govt_m = re.search(r'Govt\s*Institution\s*in\s*Vicinity\s*[:\-]?\s*(\d+)', all_text, re.IGNORECASE)
        if govt_m:
            report.solar_roof_vicinity.govt_institutions_vicinity = int(govt_m.group(1))

        # 13. Floors & Accommodation
        if re.search(r'(?:S\+UG\+3|S\s*\+\s*UG\s*\+\s*3|5\s*storied|5\s*floor)', all_text, re.IGNORECASE):
            report.accommodation.no_of_floors = 5
        elif re.search(r'(\d+)\s*(?:storied|floors?)', all_text, re.IGNORECASE):
            fl_m = re.search(r'(\d+)\s*(?:storied|floors?)', all_text, re.IGNORECASE)
            report.accommodation.no_of_floors = int(fl_m.group(1))

        toilet_m = re.search(r'Toilet\s*Available\s*[:\-]?\s*(Yes|No)', all_text, re.IGNORECASE)
        if toilet_m:
            report.accommodation.toilet_available = toilet_m.group(1).capitalize()

        lift_m = re.search(r'No\.?\s*of\s*Lifts?\s*[:\-]?\s*(\d+)', all_text, re.IGNORECASE)
        if lift_m:
            report.accommodation.no_of_lifts = int(lift_m.group(1))

        apt_m = re.search(r'Number\s*of\s*Appartments?\s*Per\s*Floor\s*[:\-]?\s*(\d+)', all_text, re.IGNORECASE)
        if apt_m:
            report.accommodation.apartments_per_floor = int(apt_m.group(1))

        em_inst_m = re.search(r'Electricity\s*Meter\s*Installed\s*[:\-]?\s*(Yes|No)', all_text, re.IGNORECASE)
        if em_inst_m:
            report.accommodation.electricity_meter_installed = em_inst_m.group(1).capitalize()

        em_num_m = re.search(r'Electricity\s*Meter\s*Number\s*[:\-]?\s*([A-Za-z0-9\-_/]+)', all_text, re.IGNORECASE)
        if em_num_m:
            report.accommodation.electricity_meter_number = em_num_m.group(1).strip()

        # 14. Legal Statutory Checks & Person Met
        doc_name_m = re.search(r'Documents\s*Name\s*[:\-]?\s*([A-Za-z\s]{3,25})', all_text, re.IGNORECASE)
        if doc_name_m:
            report.legal_checks.documents_name = doc_name_m.group(1).strip()

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

        prop_sit_m = re.search(r'Property\s*Situated\s*at\s*[:\-]?\s*(MC|Village\s*Abadi|[A-Za-z\s]+)', all_text, re.IGNORECASE)
        if prop_sit_m:
            cand = prop_sit_m.group(1).strip()
            if not cand.lower().startswith("is"):
                report.legal_checks.property_situated_at = cand

        sanc_m = re.search(r'Is\s*Construction\s*as\s*per\s*Sanction\s*Plan\s*[:\-]?\s*(Yes|No)', all_text, re.IGNORECASE)
        if sanc_m:
            report.legal_checks.is_sanction_plan_compliant = sanc_m.group(1).capitalize()

        path_m = re.search(r'Pathway\s*Clear\s*[:\-]?\s*(Yes|No)', all_text, re.IGNORECASE)
        if path_m:
            report.legal_checks.pathway_clear = path_m.group(1).capitalize()

        sanc_app_m = re.search(r'Sanction\s*Plan\s*Approval\s*Number\s*&\s*Date\s*[:\-]?\s*(Yes|No|[A-Za-z0-9\s]+)', all_text, re.IGNORECASE)
        if sanc_app_m:
            cand = sanc_app_m.group(1).strip()
            if not cand.lower().startswith("property"):
                report.legal_checks.sanction_plan_approval_no_date = cand

        disast_m = re.search(r'Property\s*(?:is\s*)?in\s*Disaster\s*Prone\s*Area\s*[:\-]?\s*(Yes|No)', all_text, re.IGNORECASE)
        if disast_m:
            report.legal_checks.is_disaster_prone = disast_m.group(1).capitalize()

        appr_road_m = re.search(r'Approach\s*to\s*Property\s*by\s*Public\s*Road\s*[:\-]?\s*(Yes|No)', all_text, re.IGNORECASE)
        if appr_road_m:
            report.legal_checks.approach_by_public_road = appr_road_m.group(1).capitalize()

        nala_m = re.search(r'Property\s*Situated\s*Near\s*(?:by\s*)?Nala\s*[:\-]?\s*(Yes|No)', all_text, re.IGNORECASE)
        if nala_m:
            report.legal_checks.near_nala = nala_m.group(1).capitalize()

        hte_m = re.search(r'Property\s*Situated\s*in\s*HTE\s*Line\s*[:\-]?\s*(Yes|No)', all_text, re.IGNORECASE)
        if hte_m:
            report.legal_checks.in_hte_line = hte_m.group(1).capitalize()

        util_m = re.search(r'Whether\s*Electricty,\s*Water,\s*Drainage\s*in\s*Vicinity\s*[:\-]?\s*(Yes|No)', all_text, re.IGNORECASE)
        if util_m:
            report.legal_checks.utilities_in_vicinity = util_m.group(1).capitalize()

        mast_m = re.search(r'Approved\s*Land\s*as\s*per\s*Master\s*Plan\s*[:\-]?\s*([A-Za-z]+)', all_text, re.IGNORECASE)
        if mast_m:
            report.legal_checks.approved_land_master_plan = mast_m.group(1).strip()

        uses_m = re.search(r'Current\s*Uses\s*of\s*Property\s*[:\-]?\s*([A-Za-z]+)', all_text, re.IGNORECASE)
        if uses_m:
            report.legal_checks.current_uses = uses_m.group(1).strip()

        opin_m = re.search(r'Opinion\s*(?:Abot|About)\s*Report\s*[:\-]?\s*(Positive|Negative)', all_text, re.IGNORECASE)
        if opin_m:
            report.legal_checks.opinion_about_report = opin_m.group(1).capitalize()

        occ250_m = re.search(r'%\s*Occupancy\s*in\s*250\s*Mtr\s*Radius\s*[:\-]?\s*([0-9%\-]+)', all_text, re.IGNORECASE)
        if occ250_m:
            report.legal_checks.occupancy_250m = occ250_m.group(1).strip()

        dev250_m = re.search(r'%\s*of\s*Development\s*in\s*250\s*Mtr\s*Radius\s*[:\-]?\s*([0-9%\-]+)', all_text, re.IGNORECASE)
        if dev250_m:
            report.legal_checks.development_250m = dev250_m.group(1).strip()

        prop_lim_m = re.search(r'Property\s*Limit\s*[:\-]?\s*([A-Za-z\s]+Limit|[A-Za-z\s]{3,30})', all_text, re.IGNORECASE)
        if prop_lim_m:
            cand = prop_lim_m.group(1).strip()
            if not cand.lower().startswith("adm"):
                report.legal_checks.property_limit = cand

        adm_m = re.search(r'ADM\s*\(Area\s*Development\s*and\s*Marketability\)\s*[:\-]?\s*(Low|Average|High|[A-Za-z\s]+)', all_text, re.IGNORECASE)
        if adm_m:
            report.legal_checks.adm = adm_m.group(1).strip()

        if "road" in all_text.lower():
            road_w_m = re.search(r'(\d+\s*(?:Ft|Feet|Meter|Mtr)\s*Wide)', all_text, re.IGNORECASE)
            if road_w_m:
                report.legal_checks.width_of_public_road = road_w_m.group(1).strip()

        # 15. Reference & Feedback
        ref_n_m = re.search(r'Reference\s*Name\s*[:\-]?\s*([A-Za-z\s]{3,30})', all_text, re.IGNORECASE)
        if ref_n_m:
            cand = ref_n_m.group(1).strip()
            if not cand.lower().startswith("reference") and not cand.lower().startswith("mobile"):
                report.reference.reference_name = cand

        phone_m = re.search(r'\b([6-9]\d{9})\b', all_text)
        if phone_m:
            report.reference.reference_mobile = phone_m.group(1)

        fb_m = re.search(r'(?:Feedback)\s*[:\-]?\s*([0-9\sA-Za-z\.\,\-toperSqyds]{3,40})(?:\n|\r|$)', all_text, re.IGNORECASE)
        if fb_m and not fb_m.group(1).strip().startswith("1.") and "subject property" not in fb_m.group(1).lower():
            report.reference.feedback = fb_m.group(1).strip()

        # Clean multi-line / whitespace artifacts
        def _clean_str(val: str) -> str:
            if not val:
                return ""
            s = str(val).strip()
            s = re.split(r'[\r\n]', s)[0].strip()
            return re.sub(r'\s+', ' ', s)

        report.address.nearest_landmark = _clean_str(report.address.nearest_landmark)
        report.address.plot_house_khasra = _clean_str(report.address.plot_house_khasra)
        report.address.floor_number = _clean_str(report.address.floor_number)
        report.boundaries.mismatch_remarks = _clean_str(report.boundaries.mismatch_remarks)
        report.solar_roof_vicinity.solar_install_location = _clean_str(report.solar_roof_vicinity.solar_install_location)
        report.solar_roof_vicinity.population_1km = _clean_str(report.solar_roof_vicinity.population_1km)
        report.legal_checks.documents_name = _clean_str(report.legal_checks.documents_name)
        report.legal_checks.property_limit = _clean_str(report.legal_checks.property_limit)
        report.reference.reference_name = _clean_str(report.reference.reference_name)
        # 16. Narrative Remarks
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
