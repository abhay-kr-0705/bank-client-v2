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
    """Extracts digital text and renders high-res page images for OCR/Vision."""
    text_content = []
    page_images = []
    try:
        doc = fitz.open(filepath)
        for page_num in range(min(len(doc), max_pages)):
            page = doc[page_num]
            txt = page.get_text()
            if txt.strip():
                text_content.append(f"--- [Page {page_num + 1}] ---\n{txt}")
            
            # Render page at 150 DPI for multimodal OCR vision
            pix = page.get_pixmap(dpi=150)
            img_bytes = pix.tobytes("jpeg")
            page_images.append(img_bytes)
        doc.close()
    except Exception as e:
        text_content.append(f"[PDF Error: {e}]")
    return "\n\n".join(text_content), page_images

class CaseExtractor:
    def __init__(self, api_key: Optional[str] = None, model_name: str = "gemini-2.5-flash"):
        self.api_key = api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
        self.model_name = model_name

    def inspect_folder(self, folder_path: str) -> Dict[str, Any]:
        """Scans folder and categorizes raw files."""
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

            ext = os.path.splitext(filename)[1].lower()
            file_type = "unknown"
            size = os.path.getsize(file_path)

            if ext == ".xlsx":
                file_type = "target_excel_template" if "sunita" in filename.lower() else "excel"
            elif ext == ".docx":
                file_type = "field_notes_docx"
                doc_text = parse_docx_file(file_path)
                aggregated_text.append(f"=== DOCX FILE: {filename} ===\n{doc_text}")
            elif ext == ".pdf":
                file_type = "property_pdf"
                pdf_text, p_images = parse_pdf_text_and_images(file_path, max_pages=20)
                aggregated_text.append(f"=== PDF FILE: {filename} ===\n{pdf_text}")
                images_to_process.extend(p_images[:4])
            elif ext in [".jpeg", ".jpg", ".png", ".webp"]:
                file_type = "site_photo"
                gps = extract_exif_gps(file_path)
                if gps and not found_gps:
                    found_gps = f"{gps[0]}, {gps[1]}"
                with open(file_path, "rb") as img_f:
                    images_to_process.append(img_f.read())

            files_info.append({
                "filename": filename,
                "path": file_path,
                "type": file_type,
                "size_bytes": size,
                "size_display": f"{round(size / 1024, 1)} KB"
            })

        return {
            "files": files_info,
            "all_text": "\n\n".join(aggregated_text),
            "images": images_to_process,
            "gps_from_exif": found_gps
        }

    def dynamic_entity_extractor(self, all_text: str, gps_exif: Optional[str] = None) -> ReportData:
        """
        Universal NLP & regex entity extractor for any Indian banking/valuation case.
        Extracts names, plot/khasra numbers, boundaries, dimensions, and addresses dynamically.
        """
        report = ReportData()
        clean_text = " ".join(all_text.split())

        # 1. Application ID
        app_id_m = re.search(r'(?:Application\s*ID|App\s*No\.?|Loan\s*No\.?|Ref\s*No\.?)\s*[:\-]?\s*([A-Za-z0-9\-_/]+)', all_text, re.IGNORECASE)
        if app_id_m:
            report.header.application_id = app_id_m.group(1).strip()
        else:
            id_fallback = re.search(r'(AP[-_]\d{5,12})', all_text, re.IGNORECASE)
            if id_fallback:
                report.header.application_id = id_fallback.group(1).strip()

        # 2. Applicant Name
        name_m = re.search(r'(?:Applicant\s*Name|Borrower\s*Name|Purchaser|Customer\s*Name)\s*[:\-]?\s*([A-Za-z\.\s]{3,40})', all_text, re.IGNORECASE)
        if name_m:
            candidate = name_m.group(1).strip()
            # clean trailing words
            candidate = re.split(r'(\n|\r|Type|Address|Age|Plot|and|Son|Wife)', candidate, flags=re.IGNORECASE)[0].strip()
            report.header.applicant_name = candidate
        else:
            purchaser_m = re.search(r'(?:Purchaser|Second\s*Party)\s*[:\-]?\s*(?:Mrs?\.?|Smt\.?|Sh\.?|Shri)?\s*([A-Za-z\s]{3,35})', all_text, re.IGNORECASE)
            if purchaser_m:
                report.header.applicant_name = purchaser_m.group(1).strip()

        # 3. Geo Tag / GPS Coordinates
        geo_m = re.search(r'(\d{2}\.\d{4,8})\s*,\s*(\d{2}\.\d{4,8})', all_text)
        if geo_m:
            report.header.geo_tag = f"{geo_m.group(1)}, {geo_m.group(2)}"
        elif gps_exif:
            report.header.geo_tag = gps_exif

        # 4. Age of Property & Structure
        age_m = re.search(r'(\d{1,2}\s*(?:Years?|Yrs?))\s*(?:old|age)?', all_text, re.IGNORECASE)
        if age_m:
            report.header.age_of_property = f"{age_m.group(1).title()}"

        if "load bearing" in all_text.lower():
            report.header.structure_type = "Load Bearing"
        elif "rcc" in all_text.lower():
            report.header.structure_type = "RCC"

        # 5. Property / Plot / Khasra Identification
        plot_m = re.search(r'(?:Plot\s*No\.?|Property\s*(?:Bearing\s*Plot\s*)?No\.?)\s*([A-Za-z0-9\-_]+)', all_text, re.IGNORECASE)
        khasra_m = re.search(r'(?:Khasra\s*No\.?|Khata\s*No\.?)\s*([0-9\s,&]+)', all_text, re.IGNORECASE)
        
        plot_val = plot_m.group(1).strip() if plot_m else ""
        if plot_val:
            report.address.plot_house_khasra = f"Property No. {plot_val}"

        # 6. Dimensions (Length & Breadth)
        dim_m = re.search(r'(\d+(?:\.\d+)?)\s*(?:x|X|\*|by|X)\s*(\d+(?:\.\d+)?)', all_text)
        if dim_m:
            d1 = float(dim_m.group(1))
            d2 = float(dim_m.group(2))
            length = max(d1, d2)
            breadth = min(d1, d2)
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
        elif report.land_measurements.land_length and report.land_measurements.land_breadth:
            calc_area = round(report.land_measurements.land_length * report.land_measurements.land_breadth, 1)
            report.land_measurements.land_area_site_sqft = f"{calc_area} Sqft"
            report.land_measurements.adopted_land_area_sqft = calc_area

        # 8. Pincode & City
        pin_m = re.search(r'\b(1100\d{2}|1200\d{2}|2013\d{2}|3020\d{2}|\d{6})\b', all_text)
        if pin_m:
            report.address.pincode = pin_m.group(1)

        # 9. Street / Landmark / Village / Colony
        street_m = re.search(r'(Gali\s*No\.?\s*[0-9A-Za-z\-_]+|Road\s*No\.?\s*[0-9A-Za-z\-_]+|Street\s*[0-9A-Za-z\-_]+)', all_text, re.IGNORECASE)
        if street_m:
            report.address.street_name = street_m.group(1).strip()

        colony_m = re.search(r'(?:Abadi\s*Known\s*as|Colony|Enclave|Garden|Vihar|Nagar)\s*[:\-]?\s*([A-Za-z\s]+(?:Extn\.?|Extension|Vihar|Nagar|Garden|Enclave))', all_text, re.IGNORECASE)
        if colony_m:
            report.address.colony_name = colony_m.group(1).strip()

        village_m = re.search(r'(?:Revenue\s*Estate\s*of\s*Village[\- ]*|Village[\- ]+)([A-Za-z]+)', all_text, re.IGNORECASE)
        if village_m:
            report.address.village_name = village_m.group(1).strip()

        # 10. Document Address & Site Address
        doc_match = re.search(r'(Property\s*(?:Bearing\s*)?Plot\s*No\.[^\n\r]*?(?:110\d{3}|\d{6}))', clean_text, re.IGNORECASE)
        if doc_match:
            report.address.address_docs = doc_match.group(1).strip()

        site_match = re.search(r'(Property\s*No\.\s*[0-9A-Za-z\-_]+,\s*Situated\s*in[^\n\r]*?(?:110\d{3}|\d{6}))', clean_text, re.IGNORECASE)
        if site_match:
            report.address.address_site = site_match.group(1).strip()

        # 11. Title Deed Boundaries (East, West, North, South)
        east_m = re.search(r'East\s*[:\-]?\s*([^,\n\r;]+)', all_text, re.IGNORECASE)
        west_m = re.search(r'West\s*[:\-]?\s*([^,\n\r;]+)', all_text, re.IGNORECASE)
        north_m = re.search(r'North\s*[:\-]?\s*([^,\n\r;]+)', all_text, re.IGNORECASE)
        south_m = re.search(r'South\s*[:\-]?\s*([^,\n\r;]+)', all_text, re.IGNORECASE)

        if east_m: report.boundaries.deed_east = east_m.group(1).strip()
        if west_m: report.boundaries.deed_west = west_m.group(1).strip()
        if north_m: report.boundaries.deed_north = north_m.group(1).strip()
        if south_m: report.boundaries.deed_south = south_m.group(1).strip()

        # Road width
        road_w_m = re.search(r'(\d+\s*(?:Ft|Feet|Meter|Mtr)\s*(?:Wide)?)', all_text, re.IGNORECASE)
        if road_w_m:
            report.legal_checks.width_of_public_road = road_w_m.group(1).strip()

        # 12. Person Met & Reference
        person_m = re.search(r'(?:Person\s*Meet|Met\s*at\s*site|Contact\s*Person)\s*[:\-]?\s*(?:Mr\.?|Mrs\.?|Sh\.?)?\s*([A-Za-z\s]{3,30})', all_text, re.IGNORECASE)
        if person_m:
            report.legal_checks.person_met = person_m.group(1).strip()

        phone_m = re.search(r'\b([6-9]\d{9})\b', all_text)
        if phone_m:
            report.reference.reference_mobile = phone_m.group(1)

        # 13. Remarks extraction (Look for numbered points 1. to 13. in DOCX or text)
        remarks_block_m = re.search(r'(1\.\s*Subject Property[\s\S]*?13\.\s*[^\n\r]+)', all_text)
        if remarks_block_m:
            report.remarks = remarks_block_m.group(1).strip() + "\n"
        else:
            # Look for any multiline bullet remarks
            numbered_m = re.findall(r'(\d+\.\s*[^\n\r]+)', all_text)
            if len(numbered_m) >= 5:
                report.remarks = "\n".join(numbered_m) + "\n"

        return report

    def extract_with_gemini(self, all_text: str, images: List[bytes]) -> Optional[ReportData]:
        """Calls Google Gemini Vision & Multimodal LLM to perform deep visual OCR."""
        if not self.api_key:
            return None

        try:
            from google import genai
            from google.genai import types

            client = genai.Client(api_key=self.api_key)
            prompt = f"""
You are an expert Indian Housing Finance & Banking Property Valuation Technical Extraction AI.
Analyze all provided documents (scanned title deeds, GPA chains, ATS, possession letters, site notes, GPS photos) and extract exact field values for the India Shelter Technical Report template.

DOCUMENT TEXT EXTRACTED:
{all_text[:35000]}

Extract and return a JSON object matching this schema:
{{
  "header": {{
    "report_title": "India Shelter Report",
    "application_id": "string",
    "applicant_name": "string",
    "property_type": "Row House | Independent Floor | Apartment | Commercial Building",
    "completion_percent": 1.0,
    "structure_type": "RCC | Load Bearing | Temporary",
    "age_of_property": "string (e.g. 08 Years)",
    "dwelling_units_owned": 1,
    "geo_tag": "lat, long string"
  }},
  "address": {{
    "street_name": "string",
    "nearest_landmark": "string",
    "village_name": "string",
    "city": "string",
    "plot_house_khasra": "string",
    "floor_number": "Entire Property",
    "colony_name": "string",
    "address_site": "Full string",
    "address_docs": "Full string",
    "pincode": "string or int",
    "district": "string"
  }},
  "boundaries": {{
    "site_east": "string",
    "site_west": "string",
    "site_north": "string",
    "site_south": "string",
    "deed_east": "string",
    "deed_west": "string",
    "deed_north": "string",
    "deed_south": "string",
    "boundary_matching": "Yes | No",
    "mismatch_remarks": "NA",
    "occupancy_status": "Seller | Borrower | Tenant | Vacant"
  }},
  "solar_roof_vicinity": {{
    "solar_install_location": "Ground | Rooftop | None",
    "shadow_free_roof_sqft": 0,
    "parapet_wall_height": 0,
    "cracks_in_roof": 0,
    "broken_above_roof": 0,
    "access_to_reach_roof": 0,
    "roof_length_sqft": 15,
    "roof_breadth_sqft": 38,
    "is_outreach": "No | Yes",
    "population_1km": "Above 5000",
    "primary_schools_1km": 1,
    "secondary_schools_1km": 1,
    "govt_institutions_vicinity": 1
  }},
  "land_measurements": {{
    "land_length": 38.0,
    "land_breadth": 15.0,
    "land_area_site_sqft": "569.7 Sqft",
    "adopted_land_area_sqft": 569.7,
    "per_unit_land_rate": 0,
    "total_land_value": "=B46*B47"
  }},
  "construction_floors": {{
    "floors": [
      {{"name": "Basement", "actual_area": 0, "permissible_area": 0, "adopted_area": 0}},
      {{"name": "Stilt Floor", "actual_area": 0, "permissible_area": 0, "adopted_area": 0}},
      {{"name": "Ground Floor", "actual_area": 0, "permissible_area": 0, "adopted_area": 0}},
      {{"name": "First Floor", "actual_area": 0, "permissible_area": 0, "adopted_area": 0}},
      {{"name": "Second Floor", "actual_area": 0, "permissible_area": 0, "adopted_area": 0}},
      {{"name": "Third Floor", "actual_area": 0, "permissible_area": 0, "adopted_area": 0}},
      {{"name": "Fouth Floor", "actual_area": 0, "permissible_area": 0, "adopted_area": 0}},
      {{"name": "Fifth Floor", "actual_area": 0, "permissible_area": 0, "adopted_area": 0}},
      {{"name": "Six Floor", "actual_area": 0, "permissible_area": 0, "adopted_area": 0}},
      {{"name": "Seven Floor", "actual_area": 0, "permissible_area": 0, "adopted_area": 0}}
    ],
    "built_up_rate": 0,
    "total_property_value": 0
  }},
  "accommodation": {{
    "no_of_floors": 5,
    "toilet_available": "Yes",
    "no_of_lifts": 0,
    "apartments_per_floor": 1,
    "electricity_meter_installed": "Yes",
    "electricity_meter_number": "NA"
  }},
  "legal_checks": {{
    "documents_name": "Other",
    "person_met": "string",
    "relation_with_owner": "string",
    "property_situated_at": "MC",
    "is_sanction_plan_compliant": "No",
    "pathway_clear": "Yes",
    "sanction_plan_approval_no_date": "No",
    "is_disaster_prone": "No",
    "approach_by_public_road": "Yes",
    "near_nala": "No",
    "in_hte_line": "No",
    "utilities_in_vicinity": "Yes",
    "approved_land_master_plan": "Residential",
    "width_of_public_road": "23 Ft Wide",
    "current_uses": "Residential",
    "opinion_about_report": "Negative | Positive",
    "occupancy_250m": "80%-90%",
    "tentative_rent": "",
    "development_250m": "80%-90%",
    "property_limit": "Within MC Limit",
    "adm": "Average"
  }},
  "reference": {{
    "reference_name": "Local Enquiry",
    "reference_mobile": "string",
    "feedback": "string"
  }},
  "remarks": "Complete 13-point numbered technical remarks string"
}}
"""
            contents = [prompt]
            for img_b in images[:4]:
                contents.append(types.Part.from_bytes(data=img_b, mime_type="image/jpeg"))

            response = client.models.generate_content(
                model=self.model_name,
                contents=contents,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.1
                )
            )

            if response and response.text:
                data_dict = json.loads(response.text)
                return ReportData(**data_dict)
        except Exception as e:
            print(f"[Gemini Vision OCR Error: {e}] - Falling back to dynamic regex NLP engine.")
            return None

    def process_case_folder(self, folder_path: str) -> ReportData:
        """Full pipeline: Ingest files -> Multimodal Vision OCR -> Dynamic NLP Entity Extractor."""
        scan_res = self.inspect_folder(folder_path)
        all_text = scan_res["all_text"]
        images = scan_res["images"]
        gps = scan_res["gps_from_exif"]

        report = None
        if self.api_key:
            report = self.extract_with_gemini(all_text, images)

        if not report:
            report = self.dynamic_entity_extractor(all_text, gps)

        report.case_name = os.path.basename(folder_path.rstrip("/\\"))
        report.raw_files_summary = scan_res["files"]
        return report
