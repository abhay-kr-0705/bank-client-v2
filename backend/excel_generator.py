import os
import tempfile
import openpyxl
from openpyxl.utils import get_column_letter
from typing import Union
from .models import ReportData

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOCAL_TEMP_DIR = os.path.join(BASE_DIR, "uploads", "temp")
os.makedirs(LOCAL_TEMP_DIR, exist_ok=True)
tempfile.tempdir = LOCAL_TEMP_DIR

TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "base_template.xlsx")

def generate_excel_report(data: Union[ReportData, dict], output_path: str) -> str:
    """
    Populates data into the exact template structure and saves to output_path.
    Preserves all styling, formulas, merged cells, and Sheet2 dropdown validation lists.
    """
    if isinstance(data, dict):
        report = ReportData(**data)
    else:
        report = data

    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Base template not found at {TEMPLATE_PATH}")

    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb["Sheet1"]

    # Header
    if report.header.report_title:
        ws["A1"] = report.header.report_title
    ws["B2"] = report.header.application_id
    ws["D2"] = report.header.applicant_name
    ws["B3"] = report.header.property_type
    ws["B4"] = report.header.completion_percent
    ws["B5"] = report.header.structure_type
    ws["B6"] = report.header.age_of_property
    ws["B7"] = report.header.dwelling_units_owned
    ws["B8"] = report.header.geo_tag

    # Address
    ws["B10"] = report.address.street_name
    ws["B11"] = report.address.nearest_landmark
    ws["B12"] = report.address.village_name
    ws["B13"] = report.address.city
    ws["B14"] = report.address.plot_house_khasra
    ws["B15"] = report.address.floor_number
    ws["B16"] = report.address.colony_name
    ws["B17"] = report.address.address_site
    ws["B18"] = report.address.address_docs
    try:
        ws["B19"] = int(report.address.pincode) if str(report.address.pincode).isdigit() else report.address.pincode
    except Exception:
        ws["B19"] = report.address.pincode
    ws["D19"] = report.address.district

    # Boundaries
    ws["A22"] = report.boundaries.site_east
    ws["B22"] = report.boundaries.site_west
    ws["C22"] = report.boundaries.site_north
    ws["D22"] = report.boundaries.site_south

    ws["A25"] = report.boundaries.deed_east
    ws["B25"] = report.boundaries.deed_west
    ws["C25"] = report.boundaries.deed_north
    ws["D25"] = report.boundaries.deed_south

    ws["B26"] = f" {report.boundaries.boundary_matching.strip()}" if report.boundaries.boundary_matching else " Yes"
    ws["B27"] = report.boundaries.mismatch_remarks
    ws["B28"] = report.boundaries.occupancy_status

    # Solar, Roof, Vicinity
    ws["B30"] = f" {report.solar_roof_vicinity.solar_install_location.strip()}" if report.solar_roof_vicinity.solar_install_location else " Ground"
    ws["B31"] = report.solar_roof_vicinity.shadow_free_roof_sqft
    ws["B32"] = report.solar_roof_vicinity.parapet_wall_height
    ws["B33"] = report.solar_roof_vicinity.cracks_in_roof
    ws["B34"] = report.solar_roof_vicinity.broken_above_roof
    ws["B35"] = report.solar_roof_vicinity.access_to_reach_roof
    ws["B37"] = report.solar_roof_vicinity.roof_length_sqft
    ws["D37"] = report.solar_roof_vicinity.roof_breadth_sqft
    ws["B38"] = report.solar_roof_vicinity.is_outreach
    ws["B39"] = report.solar_roof_vicinity.population_1km
    ws["B40"] = report.solar_roof_vicinity.primary_schools_1km
    ws["B41"] = report.solar_roof_vicinity.secondary_schools_1km
    ws["B42"] = report.solar_roof_vicinity.govt_institutions_vicinity

    # Land Measurements
    ws["B44"] = report.land_measurements.land_length
    ws["D44"] = report.land_measurements.land_breadth
    ws["B45"] = report.land_measurements.land_area_site_sqft
    ws["B46"] = report.land_measurements.adopted_land_area_sqft
    ws["B47"] = report.land_measurements.per_unit_land_rate
    ws["B48"] = "=B46*B47"

    # Construction Floors
    floor_rows = [52, 53, 54, 55, 56, 57, 58, 59, 60, 61]
    for idx, row_idx in enumerate(floor_rows):
        if idx < len(report.construction_floors.floors):
            fl = report.construction_floors.floors[idx]
            ws[f"B{row_idx}"] = fl.actual_area
            ws[f"C{row_idx}"] = fl.permissible_area
            ws[f"D{row_idx}"] = fl.adopted_area
        else:
            ws[f"B{row_idx}"] = 0
            ws[f"C{row_idx}"] = 0
            ws[f"D{row_idx}"] = 0

    ws["B62"] = "=SUM(B52:B61)"
    ws["C62"] = 0
    ws["D62"] = "=SUM(D52:D61)"
    ws["B63"] = "=D62"
    ws["B64"] = report.construction_floors.built_up_rate
    ws["B65"] = "=B63*B64"
    ws["C66"] = report.construction_floors.total_property_value

    # Accommodation
    ws["B76"] = report.accommodation.no_of_floors
    ws["D76"] = f" {report.accommodation.toilet_available.strip()}" if report.accommodation.toilet_available else " Yes"
    ws["B77"] = report.accommodation.no_of_lifts
    ws["D77"] = report.accommodation.apartments_per_floor
    ws["B78"] = f" {report.accommodation.electricity_meter_installed.strip()}" if report.accommodation.electricity_meter_installed else " Yes"
    ws["D78"] = report.accommodation.electricity_meter_number
    ws["B79"] = "=B4"
    ws["B80"] = "=B63"
    ws["B82"] = report.accommodation.proposed_per_unit_value
    ws["D82"] = report.accommodation.proposed_per_unit_area
    ws["B83"] = "=B82*D82"

    # Legal & Statutory Checks
    ws["B88"] = report.legal_checks.documents_name
    ws["B90"] = report.legal_checks.person_met
    ws["D90"] = report.legal_checks.relation_with_owner
    ws["B91"] = f" {report.legal_checks.property_situated_at.strip()}" if report.legal_checks.property_situated_at else " MC"
    ws["D91"] = report.legal_checks.is_sanction_plan_compliant
    ws["B92"] = report.legal_checks.pathway_clear
    ws["D92"] = report.legal_checks.sanction_plan_approval_no_date
    ws["B93"] = report.legal_checks.is_disaster_prone
    ws["D93"] = f" {report.legal_checks.approach_by_public_road.strip()}" if report.legal_checks.approach_by_public_road else " Yes"
    ws["B94"] = report.legal_checks.near_nala
    ws["D94"] = report.legal_checks.in_hte_line
    ws["B95"] = report.legal_checks.utilities_in_vicinity
    ws["D95"] = report.legal_checks.approved_land_master_plan
    ws["B96"] = report.legal_checks.width_of_public_road
    ws["D96"] = report.legal_checks.current_uses
    ws["B97"] = " Positive" if report.legal_checks.opinion_about_report and report.legal_checks.opinion_about_report.strip().lower() == "positive" else (report.legal_checks.opinion_about_report.strip() if report.legal_checks.opinion_about_report else "")
    ws["D97"] = report.legal_checks.occupancy_250m
    ws["B98"] = report.legal_checks.tentative_rent
    ws["D98"] = report.legal_checks.development_250m
    ws["B99"] = report.legal_checks.property_limit
    ws["D99"] = report.legal_checks.adm

    # Reference
    ws["B102"] = report.reference.reference_name
    try:
        ws["B103"] = int(report.reference.reference_mobile) if str(report.reference.reference_mobile).isdigit() else report.reference.reference_mobile
    except Exception:
        ws["B103"] = report.reference.reference_mobile
    ws["B104"] = report.reference.feedback

    # Detailed Narrative Remarks
    ws["A106"] = report.remarks

    # Apply Review Highlights and Comments for Low-Confidence or Flagged Fields
    try:
        from .smart_mapper import SmartFieldMapper
        from openpyxl.styles import PatternFill
        from openpyxl.comments import Comment

        review_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
        if getattr(report, "field_confidences", None):
            for field_key, fc in report.field_confidences.items():
                if getattr(fc, "needs_review", False):
                    coord = SmartFieldMapper.STANDARD_COORDINATE_MAP.get(field_key)
                    if coord:
                        try:
                            cell = ws[coord]
                            cell.fill = review_fill
                            conf_pct = int(getattr(fc, "confidence", 0.5) * 100)
                            reason = getattr(fc, "review_reason", None) or "Please verify this field value"
                            cell.comment = Comment(f"[REVIEW NEEDED] (Confidence: {conf_pct}%)\n{reason}", "Valuation Engine")
                        except Exception as cell_err:
                            print(f"[Highlight cell warning {coord}]: {cell_err}")
    except Exception as hl_err:
        print(f"[Excel Highlighting Warning]: {hl_err}")

    # Save
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    return output_path
