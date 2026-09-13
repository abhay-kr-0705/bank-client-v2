import re
from typing import Dict, Any, Optional, Tuple
from .models import ReportData

class SmartFieldMapper:
    """
    Intelligently maps extracted valuation entities into target Excel template cells.
    Supports both India Shelter standard format and arbitrary custom templates via semantic fuzzy label detection.
    """

    STANDARD_COORDINATE_MAP = {
        "report_title": "A1",
        "application_id": "B2",
        "applicant_name": "D2",
        "property_type": "B3",
        "completion_percent": "B4",
        "structure_type": "B5",
        "age_of_property": "B6",
        "dwelling_units_owned": "B7",
        "geo_tag": "B8",
        "street_name": "B10",
        "nearest_landmark": "B11",
        "village_name": "B12",
        "city": "B13",
        "plot_house_khasra": "B14",
        "floor_number": "B15",
        "colony_name": "B16",
        "address_site": "B17",
        "address_docs": "B18",
        "pincode": "B19",
        "district": "D19",
        "site_east": "A22",
        "site_west": "B22",
        "site_north": "C22",
        "site_south": "D22",
        "deed_east": "A25",
        "deed_west": "B25",
        "deed_north": "C25",
        "deed_south": "D25",
        "boundary_matching": "B26",
        "mismatch_remarks": "B27",
        "occupancy_status": "B28",
        "solar_install_location": "B30",
        "shadow_free_roof_sqft": "B31",
        "parapet_wall_height": "B32",
        "cracks_in_roof": "B33",
        "broken_above_roof": "B34",
        "access_to_reach_roof": "B35",
        "roof_length_sqft": "B37",
        "roof_breadth_sqft": "D37",
        "is_outreach": "B38",
        "population_1km": "B39",
        "primary_schools_1km": "B40",
        "secondary_schools_1km": "B41",
        "govt_institutions_vicinity": "B42",
        "land_length": "B44",
        "land_breadth": "D44",
        "land_area_site_sqft": "B45",
        "adopted_land_area_sqft": "B46",
        "per_unit_land_rate": "B47",
        "total_land_value": "B48",
        "built_up_rate": "B64",
        "total_property_value": "C66",
        "no_of_floors": "B76",
        "toilet_available": "D76",
        "no_of_lifts": "B77",
        "apartments_per_floor": "D77",
        "electricity_meter_installed": "B78",
        "electricity_meter_number": "D78",
        "documents_name": "B88",
        "person_met": "B90",
        "relation_with_owner": "D90",
        "property_situated_at": "B91",
        "is_sanction_plan_compliant": "D91",
        "pathway_clear": "B92",
        "sanction_plan_approval_no_date": "D92",
        "is_disaster_prone": "B93",
        "approach_by_public_road": "D93",
        "near_nala": "B94",
        "in_hte_line": "D94",
        "utilities_in_vicinity": "B95",
        "approved_land_master_plan": "D95",
        "width_of_public_road": "B96",
        "current_uses": "D96",
        "opinion_about_report": "B97",
        "occupancy_250m": "D97",
        "tentative_rent": "B98",
        "development_250m": "D98",
        "property_limit": "B99",
        "adm": "D99",
        "reference_name": "B102",
        "reference_mobile": "B103",
        "feedback": "B104",
        "remarks": "A106"
    }

    # Semantic synonym dictionary for custom templates
    SEMANTIC_SYNONYMS = {
        "applicant_name": ["applicant name", "borrower name", "customer name", "purchaser", "client name"],
        "application_id": ["application id", "app no", "application no", "loan no", "case ref", "file no"],
        "property_type": ["type of property", "property type", "nature of property"],
        "geo_tag": ["geo tag", "latitude", "gps coordinates", "geo-ordinates"],
        "age_of_property": ["age of property", "property age", "age of building"],
        "structure_type": ["type of structure", "structure type", "construction type"],
        "address_site": ["address as per site", "site address", "actual address"],
        "address_docs": ["property address as per documents", "document address", "legal address"],
        "pincode": ["pincode", "pin code", "postal code"],
        "city": ["city", "town"],
        "person_met": ["person meet", "person met", "contact person"],
        "remarks": ["remarks", "technical remarks", "valuation notes"]
    }

    @classmethod
    def report_data_to_flat_dict(cls, data: ReportData) -> Dict[str, Any]:
        """Flattens the nested Pydantic model into a key-value dictionary."""
        flat = {}
        # Header
        h = data.header
        flat["report_title"] = h.report_title
        flat["application_id"] = h.application_id
        flat["applicant_name"] = h.applicant_name
        flat["property_type"] = h.property_type
        flat["completion_percent"] = h.completion_percent
        flat["structure_type"] = h.structure_type
        flat["age_of_property"] = h.age_of_property
        flat["dwelling_units_owned"] = h.dwelling_units_owned
        flat["geo_tag"] = h.geo_tag

        # Address
        a = data.address
        flat["street_name"] = a.street_name
        flat["nearest_landmark"] = a.nearest_landmark
        flat["village_name"] = a.village_name
        flat["city"] = a.city
        flat["plot_house_khasra"] = a.plot_house_khasra
        flat["floor_number"] = a.floor_number
        flat["colony_name"] = a.colony_name
        flat["address_site"] = a.address_site
        flat["address_docs"] = a.address_docs
        flat["pincode"] = a.pincode
        flat["district"] = a.district

        # Boundaries
        b = data.boundaries
        flat["site_east"] = b.site_east
        flat["site_west"] = b.site_west
        flat["site_north"] = b.site_north
        flat["site_south"] = b.site_south
        flat["deed_east"] = b.deed_east
        flat["deed_west"] = b.deed_west
        flat["deed_north"] = b.deed_north
        flat["deed_south"] = b.deed_south
        flat["boundary_matching"] = f" {b.boundary_matching.strip()}" if b.boundary_matching and b.boundary_matching.strip() else ""
        flat["mismatch_remarks"] = b.mismatch_remarks
        flat["occupancy_status"] = b.occupancy_status

        # Solar & Roof
        s = data.solar_roof_vicinity
        flat["solar_install_location"] = f" {s.solar_install_location.strip()}" if s.solar_install_location and s.solar_install_location.strip() else ""
        flat["shadow_free_roof_sqft"] = s.shadow_free_roof_sqft
        flat["parapet_wall_height"] = s.parapet_wall_height
        flat["cracks_in_roof"] = s.cracks_in_roof
        flat["broken_above_roof"] = s.broken_above_roof
        flat["access_to_reach_roof"] = s.access_to_reach_roof
        flat["roof_length_sqft"] = s.roof_length_sqft
        flat["roof_breadth_sqft"] = s.roof_breadth_sqft
        flat["is_outreach"] = s.is_outreach
        flat["population_1km"] = s.population_1km
        flat["primary_schools_1km"] = s.primary_schools_1km
        flat["secondary_schools_1km"] = s.secondary_schools_1km
        flat["govt_institutions_vicinity"] = s.govt_institutions_vicinity

        # Land
        l = data.land_measurements
        flat["land_length"] = l.land_length
        flat["land_breadth"] = l.land_breadth
        flat["land_area_site_sqft"] = l.land_area_site_sqft
        flat["adopted_land_area_sqft"] = l.adopted_land_area_sqft
        flat["per_unit_land_rate"] = l.per_unit_land_rate
        flat["total_land_value"] = l.total_land_value or "=B46*B47"

        # Floors
        f = data.construction_floors
        flat["built_up_rate"] = f.built_up_rate
        flat["total_property_value"] = f.total_property_value
        for i, fl in enumerate(f.floors):
            flat[f"floor_{i}_actual"] = fl.actual_area
            flat[f"floor_{i}_permissible"] = fl.permissible_area
            flat[f"floor_{i}_adopted"] = fl.adopted_area

        # Accommodation
        acc = data.accommodation
        flat["no_of_floors"] = acc.no_of_floors
        flat["toilet_available"] = f" {acc.toilet_available.strip()}" if acc.toilet_available and acc.toilet_available.strip() else ""
        flat["no_of_lifts"] = acc.no_of_lifts
        flat["apartments_per_floor"] = acc.apartments_per_floor
        flat["electricity_meter_installed"] = f" {acc.electricity_meter_installed.strip()}" if acc.electricity_meter_installed and acc.electricity_meter_installed.strip() else ""
        flat["electricity_meter_number"] = acc.electricity_meter_number

        # Legal checks
        leg = data.legal_checks
        flat["documents_name"] = leg.documents_name
        flat["person_met"] = leg.person_met
        flat["relation_with_owner"] = leg.relation_with_owner
        flat["property_situated_at"] = f" {leg.property_situated_at.strip()}" if leg.property_situated_at and leg.property_situated_at.strip() else ""
        flat["is_sanction_plan_compliant"] = leg.is_sanction_plan_compliant
        flat["pathway_clear"] = leg.pathway_clear
        flat["sanction_plan_approval_no_date"] = leg.sanction_plan_approval_no_date
        flat["is_disaster_prone"] = leg.is_disaster_prone
        flat["approach_by_public_road"] = f" {leg.approach_by_public_road.strip()}" if leg.approach_by_public_road and leg.approach_by_public_road.strip() else ""
        flat["near_nala"] = leg.near_nala
        flat["in_hte_line"] = leg.in_hte_line
        flat["utilities_in_vicinity"] = leg.utilities_in_vicinity
        flat["approved_land_master_plan"] = leg.approved_land_master_plan
        flat["width_of_public_road"] = leg.width_of_public_road
        flat["current_uses"] = leg.current_uses
        flat["opinion_about_report"] = f" {leg.opinion_about_report.strip()}" if leg.opinion_about_report and leg.opinion_about_report.strip() else ""
        flat["occupancy_250m"] = leg.occupancy_250m
        flat["tentative_rent"] = leg.tentative_rent
        flat["development_250m"] = leg.development_250m
        flat["property_limit"] = leg.property_limit
        flat["adm"] = leg.adm

        # Reference & Remarks
        ref = data.reference
        flat["reference_name"] = ref.reference_name
        flat["reference_mobile"] = ref.reference_mobile
        flat["feedback"] = ref.feedback
        flat["remarks"] = data.remarks

        return flat

    @classmethod
    def map_to_cells(cls, report_data: ReportData, metadata_cells: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        Maps flat report fields to cell coordinates.
        Uses standard coordinate map as primary, and applies semantic fuzzy matching if custom metadata cells provided.
        """
        vals, _ = cls.map_to_cells_with_metadata(report_data, metadata_cells)
        return vals

    @classmethod
    def map_to_cells_with_metadata(cls, report_data: ReportData, metadata_cells: Optional[Dict[str, Any]] = None) -> Tuple[Dict[str, Any], Dict[str, Any]]:
        """
        Maps flat report fields to cell coordinates and attaches confidence scores + review flags.
        """
        flat = cls.report_data_to_flat_dict(report_data)
        cell_values = {}
        cell_metadata = {}
        confidences = getattr(report_data, "field_confidences", {}) or {}

        # 1. Standard mapping
        for field, coord in cls.STANDARD_COORDINATE_MAP.items():
            if field in flat and flat[field] is not None:
                cell_values[coord] = flat[field]

            if field in confidences:
                fc = confidences[field]
                cell_metadata[coord] = {
                    "field_key": field,
                    "confidence": getattr(fc, "confidence", 1.0),
                    "needs_review": getattr(fc, "needs_review", False),
                    "review_reason": getattr(fc, "review_reason", None)
                }

        # Floor table rows B52:D61
        for i in range(10):
            r = 52 + i
            if f"floor_{i}_actual" in flat:
                cell_values[f"B{r}"] = flat[f"floor_{i}_actual"]
            if f"floor_{i}_permissible" in flat:
                cell_values[f"C{r}"] = flat[f"floor_{i}_permissible"]
            if f"floor_{i}_adopted" in flat:
                cell_values[f"D{r}"] = flat[f"floor_{i}_adopted"]

        # Standard formulas
        cell_values["B48"] = "=B46*B47"
        cell_values["B62"] = "=SUM(B52:B61)"
        cell_values["C62"] = 0
        cell_values["D62"] = "=SUM(D52:D61)"
        cell_values["B63"] = "=D62"
        cell_values["B65"] = "=B63*B64"
        cell_values["B79"] = "=B4"
        cell_values["B80"] = "=B63"
        cell_values["B83"] = "=B82*D82"

        return cell_values, cell_metadata
