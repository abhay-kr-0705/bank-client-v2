from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any, Union

class FloorArea(BaseModel):
    name: str
    actual_area: float = 0.0
    permissible_area: float = 0.0
    adopted_area: float = 0.0

class HeaderInfo(BaseModel):
    report_title: str = "India Shelter Report"
    application_id: str = ""
    applicant_name: str = ""
    property_type: str = ""
    completion_percent: float = 0.0
    structure_type: str = ""
    age_of_property: str = ""
    dwelling_units_owned: int = 0
    geo_tag: str = ""

class AddressInfo(BaseModel):
    street_name: str = ""
    nearest_landmark: str = ""
    village_name: str = ""
    city: str = ""
    plot_house_khasra: str = ""
    floor_number: str = ""
    colony_name: str = ""
    address_site: str = ""
    address_docs: str = ""
    pincode: str = ""
    district: str = ""

class BoundariesInfo(BaseModel):
    site_east: str = ""
    site_west: str = ""
    site_north: str = ""
    site_south: str = ""
    deed_east: str = ""
    deed_west: str = ""
    deed_north: str = ""
    deed_south: str = ""
    boundary_matching: str = ""
    mismatch_remarks: str = ""
    occupancy_status: str = ""

class SolarRoofVicinity(BaseModel):
    solar_install_location: str = ""
    shadow_free_roof_sqft: float = 0.0
    parapet_wall_height: float = 0.0
    cracks_in_roof: float = 0.0
    broken_above_roof: float = 0.0
    access_to_reach_roof: float = 0.0
    roof_length_sqft: float = 0.0
    roof_breadth_sqft: float = 0.0
    is_outreach: str = ""
    population_1km: str = ""
    primary_schools_1km: int = 0
    secondary_schools_1km: int = 0
    govt_institutions_vicinity: int = 0

class LandMeasurements(BaseModel):
    land_length: float = 0.0
    land_breadth: float = 0.0
    land_area_site_sqft: str = ""
    adopted_land_area_sqft: float = 0.0
    per_unit_land_rate: float = 0.0
    total_land_value: Union[str, float, int] = "=B46*B47"

class ConstructionFloors(BaseModel):
    floors: List[FloorArea] = [
        FloorArea(name="Basement", actual_area=0.0, permissible_area=0.0, adopted_area=0.0),
        FloorArea(name="Stilt Floor", actual_area=0.0, permissible_area=0.0, adopted_area=0.0),
        FloorArea(name="Ground Floor", actual_area=0.0, permissible_area=0.0, adopted_area=0.0),
        FloorArea(name="First Floor", actual_area=0.0, permissible_area=0.0, adopted_area=0.0),
        FloorArea(name="Second Floor", actual_area=0.0, permissible_area=0.0, adopted_area=0.0),
        FloorArea(name="Third Floor", actual_area=0.0, permissible_area=0.0, adopted_area=0.0),
        FloorArea(name="Fouth Floor", actual_area=0.0, permissible_area=0.0, adopted_area=0.0),
        FloorArea(name="Fifth Floor", actual_area=0.0, permissible_area=0.0, adopted_area=0.0),
        FloorArea(name="Six Floor", actual_area=0.0, permissible_area=0.0, adopted_area=0.0),
        FloorArea(name="Seven Floor", actual_area=0.0, permissible_area=0.0, adopted_area=0.0),
    ]
    built_up_rate: float = 0.0
    total_property_value: float = 0.0

class AccommodationInfo(BaseModel):
    no_of_floors: int = 0
    toilet_available: str = ""
    no_of_lifts: int = 0
    apartments_per_floor: int = 0
    electricity_meter_installed: str = ""
    electricity_meter_number: str = ""
    proposed_per_unit_value: Optional[float] = None
    proposed_per_unit_area: Optional[float] = None
    proposed_accommodation: str = ""
    toilet_proposed: str = ""
    deviation: str = ""

class LegalStatutoryChecks(BaseModel):
    documents_name: str = ""
    other_docs: str = ""
    person_met: str = ""
    relation_with_owner: str = ""
    property_situated_at: str = ""
    is_sanction_plan_compliant: str = ""
    pathway_clear: str = ""
    sanction_plan_approval_no_date: str = ""
    is_disaster_prone: str = ""
    approach_by_public_road: str = ""
    near_nala: str = ""
    in_hte_line: str = ""
    utilities_in_vicinity: str = ""
    approved_land_master_plan: str = ""
    width_of_public_road: str = ""
    current_uses: str = ""
    opinion_about_report: str = ""
    occupancy_250m: str = ""
    tentative_rent: str = ""
    development_250m: str = ""
    property_limit: str = ""
    adm: str = ""

class ReferenceEnquiry(BaseModel):
    reference_type: str = ""
    reference_name: str = ""
    reference_mobile: str = ""
    feedback: str = ""

class FieldConfidence(BaseModel):
    field_key: str
    value: Any = None
    confidence: float = 1.0  # 0.0 to 1.0
    source_document: Optional[str] = None
    needs_review: bool = False
    review_reason: Optional[str] = None

class ReportData(BaseModel):
    case_name: str = "New Case"
    header: HeaderInfo = Field(default_factory=HeaderInfo)
    address: AddressInfo = Field(default_factory=AddressInfo)
    boundaries: BoundariesInfo = Field(default_factory=BoundariesInfo)
    solar_roof_vicinity: SolarRoofVicinity = Field(default_factory=SolarRoofVicinity)
    land_measurements: LandMeasurements = Field(default_factory=LandMeasurements)
    construction_floors: ConstructionFloors = Field(default_factory=ConstructionFloors)
    accommodation: AccommodationInfo = Field(default_factory=AccommodationInfo)
    legal_checks: LegalStatutoryChecks = Field(default_factory=LegalStatutoryChecks)
    reference: ReferenceEnquiry = Field(default_factory=ReferenceEnquiry)
    remarks: str = ""
    raw_files_summary: List[Dict[str, Any]] = Field(default_factory=list)
    field_confidences: Dict[str, FieldConfidence] = Field(default_factory=dict)
    overall_confidence: float = 1.0
    review_needed_count: int = 0
