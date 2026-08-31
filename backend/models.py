from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any

class FloorArea(BaseModel):
    name: str
    actual_area: float = 0.0
    permissible_area: float = 0.0
    adopted_area: float = 0.0

class HeaderInfo(BaseModel):
    report_title: str = "India Shelter Report"
    application_id: str = ""
    applicant_name: str = ""
    property_type: str = "Row House"
    completion_percent: float = 1.0
    structure_type: str = "RCC"
    age_of_property: str = "08 Years"
    dwelling_units_owned: int = 1
    geo_tag: str = ""

class AddressInfo(BaseModel):
    street_name: str = ""
    nearest_landmark: str = ""
    village_name: str = ""
    city: str = ""
    plot_house_khasra: str = ""
    floor_number: str = "Entire Property"
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
    boundary_matching: str = "Yes"
    mismatch_remarks: str = "NA"
    occupancy_status: str = "Seller"

class SolarRoofVicinity(BaseModel):
    solar_install_location: str = "Ground"
    shadow_free_roof_sqft: float = 0.0
    parapet_wall_height: float = 0.0
    cracks_in_roof: float = 0.0
    broken_above_roof: float = 0.0
    access_to_reach_roof: float = 0.0
    roof_length_sqft: float = 15.0
    roof_breadth_sqft: float = 38.0
    is_outreach: str = "No"
    population_1km: str = "Above 5000"
    primary_schools_1km: int = 1
    secondary_schools_1km: int = 1
    govt_institutions_vicinity: int = 1

class LandMeasurements(BaseModel):
    land_length: float = 38.0
    land_breadth: float = 15.0
    land_area_site_sqft: str = "569.7 Sqft"
    adopted_land_area_sqft: float = 569.7
    per_unit_land_rate: float = 0.0
    total_land_value: str = "=B46*B47"

class ConstructionFloors(BaseModel):
    floors: List[FloorArea] = [
        FloorArea(name="Basement", actual_area=0, permissible_area=0, adopted_area=0),
        FloorArea(name="Stilt Floor", actual_area=0, permissible_area=0, adopted_area=0),
        FloorArea(name="Ground Floor", actual_area=0, permissible_area=0, adopted_area=0),
        FloorArea(name="First Floor", actual_area=0, permissible_area=0, adopted_area=0),
        FloorArea(name="Second Floor", actual_area=0, permissible_area=0, adopted_area=0),
        FloorArea(name="Third Floor", actual_area=0, permissible_area=0, adopted_area=0),
        FloorArea(name="Fouth Floor", actual_area=0, permissible_area=0, adopted_area=0),
        FloorArea(name="Fifth Floor", actual_area=0, permissible_area=0, adopted_area=0),
        FloorArea(name="Six Floor", actual_area=0, permissible_area=0, adopted_area=0),
        FloorArea(name="Seven Floor", actual_area=0, permissible_area=0, adopted_area=0),
    ]
    built_up_rate: float = 0.0
    total_property_value: float = 0.0

class AccommodationInfo(BaseModel):
    no_of_floors: int = 5
    toilet_available: str = "Yes"
    no_of_lifts: int = 0
    apartments_per_floor: int = 1
    electricity_meter_installed: str = "Yes"
    electricity_meter_number: str = "NA"
    proposed_per_unit_value: Optional[float] = None
    proposed_per_unit_area: Optional[float] = None
    proposed_accommodation: str = ""
    toilet_proposed: str = ""
    deviation: str = ""

class LegalStatutoryChecks(BaseModel):
    documents_name: str = "Other"
    other_docs: str = ""
    person_met: str = ""
    relation_with_owner: str = ""
    property_situated_at: str = "MC"
    is_sanction_plan_compliant: str = "No"
    pathway_clear: str = "Yes"
    sanction_plan_approval_no_date: str = "No"
    is_disaster_prone: str = "No"
    approach_by_public_road: str = "Yes"
    near_nala: str = "No"
    in_hte_line: str = "No"
    utilities_in_vicinity: str = "Yes"
    approved_land_master_plan: str = "Residential"
    width_of_public_road: str = "23 Ft Wide"
    current_uses: str = "Residential"
    opinion_about_report: str = "Negative"
    occupancy_250m: str = "80%-90%"
    tentative_rent: str = ""
    development_250m: str = "80%-90%"
    property_limit: str = "Within MC Limit"
    adm: str = "Average"

class ReferenceEnquiry(BaseModel):
    reference_type: str = ""
    reference_name: str = "Local Enquiry"
    reference_mobile: str = ""
    feedback: str = ""

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
