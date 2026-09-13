"""
Ground-Truth Valuation Knowledge & Few-Shot Learning Reference
Synthesized from verified banking valuation cases (e.g. SUNITA DEVI BACHHAN KUMAR).
Provides exemplar patterns, boundary resolution heuristics, dimension normalization,
and standard 13-point narrative structures for Indian Housing Finance & Technical Reports.
"""

from typing import Dict, Any

# Exemplar Ground-Truth Case Data (Sunita Devi)
SUNITA_CASE_GROUND_TRUTH: Dict[str, Any] = {
    "case_name": "SUNITA DEVI BACHHAN KUMAR",
    "header": {
        "report_title": "India Shelter Report",
        "application_id": "AP-10524478",
        "applicant_name": "Mrs. Sunita Devi",
        "property_type": "Row House",
        "completion_percent": 1.0,
        "structure_type": "RCC",
        "age_of_property": "08 Years",
        "dwelling_units_owned": 1,
        "geo_tag": "28.627023, 77.026233"
    },
    "address": {
        "street_name": "Gali No. 14",
        "nearest_landmark": "Nearby Aggarwal Store",
        "village_name": "Nawada",
        "city": "New Delhi",
        "plot_house_khasra": "Property No. 8-B",
        "floor_number": "Entire Property",
        "colony_name": "Vipin garden Extn.",
        "address_site": "Property No. 8-B, Situated in Village-Nawada, Gali No. 14, Vipin garden Extn., Uttam Nagar, New Delhi-110059",
        "address_docs": "Property Bearing Plot No. 8-B, Out of Khasra No. 75 & 76, Situated in the Revenue Estate of Village-Nawada, Delhi State Delhi in the Abadi Known as Vipin garden Extn., Uttam Nagar, New Delhi-110059",
        "pincode": "110059",
        "district": "New Delhi"
    },
    "boundaries": {
        "site_east": "Others Property/Meter No. 21901820",
        "site_west": "Others Property/Meter No. 46215085",
        "site_north": "Road 23 Ft Wide",
        "site_south": "Others Property",
        "deed_east": "Plot No. 8-A",
        "deed_west": "Plot No. 9-A",
        "deed_north": "Road 23 Ft Wide",
        "deed_south": "Other Land",
        "boundary_matching": "Yes",
        "mismatch_remarks": "NA",
        "occupancy_status": "Seller"
    },
    "solar_roof_vicinity": {
        "solar_install_location": "Ground",
        "shadow_free_roof_sqft": 0.0,
        "parapet_wall_height": 0.0,
        "cracks_in_roof": 0.0,
        "broken_above_roof": 0.0,
        "access_to_reach_roof": 0.0,
        "roof_length_sqft": 15.0,
        "roof_breadth_sqft": 38.0,
        "is_outreach": "No",
        "population_1km": "Above 5000",
        "primary_schools_1km": 1,
        "secondary_schools_1km": 1,
        "govt_institutions_vicinity": 1
    },
    "land_measurements": {
        "land_length": 38.0,
        "land_breadth": 15.0,
        "land_area_site_sqft": "569.7 Sqft",
        "adopted_land_area_sqft": 569.7,
        "per_unit_land_rate": 0.0,
        "total_land_value": "=B46*B47"
    },
    "construction_floors": {
        "built_up_rate": 0.0,
        "total_property_value": 0.0
    },
    "accommodation": {
        "no_of_floors": 5,
        "toilet_available": "Yes",
        "no_of_lifts": 0,
        "apartments_per_floor": 1,
        "electricity_meter_installed": "Yes",
        "electricity_meter_number": "NA"
    },
    "legal_checks": {
        "documents_name": "Other",
        "person_met": "Mr. Gauarv",
        "relation_with_owner": "Applicant's Son",
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
        "opinion_about_report": "Negative",
        "occupancy_250m": "80%-90%",
        "tentative_rent": "",
        "development_250m": "80%-90%",
        "property_limit": "Within MC  Limit",
        "adm": "Average"
    },
    "reference": {
        "reference_name": "Local Enquiry",
        "reference_mobile": "9540637533",
        "feedback": "1 L to 1.10 L per Sqyds"
    },
    "remarks": """1. Subject Property is a S+UG+3 storied residential house built over a plot having area 63.3 Sq Yrd.
2. Access to the property is through Road 23 Ft Wide in North direction.
3. The subject property is 08 years old & same was found seller-occupied as on date of time of site visit.
4. The subject property has identified with help of the applicant & local enquiry, Name Board.
5. Surrounding Vicinity is approx. 80%-90%% within 250 meters radius.
6. The Owner has done 100% ground coverage over the plot and projected front side approx. total 3 ft. beyond the plot limit.
7. The subject property falls under MC Limits.
8. This is to inform you that the applicant had called the engineer to the site. However, upon arrival, the seller informed that the property’s bayana (token/advance) has not yet been completed, and therefore they are not allowing the internal visit at this stage. As a result, only the external (outside) visit was conducted.
9. Provided GPA/ATS is draft only thus required registered title documents.
10. A Soft Copy of GPA/ATS/Possesion Latter/Will Deed has been provided, dated: 12/05/2023 in favor of (1). Mr. Surender Singh S/o Mr. Ram Mahar & (2). Mrs. Sarla W/o Mr. Surender Singh for Property Bearing Plot No. 8-B, Out of Khasra No. 75 & 76, Situated in the Revenue Estate of Village-Nawada, Delhi State Delhi in the Abadi Known as Vipin Garden Extn., Uttam Nagar, New Delhi-110059 for having plot area 63.3 sqyds (15 X 38).
11. A Soft Copy of Draft ATS has been provided, undated in between of (1). Mr. Surender Singh & (2). Mrs. Sarla (Seller) & Mrs. Sunitta Devi W/o Mr. Bachhan Kumar (Purchaser) for Property Bearing Plot No. 8-B, Out of Khasra No. 75 & 76, Situated in the Revenue Estate of Village-Nawada, Delhi State Delhi in the Abadi Known as Vipin Garden Extn., Uttam Nagar, New Delhi-110059 for having plot area 63.3 sqyds  
12. The Geo-ordinates of Subject Property are 28.627023, 77.026233.
13. Value of the subject property has not been released due to above mentioned deviations."""
}

def get_few_shot_prompt_guidance() -> str:
    """Returns formatted training prompt context explaining how valuation inputs translate to output fields."""
    return """
GROUND-TRUTH DOMAIN RULES & EXTRACTION INSTRUCTIONS:
CRITICAL ZERO-HARDCODING RULE:
- ALL values MUST be extracted strictly from the provided DOCUMENT TEXT or images.
- NEVER invent, hallucinate, or copy values from examples. If a field is not mentioned in the input documents, return empty string "" or 0 for numeric fields.
- Do NOT return literal placeholder text like "AP-XXXXXXX", "string", "NA", or "null".

1. APPLICATION ID & APPLICANT:
   - Extract the exact real Application ID from the text (e.g. AP-XXXXXXXX format, Loan/Ref/File No). Extract the real identifier digits found in the document.
   - Applicant name should be titled (e.g. Mrs./Mr. <FullName>). If only purchaser name is found in draft ATS or title deed, use that name.

2. BOUNDARIES DISTINCTION:
   - 'Site Boundaries' are observed physically on site (e.g. adjacent electric meter numbers, physical landmarks, road width).
   - 'Deed Boundaries' are specified in legal title deed / GPA / ATS schedule (e.g. Plot numbers, adjacent owners, or roads as stated in the deed).
   - Approach road direction usually corresponds to the front facing boundary.

3. DIMENSIONS & AREA:
   - Dimensions are typically Length x Breadth found in deed or field notes (e.g. 15 X 38, 20 X 40, etc.).
   - Longer side = Length, Shorter side = Breadth.
   - Area calculation: If area is given in Sq. Yards, convert to Sq. Feet by multiplying by 9 (Sqft = SqYds * 9.0).
   - Adopted Land Area in sq ft should match the actual deed/site plot area in sq ft.

4. ACCOMMODATION:
   - Count the total number of floors constructed (e.g. Stilt + Upper Ground + 3 floors = 5 floors; Ground + 1 = 2 floors).
   - Ground coverage & projections: note if owner projected beyond plot limit (e.g. balcony/chhajja).

5. LEGAL & SITE CONTACT:
   - Person Met at Site: identify contact person and their relation (e.g. contact name and relation such as Applicant's Son, Spouse, Self).
   - Municipal status: record whether property is within Municipal Corporation (MC) limits or Gram Panchayat.

6. TECHNICAL REMARKS:
   - Always synthesize a structured, numbered 13-point narrative valuation remark detailing:
     1. Structure & floor count + plot area.
     2. Road access width & direction.
     3. Age of property & occupancy (Seller/Borrower/Tenant).
     4. Property identification method (Applicant/local enquiry/board).
     5. Surrounding vicinity occupancy % in 250m.
     6. Plot coverage & projections beyond limit.
     7. Municipal authority (MC limits).
     8. Inspection status (Internal vs External visit, bayana status).
     9. Title document registration status (draft ATS vs registered deed).
     10. Chain of title / previous owner details (GPA/ATS/Possession letter dates, sellers).
     11. Current agreement details (parties, undated/dated, area).
     12. GPS Geo-coordinates.
     13. Final valuation recommendation / release decision.
"""

