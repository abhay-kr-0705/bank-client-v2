"""
Validation & Confidence Decision Engine
Executes domain-specific format and consistency checks on extracted report entities:
- Application ID & Applicant Name completeness
- GPS Coordinates validity (India boundary range: 6°N - 38°N, 68°E - 98°E)
- Pincode 6-digit numeric check
- Dimensional math check (Length * Breadth vs Adopted Area)
- Boundary completeness (Site & Deed 4 directions)
- Confidence Threshold Decision:
    HIGH (>= 0.85) -> Auto-Approve
    LOW  (< 0.85)  -> HIGHLIGHT IT FOR REVIEW
"""

import re
from typing import Dict, Any, Tuple
from .models import ReportData, FieldConfidence

CONFIDENCE_THRESHOLD = 0.85

class ValuationValidator:
    @classmethod
    def validate_and_score(cls, report: ReportData) -> ReportData:
        """
        Runs format validation rules, evaluates field-level confidence,
        and assigns HIGH / LOW decisions with review reasons.
        """
        confidences = report.field_confidences or {}
        review_count = 0
        all_scores = []

        def set_field_decision(key: str, val: Any, base_conf: float, rule_valid: bool, reason: str = "", source: str = "Offline Extraction"):
            nonlocal review_count
            if not val or val == "" or val == 0:
                conf = 0.0
                needs_rev = False
            else:
                conf = base_conf if rule_valid else max(0.3, base_conf * 0.6)
                needs_rev = conf < CONFIDENCE_THRESHOLD or not rule_valid
                if needs_rev:
                    review_count += 1
                all_scores.append(conf)

            confidences[key] = FieldConfidence(
                field_key=key,
                value=val,
                confidence=round(conf, 2),
                source_document=source,
                needs_review=needs_rev,
                review_reason=reason if needs_rev else None
            )

        # 1. Header Validation
        h = report.header
        # Application ID
        app_id_valid = bool(h.application_id and len(h.application_id) >= 4)
        app_id_reason = "Application ID is missing or incomplete" if not app_id_valid and h.application_id else ""
        existing_app_conf = confidences.get("application_id").confidence if "application_id" in confidences else 0.95
        set_field_decision("application_id", h.application_id, existing_app_conf, app_id_valid, app_id_reason)

        # Applicant Name
        name_valid = bool(h.applicant_name and len(h.applicant_name.strip()) >= 3 and not h.applicant_name.lower().startswith("applicant"))
        name_reason = "Applicant name appears incomplete or placeholder" if not name_valid and h.applicant_name else ""
        existing_name_conf = confidences.get("applicant_name").confidence if "applicant_name" in confidences else 0.95
        set_field_decision("applicant_name", h.applicant_name, existing_name_conf, name_valid, name_reason)

        # Geo Tag
        geo_valid = False
        geo_reason = "Geo-tag missing or not in valid lat, lon format"
        if h.geo_tag:
            geo_m = re.search(r'(-?\d{1,2}\.\d+)\s*,\s*(-?\d{1,3}\.\d+)', h.geo_tag)
            if geo_m:
                lat, lon = float(geo_m.group(1)), float(geo_m.group(2))
                if 6.0 <= lat <= 38.0 and 68.0 <= lon <= 98.0:
                    geo_valid = True
                    geo_reason = ""
                else:
                    geo_reason = f"Coordinates ({lat}, {lon}) outside valid India bounds"
        existing_geo_conf = confidences.get("geo_tag").confidence if "geo_tag" in confidences else 0.90
        set_field_decision("geo_tag", h.geo_tag, existing_geo_conf, geo_valid, geo_reason)

        # Property Type
        pt_valid = bool(h.property_type and len(h.property_type) > 2)
        set_field_decision("property_type", h.property_type, 0.90, pt_valid, "Property type not determined")

        # 2. Address Validation
        a = report.address
        pin_valid = False
        pin_reason = "Pincode is missing"
        if a.pincode:
            clean_pin = re.sub(r'\D', '', str(a.pincode))
            if len(clean_pin) == 6:
                pin_valid = True
                pin_reason = ""
            else:
                pin_reason = f"Invalid Pincode '{a.pincode}' (must be 6 digits)"
        set_field_decision("pincode", a.pincode, 0.95, pin_valid, pin_reason)

        city_valid = bool(a.city and len(a.city) >= 3)
        set_field_decision("city", a.city, 0.90, city_valid, "City name missing or unverified")

        addr_site_valid = bool(a.address_site and len(a.address_site) >= 10)
        set_field_decision("address_site", a.address_site, 0.90, addr_site_valid, "Site address incomplete or unverified")

        addr_docs_valid = bool(a.address_docs and len(a.address_docs) >= 10)
        set_field_decision("address_docs", a.address_docs, 0.90, addr_docs_valid, "Legal deed address incomplete or unverified")

        # 3. Dimensions & Land Measurements Validation
        l = report.land_measurements
        dim_valid = False
        dim_reason = "Length and breadth dimensions missing"
        if l.land_length > 0 and l.land_breadth > 0:
            dim_valid = True
            dim_reason = ""
            # Cross-check math consistency if adopted area is provided
            calc_area = l.land_length * l.land_breadth
            if l.adopted_land_area_sqft > 0:
                diff_pct = abs(calc_area - l.adopted_land_area_sqft) / l.adopted_land_area_sqft
                if diff_pct > 0.15:
                    dim_valid = False
                    dim_reason = f"Length*Breadth ({calc_area:.1f}) deviates from adopted area ({l.adopted_land_area_sqft})"

        set_field_decision("land_length", l.land_length, 0.92, dim_valid, dim_reason)
        set_field_decision("land_breadth", l.land_breadth, 0.92, dim_valid, dim_reason)
        set_field_decision("adopted_land_area_sqft", l.adopted_land_area_sqft, 0.92, bool(l.adopted_land_area_sqft > 0), "Adopted area missing")

        # 4. Boundaries Validation
        b = report.boundaries
        site_b_count = sum(1 for val in [b.site_east, b.site_west, b.site_north, b.site_south] if val and val.strip())
        site_b_valid = site_b_count >= 3
        site_b_reason = f"Only {site_b_count}/4 site boundary directions identified" if not site_b_valid and site_b_count > 0 else ""
        for dir_key, dir_val in [("site_east", b.site_east), ("site_west", b.site_west), ("site_north", b.site_north), ("site_south", b.site_south)]:
            set_field_decision(dir_key, dir_val, 0.88, site_b_valid and bool(dir_val), site_b_reason)

        deed_b_count = sum(1 for val in [b.deed_east, b.deed_west, b.deed_north, b.deed_south] if val and val.strip())
        deed_b_valid = deed_b_count >= 3
        deed_b_reason = f"Only {deed_b_count}/4 deed boundary directions identified" if not deed_b_valid and deed_b_count > 0 else ""
        for dir_key, dir_val in [("deed_east", b.deed_east), ("deed_west", b.deed_west), ("deed_north", b.deed_north), ("deed_south", b.deed_south)]:
            set_field_decision(dir_key, dir_val, 0.88, deed_b_valid and bool(dir_val), deed_b_reason)

        # 5. Reference Mobile Validation
        ref = report.reference
        ref_valid = False
        ref_reason = ""
        if ref.reference_mobile:
            clean_phone = re.sub(r'\D', '', str(ref.reference_mobile))
            if len(clean_phone) == 10 and clean_phone[0] in '6789':
                ref_valid = True
            else:
                ref_reason = f"Mobile '{ref.reference_mobile}' is not a standard 10-digit number"
        set_field_decision("reference_mobile", ref.reference_mobile, 0.95, ref_valid, ref_reason)

        # 6. Overall Confidence Calculation
        overall = sum(all_scores) / len(all_scores) if all_scores else 0.0
        report.field_confidences = confidences
        report.overall_confidence = round(overall, 2)
        report.review_needed_count = review_count

        return report
