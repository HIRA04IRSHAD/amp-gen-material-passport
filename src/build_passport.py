import json
import re
from pathlib import Path
import shutil
import openpyxl

INPUT_JSON = Path("output/boq_extracted.json")
TEMPLATE_EXCEL = Path("AMP_Passport_Template.xlsx")
OUTPUT_EXCEL = Path("output/passport_filled.xlsx")
OUTPUT_JSON = Path("output/passport.json")
OUTPUT_META = Path("output/building_meta.json")

# Bonus B2: EPD Database for Carbon Calculations (AMBER Columns)
EPD_DATABASE = {
    "Concrete": {"carbon_factor": 0.149, "source": "ICE Database V4.1 (Concrete C35/45) via NZBG Guide V1.0"},
    "Steel": {"carbon_factor": 1.55, "source": "ICE Database V4.0 (Steel Open Sections) via NZBG Guide V1.0"},
    "Masonry": {"carbon_factor": 0.213, "source": "ICE Database V4.0 (Brick Clay) via NZBG Guide V1.0"},
    "Timber": {"carbon_factor": 0.306, "source": "ICE Database V4.0 (Timber Hardwood) via NZBG Guide V1.0"},
    "Glass": {"carbon_factor": 1.15, "source": "ICE Database V3.0 (Glass Float) / Fair Comparison Standard"},
    "Paint/Finish": {"carbon_factor": 0.95, "source": "ICE Database V3.0 (Generic Assumed)"},
    "Plaster": {"carbon_factor": 0.163, "source": "ICE Database V4.0 (Cement Mortar 1:3, general) via NZBG Guide V1.0"},
}












# Bonus B2 continued : DENSITY
DENSITY_DATABASE = {
    "Concrete": 2400,
    "Steel": 7850,
    "Masonry": 1800,
    "Timber": 750,
    "Paint/Finish": 1300,
    "Plaster": 1900,
}

NOMINAL_MIX_GRADE = {
    "1:1:2": "M20",
    "1:1.5:3": "M20",
    "1:2:4": "M15",
    "1:3:6": "M10",
    "1:4:8": "M7.5",
    "1:5:10": "M5",
}

MIX_RATIO_RE = re.compile(r'1\s*:\s*\d+(?:\.\d+)?\s*:\s*\d+(?:\.\d+)?')
MORTAR_RATIO_RE = re.compile(r'1\s*:\s*\d+(?:\.\d+)?')
IS_CODE_RE = re.compile(r'\bIS\s*[:.]?\s*\d{2,5}(?:\s*\(\s*Part\s*[\dIVX]+\s*\))?', re.IGNORECASE)
UNIT_MULTIPLIER_RE = re.compile(r'^\s*(\d+(?:\.\d+)?)\s*(sq\.?\s?m|cu\.?\s?m)\s*$', re.IGNORECASE)



FLOOR_SECTION_RULES = [
    ("Sub-Head - III, R.C.C. Work", [
        r'\bR\s*\.?\s*C\s*\.?\s*C\s*\.?\b',
        r'reinforced\s+cement\s+concrete',
        r'reinforced\s+concrete',
    ]),
    ("Sub-Head - V, Wood Work", [
        r'wood\s*work',
        r'\btimber\b',
        r'\bteak\b',
        r'\bchowkat',
        r'\bwooden\b',
        r'flush\s+door',
        r'panel(l)?ed\s+door',
        r'\bshutter\b',
    ]),
    ("Sub-Head - IV, Brick Work", [
        r'brick\s*work',
        r'burnt\s+brick',
        r'\bmasonry\b',
    ]),
    ("Sub-Head - VI, Steel & Aluminium Fittings", [
        r'\baluminium\b',
        r'\baluminum\b',
        r'mild\s+steel',
        r'\bM\s*\.?\s*S\s*\.?\b',
        r'\bgrill\b',
        r'\brailing\b',
        r'steel\s+(window|door|gate|section|fitting)',
    ]),
    ("Sub-Head - VIII, Sanitary Installation (Rain Water Pipes)", [
        r'rain\s*water\s*pipe',
        r'\bR\s*\.?\s*W\s*\.?\s*P\s*\.?\b',
        r'\bsanitary\b',
        r'down\s*pipe',
        r'\bgutter\b',
        r'\bplumb(ing)?\b',
    ]),
    ("Sub-Head - VII, Flooring & Roof Treatment", [
        r'floor(ing)?',
        r'roof\s+treatment',
        r'terrac(e|ing)',
        r'damp[- ]?proof',
        r'water\s*proof(ing)?',
        r'\bskirting\b',
        r'\bdado\b',
    ]),
    ("Sub-Head - IX, Plastering", [
        r'\bplaster',
        r'\bpointing\b',
        r'\brendering\b',
    ]),
    ("Sub-Head - X, White/Colour Washing & Painting", [
        r'white\s*wash',
        r'colou?r\s*wash',
        r'\bpainting\b',
        r'\bpaint\b',
        r'\bdistemper\b',
        r'\bprimer\b',
        r'\bvarnish\b',
        r'\benamel\b',
    ]),
    ("Sub-Head - I, Earth Work", [
        r'earth\s*work',
        r'excavat',
        r'\bfilling\b',
        r'back\s*filling',
        r'anti[- ]?termite',
    ]),
    ("Sub-Head - II, Cement Concrete Work", [
        r'\bP\s*\.?\s*C\s*\.?\s*C\s*\.?\b',
        r'cement\s+concrete',
        r'lean\s+concrete',
        r'plain\s+concrete',
    ]),
]

FLOOR_SECTION_RULES_COMPILED = [
    (label, [re.compile(p, re.IGNORECASE) for p in patterns])
    for label, patterns in FLOOR_SECTION_RULES
]

# Fallback: if no keyword in the description matched, use the material
# category (already classified upstream) as a secondary signal.
CATEGORY_TO_SECTION = {
    "Earthwork": "Sub-Head - I, Earth Work",
    "Timber": "Sub-Head - V, Wood Work",
    "Masonry": "Sub-Head - IV, Brick Work",
    "Plaster": "Sub-Head - IX, Plastering",
    "Paint/Finish": "Sub-Head - X, White/Colour Washing & Painting",
}

DEFAULT_FLOOR_SECTION = "Sub-Head - XI, Sundry Items / Site Development"


def resolve_floor_section(description, material_category, extracted_value):
    """Classify a BOQ line item into its Sub-Head section by scanning the
    item's description text for keywords (wood -> Wood Work, RCC/reinforced
    cement concrete -> R.C.C. Work, brick -> Brick Work, etc.), instead of
    relying on the printed item number.

    Order of precedence:
      1. Keyword match against the description (most reliable, most specific).
      2. The item's already-classified material_category, mapped to a section.
      3. Whatever section the extraction step (Gemini, reading the actual
         page headers) already guessed.
      4. A generic "Sundry Items" bucket as the last resort.
    """
    text = description or ""
    for label, patterns in FLOOR_SECTION_RULES_COMPILED:
        if any(p.search(text) for p in patterns):
            return label

    if material_category in CATEGORY_TO_SECTION:
        return CATEGORY_TO_SECTION[material_category]

    if extracted_value:
        return extracted_value

    return DEFAULT_FLOOR_SECTION


def confidence_label(category, mass_basis, grade_inferred):
    if category not in EPD_DATABASE:
        return "Low"
    if mass_basis is None:
        return "Low"
    if mass_basis == "derived-volume" or grade_inferred:
        return "Medium"
    return "High"

def compute_mass_kg(vol, weight, derived_qty, derived_unit, density):
    if weight not in ("", None):
        return float(weight), "direct-weight"
    if vol not in ("", None):
        return float(vol) * density, "direct-volume"
    if derived_qty not in ("", None) and derived_unit == "cum":
        return float(derived_qty) * density, "derived-volume"
    return None, None

def normalize_unit(unit_str):
    if not unit_str:
        return ""
    u = str(unit_str).lower().strip().replace('.', '')
    if u in ['cum', 'm3', 'cubic metre', 'cu m']:
        return 'cum'
    if u in ['sqm', 'm2', 'sq m']:
        return 'sqm'
    if u in ['mtr', 'm']:
        return 'm'
    if u in ['kg']:
        return 'kg'
    if u in ['each', 'nos']:
        return 'nos'
    return unit_str

def safe_float(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "null":
        return ""
    try:
        return float(val)
    except ValueError:
        return str(val)

def normalize_unit_and_qty(raw_unit, raw_qty):
    if not raw_unit:
        return raw_qty, ""
    u_raw = str(raw_unit).strip()
    if "10 cubic" in u_raw.lower():
        if isinstance(raw_qty, (int, float)):
            return round(raw_qty * 0.01, 4), "cum"
        return raw_qty, "cum"
    mult_match = UNIT_MULTIPLIER_RE.match(u_raw)
    if mult_match:
        multiplier = float(mult_match.group(1))
        base_unit = normalize_unit(mult_match.group(2))
        if isinstance(raw_qty, (int, float)):
            return round(raw_qty * multiplier, 4), base_unit
        return raw_qty, base_unit
    return raw_qty, normalize_unit(u_raw)

def extract_mix_ratio(description, existing):
    if existing:
        return existing
    if not description:
        return ""
    m = MIX_RATIO_RE.search(description)
    if m:
        return m.group(0).replace(" ", "")
    m2 = MORTAR_RATIO_RE.search(description)
    if m2:
        return m2.group(0).replace(" ", "")
    return ""

def resolve_grade(printed_grade, category, mix_ratio):
    if printed_grade:
        return printed_grade, False
    if category == "Concrete" and mix_ratio in NOMINAL_MIX_GRADE:
        return NOMINAL_MIX_GRADE[mix_ratio], True
    return "", False

def extract_code_reference(description, printed_codes):
    if printed_codes:
        return printed_codes
    if not description:
        return ""
    codes = sorted(set(m.group(0).upper().replace(" ", "") for m in IS_CODE_RE.finditer(description)))
    return ", ".join(codes)

def build_classification(category, material, grade, mix_ratio):
    parts = [p for p in [category, material] if p]
    tail = grade or mix_ratio
    if tail:
        parts.append(str(tail))
    return " > ".join(parts) if parts else ""

def compute_derived_quantity(area, thickness_mm):
    if area not in ("", None) and thickness_mm not in ("", None):
        try:
            derived = round(float(area) * (float(thickness_mm) / 1000.0), 4)
            return derived, "cum", f"Area x thickness ({thickness_mm} mm, from description)"
        except (TypeError, ValueError):
            pass
    return "", "", ""

def build_comment(category, material, mass_kg, mass_basis, factor, source, grade_inferred, grade, mix_ratio, unit_rate, total_cost):
    material = material or "Unclassified material"

    if category not in EPD_DATABASE:
        reason = ("earthwork/excavated soil; negligible embodied material carbon"
                  if category == "Earthwork"
                  else f"category '{category}' not in current EPD database; carbon not estimated")
        return f"[EXCLUDED] material: {material} — {reason}"

    if mass_kg is None:
        return (f"[EXCLUDED] material: {material} — category '{category}' has an EPD factor "
                f"but no usable weight/volume figure was available to compute mass; "
                f"carbon not estimated")

    tag = "[ASSUMED]" if category == "Paint/Finish" else "[OK]"
    basis_note = "area × thickness (derived)" if mass_basis == "derived-volume" else mass_basis.replace("-", " ")
    carbon = round(mass_kg * factor, 2)
    
    parts = [f"{tag} material: {material} — mass {round(mass_kg, 2)} kg [{basis_note}] "
             f"× {factor} kgCO2e/kg = {carbon} kgCO2e — {source}"]
    
    if grade_inferred:
        parts.append(f"Grade {grade} inferred from nominal mix {mix_ratio} (IS 456)")
        
    if unit_rate != "" or total_cost != "":
        parts.append("Cost in INR (₹)")
        
    return "; ".join(parts)

def main():
    if not INPUT_JSON.exists():
        print(f"Error: {INPUT_JSON} not found. Please run extract_boq.py first.")
        return

    with open(INPUT_JSON, 'r', encoding='utf-8') as f:
        content = json.load(f)

    if isinstance(content, dict):
        items = content.get("line_items", [])
        extracted_meta = content.get("metadata", {})
    else:
        items = content
        extracted_meta = {}

    building_meta = {
        "Project_Name": extracted_meta.get("Project_Name", "Principal's Residence (General-Modified)"),
        "Institute": extracted_meta.get("Institute", "CENTRAL BUILDING RESEARCH INSTITUTE"),
        "Location": extracted_meta.get("Location", "ROORKEE (U.P.)"),
        "Depth_of_Foundation": extracted_meta.get("Depth_of_Foundation", "0.60 mtr."),
        "Plinth_Height": extracted_meta.get("Plinth_Height", "0.45 mtr."),
        "Plinth_Area": extracted_meta.get("Plinth_Area", "90.6 Sq.m."),
        "Seismic_Zone": extracted_meta.get("Seismic_Zone", "I to IV and V"),
        "Capacity": extracted_meta.get("Capacity", "10T/Sq.m and above"),
        "Schedule_Source": extracted_meta.get("Schedule_Source", "DSR 1989"),
        "Currency": "INR (₹)",
        "Total_Line_Items": len(items)
    }

    OUTPUT_META.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_META, 'w', encoding='utf-8') as meta_f:
        json.dump(building_meta, meta_f, indent=2, ensure_ascii=False)
    print(f"Successfully generated {OUTPUT_META}")

    rows = []
    json_export = []

    for idx, item in enumerate(items, start=1):
        raw_qty = safe_float(item.get("original_quantity", ""))
        raw_unit = item.get("original_unit", "")
        raw_qty, norm_unit = normalize_unit_and_qty(raw_unit, raw_qty)

        vol, area, length, weight, count = "", "", "", "", ""
        if isinstance(raw_qty, (int, float)) and raw_qty != "":
            if norm_unit == 'cum':
                vol = raw_qty
            elif norm_unit == 'sqm':
                area = raw_qty
            elif norm_unit == 'm':
                length = raw_qty
            elif norm_unit == 'kg':
                weight = raw_qty
            elif norm_unit == 'nos':
                count = raw_qty

        description = item.get("description", "")
        category = item.get("material_category", "")
        material = item.get("material_product", "")

        mix_ratio = extract_mix_ratio(description, item.get("mix_ratio", ""))
        grade, grade_inferred = resolve_grade(item.get("grade", ""), category, mix_ratio)
        code_reference = extract_code_reference(description, item.get("standard_code_reference", ""))
        classification = build_classification(category, material, grade, mix_ratio)

        floor_section = resolve_floor_section(description, category, item.get("floor_section", ""))
        schedule_source = item.get("schedule_source", "")

        thickness_mm = safe_float(item.get("thickness_mm", ""))
        diameter_mm = safe_float(item.get("diameter_mm", ""))
        length_mm = safe_float(item.get("length_mm", ""))
        width_mm = safe_float(item.get("width_mm", ""))
        height_mm = safe_float(item.get("height_mm", ""))
        depth_mm = safe_float(item.get("depth_mm", ""))
        derived_qty, derived_unit, derived_basis = compute_derived_quantity(area, thickness_mm)
        
        unit_rate = safe_float(item.get("unit_rate", ""))
        total_cost = safe_float(item.get("total_cost", ""))

        density = DENSITY_DATABASE.get(category, "")
        mass_kg, mass_basis = (None, None)
        if density != "":
            mass_kg, mass_basis = compute_mass_kg(vol, weight, derived_qty, derived_unit, density)

        carbon_total, carbon_factor = "", ""
        if category in EPD_DATABASE:
            carbon_factor = EPD_DATABASE[category]["carbon_factor"]
            if mass_kg is not None:
                carbon_total = round(mass_kg * carbon_factor, 2)

        source = EPD_DATABASE.get(category, {}).get("source", "")
        
        comment = build_comment(category, material, mass_kg, mass_basis, carbon_factor, source, grade_inferred, grade, mix_ratio, unit_rate, total_cost)
        
        confidence_display = confidence_label(category, mass_basis, grade_inferred)

        mapped_row = {
            1: f"GMAP-{idx:04d}",
            2: item.get("boq_item_no", ""),
            5: description,
            6: floor_section,
            7: item.get("discipline", ""),
            8: material,
            9: item.get("all_materials_detected", ""),
            10: category,
            11: confidence_display,
            12: grade,
            13: mix_ratio,
            14: raw_qty,
            15: norm_unit,
            16: vol,
            17: area,
            18: length,
            19: weight,
            20: count,
            21: derived_qty,
            22: derived_unit,
            23: derived_basis,
            24: density if density != "" else "",
            25: carbon_total,
            26: carbon_factor,
            27: schedule_source,
            28: item.get("schedule_item_code", ""),
            29: code_reference,
            30: classification,
            41: length_mm,
            42: width_mm,
            43: height_mm,
            44: thickness_mm,
            45: depth_mm,
            46: diameter_mm,
            47: safe_float(item.get("unit_rate", "")),
            48: safe_float(item.get("total_cost", "")),
            49: item.get("currency", "INR"),
            50: comment
        }
        rows.append(mapped_row)

        j_item = item.copy()
        j_item['gmap_id'] = mapped_row[1]
        j_item['floor_section'] = floor_section
        j_item['floor_section_source'] = "inferred_from_description_keywords"
        j_item['schedule_source'] = schedule_source
        j_item['currency'] = item.get("currency", "INR") 
        j_item['grade'] = grade
        j_item['mix_ratio'] = mix_ratio
        j_item['standard_code_reference'] = code_reference
        j_item['classification_matched'] = classification
        j_item['original_unit'] = norm_unit
        if raw_qty != "":
            j_item['original_quantity'] = raw_qty
        j_item['material_confidence'] = confidence_display
        j_item['material_confidence_raw'] = item.get("material_confidence", "")
        j_item['routed_quantities'] = {
            "volume_cum": vol,
            "area_sqm": area,
            "length_m": length,
            "weight_kg": weight,
            "count_nos": count
        }
        j_item['derived_quantity'] = {
            "value": derived_qty,
            "unit": derived_unit,
            "basis": derived_basis
        }
        j_item['density_kg_per_m3'] = density if density != "" else None
        j_item['mass_kg'] = round(mass_kg, 2) if mass_kg is not None else None
        j_item['mass_basis'] = mass_basis
        j_item['embodied_carbon'] = carbon_total
        j_item['comment'] = comment
        json_export.append(j_item)

    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(json_export, f, indent=2)
    print(f"Successfully generated {OUTPUT_JSON}")

    if TEMPLATE_EXCEL.exists():
        shutil.copy(TEMPLATE_EXCEL, OUTPUT_EXCEL)
        wb = openpyxl.load_workbook(OUTPUT_EXCEL)
        ws = wb['Material Passport']
        ws.delete_rows(4, 3)
        start_row = 4
        for r_idx, row_data in enumerate(rows):
            current_row = start_row + r_idx
            for col_idx, val in row_data.items():
                if val != "" and val is not None:
                    cell = ws.cell(row=current_row, column=col_idx)
                    if isinstance(val, (int, float)):
                        cell.value = val
                        cell.number_format = '0.00'
                    else:
                        cell.value = val
        wb.save(OUTPUT_EXCEL)
        print(f"Successfully generated {OUTPUT_EXCEL}")
    else:
        print("Warning: Template missing!")

if __name__ == "__main__":
    main()