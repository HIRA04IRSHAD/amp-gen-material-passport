import json
import re
from pathlib import Path
import shutil
import openpyxl

INPUT_JSON = Path("output/boq_extracted.json")
TEMPLATE_EXCEL = Path("AMP_Passport_Template.xlsx")
OUTPUT_EXCEL = Path("output/passport_filled.xlsx")
OUTPUT_JSON = Path("output/passport.json")

# Values sourced via NotebookLM (ICE Database V4.0/V4.1)
EPD_DATABASE = {
    "Concrete": {"carbon_factor": 0.149, "source": "ICE Database V4.1 (Concrete C35/45) via NZBG Guide V1.0"},
    "Steel": {"carbon_factor": 1.55, "source": "ICE Database V4.0 (Steel Open Sections) via NZBG Guide V1.0"},
    "Masonry": {"carbon_factor": 0.213, "source": "ICE Database V4.0 (Brick Clay) via NZBG Guide V1.0"},
    "Timber": {"carbon_factor": 0.306, "source": "ICE Database V4.0 (Timber Hardwood) via NZBG Guide V1.0"},
    "Glass": {"carbon_factor": 1.15, "source": "ICE Database V3.0 (Glass Float) / Fair Comparison Standard"},
    "Paint/Finish": {"carbon_factor": 0.95, "source": "ICE Database V3.0 (Generic Assumed)"}
}

# Nominal-mix -> nominal grade, per IS 456:2000 conventions for ordinary
# concrete grades. Used ONLY as a fallback when the BoQ text gives a ratio
# but no named grade is printed anywhere (extract_boq.py's "grade" field
# will already be filled when a grade IS printed -- this never overrides
# that). Only applied to items already classified as Concrete, so masonry
# mortar ratios never get mislabelled with a concrete grade.
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
# Matches unit strings like "100 Sq.m" where the printed quantity is a
# count of multiples of the base unit, not the base unit itself. Kept as a
# deterministic post-process step rather than asking the model to do this
# arithmetic, since unit-multiplier parsing is exact and shouldn't depend
# on LLM math.
UNIT_MULTIPLIER_RE = re.compile(r'^\s*(\d+(?:\.\d+)?)\s*(sq\.?\s?m|cu\.?\s?m)\s*$', re.IGNORECASE)


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
    """Handles multiplier-prefixed units like "100 Sq.m" (qty is a count of
    hundreds of sqm) or "10 Cubic decimetre" (qty is already in that unit,
    x0.01 -> cum). Falls back to plain normalize_unit for everything else.
    """
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
    """extract_boq.py already fills this most of the time; this is a plain
    regex safety net for the items it missed (e.g. "1:6" mortar mentions)."""
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
    """Prefer a grade actually printed in the BoQ text (from extraction).
    Only fall back to inferring one from the nominal mix ratio -- and only
    for Concrete -- when nothing was printed. Returns (grade_text, inferred)."""
    if printed_grade:
        return printed_grade, False
    if category == "Concrete" and mix_ratio in NOMINAL_MIX_GRADE:
        return NOMINAL_MIX_GRADE[mix_ratio], True
    return "", False


def extract_code_reference(description, printed_codes):
    """Prefer what extraction already found; regex over the description is
    just a safety net for anything it missed."""
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
    """Where a BoQ item gives an area (Sq.m) and a thickness (from the
    description, e.g. "40 mm thick"), the implied material volume is a
    genuinely useful derived quantity for a material passport. Basis is
    recorded so the assumption is auditable."""
    if area not in ("", None) and thickness_mm not in ("", None):
        try:
            derived = round(float(area) * (float(thickness_mm) / 1000.0), 4)
            return derived, "cum", f"Area x thickness ({thickness_mm} mm, from description)"
        except (TypeError, ValueError):
            pass
    return "", "", ""


def main():
    if not INPUT_JSON.exists():
        print(f"Error: {INPUT_JSON} not found. Please run extract_boq.py first.")
        return

    with open(INPUT_JSON, 'r', encoding='utf-8') as f:
        items = json.load(f)

    rows = []
    json_export = []
    # Safety nets only -- extract_boq.py's prompt now asks the model to
    # carry these forward itself. These local carries just cover the odd
    # row where the model still returns null.
    last_floor_section = ""
    last_schedule_source = ""

    for idx, item in enumerate(items, start=1):
        raw_qty = safe_float(item.get("original_quantity", ""))
        raw_unit = item.get("original_unit", "")
        raw_qty, norm_unit = normalize_unit_and_qty(raw_unit, raw_qty)

        # Initialize quantity mapping columns to empty
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

        # Floor / Section: trust the model's own carry-forward first; only
        # fall back to our local last-seen value if it returned null.
        floor_section = item.get("floor_section") or last_floor_section
        last_floor_section = floor_section or last_floor_section

        # Schedule source (e.g. "DSR 1989"): same pattern.
        schedule_source = item.get("schedule_source") or last_schedule_source
        last_schedule_source = schedule_source or last_schedule_source

        thickness_mm = safe_float(item.get("thickness_mm", ""))
        diameter_mm = safe_float(item.get("diameter_mm", ""))
        length_mm = safe_float(item.get("length_mm", ""))
        width_mm = safe_float(item.get("width_mm", ""))
        height_mm = safe_float(item.get("height_mm", ""))
        depth_mm = safe_float(item.get("depth_mm", ""))
        derived_qty, derived_unit, derived_basis = compute_derived_quantity(area, thickness_mm)

        # AMBER Carbon Calculation (Bonus B2)
        carbon_total, carbon_factor = "", ""
        comment_parts = []
        if category in EPD_DATABASE and isinstance(raw_qty, (int, float)):
            carbon_factor = EPD_DATABASE[category]["carbon_factor"]
            carbon_total = round(raw_qty * carbon_factor, 2)
            comment_parts.append(f"EPD Source: {EPD_DATABASE[category]['source']}")
        if grade_inferred:
            comment_parts.append(f"Grade {grade} inferred from nominal mix {mix_ratio} (IS 456); not printed in BoQ")
        comment = "; ".join(comment_parts)

        # Exact column index mapping based on template structure for Excel
        mapped_row = {
            1: f"GMAP-{idx:04d}",                          # Column A: GMAP Id (surrogate, no external ID source given)
            2: item.get("boq_item_no", ""),                # Column B: BOQ Item No.
            5: description,                                # Column E: Description
            6: floor_section,                              # Column F: Floor / Section
            7: item.get("discipline", ""),                 # Column G: Discipline
            8: material,                                   # Column H: Material / Product
            9: item.get("all_materials_detected", ""),     # Column I: All Materials Detected
            10: category,                                  # Column J: Material Category
            11: safe_float(item.get("material_confidence", "")),  # Column K: Material Confidence
            12: grade,                                     # Column L: Grade
            13: mix_ratio,                                 # Column M: Mix Ratio
            14: raw_qty,                                   # Column N: Original Quantity
            15: norm_unit,                                 # Column O: Original Unit
            16: vol,                                       # Column P: Volume (m3)
            17: area,                                      # Column Q: Area (m2)
            18: length,                                    # Column R: Length (m)
            19: weight,                                    # Column S: Weight (kg)
            20: count,                                     # Column T: Count (Nos)
            21: derived_qty,                                # Column U: Derived Quantity
            22: derived_unit,                               # Column V: Derived Quantity Unit
            23: derived_basis,                              # Column W: Derived Quantity Basis
            25: carbon_total,                               # Column Y: Embodied Carbon A1-A3
            26: carbon_factor,                              # Column Z: GWP / kg
            27: schedule_source,                            # Column AA: Schedule (DSR/SOR)
            28: item.get("schedule_item_code", ""),         # Column AB: Schedule Item Code
            29: code_reference,                             # Column AC: Standard / Code Reference
            30: classification,                             # Column AD: Classification (Matched)
            41: length_mm,                                  # Column AO: Length (mm)
            42: width_mm,                                   # Column AP: Width (mm)
            43: height_mm,                                  # Column AQ: Height (mm)
            44: thickness_mm,                               # Column AR: Thickness (mm)
            45: depth_mm,                                   # Column AS: Depth (mm)
            46: diameter_mm,                                # Column AT: Diameter (mm)
            47: safe_float(item.get("unit_rate", "")),      # Column AU: Unit Rate (blank if not printed)
            48: safe_float(item.get("total_cost", "")),     # Column AV: Total Cost (blank if not printed)
            50: comment                                     # Column AX: Comment
            # Columns C, D (Article Number, External DB Id) are intentionally
            # left out: they need a manufacturer/vendor catalogue or an
            # external materials database to match against, and a generic
            # government DSR BoQ has neither.
        }
        rows.append(mapped_row)

        # Build clean JSON record mirroring the routed data
        j_item = item.copy()
        j_item['gmap_id'] = mapped_row[1]
        j_item['floor_section'] = floor_section
        j_item['schedule_source'] = schedule_source
        j_item['grade'] = grade
        j_item['mix_ratio'] = mix_ratio
        j_item['standard_code_reference'] = code_reference
        j_item['classification_matched'] = classification
        j_item['original_unit'] = norm_unit
        if raw_qty != "":
            j_item['original_quantity'] = raw_qty
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
        j_item['embodied_carbon'] = carbon_total
        j_item['comment'] = comment
        json_export.append(j_item)

    # 1. Save to passport.json correctly
    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(json_export, f, indent=2)
    print(f"Successfully generated {OUTPUT_JSON}")

    # 2. Populate Excel Template
    if TEMPLATE_EXCEL.exists():
        shutil.copy(TEMPLATE_EXCEL, OUTPUT_EXCEL)
        wb = openpyxl.load_workbook(OUTPUT_EXCEL)
        ws = wb['Material Passport']

        start_row = 7  # Starting safely below example rows

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
