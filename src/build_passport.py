import json
import pandas as pd
from pathlib import Path
import shutil
import openpyxl

INPUT_JSON = Path("output/boq_extracted.json")
TEMPLATE_EXCEL = Path("AMP_Passport_Template.xlsx")
OUTPUT_EXCEL = Path("output/passport_filled.xlsx")
OUTPUT_JSON = Path("output/passport.json")

def normalize_unit(unit_str):
    if not unit_str: return ""
    u = str(unit_str).lower().strip().replace('.', '')
    if u in ['cum', 'm3', 'cubic metre', 'cu m']: return 'cum'
    if u in ['sqm', 'm2', 'sq m']: return 'sqm'
    if u in ['mtr', 'm']: return 'm'
    if u in ['kg']: return 'kg'
    if u in ['each', 'nos']: return 'nos'
    return unit_str

def safe_float(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "null":
        return ""
    try:
        return float(val)
    except ValueError:
        return str(val)

def main():
    if not INPUT_JSON.exists():
        print(f"Error: {INPUT_JSON} not found. Please run extract_boq.py first.")
        return

    with open(INPUT_JSON, 'r', encoding='utf-8') as f:
        items = json.load(f)

    rows = []
    json_export = []
    
    for item in items:
        raw_qty = safe_float(item.get("original_quantity", ""))
        raw_unit = item.get("original_unit", "")
        norm_unit = normalize_unit(raw_unit)
            
        if raw_unit and "10 cubic" in str(raw_unit).lower() and isinstance(raw_qty, float):
            raw_qty = round(raw_qty * 0.01, 4)
            norm_unit = "cum"

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

        # Exact column index mapping based on template structure for Excel
        mapped_row = {
            2: item.get("boq_item_no", ""),               # Column B: BOQ Item No.
            5: item.get("description", ""),               # Column E: Description
            6: item.get("floor_section", ""),             # Column F: Floor / Section
            7: item.get("discipline", ""),                # Column G: Discipline
            8: item.get("material_product", ""),          # Column H: Material / Product
            9: item.get("all_materials_detected", ""),    # Column I: All Materials Detected
            10: item.get("material_category", ""),        # Column J: Material Category
            11: safe_float(item.get("material_confidence", "")), # Column K: Material Confidence
            13: item.get("mix_ratio", ""),                # Column M: Mix Ratio
            14: raw_qty,                                  # Column N: Original Quantity
            15: norm_unit,                                # Column O: Original Unit
            16: vol,                                      # Column P: Volume (m³)
            17: area,                                     # Column Q: Area (m²)
            18: length,                                   # Column R: Length (m)
            19: weight,                                   # Column S: Weight (kg)
            20: count,                                    # Column T: Count (Nos)
            28: item.get("schedule_item_code", ""),       # Column AB: Schedule Item Code
            44: safe_float(item.get("thickness_mm", "")), # Column AR: Thickness (mm)
            46: safe_float(item.get("diameter_mm", ""))   # Column AT: Diameter (mm)
        }
        rows.append(mapped_row)

        # Build clean JSON record mirroring the routed data
        j_item = item.copy()
        j_item['original_unit'] = norm_unit
        if raw_qty != "": j_item['original_quantity'] = raw_qty
        j_item['routed_quantities'] = {
            "volume_cum": vol,
            "area_sqm": area,
            "length_m": length,
            "weight_kg": weight,
            "count_nos": count
        }
        json_export.append(j_item)

    # 1. Save to passport.json correctly
    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(json_export, f, indent=2)
    print(f"✅ Successfully generated {OUTPUT_JSON}")

    # 2. Populate Excel Template
    if TEMPLATE_EXCEL.exists():
        shutil.copy(TEMPLATE_EXCEL, OUTPUT_EXCEL)
        wb = openpyxl.load_workbook(OUTPUT_EXCEL)
        ws = wb['Material Passport']
        
        start_row = 7 # Starting safely below example rows
            
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