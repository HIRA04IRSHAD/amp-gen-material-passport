import json
import pandas as pd
from pathlib import Path
import shutil
import openpyxl

INPUT_JSON = Path("output/boq_extracted.json")
TEMPLATE_EXCEL = Path("AMP_Passport_Template.xlsx")
OUTPUT_EXCEL = Path("output/passport_filled.xlsx")
OUTPUT_JSON = Path("output/passport.json")

def is_number(s):
    try:
        float(s)
        return True
    except (ValueError, TypeError):
        return False

def get_number(s):
    try:
        f = float(s)
        return int(f) if f.is_integer() else f
    except (ValueError, TypeError):
        return s

def main():
    if not INPUT_JSON.exists():
        print(f"Error: {INPUT_JSON} not found. Please run extract_boq.py first.")
        return

    print("Reading extracted JSON data...")
    with open(INPUT_JSON, 'r', encoding='utf-8') as f:
        items = json.load(f)

    # 1. Export JSON format
    df = pd.DataFrame(items)
    OUTPUT_EXCEL.parent.mkdir(parents=True, exist_ok=True)
    df.to_json(OUTPUT_JSON, orient="records", indent=2)

    print(f"Writing data to {OUTPUT_EXCEL}...")
    if TEMPLATE_EXCEL.exists():
        shutil.copy(TEMPLATE_EXCEL, OUTPUT_EXCEL)
        wb = openpyxl.load_workbook(OUTPUT_EXCEL)
        ws = wb['Material Passport']
        
        # 2. Find the header row dynamically
        header_row_idx = None
        for r in range(1, 10):
            vals = [str(ws.cell(row=r, column=c).value).strip().lower() for c in range(1, 20) if ws.cell(row=r, column=c).value]
            if any("description" in v for v in vals) and any("quantity" in v for v in vals):
                header_row_idx = r
                break
                
        if not header_row_idx:
            print("Error: Could not find column headers in the template.")
            return
            
        # 3. Create a map of JSON keys to exact Excel column numbers
        col_map = {}
        for c in range(1, 30):
            val = ws.cell(row=header_row_idx, column=c).value
            if val:
                val_str = str(val).strip().lower()
                if "item no" in val_str or "sl.no" in val_str: col_map["boq_item_no"] = c
                elif "description" in val_str: col_map["description"] = c
                elif "floor" in val_str or "section" in val_str: col_map["floor_section"] = c
                elif "discipline" in val_str: col_map["discipline"] = c
                elif "product" in val_str: col_map["material_product"] = c
                elif "category" in val_str: col_map["material_category"] = c
                elif "quantity" in val_str: col_map["original_quantity"] = c
                elif "unit" in val_str and "rate" not in val_str: col_map["original_unit"] = c
                elif "rate" in val_str: col_map["unit_rate"] = c
                elif "cost" in val_str or "amount" in val_str: col_map["total_cost"] = c
                elif "code" in val_str or "dsr" in val_str or "schedule" in val_str: col_map["schedule_item_code"] = c

        # 4. Find the first completely empty row to start writing
        start_row = header_row_idx + 1
        desc_col = col_map.get("description", 5)
        while ws.cell(row=start_row, column=desc_col).value is not None:
            start_row += 1
            
        # 5. Write the data securely to exact columns
        for r_idx, item in enumerate(items):
            current_row = start_row + r_idx
            for key, col_idx in col_map.items():
                val = item.get(key, "")
                if val is not None and val != "":
                    # Fix the green triangle issue permanently
                    if is_number(val):
                        val = get_number(val)
                try:
                    ws.cell(row=current_row, column=col_idx, value=val)
                except AttributeError:
                    pass

        wb.save(OUTPUT_EXCEL)
        print("✅ STEP 2 Complete! Excel mapped perfectly by column names.")
    else:
        print("Warning: Template missing!")

if __name__ == "__main__":
    main()