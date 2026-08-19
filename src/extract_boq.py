import argparse
import json
import os
import time
from pathlib import Path

import fitz  # PyMuPDF
from google import genai
from google.genai import types
from google.genai import errors as genai_errors

# ---- CONFIG ----
DPI = 350              # for the local review images only
MODEL = "gemini-3.6-flash"
MAX_RETRIES = 3
RETRY_BASE_DELAY = 10

# Final prompt: Updated with all green columns and strict anti-shifting rules
EXTRACTION_PROMPT = """
You are a highly precise Data Extraction AI reading a scanned, dot-matrix Bill of Quantities (BoQ) for a civil construction project in India.

CRITICAL ANTI-SHIFTING RULES (MUST FOLLOW):
1. HORIZONTAL TRACKING: Read strictly row by row, left to right. Never mix data between rows.
2. MISSING VALUES: If a column (like Quantity or Schedule Code) is blank, illegible, or spans multiple lines, you MUST output `null` for that field. NEVER shift a value from row N+1 into row N.
3. DECIMAL POINTS: Use the context of the unit and surrounding numbers to infer the correct decimal placement (e.g., '7.3' instead of '73').

SUB-ITEM RULES:
1. Items with multiple sub-parts (e.g., 16 i, 16 ii, or a, b, c) MUST be split into separate JSON objects (e.g., "16a", "16b"). Extract the specific quantity for EACH sub-part accurately.

Extract EVERY line item and return a JSON array of objects with EXACTLY these fields (use null if not found):
- "source_page": (Integer) The page number (1-indexed) this item appears on.
- "boq_item_no": (String) The printed item number.
- "description": (String) Full description text.
- "floor_section": (String) Extract any section header above the item like "Schedule A" or "Sub-Head I". If none, return null.
- "discipline": (String) Intelligently classify: "Civil & Sitework", "Structural", "Electrical", "Plumbing & Sanitary", "HVAC", "Finishes", "Other".
- "material_product": (String) The primary specific material mentioned (e.g., "Cement concrete", "Burnt brick", "Teak wood").
- "all_materials_detected": (String) A comma-separated list of all distinct materials found in the description (e.g., "Cement, Sand, Stone aggregate").
- "material_category": (String) Classify into a broad category: "Concrete", "Earthwork", "Steel", "Timber", "Masonry", "Paint/Finish", "Other".
- "material_confidence": (Number) A score from 0.0 to 1.0 reflecting your confidence in the material classification.
- "mix_ratio": (String) Any mix ratio mentioned like "1:2:4" or "1:6". Null if not found.
- "original_quantity": (Number) The extracted quantity.
- "original_unit": (String) The printed unit.
- "thickness_mm": (Number) Convert any mentioned thickness to mm (e.g., "40 mm thick" -> 40, "7 cm thick" -> 70). Null if none.
- "diameter_mm": (Number) Convert any mentioned diameter to mm (e.g., "100 mm dia" -> 100). Null if none.
- "unit_rate": (Number) The rate if printed, else null.
- "total_cost": (Number) The amount if printed, else null.
- "schedule_item_code": (String) The DSR/SOR reference code on the far right.

Return ONLY a JSON array of these objects without markdown blocks.
"""

def get_client() -> genai.Client:
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise EnvironmentError("GEMINI_API_KEY not set.")
    return genai.Client(api_key=api_key)

def render_review_images(pdf_path: str, review_dir: Path):
    """Render each page to a PNG locally for visual cross-checking."""
    doc = fitz.open(pdf_path)
    zoom = DPI / 72
    matrix = fitz.Matrix(zoom, zoom)
    for i, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=matrix)
        pix.save(str(review_dir / f"page_{i:02d}.png"))
    n_pages = doc.page_count
    doc.close()
    return n_pages

def extract_whole_pdf(client: genai.Client, pdf_path: str) -> list[dict]:
    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = client.models.generate_content(
                model=MODEL,
                contents=[
                    types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                    EXTRACTION_PROMPT,
                ],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.1,
                ),
            )
            return json.loads(response.text)
        except genai_errors.ClientError as e:
            if "RESOURCE_EXHAUSTED" in str(e) or "429" in str(e):
                print(f"\n  QUOTA EXHAUSTED: {e}")
                return []
            raise
        except (genai_errors.ServerError, json.JSONDecodeError):
            print(f"Server busy, retrying attempt {attempt}/{MAX_RETRIES}...")
            time.sleep(RETRY_BASE_DELAY)
    return []

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf", required=True)
    parser.add_argument("--out", default="output")
    args = parser.parse_args()

    out_dir = Path(args.out)
    review_dir = out_dir / "ocr_review"
    review_dir.mkdir(parents=True, exist_ok=True)
    
    print("Rendering page images locally via PyMuPDF...")
    n_pages = render_review_images(args.pdf, review_dir)
    print(f"  {n_pages} page(s) rendered to {review_dir}")
    
    client = get_client()
    print(f"\nSending PDF to Gemini for smart extraction (handling sub-items)...")
    all_items = extract_whole_pdf(client, args.pdf)

    combined_path = out_dir / "boq_extracted.json"
    combined_path.write_text(json.dumps(all_items, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"\nDone. {len(all_items)} line item(s) extracted.")
    print(f"-> Combined JSON saved to: {combined_path}")
    print(f"-> Local review images saved to: {review_dir}")

if __name__ == "__main__":
    main()