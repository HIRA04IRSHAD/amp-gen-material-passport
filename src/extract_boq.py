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

# Final prompt: Kept source_page for manual review, removed notes/confidence, added strict sub-item rules
EXTRACTION_PROMPT = """
You are reading a scanned, dot-matrix-printed Bill of Quantities (BoQ)
for a civil construction project in India (DSR 1989 style schedule).
The scan quality is poor, so read carefully and use context/column
alignment to resolve ambiguous characters or digits.

CRITICAL INSTRUCTION FOR SUB-ITEMS & QUANTITIES:
1. Sub-items: Items like 16, 17, 31, 32, 34, 51 have multiple sub-parts (e.g., a, b, c or i, ii, iii). 
   DO NOT merge them. You MUST split them into separate JSON objects. 
   Give them item numbers like "16a", "16b". Extract the specific quantity, unit, and description for EACH sub-part accurately.
2. Quantities: If a quantity is missing or illegible, set "original_quantity" to null.

Extract EVERY line item (numbered row) visible across ALL pages. 
For each item/sub-item, return a JSON object with EXACTLY these fields (do not add any extra fields):

- "source_page": the page number (1-indexed) this item appears on.
- "boq_item_no": the item number (e.g. "1", "16a"). Null if not present.
- "description": the full item description text for this specific row/sub-row.
- "discipline": guess one of: "Civil & Sitework", "Structural", "Electrical", "Plumbing & Sanitary", "HVAC", "Finishes", "Other".
- "material_category": generic category (e.g. "Concrete", "Masonry", "Earthwork", "Steel", "Timber", "Paint/Finish", "Other").
- "original_quantity": the quantity as a NUMBER (no commas/units), or null.
- "original_unit": The unit of measurement (e.g., "Cu.m", "Sq.m").
- "schedule_item_code": DSR/SOR reference code if present, else null.

Return ONLY a JSON array of these objects.
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