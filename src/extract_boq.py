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
You are a highly precise Data Extraction AI reading a scanned Bill of Quantities
(BoQ) for a civil construction project in India. The scan may be an old
dot-matrix printout, so read carefully.

CRITICAL ANTI-SHIFTING RULES (MUST FOLLOW):
1. HORIZONTAL TRACKING: Read strictly row by row, left to right. Never mix data between rows.
2. MISSING VALUES: If a column (like Quantity or Schedule Code) is blank, illegible, or spans multiple lines, you MUST output `null` for that field. NEVER shift a value from row N+1 into row N.
3. DECIMAL POINTS: Use the context of the unit and surrounding numbers to infer the correct decimal placement (e.g., '7.3' instead of '73').

SUB-ITEM RULES:
1. Items with multiple sub-parts (e.g., 16 i, 16 ii, or a, b, c) MUST be split into separate JSON objects (e.g., "16a", "16b"). Extract the specific quantity for EACH sub-part accurately.

DOCUMENT-LEVEL CONTEXT (read once, apply to every item):
- Look at the page header/title block (usually repeated on every page) for the
  name of the rate schedule this BoQ is priced against, e.g. "DSR 1989",
  "CPWD DSR 2019", "UPPWD SOR". Use this exact text as "schedule_source" for
  EVERY item in the document, even pages where the header is cropped or faint
  (carry it over from a page where it IS legible).
- Track section/division headers as you move down the document. These
  headers exist at TWO nested levels and you must capture the MORE SPECIFIC
  (deeper) one, not just the top-level one:
    Level 1 (broad, changes rarely): e.g. "Schedule 'A'", "Schedule 'B'".
    Level 2 (granular, changes often): e.g. "Sub-Head - I, Earth Work",
      "Sub-Head - II, Concrete Work", "Sub-Head - III, RCC Work".
  "floor_section" MUST be the Level 2 (Sub-Head) heading in effect for that
  item, e.g. "Sub-Head - I, Earth Work" -- NOT the Level 1 heading alone.
  If a page only shows a Level 1 heading with no Sub-Head printed yet,
  carry forward the most recent Sub-Head heading you have seen, the way a
  human reading top-to-bottom would. Only fall back to the Level 1 heading
  if no Sub-Head heading has appeared anywhere in the document yet.
  Do NOT let the item number reset or a Level 1 heading repeating on every
  page trick you into re-using the same Level 1 value for every single row
  in the document -- that is a sign you have missed the Sub-Head headings.

Extract EVERY line item and return a JSON array of objects with EXACTLY these fields (use null if not found):
- "source_page": (Integer) The page number (1-indexed) this item appears on.
- "boq_item_no": (String) The printed item number.
- "description": (String) Full description text.
- "floor_section": (String) The most recent section/division header in effect for this item (see DOCUMENT-LEVEL CONTEXT above). Only null if no header has appeared anywhere yet.
- "schedule_source": (String) The rate schedule name from the page header/title (see DOCUMENT-LEVEL CONTEXT above), e.g. "DSR 1989". Same value for every item unless the document itself switches schedules partway through.
- "discipline": (String) Intelligently classify: "Civil & Sitework", "Structural", "Electrical", "Plumbing & Sanitary", "HVAC", "Finishes", "Other".
- "material_product": (String) The primary specific material mentioned (e.g., "Cement concrete", "Burnt brick", "Teak wood").
- "all_materials_detected": (String) A comma-separated list of all distinct materials found in the description (e.g., "Cement, Sand, Stone aggregate").
- "material_category": (String) Classify into a broad category: "Concrete", "Earthwork", "Steel", "Timber", "Masonry", "Plaster", "Paint/Finish", "Other".
  IMPORTANT distinction: "Plaster" is for cement/sand (or lime) plastering,
  rendering, or pointing work applied to walls/ceilings/floors as a base
  layer (typically several mm to a few cm thick, described as e.g.
  "12 mm cement plaster 1:4", "15 mm thick plastering"). "Paint/Finish" is
  ONLY for actual paint, primer, distemper, white/colour washing, or
  varnish -- a thin decorative/protective coating, not a structural mortar
  layer. Do NOT put plastering or pointing items in "Paint/Finish" just
  because they are both finishing work -- they have very different
  material composition, density, and embodied carbon and must be kept
  separate.
- "material_confidence": (Number) A score from 0.0 to 1.0 reflecting your confidence in the material classification.
- "grade": (String) A named material grade ONLY if literally printed in the text (e.g., "M-15", "Fe-500D", "43 Grade OPC"). Do NOT infer or calculate a grade from a mix ratio -- leave null if no grade word is printed. Grade inference from mix ratios is handled separately downstream.
- "mix_ratio": (String) Any mix ratio mentioned like "1:2:4" or "1:6". Null if not found.
- "original_quantity": (Number) The extracted quantity exactly as printed (do not multiply or rescale it, even if the unit column says something like "100 Sq.m" or "10 Cubic decimetre" -- report the raw printed number and let "original_unit" carry the full unit text).
- "original_unit": (String) The printed unit, verbatim, including any multiplier prefix printed with it (e.g. "100 Sq.m", "10 Cubic decimetre").
- "standard_code_reference": (String) Any IS/BIS or other named standard cited in the description (e.g., "IS 456", "IS:9103"). Comma-separate if multiple. Null if none cited.
- "thickness_mm": (Number) Convert any mentioned thickness to mm (e.g., "40 mm thick" -> 40, "7 cm thick" -> 70). Null if none.
- "diameter_mm": (Number) Convert any mentioned diameter to mm (e.g., "100 mm dia" -> 100). Null if none.
- "length_mm": (Number) Convert any explicit length dimension to mm if separately stated (e.g. panel/unit length, not the BoQ quantity). Null if none.
- "width_mm": (Number) Convert any explicit width dimension to mm if stated. Null if none.
- "height_mm": (Number) Convert any explicit height dimension to mm if stated. Null if none.
- "depth_mm": (Number) Convert any explicit depth dimension to mm if stated (distinct from thickness). Null if none.
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
