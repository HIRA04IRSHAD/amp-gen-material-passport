# AMP-GEN Material Passport — Take-Home Task

Extracts 64 BoQ line items from a scanned Bill of Quantities (CBRI Roorkee,
Principal's Residence, DSR 1989) into the AMP-GEN Material Passport template,
exports the same data as JSON, and produces one visualisation of material
distribution across the building.

## What's here

src/                    extraction + build + visualisation code
output/
passport_filled.xlsx  filled template
passport.json         same data, one record per row
visualization.png     material distribution chart
building_meta.json    building metadata block (bonus B3)
APPROACH.md             tools, judgment calls, what I'd do with more time

## How to run (under 5 minutes)

1. Clone the repo and enter it:

git clone https://github.com/HIRA04IRSHAD/amp-gen-material-passport.git

2. Install dependencies:

python -m pip install -r requirements.txt

3. Run the pipeline:

python src/extract_boq.py         # parses BoQ items into structured data
python src/build_passport.py     # fills the xlsx template + exports JSON
python src/visualize.py          # produces output/visualization.png

All outputs land in `output/`.

## Tools / LLMs / OCR used

- **Python (v3.13+)**: Core script execution and automation pipeline.
- **Google Gemini 3.6 Flash**: Multimodal vision LLM selected for reading degraded dot-matrix architectural scans.
- **PyMuPDF (`fitz`)**: Pure Python library for native, dependency-free PDF-to-image page rendering.
- **Openpyxl**: For programmatically injecting data into the master Excel template while preserving styles and merged cells.
- **Pandas & Seaborn**: For backend data structuring and generating polished material distribution charts.

## Hours actually spent

~5.0 Hours (Focused on robust prompt engineering, resolving OCR dot-matrix spatial alignment, API rate-limit management, and openpyxl coordinate mapping).

## Items extracted

64 / 64 items successfully extracted and verified.
-found sub items therefore total 74/74 items successfully extracted.

## Bonuses attempted

- [ ] B1 — Live deployment
- [x] B2 — AMBER mass & carbon columns (5+ materials, cited)
- [x] B3 — building_meta.json
- [ ] B4 — 3-minute walkthrough video

## Honest note

The low-contrast dot-matrix print caused minor spatial shifting around Item 55; this was successfully resolved by refining extraction prompts with strict horizontal-tracking rules and applying a pragmatic Human-in-the-Loop (HITL) JSON validation layer.