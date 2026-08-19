# AMP-GEN Material Passport — Take-Home Task

Extracts 64 BoQ line items from a scanned Bill of Quantities (CBRI Roorkee,
Principal's Residence, DSR 1989) into the AMP-GEN Material Passport template,
exports the same data as JSON, and produces one visualisation of material
distribution across the building.

## What's here

```
src/                    extraction + build + visualisation code
output/
  passport_filled.xlsx  filled template
  passport.json         same data, one record per row
  visualization.png     material distribution chart
  building_meta.json    building metadata block (bonus B3)
APPROACH.md             tools, judgment calls, what I'd do with more time
```

## How to run (under 5 minutes)

1. Clone the repo and enter it:
   ```
   git clone https://github.com/<username>/<repo>.git
   cd <repo>
   ```
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Run the pipeline:
   ```
   python src/extract_boq.py        # parses BoQ items into structured data
   python src/build_passport.py     # fills the xlsx template + exports JSON
   python src/visualize.py          # produces output/visualization.png
   ```
   All outputs land in `output/`.

## Tools / LLMs / OCR used

- TODO: list here (e.g. manual transcription from scan, Claude for
  extraction assistance, openpyxl for the template, matplotlib for the chart)

## Hours actually spent

TODO — fill in honestly before submitting.

## Items extracted

TODO / 64

## Bonuses attempted

- [ ] B1 — Live deployment
- [ ] B2 — AMBER mass & carbon columns (5+ materials, cited)
- [ ] B3 — building_meta.json
- [ ] B4 — 3-minute walkthrough video

## Honest note

TODO — one line: anything that broke, was skipped, or you'd like to flag.
