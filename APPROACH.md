# Engineering Approach & Methodology : AMP-GEN Material Passport

## 1. Objective
Build an automated pipeline to parse a low-contrast, dot-matrix scanned BoQ (`BoQ_CBRI_Principals_Residence.pdf`), structure line items into JSON, normalize units/quantities, and inject them into the official `AMP_Passport_Template.xlsx` without breaking its styling, merged cells, or example rows.

---

## 2. Tech Stack
* **Python 3.13+** : pipeline scripting.
* **Google Gemini 3.6 Flash** : multimodal extraction; handles degraded dot-matrix typography and layout context better than traditional OCR.
* **PyMuPDF (`fitz`)** : in-memory PDF→PNG rendering, no external binary deps.
* **Pandas / Openpyxl** : Pandas for data shaping; `openpyxl` for cell-level template injection so styling, merges, and formulas survive.
* **Seaborn + Matplotlib** : material-distribution chart (`visualization.png`).

---

## 3. Key Issues & Resolutions

| Issue | Resolution |
| :--- | :--- |
| Gemini API `404`s from model deprecation, legacy SDK warnings | Migrated to the `google.genai` SDK, pinned to `gemini-3.6-flash`. |
| HTTP 429 quota exhaustion from page-by-page requests | Batched all rendered pages into a single multi-image request (13 calls → 1). |
| Spatial/column shifting on noisy dot-matrix scans (drift near Item 55) | Added strict horizontal row-tracking + mandatory `null`-on-missing rules to the prompt, plus a HITL review pass on the intermediate JSON. |
| `pandas.to_excel()` flattened the template (lost colors, merges, Instructions sheet) | Switched to `openpyxl.load_workbook()`, writing cell-by-cell from row 7 onward, preserving all template formatting. |
| `passport.json` path was defined but never written | Added the missing `json.dump` serialization block with routed quantities (volume/area/length/weight/count). |
| Raw Matplotlib charts looked unpolished | Migrated to Seaborn (`whitegrid` theme, `viridis` palette, bar annotations). |
| Broken pip launcher on Windows (Python 3.13 path mismatch) | Used `python -m pip install ...` instead of the wrapper executable. |

---

## 4. Post-Review Refinements (Density, Carbon Accuracy, Confidence, Plaster)

A later review of `passport_filled.xlsx` found the AMBER columns weren't actually reliable. Fixed in `build_passport.py` / `extract_boq.py`:

* **Density was never populated : which hid a unit bug in Embodied Carbon.** Carbon factors are kg CO2e *per kg*, but carbon was computed as `raw_qty × factor` regardless of unit, so 8 cum of concrete was scored as 8 kg (~1000x undercount). Fixed by adding a `DENSITY_DATABASE` (kg/m³) and deriving real mass first: direct weight, direct volume, or derived volume (area × thickness) × density. Rows with no usable basis now show blank carbon instead of a fabricated number.
* **Material Confidence was the LLM's self-rated score**, which skewed almost everything "High" regardless of data quality. Replaced with a deterministic rule: High = mass from direct weight/volume, Medium = mass derived via thickness or grade inferred from mix ratio, Low = no EPD match or no usable quantity.
* **Comments repeated identical boilerplate** (e.g. the same sentence on all 19 "Other" rows) with no way to tell rows apart. Now every comment names the specific material and shows the mass/basis/math behind the number; Paint/Finish (a generic, non-EPD-matched factor) is tagged `[ASSUMED]` instead of a misleading `[OK]`.
* **Cement plaster was falling under `Paint/Finish`** (no dedicated category existed), inheriting the wrong density and factor. Added a `"Plaster"` category to the extraction prompt (with explicit plaster-vs-paint guidance) and a matching EPD/density entry (1900 kg/m³, 0.163 kgCO2e/kg). Cut affected rows' carbon by ~4x and moved them to `[OK]`.

---

## 5. Future Scalability (2 More Weeks)

* **Hybrid OCR + bounding-box parsing** (Azure Document Intelligence / Textract, or PaddleOCR/LayoutLMv3) to eliminate spatial-shift risk before the LLM pass.
* **Interactive web audit dashboard** (Streamlit, or FastAPI+React) for side-by-side PDF/passport review and real-time HITL corrections instead of CLI-only scripts.
* **Direct LCA database integration** (ICE or a materials API) to compute A1–A3 carbon automatically from normalized quantities and densities, rather than a locally-maintained factor table.
