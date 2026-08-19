# Engineering Approach & Methodology — AMP-GEN Material Passport

## 1. Executive Summary & Core Objective
The objective of this assignment was to engineer an automated, robust data extraction and transformation pipeline capable of parsing a low-contrast, dot-matrix scanned Bill of Quantities (BoQ) PDF (`BoQ_CBRI_Principals_Residence.pdf`), structuring the line items into a standardized JSON format, normalizing architectural units and edge-case quantities, and programmatically injecting the parsed data into an official Master Excel template (`AMP_Passport_Template.xlsx`) without breaking template styling, merged cells, or pre-filled example headers.

---

## 2. Technology Stack & Tooling Rationale
* **Python (v3.13+):** Chosen for its extensive data manipulation ecosystem and scriptability.
* **Google Gemini 3.6 Flash:** Selected as the multimodal vision LLM engine. Unlike traditional OCR engines (like local Tesseract), Gemini excels at contextual layout understanding, semantic discipline classification, and handling degraded dot-matrix typography.
* **PyMuPDF (`fitz`):** Utilized for native, dependency-free PDF-to-image page rendering (`.png` buffer generation) directly in memory, avoiding external binary dependencies like Poppler (`pdf2image`).
* **Pandas & Openpyxl:** Pandas powered backend data manipulation and structural validation, while `openpyxl` was strictly used for cell-level Excel template injection to safely preserve master styling, formulas, and merged layout blocks.
* **Matplotlib & Seaborn:** Deployed to aggregate building materials by category and generate professional visualization artifacts (`visualization.png`).

---

## 3. What Broke / Did Not Work & Technical Solutions

* **Model Deprecation Cascades & SDK Migration:**
  * *The Issue:* My initial API calls to Gemini threw `404` and deprecation errors as older model endpoints were phased out. Simultaneously, the terminal flagged a `FutureWarning` indicating that the legacy `google.generativeai` package was deprecated.
  * *Resolution:* I paused development, uninstalled the legacy library, refactored the codebase, and migrated the entire pipeline to the modern `google.genai` SDK. I pinned the pipeline to `gemini-3.6-flash`, bypassing deprecated model variants and successfully handling the rapid version shifts characteristic of live cloud APIs.

* **Rate Limiting & Payload Optimization (429 Quota Exhaustion):**
  * *The Issue:* Making individual page-by-page extraction requests immediately hit free-tier rate limits (HTTP 429 - Too Many Requests) and caused cascading loop latencies.
  * *Resolution:* Instead of iterative request loops, I completely redesigned the ingestion architecture into a **single batched multi-image payload request**. Passing all rendered PyMuPDF image buffers simultaneously reduced 13 API calls down to 1 efficient request, cleanly staying within token and request per minute caps.

* **Spatial Tracking & Tabular Shifting on Dot-Matrix Scans:**
  * *The Issue:* Low-contrast, noisy dot-matrix typography caused spatial tracking slips around Item 55, where a missing code caused subsequent `schedule_item_code` values to shift out of alignment.
  * *Resolution:* Rather than accepting OCR/LLM hallucinations, I refined the `EXTRACTION_PROMPT` with strict horizontal row-by-row tracking rules and mandated `null` outputs for missing fields. To guarantee 100% data fidelity, I implemented a pragmatic **Human-in-the-Loop (HITL)** validation layer directly on the intermediate JSON before downstream processing.

* **Template Formatting Loss & Workbook Corruption:**
  * *The Issue:* Initially, using `pandas.to_excel()` to export the mapped data flattened the workbook entirely, destroying the official template's color-coded headers (Green/Amber/Grey), merged layout blocks, and the instructions sheet.
  * *Resolution:* I refactored the export logic to leverage `openpyxl.load_workbook()`. This allowed me to copy the master template directly, preserve all styling blocks and metadata in rows 1–4, and programmatically inject the parsed BoQ records strictly starting below the example rows (Row 7 onwards) with clean float formatting (`number_format = '0.00'`) to eliminate Excel green warning triangles.

* **Missing JSON Export Serialization:**
  * *The Issue:* The output path variable (`output/passport.json`) was defined in the script stub, but lacked explicit write logic, leaving the file ungenerated or blank during initial pipeline runs.
  * *Resolution:* I implemented a dedicated JSON serialization block (`json.dump`) that mirrors the routed structural dimensions (Volume, Area, Length, Weight, Count) alongside normalized units.

* **Visualization Styling & Matplotlib-to-Seaborn Migration:**
  * *The Issue:* Initial visualization attempts using raw Matplotlib produced basic, unstyled bar charts with harsh default color palettes, missing value annotations, and poor typography, making them unsuitable for a professional engineering deliverable.
  * *Resolution:* I migrated the entire visualization logic from standard Matplotlib to **Seaborn (`sns.barplot`)** layered on top of Matplotlib. This allowed me to apply clean aesthetic themes (`sns.set_theme(style="whitegrid")`), professional color palettes (`palette='viridis'`), and programmatic value labels directly onto the bars (`ax.annotate`) for maximum readability and visual polish.

---

## 4. Engineering Challenges Faced & Solutions Summary

| Challenge Encountered | Root Cause | Architectural Solution Implemented |
| :--- | :--- | :--- |
| **Broken Pip Launcher (`Fatal error in launcher`)** | Python path version mismatch (`3.13+`) on the local Windows machine where wrapper scripts failed to locate the interpreter. | Bypassed the wrapper executable by invoking package management cleanly via module-level execution (`python -m pip install ...`). |
| **API Rate Limiting (HTTP 429 - Too Many Requests)** | Iterative page-by-page LLM ingestion loops exhausting free-tier requests per minute (RPM). | Redesigned the data ingestion pipeline into a single, batched multi-image payload request, reducing 13 API calls down to 1 efficient call. |
| **Tabular Shifting & Alignment Errors on Dot-Matrix Scans** | Low-contrast, noisy dot-matrix print caused spatial tracking slips around Item 55, shifting subsequent `schedule_item_code` values. | Refined the `EXTRACTION_PROMPT` with explicit horizontal tracking constraints and introduced a pragmatic **Human-in-the-Loop (HITL)** QA validation layer on the intermediate JSON to guarantee 100% data fidelity. |
| **Unwritten `passport.json` Output File** | The output path variable was defined in the script stub, but lacked explicit file-write execution logic for routed template quantities. | Implemented a dedicated JSON serialization block (`json.dump`) mirroring the routed structural dimensions and normalized units. |

---

## 5. Future Scalability (With 2 More Weeks)

Given 2 additional weeks to scale this prototype into an enterprise-grade, production-ready material passport platform, I would implement the following architectural enhancements:

* **Hybrid OCR & Bounding-Box Layout Parsing:** 
  * *Current Limit:* Relying purely on vision LLMs can occasionally cause spatial tracking drift on noisy, low-contrast dot-matrix scans.
  * *Scale-Up:* Integrate dedicated enterprise document parsers (such as *Azure Document Intelligence* or *AWS Textract*) combined with fallback free-tier open-source engines (like *PaddleOCR* or *LayoutLMv3*). This ensures exact spatial bounding-box coordinate extraction for tabular rows *before* passing structured grids to the LLM, eliminating shifting errors entirely.

* **Full-Stack Interactive Web Audit Dashboard:** 
  * *Current Limit:* The pipeline runs via static CLI scripts (`build_passport.py`, `visualize.py`), requiring manual code execution for changes.
  * *Scale-Up:* Build a reactive web application using **Streamlit** (for rapid internal deployment) or **FastAPI + React** (for production enterprise use). This interface would allow estimators and auditors to visually inspect scanned PDF pages side-by-side with extracted material passports, perform real-time Human-in-the-Loop (HITL) corrections, and export verified digital passports instantly.

* **Automated Semantic Material Classification & LCA Database Integration:** 
  * *Current Limit:* Material categories and carbon metrics are inferred locally via prompt constraints.
  * *Scale-Up:* Connect the extraction pipeline directly to embodied carbon databases (such as *Inventory of Carbon and Energy (ICE)* or custom materials APIs). This would automatically compute Life Cycle Assessment (LCA) metrics (A1–A3 carbon footprints) based on the normalized quantities and densities extracted from the BoQ.## 5. Future Scalability (With 2 More Weeks)

Given 2 additional weeks to scale this prototype into an enterprise-grade, production-ready material passport platform, I would implement the following architectural enhancements:

* **Hybrid OCR & Bounding-Box Layout Parsing:** 
  * *Current Limit:* Relying purely on vision LLMs can occasionally cause spatial tracking drift on noisy, low-contrast dot-matrix scans.
  * *Scale-Up:* Integrate dedicated enterprise document parsers (such as *Azure Document Intelligence* or *AWS Textract*) combined with fallback free-tier open-source engines (like *PaddleOCR* or *LayoutLMv3*). This ensures exact spatial bounding-box coordinate extraction for tabular rows *before* passing structured grids to the LLM, eliminating shifting errors entirely.

* **Full-Stack Interactive Web Audit Dashboard:** 
  * *Current Limit:* The pipeline runs via static CLI scripts (`build_passport.py`, `visualize.py`), requiring manual code execution for changes.
  * *Scale-Up:* Build a reactive web application using **Streamlit** (for rapid internal deployment) or **FastAPI + React** (for production enterprise use). This interface would allow estimators and auditors to visually inspect scanned PDF pages side-by-side with extracted material passports, perform real-time Human-in-the-Loop (HITL) corrections, and export verified digital passports instantly.

* **Automated Semantic Material Classification & LCA Database Integration:** 
  * *Current Limit:* Material categories and carbon metrics are inferred locally via prompt constraints.
  * *Scale-Up:* Connect the extraction pipeline directly to embodied carbon databases (such as *Inventory of Carbon and Energy (ICE)* or custom materials APIs). This would automatically compute Life Cycle Assessment (LCA) metrics (A1–A3 carbon footprints) based on the normalized quantities and densities extracted from the BoQ.
