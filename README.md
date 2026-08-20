# AMP-GEN Material Passport

> **Take-Home Task — Material Passport Extraction & Visualisation**

This project extracts **64 Bill of Quantities (BoQ) line items** from a scanned Bill of Quantities for **CBRI Roorkee : Principal's Residence (DSR 1989)** and maps them into the **AMP-GEN Material Passport template**.

The pipeline also exports the extracted data as JSON and generates a visualisation showing the distribution of materials across the building.

---

## Project Structure

```text
amp-gen-material-passport/
│
├── src/
│   ├── extract_boq.py          # Extracts and structures BoQ items
│   ├── build_passport.py       # Builds the Excel passport + JSON
│   └── visualize.py            # Generates material distribution chart
│
├── output/
│   ├── passport_filled.xlsx    # Completed AMP-GEN Material Passport
│   ├── passport.json           # Extracted data in JSON format
│   ├── visualization.png       # Material distribution visualisation
│   └── building_meta.json      # Building metadata (Bonus B3)
│
├── APPROACH.md                 # Methodology, assumptions & judgment calls
├── requirements.txt            # Python dependencies
└── README.md
```

---

## How to Run

The complete pipeline can be executed in **under 5 minutes**.

### 1. Clone the repository

```bash
git clone https://github.com/HIRA04IRSHAD/amp-gen-material-passport.git
cd amp-gen-material-passport
```

### 2. Install dependencies

```bash
python -m pip install -r requirements.txt
```

### 3. Run the extraction pipeline

```bash
python src/extract_boq.py
python src/build_passport.py
python src/visualize.py
```

### 4. Generated Outputs

All generated files are placed in the `output/` directory:

* `passport_filled.xlsx` : completed Material Passport
* `passport.json` : structured JSON representation of the extracted data
* `visualization.png` : material distribution visualisation
* `building_meta.json` : building metadata

---

## Tools & Technologies

| Tool / Technology           | Purpose                                                                |
| --------------------------- | ---------------------------------------------------------------------- |
| **Python 3.13+**            | Core scripting and pipeline automation                                 |
| **Google Gemini 3.6 Flash** | Multimodal extraction from degraded scanned documents                  |
| **PyMuPDF (`fitz`)**        | PDF page rendering and image preprocessing                             |
| **OpenPyXL**                | Populating the Excel template while preserving styles and merged cells |
| **Pandas**                  | Data structuring and transformation                                    |
| **Seaborn**                 | Material distribution visualisation                                    |

### OCR / Vision Approach

The source document is a **low-contrast dot-matrix architectural scan**, making conventional OCR unreliable for preserving the spatial relationship between columns.

A multimodal vision LLM was therefore used to interpret the scanned pages while maintaining strict **horizontal tracking of BoQ columns and rows**.

---

## Extraction Results

### BoQ Coverage

**64 / 64 primary BoQ items successfully extracted and verified.**

During extraction, additional sub-items were identified, resulting in:

**74 / 74 total items successfully extracted.**

The extracted records were validated before being written to the final Excel and JSON outputs.

---

## Bonuses

| Bonus                                | Status | Details                               |
| ------------------------------------ | :----: | ------------------------------------- |
| **B1 — Live Deployment**             |    ⬜   | Not attempted                         |
| **B2 — AMBER Mass & Carbon Columns** |    ✅   | Density + mass-based Embodied Carbon for 6 materials (Concrete, Steel, Masonry, Timber, Plaster, Paint/Finish) with EPD citations |
| **B3 — Building Metadata**           |    ✅   | `building_meta.json` included         |
| **B4 — 3-Minute Walkthrough Video**  |    ⬜   | Not attempted                         |

---

## ⏱Time Spent

Approximately **5 hours** of focused development.

The majority of the time was spent on:

* Prompt engineering for degraded document extraction
* Resolving OCR / vision-based spatial alignment issues
* Handling API rate limits
* Mapping extracted fields to Excel coordinates
* Preserving the structure and formatting of the provided template
* Validating extracted JSON before final export

---

## Key Challenges & Resolution

The primary challenge was the **low-contrast dot-matrix printing** in the source BoQ.

Around **Item 55**, minor spatial shifting between columns initially caused alignment issues.

This was resolved by:

1. Refining the extraction prompts with explicit horizontal-tracking rules.
2. Enforcing structured JSON output.
3. Validating extracted records before spreadsheet generation.
4. Applying a pragmatic **Human-in-the-Loop (HITL)** validation step for ambiguous records.

This approach ensured that the final **74 extracted records** were correctly aligned and verified.

---

## Final Deliverables

The repository contains the complete extraction pipeline along with the generated deliverables:

* **Excel:** AMP-GEN Material Passport
* **JSON:** Structured representation of all extracted records
* **Visualisation:** Material distribution across the building
* **Metadata:** Building-level information
* **Documentation:** Methodology and extraction decisions in `APPROACH.md`

---

## Honest Note

The source document presented a genuine extraction challenge due to its **low-contrast dot-matrix print quality**.

Rather than relying on a single OCR pass, the workflow combined **multimodal document understanding, structured extraction, spatial alignment rules, and HITL validation**.

The minor alignment issue around Item 55 was identified and resolved during validation, resulting in a verified final dataset.

---

## Update Log

A post-submission review caught a unit bug in the Embodied Carbon
calculation (materials quantified in cum/sqm/m/nos were being scored as if
the raw quantity were already a mass in kg), which the empty Density column
was meant to prevent. Fixed by populating Density per material category and
computing carbon from an actual derived mass. Material Confidence was also
reworked to reflect data reliability (weight/volume basis, grade inference)
rather than the extraction model's self-rated score, Comments now always
name the specific material instead of repeating generic boilerplate, and a
dedicated `Plaster` category was added so cement plastering is no longer
folded into `Paint/Finish`. Full details in `APPROACH.md`, Section 5.

---

## Author

**Hira Irshad**

GitHub: [@HIRA04IRSHAD](https://github.com/HIRA04IRSHAD)
