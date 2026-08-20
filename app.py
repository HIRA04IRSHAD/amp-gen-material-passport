"""
app.py
End-to-end Streamlit app for the AMP-GEN Material Passport pipeline.

Two modes (tabs):
  1. "Run New Extraction" — upload a scanned BoQ PDF, the app runs the
     existing pipeline (src/extract_boq.py -> src/build_passport.py ->
     src/visualize.py) against it, then shows a preview + download
     buttons for passport_filled.xlsx / passport.json / building_meta.json
     / visualization.png.
  2. "View Bundled Demo Data" — browses the already-generated
     output/ folder (CBRI Roorkee sample) without needing an API key,
     so the app is still useful/demoable even without Gemini access.

Extraction requires a Gemini API key (GEMINI_API_KEY). On Streamlit
Community Cloud, set it under App settings -> Secrets as:
    GEMINI_API_KEY = "your-key-here"
Users without that configured can paste their own key in the sidebar
for their session only (never persisted, never logged).
"""

import json
import sys
import tempfile
import threading
from pathlib import Path

import pandas as pd
import streamlit as st
import plotly.express as px

ROOT = Path(__file__).parent
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

import extract_boq          # noqa: E402  (src/extract_boq.py)
import build_passport        # noqa: E402  (src/build_passport.py)
import visualize as viz_mod  # noqa: E402  (src/visualize.py)

TEMPLATE_EXCEL = ROOT / "AMP_Passport_Template.xlsx"
DEMO_OUTPUT = ROOT / "output"

# All GREEN (required) columns from the template, in template order, mapped
# to the field name they end up under in passport.json. Note: "Article
# Number", "External DB Id", and "Currency" are also green in the template
# but the pipeline never populates them (not extracted from the BoQ, no
# multi-currency handling), so they're left out here rather than showing an
# always-empty column.
GREEN_COLUMNS = [
    ("gmap_id", "GMAP Id"),
    ("boq_item_no", "BOQ Item No."),
    ("description", "Description"),
    ("floor_section", "Floor / Section"),
    ("discipline", "Discipline"),
    ("material_product", "Material / Product"),
    ("all_materials_detected", "All Materials Detected"),
    ("material_category", "Material Category"),
    ("material_confidence", "Material Confidence"),
    ("grade", "Grade"),
    ("mix_ratio", "Mix Ratio"),
    ("original_quantity", "Original Quantity"),
    ("original_unit", "Original Unit"),
    ("volume_cum", "Volume (m3)"),
    ("area_sqm", "Area (m2)"),
    ("length_m", "Length (m)"),
    ("weight_kg", "Weight (kg)"),
    ("count_nos", "Count (Nos)"),
    ("derived_qty_value", "Derived Quantity"),
    ("derived_qty_unit", "Derived Quantity Unit"),
    ("derived_qty_basis", "Derived Quantity Basis"),
    ("schedule_source", "Schedule (DSR/SOR)"),
    ("schedule_item_code", "Schedule Item Code"),
    ("standard_code_reference", "Standard / Code Reference"),
    ("classification_matched", "Classification (Matched)"),
    ("length_mm", "Length (mm)"),
    ("width_mm", "Width (mm)"),
    ("height_mm", "Height (mm)"),
    ("thickness_mm", "Thickness (mm)"),
    ("depth_mm", "Depth (mm)"),
    ("diameter_mm", "Diameter (mm)"),
    ("unit_rate", "Unit Rate"),
    ("total_cost", "Total Cost"),
    ("comment", "Comment"),
]


def _flatten_passport_records(data):
    """passport.json stores volume/area/length/weight/count under a nested
    routed_quantities dict, and derived quantity under a nested
    derived_quantity dict. Pull those up to top-level keys so they show up
    as normal DataFrame columns instead of disappearing into an object col."""
    flat = []
    for item in data:
        row = dict(item)
        rq = row.pop("routed_quantities", {}) or {}
        row["volume_cum"] = rq.get("volume_cum", "")
        row["area_sqm"] = rq.get("area_sqm", "")
        row["length_m"] = rq.get("length_m", "")
        row["weight_kg"] = rq.get("weight_kg", "")
        row["count_nos"] = rq.get("count_nos", "")
        dq = row.pop("derived_quantity", {}) or {}
        row["derived_qty_value"] = dq.get("value", "")
        row["derived_qty_unit"] = dq.get("unit", "")
        row["derived_qty_basis"] = dq.get("basis", "")
        flat.append(row)
    return flat

st.set_page_config(page_title="AMP-GEN Material Passport", page_icon="🏗️", layout="wide")

# A single pipeline run touches shared module-level globals in the src/
# scripts (they weren't written to be called concurrently), so serialize
# runs across sessions with a lock. Fine for a low-traffic demo/review app.
_PIPELINE_LOCK = threading.Lock()


# ---------------------------------------------------------------------------
# Pipeline runner
# ---------------------------------------------------------------------------
def run_pipeline(pdf_bytes: bytes, api_key: str, progress_cb=None):
    """Runs extract -> build -> visualize against an uploaded PDF in an
    isolated temp directory. Returns a dict of output file paths (or raises)."""
    work_dir = Path(tempfile.mkdtemp(prefix="ampgen_"))
    pdf_path = work_dir / "uploaded_boq.pdf"
    pdf_path.write_bytes(pdf_bytes)
    out_dir = work_dir / "output"
    review_dir = out_dir / "ocr_review"
    review_dir.mkdir(parents=True, exist_ok=True)

    def report(msg):
        if progress_cb:
            progress_cb(msg)

    with _PIPELINE_LOCK:
        # --- Step 1: render pages for review + call Gemini for extraction ---
        report("Rendering PDF pages...")
        extract_boq.render_review_images(str(pdf_path), review_dir)

        report("Extracting line items via Gemini (this can take 1-3 minutes)...")
        import os
        os.environ["GEMINI_API_KEY"] = api_key
        client = extract_boq.get_client()
        items = extract_boq.extract_whole_pdf(client, str(pdf_path))

        if not items:
            raise RuntimeError(
                "Extraction returned no items. This usually means the "
                "Gemini API key is invalid/exhausted, or the model call "
                "failed. Check the key and try again."
            )

        boq_extracted_path = out_dir / "boq_extracted.json"
        boq_extracted_path.write_text(json.dumps(items, indent=2, ensure_ascii=False), encoding="utf-8")

        # --- Step 2: build passport Excel + JSON (monkeypatch module paths) ---
        report("Building Material Passport (Excel + JSON)...")
        build_passport.INPUT_JSON = boq_extracted_path
        build_passport.TEMPLATE_EXCEL = TEMPLATE_EXCEL
        build_passport.OUTPUT_EXCEL = out_dir / "passport_filled.xlsx"
        build_passport.OUTPUT_JSON = out_dir / "passport.json"
        build_passport.OUTPUT_META = out_dir / "building_meta.json"
        build_passport.main()

        # --- Step 3: visualization ---
        report("Generating visualization...")
        viz_mod.INPUT_JSON = out_dir / "passport.json"
        viz_mod.OUTPUT_PNG = out_dir / "visualization.png"
        viz_mod.make_chart()

    report("Done.")
    return {
        "excel": out_dir / "passport_filled.xlsx",
        "json": out_dir / "passport.json",
        "meta": out_dir / "building_meta.json",
        "png": out_dir / "visualization.png",
        "work_dir": work_dir,
    }


# ---------------------------------------------------------------------------
# Shared rendering helpers
# ---------------------------------------------------------------------------
def render_results(paths: dict, key_prefix: str = "run"):
    passport_json = paths["json"]
    if not passport_json.exists():
        st.error("passport.json was not generated. Check the log above.")
        return

    with open(passport_json, "r", encoding="utf-8") as f:
        data = json.load(f)
    df = pd.DataFrame(_flatten_passport_records(data))
    if "material_category" in df.columns:
        df["material_category"] = df["material_category"].replace("", "Not Specified").fillna("Not Specified")

    meta = {}
    if paths["meta"].exists():
        meta = json.loads(paths["meta"].read_text(encoding="utf-8"))

    if meta:
        st.subheader("Building Metadata")
        cols = st.columns(4)
        for i, (k, v) in enumerate(meta.items()):
            cols[i % 4].metric(k.replace("_", " "), v)
        st.divider()

    k1, k2, k3 = st.columns(3)
    k1.metric("Total Items Extracted", len(df))
    k2.metric("Material Categories", df["material_category"].nunique() if "material_category" in df.columns else 0)
    if "embodied_carbon" in df.columns:
        total_carbon = pd.to_numeric(df["embodied_carbon"], errors="coerce").sum()
        k3.metric("Total Embodied Carbon (kg CO2e)", f"{total_carbon:,.0f}")

    if "material_category" in df.columns:
        cat_counts = df["material_category"].value_counts().reset_index()
        cat_counts.columns = ["Material Category", "Count"]
        fig = px.bar(
            cat_counts.sort_values("Count"), x="Count", y="Material Category", orientation="h",
            title="Material Distribution Across Building (by Category)",
            color="Count", color_continuous_scale="viridis",
        )
        fig.update_layout(coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True, key=f"{key_prefix}_chart")

    st.subheader("Extracted Line Items (all required / green columns)")
    display_pairs = [(key, label) for key, label in GREEN_COLUMNS if key in df.columns]
    display_cols = [key for key, _ in display_pairs]
    display_df = df[display_cols].rename(columns=dict(display_pairs))
    st.dataframe(display_df, use_container_width=True, height=420, key=f"{key_prefix}_table")

    st.subheader("Downloads")
    c1, c2, c3, c4 = st.columns(4)
    if paths["excel"].exists():
        with open(paths["excel"], "rb") as f:
            c1.download_button("⬇️ passport_filled.xlsx", f, file_name="passport_filled.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"{key_prefix}_dl_excel")
    if paths["json"].exists():
        with open(paths["json"], "rb") as f:
            c2.download_button("⬇️ passport.json", f, file_name="passport.json", mime="application/json",
                                key=f"{key_prefix}_dl_json")
    if paths["meta"].exists():
        with open(paths["meta"], "rb") as f:
            c3.download_button("⬇️ building_meta.json", f, file_name="building_meta.json", mime="application/json",
                                key=f"{key_prefix}_dl_meta")
    if paths["png"].exists():
        with open(paths["png"], "rb") as f:
            c4.download_button("⬇️ visualization.png", f, file_name="visualization.png", mime="image/png",
                                key=f"{key_prefix}_dl_png")


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("🏗️ AMP-GEN Material Passport")
st.caption("Upload a scanned BoQ PDF → extraction → Material Passport (Excel + JSON) → visualization.")

tab_run, tab_demo = st.tabs(["🚀 Run New Extraction", "📊 View Bundled Demo Data"])

with tab_run:
    st.sidebar.header("Gemini API Key")
    try:
        secret_key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        # No secrets.toml file at all (e.g. local dev without one configured)
        secret_key = ""
    if secret_key:
        st.sidebar.success("Using API key from app secrets.")
        api_key = secret_key
    else:
        api_key = st.sidebar.text_input(
            "Enter your GEMINI_API_KEY", type="password",
            help="Used only for this session, never stored or logged.",
        )

    uploaded_pdf = st.file_uploader("Upload a scanned BoQ PDF", type=["pdf"])
    run_clicked = st.button("Run extraction pipeline", type="primary", disabled=not uploaded_pdf)

    if run_clicked:
        if not api_key:
            st.error("Please provide a Gemini API key in the sidebar first.")
        else:
            status = st.empty()
            try:
                with st.spinner("Running pipeline..."):
                    result_paths = run_pipeline(
                        uploaded_pdf.getvalue(), api_key,
                        progress_cb=lambda m: status.info(m),
                    )
                status.success("Pipeline complete.")
                st.session_state["result_paths"] = {k: str(v) for k, v in result_paths.items()}
            except Exception as e:
                status.error(f"Pipeline failed: {e}")

    if "result_paths" in st.session_state:
        st.divider()
        render_results({k: Path(v) for k, v in st.session_state["result_paths"].items()}, key_prefix="run")

with tab_demo:
    st.caption("Pre-generated output for CBRI Roorkee — Principal's Residence (DSR 1989). No API key needed.")
    render_results({
        "excel": DEMO_OUTPUT / "passport_filled.xlsx",
        "json": DEMO_OUTPUT / "passport.json",
        "meta": DEMO_OUTPUT / "building_meta.json",
        "png": DEMO_OUTPUT / "visualization.png",
    }, key_prefix="demo")