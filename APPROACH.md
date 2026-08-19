# Approach

## What tools I picked and why

TODO. e.g.:
- Manually transcribed the BoQ scan (dot-matrix print, 64 items) rather than
  running OCR — the scan quality made OCR error-prone for numeric columns,
  and the item count is small enough that manual transcription was faster
  and more reliable than fixing OCR errors.
- openpyxl to read/write the xlsx template directly, preserving the existing
  colour-coded formatting.
- matplotlib for the material distribution chart (static PNG, simplest to
  review without a server).

## What worked

TODO.

## What did not work / was skipped

TODO. e.g. AMBER (Density / Carbon / GWP) columns left blank except for the
5+ filled for bonus B2, since most DSR 1989 items don't map cleanly to a
single EPD without assumptions I wasn't confident citing.

## What I'd do with two more weeks

TODO. e.g.:
- Build a proper OCR + LLM extraction pipeline instead of manual
  transcription, so it generalises to other BoQ scans.
- Fill AMBER columns for all materials with a proper EPD lookup table
  instead of just 5.
- Add validation (unit conversion sanity checks, DSR code cross-reference).
- Deploy a small Streamlit app for browsing the passport (bonus B1).
