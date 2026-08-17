# Seller Input vs AI Enrichment — QC Comparison Tool

A reusable Streamlit tool for the manual QC team. Upload the Seller/Vendor
input file and the AI-enriched output file for any batch, and it produces a
single Excel report with every deviating attribute highlighted in yellow,
plus a summary of how many SKUs/attributes are affected.

## How it works

1. **Upload both files** (Input and AI Output) — `.xlsx`, `.xls`, or `.csv`.
2. **Pick the data sheet** if a file has multiple sheets (e.g. reference/
   lookup sheets like "masterdata" are ignored by default but selectable).
3. **Confirm the matching key** — auto-detected as `styleId` when present,
   otherwise pick the right ID column from the dropdown.
4. **Compares every shared attribute by default** — including image/URL
   columns. If you want to skip specific columns, use the optional
   "Exclude specific columns" control before running.
5. Click **Run Comparison** to see the summary inline (with a color legend),
   then **download** the highlighted `.xlsx` report.

## Color legend (shown in-app and in the Excel report)

| Color | Meaning |
|---|---|
| 🟨 Yellow | **Value Mismatch** — both Input and AI Output have data, but the values differ. Needs manual review/correction. |
| 🟧 Amber | **Missing in AI Output** — the AI didn't generate this field at all. The Input value has been **automatically copied into the AI Output cell** so no data is lost — the team just needs to confirm it or replace it with a proper AI-generated value. |

## Output report structure (3 sheets, same Excel file)

- **Summary** — legend, SKUs compared, attributes compared, SKUs with at
  least one deviation, SKUs with at least one backfilled field, and a full
  attribute-by-attribute breakdown split by **Value Mismatch** count vs
  **Missing in AI Output** count, sorted by total deviations.
- **Input Data** — the Input file's full original column structure and
  values, unchanged, for every matched SKU. Nothing is dropped or
  reordered. Cells that mismatch the AI Output are highlighted yellow.
- **AI Output Data** — the Output file's full original column structure
  and values, unchanged, for every matched SKU. Workflow-only columns
  unique to the Output file (e.g. `Actions`, `AM`, `Detected Taxonomy`,
  `SKU Id`) pass through as-is with no highlight since there's nothing in
  the Input to compare them against. Every attribute shared with the
  Input — including image/URL fields — is compared: mismatches are
  highlighted yellow, and cells the AI left blank that were auto-filled
  from the Input are highlighted amber.

Both data sheets mirror their source file's own structure exactly — the
only difference from the original files is the highlighting layered on top
and, in the AI Output sheet, blank cells being filled in from Input where
the AI produced nothing.

SKUs present in only one of the two files are reported separately in the
Summary sheet and excluded from the row-by-row comparison (since there's
nothing to compare against).

## Running it

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (typically `http://localhost:8501`)
in a browser. Share this on an internal server so the whole QC team can
access it via a shared URL, or run it locally per-user.

## Notes / assumptions

- Comparison ignores case and extra whitespace by default (toggleable) —
  e.g. `"Red "` vs `"red"` is not flagged as a deviation unless you turn
  this off.
- Image/URL columns (Front Image, Back Image, BIS Certificate Image URL,
  etc.) are **compared like any other attribute by default**. If you want
  to skip specific columns, use the "Exclude specific columns" multiselect
  in the app before running the comparison.
- The key column is normalized so `39700559` (number) and `"39700559"`
  (text) are treated as the same SKU.
