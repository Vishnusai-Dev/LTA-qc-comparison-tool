# Rubick QC & Content Tools

A reusable Streamlit app with two tools, switchable from the sidebar:

1. **QC Comparison** — compares Seller/Vendor Input vs AI Enrichment Output
   per SKU and highlights deviations for manual review.
2. **Generate Content** — generates DN, LVN, PD, and Size & Fit text per
   SKU from category-specific nomenclature rules (Backpacks, Trolley Bag,
   Handbags, Laptop Bag), and writes the results back into your file.

---

## Tool 1: QC Comparison

### How it works

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

### Color legend (shown in-app and in the Excel report)

| Color | Meaning |
|---|---|
| 🟨 Yellow | **Value Mismatch** — both Input and AI Output have data, but the values differ. Needs manual review/correction. |
| 🟧 Amber | **Missing in AI Output** — the AI didn't generate this field at all. The Input value has been **automatically copied into the AI Output cell** so no data is lost — the team just needs to confirm it or replace it with a proper AI-generated value. |

### Output report structure (3 sheets, same Excel file)

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

### Notes / assumptions

- Comparison ignores case and extra whitespace by default (toggleable) —
  e.g. `"Red "` vs `"red"` is not flagged as a deviation unless you turn
  this off.
- Image/URL columns (Front Image, Back Image, BIS Certificate Image URL,
  etc.) are **compared like any other attribute by default**. If you want
  to skip specific columns, use the "Exclude specific columns" multiselect
  in the app before running the comparison.
- The key column is normalized so `39700559` (number) and `"39700559"`
  (text) are treated as the same SKU.

---

## Tool 2: Generate Content

Generates listing content for four categories — **Backpacks, Trolley Bag,
Handbags, Laptop Bag** — following the team's nomenclature rules:

- **DN** (Display Name) → written to `productDisplayName`
- **LVN** (List View Name) → written to `listViewName`
- **PD** (Product Details, bullet-style) → written to `Product Details`
- **Size & Fit** → written to `sizeAndFitDescription`

### How it works

1. **Choose a category.**
2. **Upload your file** (one row per SKU, attribute columns as headers —
   the same kind of file used for QC). Pick a sheet if there are multiple.
3. **Confirm field mapping** — the app auto-guesses which of your columns
   correspond to each nomenclature field (e.g. "Base Colour" → `Prominent
   Colour` if that's what your file calls it) and shows every guess for
   you to review/correct via dropdowns before generating.
4. **Confirm target columns** — defaults to `productDisplayName`,
   `listViewName`, `Product Details`, `sizeAndFitDescription` if present
   in your file; created automatically if missing.
5. Optionally set a **max LVN length** — if set, lowest-priority fields
   (per the nomenclature's stated preference order) are dropped first
   until LVN fits.
6. Click **Generate Content**. Generated cells are highlighted green in
   the downloaded file so the team can see what changed at a glance.
   Fields with no source data are simply omitted from the generated text
   rather than printed blank.

### Nomenclature rules implemented

Each category has its own DN/LVN word order, PD bullet list, and Size &
Fit fields, transcribed directly from the nomenclature the team supplied
(see `content_generator.py` — `CATEGORY_FIELDS`, `build_dn`, `build_lvn`,
`build_pd`, `build_size_fit`). A few points worth knowing:

- Fields marked "(Not to be mentioned if value is NA)" in the source
  nomenclature — and in fact **every** field — is dropped from PD/Size &
  Fit if blank, rather than printed with an empty value.
- The "Warranty provided by brand owner/manufacturer" disclaimer line is
  always appended at the end of PD, per every category's nomenclature.
- Colour is combined as `Base- Colour1 & Colour2` (only the parts that
  have data).
- Trolley Bag's "Pockets" line combines External + Internal pocket counts
  into one line.
- Handbags' "Type" PD line combines Surface Styling + Type.
- Laptop Bag's Size & Fit includes a fixed-phrase line: `Laptop sleeve can
  hold "X" laptop`, using the Laptop Size field.
- LVN truncation (if you set a max length) drops fields in the *reverse*
  of the nomenclature's stated preference order — lowest priority first.

Because real column headers can vary slightly from the nomenclature's
literal labels (e.g. "Base Colour" vs. "Prominent Colour" in one sample
file), **always review the field mapping step before generating** — it's
the single most important check, since it drives every SKU's output.

---

## Running it

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (typically `http://localhost:8501`)
in a browser. Share this on an internal server so the whole QC team can
access it via a shared URL, or run it locally per-user.
