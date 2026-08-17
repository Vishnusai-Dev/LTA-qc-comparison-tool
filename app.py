import io
import pandas as pd
import streamlit as st

from compare_engine import (
    detect_key_column,
    detect_image_url_columns,
    common_attribute_columns,
    run_comparison,
    build_report_workbook,
)

st.set_page_config(page_title="Seller Input vs AI Enrichment — QC Tool", layout="wide")

NAVY = "#1F2A44"
ORANGE = "#E8720C"

st.markdown(
    f"""
    <div style="background-color:{NAVY};padding:18px 24px;border-radius:6px;margin-bottom:14px;">
        <h2 style="color:white;margin:0;">Seller Input vs AI Enrichment — QC Comparison Tool</h2>
        <p style="color:#D6DCE5;margin:4px 0 0 0;">
            Upload the Seller/Vendor input file and the AI-enriched output file.
            The tool matches SKUs, compares every shared attribute, and highlights
            deviations for manual review. If a field is missing in the AI Output,
            the Input value is copied in automatically so no data is lost.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div style="display:flex;gap:24px;margin-bottom:20px;font-size:14px;">
        <div style="display:flex;align-items:center;gap:8px;">
            <span style="width:16px;height:16px;background-color:#FFF2AC;border:1px solid #ccc;display:inline-block;"></span>
            <span><b>Value Mismatch</b> — both files have data, but values differ. Review and correct.</span>
        </div>
        <div style="display:flex;align-items:center;gap:8px;">
            <span style="width:16px;height:16px;background-color:#FFC97A;border:1px solid #ccc;display:inline-block;"></span>
            <span><b>Missing in AI Output</b> — AI didn't generate this field; Input value auto-filled here. Confirm or replace.</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


def load_sheets(uploaded_file):
    """Return dict of {sheet_name: DataFrame} for an uploaded xlsx/csv."""
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
        return {"Sheet1": df}
    xls = pd.ExcelFile(uploaded_file)
    return {sn: xls.parse(sn) for sn in xls.sheet_names}


def pick_sheet(label, sheets_dict, key_prefix):
    """UI for picking which sheet holds the actual SKU data."""
    sheet_names = list(sheets_dict.keys())
    if len(sheet_names) == 1:
        chosen = sheet_names[0]
    else:
        # heuristic default: sheet with the most columns (data sheets tend to
        # be wide; pure reference/lookup sheets are often narrow but very long,
        # or vice versa — so we just show row/col counts and let the user pick)
        info = [f"{sn} ({sheets_dict[sn].shape[0]} rows x {sheets_dict[sn].shape[1]} cols)" for sn in sheet_names]
        default_idx = max(range(len(sheet_names)), key=lambda i: sheets_dict[sheet_names[i]].shape[1])
        display_choice = st.selectbox(
            f"{label} — multiple sheets found, choose the one with SKU data:",
            options=info,
            index=default_idx,
            key=f"{key_prefix}_sheet",
        )
        chosen = sheet_names[info.index(display_choice)]
    return chosen, sheets_dict[chosen]


col1, col2 = st.columns(2)
with col1:
    input_file = st.file_uploader("1. Seller / Vendor Input file", type=["xlsx", "xls", "csv"], key="input_upl")
with col2:
    output_file = st.file_uploader("2. AI Enrichment Output file", type=["xlsx", "xls", "csv"], key="output_upl")

if input_file and output_file:
    input_sheets = load_sheets(input_file)
    output_sheets = load_sheets(output_file)

    st.markdown("### Select data sheets")
    scol1, scol2 = st.columns(2)
    with scol1:
        in_sheet_name, input_df = pick_sheet("Input file", input_sheets, "in")
    with scol2:
        out_sheet_name, output_df = pick_sheet("Output file", output_sheets, "out")

    input_df.columns = [str(c).strip() for c in input_df.columns]
    output_df.columns = [str(c).strip() for c in output_df.columns]

    default_key = detect_key_column(list(input_df.columns), list(output_df.columns))
    common_cols = [c for c in input_df.columns if c in list(output_df.columns)]

    if not common_cols:
        st.error("No matching column headers found between the two files. Please check the sheets selected above.")
        st.stop()

    st.markdown("### Matching key")
    key_col = st.selectbox(
        "Column used to match SKUs between the two files:",
        options=common_cols,
        index=common_cols.index(default_key) if default_key in common_cols else 0,
    )

    auto_excluded = detect_image_url_columns([c for c in common_cols if c != key_col])
    with st.expander(f"Columns auto-excluded from comparison ({len(auto_excluded)}) — image/URL fields"):
        st.write(", ".join(auto_excluded) if auto_excluded else "None detected")

    attribute_cols = common_attribute_columns(
        list(input_df.columns), list(output_df.columns), key_col, auto_excluded
    )

    ignore_case_whitespace = st.checkbox(
        "Ignore case & extra whitespace differences (recommended)", value=True
    )

    st.caption(
        f"Comparing **{len(attribute_cols)}** shared attributes across matched SKUs. "
        f"Key: **{key_col}**. Input sheet: *{in_sheet_name}* ({input_df.shape[0]} rows). "
        f"Output sheet: *{out_sheet_name}* ({output_df.shape[0]} rows)."
    )

    if st.button("Run Comparison", type="primary"):
        with st.spinner("Comparing SKUs and building report..."):
            result = run_comparison(
                input_df, output_df, key_col, attribute_cols, auto_excluded, ignore_case_whitespace
            )
            report_bytes = build_report_workbook(result)

        st.success("Comparison complete.")

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("SKUs Compared", result.total_skus)
        m2.metric("Attributes Compared", result.total_attributes)
        m3.metric("SKUs with ≥1 Deviation", result.skus_with_deviation)
        m4.metric("SKUs with ≥1 Backfilled Field", result.skus_with_backfill)
        attrs_with_dev = int((result.attribute_summary["Total Deviating"] > 0).sum())
        m5.metric("Attributes with ≥1 Deviation", attrs_with_dev)

        if result.input_only_keys or result.output_only_keys:
            st.warning(
                f"{len(result.input_only_keys)} SKU(s) found only in the Input file, "
                f"{len(result.output_only_keys)} SKU(s) found only in the Output file. "
                "These were excluded from the comparison."
            )

        st.markdown("#### Attribute-level deviation breakdown")
        display_df = result.attribute_summary.copy()
        display_df["% Deviating"] = display_df["% Deviating"].round(1)
        st.dataframe(display_df, use_container_width=True, hide_index=True)

        deviating_only = display_df[display_df["Total Deviating"] > 0]
        if not deviating_only.empty:
            st.bar_chart(
                deviating_only.set_index("Attribute")[["Value Mismatch", "Missing in AI Output"]]
            )

        st.download_button(
            label="Download Highlighted Comparison Report (.xlsx)",
            data=report_bytes,
            file_name="QC_Comparison_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Upload both files to begin.")
