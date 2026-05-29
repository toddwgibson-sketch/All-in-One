#!/usr/bin/env python3
"""
AIO Validation Formatter (All-In-One)

The single tool that can handle ANY hall, ANY fabric type, ANY source.

It reads the cutsheet first (the cutsheet knows what this is),
then automatically chooses and runs the correct rich formatting logic
(GFAB, HOPS, CFAB, IPR, future LV Portal, etc.).

This is the recommended starting point for most users.
"""

import streamlit as st
import sys
from pathlib import Path
import tempfile
import os

# Robust import for both local run and Streamlit Cloud
root_dir = Path(__file__).parent.parent
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

try:
    from logic import aio_logic as aio
except ImportError as e:
    st.error(f"Failed to import logic.aio_logic: {e}")
    st.stop()

st.set_page_config(
    page_title="AIO Validation Formatter",
    page_icon="🧠",
    layout="wide"
)

st.title("🧠 All In One Validation Formatter")
st.caption("Any Hall • Any Fabric (GFAB/CFAB/HOPS/LV-Portal/etc.) • Slack or LVV Portal • One Button")

st.markdown("""
**This is the main tool you should use.**

Upload your raw validation exports + the relevant **cutsheet**.  
The AIO reads the cutsheet to understand what kind of fabric/pathway this is 
(GFAB, CFAB/HOPS, T3-T2 vs T2-T1-T0, LV Portal style, Slack export, etc.), 
then automatically dispatches to the correct rich processing logic we have built.

No more guessing which script to run.
""")

# =============================================================================
# Uploads (must be outside the form so the button can react to file selection)
# =============================================================================
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Raw Validation Files")
    input_files = st.file_uploader(
        "Upload one or more validation export files (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="aio_inputs",
        help="These are the messy files coming out of the hall testing / Slack / LVV Portal"
    )

with col2:
    st.subheader("2. Cutsheet (Required)")
    cutsheet_file = st.file_uploader(
        "The cutsheet that describes this fabric (this file tells the AIO what to do)",
        type=["xlsx"],
        key="aio_cutsheet",
        help="The cutsheet contains the intelligence about hall, fabric type, and source system"
    )

# =============================================================================
# Submit Button (inside a minimal form)
# =============================================================================
with st.form("aio_form", clear_on_submit=False):
    submitted = st.form_submit_button(
        "🚀 Detect & Process (AIO)",
        type="primary",
        disabled=not (input_files and cutsheet_file)
    )

if not (input_files and cutsheet_file):
    st.info("Please upload at least one validation file and a cutsheet to enable the button above.")

# =============================================================================
# Processing + Detection
# =============================================================================
if submitted:
    with st.spinner("Inspecting cutsheet and detecting configuration..."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                # Save inputs
                input_paths = []
                for i, f in enumerate(input_files):
                    p = os.path.join(tmpdir, f"raw_{i}.xlsx")
                    with open(p, "wb") as out:
                        out.write(f.getbuffer())
                    input_paths.append(p)

                # Save cutsheet
                cutsheet_path = os.path.join(tmpdir, "cutsheet.xlsx")
                with open(cutsheet_path, "wb") as out:
                    out.write(cutsheet_file.getbuffer())

                # === THE KEY STEP ===
                detection = aio.inspect_cutsheet(cutsheet_path)

                st.subheader("🔍 Cutsheet Detection Results")
                st.info(detection.summary())

                if detection.reasons:
                    with st.expander("Why the AIO thinks this:", expanded=False):
                        for r in detection.reasons:
                            st.write(f"• {r}")

                if detection.cutsheet_metadata:
                    with st.expander("Cutsheet metadata (first sheet + sheets)", expanded=False):
                        st.json(detection.cutsheet_metadata)

                # Manual override
                st.divider()
                st.subheader("Optional Override")

                processor_options = ["auto (use detection)", "gfab", "hops", "cfab", "ipr"]
                chosen = st.selectbox(
                    "Force a specific processor (only if detection is wrong)",
                    processor_options,
                    index=0,
                    key="aio_override"
                )
                override = None if chosen == "auto (use detection)" else chosen

                if st.button("✅ Confirm and Generate Report", type="primary"):
                    with st.spinner(f"Running {chosen if override else detection.recommended_processor} logic..."):
                        if len(input_paths) == 1:
                            xlsx_bytes, fname, info = aio.process_aio_validation(
                                input_paths, cutsheet_path, override_processor=override
                            )

                            # === ANALYTICS DASHBOARD (pre-download) ===
                            st.success(f"✅ Report generated: **{fname}**")

                            try:
                                from io import BytesIO
                                from openpyxl import load_workbook

                                wb_preview = load_workbook(BytesIO(xlsx_bytes), data_only=False)

                                st.subheader("📊 Quick Analytics — See the damage before you download")

                                # Basic counts
                                total_sheets = len(wb_preview.sheetnames)
                                data_rows = 0
                                tab_counts = {}

                                for sheet_name in wb_preview.sheetnames:
                                    ws = wb_preview[sheet_name]
                                    row_count = max(0, ws.max_row - 1)  # exclude header
                                    tab_counts[sheet_name] = row_count
                                    if sheet_name.lower() != "summary":
                                        data_rows += row_count

                                col1, col2, col3 = st.columns(3)
                                col1.metric("Total Data Rows", f"{data_rows:,}")
                                col2.metric("Tabs in Report", total_sheets)
                                col3.metric("Biggest Tab", max(tab_counts, key=tab_counts.get) if tab_counts else "N/A")

                                # Tab breakdown table + chart
                                st.markdown("**Issues by Category**")

                                # Filter out Summary and very small tabs for clarity
                                display_counts = {k: v for k, v in tab_counts.items() if v > 0 and k.lower() != "summary"}

                                if display_counts:
                                    import pandas as pd
                                    df_counts = pd.DataFrame(
                                        list(display_counts.items()),
                                        columns=["Tab", "Rows"]
                                    ).sort_values("Rows", ascending=False)

                                    st.dataframe(df_counts, use_container_width=True, hide_index=True)

                                    st.bar_chart(df_counts.set_index("Tab")["Rows"])
                                else:
                                    st.info("No significant data tabs found.")

                                # Show the actual Summary tab content if it exists (very useful)
                                if "Summary" in wb_preview.sheetnames:
                                    ws_sum = wb_preview["Summary"]
                                    st.markdown("**Summary Tab Preview** (what will be in the Excel)")
                                    summary_data = []
                                    for row in ws_sum.iter_rows(min_row=1, max_row=min(12, ws_sum.max_row), values_only=True):
                                        summary_data.append([str(c) if c is not None else "" for c in row])
                                    if summary_data:
                                        st.table(summary_data)

                            except Exception as preview_err:
                                st.warning(f"Could not generate preview analytics: {preview_err}")
                                st.info("The file is still valid — you can download it below.")

                            # Download
                            st.download_button(
                                "📥 Download Formatted Report",
                                xlsx_bytes,
                                fname,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )

                        else:
                            zip_bytes, info = aio.process_multiple_aio_validation(
                                input_paths, cutsheet_path, override_processor=override
                            )
                            st.success(f"Generated ZIP with {len(input_paths)} formatted reports")
                            st.download_button(
                                "📥 Download All Formatted Reports (ZIP)",
                                zip_bytes,
                                "AIO_Formatted_Reports.zip",
                                mime="application/zip",
                                use_container_width=True
                            )

                        with st.expander("Processing details", expanded=False):
                            st.json({k: str(v) for k, v in info.items()})

        except Exception as e:
            st.error("AIO processing failed")
            st.exception(e)

# =============================================================================
# Sidebar help
# =============================================================================
with st.sidebar:
    st.header("How the AIO Works")
    st.markdown("""
    1. You upload raw validation files + the cutsheet.
    2. The AIO **inspects the cutsheet** (not just the raw files).
    3. It extracts:
       - Which hall (JPB19, JBP15, SYD20...)
       - Fabric type (GFAB / CFAB / HOPS / LV Portal...)
       - Source (Slack export vs LVV Portal)
    4. It automatically picks the correct rich formatter we built.
    5. You get the properly formatted, actionable report.

    The cutsheet is the single source of truth.  
    This is why we can support **any room**.
    """)

    st.divider()
    st.caption("Currently supports: GFAB, HOPS, CFAB, IPR (more coming as we add halls)")

st.caption("All In One • Master logic from all cleaned formatters • Old scripts in legacy/")

st.markdown("---")
st.caption("Have an idea to make this even better? Go back to the main **app.py** page and use the **💡 Suggest a Feature** section at the bottom. It creates a ready-to-submit GitHub issue for you.")
