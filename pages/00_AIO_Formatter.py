# -*- coding: utf-8 -*-
"""
AIO Validation Formatter (All-In-One)

The single tool that can handle ANY hall, ANY fabric type, ANY source.

It reads the cutsheet first (the cutsheet knows what this is),
then automatically chooses and runs the correct rich formatting logic.
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
st.caption("Any Hall • Any Fabric • Any Source • One Tool")

st.markdown("""
**This is the main tool you should use.**

Upload your raw validation exports + the relevant **cutsheet**.  
The AIO reads the cutsheet to understand what kind of fabric/pathway this is,  
then automatically dispatches to the correct rich processing logic.

No more guessing which script to run.
""")

# =============================================================================
# Session State Initialization
# =============================================================================
if "aio_stage" not in st.session_state:
    st.session_state.aio_stage = "upload"
if "aio_detection" not in st.session_state:
    st.session_state.aio_detection = None
if "aio_input_bytes" not in st.session_state:
    st.session_state.aio_input_bytes = []
if "aio_cutsheet_bytes" not in st.session_state:
    st.session_state.aio_cutsheet_bytes = None
if "aio_override" not in st.session_state:
    st.session_state.aio_override = None
if "aio_result" not in st.session_state:
    st.session_state.aio_result = None

# =============================================================================
# STAGE 1: Upload
# =============================================================================
if st.session_state.aio_stage == "upload":
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1. Raw Validation Files")
        input_files = st.file_uploader(
            "Upload one or more validation export files (.xlsx)",
            type=["xlsx"],
            accept_multiple_files=True,
            key="aio_inputs",
            help="These are the messy files coming out of hall testing / Slack / LVV Portal"
        )

    with col2:
        st.subheader("2. Cutsheet (Required)")
        cutsheet_file = st.file_uploader(
            "The cutsheet that describes this fabric (this file tells the AIO what to do)",
            type=["xlsx"],
            key="aio_cutsheet",
            help="The cutsheet contains the intelligence about hall, fabric type, and source system"
        )

    if not (input_files and cutsheet_file):
        st.info("Please upload at least one validation file and a cutsheet.")

    if st.button("🚀 Detect & Process (AIO)", type="primary", disabled=not (input_files and cutsheet_file)):
        st.session_state.aio_input_bytes = [(f.name, f.getbuffer().tobytes()) for f in input_files]
        st.session_state.aio_cutsheet_bytes = (cutsheet_file.name, cutsheet_file.getbuffer().tobytes())

        with st.spinner("Inspecting cutsheet..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                cutsheet_path = os.path.join(tmpdir, "cutsheet.xlsx")
                with open(cutsheet_path, "wb") as out:
                    out.write(st.session_state.aio_cutsheet_bytes[1])

                detection = aio.inspect_cutsheet(cutsheet_path)
                st.session_state.aio_detection = detection
                st.session_state.aio_stage = "detected"
                st.rerun()

# =============================================================================
# STAGE 2: Detection + Override
# =============================================================================
elif st.session_state.aio_stage == "detected":
    detection = st.session_state.aio_detection

    st.subheader("🔍 Cutsheet Detection Results")
    st.info(detection.summary())

    if detection.reasons:
        with st.expander("Why the AIO thinks this:", expanded=False):
            for r in detection.reasons:
                st.write(f"• {r}")

    st.divider()
    st.subheader("Optional Override")

    processor_options = ["auto (use detection)", "gfab", "hops", "cfab", "ipr", "lv_portal"]
    chosen = st.selectbox(
        "Force a specific processor",
        processor_options,
        index=0,
        key="aio_override_select"
    )

    if st.button("✅ Confirm and Generate Report", type="primary"):
        override = None if chosen == "auto (use detection)" else chosen
        st.session_state.aio_override = override

        processor_to_use = override or detection.recommended_processor

        supported = ["gfab", "hops", "cfab", "ipr", "lv_portal"]
        if processor_to_use not in supported:
            st.error(f"Processor '{processor_to_use}' is not yet fully wired in the AIO.")
            st.stop()

        with st.spinner(f"Running {processor_to_use} logic..."):
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    input_paths = []
                    for i, (name, data) in enumerate(st.session_state.aio_input_bytes):
                        p = os.path.join(tmpdir, f"input_{i}.xlsx")
                        with open(p, "wb") as out:
                            out.write(data)
                        input_paths.append(p)

                    cutsheet_path = os.path.join(tmpdir, "cutsheet.xlsx")
                    with open(cutsheet_path, "wb") as out:
                        out.write(st.session_state.aio_cutsheet_bytes[1])

                    if len(input_paths) == 1:
                        result = aio.process_aio_validation(
                            input_paths[0], cutsheet_path, override_processor=override
                        )
                        if isinstance(result, tuple) and len(result) == 2:
                            xlsx_bytes, fname = result
                            info = {}
                        else:
                            xlsx_bytes, fname, info = result
                        st.session_state.aio_result = (xlsx_bytes, fname, info, False)
                    else:
                        result = aio.process_multiple_aio_validation(
                            input_paths, cutsheet_path, override_processor=override
                        )
                        if isinstance(result, (bytes, bytearray)):
                            zip_bytes = result
                            info = {}
                        else:
                            zip_bytes, info = result
                        st.session_state.aio_result = (zip_bytes, "AIO_Formatted_Reports.zip", info, True)

                    st.session_state.aio_stage = "processed"
                    st.rerun()

            except Exception as proc_err:
                st.error("Processing failed for this file type.")
                st.exception(proc_err)

    if st.button("← Start Over"):
        for key in list(st.session_state.keys()):
            if key.startswith("aio_"):
                del st.session_state[key]
        st.rerun()

# =============================================================================
# STAGE 3: Results
# =============================================================================
elif st.session_state.aio_stage == "processed":
    result = st.session_state.aio_result

    if isinstance(result, (list, tuple)) and len(result) >= 1:
        data = result[0]
        fname = result[1] if len(result) >= 2 else "Report.xlsx"
        info = result[2] if len(result) >= 3 else {}
        is_multi = len(result) >= 4 and bool(result[3])

        if is_multi:
            st.success(f"Generated ZIP with multiple reports")
            st.download_button(
                "📥 Download All Formatted Reports (ZIP)",
                data,
                fname,
                mime="application/zip",
                use_container_width=True
            )
        else:
            xlsx_bytes = data
            st.success(f"✅ Report generated: **{fname}**")

            try:
                from io import BytesIO
                from openpyxl import load_workbook

                wb_preview = load_workbook(BytesIO(xlsx_bytes), data_only=False)

                st.subheader("📊 Quick Analytics — See the damage before you download")

                total_sheets = len(wb_preview.sheetnames)
                data_rows = 0
                tab_counts = {}

                for sheet_name in wb_preview.sheetnames:
                    ws = wb_preview[sheet_name]
                    row_count = max(0, ws.max_row - 1)
                    tab_counts[sheet_name] = row_count
                    if sheet_name.lower() != "summary":
                        data_rows += row_count

                col1, col2, col3 = st.columns(3)
                col1.metric("Total Data Rows", f"{data_rows:,}")
                col2.metric("Tabs in Report", total_sheets)
                col3.metric("Biggest Tab", max(tab_counts, key=tab_counts.get) if tab_counts else "N/A")

                st.markdown("**Issues by Category**")
                display_counts = {k: v for k, v in tab_counts.items() if v > 0 and k.lower() != "summary"}

                if display_counts:
                    import pandas as pd
                    df_counts = pd.DataFrame(list(display_counts.items()), columns=["Tab", "Rows"]).sort_values("Rows", ascending=False)
                    st.dataframe(df_counts, use_container_width=True, hide_index=True)
                    st.bar_chart(df_counts.set_index("Tab")["Rows"])
                else:
                    st.info("No significant data tabs found.")

                if "Summary" in wb_preview.sheetnames:
                    ws_sum = wb_preview["Summary"]
                    st.markdown("**Summary Tab Preview**")
                    summary_data = []
                    for row in ws_sum.iter_rows(min_row=1, max_row=min(12, ws_sum.max_row), values_only=True):
                        summary_data.append([str(c) if c is not None else "" for c in row])
                    if summary_data:
                        st.table(summary_data)

            except Exception as preview_err:
                st.warning(f"Could not generate full preview analytics: {preview_err}")

            st.download_button(
                "📥 Download Formatted Report",
                xlsx_bytes,
                fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.error("Unexpected result format.")
        if st.button("← Back to Upload"):
            for key in list(st.session_state.keys()):
                if key.startswith("aio_"):
                    del st.session_state[key]
            st.rerun()

    if st.button("← Start New Analysis"):
        for key in list(st.session_state.keys()):
            if key.startswith("aio_"):
                del st.session_state[key]
        st.rerun()
