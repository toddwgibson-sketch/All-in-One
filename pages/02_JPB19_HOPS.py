#!/usr/bin/env python3
"""
JPB19 HOPS Formatter - Clean Streamlit Page

GPU / HOPS validation reports (LLDP Mismatch (GPU), Optic Errors, FEC_BER, Interface Down).

Core logic lives in ../jpb19_hops_logic.py
"""

import streamlit as st
import sys
from pathlib import Path
import tempfile
import os

sys.path.append(str(Path(__file__).parent.parent))

try:
    from logic import hops_logic as logic
except ImportError as e:
    st.error(f"Failed to import logic.hops_logic: {e}")
    st.stop()

st.set_page_config(page_title="JPB19 HOPS Formatter", page_icon="🖥️", layout="wide")
st.title("HOPS / GPU Formatter (Specialized)")
st.caption("Direct access to full HOPS logic • Use AIO for most cases")

with st.sidebar:
    st.header("How to use")
    st.markdown("""
    1. Upload the **JPB19 Combined Cutsheet** (the special one with DeviceA/PortA + DeviceB/PortB).
    2. Upload one or more **HOPS validation result files** (the ones containing LLDP Mismatch (GPU), Optic Errors (GPU), etc.).
    3. Click **Process**.
    4. Download the ZIP of formatted reports.

    The logic applies the full original transformations:
    - Row filtering (junk + non-compute)
    - A/B side enrichment via canon_slot / canon_swp
    - Heavy column reordering + pink "Current" blocks
    - "maybe CT off" grey marking
    - Summary refresh + full workbook formatting
    - Reciprocal mismatch pair highlighting
    """)
    st.info("Uses the exact same business logic as the original HOPS CODE19.py")

# =============================================================================
# Form
# =============================================================================
with st.form("jpb19_hops_form", clear_on_submit=False):
    st.subheader("1. JPB19 Combined Cutsheet (HOPS)")
    cutsheet_file = st.file_uploader(
        "Combined cutsheet with DeviceA/PortA + DeviceB/PortB columns (.xlsx)",
        type=["xlsx"],
        key="jpb19_hops_cutsheet_v1"
    )

    st.subheader("2. HOPS Validation Result Files")
    input_files = st.file_uploader(
        "One or more HOPS validation exports (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="jpb19_hops_inputs_v1"
    )

    submitted = st.form_submit_button(
        "🚀 Process HOPS Files",
        type="primary",
        disabled=not (input_files and cutsheet_file)
    )

# =============================================================================
# Processing
# =============================================================================
if submitted:
    with st.spinner("Processing HOPS files..."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                # Save cutsheet
                cutsheet_path = os.path.join(tmpdir, "combined_cutsheet.xlsx")
                with open(cutsheet_path, "wb") as out:
                    out.write(cutsheet_file.getbuffer())

                # Save inputs
                input_paths = []
                for i, f in enumerate(input_files):
                    p = os.path.join(tmpdir, f"input_{i}.xlsx")
                    with open(p, "wb") as out:
                        out.write(f.getbuffer())
                    input_paths.append(p)

                progress = st.progress(0, text="Starting HOPS processing...")
                results = []

                for idx, path in enumerate(input_paths, 1):
                    fname = os.path.basename(path)
                    progress.progress(
                        (idx - 1) / len(input_paths),
                        text=f"Processing {idx}/{len(input_paths)}: {fname}"
                    )
                    xlsx_bytes, out_name = logic.process_hops_validation(path, cutsheet_path)
                    results.append((out_name, xlsx_bytes))

                progress.progress(1.0, text="Building ZIP...")

                # Build zip
                import zipfile as zf
                zip_buffer = os.path.join(tmpdir, "hops_out.zip")
                with zf.ZipFile(zip_buffer, "w", zf.ZIP_DEFLATED) as zipf:
                    for out_name, data in results:
                        zipf.writestr(out_name, data)

                with open(zip_buffer, "rb") as f:
                    zip_bytes = f.read()

                st.success(f"Processed {len(input_paths)} HOPS file(s) successfully!")

                with st.expander("Output file names (rack-based where possible)", expanded=True):
                    for out_name, _ in results:
                        st.write(f"• {out_name}")

                st.download_button(
                    label="📥 Download All Formatted HOPS Reports (ZIP)",
                    data=zip_bytes,
                    file_name="JPB19_HOPS_Formatted.zip",
                    mime="application/zip",
                    use_container_width=True
                )

        except Exception as e:
            st.error("HOPS processing failed")
            st.exception(e)

st.caption("Core logic: jpb19_hops_logic.py • One page in the JPB19 Reports suite")
