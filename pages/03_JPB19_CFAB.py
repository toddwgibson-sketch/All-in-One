#!/usr/bin/env python3
"""
JPB19 CFAB Formatter - Clean Streamlit Page

Full rich port of the original CFAB validation formatter.

Core logic lives in ../jpb19_cfab_logic.py
"""

import streamlit as st
import sys
from pathlib import Path
import tempfile
import os

sys.path.append(str(Path(__file__).parent.parent))

try:
    from logic import cfab_logic as logic
except ImportError as e:
    st.error(f"Failed to import logic.cfab_logic: {e}")
    st.stop()

st.set_page_config(page_title="JPB19 CFAB Formatter", page_icon="🔌", layout="wide")
st.title("CFAB Formatter (Specialized)")
st.caption("Direct access • T3-T2 / T2-T1-T0 logic • Use AIO for most cases")

with st.sidebar:
    st.header("How to use")
    st.markdown("""
    1. Upload the **CFAB cutsheet** (the one that defines long vs short cable paths).
    2. Upload one or more **CFAB validation result files**.
    3. Click **Process**.
    4. Download the ZIP.

    Full transformations applied (matching original):
    - Split LLDP → Downlink / Mismatch
    - Cutsheet enrichment (Hop/PP columns + peer info)
    - Automatic T3-T2 (long) vs T2-T1-T0 (short) split on all tabs
    - Pink "Current ..." B-side on Mismatch tabs
    - Grey-out matched Optic rows
    - Reciprocal pair highlighting
    - Summary with live formulas
    """)
    st.info("Complete high-fidelity port of CFAB_CODE19NEW.py")

with st.form("jpb19_cfab_form", clear_on_submit=False):
    st.subheader("1. CFAB Cutsheet")
    cutsheet_file = st.file_uploader(
        "CFAB cutsheet (long/short path definitions) (.xlsx)",
        type=["xlsx"],
        key="jpb19_cfab_cutsheet_v1"
    )

    st.subheader("2. CFAB Validation Files")
    input_files = st.file_uploader(
        "One or more CFAB validation exports (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="jpb19_cfab_inputs_v1"
    )

    submitted = st.form_submit_button(
        "🚀 Process CFAB Files",
        type="primary",
        disabled=not (input_files and cutsheet_file)
    )

if submitted:
    with st.spinner("Processing CFAB files..."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                cutsheet_path = os.path.join(tmpdir, "cutsheet.xlsx")
                with open(cutsheet_path, "wb") as out:
                    out.write(cutsheet_file.getbuffer())

                input_paths = []
                for i, f in enumerate(input_files):
                    p = os.path.join(tmpdir, f"input_{i}.xlsx")
                    with open(p, "wb") as out:
                        out.write(f.getbuffer())
                    input_paths.append(p)

                progress = st.progress(0, text="Starting...")
                results = []

                for idx, path in enumerate(input_paths, 1):
                    fname = os.path.basename(path)
                    progress.progress((idx-1)/len(input_paths), text=f"Processing {idx}/{len(input_paths)}: {fname}")
                    xlsx_bytes, out_name = logic.process_cfab_validation(path, cutsheet_path)
                    results.append((out_name, xlsx_bytes))

                progress.progress(1.0, text="Building ZIP...")

                import zipfile as zf
                zip_buffer = os.path.join(tmpdir, "cfab_out.zip")
                with zf.ZipFile(zip_buffer, "w", zf.ZIP_DEFLATED) as zipf:
                    for out_name, data in results:
                        zipf.writestr(out_name, data)

                with open(zip_buffer, "rb") as f:
                    zip_bytes = f.read()

                st.success(f"Processed {len(input_paths)} CFAB file(s) successfully!")

                with st.expander("Output file names (rack-based)", expanded=True):
                    for out_name, _ in results:
                        st.write(f"• {out_name}")

                st.download_button(
                    label="📥 Download All Formatted CFAB Reports (ZIP)",
                    data=zip_bytes,
                    file_name="JPB19_CFAB_Formatted.zip",
                    mime="application/zip",
                    use_container_width=True
                )
        except Exception as e:
            st.error("CFAB processing failed")
            st.exception(e)

st.caption("Core logic: jpb19_cfab_logic.py • One page in the JPB19 Reports suite")
