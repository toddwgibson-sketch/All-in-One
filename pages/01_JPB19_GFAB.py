#!/usr/bin/env python3
"""
JPB19 GFAB Formatter - Clean Streamlit Page

UI layer only. Lives in pages/01_JPB19_GFAB.py
Core logic lives in ../jpb19_gfab_logic.py (project root).
"""

import streamlit as st
import sys
from pathlib import Path
import tempfile
import os

sys.path.append(str(Path(__file__).parent.parent))

try:
    from logic import gfab_logic as logic
except ImportError as e:
    st.error(f"Failed to import logic.gfab_logic: {e}")
    st.stop()

st.set_page_config(page_title="JPB19 GFAB Formatter", page_icon="🗂️", layout="wide")
st.title("GFAB Formatter (Specialized)")
st.caption("Direct access to full GFAB logic • Use AIO for most cases")

with st.sidebar:
    st.header("How to use")
    st.markdown("""
    1. Upload one or more **GFAB validation exports** from the hall.
    2. Upload the **JPB19 combined cutsheet**.
    3. Click **Process**.
    4. Download the formatted output(s).
    """)
    st.info("Part of the unified JPB19 Reports app. Same clean architecture as the other tools.")

# =============================================================================
# Form (prevents widget-rerun loops)
# =============================================================================
with st.form("jpb19_gfab_form", clear_on_submit=False):
    st.subheader("1. GFAB Validation Exports")
    input_files = st.file_uploader(
        "Upload one or more GFAB validation result files (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="jpb19_gfab_inputs_v3"
    )

    st.subheader("2. JPB19 Cutsheet")
    cutsheet_file = st.file_uploader(
        "JPB19 Combined Cutsheet (.xlsx)",
        type=["xlsx"],
        key="jpb19_gfab_cutsheet_v3"
    )

    submitted = st.form_submit_button(
        "🚀 Process Files",
        type="primary",
        disabled=not (input_files and cutsheet_file)
    )

# =============================================================================
# Processing (only runs on fresh form submit)
# =============================================================================
if submitted:
    with st.spinner("Processing..."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                input_paths = []
                for i, f in enumerate(input_files):
                    p = os.path.join(tmpdir, f"input_{i}.xlsx")
                    with open(p, "wb") as out:
                        out.write(f.getbuffer())
                    input_paths.append(p)

                cutsheet_path = os.path.join(tmpdir, "cutsheet.xlsx")
                with open(cutsheet_path, "wb") as out:
                    out.write(cutsheet_file.getbuffer())

                progress = st.progress(0, text="Starting...")
                results = []

                # Process one-by-one so we can show per-file status + final names
                for idx, path in enumerate(input_paths, 1):
                    fname = os.path.basename(path)
                    progress.progress(
                        (idx - 1) / len(input_paths),
                        text=f"Processing {idx}/{len(input_paths)}: {fname}"
                    )
                    xlsx_bytes, out_name = logic.process_gfab_validation(path, cutsheet_path)
                    results.append((out_name, xlsx_bytes))

                progress.progress(1.0, text="Building ZIP...")

                # Build zip using the nice per-file names the logic already chose
                import zipfile as zf
                zip_buffer = os.path.join(tmpdir, "out.zip")
                with zf.ZipFile(zip_buffer, "w", zf.ZIP_DEFLATED) as zipf:
                    for out_name, data in results:
                        zipf.writestr(out_name, data)

                with open(zip_buffer, "rb") as f:
                    zip_bytes = f.read()

                st.success(f"Processed {len(input_paths)} file(s) successfully!")

                # Show what we actually produced (very useful for validators)
                with st.expander("Output file names (rack-based where possible)", expanded=True):
                    for out_name, _ in results:
                        st.write(f"• {out_name}")

                st.download_button(
                    label="📥 Download All Formatted Reports (ZIP)",
                    data=zip_bytes,
                    file_name="JPB19_GFAB_Formatted.zip",
                    mime="application/zip",
                    use_container_width=True
                )

        except Exception as e:
            st.error("Processing failed")
            st.exception(e)

st.caption("Core logic: jpb19_gfab_logic.py • This is one page in the JPB19 Reports suite")