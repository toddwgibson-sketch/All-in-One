#!/usr/bin/env python3
"""
JPB19 IPR Formatter - Clean Streamlit Page

Combines per-building audit files + enriches with cutsheet (PP / Other End).

Core logic lives in ../jpb19_ipr_logic.py
"""

import streamlit as st
import sys
from pathlib import Path
import tempfile
import os

sys.path.append(str(Path(__file__).parent.parent))

try:
    from logic import ipr_logic as logic
except ImportError as e:
    st.error(f"Failed to import logic.ipr_logic: {e}")
    st.stop()

st.set_page_config(page_title="JPB19 IPR Formatter", page_icon="📋", layout="wide")
st.title("IPR Combiner (Specialized)")
st.caption("Direct access • Multi-file + enrichment • Use AIO for most cases")

with st.sidebar:
    st.markdown("""
    Upload one or more per-building audit exports + the IPR cutsheet.
    The tool combines them, enriches with PP / Other End data from the cutsheet,
    applies consistent styling, and produces a single downloadable workbook.
    """)
    st.info("Clean port following the same architecture as GFAB / HOPS / CFAB")

with st.form("jpb19_ipr_form", clear_on_submit=False):
    st.subheader("1. IPR Cutsheet")
    cutsheet_file = st.file_uploader(
        "Cutsheet (5-col: Device A, Rack A, PP, Device B, Rack B) (.xlsx)",
        type=["xlsx"],
        key="jpb19_ipr_cutsheet"
    )

    st.subheader("2. Per-Building Audit Files")
    input_files = st.file_uploader(
        "One or more audit exports to combine (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="jpb19_ipr_inputs"
    )

    submitted = st.form_submit_button(
        "🚀 Process & Combine",
        type="primary",
        disabled=not (input_files and cutsheet_file)
    )

if submitted:
    with st.spinner("Combining and formatting IPR files..."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                cutsheet_path = os.path.join(tmpdir, "cutsheet.xlsx")
                with open(cutsheet_path, "wb") as out:
                    out.write(cutsheet_file.getbuffer())

                input_paths = []
                for i, f in enumerate(input_files):
                    p = os.path.join(tmpdir, f"in_{i}.xlsx")
                    with open(p, "wb") as out:
                        out.write(f.getbuffer())
                    input_paths.append(p)

                if len(input_paths) == 1:
                    xlsx_bytes, out_name = logic.process_ipr_validation(input_paths[0], cutsheet_path)
                    st.success("Processed successfully!")
                    st.download_button("📥 Download Formatted IPR Report",
                                       xlsx_bytes, out_name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    zip_bytes = logic.process_multiple_ipr_files(input_paths, cutsheet_path)
                    st.success(f"Combined and processed {len(input_paths)} files!")
                    st.download_button("📥 Download Combined IPR Reports (ZIP)",
                                       zip_bytes, "JPB19_IPR_Combined.zip", "application/zip")

        except Exception as e:
            st.error("IPR processing failed")
            st.exception(e)

st.caption("Core logic: jpb19_ipr_logic.py • One page in the JPB19 Reports suite")
