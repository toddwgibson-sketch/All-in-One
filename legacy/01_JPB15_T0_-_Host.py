#!/usr/bin/env python3
"""
JBP15 LV Portal Formatter - Streamlit Page

This file goes in the pages/ folder.

The core logic lives in lv_portal_logic.py (in the project root).
"""

import streamlit as st
import sys
from pathlib import Path
import tempfile
import os

# Add parent directory so we can import the logic
sys.path.append(str(Path(__file__).parent.parent))

try:
    import lv_portal_logic as logic
except ImportError as e:
    st.error(
        "Failed to import lv_portal_logic.py\n\n"
        "Make sure the file 'lv_portal_logic.py' exists in the main project folder "
        "(one level above the pages/ folder).\n\n"
        f"Error: {e}"
    )
    st.stop()

st.set_page_config(page_title="LV Portal Formatter - JBP15", page_icon="🔌", layout="wide")
st.title("LV Portal Validation Formatter — JBP15 T0→Host")
st.caption("Streamlit version • Clean two-file setup")

with st.sidebar:
    st.header("Instructions")
    st.markdown("""
    1. Upload one or more **GPU/Compute Cutsheets**.
    2. Upload the main **LV Portal Validation Export**.
    3. (Optional) Upload a previous formatted report for History.
    4. Click **Process**.
    """)

# Use a form so all options are collected before processing
with st.form("lv_portal_form"):
    st.subheader("1. GPU / Compute Cutsheets")
    gpu_files = st.file_uploader(
        "Upload one or more GPU/Compute cutsheets",
        type=["xlsx"],
        accept_multiple_files=True,
        key="gpu_cutsheets_clean"
    )

    st.subheader("2. LV Portal Validation Export")
    portal_file = st.file_uploader(
        "LV Portal Validation Export (.xlsx)",
        type=["xlsx"],
        key="portal_report_clean"
    )

    st.subheader("3. Previous Report (Optional)")
    prev_file = st.file_uploader(
        "Previous formatted report (for History)",
        type=["xlsx"],
        key="prev_report_clean"
    )

    submitted = st.form_submit_button("🚀 Process", type="primary")

if submitted:
    if not gpu_files or not portal_file:
        st.error("Please upload at least one GPU cutsheet and the main LV Portal report.")
        st.stop()

    with st.spinner("Processing... This can take 1-3 minutes on large reports."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                # Save GPU cutsheets
                gpu_paths = []
                for i, f in enumerate(gpu_files):
                    p = os.path.join(tmpdir, f"gpu_{i}.xlsx")
                    with open(p, "wb") as out:
                        out.write(f.getbuffer())
                    gpu_paths.append(p)

                # Save main report
                report_path = os.path.join(tmpdir, "report.xlsx")
                with open(report_path, "wb") as out:
                    out.write(portal_file.getbuffer())

                # Save previous report if provided
                prev_path = None
                if prev_file:
                    prev_path = os.path.join(tmpdir, "prev.xlsx")
                    with open(prev_path, "wb") as out:
                        out.write(prev_file.getbuffer())

                # --- Safe monkey-patching of the original pickers ---
                original_pick = getattr(logic, 'pick_file', None)
                original_pick_multi = getattr(logic, 'pick_multiple_files', None)

                if original_pick is None or original_pick_multi is None:
                    st.error(
                        "lv_portal_logic.py does not expose `pick_file` or `pick_multiple_files`.\n\n"
                        "Please make sure you renamed the *original* full logic script to `lv_portal_logic.py`."
                    )
                    st.stop()

                def fake_pick_multi(title, *a, **k):
                    if "gpu" in title.lower() or "compute" in title.lower():
                        return gpu_paths
                    return []

                def fake_pick(title, *a, **k):
                    title_lower = title.lower()
                    if "lv portal" in title_lower or "validation" in title_lower:
                        return report_path
                    if "previous" in title_lower or "history" in title_lower:
                        return prev_path
                    return None

                logic.pick_multiple_files = fake_pick_multi
                logic.pick_file = fake_pick

                # Prevent the original main() from killing the Streamlit session
                original_exit = sys.exit
                sys.exit = lambda *a, **k: None

                try:
                    logic.main()
                finally:
                    sys.exit = original_exit
                    logic.pick_file = original_pick
                    logic.pick_multiple_files = original_pick_multi

                # Try to find the output file created by the original logic
                base = os.path.splitext(report_path)[0]
                out_path = base + "_formatted.xlsx"

                if os.path.exists(out_path):
                    with open(out_path, "rb") as f:
                        result_bytes = f.read()

                    st.success("Processing complete!")
                    st.download_button(
                        label="📥 Download Formatted Report",
                        data=result_bytes,
                        file_name=os.path.basename(out_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.error("The logic ran but no output file was created. Check the logs above.")

        except Exception as e:
            st.error("Processing failed")
            st.exception(e)

st.caption("Core logic lives in lv_portal_logic.py (project root) • UI lives in pages/")