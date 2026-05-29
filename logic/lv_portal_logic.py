# -*- coding: utf-8 -*-
"""
lv_portal_logic.py

Clean implementation for JBP15 (and similar) T0-to-Host / LV Portal reports.
"""

from __future__ import annotations

import io
from typing import List, Optional, Tuple

from openpyxl import load_workbook


def process_lv_portal_validation(
    input_path: str,
    cutsheet_path: str,
    output_name: Optional[str] = None,
) -> Tuple[bytes, str]:
    """
    Process one LV Portal / T0-to-Host validation export.
    Full rich logic is being integrated.
    """
    wb = load_workbook(input_path)
    # TODO: Full T0-to-Host LV Portal pipeline

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    if output_name is None:
        output_name = "LV_Portal_T0_Formatted.xlsx"

    return buf.getvalue(), output_name


def process_multiple_lv_portal_files(
    input_paths: List[str],
    cutsheet_path: str,
) -> bytes:
    """Process multiple LV Portal files and return a ZIP."""
    import zipfile

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, path in enumerate(input_paths):
            data, name = process_lv_portal_validation(path, cutsheet_path)
            zf.writestr(name or f"lv_portal_{i}.xlsx", data)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()
