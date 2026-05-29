# -*- coding: utf-8 -*-
"""
gfab_logic.py

Clean, importable core logic for GFAB Validation Formatter.

This module follows the standard AIO processor interface.
All original GFAB formatting behaviour will be restored here.
"""

from __future__ import annotations

import io
from typing import List, Optional, Tuple

from openpyxl import load_workbook


def process_gfab_validation(
    input_path: str,
    cutsheet_path: str,
    output_name: Optional[str] = None,
) -> Tuple[bytes, str]:
    """
    Process a single GFAB validation file.
    Currently returns the input file as-is (placeholder).
    Full rich logic will be restored.
    """
    wb = load_workbook(input_path)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    if output_name is None:
        output_name = "GFAB_Formatted.xlsx"

    return buf.getvalue(), output_name


def process_multiple_files(
    input_paths: List[str],
    cutsheet_path: str,
) -> bytes:
    """
    Process multiple GFAB files and return a ZIP.
    """
    import zipfile

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, path in enumerate(input_paths):
            data, name = process_gfab_validation(path, cutsheet_path)
            zf.writestr(name or f"gfab_{i}.xlsx", data)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()
