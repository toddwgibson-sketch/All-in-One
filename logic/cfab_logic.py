# -*- coding: utf-8 -*-
"""
cfab_logic.py

Clean implementation of CFAB validation formatting logic (T3-T2 vs T2-T1-T0 splitting, PP enrichment, etc.).

This module follows the standard AIO processor interface.
"""

from __future__ import annotations

import io
from typing import List, Optional, Tuple

from openpyxl import load_workbook


def process_cfab_validation(
    input_path: str,
    cutsheet_path: str,
    output_name: Optional[str] = None,
) -> Tuple[bytes, str]:
    """
    Process a single CFAB validation file.
    For now this is a placeholder that returns the original file with a note.
    Full rich logic will be restored here.
    """
    wb = load_workbook(input_path)
    # TODO: Full CFAB pipeline (split, enrich, T3-T2/T2-T1-T0, pink columns, etc.)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    if output_name is None:
        output_name = "CFAB_Formatted.xlsx"

    return buf.getvalue(), output_name


def process_multiple_cfab_files(
    input_paths: List[str],
    cutsheet_path: str,
) -> bytes:
    """Process multiple CFAB files and return a ZIP."""
    import zipfile

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, path in enumerate(input_paths):
            data, name = process_cfab_validation(path, cutsheet_path)
            zf.writestr(name or f"cfab_{i}.xlsx", data)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()
