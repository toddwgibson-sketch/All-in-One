# -*- coding: utf-8 -*-
"""
hops_logic.py

Clean implementation of HOPS/GPU validation formatting logic.
"""

from __future__ import annotations

import io
from typing import List, Optional, Tuple

from openpyxl import load_workbook


def process_hops_validation(
    input_path: str,
    cutsheet_path: str,
    output_name: Optional[str] = None,
) -> Tuple[bytes, str]:
    """
    Process a single HOPS validation file.
    Placeholder until full rich logic is restored.
    """
    wb = load_workbook(input_path)
    # TODO: Full HOPS pipeline (combined cutsheet lookups, GPU sheets, etc.)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    if output_name is None:
        output_name = "HOPS_Formatted.xlsx"

    return buf.getvalue(), output_name


def process_multiple_hops_files(
    input_paths: List[str],
    cutsheet_path: str,
) -> bytes:
    """Process multiple HOPS files and return a ZIP."""
    import zipfile

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, path in enumerate(input_paths):
            data, name = process_hops_validation(path, cutsheet_path)
            zf.writestr(name or f"hops_{i}.xlsx", data)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()
