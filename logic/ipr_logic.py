# -*- coding: utf-8 -*-
"""
ipr_logic.py

Clean implementation of IPR audit combiner + cutsheet enrichment.
"""

from __future__ import annotations

import io
from typing import List, Optional, Tuple

from openpyxl import load_workbook


def process_ipr_validation(
    input_path: str,
    cutsheet_path: Optional[str] = None,
    output_name: Optional[str] = None,
) -> Tuple[bytes, str]:
    """
    Process a single IPR file.
    Placeholder.
    """
    wb = load_workbook(input_path)
    # TODO: Full IPR combining + enrichment logic

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    if output_name is None:
        output_name = "IPR_Formatted.xlsx"

    return buf.getvalue(), output_name


def process_multiple_ipr_files(
    input_paths: List[str],
    cutsheet_path: Optional[str] = None,
) -> bytes:
    """Process multiple IPR files and return a ZIP."""
    import zipfile

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, path in enumerate(input_paths):
            data, name = process_ipr_validation(path, cutsheet_path)
            zf.writestr(name or f"ipr_{i}.xlsx", data)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()
