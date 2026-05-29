# -*- coding: utf-8 -*-
"""
aio_logic.py

ALL-IN-ONE (AIO) Validation Formatter - Master Logic

This is the universal entry point for any hall / any fabric type.

The CUTSHEET is the source of truth for what kind of validation this is.
We inspect the cutsheet first, then dispatch to the correct rich processor.

IMPORTANT ARCHITECTURAL RULE:
    Do NOT use multiprocessing, ProcessPoolExecutor, or concurrent.futures
    in any processor. This app runs on Streamlit Cloud, which has limited
    and problematic support for multiprocessing. It also breaks Streamlit's
    session state model.

Public API:
    inspect_cutsheet(cutsheet_path) -> DetectionResult
    process_aio_validation(...) -> (bytes, filename)
    process_multiple_aio_validation(...) -> zip_bytes
"""

from __future__ import annotations

import io
import os
import zipfile
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any, Callable, Tuple

import pandas as pd
from openpyxl import load_workbook

# Import the individual processors
from . import gfab_logic
from . import hops_logic
from . import cfab_logic
from . import ipr_logic
from . import lv_portal_logic


# =============================================================================
# DETECTION RESULT
# =============================================================================

@dataclass
class DetectionResult:
    hall: str = "UNKNOWN"
    fabric_type: str = "UNKNOWN"
    source: str = "UNKNOWN"
    recommended_processor: str = "unknown"
    confidence: float = 0.0
    reasons: List[str] = field(default_factory=list)
    cutsheet_metadata: Dict[str, Any] = field(default_factory=dict)

    def summary(self) -> str:
        return (
            f"Hall: {self.hall} | Type: {self.fabric_type} | "
            f"Source: {self.source} | Processor: {self.recommended_processor} "
            f"(confidence: {self.confidence:.0%})"
        )


# =============================================================================
# CUTSHEET INSPECTION
# =============================================================================

def inspect_cutsheet(cutsheet_path: str) -> DetectionResult:
    """
    Inspect a cutsheet and figure out what kind of validation this is.
    This is the intelligence layer of the AIO.
    """
    result = DetectionResult()
    reasons = []

    try:
        wb = load_workbook(cutsheet_path, read_only=True, data_only=True)
        first_sheet = wb[wb.sheetnames[0]]
    except Exception as e:
        result.reasons.append(f"Could not read cutsheet: {e}")
        return result

    # Collect text for keyword scanning
    text_blob = ""
    sheet_names = [s.lower() for s in wb.sheetnames]

    for row in first_sheet.iter_rows(min_row=1, max_row=40, values_only=True):
        for cell in row:
            if cell:
                text_blob += " " + str(cell).lower()

    text_lower = text_blob.lower()
    fname = os.path.basename(cutsheet_path).lower()

    # Hall detection
    if "jpb19" in text_lower or "jpb19" in fname:
        result.hall = "JPB19"
        reasons.append("Found JPB19 indicators")
    elif "jbp15" in text_lower or "jbp15" in fname:
        result.hall = "JBP15"
        reasons.append("Found JBP15 indicators")

    # Fabric type detection (priority order)
    if any(k in text_lower for k in ["hops", "gpu", "ll dp mismatch (gpu)"]):
        result.fabric_type = "HOPS"
        result.recommended_processor = "hops"
        reasons.append("HOPS/GPU indicators found")

    elif any(k in text_lower for k in ["t3-t2", "t2-t1-t0", "cfab"]):
        result.fabric_type = "CFAB"
        result.recommended_processor = "cfab"
        reasons.append("CFAB style detected")

    elif any(k in text_lower for k in ["t0-to-host", "t0 to host", "qfabt0", "t0 switch port"]):
        result.fabric_type = "LV_PORTAL_T0"
        result.recommended_processor = "lv_portal"
        reasons.append("T0-to-Host / LV Portal style detected")

    elif any(k in text_lower for k in ["gfab", "full_path_lldp"]):
        if result.fabric_type == "UNKNOWN":
            result.fabric_type = "GFAB"
            result.recommended_processor = "gfab"
            reasons.append("GFAB indicators found")

    # Filename fallback for hall
    if result.hall == "UNKNOWN":
        if "jpb19" in fname:
            result.hall = "JPB19"
        elif "jbp15" in fname:
            result.hall = "JBP15"

    # Confidence
    confidence = 0.4
    if result.hall != "UNKNOWN":
        confidence += 0.3
    if result.fabric_type != "UNKNOWN":
        confidence += 0.3
    result.confidence = min(confidence, 1.0)

    result.reasons = reasons
    return result


# =============================================================================
# PROCESSOR REGISTRY
# =============================================================================

_PROCESSORS: Dict[str, Callable] = {
    "gfab": gfab_logic.process_gfab_validation,
    "hops": hops_logic.process_hops_validation,
    "cfab": cfab_logic.process_cfab_validation,
    "ipr": ipr_logic.process_ipr_validation,
    "lv_portal": lv_portal_logic.process_lv_portal_validation,
}


def get_processor(detection: DetectionResult) -> Optional[Callable]:
    name = detection.recommended_processor
    return _PROCESSORS.get(name)


# =============================================================================
# PUBLIC API
# =============================================================================

def process_aio_validation(
    input_path: str,
    cutsheet_path: str,
    override_processor: Optional[str] = None,
) -> Tuple[bytes, str]:
    processor_name = override_processor or "gfab"  # safe default
    processor = _PROCESSORS.get(processor_name)

    if processor is None:
        # Fallback to gfab
        processor = gfab_logic.process_gfab_validation
        processor_name = "gfab (fallback)"

    result = processor(input_path, cutsheet_path)
    if isinstance(result, tuple) and len(result) == 2:
        return (*result, {})   # normalize to (bytes, name, info)
    return result


def process_multiple_aio_validation(
    input_paths: List[str],
    cutsheet_path: str,
    override_processor: Optional[str] = None,
) -> bytes:
    processor_name = override_processor or "gfab"
    processor = _PROCESSORS.get(processor_name)

    if processor is None:
        processor = gfab_logic.process_multiple_files

    # Try to find a dedicated multi-file function on the processor module
    multi_fn = getattr(processor, "process_multiple_files", None)

    if callable(multi_fn):
        result = multi_fn(input_paths, cutsheet_path)
        if isinstance(result, (bytes, bytearray)):
            return (result, "AIO_Formatted_Reports.zip", {})
        return result

    # Fallback: process files one by one using the single-file function
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, path in enumerate(input_paths):
            data, name = process_aio_validation(path, cutsheet_path, override_processor)
            zf.writestr(name or f"report_{i}.xlsx", data)
    zip_buffer.seek(0)
    return (zip_buffer.getvalue(), "AIO_Formatted_Reports.zip", {})
