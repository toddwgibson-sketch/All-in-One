# -*- coding: utf-8 -*-
"""
All In One Validation Formatter - Logic Package

This package contains all the rich processing logic.

Main entry point for most users:
    from logic import aio_logic as aio

Adding a new processor is straightforward:
1. Create new_logic.py with the standard interface
2. Import it here
3. Register it in aio_logic.py's PROCESSORS dict
"""

from .aio_logic import (
    inspect_cutsheet,
    DetectionResult,
    process_aio_validation,
    process_multiple_aio_validation,
    get_processor,
)

# Individual processors (exposed for direct use if needed)
from . import gfab_logic
from . import hops_logic
from . import cfab_logic
from . import ipr_logic
from . import lv_portal_logic

__all__ = [
    "inspect_cutsheet",
    "DetectionResult",
    "process_aio_validation",
    "process_multiple_aio_validation",
    "get_processor",
    "gfab_logic",
    "hops_logic",
    "cfab_logic",
    "ipr_logic",
    "lv_portal_logic",
]
