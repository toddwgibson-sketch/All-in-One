"""
All In One Validation Formatter - Logic Package

This package contains all the rich processing logic.

Main entry point for most users:
    from logic.aio_logic import inspect_cutsheet, process_aio_validation, process_multiple_aio_validation

Adding a new processor is straightforward:
1. Create new_logic.py with the standard interface (process_xxx_validation + process_multiple_xxx_files)
2. Import it here
3. Register it in aio_logic.py's PROCESSORS dict + improve inspect_cutsheet detection rules

This structure makes it easy for the team to contribute new hall/fabric support.
"""

from .aio_logic import (
    inspect_cutsheet,
    DetectionResult,
    process_aio_validation,
    process_multiple_aio_validation,
    get_processor,
)

# Also expose the individual processors for advanced use / testing
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
]
