"""Read-only style classes of ``compatibility="openpyxl"`` workbooks
(the counterparts of ``openpyxl.styles``), for ``isinstance`` checks."""

from .._python_calamine import (
    Alignment,
    Border,
    Color,
    Font,
    GradientFill,
    PatternFill,
    Protection,
    Side,
    Stop,
)
from . import numbers
from .numbers import is_date_format

__all__ = (
    "Alignment",
    "Border",
    "Color",
    "Font",
    "GradientFill",
    "PatternFill",
    "Protection",
    "Side",
    "Stop",
    "is_date_format",
    "numbers",
)
