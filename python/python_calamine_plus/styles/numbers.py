"""Number format helpers (``openpyxl.styles.numbers``), implemented in Rust."""

from .._python_calamine import builtin_formats, is_date_format, is_timedelta_format

BUILTIN_FORMATS: dict[int, str] = builtin_formats()
BUILTIN_FORMATS_MAX_SIZE = 164
BUILTIN_FORMATS_REVERSE: dict[str, int] = {v: k for k, v in BUILTIN_FORMATS.items()}

__all__ = (
    "BUILTIN_FORMATS",
    "BUILTIN_FORMATS_MAX_SIZE",
    "BUILTIN_FORMATS_REVERSE",
    "is_date_format",
    "is_timedelta_format",
)
