"""Excel date helpers (``openpyxl.utils.datetime``), implemented in Rust."""

from .._python_calamine import (
    MAC_EPOCH,
    WINDOWS_EPOCH,
    from_excel,
    from_ISO8601,
    serial_to_datetime,
    to_excel,
)

CALENDAR_WINDOWS_1900 = WINDOWS_EPOCH
CALENDAR_MAC_1904 = MAC_EPOCH

__all__ = (
    "CALENDAR_MAC_1904",
    "CALENDAR_WINDOWS_1900",
    "MAC_EPOCH",
    "WINDOWS_EPOCH",
    "from_ISO8601",
    "from_excel",
    "serial_to_datetime",
    "to_excel",
)
