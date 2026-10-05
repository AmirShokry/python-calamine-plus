"""Streaming Excel reader (python-calamine fork).

``utils``, ``styles`` and the ``openpyxl`` drop-in package mirror openpyxl's modules for
``load_workbook(..., compatibility="openpyxl")``.
"""

from . import openpyxl, styles, utils
from ._python_calamine import (
    CalamineCell,
    CalamineError,
    CalamineSheet,
    CalamineSheetStream,
    CalamineTable,
    CalamineWorkbook,
    CompatibilityNotSupported,
    PasswordError,
    SheetMetadata,
    SheetTypeEnum,
    SheetVisibleEnum,
    StreamingNotSupported,
    StreamInvalidated,
    TableNotFound,
    TablesNotLoaded,
    TablesNotSupported,
    WorkbookClosed,
    WorksheetNotFound,
    XmlError,
    ZipError,
    load_workbook,
)

__all__ = (
    "CalamineCell",
    "CalamineError",
    "CalamineSheet",
    "CalamineSheetStream",
    "CalamineTable",
    "CalamineWorkbook",
    "CompatibilityNotSupported",
    "PasswordError",
    "SheetMetadata",
    "SheetTypeEnum",
    "SheetVisibleEnum",
    "StreamingNotSupported",
    "StreamInvalidated",
    "TableNotFound",
    "TablesNotLoaded",
    "TablesNotSupported",
    "WorkbookClosed",
    "WorksheetNotFound",
    "XmlError",
    "ZipError",
    "load_workbook",
    "openpyxl",
    "styles",
    "utils",
)
