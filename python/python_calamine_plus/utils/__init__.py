"""openpyxl's utilities (``openpyxl.utils``), implemented in Rust.

``from python_calamine_plus.utils import get_column_letter`` works like
``from openpyxl.utils import get_column_letter``.
"""

from .._python_calamine import (
    CellCoordinatesException,
    absolute_coordinate,
    cols_from_range,
    column_index_from_string,
    coordinate_from_row_col,
    coordinate_from_string,
    coordinate_from_tuple,
    coordinate_to_tuple,
    get_column_interval,
    get_column_letter,
    quote_sheetname,
    range_boundaries,
    range_to_tuple,
    rows_from_range,
    split_print_areas,
)
from . import cell, datetime, exceptions

__all__ = (
    "CellCoordinatesException",
    "absolute_coordinate",
    "cell",
    "cols_from_range",
    "column_index_from_string",
    "coordinate_from_row_col",
    "coordinate_from_string",
    "coordinate_from_tuple",
    "coordinate_to_tuple",
    "datetime",
    "exceptions",
    "get_column_interval",
    "get_column_letter",
    "quote_sheetname",
    "range_boundaries",
    "range_to_tuple",
    "rows_from_range",
    "split_print_areas",
)
