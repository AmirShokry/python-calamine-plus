"""Coordinate helpers (``openpyxl.utils.cell``), implemented in Rust."""

from .._python_calamine import (
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

__all__ = (
    "absolute_coordinate",
    "cols_from_range",
    "column_index_from_string",
    "coordinate_from_row_col",
    "coordinate_from_string",
    "coordinate_from_tuple",
    "coordinate_to_tuple",
    "get_column_interval",
    "get_column_letter",
    "quote_sheetname",
    "range_boundaries",
    "range_to_tuple",
    "rows_from_range",
    "split_print_areas",
)
