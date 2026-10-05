"""Drop-in replacement for openpyxl's reading API.

Change ``import openpyxl`` to ``from python_calamine_plus import openpyxl``: this
``load_workbook`` is ``python_calamine_plus.load_workbook(...,
compatibility="openpyxl")`` with openpyxl's signature, and ``utils`` / ``styles``
mirror ``openpyxl.utils`` / ``openpyxl.styles``. Workbooks are read-only.
"""

from __future__ import annotations

import os
import typing

from .. import _python_calamine
from .._python_calamine import Workbook
from . import styles, utils

if typing.TYPE_CHECKING:
    from .._python_calamine import ReadBuffer


def load_workbook(
    filename: str | os.PathLike[str] | ReadBuffer,
    read_only: bool = False,
    keep_vba: bool = False,
    data_only: bool = False,
    keep_links: bool = True,
    rich_text: bool = False,
    *,
    formula_and_value: bool = False,
    read_comments: bool | None = None,
    read_hyperlinks: bool | None = None,
    read_merged_cells: bool | None = None,
) -> Workbook:
    """Open a workbook like ``openpyxl.load_workbook``.

    ``keep_vba``, ``keep_links`` and ``rich_text`` must keep their defaults: the first
    two only matter when saving, and rich text values are not supported.

    Extensions (keyword-only, not in openpyxl): ``formula_and_value=True`` gives cells
    ``formula`` and ``cached_value`` from the same pass; with ``read_only=True``,
    ``read_comments`` / ``read_hyperlinks`` / ``read_merged_cells`` give cells
    ``comment`` / ``hyperlink`` / ``merged_range`` like the native rich stream.
    """
    extras = {
        name: value
        for name, value in (
            ("read_comments", read_comments),
            ("read_hyperlinks", read_hyperlinks),
            ("read_merged_cells", read_merged_cells),
        )
        if value is not None
    }
    return _python_calamine.load_workbook(
        filename,
        compatibility="openpyxl",
        read_only=read_only,
        keep_vba=keep_vba,
        data_only=data_only,
        keep_links=keep_links,
        rich_text=rich_text,
        formula_and_value=formula_and_value,
        **extras,
    )


__all__ = ("Workbook", "load_workbook", "styles", "utils")
