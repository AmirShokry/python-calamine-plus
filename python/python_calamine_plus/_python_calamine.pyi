# Some documentations from upstream under MIT License. See authors in https://github.com/tafia/calamine
from __future__ import annotations

import datetime
import enum
import os
import types
import typing

@typing.type_check_only
class ReadBuffer(typing.Protocol):
    def seek(self, __offset: int, __whence: int = ...) -> int: ...
    def read(self, __size: int = ...) -> bytes | None: ...

@typing.final
class SheetTypeEnum(enum.Enum):
    WorkSheet = ...
    DialogSheet = ...
    MacroSheet = ...
    ChartSheet = ...
    Vba = ...

@typing.final
class SheetVisibleEnum(enum.Enum):
    Visible = ...
    """Visible."""
    Hidden = ...
    """Hidden."""
    VeryHidden = ...
    """The sheet is hidden and cannot be displayed using the user interface. It is supported only by Excel formats."""

@typing.final
class SheetMetadata:
    name: str
    """Name of sheet."""
    typ: SheetTypeEnum
    """Type of sheet.

    Only Excel formats support this. Default value for ODS is `WorkSheet`.
    """
    visible: SheetVisibleEnum
    """Visible of sheet."""

    def __new__(
        cls, name: str, typ: SheetTypeEnum, visible: SheetVisibleEnum
    ) -> SheetMetadata: ...

@typing.final
class CalamineSheet:
    name: str
    @property
    def height(self) -> int:
        """Get the row height of a sheet data.

        The height is defined as the number of rows between the start and end positions.
        """

    @property
    def width(self) -> int:
        """Get the column width of a sheet data.

        The width is defined as the number of columns between the start and end positions.
        """

    @property
    def total_height(self) -> int: ...
    @property
    def total_width(self) -> int: ...
    @property
    def start(self) -> tuple[int, int] | None:
        """Get top left cell position of a sheet data."""

    @property
    def end(self) -> tuple[int, int] | None:
        """Get bottom right cell position of a sheet data."""

    def to_python(
        self, skip_empty_area: bool = True, nrows: int | None = None
    ) -> list[
        list[
            int
            | float
            | str
            | bool
            | datetime.time
            | datetime.date
            | datetime.datetime
            | datetime.timedelta
        ]
    ]:
        """Returning data from sheet as list of lists.

        Args:
            skip_empty_area (bool):
                By default, calamine skips empty rows/cols before data.
                For suppress this behaviour, set `skip_empty_area` to `False`.
        """

    def iter_rows(
        self,
    ) -> typing.Iterator[
        list[
            int
            | float
            | str
            | bool
            | datetime.time
            | datetime.date
            | datetime.datetime
            | datetime.timedelta
        ]
    ]:
        """Returning data from sheet as iterator of lists."""

    @property
    def merged_cell_ranges(
        self,
    ) -> list[tuple[tuple[int, int], tuple[int, int]]] | None:
        """Return a copy of merged cell ranges.

        Support only for xlsx/xls.

        Returns:
            list of merged cell ranges (tuple[start coordinate, end coordinate]) or None for unsupported format
        """

_CellValueT: typing.TypeAlias = (
    int
    | float
    | str
    | bool
    | datetime.time
    | datetime.date
    | datetime.datetime
    | datetime.timedelta
)
_MergedCellRange: typing.TypeAlias = tuple[tuple[int, int], tuple[int, int]]

@typing.final
class CalamineCell:
    """A cell yielded by a stream started with ``rich=True``."""

    @property
    def value(self) -> _CellValueT: ...
    @property
    def row(self) -> int:
        """0-based row index."""

    @property
    def column(self) -> int:
        """0-based column index."""

    @property
    def style_id(self) -> int:
        """Index into ``CalamineSheetStream.styles`` (0 is the workbook default)."""

    @property
    def style(self) -> dict[str, typing.Any] | None:
        """The cell format: ``number_format``, ``number_format_id``, ``font``,
        ``fill``, ``border``, ``alignment`` and ``protection``.

        The same dict object is shared by every cell with this format, so treat
        it as read-only. Colors are dicts with ``rgb``/``theme``/``tint``/
        ``indexed``/``auto`` keys as stored in the file (themes are not resolved).
        """

    @property
    def comment(self) -> dict[str, str | None] | None:
        """``{"author": ..., "text": ...}`` or ``None``."""

    @property
    def merged_range(self) -> _MergedCellRange | None:
        """``((first_row, first_col), (last_row, last_col))`` if the cell is merged."""

    @property
    def formula(self) -> str | None:
        """Formula text with a leading ``=`` (e.g. ``"=SUM(A1:A3)"``), or ``None``.

        Shared formulas are expanded per cell. ``value`` holds the result saved
        with the file; it is ``""`` if the writer never calculated it.
        """

    @property
    def hyperlink(self) -> dict[str, typing.Any] | None:
        """``{"range", "target", "location", "display", "tooltip"}`` or ``None``.

        ``target`` is an external URL/file, ``location`` an in-workbook reference.
        """

    @property
    def data_type(self) -> str:
        """openpyxl's data type of the value: ``n``, ``s``, ``b``, ``e`` or ``d``."""

    @property
    def is_date(self) -> bool:
        """The value is a date, time or duration."""

    @property
    def coordinate(self) -> str:
        """Excel reference, e.g. ``"B3"``."""

    @property
    def formula_type(self) -> str | None:
        """``normal``, ``shared``, ``array`` or ``dataTable``; ``None`` without a formula."""

    @property
    def formula_range(self) -> str | None:
        """Range of an array formula (openpyxl's ``ArrayFormula.ref``), on its anchor cell."""

    @property
    def formula_attributes(self) -> dict[str, typing.Any] | None:
        """Attributes of an array / data table formula (``ref``, ``r1``, ``dt2_d``, ...)."""

    @property
    def rich_text(self) -> list[dict[str, typing.Any]] | None:
        """Formatting runs ``[{"text": ..., "font": {...} | None}, ...]`` for rich
        text cells (shared or inline strings), else ``None``. ``value`` is the
        concatenated text."""

@typing.final
class CalamineSheetStream:
    """Row-by-row streaming reader over an xlsx worksheet.

    Cells are read lazily from the file, so the sheet is never fully loaded
    into memory. Rows start at row 0 / column 0; missing cells are ``""``.

    With ``rich=True`` each row is a list of ``CalamineCell`` (value, style,
    comment, merged range, formula, hyperlink, rich text) instead of plain
    values, and the stream exposes the sheet's metadata. The rows are the same,
    except that rich streams also yield rows whose only content is a formula
    without a saved result (files never calculated by Excel).

    Rows and columns are 0-based; ``min_row``/``max_row``/``min_col``/``max_col``
    (inclusive) restrict the stream, and it stops reading after ``max_row``.

    The stream reuses the workbook's already-loaded shared strings. A workbook
    has at most one active stream: starting another stream, loading a sheet or
    table, or closing the workbook invalidates it (``StreamInvalidated`` /
    ``WorkbookClosed`` on the next read).
    """

    @property
    def name(self) -> str: ...
    @property
    def rich(self) -> bool: ...
    @property
    def merged_cell_ranges(self) -> list[_MergedCellRange] | None:
        """Merged ranges of the sheet (``rich=True`` only, else ``None``)."""

    @property
    def comments(self) -> dict[tuple[int, int], dict[str, str | None]] | None:
        """All comments of the sheet keyed by ``(row, column)``, including those
        outside the streamed rows (``rich=True`` only, else ``None``)."""

    @property
    def styles(self) -> list[dict[str, typing.Any]] | None:
        """The workbook's cell formats indexed by ``CalamineCell.style_id``
        (``rich=True`` only, else ``None``)."""

    @property
    def hyperlinks(self) -> list[dict[str, typing.Any]] | None:
        """All hyperlinks of the sheet (``rich=True`` only, else ``None``)."""

    @property
    def conditional_formats(self) -> list[dict[str, typing.Any]] | None:
        """Conditional formatting: ``[{"ranges": [...], "rules": [{"type",
        "priority", "operator", "formulas", "stop_if_true", "text", "dxf_id",
        "style", "values", "colors", "icon_set", "extra"}, ...]}, ...]``.
        ``style`` is the differential format applied when the rule matches
        (``rich=True`` only, else ``None``)."""

    @property
    def data_validations(self) -> list[dict[str, typing.Any]] | None:
        """Data validation rules: ``{"ranges", "type", "operator", "formula1",
        "formula2", "allow_blank", "in_cell_dropdown", "show_input_message",
        "show_error_message", "error_style", "error_title", "error",
        "prompt_title", "prompt"}`` (``rich=True`` only, else ``None``)."""

    @property
    def column_dimensions(self) -> list[dict[str, typing.Any]] | None:
        """Column spans with ``{"min", "max", "width", "custom_width", "hidden",
        "outline_level", "collapsed", "style_id"}`` (0-based, inclusive;
        ``rich=True`` only, else ``None``)."""

    @property
    def sheet_settings(self) -> dict[str, typing.Any] | None:
        """``freeze_panes`` (e.g. ``"B2"``), ``frozen_rows``, ``frozen_columns``,
        ``auto_filter``, ``auto_filter_details`` (filter columns, sort),
        ``sort_state``, ``protection``, ``page_setup``, ``page_margins``,
        ``print_options``, ``print_area``, ``print_titles``, ``header_footer``,
        ``sheet_view``, ``views`` (all views with pane and selections),
        ``custom_views``, ``sheet_properties``, ``outline_properties``,
        ``page_setup_properties``, ``sheet_format``, ``tab_color``,
        ``default_row_height`` and ``default_column_width`` (``rich=True`` only,
        else ``None``). Attribute dicts use snake_case keys with typed values."""

    @property
    def dimension(self) -> str | None:
        """``<dimension>`` reference declared by the writer, e.g. ``"A1:H100"``
        (``rich=True`` only)."""

    @property
    def tables(self) -> list[dict[str, typing.Any]] | None:
        """Tables (list objects): attributes (``name``, ``display_name``, ``ref``,
        ``header_row_count``, ...), ``columns``, ``style`` and ``auto_filter``
        (``rich=True`` only)."""

    @property
    def row_dimensions(self) -> dict[int, dict[str, typing.Any]] | None:
        """All rows with non-default height / hidden / outline / style keyed by row
        index, including rows outside the streamed data (``rich=True`` only)."""

    @property
    def row_breaks(self) -> list[dict[str, typing.Any]] | None:
        """Manual row page breaks (``id``, ``min``, ``max``, ``man``) (``rich=True`` only)."""

    @property
    def column_breaks(self) -> list[dict[str, typing.Any]] | None:
        """Manual column page breaks (``rich=True`` only)."""

    @property
    def scenarios(self) -> dict[str, typing.Any] | None:
        """What-if scenarios, or ``None`` (also when not ``rich=True``)."""

    @property
    def row_info(self) -> dict[str, typing.Any] | None:
        """``{"row", "height", "custom_height", "hidden", "outline_level",
        "collapsed", "style_id"}`` of the row last yielded (``rich=True`` only)."""

    def __iter__(self) -> CalamineSheetStream: ...
    def __next__(self) -> list[_CellValueT] | list[CalamineCell]: ...

@typing.final
class CalamineTable:
    name: str
    """Get the name of the table."""
    sheet: str
    """Get the name of the parent worksheet for a table."""
    columns: list[str]
    """Get the header names of the table columns.

    In Excel table headers can be hidden but the table will still have
    column header names.
    """
    @property
    def height(self) -> int:
        """Get the row height of a table data.

        The height is defined as the number of rows between the start and end positions.
        """

    @property
    def width(self) -> int:
        """Get the column width of a table data.

        The width is defined as the number of columns between the start and end positions.
        """

    @property
    def start(self) -> tuple[int, int] | None:
        """Get top left cell position of a table data."""

    @property
    def end(self) -> tuple[int, int] | None:
        """Get bottom right cell position of a table data."""

    def to_python(
        self,
    ) -> list[
        list[
            int
            | float
            | str
            | bool
            | datetime.time
            | datetime.date
            | datetime.datetime
            | datetime.timedelta
        ]
    ]:
        """Returning data from table as list of lists."""

@typing.final
class CalamineWorkbook:
    path: str | None
    """Path to file. `None` if bytes was loaded."""
    sheet_names: list[str]
    """All sheet names of this workbook, in workbook order."""
    sheets_metadata: list[SheetMetadata]
    """All sheets metadata of this workbook, in workbook order."""
    table_names: list[str] | None
    """All table names of this workbook."""
    defined_names: list[dict[str, typing.Any]]
    """Defined names: ``{"name", "value", "sheet", "hidden", "comment"}``. ``sheet`` is the
    sheet a name is local to (``None`` for workbook scope; always ``None`` for
    non-xlsx formats)."""
    properties: dict[str, typing.Any]
    """Document properties (``creator``, ``title``, ``created``, ``modified``,
    ``last_modified_by``, ``company``, ...); dates are naive UTC ``datetime``.
    Empty for non-xlsx formats."""
    custom_properties: dict[str, typing.Any]
    """Custom document properties, typed (str / int / float / bool / datetime)."""
    epoch_1904: bool
    """Dates use the 1904 date system."""
    code_name: str | None
    """VBA code name of the workbook."""
    workbook_properties: dict[str, typing.Any]
    """``<workbookPr>`` attributes (``date1904``, ``code_name``, ...)."""
    active_sheet_index: int
    """Index of the active sheet (openpyxl's ``wb.active``)."""
    workbook_views: list[dict[str, typing.Any]]
    """Workbook window views (``active_tab``, ``first_sheet``, ``visibility``, ...)."""
    calculation: dict[str, typing.Any] | None
    """Calculation properties (``calc_id``, ``full_calc_on_load``, ...)."""
    workbook_protection: dict[str, typing.Any] | None
    """Workbook protection (``lock_structure``, ``lock_windows``, hashes)."""
    file_version: dict[str, typing.Any] | None
    """``<fileVersion>`` attributes (``app_name``, ``last_edited``, ...)."""
    theme_colors: dict[str, str]
    """Theme color scheme: ``{"dk1": "000000", "lt1": "FFFFFF", "accent1": ...}``."""
    theme_xml: bytes | None
    """Raw theme XML (openpyxl's ``loaded_theme``)."""
    is_template: bool
    """The file is a template (``.xltx`` / ``.xltm``)."""
    has_macros: bool
    """The file is macro-enabled (``.xlsm`` / ``.xltm``)."""
    named_styles: list[dict[str, typing.Any]]
    """Named cell styles: ``{"name", "builtin_id", "hidden"}``."""
    @classmethod
    def from_object(
        cls, path_or_filelike: str | os.PathLike | ReadBuffer, load_tables: bool = False
    ) -> "CalamineWorkbook":
        """Determining type of pyobject and reading from it.

        Args:
            path_or_filelike (str | os.PathLike | ReadBuffer): path to file or IO (must implement read/seek methods).
            load_tables (bool): load Excel tables (supported for XLSX only).
        """

    @classmethod
    def from_path(
        cls, path: str | os.PathLike, load_tables: bool = False
    ) -> "CalamineWorkbook":
        """Reading file from path.

        Args:
            path (str | os.PathLike): path to file.
            load_tables (bool): load Excel tables (supported for XLSX only).
        """

    @classmethod
    def from_filelike(
        cls, filelike: ReadBuffer, load_tables: bool = False
    ) -> "CalamineWorkbook":
        """Reading file from IO.

        Args:
            filelike : IO (must implement read/seek methods).
            load_tables (bool): load Excel tables (supported for XLSX only).
        """

    def close(self) -> None:
        """Close the workbook.

        Drop internal rust structure from workbook (and close the file under the hood).
        `get_sheet_by_name`/`get_sheet_by_index` will raise WorkbookClosed after calling that method.

        Raises:
            WorkbookClosed: If workbook already closed.
        """

    def __enter__(self) -> "CalamineWorkbook": ...
    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: types.TracebackType | None,
    ) -> None: ...
    def get_sheet_by_name(self, name: str) -> CalamineSheet:
        """Get worksheet by name.

        Args:
            name(str): name of worksheet

        Returns:
            CalamineSheet

        Raises:
            WorkbookClosed: If workbook already closed.
            WorksheetNotFound: If worksheet not found in workbook.
        """

    def get_sheet_by_index(self, index: int) -> CalamineSheet:
        """Get worksheet by index.

        Args:
            index(int): index of worksheet

        Returns:
            CalamineSheet

        Raises:
            WorkbookClosed: If workbook already closed.
            WorksheetNotFound: If worksheet not found in workbook.
        """

    def stream_sheet_by_name(
        self,
        name: str,
        rich: bool = False,
        min_row: int = 0,
        max_row: int | None = None,
        min_col: int = 0,
        max_col: int | None = None,
        ref: str | None = None,
    ) -> CalamineSheetStream:
        """Stream worksheet rows by name without loading the whole sheet.

        Only supported for xlsx workbooks. Invalidates any active stream.

        With ``rich=True`` rows are lists of ``CalamineCell`` with style, comment,
        merged range, formula, hyperlink and rich text, and the stream exposes the
        sheet's metadata. That metadata (merged ranges, hyperlinks, conditional
        formats, ...) is stored after the cells, so it costs one extra pass over
        the sheet when the stream starts.

        Args:
            name(str): name of worksheet
            rich(bool): yield ``CalamineCell`` objects and expose sheet metadata
            min_row(int): first row to stream (0-based)
            max_row(int | None): last row to stream (inclusive); reading stops after it
            min_col(int): first column to stream (0-based)
            max_col(int | None): last column (inclusive); rows are padded to it
            ref(str | None): A1-style range (``"B2:D10"``) or cell (``"B2"``); overrides the bounds

        Returns:
            CalamineSheetStream

        Raises:
            WorkbookClosed: If workbook already closed.
            WorksheetNotFound: If worksheet not found in workbook.
            StreamingNotSupported: If the workbook is not xlsx / xlsm (or xltx / xltm).
        """

    def stream_sheet_by_index(
        self,
        index: int,
        rich: bool = False,
        min_row: int = 0,
        max_row: int | None = None,
        min_col: int = 0,
        max_col: int | None = None,
        ref: str | None = None,
    ) -> CalamineSheetStream:
        """Stream worksheet rows by index without loading the whole sheet.

        Only supported for xlsx workbooks. Invalidates any active stream.

        With ``rich=True`` rows are lists of ``CalamineCell`` with style, comment,
        merged range, formula, hyperlink and rich text, and the stream exposes the
        sheet's metadata. That metadata (merged ranges, hyperlinks, conditional
        formats, ...) is stored after the cells, so it costs one extra pass over
        the sheet when the stream starts.

        Args:
            index(int): index of worksheet
            rich(bool): yield ``CalamineCell`` objects and expose sheet metadata
            min_row(int): first row to stream (0-based)
            max_row(int | None): last row to stream (inclusive); reading stops after it
            min_col(int): first column to stream (0-based)
            max_col(int | None): last column (inclusive); rows are padded to it
            ref(str | None): A1-style range (``"B2:D10"``) or cell (``"B2"``); overrides the bounds

        Returns:
            CalamineSheetStream

        Raises:
            WorkbookClosed: If workbook already closed.
            WorksheetNotFound: If worksheet not found in workbook.
            StreamingNotSupported: If the workbook is not xlsx / xlsm (or xltx / xltm).
        """

    def get_table_by_name(self, name: str) -> CalamineTable:
        """Get table by name.

        Args:
            name(str): name of table

        Returns:
            CalamineTable

        Raises:
            WorkbookClosed: If workbook already closed.
            TableNotFound: If table not found in workbook.
        """

class CalamineError(Exception): ...
class PasswordError(CalamineError): ...
class WorksheetNotFound(CalamineError): ...
class XmlError(CalamineError): ...
class ZipError(CalamineError): ...
class WorkbookClosed(CalamineError): ...
class TablesNotLoaded(CalamineError): ...
class TablesNotSupported(CalamineError): ...
class TableNotFound(CalamineError): ...
class StreamingNotSupported(CalamineError): ...
class StreamInvalidated(CalamineError): ...

def load_workbook(
    path_or_filelike: str | os.PathLike | ReadBuffer, load_tables: bool = False
) -> CalamineWorkbook:
    """Determining type of pyobject and reading from it.

    Args:
        path_or_filelike (str | os.PathLike | ReadBuffer): path to file or IO (must implement read/seek methods).
        load_tables (bool): load Excel tables (supported for XLSX only).
    """

__all__ = [
    "CalamineCell",
    "CalamineError",
    "CalamineSheet",
    "CalamineSheetStream",
    "CalamineTable",
    "CalamineWorkbook",
    "PasswordError",
    "SheetMetadata",
    "SheetTypeEnum",
    "SheetVisibleEnum",
    "StreamInvalidated",
    "StreamingNotSupported",
    "TableNotFound",
    "TablesNotLoaded",
    "TablesNotSupported",
    "WorkbookClosed",
    "WorksheetNotFound",
    "XmlError",
    "ZipError",
    "load_workbook",
]
