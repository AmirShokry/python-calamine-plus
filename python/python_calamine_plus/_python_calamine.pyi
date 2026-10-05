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

@typing.overload
def load_workbook(
    path_or_filelike: str | os.PathLike | ReadBuffer,
    load_tables: bool = False,
    *,
    compatibility: None = None,
) -> CalamineWorkbook:
    """Determining type of pyobject and reading from it.

    Args:
        path_or_filelike (str | os.PathLike | ReadBuffer): path to file or IO (must implement read/seek methods).
        load_tables (bool): load Excel tables (supported for XLSX only).
    """

@typing.overload
def load_workbook(
    path_or_filelike: str | os.PathLike | ReadBuffer,
    load_tables: bool = False,
    *,
    compatibility: typing.Literal["openpyxl"],
    read_only: bool | None = None,
    data_only: bool | None = None,
    keep_vba: bool | None = None,
    keep_links: bool | None = None,
    rich_text: bool | None = None,
    formula_and_value: bool | None = None,
    read_comments: bool | None = None,
    read_hyperlinks: bool | None = None,
    read_merged_cells: bool | None = None,
) -> Workbook:
    """Open a workbook with openpyxl's reading behavior.

    With ``compatibility="openpyxl"`` the result is a read-only, openpyxl-shaped
    ``Workbook`` (xlsx / xlsm / xltx / xltm only): ``wb["Sheet"]``, ``ws["A1"]``,
    ``ws.iter_rows(min_row=1, values_only=True)``, ``cell.font.b``, ... with openpyxl's
    values (``None`` for empty cells, ``datetime`` for dates, formula text unless
    ``data_only=True``) and openpyxl's ``read_only`` semantics. ``keep_vba`` /
    ``keep_links`` / ``rich_text`` must keep openpyxl's defaults.

    Extensions (not in openpyxl): ``formula_and_value=True`` gives every cell
    ``formula`` and ``cached_value`` read in the same pass (``value`` still follows
    ``data_only``). With ``read_only=True``, ``read_comments`` / ``read_hyperlinks`` /
    ``read_merged_cells`` give read-only cells ``comment`` / ``hyperlink`` /
    ``merged_range`` like the native ``rich=True`` stream (normal mode always reads
    them, like openpyxl; ``False`` there raises ``ValueError``).

    Raises:
        CompatibilityNotSupported: for .xls / .xlsb / .ods files or ``rich_text=True``.
        TypeError: openpyxl options without ``compatibility="openpyxl"``.
    """

# ---- compatibility="openpyxl" ------------------------------------------------------

_OpenpyxlValue: typing.TypeAlias = (
    int
    | float
    | str
    | bool
    | datetime.datetime
    | datetime.date
    | datetime.time
    | datetime.timedelta
    | ArrayFormula
    | DataTableFormula
    | None
)

class CompatibilityNotSupported(CalamineError): ...
class CellCoordinatesException(Exception): ...

@typing.final
class Color:
    """``type`` is ``rgb``, ``theme``, ``indexed`` or ``auto``; only that attribute is
    set (theme / indexed colors are not resolved to RGB)."""

    @property
    def rgb(self) -> str | None: ...
    @property
    def indexed(self) -> int | None: ...
    @property
    def theme(self) -> int | None: ...
    @property
    def auto(self) -> bool | None: ...
    @property
    def tint(self) -> float: ...
    @property
    def type(self) -> str: ...
    @property
    def value(self) -> str | int | bool | None: ...
    @property
    def index(self) -> str | int | bool | None: ...

@typing.final
class Font:
    @property
    def name(self) -> str | None: ...
    @property
    def sz(self) -> float | None: ...
    @property
    def size(self) -> float | None: ...
    @property
    def b(self) -> bool: ...
    @property
    def bold(self) -> bool: ...
    @property
    def i(self) -> bool: ...
    @property
    def italic(self) -> bool: ...
    @property
    def u(self) -> str | None: ...
    @property
    def underline(self) -> str | None: ...
    @property
    def strike(self) -> bool | None: ...
    @property
    def strikethrough(self) -> bool | None: ...
    @property
    def vertAlign(self) -> str | None: ...
    @property
    def color(self) -> Color | None: ...
    @property
    def family(self) -> float | None: ...
    @property
    def charset(self) -> int | None: ...
    @property
    def scheme(self) -> str | None: ...
    @property
    def outline(self) -> bool | None: ...
    @property
    def shadow(self) -> bool | None: ...
    @property
    def condense(self) -> bool | None: ...
    @property
    def extend(self) -> bool | None: ...

@typing.final
class PatternFill:
    @property
    def patternType(self) -> str | None: ...
    @property
    def fill_type(self) -> str | None: ...
    @property
    def fgColor(self) -> Color: ...
    @property
    def start_color(self) -> Color: ...
    @property
    def bgColor(self) -> Color: ...
    @property
    def end_color(self) -> Color: ...
    @property
    def tagname(self) -> str: ...

@typing.final
class Stop:
    @property
    def color(self) -> Color: ...
    @property
    def position(self) -> float: ...

@typing.final
class GradientFill:
    @property
    def type(self) -> str: ...
    @property
    def fill_type(self) -> str: ...
    @property
    def degree(self) -> float: ...
    @property
    def left(self) -> float: ...
    @property
    def right(self) -> float: ...
    @property
    def top(self) -> float: ...
    @property
    def bottom(self) -> float: ...
    @property
    def stop(self) -> list[Stop]: ...
    @property
    def tagname(self) -> str: ...

@typing.final
class Side:
    @property
    def style(self) -> str | None: ...
    @property
    def border_style(self) -> str | None: ...
    @property
    def color(self) -> Color | None: ...

@typing.final
class Border:
    """Edges absent from the file are ``None``."""

    @property
    def left(self) -> Side | None: ...
    @property
    def right(self) -> Side | None: ...
    @property
    def top(self) -> Side | None: ...
    @property
    def bottom(self) -> Side | None: ...
    @property
    def diagonal(self) -> Side | None: ...
    @property
    def vertical(self) -> Side | None: ...
    @property
    def horizontal(self) -> Side | None: ...
    @property
    def start(self) -> None: ...
    @property
    def end(self) -> None: ...
    @property
    def diagonal_direction(self) -> None: ...
    @property
    def diagonalUp(self) -> bool: ...
    @property
    def diagonalDown(self) -> bool: ...
    @property
    def outline(self) -> bool: ...

@typing.final
class Alignment:
    @property
    def horizontal(self) -> str | None: ...
    @property
    def vertical(self) -> str | None: ...
    @property
    def textRotation(self) -> int: ...
    @property
    def text_rotation(self) -> int: ...
    @property
    def wrapText(self) -> bool | None: ...
    @property
    def wrap_text(self) -> bool | None: ...
    @property
    def shrinkToFit(self) -> bool | None: ...
    @property
    def shrink_to_fit(self) -> bool | None: ...
    @property
    def indent(self) -> float: ...
    @property
    def relativeIndent(self) -> float: ...
    @property
    def justifyLastLine(self) -> bool | None: ...
    @property
    def readingOrder(self) -> float: ...

@typing.final
class Protection:
    @property
    def locked(self) -> bool: ...
    @property
    def hidden(self) -> bool: ...

@typing.final
class Comment:
    @property
    def text(self) -> str: ...
    @property
    def content(self) -> str: ...
    @property
    def author(self) -> str | None: ...
    @property
    def width(self) -> int: ...
    @property
    def height(self) -> int: ...

@typing.final
class Hyperlink:
    @property
    def ref(self) -> str:
        """Coordinate of the cell the link is bound to."""

    @property
    def target(self) -> str | None: ...
    @property
    def location(self) -> str | None: ...
    @property
    def tooltip(self) -> str | None: ...
    @property
    def display(self) -> str | None: ...
    @property
    def id(self) -> None:
        """The relationship id is not read (always ``None``)."""

@typing.final
class ArrayFormula:
    @property
    def ref(self) -> str | None: ...
    @property
    def text(self) -> str: ...
    @property
    def t(self) -> str: ...
    def __iter__(self) -> typing.Iterator[tuple[str, str]]: ...

@typing.final
class DataTableFormula:
    """Attribute values as stored (``"1"``), ``False`` / ``None`` when absent."""

    @property
    def t(self) -> str: ...
    @property
    def ref(self) -> str | None: ...
    @property
    def ca(self) -> str | bool: ...
    @property
    def dt2D(self) -> str | bool: ...
    @property
    def dtr(self) -> str | bool: ...
    @property
    def r1(self) -> str | None: ...
    @property
    def r2(self) -> str | None: ...
    @property
    def del1(self) -> str | bool: ...
    @property
    def del2(self) -> str | bool: ...

@typing.final
class MergedCellRange:
    @property
    def min_row(self) -> int: ...
    @property
    def min_col(self) -> int: ...
    @property
    def max_row(self) -> int: ...
    @property
    def max_col(self) -> int: ...
    @property
    def bounds(self) -> tuple[int, int, int, int]:
        """``(min_col, min_row, max_col, max_row)``."""

    @property
    def coord(self) -> str: ...
    @property
    def ref(self) -> str: ...
    @property
    def title(self) -> None: ...
    @property
    def size(self) -> dict[str, int]: ...
    @property
    def top(self) -> list[tuple[int, int]]: ...
    @property
    def bottom(self) -> list[tuple[int, int]]: ...
    @property
    def left(self) -> list[tuple[int, int]]: ...
    @property
    def right(self) -> list[tuple[int, int]]: ...
    @property
    def cells(self) -> typing.Iterator[tuple[int, int]]: ...
    @property
    def rows(self) -> typing.Iterator[list[tuple[int, int]]]: ...
    @property
    def cols(self) -> typing.Iterator[list[tuple[int, int]]]: ...
    @property
    def start_cell(self) -> Cell: ...
    def __contains__(self, coord: str, /) -> bool: ...

@typing.final
class MultiCellRange:
    @property
    def ranges(self) -> list[MergedCellRange]:
        """In file order (openpyxl uses a set, whose order is not stable)."""

    def sorted(self) -> list[MergedCellRange]: ...
    def __contains__(self, coord: str | MergedCellRange, /) -> bool: ...
    def __iter__(self) -> typing.Iterator[MergedCellRange]: ...
    def __len__(self) -> int: ...

@typing.final
class RowDimension:
    @property
    def index(self) -> int: ...
    @property
    def r(self) -> int: ...
    @property
    def ht(self) -> float | None: ...
    @property
    def height(self) -> float | None: ...
    @property
    def customHeight(self) -> bool: ...
    @property
    def hidden(self) -> bool: ...
    @property
    def outlineLevel(self) -> int: ...
    @property
    def outline_level(self) -> int: ...
    @property
    def collapsed(self) -> bool: ...
    @property
    def customFormat(self) -> bool: ...
    @property
    def thickBot(self) -> bool: ...
    @property
    def thickTop(self) -> bool: ...

@typing.final
class ColumnDimension:
    @property
    def index(self) -> str: ...
    @property
    def width(self) -> float: ...
    @property
    def customWidth(self) -> bool: ...
    @property
    def bestFit(self) -> bool: ...
    @property
    def auto_size(self) -> bool: ...
    @property
    def hidden(self) -> bool: ...
    @property
    def outlineLevel(self) -> int: ...
    @property
    def outline_level(self) -> int: ...
    @property
    def collapsed(self) -> bool: ...
    @property
    def min(self) -> int | None: ...
    @property
    def max(self) -> int | None: ...
    @property
    def range(self) -> str | None: ...

_DimT = typing.TypeVar("_DimT", RowDimension, ColumnDimension)

@typing.final
class DimensionHolder(typing.Generic[_DimT]):
    """Read-only mapping; missing keys return a default dimension (not stored)."""

    def __getitem__(self, key: int | str, /) -> _DimT: ...
    def get(self, key: int | str, default: typing.Any = None) -> typing.Any: ...
    def __contains__(self, key: object, /) -> bool: ...
    def __len__(self) -> int: ...
    def __iter__(self) -> typing.Iterator[typing.Any]: ...
    def keys(self) -> list[typing.Any]: ...
    def values(self) -> list[_DimT]: ...
    def items(self) -> list[tuple[typing.Any, _DimT]]: ...

@typing.final
class DefinedName:
    @property
    def name(self) -> str: ...
    @property
    def value(self) -> str: ...
    @property
    def attr_text(self) -> str: ...
    @property
    def localSheetId(self) -> int | None: ...
    @property
    def hidden(self) -> bool | None: ...
    @property
    def comment(self) -> str | None: ...
    @property
    def is_reserved(self) -> str | None: ...
    @property
    def is_external(self) -> bool: ...
    @property
    def destinations(self) -> typing.Iterator[tuple[str, str]]: ...

@typing.final
class DocumentProperties:
    """Properties the file does not store are ``None``."""

    @property
    def creator(self) -> str | None: ...
    @property
    def title(self) -> str | None: ...
    @property
    def description(self) -> str | None: ...
    @property
    def subject(self) -> str | None: ...
    @property
    def identifier(self) -> str | None: ...
    @property
    def language(self) -> str | None: ...
    @property
    def created(self) -> datetime.datetime | None: ...
    @property
    def modified(self) -> datetime.datetime | None: ...
    @property
    def lastModifiedBy(self) -> str | None: ...
    @property
    def last_modified_by(self) -> str | None: ...
    @property
    def category(self) -> str | None: ...
    @property
    def contentStatus(self) -> str | None: ...
    @property
    def version(self) -> str | None: ...
    @property
    def revision(self) -> str | None: ...
    @property
    def keywords(self) -> str | None: ...
    @property
    def lastPrinted(self) -> datetime.datetime | None: ...

@typing.final
class CustomProperty:
    @property
    def name(self) -> str: ...
    @property
    def value(self) -> typing.Any: ...

@typing.final
class TableStyleInfo:
    @property
    def name(self) -> str | None: ...
    @property
    def showFirstColumn(self) -> bool | None: ...
    @property
    def showLastColumn(self) -> bool | None: ...
    @property
    def showRowStripes(self) -> bool | None: ...
    @property
    def showColumnStripes(self) -> bool | None: ...

@typing.final
class TableColumn:
    @property
    def id(self) -> int | None: ...
    @property
    def name(self) -> str | None: ...
    @property
    def totalsRowFunction(self) -> str | None: ...
    @property
    def totalsRowLabel(self) -> str | None: ...

@typing.final
class AutoFilter:
    @property
    def ref(self) -> str | None: ...

@typing.final
class Table:
    @property
    def id(self) -> int | None: ...
    @property
    def name(self) -> str | None: ...
    @property
    def displayName(self) -> str | None: ...
    @property
    def ref(self) -> str | None: ...
    @property
    def headerRowCount(self) -> int: ...
    @property
    def totalsRowCount(self) -> int | None: ...
    @property
    def totalsRowShown(self) -> bool | None: ...
    @property
    def tableType(self) -> str | None: ...
    @property
    def comment(self) -> str | None: ...
    @property
    def tableStyleInfo(self) -> TableStyleInfo | None: ...
    @property
    def tableColumns(self) -> list[TableColumn]: ...
    @property
    def autoFilter(self) -> AutoFilter | None: ...
    @property
    def column_names(self) -> list[str | None]: ...

@typing.final
class TableList:
    """Tables by name; ``items()`` returns ``(name, ref)`` pairs like openpyxl."""

    def __getitem__(self, key: str, /) -> Table: ...
    def get(self, key: str, default: typing.Any = None) -> typing.Any: ...
    def __contains__(self, key: str, /) -> bool: ...
    def __len__(self) -> int: ...
    def __iter__(self) -> typing.Iterator[str]: ...
    def keys(self) -> list[str]: ...
    def values(self) -> list[Table]: ...
    def items(self) -> list[tuple[str, str | None]]: ...

@typing.final
class WorksheetProperties:
    @property
    def tabColor(self) -> Color | None: ...
    @property
    def codeName(self) -> str | None: ...

@typing.type_check_only
class _CellBase:
    @property
    def row(self) -> int:
        """1-based row."""

    @property
    def column(self) -> int:
        """1-based column."""

    @property
    def col_idx(self) -> int: ...
    @property
    def coordinate(self) -> str: ...
    @property
    def column_letter(self) -> str: ...
    @property
    def parent(self) -> Worksheet: ...
    @property
    def number_format(self) -> str: ...
    @property
    def font(self) -> Font: ...
    @property
    def fill(self) -> PatternFill | GradientFill: ...
    @property
    def border(self) -> Border: ...
    @property
    def alignment(self) -> Alignment: ...
    @property
    def protection(self) -> Protection: ...
    @property
    def style(self) -> str:
        """Name of the cell's named style."""

    @property
    def quotePrefix(self) -> bool: ...
    @property
    def pivotButton(self) -> bool: ...
    @property
    def has_style(self) -> bool: ...
    @property
    def is_date(self) -> bool: ...

@typing.final
class Cell(_CellBase):
    """A cell of a normal-mode worksheet (openpyxl's ``Cell``)."""

    @property
    def value(self) -> _OpenpyxlValue: ...
    @property
    def internal_value(self) -> _OpenpyxlValue: ...
    @property
    def data_type(self) -> str:
        """``n``, ``s``, ``b``, ``e``, ``d`` or ``f`` (formula, unless ``data_only``)."""

    @property
    def style_id(self) -> int:
        """Workbook-local style id (not comparable with openpyxl's)."""

    @property
    def comment(self) -> Comment | None: ...
    @property
    def hyperlink(self) -> Hyperlink | None: ...
    @property
    def encoding(self) -> str: ...
    @property
    def base_date(self) -> datetime.datetime: ...
    @property
    def formula(self) -> str | ArrayFormula | DataTableFormula | None:
        """The formula, read with the value (``formula_and_value=True``)."""

    @property
    def cached_value(self) -> _OpenpyxlValue:
        """The result saved in the file (``formula_and_value=True``)."""

    def offset(self, row: int = 0, column: int = 0) -> Cell | MergedCell: ...

@typing.final
class ReadOnlyCell(_CellBase):
    """A cell of a read-only worksheet (openpyxl's ``ReadOnlyCell``)."""

    @property
    def value(self) -> _OpenpyxlValue: ...
    @property
    def internal_value(self) -> _OpenpyxlValue: ...
    @property
    def data_type(self) -> str: ...
    @property
    def formula(self) -> str | ArrayFormula | DataTableFormula | None:
        """The formula, read with the value (``formula_and_value=True``)."""

    @property
    def cached_value(self) -> _OpenpyxlValue:
        """The result saved in the file (``formula_and_value=True``)."""

    @property
    def comment(self) -> Comment | None:
        """The cell's comment (``read_comments=True``)."""

    @property
    def hyperlink(self) -> Hyperlink | None:
        """The cell's hyperlink (``read_hyperlinks=True``)."""

    @property
    def merged_range(self) -> MergedCellRange | None:
        """The merged range containing the cell (``read_merged_cells=True``)."""

@typing.final
class MergedCell(_CellBase):
    """A non-anchor cell of a merged range: value ``None``, outer-edge borders."""

    @property
    def value(self) -> None: ...
    @property
    def internal_value(self) -> None: ...
    @property
    def data_type(self) -> str: ...
    @property
    def comment(self) -> None: ...
    @property
    def hyperlink(self) -> None: ...
    @property
    def formula(self) -> None:
        """Always ``None`` (``formula_and_value=True``)."""

    @property
    def cached_value(self) -> None:
        """Always ``None`` (``formula_and_value=True``)."""

@typing.final
class EmptyCell:
    """Padding cell of read-only rows (openpyxl's ``EmptyCell``; no position)."""

    @property
    def value(self) -> None: ...
    @property
    def is_date(self) -> bool: ...
    @property
    def font(self) -> None: ...
    @property
    def border(self) -> None: ...
    @property
    def fill(self) -> None: ...
    @property
    def number_format(self) -> None: ...
    @property
    def alignment(self) -> None: ...
    @property
    def data_type(self) -> str: ...

@typing.final
class RowIterator:
    """Rows as tuples of cells (or of values with ``values_only=True``). Every
    iterator reads the sheet independently of other iterators."""

    def __iter__(self) -> RowIterator: ...
    def __next__(self) -> tuple[typing.Any, ...]: ...

@typing.final
class Worksheet:
    """A worksheet of a ``compatibility="openpyxl"`` workbook: openpyxl's
    ``Worksheet``, or ``ReadOnlyWorksheet`` semantics with ``read_only=True``.
    Rows, columns and coordinates are 1-based."""

    @property
    def title(self) -> str: ...
    @property
    def sheet_state(self) -> str:
        """``visible``, ``hidden`` or ``veryHidden``."""

    @property
    def parent(self) -> Workbook: ...
    @property
    def min_row(self) -> int: ...
    @property
    def min_column(self) -> int: ...
    @property
    def max_row(self) -> int | None:
        """Normal mode: last row of the stored cells, merged ranges, hyperlinks and
        comments (like openpyxl). Read-only mode: from the declared ``<dimension>``
        (``None`` when the file declares none)."""

    @property
    def max_column(self) -> int | None: ...
    @property
    def dimensions(self) -> str: ...
    def calculate_dimension(self, force: bool = False) -> str: ...
    def reset_dimensions(self) -> None:
        """Read-only mode only: forget the declared dimension."""

    def iter_rows(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        values_only: bool = False,
    ) -> RowIterator: ...
    @property
    def rows(self) -> RowIterator: ...
    @property
    def values(self) -> RowIterator: ...
    def __iter__(self) -> RowIterator: ...
    def iter_cols(
        self,
        min_col: int | None = None,
        max_col: int | None = None,
        min_row: int | None = None,
        max_row: int | None = None,
        values_only: bool = False,
    ) -> typing.Iterator[tuple[typing.Any, ...]]:
        """Normal mode only (like openpyxl). Reads the requested block in one pass."""

    @property
    def columns(self) -> typing.Iterator[tuple[typing.Any, ...]]: ...
    def cell(self, row: int, column: int, value: None = None) -> typing.Any: ...
    def __getitem__(self, key: str | int | slice, /) -> typing.Any: ...
    @property
    def merged_cells(self) -> MultiCellRange: ...
    @property
    def freeze_panes(self) -> str | None: ...
    @property
    def print_area(self) -> str:
        """``"'Sheet'!$A$1:$C$10"`` or ``""`` (also for non-static print areas)."""

    @property
    def print_title_rows(self) -> str | None: ...
    @property
    def print_title_cols(self) -> str | None: ...
    @property
    def print_titles(self) -> str: ...
    @property
    def defined_names(self) -> dict[str, DefinedName]: ...
    @property
    def row_dimensions(self) -> DimensionHolder[RowDimension]: ...
    @property
    def column_dimensions(self) -> DimensionHolder[ColumnDimension]: ...
    @property
    def sheet_properties(self) -> WorksheetProperties: ...
    @property
    def auto_filter(self) -> AutoFilter: ...
    @property
    def tables(self) -> TableList: ...

@typing.final
class Chartsheet:
    @property
    def title(self) -> str: ...
    @property
    def sheet_state(self) -> str: ...
    @property
    def parent(self) -> Workbook: ...

@typing.final
class Workbook:
    """A read-only, openpyxl-shaped workbook (``load_workbook(...,
    compatibility="openpyxl")``)."""

    @property
    def sheetnames(self) -> list[str]: ...
    @property
    def worksheets(self) -> list[Worksheet]: ...
    @property
    def chartsheets(self) -> list[Chartsheet]: ...
    @property
    def active(self) -> Worksheet | Chartsheet | None: ...
    def __getitem__(self, key: str, /) -> Worksheet: ...
    def __contains__(self, key: str, /) -> bool: ...
    def __iter__(self) -> typing.Iterator[Worksheet]: ...
    def index(self, worksheet: Worksheet | Chartsheet) -> int: ...
    @property
    def read_only(self) -> bool: ...
    @property
    def data_only(self) -> bool: ...
    @property
    def epoch(self) -> datetime.datetime: ...
    @property
    def excel_base_date(self) -> datetime.datetime: ...
    @property
    def template(self) -> bool: ...
    @property
    def code_name(self) -> str | None: ...
    @property
    def vba_archive(self) -> None: ...
    @property
    def loaded_theme(self) -> bytes | None: ...
    @property
    def properties(self) -> DocumentProperties: ...
    @property
    def custom_doc_props(self) -> list[CustomProperty]: ...
    @property
    def defined_names(self) -> dict[str, DefinedName]: ...
    @property
    def named_styles(self) -> list[str]: ...
    @property
    def style_names(self) -> list[str]: ...
    def close(self) -> None:
        """Release the file. Reading more cells afterwards raises ``WorkbookClosed``."""

    def __enter__(self) -> Workbook: ...
    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: types.TracebackType | None,
    ) -> None: ...

# ---- openpyxl utilities ------------------------------------------------------------

WINDOWS_EPOCH: datetime.datetime
MAC_EPOCH: datetime.datetime

def get_column_letter(col_idx: int) -> str:
    """Excel column letters of a 1-based column index (``28`` -> ``"AB"``)."""

def column_index_from_string(col: str) -> int:
    """1-based column index of one to three letters (``"AB"`` -> ``28``)."""

def coordinate_from_string(coord_string: str) -> tuple[str, int]:
    """``"B12"`` -> ``("B", 12)``."""

def coordinate_to_tuple(coordinate: str) -> tuple[int, int]:
    """``"C7"`` -> ``(7, 3)`` (1-based row, column)."""

def coordinate_from_row_col(row: int, column: int) -> str:
    """A1 coordinate of a 1-based row and column (``(7, 3)`` -> ``"C7"``)."""

def coordinate_from_tuple(row_column: tuple[int, int]) -> str:
    """Inverse of ``coordinate_to_tuple``: ``(7, 3)`` -> ``"C7"``."""

def absolute_coordinate(coord_string: str) -> str: ...
def range_boundaries(
    range_string: str,
) -> tuple[int | None, int | None, int | None, int | None]:
    """``"A1:C3"`` -> ``(min_col, min_row, max_col, max_row)``, 1-based."""

def range_to_tuple(
    range_string: str,
) -> tuple[str, tuple[int | None, int | None, int | None, int | None]]: ...
def quote_sheetname(sheetname: str) -> str: ...
def rows_from_range(range_string: str) -> typing.Iterator[tuple[str, ...]]: ...
def cols_from_range(range_string: str) -> typing.Iterator[tuple[str, ...]]: ...
def get_column_interval(start: str | int, end: str | int) -> list[str]: ...
def split_print_areas(value: str) -> list[str]:
    """Static A1 ranges of a print area (``"'S'!$A$1:$C$10,'S'!$E$1:$F$2"`` ->
    ``["A1:C10", "E1:F2"]``); ``ValueError`` for formulas and other non-static values.
    """

def is_date_format(fmt: str | None) -> bool: ...
def is_timedelta_format(fmt: str | None) -> bool: ...
def builtin_formats() -> dict[int, str]: ...
def from_excel(
    value: float | None,
    epoch: datetime.datetime | None = None,
    timedelta: bool = False,
) -> datetime.datetime | datetime.time | datetime.timedelta | None:
    """Excel serial -> ``datetime`` (``time`` below 1; ``timedelta`` with ``timedelta=True``)."""

def serial_to_datetime(serial: float, epoch_1904: bool = False) -> datetime.datetime:
    """Excel serial -> ``datetime`` (always a ``datetime``)."""

def to_excel(
    dt: datetime.date | datetime.datetime | datetime.time | datetime.timedelta,
    epoch: datetime.datetime | None = None,
) -> float: ...
def from_ISO8601(
    formatted_string: str,
) -> datetime.datetime | datetime.date | datetime.time | datetime.timedelta | None: ...

__all__ = [
    "Alignment",
    "ArrayFormula",
    "AutoFilter",
    "Border",
    "CalamineCell",
    "CalamineError",
    "CalamineSheet",
    "CalamineSheetStream",
    "CalamineTable",
    "CalamineWorkbook",
    "Cell",
    "CellCoordinatesException",
    "Chartsheet",
    "Color",
    "ColumnDimension",
    "Comment",
    "CompatibilityNotSupported",
    "CustomProperty",
    "DataTableFormula",
    "DefinedName",
    "DimensionHolder",
    "DocumentProperties",
    "EmptyCell",
    "Font",
    "GradientFill",
    "Hyperlink",
    "MAC_EPOCH",
    "MergedCell",
    "MergedCellRange",
    "MultiCellRange",
    "PasswordError",
    "PatternFill",
    "Protection",
    "ReadOnlyCell",
    "RowDimension",
    "RowIterator",
    "SheetMetadata",
    "SheetTypeEnum",
    "SheetVisibleEnum",
    "Side",
    "Stop",
    "StreamInvalidated",
    "StreamingNotSupported",
    "Table",
    "TableColumn",
    "TableList",
    "TableNotFound",
    "TableStyleInfo",
    "TablesNotLoaded",
    "TablesNotSupported",
    "WINDOWS_EPOCH",
    "Workbook",
    "WorkbookClosed",
    "Worksheet",
    "WorksheetNotFound",
    "WorksheetProperties",
    "XmlError",
    "ZipError",
    "absolute_coordinate",
    "builtin_formats",
    "cols_from_range",
    "column_index_from_string",
    "coordinate_from_row_col",
    "coordinate_from_string",
    "coordinate_from_tuple",
    "coordinate_to_tuple",
    "from_ISO8601",
    "from_excel",
    "get_column_interval",
    "get_column_letter",
    "is_date_format",
    "is_timedelta_format",
    "load_workbook",
    "quote_sheetname",
    "range_boundaries",
    "range_to_tuple",
    "rows_from_range",
    "serial_to_datetime",
    "split_print_areas",
    "to_excel",
]
