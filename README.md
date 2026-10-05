# python-calamine-plus

[![PyPI - Version](https://img.shields.io/pypi/v/python-calamine-plus)](https://pypi.org/project/python-calamine-plus/)
![Python versions](https://img.shields.io/pypi/pyversions/python-calamine-plus)

> **This is a fork.** It is based on [python-calamine](https://github.com/dimastbk/python-calamine)
> by Dmitriy ([@dimastbk](https://github.com/dimastbk)), the Python binding for
> [calamine](https://github.com/tafia/calamine), the Rust spreadsheet reader by Johann Tuffe
> ([@tafia](https://github.com/tafia)). The fork is based on python-calamine 0.8.2 and ships a
> modified copy of calamine 0.36.1 in [`vendor/calamine`](vendor/calamine). It is not affiliated
> with or endorsed by the upstream projects. All credit for the original work goes to their authors.

## What the fork adds

python-calamine is fast, but it loads a whole sheet into memory before you can read a single
row. **python-calamine-plus adds true row-by-row streaming for `.xlsx` / `.xlsm` files.** Rows
are decoded straight from the compressed file as you iterate, so memory stays flat no matter how
big the sheet is.

On top of that, an optional **`rich=True`** mode returns everything openpyxl can *read* (cell
styles, formulas, comments, merged cells, hyperlinks, rich text, conditional formatting, data
validation, tables, column/row dimensions, sheet views, page setup, workbook properties, ...).
It does this while still streaming, and roughly 6–10× faster than openpyxl's read-only mode.

| | What you get |
|---|---|
| **Streaming** | `wb.stream_sheet_by_name(name)` yields one row at a time (a list of values) in constant memory. You can read part of a sheet with `min_row` / `max_row` / `min_col` / `max_col` or `ref="B2:D10"`, and reading stops early after `max_row`. |
| **Rich streaming** | `rich=True` yields `CalamineCell` objects with value, style, formula, comment, hyperlink, merged range, rich text and data type. The stream also exposes the sheet's metadata. |
| **Workbook metadata** | Defined names (with sheet scope), document and custom properties, theme, named styles, active sheet, protection, calculation settings, 1904 date system, template / macro flags. |
| **Value fixes** | Excel errors come back as `"#DIV/0!"`, `"#N/A"`, ... (upstream returns `""`), and whole numbers as `int` (upstream returns `float`), matching openpyxl. |
| **openpyxl compatibility** | `load_workbook(path, compatibility="openpyxl")` returns an openpyxl-shaped, read-only workbook (`wb["Sheet"]["A1"].font.b`, `ws.iter_rows(min_row=2, values_only=True)`, `data_only=`, `read_only=`, merged cells, ...), implemented in Rust and still streaming. openpyxl's utilities (`get_column_letter`, `range_boundaries`, `from_excel`, ...) are included. See [openpyxl compatibility](#openpyxl-compatibility). |

Everything from python-calamine still works unchanged (`get_sheet_by_name()`, `to_python()`,
`iter_rows()`, tables, `.xls` / `.xlsb` / `.ods` support, ...).

## Installation

```shell
pip install python-calamine-plus
```

Pre-built wheels are published for **Windows** (x64, x86, ARM64), **macOS** (Intel and Apple
Silicon), and **Linux** (x86_64, aarch64, i686, armv7, ppc64le, s390x; glibc and musl/Alpine), for
CPython 3.10 – 3.14 (including free-threaded 3.14t) and PyPy. **No Rust toolchain is needed.**

The import name is **`python_calamine_plus`**, so the package can be installed next to the
original `python-calamine` (which `pandas[excel]` installs) without conflicts:

```python
from python_calamine_plus import CalamineWorkbook
```

## Usage

### Stream rows in constant memory

```python
from python_calamine_plus import CalamineWorkbook

wb = CalamineWorkbook.from_path("big.xlsx")          # or .from_filelike(...) / .from_object(...)
for row in wb.stream_sheet_by_name("Ledger"):        # or stream_sheet_by_index(0)
    ...                                              # row is a list: [str | int | float | bool | date | datetime | time | timedelta]
```

Rows start at row 1 / column A (index 0). Missing cells come back as `""`, the same as
upstream's `iter_rows()`.

### Read only part of a sheet

```python
wb.stream_sheet_by_name("Ledger", min_row=1000, max_row=1999, min_col=0, max_col=4)  # 0-based, inclusive
wb.stream_sheet_by_name("Ledger", ref="B2:D10")      # A1 notation, like openpyxl's ws["B2:D10"]
```

### Rich mode: styles, formulas, comments, ...

```python
stream = wb.stream_sheet_by_name("Ledger", rich=True)
for row in stream:
    for cell in row:                                 # CalamineCell
        cell.value, cell.coordinate                  # 42, "C7"
        cell.formula                                 # "=SUM(C2:C6)" (shared formulas expanded per cell)
        cell.style["font"]["bold"]                   # also: fill, border, alignment, number_format, protection, named_style
        cell.comment                                 # {"author": ..., "text": ...} or None
        cell.hyperlink                               # {"target": ..., "location": ..., "tooltip": ...} or None
        cell.merged_range                            # ((first_row, first_col), (last_row, last_col)) or None
        cell.rich_text                               # [{"text": ..., "font": {...}}, ...] or None
        cell.data_type, cell.is_date                 # "n" / "s" / "b" / "e" / "d", like openpyxl
    stream.row_info                                  # height / hidden / outline level of the row just yielded

# Sheet-level metadata (available as soon as the stream starts):
stream.merged_cell_ranges, stream.hyperlinks, stream.comments
stream.conditional_formats, stream.data_validations, stream.tables
stream.column_dimensions, stream.row_dimensions, stream.row_breaks, stream.column_breaks
stream.sheet_settings     # freeze panes, auto filter (+ sort), protection, page setup, margins,
                          # print area / titles, header / footer, views, tab color, ...
```

### Workbook metadata

```python
wb.defined_names          # [{"name", "value", "sheet", "hidden", "comment"}, ...]
wb.properties             # {"creator": ..., "created": datetime(...), "title": ..., ...}
wb.custom_properties      # {"Project": "X", "Count": 3, ...}
wb.named_styles, wb.theme_colors, wb.theme_xml
wb.active_sheet_index, wb.workbook_protection, wb.calculation
wb.epoch_1904, wb.is_template, wb.has_macros
```

Type hints for the whole API are included (`python_calamine_plus/_python_calamine.pyi`).

### Original (whole-sheet) API

```python
wb = CalamineWorkbook.from_path("file.xlsx")
wb.sheet_names                                       # ["Sheet1", "Sheet2"]
wb.get_sheet_by_name("Sheet1").to_python()           # list of rows, whole sheet in memory
wb.get_sheet_by_name("Sheet1").to_python(skip_empty_area=False)
```

## openpyxl compatibility

Pass **`compatibility="openpyxl"`** to `load_workbook` to get openpyxl's reading API and
values. Code written for openpyxl works with only the import changed:

```python
from python_calamine_plus import load_workbook

wb = load_workbook("book.xlsx", compatibility="openpyxl")   # also: read_only=True, data_only=True
ws = wb["Ledger"]                                           # wb.sheetnames, wb.active, wb.worksheets
ws["A1"].value, ws["A1"].font.b, ws["A1"].fill.fgColor.rgb, ws["B2"].number_format
ws.cell(row=2, column=3).value
for row in ws.iter_rows(min_row=2, max_col=4, values_only=True):   # 1-based, like openpyxl
    ...
ws.max_row, ws.dimensions, ws.merged_cells, ws.freeze_panes, ws.print_area
```

or, as a drop-in module with openpyxl's `load_workbook` signature:

```python
from python_calamine_plus import openpyxl          # instead of `import openpyxl`
wb = openpyxl.load_workbook("book.xlsx", read_only=True, data_only=True)
from python_calamine_plus.openpyxl.utils import get_column_letter
```

Without the flag nothing changes: `load_workbook(path)` still returns a `CalamineWorkbook`
with the native API and native values described above.

### What the profile provides

The behavior is checked against openpyxl 3.1 (the test oracle) on hand-written edge cases
and on 472 real-world workbooks.

| Area | Behavior |
|---|---|
| **Loading** | `load_workbook(path_or_filelike, compatibility="openpyxl", read_only=False, data_only=False)`. `keep_vba` / `keep_links` / `rich_text` are accepted only at openpyxl's defaults (they affect saving, or need `CellRichText`); openpyxl's options without the flag raise `TypeError`. `.xls` / `.xlsb` / `.ods` raise `CompatibilityNotSupported` (openpyxl reads xlsx only). |
| **Workbook** | `sheetnames`, `worksheets`, `chartsheets`, `active`, `wb[name]` (`KeyError` like openpyxl), `in`, iteration, `index()`, `read_only`, `data_only`, `epoch` (1900 / 1904), `template`, `code_name`, `properties` (core document properties), `custom_doc_props`, `defined_names`, `named_styles`, `loaded_theme`, `close()`, context manager. |
| **Worksheet** | `title`, `sheet_state`, `parent`, `ws["A1"]` / `ws["A1:C3"]` / `ws["A"]` / `ws["A:C"]` / `ws[3]` / `ws["2:4"]` / `ws[2:4]`, `cell()`, `iter_rows()`, `rows`, `values`, `iter_cols()` / `columns` (normal mode only, like openpyxl), `min_row` / `max_row` / `min_column` / `max_column`, `dimensions`, `calculate_dimension(force=)`, `merged_cells`, `row_dimensions`, `column_dimensions`, `freeze_panes`, `print_area`, `print_title_rows` / `print_title_cols`, `print_titles`, `defined_names` (sheet scope), `tables`, `auto_filter.ref`, `sheet_properties.tabColor` / `.codeName`. |
| **Cells** | `Cell` / `ReadOnlyCell` / `MergedCell` / `EmptyCell` with `value`, `data_type`, `row`, `column` (1-based), `coordinate`, `column_letter`, `number_format`, `is_date`, `font`, `fill`, `border`, `alignment`, `protection`, `style`, `has_style`, `comment`, `hyperlink`, `offset()`. Objects are read-only. |
| **Values** | Empty cells are `None`; an explicitly stored empty string is `""`. Whole numbers written without `.`/`e` are `int`, others `float`. Numbers with a date format become `datetime` (`time` below 1, `timedelta` for `[h]:mm`-style formats) using the workbook's epoch and openpyxl's date-format rules; out-of-range serials become the error `"#VALUE!"`. ISO dates (`t="d"`) become `date` / `time` / `datetime` / `timedelta`. Errors are strings (`"#DIV/0!"`, `data_type "e"`), booleans `bool`. |
| **Formulas** | Default (formula view): formula cells have `data_type "f"` and the formula text as value (`"=SUM(A1:A3)"`; shared formulas expanded per cell; `ArrayFormula(ref, text)` / `DataTableFormula` objects for array / data table formulas). `data_only=True`: the result saved in the file, typed as above, or `None` if the file has no saved result. **Nothing is recalculated.** |
| **Normal mode** (`read_only=False`) | Like openpyxl's `Worksheet`: bounds cover every stored cell (including styled empty cells) plus merged ranges, hyperlinks and comments; non-anchor cells of merged ranges are `MergedCell`s (value `None`, the range's outer-edge borders, the anchor's protection); hyperlinks and comments are bound to cells (a link on an empty cell makes the link its value, as in openpyxl). Iteration is rectangular from row 1 / column 1. |
| **Read-only mode** (`read_only=True`) | Like openpyxl's `ReadOnlyWorksheet`: bounds come from the sheet's declared `<dimension>` (`None` if absent; `calculate_dimension(force=True)` scans), rows are padded with `EmptyCell` (or `None` with `values_only`), comments / hyperlinks / merges are not applied to cells. No sheet metadata is read unless you ask for it. |
| **Styles** | `Font` (`name`, `sz`/`size`, `b`/`bold`, `i`/`italic`, `u`/`underline`, `strike`, `vertAlign`, `color`, `family`, `charset`, `scheme`, ...), `PatternFill` (`patternType`/`fill_type`, `fgColor`/`start_color`, `bgColor`/`end_color`), `GradientFill`, `Border` / `Side`, `Alignment`, `Protection`, `Color` (`type` = `rgb` / `theme` / `indexed` / `auto`, `tint`). Attributes the file does not set are `None` or openpyxl's defaults (e.g. `fgColor.rgb == "00000000"`). Theme and indexed colors are **not** resolved to RGB. Number formats use openpyxl's built-in format table. |

### Extensions: formulas with values, and read-only metadata

Opt-in keywords that go beyond openpyxl (the defaults are openpyxl's behavior):

```python
wb = load_workbook("book.xlsx", compatibility="openpyxl", read_only=True,
                   formula_and_value=True,                  # formula and saved result in one pass
                   read_comments=True, read_hyperlinks=True, read_merged_cells=True)
for row in wb["Ledger"].iter_rows(min_row=2):
    for cell in row:
        cell.value                        # follows data_only, as in openpyxl
        cell.formula, cell.cached_value   # "=SUM(B2:B9)", 1234.5 (None / None where absent)
        cell.comment, cell.hyperlink, cell.merged_range
```

* **`formula_and_value=True`** (normal and read-only mode): every cell also has `formula`
  (what the formula view would return: text, `ArrayFormula` or `DataTableFormula`, `None`
  without a formula) and `cached_value` (what `data_only=True` would return). Both come from the
  same streaming pass; `value` / `data_type` still follow `data_only`.
* **`read_comments` / `read_hyperlinks` / `read_merged_cells`** (read-only mode): read-only cells
  get `comment`, `hyperlink` and `merged_range` the way the native `rich=True` stream attaches
  them: every cell of a link's range and every cell of a merged range, and positions that have a
  comment or link but no stored cell get a `ReadOnlyCell` (value `None`) instead of an
  `EmptyCell`. Like the native rich stream, this reads the sheet's metadata with one extra pass
  when iteration starts; values-only iteration does not read it. Normal mode always binds them
  the openpyxl way, so `False` there raises `ValueError`.
* Without a flag, these attributes raise `AttributeError` naming the flag to pass.

### How it maps to native streaming

Both APIs stream rows; neither loads a whole sheet.

| Native: `wb.stream_sheet_by_name(name, rich=True)` | `compatibility="openpyxl"`: `ws.iter_rows()` |
|---|---|
| `CalamineCell`, `row` / `column` 0-based | `Cell` / `ReadOnlyCell`, 1-based |
| `value` (saved result, `""` when empty) | `value` (openpyxl types, `None` when empty; formula text unless `data_only=True`) |
| `formula` + `value` together | `formula` + `cached_value` with `formula_and_value=True` |
| `formula_type`, `formula_range` | `ArrayFormula` / `DataTableFormula` objects |
| `style` (dicts: `style["font"]["bold"]`) | `font`, `fill`, `border`, `alignment`, `protection`, `number_format` (objects: `font.b`) |
| `comment`, `hyperlink`, `merged_range` | normal mode always; read-only mode with `read_comments` / `read_hyperlinks` / `read_merged_cells` |
| `rich_text`, `stream.sheet_settings`, conditional formats, data validation | not provided |

### openpyxl utilities (implemented in Rust)

Importable like openpyxl's, from `python_calamine_plus.utils` (also `utils.cell`,
`utils.datetime`), `python_calamine_plus.styles.numbers`, or the `python_calamine_plus.openpyxl`
drop-in package:

| Function | |
|---|---|
| `get_column_letter(28)` | `"AB"` (1-based column index to letters) |
| `column_index_from_string("AB")` | `28` (one to three letters) |
| `coordinate_to_tuple("C7")` | `(7, 3)` (1-based row, column) |
| `coordinate_from_row_col(7, 3)`, `coordinate_from_tuple((7, 3))` | `"C7"` (not in openpyxl) |
| `coordinate_from_string("B12")` | `("B", 12)` |
| `range_boundaries("A1:C3")` | `(1, 1, 3, 3)`: `(min_col, min_row, max_col, max_row)` |
| `range_to_tuple`, `absolute_coordinate`, `quote_sheetname`, `rows_from_range`, `cols_from_range`, `get_column_interval` | as in openpyxl |
| `split_print_areas("'S'!$A$1:$C$10,'S'!$E$1:$F$2")` | `["A1:C10", "E1:F2"]`; `ValueError` for formulas / non-static values (not in openpyxl) |
| `from_excel(45351.5)`, `to_excel(dt)`, `from_ISO8601(text)` | as in openpyxl (`epoch=`, `timedelta=`) |
| `serial_to_datetime(45351, epoch_1904=False)` | always a `datetime` (not in openpyxl) |
| `is_date_format(fmt)`, `is_timedelta_format(fmt)`, `BUILTIN_FORMATS` | as in `openpyxl.styles.numbers` |

### Streaming, iterators and lifecycle

* Every iterator (`iter_rows`, `rows`, `values`, `for row in ws`) has its **own** reader:
  iterators over the same or different sheets never invalidate each other, and random access
  (`ws["A1"]`, `ws.cell()`) may be interleaved with iteration. (The native API keeps its
  one-active-stream rule.)
* Nothing is materialized: rows are decoded as they are iterated. Random access goes through a
  bounded row cache (at most 4,096 rows / 1M cells); reading backwards restarts the reader.
  `iter_cols()` reads the requested block in one pass and holds that block.
* Normal mode reads a sheet's metadata (merged ranges, hyperlinks, comments, bounds) with one
  extra pass the first time it is needed. Read-only mode reads it only when you access it
  (`merged_cells`, `row_dimensions`, ...); values-only read-only iteration never reads it.
* `wb.close()` (or leaving `with`) releases the file. Cells and values already read stay
  usable; reading more (`next(iterator)`, `ws["A1"]`, `iter_rows()`) raises `WorkbookClosed`.
  An iterator created before `close()` releases its own file handle on its next `next()` call
  or when it is garbage-collected.

### Differences from openpyxl in this profile

These are deliberate; everything else listed above matches openpyxl's observable behavior.

* **Read-only objects.** Nothing can be assigned or saved (`AttributeError` / `TypeError`).
* **Accessing a cell does not create it**: `ws["Z99"]` returns an empty cell, but `max_row` /
  `max_column` / `dimensions` do not grow (openpyxl's normal mode adds the cell). Cell objects
  are created per access, so `ws["A1"] is ws["A1"]` is `False`.
* **Normal mode is lazy**: openpyxl reads everything into memory at load; here data is read on
  demand, so cells cannot be read after `wb.close()`.
* **Unset color attributes are `None`** (openpyxl returns a descriptor object, e.g. the `rgb` of
  a theme color prints as `"Values must be of type <class 'str'>"`). Theme colors are not resolved.
* **Document properties the file does not store are `None`** (openpyxl fills in
  `creator="openpyxl"` and the current time for `created` / `modified`).
* `ws.merged_cells.ranges` is a list in file order (openpyxl: a set). `cell.style_id` is a
  workbook-local id, not openpyxl's.
* `cell.style` (the named style) is the one the file links to the cell's format. openpyxl looks
  it up by position in its sorted list instead, which gives a different name (e.g. `Comma 3`
  instead of `Comma`) or raises `IndexError` when the file's style ids are not contiguous.
* `has_style` is `True` for any non-default cell format id; `Hyperlink.id` (the relationship id)
  is always `None`. Read-only worksheets also expose metadata openpyxl's `ReadOnlyWorksheet`
  lacks (`merged_cells`, `print_area`, ...).
* **Not provided** (accessing them raises `AttributeError` naming the alternative): conditional
  formatting, data validation, page setup / margins / print options, header & footer, sheet
  views and protection, page breaks, scenarios, rich text values (`rich_text=True`), charts,
  images and pivot tables. The native `stream_sheet_by_name(name, rich=True)` returns most of
  these (`sheet_settings`, `conditional_formats`, `data_validations`, `rich_text`, ...).

## Benchmarks

Measured on Windows 10, Python 3.13 (peak memory = increase of the process's peak working set).
Each library read every row of every sheet.

**`large_sample.xlsx`**: 94 MB file, 25 sheets, 470,300 rows, 11.2 M cells, 5.6 M formulas

| | Peak memory | Time |
|---|---:|---:|
| **python-calamine-plus** (plain) | **35 MB** | 9.9 s |
| **python-calamine-plus** (`rich=True`) | **41 MB** | 21.1 s |
| **`compatibility="openpyxl"`, `read_only=True`, `values_only=True`** | **37 MB** | 14.9 s |
| **`compatibility="openpyxl"`, `read_only=True`, reading each cell's font / fill / border / format** | **37 MB** | 24.0 s |
| **`compatibility="openpyxl"`, normal mode, cells + styles + comments + hyperlinks** | **38 MB** | 34.7 s |
| python-calamine (whole sheet) | 256 MB | 12.1 s |
| fastexcel | 256 MB | 8.4 s |
| openpyxl `read_only=True` | 101 MB | 135.9 s |

**`100mb.xlsx`**: 101 MB file, 1.1 M rows, text-heavy (153 MB of shared strings)

| | Peak memory | Time |
|---|---:|---:|
| **python-calamine-plus** (plain) | **187 MB** | 8.3 s |
| **python-calamine-plus** (`rich=True`) | **188 MB** | 11.6 s |
| python-calamine (whole sheet) | 1,263 MB | 11.4 s |
| fastexcel | 2,049 MB | 10.9 s |
| openpyxl `read_only=True` | 320 MB | 110.6 s |

Most of the remaining memory is the workbook's shared-strings table, which every reader
(including openpyxl) has to load in full, because cells refer to strings by index.

**Correctness:** on `large_sample.xlsx`, all 11.2 M values, 5.6 M formulas and 9.9 M cell styles
match openpyxl, and so does the sheet and workbook metadata. The test suite also checks openpyxl
parity, and the same output was verified byte-for-byte on Windows and Linux.

## Supported formats

| Format | Streaming (`stream_sheet_*`) | Whole-sheet API (`get_sheet_*`) | `compatibility="openpyxl"` |
|---|:---:|:---:|:---:|
| `.xlsx`, `.xlsm`, `.xltx`, `.xltm` | ✅ (plain and `rich=True`) | ✅ | ✅ |
| `.xls`, `.xlsb`, `.ods` | ❌ raises `StreamingNotSupported` | ✅ (as upstream) | ❌ raises `CompatibilityNotSupported` |

## Differences from openpyxl (native API)

These apply to the native API; `compatibility="openpyxl"` follows openpyxl instead.

* **Dates** without a time come back as `datetime.date` (openpyxl returns `datetime` at midnight).
* **Empty cells** are `""` (openpyxl: `None`).
* **Values and formulas** are both available in one pass: `cell.value` is the result saved in
  the file and `cell.formula` is the formula text. Formulas are not recalculated, so files never
  saved by Excel have no saved results.
* Attribute dicts (page setup, protection, views, ...) use snake_case keys with typed values,
  e.g. `{"show_grid_lines": False}`.
* openpyxl's read-only mode has no comments, merged cells or hyperlinks; `rich=True` has them all.

## Limitations

* **Reading only.** No writing or editing. Charts, images and pivot tables are not read.
* **One active stream per workbook.** Starting another stream, loading a sheet or closing the
  workbook invalidates the current stream (`StreamInvalidated` / `WorkbookClosed` on the next
  read). Open the workbook twice to read two sheets side by side.
* Cells must be stored in row order, which Excel and all common writers do.
* Theme colors are returned as stored (`{"theme": 4, "tint": 0.4}`), not resolved to RGB
  (openpyxl doesn't resolve them either). Extended (`x14`) conditional-format details are skipped;
  the base rule is returned.
* `rich=True` reads the sheet's metadata (merged cells, hyperlinks, ...) with one extra pass over
  the sheet when the stream starts, because Excel stores it after the cell data.

## Development

Building from source needs Rust ([rustup.rs](https://rustup.rs/)).

```shell
git clone https://github.com/AmirShokry/python-calamine-plus.git
cd python-calamine-plus
python -m venv .venv && . .venv/bin/activate     # Windows: .venv\Scripts\activate
pip install --group dev -e .                       # pip 25.1+
pre-commit run --all-files
pytest
```

## License and credits

MIT, the same as the upstream projects. See [LICENSE](LICENSE) and
[vendor/calamine/LICENSE-MIT.md](vendor/calamine/LICENSE-MIT.md).

* [python-calamine](https://github.com/dimastbk/python-calamine): Dmitriy and contributors
* [calamine](https://github.com/tafia/calamine): Johann Tuffe and contributors
* Built with [PyO3](https://github.com/PyO3/pyo3) and [maturin](https://github.com/PyO3/maturin)
