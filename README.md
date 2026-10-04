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

## Benchmarks

Measured on Windows 10, Python 3.13 (peak memory = increase of the process's peak working set).
Each library read every row of every sheet.

**`large_sample.xlsx`**: 94 MB file, 25 sheets, 470,300 rows, 11.2 M cells, 5.6 M formulas

| | Peak memory | Time |
|---|---:|---:|
| **python-calamine-plus** (plain) | **35 MB** | 9.9 s |
| **python-calamine-plus** (`rich=True`) | **41 MB** | 21.1 s |
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

| Format | Streaming (`stream_sheet_*`) | Whole-sheet API (`get_sheet_*`) |
|---|:---:|:---:|
| `.xlsx`, `.xlsm`, `.xltx`, `.xltm` | ✅ (plain and `rich=True`) | ✅ |
| `.xls`, `.xlsb`, `.ods` | ❌ raises `StreamingNotSupported` | ✅ (as upstream) |

## Differences from openpyxl

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
