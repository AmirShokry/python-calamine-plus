# Changelog

## 0.2.0 (2026-10-05)

### Added

- **openpyxl compatibility profile**: `load_workbook(path_or_filelike, compatibility="openpyxl")`
  returns a read-only, openpyxl-shaped `Workbook` (`wb["Sheet"]`, `ws["A1"]`, `ws.cell()`,
  `ws.iter_rows(..., values_only=True)`, `cell.font` / `fill` / `border` / `alignment` /
  `protection` / `number_format`, merged cells, comments, hyperlinks, row / column dimensions,
  print settings, defined names, document properties, tables), with openpyxl's `read_only` and
  `data_only` semantics and openpyxl's values (`None` for empty cells, `datetime` for dates,
  formula text unless `data_only=True`). It is implemented in Rust, streams rows instead of
  loading sheets, and gives every iterator its own reader. Checked against openpyxl 3.1 as the
  test oracle. See "openpyxl compatibility" in the README for the contract and the documented
  differences.
- Opt-in extensions of the profile: `formula_and_value=True` (cells also get `formula` and
  `cached_value`, read in the same pass) and, for `read_only=True`, `read_comments` /
  `read_hyperlinks` / `read_merged_cells` (cells get `comment` / `hyperlink` / `merged_range`
  like the native `rich=True` stream).
- `python_calamine_plus.openpyxl`: a drop-in module with openpyxl's `load_workbook` signature and
  `utils` / `styles` subpackages.
- openpyxl's utilities, implemented in Rust, under openpyxl's import paths
  (`python_calamine_plus.utils`, `.utils.cell`, `.utils.datetime`, `.styles.numbers`):
  `get_column_letter`, `column_index_from_string`, `coordinate_from_string`,
  `coordinate_to_tuple`, `absolute_coordinate`, `range_boundaries`, `range_to_tuple`,
  `quote_sheetname`, `rows_from_range`, `cols_from_range`, `get_column_interval`, `from_excel`,
  `to_excel`, `from_ISO8601`, `is_date_format`, `is_timedelta_format`, `BUILTIN_FORMATS`; plus
  `coordinate_from_row_col` / `coordinate_from_tuple`, `serial_to_datetime` and
  `split_print_areas`.
- `CompatibilityNotSupported` exception (`.xls` / `.xlsb` / `.ods` files and `rich_text=True`
  with `compatibility="openpyxl"`).

### Fixed

- Hyperlink `target`, `location`, `display` and `tooltip` are XML-decoded (`&amp;` -> `&`,
  `&apos;` -> `'`), in `CalamineCell.hyperlink`, `CalamineSheetStream.hyperlinks` and
  `compatibility="openpyxl"`. Previously a URL such as `https://example.com/?a=1&b=2` came back
  as `...a=1&amp;b=2`, and a location such as `'My Sheet'!A1` as `&apos;My Sheet&apos;!A1`.
- Shared formulas are translated relative to the cell that holds the formula text (as Excel and
  openpyxl do), not the first cell of the shared range. When the two differ, every derived
  formula was shifted (e.g. `CP$4` instead of `CO$4`), in `CalamineCell.formula` and the
  whole-sheet formula API.
- `CalamineSheet.iter_rows()` on an empty sheet yields no rows instead of raising
  `PanicException`.

### Unchanged

- Without `compatibility`, `load_workbook` and the rest of the native API (`CalamineWorkbook`,
  streams, rich streams, whole-sheet API) behave as before; apart from the fixes above,
  their outputs were compared with the previous build. openpyxl-only options (`read_only`,
  `data_only`, ...) passed without `compatibility="openpyxl"` raise `TypeError`.
