use std::collections::{HashMap, VecDeque};
use std::fs::File;
use std::io::{BufReader, Cursor};
use std::sync::Arc;

use calamine::{
    Comment, DataRef, Dimensions, Error as CalamineCrateError, RowAttributes, StyleSheet, TextRun,
    WorksheetInfo, XlsxCellReader,
};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use crate::types::convert::{
    attrs_to_py, breaks_to_py, column_letters, columns_to_py, conditional_formats_to_py,
    data_validations_to_py, hyperlink_to_py, rich_text_to_py, row_dimensions_to_py, row_info_to_py,
    scenarios_to_py, sheet_settings_to_py, style_to_py, tables_to_py,
};
use crate::{CalamineWorkbook, CellValue, Error};

type MergedCellRange = ((u32, u32), (u32, u32));

/// A cell as read from the sheet XML.
struct RawCell {
    pos: (u32, u32),
    // Borrows the workbook's shared strings, like the reader.
    value: DataRef<'static>,
    style: u32,
    /// Formula text without the leading `=` (rich streams only).
    formula: Option<String>,
    /// Shared string index (rich streams only).
    sst: Option<usize>,
    /// Attributes of the cell's row, on the first cell read of each row (rich only).
    row_attrs: Option<RowAttributes>,
    /// Formatting runs of an inline rich string (rich only).
    runs: Option<Vec<TextRun>>,
    /// `normal`, `shared`, `array` or `dataTable` (rich only).
    formula_kind: Option<&'static str>,
    /// `<f>` attributes of array / data table formulas (rich only).
    formula_attrs: Option<Vec<(String, String)>>,
}

fn formula_kind(kind: &str) -> &'static str {
    match kind {
        "shared" => "shared",
        "array" => "array",
        "dataTable" => "dataTable",
        _ => "normal",
    }
}

/// openpyxl's `Cell.data_type` of a value (`n`, `s`, `b`, `e`, `d`).
fn data_type(value: &DataRef<'_>) -> &'static str {
    match value {
        DataRef::String(_) | DataRef::SharedString(_) => "s",
        DataRef::Bool(_) => "b",
        DataRef::Error(_) => "e",
        DataRef::DateTime(_) | DataRef::DateTimeIso(_) | DataRef::DurationIso(_) => "d",
        _ => "n",
    }
}

impl RawCell {
    /// Whether the cell holds anything worth a row: a value, or (rich) a formula
    /// whose result was never saved, e.g. in files written by openpyxl.
    fn has_content(&self) -> bool {
        !matches!(self.value, DataRef::Empty) || self.formula.is_some()
    }
}

/// A cell reader borrowing from the workbook's own `Xlsx`. The `'static` lifetime is
/// a lie upheld by `CalamineWorkbook`: it owns the `StreamState` and drops it before
/// it touches, replaces or drops the `Xlsx` the reader borrows from.
pub(crate) enum CellReader {
    File(XlsxCellReader<'static, BufReader<File>>),
    FileLike(XlsxCellReader<'static, Cursor<Vec<u8>>>),
}

macro_rules! read_cell {
    ($r:expr, $rich:expr) => {{
        if $rich {
            $r.next_cell_with_formula().map(|c| {
                c.map(|c| RawCell {
                    pos: c.pos,
                    value: c.value,
                    style: $r.last_style_id(),
                    formula: c.formula,
                    sst: $r.last_shared_string_index(),
                    row_attrs: None,
                    runs: $r.take_inline_runs(),
                    formula_kind: $r.last_formula_kind().map(formula_kind),
                    formula_attrs: Some($r.last_formula_attributes())
                        .filter(|a| !a.is_empty())
                        .map(|a| a.to_vec()),
                })
            })
        } else {
            $r.next_cell().map(|c| {
                c.map(|c| RawCell {
                    pos: c.get_position(),
                    value: c.get_value().clone(),
                    style: $r.last_style_id(),
                    formula: None,
                    sst: None,
                    row_attrs: None,
                    runs: None,
                    formula_kind: None,
                    formula_attrs: None,
                })
            })
        }
    }};
}

macro_rules! each_reader {
    ($self:expr, $r:ident => $body:expr) => {
        match $self {
            CellReader::File($r) => $body,
            CellReader::FileLike($r) => $body,
        }
    };
}

impl CellReader {
    /// Next cell with its style id; formula and shared string index when `rich`.
    fn next_cell(&mut self, rich: bool) -> Result<Option<RawCell>, Error> {
        each_reader!(self, r => read_cell!(r, rich))
            .map_err(|e| Error::Calamine(CalamineCrateError::Xlsx(e)))
    }

    fn current_row_attributes(&self) -> Option<RowAttributes> {
        each_reader!(self, r => r.current_row_attributes().cloned())
    }

    fn take_cellless_rows(&mut self) -> Vec<RowAttributes> {
        each_reader!(self, r => r.take_cellless_rows())
    }

    fn dimensions(&self) -> Dimensions {
        each_reader!(self, r => r.dimensions())
    }

    fn set_capture_inline_runs(&mut self, capture: bool) {
        each_reader!(self, r => r.set_capture_inline_runs(capture))
    }
}

/// All cells the file stores for one row.
struct PhysicalRow {
    index: u32,
    cells: Vec<RawCell>,
    has_content: bool,
    attrs: Option<RowAttributes>,
}

/// Rows/columns to stream (0-based, inclusive).
#[derive(Clone, Copy, Default)]
pub(crate) struct StreamRange {
    pub(crate) min_row: u32,
    pub(crate) max_row: Option<u32>,
    pub(crate) min_col: u32,
    pub(crate) max_col: Option<u32>,
}

/// Per-cell details of a rich stream.
#[derive(Clone)]
pub(crate) struct CellExtra {
    style: u32,
    formula: Option<String>,
    sst: Option<usize>,
    runs: Option<Vec<TextRun>>,
    data_type: &'static str,
    formula_kind: Option<&'static str>,
    formula_attrs: Option<Vec<(String, String)>>,
}

impl Default for CellExtra {
    fn default() -> Self {
        Self {
            style: 0,
            formula: None,
            sst: None,
            runs: None,
            data_type: "n",
            formula_kind: None,
            formula_attrs: None,
        }
    }
}

/// One row as yielded to Python: values, plus per-cell details when streaming rich.
pub(crate) struct StreamRow {
    pub(crate) index: u32,
    pub(crate) first_col: u32,
    pub(crate) values: Vec<CellValue>,
    pub(crate) extras: Vec<CellExtra>,
    pub(crate) attrs: Option<RowAttributes>,
}

/// Position of an in-progress stream over one worksheet, owned by the workbook.
pub(crate) struct StreamState {
    pub(crate) id: u64,
    rich: bool,
    range: StreamRange,
    // Both borrow the workbook's shared strings, like the reader.
    pending: Option<RawCell>,
    queue: VecDeque<PhysicalRow>,
    // Attributes of rows without cells (rich only), in row order.
    cellless: VecDeque<RowAttributes>,
    last_cell_row: Option<u32>,
    // `None` for non-worksheets (e.g. chart sheets), which stream as empty, like `get_sheet_by_name`.
    reader: Option<CellReader>,
    width: usize,
    next_row: u32,
    eof: bool,
}

impl StreamState {
    pub(crate) fn new(
        id: u64,
        mut reader: Option<CellReader>,
        rich: bool,
        range: StreamRange,
    ) -> Self {
        if let Some(r) = reader.as_mut() {
            r.set_capture_inline_runs(rich);
        }
        // The <dimension> tag is only a hint for the initial row width; rows grow
        // if a cell lies beyond it.
        let width = match &reader {
            Some(r) if r.dimensions().len() > 0 => r.dimensions().end.1 as usize + 1,
            _ => 0,
        };
        Self {
            id,
            rich,
            range,
            pending: None,
            queue: VecDeque::new(),
            cellless: VecDeque::new(),
            last_cell_row: None,
            reader,
            width,
            next_row: 0,
            eof: false,
        }
    }

    fn next_cell(&mut self) -> Result<Option<RawCell>, Error> {
        if let Some(cell) = self.pending.take() {
            return Ok(Some(cell));
        }
        let Some(reader) = self.reader.as_mut().filter(|_| !self.eof) else {
            return Ok(None);
        };
        loop {
            let cell = reader.next_cell(self.rich)?;
            if self.rich {
                self.cellless.extend(reader.take_cellless_rows());
            }
            match cell {
                // Value-less cells (formatted but empty) only matter for their style.
                Some(c) if !self.rich && matches!(c.value, DataRef::Empty) => continue,
                Some(mut cell) => {
                    if self.rich && self.last_cell_row != Some(cell.pos.0) {
                        cell.row_attrs = reader.current_row_attributes();
                        self.last_cell_row = Some(cell.pos.0);
                    }
                    return Ok(Some(cell));
                }
                None => {
                    self.eof = true;
                    return Ok(None);
                }
            }
        }
    }

    /// Reads all cells of the next row stored in the file.
    fn read_physical_row(&mut self) -> Result<Option<PhysicalRow>, Error> {
        let Some(mut first) = self.next_cell()? else {
            return Ok(None);
        };
        let index = first.pos.0;
        let mut row = PhysicalRow {
            index,
            has_content: first.has_content(),
            attrs: first.row_attrs.take(),
            cells: vec![first],
        };
        while let Some(cell) = self.next_cell()? {
            let r = cell.pos.0;
            if r == index {
                row.has_content |= cell.has_content();
                row.cells.push(cell);
            } else if r > index {
                self.pending = Some(cell);
                break;
            } else {
                return Err(Error::Calamine(CalamineCrateError::Msg(
                    "cells are not stored in row order; streaming is not possible for this sheet",
                )));
            }
        }
        Ok(Some(row))
    }

    /// Queues rows up to (and including) the next one with content. Rows without
    /// content are kept (for their styles) only if a row with content follows them.
    fn fill_queue(&mut self) -> Result<(), Error> {
        loop {
            match self.read_physical_row()? {
                None => {
                    self.queue.clear();
                    return Ok(());
                }
                Some(row) => {
                    if let Some(last) = self.queue.back() {
                        if row.index <= last.index {
                            return Err(Error::Calamine(CalamineCrateError::Msg(
                                "rows are not stored in order; streaming is not possible for this sheet",
                            )));
                        }
                    }
                    let has_content = row.has_content;
                    self.queue.push_back(row);
                    if has_content {
                        return Ok(());
                    }
                }
            }
        }
    }

    /// Returns the next row, or `None` once the sheet (or the requested range) is exhausted.
    pub(crate) fn next_row(&mut self) -> Result<Option<StreamRow>, Error> {
        let range = self.range;
        loop {
            if range.max_row.is_some_and(|max| self.next_row > max) {
                // Stop early: the rest of the sheet is never read.
                self.queue.clear();
                return Ok(None);
            }
            if self.queue.is_empty() {
                self.fill_queue()?;
            }
            let Some(front) = self.queue.front() else {
                return Ok(None);
            };
            if self.next_row < range.min_row {
                // Skip rows before the range without converting their values.
                if front.index < range.min_row {
                    self.queue.pop_front();
                } else {
                    self.next_row = range.min_row;
                }
                continue;
            }
            break;
        }
        let front_index = self.queue.front().expect("checked").index;

        let index = self.next_row;
        self.next_row += 1;
        let first_col = range.min_col;
        let width = match range.max_col {
            Some(max) => (max as usize + 1).saturating_sub(first_col as usize),
            None => self.width.saturating_sub(first_col as usize),
        };
        let rich_width = if self.rich { width } else { 0 };
        let mut out = StreamRow {
            index,
            first_col,
            values: vec![CellValue::Empty; width],
            extras: vec![CellExtra::default(); rich_width],
            attrs: None,
        };

        // Attributes of rows without cells, for gap rows.
        while self.cellless.front().is_some_and(|a| a.index < index) {
            self.cellless.pop_front();
        }
        let cellless_attrs = self
            .cellless
            .front()
            .is_some_and(|a| a.index == index)
            .then(|| self.cellless.pop_front())
            .flatten();

        if front_index > index {
            // Gap in the sheet: emit an empty row.
            out.attrs = cellless_attrs;
            return Ok(Some(out));
        }

        let row = self.queue.pop_front().expect("front exists");
        out.attrs = row.attrs.or(cellless_attrs);
        for cell in row.cells {
            let col = cell.pos.1;
            self.width = self.width.max(col as usize + 1);
            if col < first_col || range.max_col.is_some_and(|max| col > max) {
                continue;
            }
            let i = (col - first_col) as usize;
            if i >= out.values.len() {
                out.values.resize(i + 1, CellValue::Empty);
                if self.rich {
                    out.extras.resize(i + 1, CellExtra::default());
                }
            }
            out.values[i] = CellValue::from(&cell.value);
            if self.rich {
                out.extras[i] = CellExtra {
                    style: cell.style,
                    formula: cell.formula.map(|f| format!("={f}")),
                    sst: cell.sst,
                    runs: cell.runs,
                    data_type: data_type(&cell.value),
                    formula_kind: cell.formula_kind,
                    formula_attrs: cell.formula_attrs,
                };
            }
        }
        Ok(Some(out))
    }
}

/// Extra sheet information loaded when a stream is started with `rich=True`.
pub(crate) struct RichInfo {
    pub(crate) info: WorksheetInfo,
    pub(crate) comments: Vec<Comment>,
    pub(crate) stylesheet: Arc<StyleSheet>,
    pub(crate) rich_strings: Arc<HashMap<usize, Vec<TextRun>>>,
    /// The sheet's `_xlnm.Print_Area` / `_xlnm.Print_Titles` defined names.
    pub(crate) print_area: Option<String>,
    pub(crate) print_titles: Option<String>,
}

/// Finds the range (merged cells, hyperlinks) containing a cell while rows are
/// visited in increasing order.
struct RangeCursor {
    ranges: Vec<(MergedCellRange, usize)>,
    next: usize,
    active: Vec<(MergedCellRange, usize)>,
}

impl RangeCursor {
    fn new(mut ranges: Vec<(MergedCellRange, usize)>) -> Self {
        ranges.sort();
        Self {
            ranges,
            next: 0,
            active: Vec::new(),
        }
    }

    fn enter_row(&mut self, row: u32) {
        self.active.retain(|(m, _)| m.1 .0 >= row);
        while let Some(&(m, i)) = self.ranges.get(self.next) {
            if m.0 .0 > row {
                break;
            }
            if m.1 .0 >= row {
                self.active.push((m, i));
            }
            self.next += 1;
        }
    }

    fn find(&self, col: u32) -> Option<(MergedCellRange, usize)> {
        self.active
            .iter()
            .find(|(m, _)| m.0 .1 <= col && col <= m.1 .1)
            .copied()
    }
}

/// Python-side view of `RichInfo`.
struct RichView {
    info: WorksheetInfo,
    stylesheet: Arc<StyleSheet>,
    styles: Vec<Py<PyDict>>,
    comments: HashMap<(u32, u32), Py<PyDict>>,
    merged: RangeCursor,
    hyperlinks: Vec<Py<PyDict>>,
    hyperlink_cursor: RangeCursor,
    rich_strings: Arc<HashMap<usize, Vec<TextRun>>>,
    rich_cache: HashMap<usize, Py<PyList>>,
    // Index and attributes of the row last yielded; converted on access.
    row_info: Option<(u32, Option<RowAttributes>)>,
    print_area: Option<String>,
    print_titles: Option<String>,
}

impl RichView {
    fn new(py: Python<'_>, rich: RichInfo) -> PyResult<Self> {
        let styles = rich
            .stylesheet
            .cell_styles
            .iter()
            .map(|s| style_to_py(py, s).map(Bound::unbind))
            .collect::<PyResult<_>>()?;
        let mut comments = HashMap::with_capacity(rich.comments.len());
        for c in rich.comments {
            let d = PyDict::new(py);
            d.set_item("author", c.author)?;
            d.set_item("text", c.text)?;
            comments.insert(c.pos, d.unbind());
        }
        let merged = RangeCursor::new(
            rich.info
                .merged
                .iter()
                .enumerate()
                .map(|(i, d)| ((d.start, d.end), i))
                .collect(),
        );
        let hyperlinks = rich
            .info
            .hyperlinks
            .iter()
            .map(|h| hyperlink_to_py(py, h).map(Bound::unbind))
            .collect::<PyResult<Vec<_>>>()?;
        let hyperlink_cursor = RangeCursor::new(
            rich.info
                .hyperlinks
                .iter()
                .enumerate()
                .map(|(i, h)| ((h.range.start, h.range.end), i))
                .collect(),
        );
        Ok(Self {
            info: rich.info,
            stylesheet: rich.stylesheet,
            styles,
            comments,
            merged,
            hyperlinks,
            hyperlink_cursor,
            rich_strings: rich.rich_strings,
            rich_cache: HashMap::new(),
            row_info: None,
            print_area: rich.print_area,
            print_titles: rich.print_titles,
        })
    }

    fn rich_text(&mut self, py: Python<'_>, sst: Option<usize>) -> PyResult<Option<Py<PyList>>> {
        let Some(i) = sst.filter(|_| !self.rich_strings.is_empty()) else {
            return Ok(None);
        };
        if let Some(cached) = self.rich_cache.get(&i) {
            return Ok(Some(cached.clone_ref(py)));
        }
        let Some(runs) = self.rich_strings.get(&i) else {
            return Ok(None);
        };
        let list = rich_text_to_py(py, runs)?.unbind();
        self.rich_cache.insert(i, list.clone_ref(py));
        Ok(Some(list))
    }
}

/// A cell yielded by a stream started with `rich=True`.
#[pyclass(frozen)]
pub struct CalamineCell {
    #[pyo3(get)]
    value: Py<PyAny>,
    #[pyo3(get)]
    row: u32,
    #[pyo3(get)]
    column: u32,
    #[pyo3(get)]
    style_id: u32,
    #[pyo3(get)]
    style: Option<Py<PyDict>>,
    #[pyo3(get)]
    comment: Option<Py<PyDict>>,
    #[pyo3(get)]
    merged_range: Option<MergedCellRange>,
    /// Formula text with a leading `=` (e.g. `=SUM(A1:A3)`), or `None`.
    #[pyo3(get)]
    formula: Option<String>,
    #[pyo3(get)]
    hyperlink: Option<Py<PyDict>>,
    #[pyo3(get)]
    rich_text: Option<Py<PyList>>,
    /// openpyxl's data type of the value: `n`, `s`, `b`, `e` or `d`.
    #[pyo3(get)]
    data_type: &'static str,
    /// `normal`, `shared`, `array` or `dataTable`; `None` without a formula.
    #[pyo3(get)]
    formula_type: Option<&'static str>,
    /// Range of an array formula (openpyxl's `ArrayFormula.ref`), on its anchor cell.
    #[pyo3(get)]
    formula_range: Option<String>,
    /// Attributes of an array / data table formula (`ref`, `r1`, `dt2D`, ...).
    #[pyo3(get)]
    formula_attributes: Option<Py<PyDict>>,
}

#[pymethods]
impl CalamineCell {
    /// Excel reference, e.g. `B3`.
    #[getter]
    fn coordinate(&self) -> String {
        format!("{}{}", column_letters(self.column), self.row + 1)
    }

    /// The value is a date / time / duration (openpyxl's `is_date`).
    #[getter]
    fn is_date(&self) -> bool {
        self.data_type == "d"
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        let mut extra = String::new();
        if let Some(f) = &self.formula {
            extra.push_str(&format!(", formula={f:?}"));
        }
        if self.comment.is_some() {
            extra.push_str(", comment=...");
        }
        if self.hyperlink.is_some() {
            extra.push_str(", hyperlink=...");
        }
        if self.rich_text.is_some() {
            extra.push_str(", rich_text=...");
        }
        if let Some(m) = self.merged_range {
            extra.push_str(&format!(", merged_range={m:?}"));
        }
        Ok(format!(
            "CalamineCell(row={}, column={}, value={}, style_id={}{extra})",
            self.row,
            self.column,
            self.value.bind(py).repr()?,
            self.style_id,
        ))
    }
}

/// Row-by-row streaming reader over a single xlsx worksheet.
///
/// Unlike `CalamineSheet`, the sheet is never materialized as a `Range`: cells are
/// pulled from the zipped XML one at a time, so memory stays bounded by one row
/// (plus the workbook's shared strings table, which is reused, not reloaded).
///
/// Rows are yielded from row 0 and column 0 (or the requested range), with empty
/// rows/cells filled with `""`. With `rich=True`, rows are lists of `CalamineCell`
/// and the stream exposes the sheet's metadata. A workbook has at most one active
/// stream; starting another one, loading a sheet or closing the workbook invalidates it.
#[pyclass]
pub struct CalamineSheetStream {
    #[pyo3(get)]
    name: String,
    workbook: Py<CalamineWorkbook>,
    id: u64,
    done: bool,
    rich: Option<RichView>,
}

impl CalamineSheetStream {
    pub(crate) fn new(
        py: Python<'_>,
        name: String,
        workbook: Py<CalamineWorkbook>,
        id: u64,
        rich: Option<RichInfo>,
    ) -> PyResult<Self> {
        Ok(Self {
            name,
            workbook,
            id,
            done: false,
            rich: rich.map(|info| RichView::new(py, info)).transpose()?,
        })
    }

    fn rich_row<'py>(&mut self, py: Python<'py>, row: StreamRow) -> PyResult<Bound<'py, PyList>> {
        let view = self.rich.as_mut().expect("rich stream");
        view.merged.enter_row(row.index);
        view.hyperlink_cursor.enter_row(row.index);
        view.row_info = Some((row.index, row.attrs));
        let list = PyList::empty(py);
        let cells = row.values.into_iter().zip(row.extras).enumerate();
        for (i, (value, extra)) in cells {
            let col = row.first_col + i as u32;
            let style_id = extra.style;
            let formula_range = extra
                .formula_attrs
                .as_ref()
                .filter(|_| extra.formula_kind == Some("array"))
                .and_then(|a| a.iter().find(|(k, _)| k == "ref"))
                .map(|(_, v)| v.clone());
            let formula_attributes = extra
                .formula_attrs
                .as_ref()
                .map(|a| attrs_to_py(py, a).map(Bound::unbind))
                .transpose()?;
            let cell = CalamineCell {
                value: value.into_pyobject(py)?.unbind(),
                row: row.index,
                column: col,
                style_id,
                style: view.styles.get(style_id as usize).map(|s| s.clone_ref(py)),
                comment: view
                    .comments
                    .get(&(row.index, col))
                    .map(|c| c.clone_ref(py)),
                merged_range: view.merged.find(col).map(|(m, _)| m),
                formula: extra.formula,
                hyperlink: view
                    .hyperlink_cursor
                    .find(col)
                    .and_then(|(_, h)| view.hyperlinks.get(h))
                    .map(|h| h.clone_ref(py)),
                rich_text: match extra.runs {
                    Some(runs) => Some(rich_text_to_py(py, &runs)?.unbind()),
                    None => view.rich_text(py, extra.sst)?,
                },
                data_type: extra.data_type,
                formula_type: extra.formula_kind,
                formula_range,
                formula_attributes,
            };
            list.append(Py::new(py, cell)?)?;
        }
        Ok(list)
    }
}

#[pymethods]
impl CalamineSheetStream {
    fn __repr__(&self) -> String {
        format!(
            "CalamineSheetStream(name='{}', rich={})",
            self.name,
            self.rich.is_some()
        )
    }

    #[getter]
    fn rich(&self) -> bool {
        self.rich.is_some()
    }

    /// Merged cell ranges of the sheet (`rich=True` only).
    #[getter]
    fn merged_cell_ranges(&self) -> Option<Vec<MergedCellRange>> {
        self.rich
            .as_ref()
            .map(|r| r.info.merged.iter().map(|d| (d.start, d.end)).collect())
    }

    /// Comments keyed by `(row, column)` (`rich=True` only), including comments
    /// on cells outside the streamed rows.
    #[getter]
    fn comments<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyDict>>> {
        let Some(view) = &self.rich else {
            return Ok(None);
        };
        // Sheet order (row, then column), independent of the map's hash order.
        let mut positions: Vec<&(u32, u32)> = view.comments.keys().collect();
        positions.sort();
        let d = PyDict::new(py);
        for pos in positions {
            d.set_item(pos, view.comments[pos].bind(py))?;
        }
        Ok(Some(d))
    }

    /// The workbook's cell formats, indexed by `CalamineCell.style_id` (`rich=True` only).
    #[getter]
    fn styles<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyList>>> {
        let Some(view) = &self.rich else {
            return Ok(None);
        };
        Ok(Some(PyList::new(
            py,
            view.styles.iter().map(|s| s.bind(py)),
        )?))
    }

    /// Hyperlinks of the sheet (`rich=True` only).
    #[getter]
    fn hyperlinks<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyList>>> {
        let Some(view) = &self.rich else {
            return Ok(None);
        };
        Ok(Some(PyList::new(
            py,
            view.hyperlinks.iter().map(|h| h.bind(py)),
        )?))
    }

    /// Conditional formatting of the sheet (`rich=True` only).
    #[getter]
    fn conditional_formats<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyList>>> {
        let Some(view) = &self.rich else {
            return Ok(None);
        };
        conditional_formats_to_py(
            py,
            &view.info.conditional_formats,
            &view.stylesheet.differential_styles,
        )
        .map(Some)
    }

    /// Data validation rules of the sheet (`rich=True` only).
    #[getter]
    fn data_validations<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyList>>> {
        let Some(view) = &self.rich else {
            return Ok(None);
        };
        data_validations_to_py(py, &view.info.data_validations).map(Some)
    }

    /// Column widths / visibility / outline levels (`rich=True` only).
    #[getter]
    fn column_dimensions<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyList>>> {
        let Some(view) = &self.rich else {
            return Ok(None);
        };
        columns_to_py(py, &view.info.columns).map(Some)
    }

    /// Freeze panes, auto filter, protection, page setup, view and tab color (`rich=True` only).
    #[getter]
    fn sheet_settings<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyDict>>> {
        let Some(view) = &self.rich else {
            return Ok(None);
        };
        sheet_settings_to_py(
            py,
            &view.info,
            view.print_area.as_deref(),
            view.print_titles.as_deref(),
        )
        .map(Some)
    }

    /// `<dimension>` reference declared by the writer, e.g. `A1:H100` (`rich=True` only).
    #[getter]
    fn dimension(&self) -> Option<String> {
        self.rich.as_ref().and_then(|v| v.info.dimension.clone())
    }

    /// Tables (list objects) of the sheet (`rich=True` only).
    #[getter]
    fn tables<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyList>>> {
        let Some(view) = &self.rich else {
            return Ok(None);
        };
        tables_to_py(py, &view.info.tables).map(Some)
    }

    /// All rows with non-default height / hidden / outline / style, keyed by row
    /// index, including rows outside the streamed data (`rich=True` only).
    #[getter]
    fn row_dimensions<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyDict>>> {
        let Some(view) = &self.rich else {
            return Ok(None);
        };
        row_dimensions_to_py(py, &view.info.rows).map(Some)
    }

    /// Manual row page breaks (`rich=True` only).
    #[getter]
    fn row_breaks<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyList>>> {
        let Some(view) = &self.rich else {
            return Ok(None);
        };
        breaks_to_py(py, &view.info.row_breaks).map(Some)
    }

    /// Manual column page breaks (`rich=True` only).
    #[getter]
    fn column_breaks<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyList>>> {
        let Some(view) = &self.rich else {
            return Ok(None);
        };
        breaks_to_py(py, &view.info.column_breaks).map(Some)
    }

    /// What-if scenarios (`rich=True` only; `None` also when the sheet has none).
    #[getter]
    fn scenarios<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyDict>>> {
        let Some(view) = &self.rich else {
            return Ok(None);
        };
        scenarios_to_py(py, &view.info.scenarios)
    }

    /// Height / hidden / outline level of the row last yielded (`rich=True` only).
    #[getter]
    fn row_info<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyDict>>> {
        self.rich
            .as_ref()
            .and_then(|v| v.row_info.as_ref())
            .map(|(row, attrs)| row_info_to_py(py, *row, attrs.as_ref()))
            .transpose()
    }

    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__<'py>(&mut self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyList>>> {
        if self.done {
            return Ok(None);
        }
        let row = self
            .workbook
            .bind(py)
            .borrow_mut()
            .next_stream_row(self.id)?;
        match row {
            None => {
                self.done = true;
                Ok(None)
            }
            Some(row) if self.rich.is_some() => Ok(Some(self.rich_row(py, row)?)),
            Some(row) => Ok(Some(PyList::new(py, row.values)?)),
        }
    }
}

impl Drop for CalamineSheetStream {
    fn drop(&mut self) {
        // Release the reader's buffers as soon as the stream is gone.
        if !self.done {
            Python::attach(|py| {
                if let Ok(mut wb) = self.workbook.bind(py).try_borrow_mut() {
                    wb.end_stream(self.id);
                }
            });
        }
    }
}
