//! openpyxl-shaped worksheets, cells and row iterators.
//!
//! Normal mode (`read_only=False`) follows openpyxl's `Worksheet`: bounds are the
//! cells the file stores plus merged ranges, hyperlinks and comments; merged
//! placeholders are `MergedCell`s; missing cells are empty `Cell`s. Read-only mode
//! follows `ReadOnlyWorksheet`: bounds come from the declared `<dimension>`, rows are
//! padded with `EmptyCell`, and comments, hyperlinks and merges are not applied.
//!
//! Nothing is materialized: rows are streamed, random access goes through a bounded
//! row cache, and every iterator has its own reader.

use std::collections::{BTreeMap, HashMap, VecDeque};
use std::sync::{Arc, Mutex};

use calamine::{Dimensions, WorksheetInfo};
use pyo3::exceptions::{PyAttributeError, PyIndexError, PyValueError};
use pyo3::prelude::*;
use pyo3::sync::PyOnceLock;
use pyo3::types::{PyDict, PyIterator, PyList, PySlice, PyTuple};
use pyo3::{PyTraverseError, PyVisit};

use super::objects::{
    AutoFilter, ColumnDimension, Comment, DefinedName, DimensionHolder, Hyperlink, MergedCellRange,
    MultiCellRange, RowDimension, Table, TableColumn, TableList, TableStyleInfo,
    WorksheetProperties,
};
use super::reader::{with_main, CCell, CExtra, CRow, CVal, Shared, SheetReader, ZERO_STYLE};
use super::styles::{edge, Border, Color, Protection, Side};
use super::utils::{column_letter, coordinate, quote_sheetname, range_boundaries, sheet_ranges};
use super::workbook::Workbook;
use crate::Error;

/// 1-based `(min_row, min_col, max_row, max_col)`.
pub(crate) type Bounds = (u32, u32, u32, u32);

const MAX_ROW: u32 = 1_048_576;
/// Hyperlink ranges larger than this are resolved by scanning instead of per-cell entries.
const LINK_EXPAND_LIMIT: u64 = 100_000;
/// Merges taller than this are checked by scanning instead of per-row entries.
const MERGE_INDEX_LIMIT: u32 = 10_000;
/// Row cache limits for random access (`ws.cell`, `ws["A1"]`).
const CACHE_ROWS: usize = 4096;
const CACHE_CELLS: usize = 1_000_000;

fn union(a: Option<Bounds>, b: Bounds) -> Option<Bounds> {
    Some(match a {
        None => b,
        Some(a) => (a.0.min(b.0), a.1.min(b.1), a.2.max(b.2), a.3.max(b.3)),
    })
}

fn dims_bounds(d: &Dimensions) -> Bounds {
    (d.start.0 + 1, d.start.1 + 1, d.end.0 + 1, d.end.1 + 1)
}

fn in_bounds(b: Bounds, r: u32, c: u32) -> bool {
    b.0 <= r && r <= b.2 && b.1 <= c && c <= b.3
}

/// Python's `x or default` for optional 1-based indexes (0 counts as unset).
fn or_none(v: Option<u32>) -> Option<u32> {
    v.filter(|v| *v != 0)
}

// ---- sheet metadata ----------------------------------------------------------------

/// Effective borders of a merged range (openpyxl's `MergedCellRange.format`).
struct MergedFormat {
    /// Border of the top-left cell after `_get_borders`.
    anchor_border: Border,
    protection: Protection,
    /// Sides propagated to the top, left, right and bottom edges.
    edges: [(usize, Option<Side>); 4],
}

struct LinkData {
    target: Option<String>,
    location: Option<String>,
    tooltip: Option<String>,
    display: Option<String>,
}

impl LinkData {
    fn from_raw(h: &calamine::Hyperlink) -> Self {
        Self {
            target: h.target.clone(),
            location: h.location.clone(),
            tooltip: h.tooltip.clone(),
            display: h.displayed_text.clone(),
        }
    }

    fn to_py(&self, py: Python<'_>, coord: String) -> PyResult<Py<Hyperlink>> {
        Py::new(
            py,
            Hyperlink {
                reference: coord,
                target: self.target.clone(),
                location: self.location.clone(),
                tooltip: self.tooltip.clone(),
                display: self.display.clone(),
            },
        )
    }
}

/// Metadata of a sheet, read with one pass over the sheet XML (plus, in normal mode
/// with merged cells, one pass over the rows holding merge anchors).
pub(crate) struct SheetMeta {
    info: WorksheetInfo,
    merged: Vec<Bounds>,
    merged_by_row: HashMap<u32, Vec<usize>>,
    tall_merged: Vec<usize>,
    merged_fmt: Vec<MergedFormat>,
    comments: HashMap<(u32, u32), (String, Option<String>)>,
    links: Vec<LinkData>,
    link_cells: HashMap<(u32, u32), usize>,
    link_ranges: Vec<(Bounds, usize)>,
    /// Bounds of all cells openpyxl's normal mode would hold.
    bounds: Option<Bounds>,
    /// Normal (openpyxl `Worksheet`) binding rules, else read-only (native) rules.
    normal: bool,
}

impl SheetMeta {
    fn load(
        shared: &Arc<Shared>,
        title: &str,
        part: Option<&str>,
        normal: bool,
    ) -> Result<Self, Error> {
        let (info, comments) = shared.with_main(|m| {
            with_main!(m, x => {
                let info = if normal {
                    x.worksheet_info_with_bounds(title)?
                } else {
                    x.worksheet_info(title)?
                };
                Ok((info, x.worksheet_comments(title)?))
            })
        })?;
        let merged: Vec<Bounds> = info.merged.iter().map(dims_bounds).collect();
        let mut merged_by_row: HashMap<u32, Vec<usize>> = HashMap::new();
        let mut tall_merged = Vec::new();
        for (i, m) in merged.iter().enumerate() {
            if m.2 - m.0 >= MERGE_INDEX_LIMIT {
                tall_merged.push(i);
            } else {
                for r in m.0..=m.2 {
                    merged_by_row.entry(r).or_default().push(i);
                }
            }
        }
        let mut meta = SheetMeta {
            merged,
            merged_by_row,
            tall_merged,
            merged_fmt: Vec::new(),
            comments: HashMap::new(),
            links: Vec::new(),
            link_cells: HashMap::new(),
            link_ranges: Vec::new(),
            bounds: None,
            normal,
            info,
        };
        if !normal {
            // Read-only mode binds nothing by default; with `read_comments` /
            // `read_hyperlinks` cells get them like the native rich stream does: every
            // cell of a link's range (merged cells included) and every comment.
            let links = std::mem::take(&mut meta.info.hyperlinks);
            for (i, h) in links.iter().enumerate() {
                let b = dims_bounds(&h.range);
                if (b.2 - b.0 + 1) as u64 * (b.3 - b.1 + 1) as u64 <= LINK_EXPAND_LIMIT {
                    for r in b.0..=b.2 {
                        for c in b.1..=b.3 {
                            meta.link_cells.insert((r, c), i);
                        }
                    }
                } else {
                    meta.link_ranges.push((b, i));
                }
                meta.links.push(LinkData::from_raw(h));
            }
            meta.info.hyperlinks = links;
            for c in comments {
                meta.comments
                    .insert((c.pos.0 + 1, c.pos.1 + 1), (c.text, c.author));
            }
            return Ok(meta);
        }

        let mut bounds = meta.info.cell_bounds.as_ref().map(dims_bounds);
        for m in &meta.merged {
            bounds = union(bounds, *m);
        }
        // Hyperlinks, in file order: later links win, merged placeholders are skipped
        // (single-cell links on them move to the anchor), as in openpyxl's reader.
        let links = std::mem::take(&mut meta.info.hyperlinks);
        for (i, h) in links.iter().enumerate() {
            let b = dims_bounds(&h.range);
            if (b.0, b.1) == (b.2, b.3) {
                let pos = match meta.merge_at(b.0, b.1) {
                    Some(m) if (meta.merged[m].0, meta.merged[m].1) != (b.0, b.1) => {
                        (meta.merged[m].0, meta.merged[m].1)
                    }
                    _ => (b.0, b.1),
                };
                meta.link_cells.insert(pos, i);
                bounds = union(bounds, (pos.0, pos.1, pos.0, pos.1));
            } else {
                let size = (b.2 - b.0 + 1) as u64 * (b.3 - b.1 + 1) as u64;
                if size <= LINK_EXPAND_LIMIT {
                    for r in b.0..=b.2 {
                        for c in b.1..=b.3 {
                            if !meta.is_placeholder(r, c) {
                                meta.link_cells.insert((r, c), i);
                            }
                        }
                    }
                } else {
                    meta.link_ranges.push((b, i));
                }
                bounds = union(bounds, b);
            }
            meta.links.push(LinkData::from_raw(h));
        }
        meta.info.hyperlinks = links;
        for c in comments {
            let (r, col) = (c.pos.0 + 1, c.pos.1 + 1);
            if !meta.is_placeholder(r, col) {
                meta.comments.insert((r, col), (c.text, c.author));
                bounds = union(bounds, (r, col, r, col));
            }
        }
        meta.bounds = bounds;
        if !meta.merged.is_empty() {
            meta.merged_fmt = merged_formats(shared, part, &meta.merged)?;
        }
        Ok(meta)
    }

    /// Index of the merged range containing a cell.
    fn merge_at(&self, r: u32, c: u32) -> Option<usize> {
        if let Some(list) = self.merged_by_row.get(&r) {
            if let Some(i) = list.iter().find(|i| in_bounds(self.merged[**i], r, c)) {
                return Some(*i);
            }
        }
        self.tall_merged
            .iter()
            .find(|i| in_bounds(self.merged[**i], r, c))
            .copied()
    }

    fn is_placeholder(&self, r: u32, c: u32) -> bool {
        self.merge_at(r, c)
            .is_some_and(|m| (self.merged[m].0, self.merged[m].1) != (r, c))
    }

    fn link_at(&self, r: u32, c: u32) -> Option<&LinkData> {
        let a = self.link_cells.get(&(r, c)).copied();
        let b = self
            .link_ranges
            .iter()
            .filter(|(b, _)| in_bounds(*b, r, c) && !(self.normal && self.is_placeholder(r, c)))
            .map(|(_, i)| *i)
            .max();
        a.max(b).map(|i| &self.links[i])
    }

    /// Border of a cell of merged range `m` (`base` is its own border before formatting).
    fn merged_border(&self, m: usize, r: u32, c: u32, base: &Border) -> Border {
        let b = self.merged[m];
        let mut border = base.clone();
        for (edge_idx, side) in &self.merged_fmt[m].edges {
            let Some(side) = side else { continue };
            let on_edge = match *edge_idx {
                edge::TOP => r == b.0,
                edge::LEFT => c == b.1,
                edge::RIGHT => c == b.3,
                _ => r == b.2,
            };
            if on_edge {
                border = border.add(&Border::single(*edge_idx, Some(side.clone())));
            }
        }
        border
    }
}

/// openpyxl's `_get_borders` + `format()` for every merged range: needs the styles of
/// each range's top-left and bottom-right cells, read in one forward pass.
fn merged_formats(
    shared: &Arc<Shared>,
    part: Option<&str>,
    merged: &[Bounds],
) -> Result<Vec<MergedFormat>, Error> {
    let mut wanted: BTreeMap<u32, Vec<(u32, usize, bool)>> = BTreeMap::new();
    for (i, m) in merged.iter().enumerate() {
        wanted.entry(m.0).or_default().push((m.1, i, true));
        wanted.entry(m.2).or_default().push((m.3, i, false));
    }
    let mut anchor: Vec<Option<u32>> = vec![None; merged.len()];
    let mut end: Vec<Option<u32>> = vec![None; merged.len()];
    let last = wanted.keys().next_back().copied().unwrap_or(0);
    let mut reader = SheetReader::open(shared, part)?;
    while let Some(row) = reader.next_row()? {
        if row.index > last {
            break;
        }
        if let Some(list) = wanted.get(&row.index) {
            for (col, i, is_anchor) in list {
                if let Some(c) = row.get(*col) {
                    if *is_anchor {
                        anchor[*i] = Some(c.style);
                    } else {
                        end[*i] = Some(c.style);
                    }
                }
            }
        }
    }
    Ok((0..merged.len())
        .map(|i| {
            let style = shared.style(anchor[i].unwrap_or(ZERO_STYLE));
            let mut anchor_border = style.border.clone();
            if let Some(e) = end[i] {
                let eb = &shared.style(e).border;
                let mut add = Border::default();
                add.sides[edge::RIGHT] = eb.side(edge::RIGHT).cloned();
                add.sides[edge::BOTTOM] = eb.side(edge::BOTTOM).cloned();
                anchor_border = anchor_border.add(&add);
            }
            let edges = [edge::TOP, edge::LEFT, edge::RIGHT, edge::BOTTOM].map(|e| {
                // openpyxl skips edges whose side has no style; a missing side adds nothing.
                let side = anchor_border.side(e).filter(|s| s.style.is_some()).cloned();
                (e, side)
            });
            MergedFormat {
                anchor_border,
                protection: style.protection().clone(),
                edges,
            }
        })
        .collect())
}

// ---- random access -----------------------------------------------------------------

/// A bounded, row-oriented cache in front of a forward-only reader.
#[derive(Default)]
struct RowCache {
    reader: Option<SheetReader>,
    /// Every stored row with an index in `[window_start, scanned]` is in `rows`.
    window_start: u32,
    scanned: u32,
    rows: BTreeMap<u32, Arc<CRow>>,
    cells: usize,
}

impl RowCache {
    fn row(
        &mut self,
        shared: &Arc<Shared>,
        part: Option<&str>,
        r: u32,
    ) -> Result<Option<Arc<CRow>>, Error> {
        shared.check_open()?;
        if self.reader.is_none() || r < self.window_start {
            self.reader = Some(SheetReader::open(shared, part)?);
            self.rows.clear();
            self.cells = 0;
            self.window_start = r;
            self.scanned = 0;
        }
        while self.scanned < r {
            match self.reader.as_mut().expect("opened").next_row()? {
                None => self.scanned = u32::MAX,
                Some(row) => {
                    self.scanned = row.index;
                    if row.index >= self.window_start {
                        self.cells += row.cells.len();
                        self.rows.insert(row.index, Arc::new(row));
                        while self.rows.len() > CACHE_ROWS || self.cells > CACHE_CELLS {
                            let (k, v) = self.rows.pop_first().expect("not empty");
                            self.cells -= v.cells.len();
                            self.window_start = k + 1;
                        }
                    }
                }
            }
        }
        Ok(self.rows.get(&r).cloned())
    }
}

struct SheetInner {
    meta: Option<Arc<SheetMeta>>,
    /// Read-only: declared `(min_col, min_row, max_col, max_row)`; `None` until read.
    declared: Option<(u32, u32, Option<u32>, Option<u32>)>,
    cache: RowCache,
}

/// Print settings and sheet-scoped names, from the workbook's defined names.
#[derive(Default, Clone)]
pub(crate) struct SheetNames {
    pub(crate) print_area: Option<String>,
    pub(crate) print_titles: Option<String>,
    pub(crate) names: Vec<DefinedName>,
}

// ---- Worksheet ---------------------------------------------------------------------

/// A worksheet of a workbook opened with `compatibility="openpyxl"` (openpyxl's
/// `Worksheet`, or `ReadOnlyWorksheet` with `read_only=True`).
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.worksheet")]
pub struct Worksheet {
    wb: Py<Workbook>,
    shared: Arc<Shared>,
    #[pyo3(get)]
    title: String,
    part: Option<String>,
    #[pyo3(get)]
    sheet_state: &'static str,
    names: SheetNames,
    inner: Mutex<SheetInner>,
}

/// openpyxl attributes this profile does not provide, with what to use instead.
const UNSUPPORTED: &[&str] = &[
    "page_setup",
    "page_margins",
    "print_options",
    "HeaderFooter",
    "oddHeader",
    "oddFooter",
    "evenHeader",
    "evenFooter",
    "firstHeader",
    "firstFooter",
    "conditional_formatting",
    "data_validations",
    "views",
    "sheet_view",
    "protection",
    "sheet_format",
    "row_breaks",
    "col_breaks",
    "page_breaks",
    "scenarios",
    "legacy_drawing",
    "print_title_cols_range",
];

impl Worksheet {
    pub(crate) fn new(
        wb: Py<Workbook>,
        shared: Arc<Shared>,
        title: String,
        part: Option<String>,
        sheet_state: &'static str,
        names: SheetNames,
    ) -> Self {
        Self {
            wb,
            shared,
            title,
            part,
            sheet_state,
            names,
            inner: Mutex::new(SheetInner {
                meta: None,
                declared: None,
                cache: RowCache::default(),
            }),
        }
    }

    fn read_only(&self) -> bool {
        self.shared.read_only
    }

    /// Drops the random-access reader (and its file handle) and cached rows.
    pub(crate) fn release(&self) {
        let mut inner = self.lock();
        inner.cache = RowCache::default();
    }

    fn lock(&self) -> std::sync::MutexGuard<'_, SheetInner> {
        self.inner.lock().expect("not poisoned")
    }

    /// Sheet metadata (loaded on first use).
    fn meta(&self, py: Python<'_>) -> PyResult<Arc<SheetMeta>> {
        self.shared.check_open()?;
        if let Some(m) = &self.lock().meta {
            return Ok(Arc::clone(m));
        }
        let normal = !self.read_only();
        let meta =
            py.detach(|| SheetMeta::load(&self.shared, &self.title, self.part.as_deref(), normal))?;
        let meta = Arc::new(meta);
        self.lock().meta = Some(Arc::clone(&meta));
        Ok(meta)
    }

    /// Read-only bounds: `(min_col, min_row, max_col, max_row)` as openpyxl's
    /// `ReadOnlyWorksheet` has them (from `<dimension>`).
    fn declared(&self, py: Python<'_>) -> PyResult<(u32, u32, Option<u32>, Option<u32>)> {
        if let Some(d) = self.lock().declared {
            return Ok(d);
        }
        let d = py.detach(|| -> Result<_, Error> {
            let reader = SheetReader::open(&self.shared, self.part.as_deref())?;
            Ok(match reader.declared_dimension() {
                Some((r0, c0, r1, c1)) => (c0, r0, Some(c1), Some(r1)),
                None => (1, 1, None, None),
            })
        })?;
        self.lock().declared = Some(d);
        Ok(d)
    }

    /// Normal-mode bounds (`None` for a sheet without cells).
    fn bounds(&self, py: Python<'_>) -> PyResult<Option<Bounds>> {
        Ok(self.meta(py)?.bounds)
    }

    fn stored_row(&self, py: Python<'_>, r: u32) -> PyResult<Option<Arc<CRow>>> {
        let shared = &self.shared;
        let part = self.part.as_deref();
        Ok(py.detach(|| self.lock().cache.row(shared, part, r))?)
    }

    fn check_cell_args(row: i64, column: i64) -> PyResult<(u32, u32)> {
        if row < 1 || column < 1 {
            return Err(PyValueError::new_err(
                "Row or column values must be at least 1",
            ));
        }
        if row > MAX_ROW as i64 {
            return Err(PyValueError::new_err(format!(
                "Row numbers must be between 1 and 1048576. Row number supplied was {row}"
            )));
        }
        if column > super::utils::MAX_COLUMN as i64 {
            return Err(PyValueError::new_err(format!(
                "Invalid column index {column}"
            )));
        }
        Ok((row as u32, column as u32))
    }

    /// `ws.cell(row, column)`.
    fn get_cell(slf: &Bound<'_, Self>, r: u32, c: u32) -> PyResult<Py<PyAny>> {
        let py = slf.py();
        let ws = slf.get();
        let row = ws.stored_row(py, r)?;
        let stored = row.as_ref().and_then(|row| row.get(c));
        if ws.read_only() {
            let meta = match ws.shared.opts.any_metadata() {
                true => Some(ws.meta(py)?),
                false => None,
            };
            return ro_cell(py, slf, r, c, stored, meta.as_deref());
        }
        let meta = ws.meta(py)?;
        normal_cell(py, slf, &meta, r, c, stored, false)
    }

    fn rows_iter(
        slf: &Bound<'_, Self>,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        values_only: bool,
    ) -> PyResult<RowIter> {
        let py = slf.py();
        let ws = slf.get();
        ws.shared.check_open()?;
        let (min_row, max_row, min_col, max_col) = (
            or_none(min_row),
            or_none(max_row),
            or_none(min_col),
            or_none(max_col),
        );
        let any = min_row.is_some() || max_row.is_some() || min_col.is_some() || max_col.is_some();
        if ws.read_only() {
            let (_, _, dmax_col, dmax_row) = ws.declared(py)?;
            let min_col = min_col.unwrap_or(1);
            let max_col = max_col.or(dmax_col);
            let max_row = max_row.or(dmax_row);
            let mut it = RowIter::new(
                slf,
                None,
                values_only,
                min_row.unwrap_or(1),
                max_row,
                min_col,
                max_col,
            );
            if !values_only && ws.shared.opts.any_metadata() {
                it.ro_meta = Some(ws.meta(py)?);
            }
            return Ok(it);
        }
        let meta = ws.meta(py)?;
        if meta.bounds.is_none() && !any {
            return Ok(RowIter::empty(slf));
        }
        let (bmax_row, bmax_col) = meta.bounds.map_or((1, 1), |b| (b.2, b.3));
        Ok(RowIter::new(
            slf,
            Some(meta),
            values_only,
            min_row.unwrap_or(1),
            Some(max_row.unwrap_or(bmax_row)),
            min_col.unwrap_or(1),
            Some(max_col.unwrap_or(bmax_col)),
        ))
    }

    fn cols_list<'py>(
        slf: &Bound<'py, Self>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        min_row: Option<u32>,
        max_row: Option<u32>,
        values_only: bool,
    ) -> PyResult<Bound<'py, PyList>> {
        let py = slf.py();
        let ws = slf.get();
        if ws.read_only() {
            return Err(PyAttributeError::new_err(
                "'ReadOnlyWorksheet' object has no attribute 'iter_cols'",
            ));
        }
        let meta = ws.meta(py)?;
        let (min_row, max_row, min_col, max_col) = (
            or_none(min_row),
            or_none(max_row),
            or_none(min_col),
            or_none(max_col),
        );
        let any = min_row.is_some() || max_row.is_some() || min_col.is_some() || max_col.is_some();
        let out = PyList::empty(py);
        if meta.bounds.is_none() && !any {
            return Ok(out);
        }
        let (bmax_row, bmax_col) = meta.bounds.map_or((1, 1), |b| (b.2, b.3));
        let (r0, r1) = (min_row.unwrap_or(1), max_row.unwrap_or(bmax_row));
        let (c0, c1) = (min_col.unwrap_or(1), max_col.unwrap_or(bmax_col));
        if c0 > c1 {
            return Ok(out);
        }
        let mut columns: Vec<Vec<Py<PyAny>>> = (c0..=c1).map(|_| Vec::new()).collect();
        let mut rows = RowIter::new(slf, Some(meta), values_only, r0, Some(r1), c0, Some(c1));
        while let Some(row) = rows.next_items(py)? {
            for (i, item) in row.into_iter().enumerate() {
                columns[i].push(item);
            }
        }
        for col in columns {
            out.append(PyTuple::new(py, col)?)?;
        }
        Ok(out)
    }

    fn print_titles_parts(&self) -> (Option<String>, Option<String>) {
        let Some(v) = &self.names.print_titles else {
            return (None, None);
        };
        let (mut rows, mut cols) = (None, None);
        for (_, part) in sheet_ranges(v) {
            let plain = part.replace('$', "");
            let Some((a, b)) = plain.split_once(':') else {
                continue;
            };
            if a.chars().all(|c| c.is_ascii_digit())
                && b.chars().all(|c| c.is_ascii_digit())
                && !a.is_empty()
                && !b.is_empty()
            {
                rows = Some(format!("${a}:${b}"));
            } else if (1..=3).contains(&a.len())
                && (1..=3).contains(&b.len())
                && a.chars().all(|c| c.is_ascii_alphabetic())
                && b.chars().all(|c| c.is_ascii_alphabetic())
            {
                cols = Some(format!("${a}:${b}"));
            }
        }
        (rows, cols)
    }
}

/// Cells and values of one row of a normal-mode worksheet.
#[allow(clippy::too_many_arguments)]
fn normal_row_items(
    py: Python<'_>,
    ws: &Bound<'_, Worksheet>,
    meta: &SheetMeta,
    r: u32,
    c0: u32,
    c1: u32,
    row: Option<&CRow>,
    values_only: bool,
) -> PyResult<Vec<Py<PyAny>>> {
    let mut out = Vec::with_capacity((c1 + 1).saturating_sub(c0) as usize);
    for c in c0..=c1 {
        let stored = row.and_then(|row| row.get(c));
        out.push(normal_cell(py, ws, meta, r, c, stored, values_only)?);
    }
    Ok(out)
}

/// A normal-mode cell (or its value): merged placeholders, hyperlinks and comments
/// applied like openpyxl's reader.
fn normal_cell(
    py: Python<'_>,
    ws: &Bound<'_, Worksheet>,
    meta: &SheetMeta,
    r: u32,
    c: u32,
    stored: Option<&CCell>,
    values_only: bool,
) -> PyResult<Py<PyAny>> {
    let merge = meta.merge_at(r, c);
    if let Some(m) = merge {
        let b = meta.merged[m];
        if (b.0, b.1) != (r, c) {
            if values_only {
                return Ok(py.None());
            }
            let shared = &ws.get().shared;
            let border = meta.merged_border(m, r, c, &shared.style(ZERO_STYLE).border);
            return Ok(Py::new(
                py,
                MergedCell {
                    ws: ws.clone().unbind(),
                    row: r,
                    column: c,
                    border,
                    protection: meta.merged_fmt[m].protection.clone(),
                },
            )?
            .into_any());
        }
    }
    let (mut data_type, style) = stored.map_or(("n", ZERO_STYLE), |s| (s.data_type, s.style));
    let link = meta.link_at(r, c);
    // openpyxl's hyperlink setter fills an empty cell with the link.
    let filled = match link {
        Some(l) if stored.is_none_or(|s| s.value.is_none()) => {
            let text = l.target.clone().or_else(|| l.location.clone());
            data_type = match text.as_deref() {
                None => "n",
                Some(t) if t.len() > 1 && t.starts_with('=') => "f",
                Some(t) if ERROR_CODES.contains(&t) => "e",
                Some(_) => "s",
            };
            Some(text.map_or(CVal::None, CVal::Str))
        }
        _ => None,
    };
    // Converted straight from the stored cell (no copy of the value).
    let value = match (&filled, stored) {
        (Some(v), _) => v.to_py(py)?,
        (None, Some(s)) => s.value.to_py(py)?,
        (None, None) => py.None(),
    };
    if values_only {
        return Ok(value);
    }
    let hyperlink = link.map(|l| l.to_py(py, coordinate(r, c))).transpose()?;
    let comment = meta
        .comments
        .get(&(r, c))
        .map(|(text, author)| Py::new(py, Comment::new(text.clone(), author.clone())))
        .transpose()?;
    let border = merge.map(|m| meta.merged_border(m, r, c, &meta.merged_fmt[m].anchor_border));
    Ok(Py::new(
        py,
        Cell {
            ws: ws.clone().unbind(),
            row: r,
            column: c,
            value,
            data_type,
            style_id: style,
            border,
            comment,
            hyperlink,
            extra: stored.and_then(|s| s.extra.clone()),
        },
    )?
    .into_any())
}

const ERROR_CODES: &[&str] = &[
    "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
];

/// A read-only cell, or `EmptyCell` where the file stores none. With `read_comments`
/// / `read_hyperlinks` / `read_merged_cells` (`meta`), cells get those like the native
/// rich stream, and positions with one of them but no stored cell get a cell too.
fn ro_cell(
    py: Python<'_>,
    ws: &Bound<'_, Worksheet>,
    r: u32,
    c: u32,
    stored: Option<&CCell>,
    meta: Option<&SheetMeta>,
) -> PyResult<Py<PyAny>> {
    let opts = ws.get().shared.opts;
    let (mut comment, mut hyperlink, mut merged_range) = (None, None, None);
    if let Some(m) = meta {
        if opts.comments {
            comment = m
                .comments
                .get(&(r, c))
                .map(|(text, author)| Py::new(py, Comment::new(text.clone(), author.clone())))
                .transpose()?;
        }
        if opts.hyperlinks {
            hyperlink = m
                .link_at(r, c)
                .map(|l| l.to_py(py, coordinate(r, c)))
                .transpose()?;
        }
        if opts.merged {
            merged_range = m
                .merge_at(r, c)
                .map(|i| {
                    Py::new(
                        py,
                        MergedCellRange {
                            bounds: m.merged[i],
                            ws: Some(ws.clone().into_any().unbind()),
                        },
                    )
                })
                .transpose()?;
        }
    }
    if stored.is_none() && comment.is_none() && hyperlink.is_none() && merged_range.is_none() {
        return empty_cell(py);
    }
    Ok(Py::new(
        py,
        ReadOnlyCell {
            ws: ws.clone().unbind(),
            row: r,
            column: c,
            value: stored.map_or_else(|| Ok(py.None()), |s| s.value.to_py(py))?,
            data_type: stored.map_or("n", |s| s.data_type),
            style_id: stored.map_or(ZERO_STYLE, |s| s.style),
            extra: stored.and_then(|s| s.extra.clone()),
            comment,
            hyperlink,
            merged_range,
        },
    )?
    .into_any())
}

fn flag_error(cls: &str, attr: &str, flag: &str) -> PyErr {
    PyAttributeError::new_err(format!(
        "'{cls}' object has no attribute '{attr}' (load the workbook with {flag}=True)"
    ))
}

/// `cell.formula` / `cell.cached_value` with `formula_and_value=True`.
fn extra_value(
    py: Python<'_>,
    ws: &Py<Worksheet>,
    extra: &Option<Box<CExtra>>,
    cls: &str,
    attr: &str,
    formula: bool,
) -> PyResult<Py<PyAny>> {
    if !ws.bind(py).get().shared.opts.formula_and_value {
        return Err(flag_error(cls, attr, "formula_and_value"));
    }
    match extra {
        Some(e) if formula => e.formula.to_py(py),
        Some(e) => e.cached.to_py(py),
        None => Ok(py.None()),
    }
}

static EMPTY: PyOnceLock<Py<EmptyCell>> = PyOnceLock::new();

fn empty_cell(py: Python<'_>) -> PyResult<Py<PyAny>> {
    Ok(EMPTY
        .get_or_try_init(py, || Py::new(py, EmptyCell {}))?
        .clone_ref(py)
        .into_any())
}

#[pymethods]
impl Worksheet {
    fn __traverse__(&self, visit: PyVisit<'_>) -> Result<(), PyTraverseError> {
        visit.call(&self.wb)
    }

    fn __repr__(&self) -> String {
        if self.read_only() {
            format!("<ReadOnlyWorksheet \"{}\">", self.title)
        } else {
            format!("<Worksheet \"{}\">", self.title)
        }
    }

    #[getter]
    fn parent(&self, py: Python<'_>) -> Py<Workbook> {
        self.wb.clone_ref(py)
    }

    /// Diagnostics: the sheet's metadata (merged cells, hyperlinks, ...) has been read.
    #[getter]
    fn _metadata_loaded(&self) -> bool {
        self.lock().meta.is_some()
    }

    #[getter]
    fn min_row(&self, py: Python<'_>) -> PyResult<u32> {
        if self.read_only() {
            return Ok(self.declared(py)?.1);
        }
        Ok(self.bounds(py)?.map_or(1, |b| b.0))
    }

    #[getter]
    fn min_column(&self, py: Python<'_>) -> PyResult<u32> {
        if self.read_only() {
            return Ok(self.declared(py)?.0);
        }
        Ok(self.bounds(py)?.map_or(1, |b| b.1))
    }

    #[getter]
    fn max_row(&self, py: Python<'_>) -> PyResult<Option<u32>> {
        if self.read_only() {
            return Ok(self.declared(py)?.3);
        }
        Ok(Some(self.bounds(py)?.map_or(1, |b| b.2)))
    }

    #[getter]
    fn max_column(&self, py: Python<'_>) -> PyResult<Option<u32>> {
        if self.read_only() {
            return Ok(self.declared(py)?.2);
        }
        Ok(Some(self.bounds(py)?.map_or(1, |b| b.3)))
    }

    /// Bounding range of the cells (normal mode) or of the declared dimension
    /// (read-only mode, where `force=True` scans the sheet when none is declared).
    #[pyo3(signature = (force=false))]
    fn calculate_dimension(slf: &Bound<'_, Self>, force: bool) -> PyResult<String> {
        let py = slf.py();
        let ws = slf.get();
        if !ws.read_only() {
            return Ok(match ws.bounds(py)? {
                Some(b) => format!("{}:{}", coordinate(b.0, b.1), coordinate(b.2, b.3)),
                None => "A1:A1".into(),
            });
        }
        let (min_col, min_row, mut max_col, mut max_row) = ws.declared(py)?;
        if max_col.is_none() || max_row.is_none() {
            if !force {
                return Err(PyValueError::new_err(
                    "Worksheet is unsized, use calculate_dimension(force=True)",
                ));
            }
            // openpyxl's `_calculate_dimension`: the last stored cell of each row.
            let (mut last_col, mut last_row) = (0u32, None);
            let shared = &ws.shared;
            let part = ws.part.as_deref();
            py.detach(|| -> Result<(), Error> {
                let mut reader = SheetReader::open(shared, part)?;
                while let Some(row) = reader.next_row()? {
                    last_col = last_col.max(row.last_col_in_file);
                    last_row = Some(row.index);
                }
                Ok(())
            })?;
            max_col = Some(last_col);
            max_row = last_row;
            ws.lock().declared = Some((min_col, min_row, max_col, max_row));
        }
        Ok(format!(
            "{}{}:{}{}",
            column_letter(min_col),
            min_row,
            column_letter(max_col.unwrap_or(0).max(1)),
            max_row.unwrap_or(0)
        ))
    }

    /// Forget the declared dimension (read-only mode), like openpyxl's `reset_dimensions`.
    fn reset_dimensions(&self, py: Python<'_>) -> PyResult<()> {
        if !self.read_only() {
            return Err(PyAttributeError::new_err(
                "'Worksheet' object has no attribute 'reset_dimensions'",
            ));
        }
        let (c0, r0, _, _) = self.declared(py)?;
        self.lock().declared = Some((c0, r0, None, None));
        Ok(())
    }

    #[getter]
    fn dimensions(slf: &Bound<'_, Self>) -> PyResult<String> {
        Self::calculate_dimension(slf, false)
    }

    #[pyo3(signature = (min_row=None, max_row=None, min_col=None, max_col=None, values_only=false))]
    fn iter_rows(
        slf: &Bound<'_, Self>,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        values_only: bool,
    ) -> PyResult<RowIter> {
        Self::rows_iter(slf, min_row, max_row, min_col, max_col, values_only)
    }

    #[getter]
    fn rows(slf: &Bound<'_, Self>) -> PyResult<RowIter> {
        Self::rows_iter(slf, None, None, None, None, false)
    }

    #[getter]
    fn values(slf: &Bound<'_, Self>) -> PyResult<RowIter> {
        Self::rows_iter(slf, None, None, None, None, true)
    }

    fn __iter__(slf: &Bound<'_, Self>) -> PyResult<RowIter> {
        Self::rows_iter(slf, None, None, None, None, false)
    }

    /// Columns of cells (normal mode only). The requested block is read in one pass.
    #[pyo3(signature = (min_col=None, max_col=None, min_row=None, max_row=None, values_only=false))]
    fn iter_cols<'py>(
        slf: &Bound<'py, Self>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        min_row: Option<u32>,
        max_row: Option<u32>,
        values_only: bool,
    ) -> PyResult<Bound<'py, PyIterator>> {
        Self::cols_list(slf, min_col, max_col, min_row, max_row, values_only)?.try_iter()
    }

    #[getter]
    fn columns<'py>(slf: &Bound<'py, Self>) -> PyResult<Bound<'py, PyIterator>> {
        Self::iter_cols(slf, None, None, None, None, false)
    }

    /// The cell at a 1-based row and column. Cells are read-only: passing `value` raises.
    #[pyo3(signature = (row, column, value=None))]
    fn cell(
        slf: &Bound<'_, Self>,
        row: i64,
        column: i64,
        value: Option<Bound<'_, PyAny>>,
    ) -> PyResult<Py<PyAny>> {
        if value.is_some_and(|v| !v.is_none()) {
            return Err(PyAttributeError::new_err("Cell is read only"));
        }
        if row < 1 || column < 1 {
            return Err(PyValueError::new_err(
                "Row or column values must be at least 1",
            ));
        }
        let (r, c) = Self::check_cell_args(row, column)?;
        Self::get_cell(slf, r, c)
    }

    fn __getitem__(slf: &Bound<'_, Self>, key: &Bound<'_, PyAny>) -> PyResult<Py<PyAny>> {
        let py = slf.py();
        let key: String = if let Ok(s) = key.cast::<PySlice>() {
            let (start, stop) = (s.getattr("start")?, s.getattr("stop")?);
            if !start.is_truthy()? || !stop.is_truthy()? {
                return Err(PyIndexError::new_err(format!(
                    "{} is not a valid coordinate or range",
                    key.repr()?
                )));
            }
            format!("{}:{}", start.str()?, stop.str()?)
        } else if let Ok(i) = key.extract::<i64>() {
            i.to_string()
        } else {
            key.extract()?
        };
        let (min_col, min_row, max_col, max_row) = range_boundaries(&key)?;
        if min_col.is_none() && min_row.is_none() && max_col.is_none() && max_row.is_none() {
            return Err(PyIndexError::new_err(format!(
                "{key} is not a valid coordinate or range"
            )));
        }
        if min_row.is_none() {
            let cols = Self::cols_list(slf, min_col, max_col, None, None, false)?;
            if min_col == max_col {
                return Ok(cols.get_item(0)?.unbind());
            }
            return Ok(PyTuple::new(py, cols)?.into_any().unbind());
        }
        if min_col.is_none() {
            let max_column = slf.get().max_column(py)?;
            let rows: Vec<Py<PyAny>> =
                Self::rows_iter(slf, min_row, max_row, None, max_column, false)?.collect(py)?;
            if min_row == max_row {
                return rows
                    .into_iter()
                    .next()
                    .ok_or_else(|| PyIndexError::new_err("tuple index out of range"));
            }
            return Ok(PyTuple::new(py, rows)?.into_any().unbind());
        }
        if !key.contains(':') {
            let (r, c) = (min_row.expect("set"), min_col.expect("set"));
            let (r, c) = Self::check_cell_args(r as i64, c as i64)?;
            return Self::get_cell(slf, r, c);
        }
        let rows: Vec<Py<PyAny>> =
            Self::rows_iter(slf, min_row, max_row, min_col, max_col, false)?.collect(py)?;
        Ok(PyTuple::new(py, rows)?.into_any().unbind())
    }

    /// Merged ranges (`MultiCellRange`).
    #[getter]
    fn merged_cells(slf: &Bound<'_, Self>) -> PyResult<MultiCellRange> {
        let py = slf.py();
        let meta = slf.get().meta(py)?;
        let ranges = meta
            .merged
            .iter()
            .map(|b| {
                Py::new(
                    py,
                    MergedCellRange {
                        bounds: *b,
                        ws: Some(slf.clone().into_any().unbind()),
                    },
                )
            })
            .collect::<PyResult<_>>()?;
        Ok(MultiCellRange { ranges })
    }

    /// `"B2"` (the pane's top-left cell) or `None`.
    #[getter]
    fn freeze_panes(&self, py: Python<'_>) -> PyResult<Option<String>> {
        let meta = self.meta(py)?;
        Ok(meta
            .info
            .views
            .first()
            .and_then(|v| v.pane.as_ref())
            .and_then(|p| p.iter().find(|(k, _)| k == "topLeftCell"))
            .map(|(_, v)| v.clone()))
    }

    /// Print area like openpyxl: `"'Sheet'!$A$1:$C$10"` (comma-separated ranges), or
    /// `""` when none (or when the defined name is not a list of static ranges).
    #[getter]
    fn print_area(&self) -> PyResult<String> {
        let Some(v) = &self.names.print_area else {
            return Ok(String::new());
        };
        let Ok(ranges) = super::utils::split_print_areas(v) else {
            return Ok(String::new());
        };
        let mut parsed: Vec<(u32, u32, u32, u32)> = Vec::new();
        for r in &ranges {
            match range_boundaries(r)? {
                (Some(c0), Some(r0), Some(c1), Some(r1)) => parsed.push((c0, r0, c1, r1)),
                _ => return Ok(String::new()),
            }
        }
        parsed.sort();
        let title = quote_sheetname(&self.title);
        Ok(parsed
            .into_iter()
            .map(|(c0, r0, c1, r1)| {
                let a = format!("${}${}", column_letter(c0), r0);
                if (c0, r0) == (c1, r1) {
                    format!("{title}!{a}")
                } else {
                    format!("{title}!{a}:${}${}", column_letter(c1), r1)
                }
            })
            .collect::<Vec<_>>()
            .join(","))
    }

    #[getter]
    fn print_title_rows(&self) -> Option<String> {
        self.print_titles_parts().0
    }

    #[getter]
    fn print_title_cols(&self) -> Option<String> {
        self.print_titles_parts().1
    }

    #[getter]
    fn print_titles(&self) -> String {
        let (rows, cols) = self.print_titles_parts();
        let title = quote_sheetname(&self.title);
        [rows, cols]
            .into_iter()
            .flatten()
            .map(|v| format!("{title}!{v}"))
            .collect::<Vec<_>>()
            .join(",")
    }

    /// Sheet-scoped defined names (print settings excluded), by name.
    #[getter]
    fn defined_names<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        let d = PyDict::new(py);
        for n in &self.names.names {
            d.set_item(&n.name, Py::new(py, n.clone())?)?;
        }
        Ok(d)
    }

    #[getter]
    fn row_dimensions(&self, py: Python<'_>) -> PyResult<DimensionHolder> {
        let meta = self.meta(py)?;
        let mut keys = Vec::new();
        let mut values = Vec::new();
        for r in &meta.info.rows {
            let index = r.index + 1;
            keys.push(index.into_pyobject(py)?.into_any().unbind());
            values.push(
                Py::new(
                    py,
                    RowDimension {
                        index,
                        ht: r.height,
                        hidden: r.hidden,
                        outline_level: r.outline_level as u32,
                        collapsed: r.collapsed,
                        style: r.style,
                    },
                )?
                .into_any(),
            );
        }
        Ok(DimensionHolder {
            keys,
            values,
            rows: true,
        })
    }

    #[getter]
    fn column_dimensions(&self, py: Python<'_>) -> PyResult<DimensionHolder> {
        let meta = self.meta(py)?;
        let mut entries: Vec<(String, ColumnDimension)> = Vec::new();
        for c in &meta.info.columns {
            let (min, max) = (c.min + 1, c.max + 1);
            let index = column_letter(min);
            let dim = ColumnDimension {
                index: index.clone(),
                width: c.width.unwrap_or(13.0),
                best_fit: c.best_fit,
                hidden: c.hidden,
                outline_level: c.outline_level as u32,
                collapsed: c.collapsed,
                min: Some(min),
                max: Some(max),
            };
            match entries.iter_mut().find(|(k, _)| *k == index) {
                Some(e) => e.1 = dim,
                None => entries.push((index, dim)),
            }
        }
        let mut keys = Vec::new();
        let mut values = Vec::new();
        for (k, v) in entries {
            keys.push(k.into_pyobject(py)?.into_any().unbind());
            values.push(Py::new(py, v)?.into_any());
        }
        Ok(DimensionHolder {
            keys,
            values,
            rows: false,
        })
    }

    /// `tabColor` and `codeName`.
    #[getter]
    fn sheet_properties(&self, py: Python<'_>) -> PyResult<WorksheetProperties> {
        let meta = self.meta(py)?;
        Ok(WorksheetProperties {
            tab_color: meta.info.tab_color.as_ref().map(Color::from_style),
            code_name: meta
                .info
                .sheet_properties
                .iter()
                .find(|(k, _)| k == "codeName")
                .map(|(_, v)| v.clone()),
        })
    }

    /// The sheet's auto filter (only `ref`).
    #[getter]
    fn auto_filter(&self, py: Python<'_>) -> PyResult<AutoFilter> {
        Ok(AutoFilter {
            reference: self.meta(py)?.info.auto_filter.clone(),
        })
    }

    /// Tables by name.
    #[getter]
    fn tables(&self, py: Python<'_>) -> PyResult<TableList> {
        let meta = self.meta(py)?;
        let mut tables = Vec::new();
        for t in &meta.info.tables {
            let get = |k: &str| t.attrs.iter().find(|(a, _)| a == k).map(|(_, v)| v.clone());
            let int = |k: &str| get(k).and_then(|v| v.parse::<i64>().ok());
            let flag = |a: &[(String, String)], k: &str| {
                a.iter()
                    .find(|(x, _)| x == k)
                    .map(|(_, v)| v == "1" || v == "true")
            };
            let style = t
                .style
                .as_ref()
                .map(|s| {
                    Py::new(
                        py,
                        TableStyleInfo {
                            name: s.iter().find(|(k, _)| k == "name").map(|(_, v)| v.clone()),
                            show_first_column: flag(s, "showFirstColumn"),
                            show_last_column: flag(s, "showLastColumn"),
                            show_row_stripes: flag(s, "showRowStripes"),
                            show_column_stripes: flag(s, "showColumnStripes"),
                        },
                    )
                })
                .transpose()?;
            let columns = t
                .columns
                .iter()
                .map(|c| {
                    let get =
                        |k: &str| c.attrs.iter().find(|(a, _)| a == k).map(|(_, v)| v.clone());
                    Py::new(
                        py,
                        TableColumn {
                            id: get("id").and_then(|v| v.parse().ok()),
                            name: get("name"),
                            totals_row_function: get("totalsRowFunction"),
                            totals_row_label: get("totalsRowLabel"),
                        },
                    )
                })
                .collect::<PyResult<_>>()?;
            let auto_filter = t
                .auto_filter
                .as_ref()
                .map(|a| {
                    Py::new(
                        py,
                        AutoFilter {
                            reference: a.range.clone(),
                        },
                    )
                })
                .transpose()?;
            let name = get("name")
                .or_else(|| get("displayName"))
                .unwrap_or_default();
            let table = Table {
                id: int("id"),
                name: get("name"),
                display_name: get("displayName"),
                reference: get("ref"),
                header_row_count: int("headerRowCount").unwrap_or(1),
                totals_row_count: int("totalsRowCount"),
                totals_row_shown: flag(&t.attrs, "totalsRowShown"),
                table_type: get("tableType"),
                comment: get("comment"),
                style,
                columns,
                auto_filter,
            };
            tables.push((name, Py::new(py, table)?));
        }
        Ok(TableList { tables })
    }

    fn __getattr__(&self, name: &str) -> PyResult<Py<PyAny>> {
        if UNSUPPORTED.contains(&name) {
            return Err(PyAttributeError::new_err(format!(
                "Worksheet.{name} is not provided with compatibility='openpyxl'; read it with \
                 CalamineWorkbook.stream_sheet_by_name(title, rich=True) (sheet_settings, \
                 conditional_formats, data_validations, row_breaks, ...)"
            )));
        }
        let cls = if self.read_only() {
            "ReadOnlyWorksheet"
        } else {
            "Worksheet"
        };
        Err(PyAttributeError::new_err(format!(
            "'{cls}' object has no attribute '{name}'"
        )))
    }
}

/// A chartsheet (no cells).
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.chartsheet")]
pub struct Chartsheet {
    pub(crate) wb: Py<Workbook>,
    #[pyo3(get)]
    pub(crate) title: String,
    #[pyo3(get)]
    pub(crate) sheet_state: &'static str,
}

#[pymethods]
impl Chartsheet {
    fn __traverse__(&self, visit: PyVisit<'_>) -> Result<(), PyTraverseError> {
        visit.call(&self.wb)
    }
    #[getter]
    fn parent(&self, py: Python<'_>) -> Py<Workbook> {
        self.wb.clone_ref(py)
    }
    fn __repr__(&self) -> String {
        format!("<Chartsheet \"{}\">", self.title)
    }
}

// ---- row iteration -----------------------------------------------------------------

/// Output of the read-only row algorithm.
enum RoOut {
    /// A row without stored cells (its index).
    Empty(u32),
    Row(CRow),
}

/// Iterator returned by `iter_rows` / `rows` / `values` / `iter(ws)`, with its own
/// reader: iterators never invalidate each other.
#[pyclass(
    name = "RowIterator",
    module = "python_calamine_plus.openpyxl.worksheet"
)]
pub struct RowIter {
    ws: Py<Worksheet>,
    /// `None` in read-only mode.
    meta: Option<Arc<SheetMeta>>,
    /// Read-only mode with `read_comments` / `read_hyperlinks` / `read_merged_cells`.
    ro_meta: Option<Arc<SheetMeta>>,
    reader: Option<SheetReader>,
    values_only: bool,
    max_row: Option<u32>,
    min_col: u32,
    max_col: Option<u32>,
    done: bool,
    // Normal mode.
    next_row: u32,
    lookahead: Option<CRow>,
    reader_eof: bool,
    // Read-only mode (openpyxl's `_cells_by_row`).
    counter: u32,
    idx: u32,
    queue: VecDeque<RoOut>,
    trailing_checked: bool,
}

impl RowIter {
    #[allow(clippy::too_many_arguments)]
    fn new(
        ws: &Bound<'_, Worksheet>,
        meta: Option<Arc<SheetMeta>>,
        values_only: bool,
        min_row: u32,
        max_row: Option<u32>,
        min_col: u32,
        max_col: Option<u32>,
    ) -> Self {
        Self {
            ws: ws.clone().unbind(),
            meta,
            ro_meta: None,
            reader: None,
            values_only,
            max_row,
            min_col,
            max_col,
            done: false,
            next_row: min_row,
            lookahead: None,
            reader_eof: false,
            counter: min_row,
            idx: 1,
            queue: VecDeque::new(),
            trailing_checked: false,
        }
    }

    fn empty(ws: &Bound<'_, Worksheet>) -> Self {
        let mut it = Self::new(ws, None, false, 1, None, 1, None);
        it.done = true;
        it
    }

    fn reader(&mut self, py: Python<'_>) -> PyResult<&mut SheetReader> {
        if self.reader.is_none() {
            let ws = self.ws.bind(py).get();
            let r = py.detach(|| SheetReader::open(&ws.shared, ws.part.as_deref()))?;
            self.reader = Some(r);
        }
        Ok(self.reader.as_mut().expect("opened"))
    }

    fn read_row(&mut self, py: Python<'_>) -> PyResult<Option<CRow>> {
        let reader = self.reader(py)?;
        Ok(py.detach(|| reader.next_row())?)
    }

    fn collect(mut self, py: Python<'_>) -> PyResult<Vec<Py<PyAny>>> {
        let mut out = Vec::new();
        while let Some(row) = self.next_items(py)? {
            out.push(PyTuple::new(py, row)?.into_any().unbind());
        }
        Ok(out)
    }

    /// The next row as a list of cells (or values).
    fn next_items(&mut self, py: Python<'_>) -> PyResult<Option<Vec<Py<PyAny>>>> {
        if self.done {
            return Ok(None);
        }
        let ws = self.ws.bind(py).clone();
        if let Err(e) = ws.get().shared.check_open() {
            self.done = true;
            self.reader = None;
            return Err(e.into());
        }
        let out = if self.meta.is_some() {
            self.next_normal(py, &ws)
        } else {
            self.next_read_only(py, &ws)
        };
        if !matches!(out, Ok(Some(_))) {
            self.done = true;
            self.reader = None;
        }
        out
    }

    fn next_normal(
        &mut self,
        py: Python<'_>,
        ws: &Bound<'_, Worksheet>,
    ) -> PyResult<Option<Vec<Py<PyAny>>>> {
        let r = self.next_row;
        let (max_row, max_col) = (self.max_row.expect("normal"), self.max_col.expect("normal"));
        if r > max_row {
            return Ok(None);
        }
        self.next_row += 1;
        while !self.reader_eof && self.lookahead.as_ref().is_none_or(|l| l.index < r) {
            match self.read_row(py)? {
                Some(row) => self.lookahead = Some(row),
                None => {
                    self.reader_eof = true;
                    self.lookahead = None;
                }
            }
        }
        let row = self.lookahead.as_ref().filter(|l| l.index == r);
        let meta = Arc::clone(self.meta.as_ref().expect("normal"));
        normal_row_items(
            py,
            ws,
            &meta,
            r,
            self.min_col,
            max_col,
            row,
            self.values_only,
        )
        .map(Some)
    }

    fn next_read_only(
        &mut self,
        py: Python<'_>,
        ws: &Bound<'_, Worksheet>,
    ) -> PyResult<Option<Vec<Py<PyAny>>>> {
        while self.queue.is_empty() {
            if self.trailing_checked {
                return Ok(None);
            }
            // The next `<row>` (cells, or the trailing row without cells).
            let item = match self.read_row(py)? {
                Some(row) => Some((row.index, Some(row))),
                None => {
                    self.trailing_checked = true;
                    let last = self.reader.as_ref().and_then(|r| r.last_row_element());
                    last.filter(|l| *l > self.idx_seen()).map(|l| (l, None))
                }
            };
            let Some((idx, row)) = item else {
                // End of the sheet: pad to `max_row` only if it stopped the loop.
                if let Some(max_row) = self.max_row {
                    if max_row < self.idx {
                        for i in self.counter..=max_row {
                            self.queue.push_back(RoOut::Empty(i));
                        }
                        self.counter = max_row + 1;
                    }
                }
                break;
            };
            self.idx = idx;
            if self.max_row.is_some_and(|m| idx > m) {
                let max_row = self.max_row.expect("checked");
                for i in self.counter..=max_row {
                    self.queue.push_back(RoOut::Empty(i));
                }
                self.trailing_checked = true;
                break;
            }
            for i in self.counter..idx {
                self.queue.push_back(RoOut::Empty(i));
            }
            self.counter = self.counter.max(idx);
            if self.counter <= idx {
                self.queue.push_back(match row {
                    Some(r) => RoOut::Row(r),
                    None => RoOut::Empty(idx),
                });
                self.counter += 1;
            }
        }
        let Some(out) = self.queue.pop_front() else {
            return Ok(None);
        };
        let min_col = self.min_col;
        let (index, row, max_col) = match out {
            RoOut::Empty(i) => match self.max_col {
                Some(max_col) => (i, None, max_col),
                None => return Ok(Some(Vec::new())),
            },
            RoOut::Row(row) => {
                let max_col = self.max_col.unwrap_or(row.last_col_in_file);
                (row.index, Some(row), max_col)
            }
        };
        let width = (max_col + 1).saturating_sub(min_col);
        let meta = self.ro_meta.as_deref();
        let mut items = Vec::with_capacity(width as usize);
        for c in min_col..min_col + width {
            let stored = row.as_ref().and_then(|r| r.get(c));
            items.push(if self.values_only {
                stored.map_or_else(|| Ok(py.None()), |cell| cell.value.to_py(py))?
            } else {
                ro_cell(py, ws, index, c, stored, meta)?
            });
        }
        Ok(Some(items))
    }

    /// Index of the last `<row>` that produced output or was skipped.
    fn idx_seen(&self) -> u32 {
        self.idx
    }
}

#[pymethods]
impl RowIter {
    fn __traverse__(&self, visit: PyVisit<'_>) -> Result<(), PyTraverseError> {
        visit.call(&self.ws)
    }

    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__<'py>(&mut self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyTuple>>> {
        match self.next_items(py)? {
            Some(items) => Ok(Some(PyTuple::new(py, items)?)),
            None => Ok(None),
        }
    }
}

// ---- cells -------------------------------------------------------------------------

fn style_of(ws: &Worksheet, id: u32) -> &super::styles::CompatStyle {
    ws.shared.style(id)
}

/// A cell of a normal-mode worksheet, like openpyxl's `Cell`.
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.cell")]
pub struct Cell {
    ws: Py<Worksheet>,
    #[pyo3(get)]
    row: u32,
    #[pyo3(get)]
    column: u32,
    #[pyo3(get)]
    value: Py<PyAny>,
    #[pyo3(get)]
    data_type: &'static str,
    style_id: u32,
    /// Border of a merged range's top-left cell (after openpyxl's merge formatting).
    border: Option<Border>,
    #[pyo3(get)]
    comment: Option<Py<Comment>>,
    #[pyo3(get)]
    hyperlink: Option<Py<Hyperlink>>,
    extra: Option<Box<CExtra>>,
}

/// One `#[pymethods]` block per cell class: shared position / style getters plus
/// the class's own methods.
macro_rules! style_getters {
    ($ty:ident { $($body:tt)* }) => {
        #[pymethods]
        impl $ty {
            $($body)*

            #[getter]
            fn coordinate(&self) -> String {
                coordinate(self.row, self.column)
            }
            #[getter]
            fn column_letter(&self) -> String {
                column_letter(self.column)
            }
            #[getter]
            fn col_idx(&self) -> u32 {
                self.column
            }
            #[getter]
            fn parent(&self, py: Python<'_>) -> Py<Worksheet> {
                self.ws.clone_ref(py)
            }
            #[getter]
            fn font(&self, py: Python<'_>) -> PyResult<Py<super::styles::Font>> {
                let ws = self.ws.bind(py).get();
                Ok(ws.shared.style_py(py, self.style_index())?.font.clone_ref(py))
            }
            #[getter]
            fn fill(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
                let ws = self.ws.bind(py).get();
                Ok(ws.shared.style_py(py, self.style_index())?.fill.clone_ref(py))
            }
            #[getter]
            fn alignment(&self, py: Python<'_>) -> PyResult<Py<super::styles::Alignment>> {
                let ws = self.ws.bind(py).get();
                Ok(ws.shared.style_py(py, self.style_index())?.alignment.clone_ref(py))
            }
            #[getter]
            fn number_format(&self, py: Python<'_>) -> PyResult<Py<pyo3::types::PyString>> {
                let ws = self.ws.bind(py).get();
                Ok(ws.shared.style_py(py, self.style_index())?.number_format.clone_ref(py))
            }
            #[getter]
            fn style(&self, py: Python<'_>) -> PyResult<Py<pyo3::types::PyString>> {
                let ws = self.ws.bind(py).get();
                Ok(ws.shared.style_py(py, self.style_index())?.named_style.clone_ref(py))
            }
            #[getter(quotePrefix)]
            fn quote_prefix(&self, py: Python<'_>) -> bool {
                style_of(self.ws.bind(py).get(), self.style_index()).quote_prefix
            }
            #[getter(pivotButton)]
            fn pivot_button(&self, py: Python<'_>) -> bool {
                style_of(self.ws.bind(py).get(), self.style_index()).pivot_button
            }
            fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
                let title = pyo3::types::PyString::new(py, &self.ws.bind(py).get().title);
                Ok(format!(
                    "<{} {}.{}>",
                    stringify!($ty),
                    title.repr()?,
                    coordinate(self.row, self.column)
                ))
            }
            fn __traverse__(&self, visit: PyVisit<'_>) -> Result<(), PyTraverseError> {
                visit.call(&self.ws)
            }
        }
    };
}

impl Cell {
    fn style_index(&self) -> u32 {
        self.style_id
    }
}
style_getters!(Cell {
    /// The formula (`formula_and_value=True`), `None` without one.
    #[getter]
    fn formula(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        extra_value(py, &self.ws, &self.extra, "Cell", "formula", true)
    }
    /// The result saved in the file (`formula_and_value=True`).
    #[getter]
    fn cached_value(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        extra_value(py, &self.ws, &self.extra, "Cell", "cached_value", false)
    }
    /// Workbook-local format id (0 for cells the file does not store).
    #[getter(style_id)]
    fn py_style_id(&self) -> u32 {
        if self.style_id == ZERO_STYLE { 0 } else { self.style_id }
    }
    #[getter]
    fn internal_value(&self, py: Python<'_>) -> Py<PyAny> {
        self.value.clone_ref(py)
    }
    #[getter]
    fn has_style(&self, py: Python<'_>) -> bool {
        style_of(self.ws.bind(py).get(), self.style_id).has_style
    }
    #[getter]
    fn border(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        if let Some(b) = &self.border {
            return Ok(Py::new(py, b.clone())?.into_any());
        }
        let ws = self.ws.bind(py).get();
        Ok(ws.shared.style_py(py, self.style_id)?.border.clone_ref(py).into_any())
    }
    #[getter]
    fn protection(&self, py: Python<'_>) -> PyResult<Py<Protection>> {
        let ws = self.ws.bind(py).get();
        Ok(ws.shared.style_py(py, self.style_id)?.protection.clone_ref(py))
    }
    #[getter]
    fn is_date(&self, py: Python<'_>) -> bool {
        self.data_type == "d"
            || (self.data_type == "n" && style_of(self.ws.bind(py).get(), self.style_id).is_date)
    }
    #[getter]
    fn encoding(&self) -> &'static str {
        "utf-8"
    }
    #[getter]
    fn base_date(&self, py: Python<'_>) -> chrono::NaiveDateTime {
        self.ws.bind(py).get().shared.epoch
    }
    /// The cell `row` rows below and `column` columns right of this one.
    #[pyo3(signature = (row=0, column=0))]
    fn offset(&self, py: Python<'_>, row: i64, column: i64) -> PyResult<Py<PyAny>> {
        let ws = self.ws.bind(py);
        Worksheet::cell(ws, self.row as i64 + row, self.column as i64 + column, None)
    }
});

/// A cell of a read-only worksheet, like openpyxl's `ReadOnlyCell`.
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.cell")]
pub struct ReadOnlyCell {
    ws: Py<Worksheet>,
    #[pyo3(get)]
    row: u32,
    #[pyo3(get)]
    column: u32,
    #[pyo3(get)]
    value: Py<PyAny>,
    #[pyo3(get)]
    data_type: &'static str,
    style_id: u32,
    extra: Option<Box<CExtra>>,
    comment: Option<Py<Comment>>,
    hyperlink: Option<Py<Hyperlink>>,
    merged_range: Option<Py<MergedCellRange>>,
}

impl ReadOnlyCell {
    fn style_index(&self) -> u32 {
        self.style_id
    }
}
style_getters!(ReadOnlyCell {
    #[getter]
    fn internal_value(&self, py: Python<'_>) -> Py<PyAny> {
        self.value.clone_ref(py)
    }
    #[getter]
    fn has_style(&self) -> bool {
        self.style_id != 0 && self.style_id != ZERO_STYLE
    }
    /// The formula (`formula_and_value=True`), `None` without one.
    #[getter]
    fn formula(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        extra_value(py, &self.ws, &self.extra, "ReadOnlyCell", "formula", true)
    }
    /// The result saved in the file (`formula_and_value=True`).
    #[getter]
    fn cached_value(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        extra_value(py, &self.ws, &self.extra, "ReadOnlyCell", "cached_value", false)
    }
    /// The cell's comment (`read_comments=True`).
    #[getter]
    fn comment(&self, py: Python<'_>) -> PyResult<Option<Py<Comment>>> {
        if !self.ws.bind(py).get().shared.opts.comments {
            return Err(flag_error("ReadOnlyCell", "comment", "read_comments"));
        }
        Ok(self.comment.as_ref().map(|c| c.clone_ref(py)))
    }
    /// The cell's hyperlink (`read_hyperlinks=True`).
    #[getter]
    fn hyperlink(&self, py: Python<'_>) -> PyResult<Option<Py<Hyperlink>>> {
        if !self.ws.bind(py).get().shared.opts.hyperlinks {
            return Err(flag_error("ReadOnlyCell", "hyperlink", "read_hyperlinks"));
        }
        Ok(self.hyperlink.as_ref().map(|c| c.clone_ref(py)))
    }
    /// The merged range containing the cell (`read_merged_cells=True`).
    #[getter]
    fn merged_range(&self, py: Python<'_>) -> PyResult<Option<Py<MergedCellRange>>> {
        if !self.ws.bind(py).get().shared.opts.merged {
            return Err(flag_error("ReadOnlyCell", "merged_range", "read_merged_cells"));
        }
        Ok(self.merged_range.as_ref().map(|c| c.clone_ref(py)))
    }
    #[getter]
    fn border(&self, py: Python<'_>) -> PyResult<Py<Border>> {
        let ws = self.ws.bind(py).get();
        Ok(ws.shared.style_py(py, self.style_id)?.border.clone_ref(py))
    }
    #[getter]
    fn protection(&self, py: Python<'_>) -> PyResult<Py<Protection>> {
        let ws = self.ws.bind(py).get();
        Ok(ws.shared.style_py(py, self.style_id)?.protection.clone_ref(py))
    }
    #[getter]
    fn is_date(&self, py: Python<'_>) -> bool {
        self.data_type == "d"
            || (self.data_type == "n" && style_of(self.ws.bind(py).get(), self.style_id).is_date)
    }
});

/// A non-anchor cell of a merged range, like openpyxl's `MergedCell` (value `None`,
/// borders of the range's outer edges).
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.cell")]
pub struct MergedCell {
    ws: Py<Worksheet>,
    #[pyo3(get)]
    row: u32,
    #[pyo3(get)]
    column: u32,
    border: Border,
    protection: Protection,
}

impl MergedCell {
    fn style_index(&self) -> u32 {
        ZERO_STYLE
    }
}
style_getters!(MergedCell {
    /// Always `None` (`formula_and_value=True`).
    #[getter]
    fn formula(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        extra_value(py, &self.ws, &None, "MergedCell", "formula", true)
    }
    /// Always `None` (`formula_and_value=True`).
    #[getter]
    fn cached_value(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        extra_value(py, &self.ws, &None, "MergedCell", "cached_value", false)
    }
    #[getter]
    fn value(&self) -> Option<Py<PyAny>> {
        None
    }
    #[getter]
    fn internal_value(&self) -> Option<Py<PyAny>> {
        None
    }
    #[getter]
    fn data_type(&self) -> &'static str {
        "n"
    }
    #[getter]
    fn comment(&self) -> Option<Py<PyAny>> {
        None
    }
    #[getter]
    fn hyperlink(&self) -> Option<Py<PyAny>> {
        None
    }
    #[getter]
    fn has_style(&self) -> bool {
        false
    }
    #[getter]
    fn is_date(&self) -> bool {
        false
    }
    #[getter]
    fn border(&self, py: Python<'_>) -> PyResult<Py<Border>> {
        Py::new(py, self.border.clone())
    }
    #[getter]
    fn protection(&self, py: Python<'_>) -> PyResult<Py<Protection>> {
        Py::new(py, self.protection.clone())
    }
});

/// Padding cell of read-only rows, like openpyxl's `EmptyCell` (no position).
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.cell")]
pub struct EmptyCell {}

#[pymethods]
impl EmptyCell {
    #[getter]
    fn value(&self) -> Option<Py<PyAny>> {
        None
    }
    #[getter]
    fn is_date(&self) -> bool {
        false
    }
    #[getter]
    fn font(&self) -> Option<Py<PyAny>> {
        None
    }
    #[getter]
    fn border(&self) -> Option<Py<PyAny>> {
        None
    }
    #[getter]
    fn fill(&self) -> Option<Py<PyAny>> {
        None
    }
    #[getter]
    fn number_format(&self) -> Option<String> {
        None
    }
    #[getter]
    fn alignment(&self) -> Option<Py<PyAny>> {
        None
    }
    #[getter]
    fn data_type(&self) -> &'static str {
        "n"
    }
    fn __repr__(&self) -> &'static str {
        "<EmptyCell>"
    }
}

pub(crate) fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Worksheet>()?;
    m.add_class::<Chartsheet>()?;
    m.add_class::<RowIter>()?;
    m.add_class::<Cell>()?;
    m.add_class::<ReadOnlyCell>()?;
    m.add_class::<MergedCell>()?;
    m.add_class::<EmptyCell>()?;
    Ok(())
}
