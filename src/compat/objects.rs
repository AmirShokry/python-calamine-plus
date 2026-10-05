//! Read-only value objects shaped like openpyxl's: comments, hyperlinks, formulas,
//! merged ranges, row/column dimensions, defined names, document properties, tables.

use chrono::NaiveDateTime;
use pyo3::exceptions::{PyKeyError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyIterator, PyList, PyTuple};

use super::styles::Color;
use super::utils::{column_letter, coordinate, range_boundaries};

/// `(min_row, min_col, max_row, max_col)`, 1-based and inclusive.
pub(crate) type Bounds = (u32, u32, u32, u32);

/// Bounds of a cell or range coordinate (`"B2"` / `"B2:D4"`, `$` allowed).
pub(crate) fn parse_bounds(coord: &str) -> PyResult<Bounds> {
    match range_boundaries(coord)? {
        (Some(c0), Some(r0), Some(c1), Some(r1)) => {
            Ok((r0.min(r1), c0.min(c1), r0.max(r1), c0.max(c1)))
        }
        _ => Err(PyValueError::new_err(format!(
            "{coord} is not a valid coordinate or range"
        ))),
    }
}

pub(crate) fn bounds_coord(b: Bounds) -> String {
    if (b.0, b.1) == (b.2, b.3) {
        coordinate(b.0, b.1)
    } else {
        format!("{}:{}", coordinate(b.0, b.1), coordinate(b.2, b.3))
    }
}

// ---- Comment / Hyperlink -----------------------------------------------------------

/// A cell comment (note), like openpyxl's `Comment`.
#[pyclass(
    frozen,
    skip_from_py_object,
    module = "python_calamine_plus.openpyxl.comments"
)]
#[derive(Clone)]
pub struct Comment {
    #[pyo3(get)]
    pub(crate) content: String,
    #[pyo3(get)]
    pub(crate) author: Option<String>,
    #[pyo3(get)]
    height: u32,
    #[pyo3(get)]
    width: u32,
}

impl Comment {
    pub(crate) fn new(text: String, author: Option<String>) -> Self {
        Self {
            content: text,
            author,
            height: 79,
            width: 144,
        }
    }
}

#[pymethods]
impl Comment {
    #[getter]
    fn text(&self) -> String {
        self.content.clone()
    }

    fn __repr__(&self) -> String {
        format!(
            "Comment: {} by {}",
            self.content,
            self.author.as_deref().unwrap_or("None")
        )
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        other
            .cast::<Comment>()
            .map(|o| o.get().content == self.content && o.get().author == self.author)
            .unwrap_or(false)
    }
}

/// A hyperlink bound to a cell, like openpyxl's `Hyperlink` (`ref` is the cell's coordinate).
#[pyclass(
    frozen,
    skip_from_py_object,
    module = "python_calamine_plus.openpyxl.worksheet"
)]
#[derive(Clone)]
pub struct Hyperlink {
    #[pyo3(get, name = "ref")]
    pub(crate) reference: String,
    #[pyo3(get)]
    pub(crate) target: Option<String>,
    #[pyo3(get)]
    pub(crate) location: Option<String>,
    #[pyo3(get)]
    pub(crate) tooltip: Option<String>,
    #[pyo3(get)]
    pub(crate) display: Option<String>,
}

#[pymethods]
impl Hyperlink {
    /// The relationship id is not read (always `None`).
    #[getter]
    fn id(&self) -> Option<String> {
        None
    }

    fn __repr__(&self) -> String {
        format!(
            "<Hyperlink ref={:?} target={:?} location={:?}>",
            self.reference, self.target, self.location
        )
    }
}

// ---- formulas ----------------------------------------------------------------------

/// An array formula (`cell.value` of its anchor cell in formula view), like openpyxl's.
#[pyclass(
    frozen,
    skip_from_py_object,
    module = "python_calamine_plus.openpyxl.worksheet.formula"
)]
#[derive(Clone)]
pub struct ArrayFormula {
    #[pyo3(get, name = "ref")]
    pub(crate) reference: Option<String>,
    #[pyo3(get)]
    pub(crate) text: String,
}

#[pymethods]
impl ArrayFormula {
    #[getter]
    fn t(&self) -> &'static str {
        "array"
    }

    fn __iter__<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyIterator>> {
        let pairs = PyList::new(py, [("t", "array".to_string())])?;
        if let Some(r) = &self.reference {
            pairs.append(("ref", r))?;
        }
        pairs.try_iter()
    }

    fn __repr__(&self) -> String {
        format!(
            "<ArrayFormula ref={:?} text={:?}>",
            self.reference, self.text
        )
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        other
            .cast::<ArrayFormula>()
            .map(|o| o.get().reference == self.reference && o.get().text == self.text)
            .unwrap_or(false)
    }
}

/// A data table formula, like openpyxl's `DataTableFormula` (attribute values as
/// stored, e.g. `dt2D == "1"`).
#[pyclass(
    frozen,
    skip_from_py_object,
    module = "python_calamine_plus.openpyxl.worksheet.formula"
)]
#[derive(Clone)]
pub struct DataTableFormula {
    attrs: Vec<(String, String)>,
}

impl DataTableFormula {
    pub(crate) fn new(attrs: Vec<(String, String)>) -> Self {
        Self { attrs }
    }

    fn get(&self, py: Python<'_>, key: &str, default_false: bool) -> PyResult<Py<PyAny>> {
        match self.attrs.iter().find(|(k, _)| k == key) {
            Some((_, v)) => Ok(v.into_pyobject(py)?.into_any().unbind()),
            None if default_false => Ok(false.into_pyobject(py)?.to_owned().into_any().unbind()),
            None => Ok(py.None()),
        }
    }
}

#[pymethods]
impl DataTableFormula {
    #[getter]
    fn t(&self) -> &'static str {
        "dataTable"
    }
    #[getter(r#ref)]
    fn reference(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        self.get(py, "ref", false)
    }
    #[getter]
    fn ca(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        self.get(py, "ca", true)
    }
    #[getter(dt2D)]
    fn dt2d(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        self.get(py, "dt2D", true)
    }
    #[getter]
    fn dtr(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        self.get(py, "dtr", true)
    }
    #[getter]
    fn r1(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        self.get(py, "r1", false)
    }
    #[getter]
    fn r2(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        self.get(py, "r2", false)
    }
    #[getter]
    fn del1(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        self.get(py, "del1", true)
    }
    #[getter]
    fn del2(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        self.get(py, "del2", true)
    }

    fn __repr__(&self) -> String {
        format!("<DataTableFormula {:?}>", self.attrs)
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        other
            .cast::<DataTableFormula>()
            .map(|o| o.get().attrs == self.attrs)
            .unwrap_or(false)
    }
}

// ---- merged ranges -----------------------------------------------------------------

/// A merged cell range, like openpyxl's `MergedCellRange`.
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.worksheet.merge")]
pub struct MergedCellRange {
    pub(crate) bounds: Bounds,
    /// The worksheet (for `start_cell`), when known.
    pub(crate) ws: Option<Py<PyAny>>,
}

#[pymethods]
impl MergedCellRange {
    #[getter]
    fn min_row(&self) -> u32 {
        self.bounds.0
    }
    #[getter]
    fn min_col(&self) -> u32 {
        self.bounds.1
    }
    #[getter]
    fn max_row(&self) -> u32 {
        self.bounds.2
    }
    #[getter]
    fn max_col(&self) -> u32 {
        self.bounds.3
    }
    /// `(min_col, min_row, max_col, max_row)`.
    #[getter(bounds)]
    fn py_bounds(&self) -> (u32, u32, u32, u32) {
        (self.bounds.1, self.bounds.0, self.bounds.3, self.bounds.2)
    }
    #[getter]
    fn coord(&self) -> String {
        bounds_coord(self.bounds)
    }
    #[getter(r#ref)]
    fn reference(&self) -> String {
        self.coord()
    }
    #[getter]
    fn title(&self) -> Option<String> {
        None
    }
    #[getter]
    fn size<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, pyo3::types::PyDict>> {
        let d = pyo3::types::PyDict::new(py);
        d.set_item("columns", self.bounds.3 - self.bounds.1 + 1)?;
        d.set_item("rows", self.bounds.2 - self.bounds.0 + 1)?;
        Ok(d)
    }
    #[getter]
    fn top(&self) -> Vec<(u32, u32)> {
        (self.bounds.1..=self.bounds.3)
            .map(|c| (self.bounds.0, c))
            .collect()
    }
    #[getter]
    fn bottom(&self) -> Vec<(u32, u32)> {
        (self.bounds.1..=self.bounds.3)
            .map(|c| (self.bounds.2, c))
            .collect()
    }
    #[getter]
    fn left(&self) -> Vec<(u32, u32)> {
        (self.bounds.0..=self.bounds.2)
            .map(|r| (r, self.bounds.1))
            .collect()
    }
    #[getter]
    fn right(&self) -> Vec<(u32, u32)> {
        (self.bounds.0..=self.bounds.2)
            .map(|r| (r, self.bounds.3))
            .collect()
    }
    /// All `(row, column)` positions, row by row.
    #[getter]
    fn cells<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyIterator>> {
        let (r0, c0, r1, c1) = self.bounds;
        let all: Vec<(u32, u32)> = (r0..=r1)
            .flat_map(|r| (c0..=c1).map(move |c| (r, c)))
            .collect();
        PyList::new(py, all)?.try_iter()
    }
    #[getter]
    fn rows<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyIterator>> {
        let (r0, c0, r1, c1) = self.bounds;
        let rows = PyList::empty(py);
        for r in r0..=r1 {
            rows.append(PyList::new(py, (c0..=c1).map(|c| (r, c)))?)?;
        }
        rows.try_iter()
    }
    #[getter]
    fn cols<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyIterator>> {
        let (r0, c0, r1, c1) = self.bounds;
        let cols = PyList::empty(py);
        for c in c0..=c1 {
            cols.append(PyList::new(py, (r0..=r1).map(|r| (r, c)))?)?;
        }
        cols.try_iter()
    }
    /// The top-left cell of the range.
    #[getter]
    fn start_cell(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        match &self.ws {
            Some(ws) => Ok(ws
                .bind(py)
                .call_method1("cell", (self.bounds.0, self.bounds.1))?
                .unbind()),
            None => Ok(py.None()),
        }
    }

    fn __contains__(&self, coord: &str) -> PyResult<bool> {
        let b = parse_bounds(coord)?;
        Ok(contains(self.bounds, b))
    }

    fn __str__(&self) -> String {
        self.coord()
    }

    fn __repr__(&self) -> String {
        format!("<MergedCellRange {}>", self.coord())
    }

    fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
        if let Ok(s) = other.extract::<String>() {
            return parse_bounds(&s).is_ok_and(|b| b == self.bounds);
        }
        other
            .cast::<MergedCellRange>()
            .map(|o| o.get().bounds == self.bounds)
            .unwrap_or(false)
    }

    fn __hash__(&self) -> u64 {
        let (a, b, c, d) = self.bounds;
        ((a as u64) << 48) ^ ((b as u64) << 32) ^ ((c as u64) << 16) ^ d as u64
    }
}

fn contains(outer: Bounds, inner: Bounds) -> bool {
    outer.0 <= inner.0 && outer.1 <= inner.1 && inner.2 <= outer.2 && inner.3 <= outer.3
}

/// The merged ranges of a sheet, like openpyxl's `MultiCellRange`. `ranges` is a
/// list in file order (openpyxl uses a set, whose order is not stable).
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.worksheet.cell_range")]
pub struct MultiCellRange {
    pub(crate) ranges: Vec<Py<MergedCellRange>>,
}

impl MultiCellRange {
    fn sorted_bounds(&self, py: Python<'_>) -> Vec<Bounds> {
        let mut b: Vec<Bounds> = self
            .ranges
            .iter()
            .map(|r| r.bind(py).get().bounds)
            .collect();
        // openpyxl's `sorted()`: by min_col, min_row, max_col, max_row.
        b.sort_by_key(|b| (b.1, b.0, b.3, b.2));
        b
    }
}

#[pymethods]
impl MultiCellRange {
    #[getter]
    fn ranges<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        PyList::new(py, self.ranges.iter().map(|r| r.clone_ref(py)))
    }

    /// Ranges sorted by `(min_col, min_row, max_col, max_row)`.
    fn sorted<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let mut ranges: Vec<Py<MergedCellRange>> =
            self.ranges.iter().map(|r| r.clone_ref(py)).collect();
        ranges.sort_by_key(|r| {
            let b = r.bind(py).get().bounds;
            (b.1, b.0, b.3, b.2)
        });
        PyList::new(py, ranges)
    }

    fn __contains__(&self, py: Python<'_>, coord: &Bound<'_, PyAny>) -> PyResult<bool> {
        let b = if let Ok(r) = coord.cast::<MergedCellRange>() {
            r.get().bounds
        } else {
            parse_bounds(&coord.extract::<String>()?)?
        };
        Ok(self
            .ranges
            .iter()
            .any(|r| contains(r.bind(py).get().bounds, b)))
    }

    fn __iter__<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyIterator>> {
        self.ranges(py)?.try_iter()
    }

    fn __len__(&self) -> usize {
        self.ranges.len()
    }

    fn __bool__(&self) -> bool {
        !self.ranges.is_empty()
    }

    fn __str__(&self, py: Python<'_>) -> String {
        self.sorted_bounds(py)
            .into_iter()
            .map(bounds_coord)
            .collect::<Vec<_>>()
            .join(" ")
    }

    fn __repr__(&self, py: Python<'_>) -> String {
        format!("<MultiCellRange [{}]>", self.__str__(py))
    }
}

// ---- row / column dimensions -------------------------------------------------------

/// Row display properties, like openpyxl's `RowDimension`.
#[pyclass(
    frozen,
    skip_from_py_object,
    module = "python_calamine_plus.openpyxl.worksheet.dimensions"
)]
#[derive(Clone, Default)]
pub struct RowDimension {
    #[pyo3(get)]
    pub(crate) index: u32,
    #[pyo3(get)]
    pub(crate) ht: Option<f64>,
    #[pyo3(get)]
    pub(crate) hidden: bool,
    #[pyo3(get, name = "outlineLevel")]
    pub(crate) outline_level: u32,
    #[pyo3(get)]
    pub(crate) collapsed: bool,
    /// Style id of the row format (when the row has one).
    pub(crate) style: Option<u32>,
}

#[pymethods]
impl RowDimension {
    #[getter]
    fn r(&self) -> u32 {
        self.index
    }
    #[getter]
    fn height(&self) -> Option<f64> {
        self.ht
    }
    #[getter(customHeight)]
    fn custom_height(&self) -> bool {
        self.ht.is_some()
    }
    #[getter]
    fn outline_level(&self) -> u32 {
        self.outline_level
    }
    #[getter(customFormat)]
    fn custom_format(&self) -> bool {
        self.style.is_some_and(|s| s != 0)
    }
    #[getter(thickBot)]
    fn thick_bot(&self) -> bool {
        false
    }
    #[getter(thickTop)]
    fn thick_top(&self) -> bool {
        false
    }
    fn __repr__(&self) -> String {
        format!(
            "<RowDimension index={} ht={:?} hidden={} outlineLevel={}>",
            self.index, self.ht, self.hidden, self.outline_level
        )
    }
}

/// Column display properties, like openpyxl's `ColumnDimension` (`index` is the
/// letter of the first column of the span).
#[pyclass(
    frozen,
    skip_from_py_object,
    module = "python_calamine_plus.openpyxl.worksheet.dimensions"
)]
#[derive(Clone)]
pub struct ColumnDimension {
    #[pyo3(get)]
    pub(crate) index: String,
    #[pyo3(get)]
    pub(crate) width: f64,
    #[pyo3(get, name = "bestFit")]
    pub(crate) best_fit: bool,
    #[pyo3(get)]
    pub(crate) hidden: bool,
    #[pyo3(get, name = "outlineLevel")]
    pub(crate) outline_level: u32,
    #[pyo3(get)]
    pub(crate) collapsed: bool,
    #[pyo3(get)]
    pub(crate) min: Option<u32>,
    #[pyo3(get)]
    pub(crate) max: Option<u32>,
}

impl ColumnDimension {
    pub(crate) fn default_for(index: String) -> Self {
        Self {
            index,
            width: 13.0,
            best_fit: false,
            hidden: false,
            outline_level: 0,
            collapsed: false,
            min: None,
            max: None,
        }
    }
}

#[pymethods]
impl ColumnDimension {
    #[getter(customWidth)]
    fn custom_width(&self) -> bool {
        self.width != 0.0
    }
    #[getter]
    fn auto_size(&self) -> bool {
        self.best_fit
    }
    #[getter]
    fn outline_level(&self) -> u32 {
        self.outline_level
    }
    /// The columns covered, e.g. `"E:G"`.
    #[getter]
    fn range(&self) -> Option<String> {
        Some(format!(
            "{}:{}",
            column_letter(self.min?),
            column_letter(self.max?)
        ))
    }
    fn __repr__(&self) -> String {
        format!(
            "<ColumnDimension index={:?} width={} hidden={} min={:?} max={:?}>",
            self.index, self.width, self.hidden, self.min, self.max
        )
    }
}

/// `ws.row_dimensions` / `ws.column_dimensions`: a read-only mapping that returns a
/// default dimension for keys the file does not define (without storing it).
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.worksheet.dimensions")]
pub struct DimensionHolder {
    pub(crate) keys: Vec<Py<PyAny>>,
    pub(crate) values: Vec<Py<PyAny>>,
    pub(crate) rows: bool,
}

impl DimensionHolder {
    fn find(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Option<usize>> {
        for (i, k) in self.keys.iter().enumerate() {
            if k.bind(py).eq(key)? {
                return Ok(Some(i));
            }
        }
        Ok(None)
    }
}

#[pymethods]
impl DimensionHolder {
    fn __getitem__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<Py<PyAny>> {
        if let Some(i) = self.find(py, key)? {
            return Ok(self.values[i].clone_ref(py));
        }
        if self.rows {
            let index: u32 = key.extract()?;
            Ok(Py::new(
                py,
                RowDimension {
                    index,
                    ..Default::default()
                },
            )?
            .into_any())
        } else {
            let index: String = key.extract()?;
            Ok(Py::new(py, ColumnDimension::default_for(index))?.into_any())
        }
    }

    #[pyo3(signature = (key, default=None))]
    fn get(
        &self,
        py: Python<'_>,
        key: &Bound<'_, PyAny>,
        default: Option<Py<PyAny>>,
    ) -> PyResult<Option<Py<PyAny>>> {
        Ok(match self.find(py, key)? {
            Some(i) => Some(self.values[i].clone_ref(py)),
            None => default,
        })
    }

    fn __contains__(&self, py: Python<'_>, key: &Bound<'_, PyAny>) -> PyResult<bool> {
        Ok(self.find(py, key)?.is_some())
    }

    fn __len__(&self) -> usize {
        self.keys.len()
    }

    fn __iter__<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyIterator>> {
        self.keys(py)?.try_iter()
    }

    fn keys<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        PyList::new(py, self.keys.iter().map(|k| k.clone_ref(py)))
    }

    fn values<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        PyList::new(py, self.values.iter().map(|k| k.clone_ref(py)))
    }

    fn items<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let out = PyList::empty(py);
        for (k, v) in self.keys.iter().zip(&self.values) {
            out.append(PyTuple::new(py, [k.clone_ref(py), v.clone_ref(py)])?)?;
        }
        Ok(out)
    }

    fn __repr__(&self) -> String {
        format!("<DimensionHolder {} entries>", self.keys.len())
    }
}

// ---- defined names -----------------------------------------------------------------

/// A defined name, like openpyxl's `DefinedName`.
#[pyclass(
    frozen,
    skip_from_py_object,
    module = "python_calamine_plus.openpyxl.workbook.defined_name"
)]
#[derive(Clone)]
pub struct DefinedName {
    #[pyo3(get)]
    pub(crate) name: String,
    #[pyo3(get)]
    pub(crate) value: String,
    #[pyo3(get, name = "localSheetId")]
    pub(crate) local_sheet_id: Option<usize>,
    #[pyo3(get)]
    pub(crate) hidden: Option<bool>,
    #[pyo3(get)]
    pub(crate) comment: Option<String>,
}

const RESERVED: &[&str] = &[
    "Print_Area",
    "Print_Titles",
    "Criteria",
    "_FilterDatabase",
    "Extract",
    "Consolidate_Area",
    "Sheet_Title",
];

impl DefinedName {
    /// openpyxl's `is_reserved`: the name after `_xlnm.` for Excel's reserved names.
    pub(crate) fn reserved(&self) -> Option<&'static str> {
        let rest = self.name.strip_prefix("_xlnm.")?;
        RESERVED.iter().find(|r| rest.starts_with(**r)).copied()
    }
}

#[pymethods]
impl DefinedName {
    #[getter]
    fn attr_text(&self) -> String {
        self.value.clone()
    }

    #[getter]
    fn is_reserved(&self) -> Option<&'static str> {
        self.reserved()
    }

    #[getter]
    fn is_external(&self) -> bool {
        self.value.starts_with('[')
    }

    /// `(sheet title, cell range)` of each static range in the value (none for
    /// formulas and constants).
    #[getter]
    fn destinations<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyIterator>> {
        let out = PyList::empty(py);
        for (title, cells) in super::utils::sheet_ranges(&self.value) {
            let Some(title) = title else { continue };
            if !cells.is_empty()
                && range_boundaries(&cells).is_ok_and(|b| b.0.is_some() || b.1.is_some())
            {
                out.append((title, cells))?;
            }
        }
        out.try_iter()
    }

    fn __repr__(&self) -> String {
        format!("<DefinedName name={:?} value={:?}>", self.name, self.value)
    }
}

// ---- document properties -----------------------------------------------------------

/// Core document properties, like openpyxl's `DocumentProperties`. Properties the
/// file does not store are `None` (openpyxl fills in some defaults when reading).
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.packaging.core")]
#[derive(Default)]
pub struct DocumentProperties {
    #[pyo3(get)]
    pub(crate) creator: Option<String>,
    #[pyo3(get)]
    pub(crate) title: Option<String>,
    #[pyo3(get)]
    pub(crate) description: Option<String>,
    #[pyo3(get)]
    pub(crate) subject: Option<String>,
    #[pyo3(get)]
    pub(crate) identifier: Option<String>,
    #[pyo3(get)]
    pub(crate) language: Option<String>,
    #[pyo3(get)]
    pub(crate) created: Option<NaiveDateTime>,
    #[pyo3(get)]
    pub(crate) modified: Option<NaiveDateTime>,
    #[pyo3(get, name = "lastModifiedBy")]
    pub(crate) last_modified_by: Option<String>,
    #[pyo3(get)]
    pub(crate) category: Option<String>,
    #[pyo3(get, name = "contentStatus")]
    pub(crate) content_status: Option<String>,
    #[pyo3(get)]
    pub(crate) version: Option<String>,
    #[pyo3(get)]
    pub(crate) revision: Option<String>,
    #[pyo3(get)]
    pub(crate) keywords: Option<String>,
    #[pyo3(get, name = "lastPrinted")]
    pub(crate) last_printed: Option<NaiveDateTime>,
}

#[pymethods]
impl DocumentProperties {
    #[getter]
    fn last_modified_by(&self) -> Option<String> {
        self.last_modified_by.clone()
    }
    fn __repr__(&self) -> String {
        format!(
            "<DocumentProperties creator={:?} title={:?}>",
            self.creator, self.title
        )
    }
}

/// A custom document property (`wb.custom_doc_props`).
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.packaging.custom")]
pub struct CustomProperty {
    #[pyo3(get)]
    pub(crate) name: String,
    #[pyo3(get)]
    pub(crate) value: Py<PyAny>,
}

#[pymethods]
impl CustomProperty {
    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        Ok(format!(
            "<CustomProperty name={:?} value={}>",
            self.name,
            self.value.bind(py).repr()?
        ))
    }
}

// ---- tables, auto filter, sheet properties ----------------------------------------

/// `<tableStyleInfo>` of a table.
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.worksheet.table")]
pub struct TableStyleInfo {
    #[pyo3(get)]
    pub(crate) name: Option<String>,
    #[pyo3(get, name = "showFirstColumn")]
    pub(crate) show_first_column: Option<bool>,
    #[pyo3(get, name = "showLastColumn")]
    pub(crate) show_last_column: Option<bool>,
    #[pyo3(get, name = "showRowStripes")]
    pub(crate) show_row_stripes: Option<bool>,
    #[pyo3(get, name = "showColumnStripes")]
    pub(crate) show_column_stripes: Option<bool>,
}

/// A column of a table.
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.worksheet.table")]
pub struct TableColumn {
    #[pyo3(get)]
    pub(crate) id: Option<i64>,
    #[pyo3(get)]
    pub(crate) name: Option<String>,
    #[pyo3(get, name = "totalsRowFunction")]
    pub(crate) totals_row_function: Option<String>,
    #[pyo3(get, name = "totalsRowLabel")]
    pub(crate) totals_row_label: Option<String>,
}

/// `ws.auto_filter` (only `ref` is provided).
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.worksheet.filters")]
pub struct AutoFilter {
    #[pyo3(get, name = "ref")]
    pub(crate) reference: Option<String>,
}

#[pymethods]
impl AutoFilter {
    fn __repr__(&self) -> String {
        format!("<AutoFilter ref={:?}>", self.reference)
    }
}

/// A table (list object), like openpyxl's `Table`.
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.worksheet.table")]
pub struct Table {
    #[pyo3(get)]
    pub(crate) id: Option<i64>,
    #[pyo3(get)]
    pub(crate) name: Option<String>,
    #[pyo3(get, name = "displayName")]
    pub(crate) display_name: Option<String>,
    #[pyo3(get, name = "ref")]
    pub(crate) reference: Option<String>,
    #[pyo3(get, name = "headerRowCount")]
    pub(crate) header_row_count: i64,
    #[pyo3(get, name = "totalsRowCount")]
    pub(crate) totals_row_count: Option<i64>,
    #[pyo3(get, name = "totalsRowShown")]
    pub(crate) totals_row_shown: Option<bool>,
    #[pyo3(get, name = "tableType")]
    pub(crate) table_type: Option<String>,
    #[pyo3(get)]
    pub(crate) comment: Option<String>,
    #[pyo3(get, name = "tableStyleInfo")]
    pub(crate) style: Option<Py<TableStyleInfo>>,
    #[pyo3(get, name = "tableColumns")]
    pub(crate) columns: Vec<Py<TableColumn>>,
    #[pyo3(get, name = "autoFilter")]
    pub(crate) auto_filter: Option<Py<AutoFilter>>,
}

#[pymethods]
impl Table {
    #[getter]
    fn column_names(&self, py: Python<'_>) -> Vec<Option<String>> {
        self.columns
            .iter()
            .map(|c| c.bind(py).get().name.clone())
            .collect()
    }
    fn __repr__(&self) -> String {
        format!("<Table name={:?} ref={:?}>", self.name, self.reference)
    }
}

/// `ws.tables`: tables by name. Like openpyxl's `TableList`, `items()` returns
/// `(name, ref)` pairs.
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.worksheet.table")]
pub struct TableList {
    pub(crate) tables: Vec<(String, Py<Table>)>,
}

#[pymethods]
impl TableList {
    fn __getitem__(&self, py: Python<'_>, key: &str) -> PyResult<Py<Table>> {
        self.tables
            .iter()
            .find(|(n, _)| n == key)
            .map(|(_, t)| t.clone_ref(py))
            .ok_or_else(|| PyKeyError::new_err(key.to_string()))
    }
    #[pyo3(signature = (key, default=None))]
    fn get(&self, py: Python<'_>, key: &str, default: Option<Py<PyAny>>) -> Option<Py<PyAny>> {
        self.tables
            .iter()
            .find(|(n, _)| n == key)
            .map(|(_, t)| t.clone_ref(py).into_any())
            .or(default)
    }
    fn __contains__(&self, key: &str) -> bool {
        self.tables.iter().any(|(n, _)| n == key)
    }
    fn __len__(&self) -> usize {
        self.tables.len()
    }
    fn __iter__<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyIterator>> {
        PyList::new(py, self.tables.iter().map(|(n, _)| n))?.try_iter()
    }
    fn keys(&self) -> Vec<String> {
        self.tables.iter().map(|(n, _)| n.clone()).collect()
    }
    fn values(&self, py: Python<'_>) -> Vec<Py<Table>> {
        self.tables.iter().map(|(_, t)| t.clone_ref(py)).collect()
    }
    fn items(&self, py: Python<'_>) -> Vec<(String, Option<String>)> {
        self.tables
            .iter()
            .map(|(n, t)| (n.clone(), t.bind(py).get().reference.clone()))
            .collect()
    }
}

/// `ws.sheet_properties` (`tabColor` and `codeName` only).
#[pyclass(frozen, module = "python_calamine_plus.openpyxl.worksheet.properties")]
pub struct WorksheetProperties {
    #[pyo3(get, name = "tabColor")]
    pub(crate) tab_color: Option<Color>,
    #[pyo3(get, name = "codeName")]
    pub(crate) code_name: Option<String>,
}

pub(crate) fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Comment>()?;
    m.add_class::<Hyperlink>()?;
    m.add_class::<ArrayFormula>()?;
    m.add_class::<DataTableFormula>()?;
    m.add_class::<MergedCellRange>()?;
    m.add_class::<MultiCellRange>()?;
    m.add_class::<RowDimension>()?;
    m.add_class::<ColumnDimension>()?;
    m.add_class::<DimensionHolder>()?;
    m.add_class::<DefinedName>()?;
    m.add_class::<DocumentProperties>()?;
    m.add_class::<CustomProperty>()?;
    m.add_class::<Table>()?;
    m.add_class::<TableColumn>()?;
    m.add_class::<TableStyleInfo>()?;
    m.add_class::<TableList>()?;
    m.add_class::<AutoFilter>()?;
    m.add_class::<WorksheetProperties>()?;
    Ok(())
}
