use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek};
use std::path::PathBuf;
use std::sync::Arc;

use calamine::{
    open_workbook_auto, open_workbook_auto_from_rs, Comment, DefinedName,
    Error as CalamineCrateError, Reader, Sheets, StyleSheet, TextRun, WorkbookInfo, WorksheetInfo,
    Xlsx, XlsxCellReader, XlsxError,
};
use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList, PyType};
use pyo3_file::PyFileLikeObject;

use crate::types::convert::{
    attrs_to_py, custom_properties_to_py, named_styles_to_py, parse_a1, parse_w3c_datetime,
};
use crate::types::stream::{CellReader, RichInfo, StreamRange, StreamRow, StreamState};
use crate::{
    CalamineSheet, CalamineSheetStream, CalamineTable, Error, SheetMetadata, WorksheetNotFound,
};

enum SheetsEnum {
    File(Sheets<BufReader<File>>),
    FileLike(Sheets<Cursor<Vec<u8>>>),
    None,
}

enum WorkbookType {
    Xls,
    Xlsx,
    Xlsb,
    Ods,
}

impl From<&SheetsEnum> for WorkbookType {
    fn from(sheets: &SheetsEnum) -> Self {
        match sheets {
            SheetsEnum::File(f) => match f {
                Sheets::Xls(_) => WorkbookType::Xls,
                Sheets::Xlsx(_) => WorkbookType::Xlsx,
                Sheets::Xlsb(_) => WorkbookType::Xlsb,
                Sheets::Ods(_) => WorkbookType::Ods,
            },
            SheetsEnum::FileLike(f) => match f {
                Sheets::Xls(_) => WorkbookType::Xls,
                Sheets::Xlsx(_) => WorkbookType::Xlsx,
                Sheets::Xlsb(_) => WorkbookType::Xlsb,
                Sheets::Ods(_) => WorkbookType::Ods,
            },
            SheetsEnum::None => unreachable!(),
        }
    }
}

impl SheetsEnum {
    fn sheets_metadata(&self) -> Vec<SheetMetadata> {
        match self {
            SheetsEnum::File(f) => f.sheets_metadata(),
            SheetsEnum::FileLike(f) => f.sheets_metadata(),
            SheetsEnum::None => unreachable!(),
        }
        .iter()
        .map(|s| SheetMetadata::new(s.name.clone(), s.typ, s.visible))
        .collect()
    }

    fn sheet_names(&self) -> Vec<String> {
        match self {
            SheetsEnum::File(f) => f.sheet_names(),
            SheetsEnum::FileLike(f) => f.sheet_names(),
            SheetsEnum::None => unreachable!(),
        }
    }

    fn worksheet_range(&mut self, name: &str) -> Result<calamine::Range<calamine::Data>, Error> {
        match self {
            SheetsEnum::File(f) => f.worksheet_range(name).map_err(Error::Calamine),
            SheetsEnum::FileLike(f) => f.worksheet_range(name).map_err(Error::Calamine),
            SheetsEnum::None => Err(Error::WorkbookClosed),
        }
    }

    fn worksheet_merge_cells(
        &mut self,
        name: &str,
    ) -> Result<Option<Vec<calamine::Dimensions>>, Error> {
        match self {
            SheetsEnum::File(f) => match f {
                Sheets::Xls(xls_f) => xls_f
                    .merge_cells_by_sheet_name(name)
                    .map(Some)
                    .map_err(CalamineCrateError::Xls)
                    .map_err(Error::Calamine),
                Sheets::Xlsx(xlsx_f) => xlsx_f
                    .merge_cells_by_sheet_name(name)
                    .map(Some)
                    .map_err(CalamineCrateError::Xlsx)
                    .map_err(Error::Calamine),
                _ => Ok(None),
            },
            SheetsEnum::FileLike(f) => match f {
                Sheets::Xls(xls_f) => xls_f
                    .merge_cells_by_sheet_name(name)
                    .map(Some)
                    .map_err(CalamineCrateError::Xls)
                    .map_err(Error::Calamine),
                Sheets::Xlsx(xlsx_f) => xlsx_f
                    .merge_cells_by_sheet_name(name)
                    .map(Some)
                    .map_err(CalamineCrateError::Xlsx)
                    .map_err(Error::Calamine),
                _ => Ok(None),
            },
            SheetsEnum::None => Err(Error::WorkbookClosed),
        }
    }

    fn worksheet_comments(&mut self, name: &str) -> Result<Vec<Comment>, Error> {
        match self {
            SheetsEnum::File(Sheets::Xlsx(x)) => x.worksheet_comments(name),
            SheetsEnum::FileLike(Sheets::Xlsx(x)) => x.worksheet_comments(name),
            SheetsEnum::None => return Err(Error::WorkbookClosed),
            _ => Ok(Vec::new()),
        }
        .map_err(|e| Error::Calamine(CalamineCrateError::Xlsx(e)))
    }

    fn stylesheet(&mut self) -> Result<StyleSheet, Error> {
        match self {
            SheetsEnum::File(Sheets::Xlsx(x)) => x.stylesheet(),
            SheetsEnum::FileLike(Sheets::Xlsx(x)) => x.stylesheet(),
            SheetsEnum::None => return Err(Error::WorkbookClosed),
            _ => Ok(StyleSheet::default()),
        }
        .map_err(|e| Error::Calamine(CalamineCrateError::Xlsx(e)))
    }

    fn worksheet_info(&mut self, name: &str) -> Result<WorksheetInfo, Error> {
        match self {
            SheetsEnum::File(Sheets::Xlsx(x)) => x.worksheet_info(name),
            SheetsEnum::FileLike(Sheets::Xlsx(x)) => x.worksheet_info(name),
            SheetsEnum::None => return Err(Error::WorkbookClosed),
            _ => Ok(WorksheetInfo::default()),
        }
        .map_err(|e| Error::Calamine(CalamineCrateError::Xlsx(e)))
    }

    fn rich_shared_strings(&mut self) -> Result<HashMap<usize, Vec<TextRun>>, Error> {
        match self {
            SheetsEnum::File(Sheets::Xlsx(x)) => x.rich_shared_strings(),
            SheetsEnum::FileLike(Sheets::Xlsx(x)) => x.rich_shared_strings(),
            SheetsEnum::None => return Err(Error::WorkbookClosed),
            _ => Ok(HashMap::new()),
        }
        .map_err(|e| Error::Calamine(CalamineCrateError::Xlsx(e)))
    }

    /// Defined names; with sheet scope for xlsx, workbook scope only for other formats.
    fn defined_names(&mut self) -> Result<Vec<DefinedName>, Error> {
        let unscoped = |names: &[(String, String)]| {
            names
                .iter()
                .map(|(name, value)| DefinedName {
                    name: name.clone(),
                    value: value.clone(),
                    sheet: None,
                    hidden: false,
                    comment: None,
                })
                .collect()
        };
        match self {
            SheetsEnum::File(Sheets::Xlsx(x)) => x.defined_names_scoped(),
            SheetsEnum::FileLike(Sheets::Xlsx(x)) => x.defined_names_scoped(),
            SheetsEnum::File(f) => Ok(unscoped(f.defined_names())),
            SheetsEnum::FileLike(f) => Ok(unscoped(f.defined_names())),
            SheetsEnum::None => return Err(Error::WorkbookClosed),
        }
        .map_err(|e| Error::Calamine(CalamineCrateError::Xlsx(e)))
    }

    fn workbook_info(&mut self) -> Result<WorkbookInfo, Error> {
        match self {
            SheetsEnum::File(Sheets::Xlsx(x)) => x.workbook_info(),
            SheetsEnum::FileLike(Sheets::Xlsx(x)) => x.workbook_info(),
            SheetsEnum::None => return Err(Error::WorkbookClosed),
            _ => Ok(WorkbookInfo::default()),
        }
        .map_err(|e| Error::Calamine(CalamineCrateError::Xlsx(e)))
    }

    /// Workbook-level metadata, read once at open (small parts only).
    fn metadata(&mut self) -> Result<WorkbookMeta, Error> {
        Ok(WorkbookMeta {
            names: self.defined_names()?,
            properties: self.document_properties()?,
            info: self.workbook_info()?,
            stylesheet: Arc::new(self.stylesheet()?),
        })
    }

    fn document_properties(&mut self) -> Result<Vec<(String, String)>, Error> {
        match self {
            SheetsEnum::File(Sheets::Xlsx(x)) => x.document_properties(),
            SheetsEnum::FileLike(Sheets::Xlsx(x)) => x.document_properties(),
            SheetsEnum::None => return Err(Error::WorkbookClosed),
            _ => Ok(Vec::new()),
        }
        .map_err(|e| Error::Calamine(CalamineCrateError::Xlsx(e)))
    }

    fn load_tables(&mut self) -> Result<(), Error> {
        match self {
            SheetsEnum::File(f) => match f {
                Sheets::Xlsx(xlsx_f) => xlsx_f
                    .load_tables()
                    .map_err(CalamineCrateError::Xlsx)
                    .map_err(Error::Calamine),
                _ => Err(Error::TablesNotSupported),
            },
            SheetsEnum::FileLike(f) => match f {
                Sheets::Xlsx(xlsx_f) => xlsx_f
                    .load_tables()
                    .map_err(CalamineCrateError::Xlsx)
                    .map_err(Error::Calamine),
                _ => Err(Error::TablesNotSupported),
            },
            SheetsEnum::None => Err(Error::WorkbookClosed),
        }
    }

    fn table_names(&self) -> Result<Vec<String>, Error> {
        match self {
            SheetsEnum::File(f) => match f {
                Sheets::Xlsx(xlsx_f) => Ok(xlsx_f.table_names()),
                _ => Err(Error::TablesNotSupported),
            },
            SheetsEnum::FileLike(f) => match f {
                Sheets::Xlsx(xlsx_f) => Ok(xlsx_f.table_names()),
                _ => Err(Error::TablesNotSupported),
            },
            SheetsEnum::None => Err(Error::WorkbookClosed),
        }
        .map(|v| {
            v.iter()
                .map(|s| s.to_owned().to_owned())
                .collect::<Vec<String>>()
        })
    }

    fn get_table_by_name(&mut self, name: &str) -> Result<CalamineTable, Error> {
        match self {
            SheetsEnum::File(f) => match f {
                Sheets::Xlsx(xlsx_f) => xlsx_f
                    .table_by_name(name)
                    .map_err(CalamineCrateError::Xlsx)
                    .map_err(Error::Calamine)
                    .map(|t| {
                        CalamineTable::new(
                            t.name().to_owned(),
                            t.sheet_name().to_owned(),
                            t.columns().iter().map(|s| s.to_owned()).collect(),
                            t.data().to_owned(),
                        )
                    }),
                _ => Err(Error::TablesNotSupported),
            },
            SheetsEnum::FileLike(f) => match f {
                Sheets::Xlsx(xlsx_f) => xlsx_f
                    .table_by_name(name)
                    .map_err(CalamineCrateError::Xlsx)
                    .map_err(Error::Calamine)
                    .map(|t| {
                        CalamineTable::new(
                            t.name().to_owned(),
                            t.sheet_name().to_owned(),
                            t.columns().iter().map(|s| s.to_owned()).collect(),
                            t.data().to_owned(),
                        )
                    }),
                _ => Err(Error::TablesNotSupported),
            },
            SheetsEnum::None => Err(Error::WorkbookClosed),
        }
    }
}

/// Opens a cell reader for `name`; `None` for non-worksheets (e.g. chart sheets).
fn cells_reader<'a, RS: Read + Seek>(
    xlsx: &'a mut Xlsx<RS>,
    name: &str,
) -> Result<Option<XlsxCellReader<'a, RS>>, Error> {
    match xlsx.worksheet_cells_reader(name) {
        Ok(reader) => Ok(Some(reader)),
        Err(XlsxError::NotAWorksheet(_)) => Ok(None),
        Err(e) => Err(Error::Calamine(CalamineCrateError::Xlsx(e))),
    }
}

/// Workbook-level metadata openpyxl exposes.
struct WorkbookMeta {
    names: Vec<DefinedName>,
    properties: Vec<(String, String)>,
    info: WorkbookInfo,
    stylesheet: Arc<StyleSheet>,
}

#[pyclass]
pub struct CalamineWorkbook {
    #[pyo3(get)]
    path: Option<String>,
    workbook_type: WorkbookType,
    // The active stream's reader borrows from `sheets`, so it is declared first
    // (dropped first), and cleared before any other access to `sheets`.
    stream: Option<StreamState>,
    stream_seq: u64,
    // Parsed styles.xml and rich text runs, loaded by the first `rich=True` stream.
    stylesheet: Option<Arc<StyleSheet>>,
    rich_strings: Option<Arc<HashMap<usize, Vec<TextRun>>>>,
    // Small parts read at open, so they never conflict with an active stream. A
    // failure is reported when the metadata is accessed, not when opening.
    meta: Result<WorkbookMeta, String>,
    sheets: SheetsEnum,
    #[pyo3(get)]
    sheets_metadata: Vec<SheetMetadata>,
    #[pyo3(get)]
    sheet_names: Vec<String>,
    table_names: Option<Vec<String>>,
}

#[pymethods]
impl CalamineWorkbook {
    fn __repr__(&self) -> PyResult<String> {
        match &self.path {
            Some(path) => Ok(format!("CalamineWorkbook(path='{path}')")),
            None => Ok("CalamineWorkbook(path='bytes')".to_string()),
        }
    }

    #[classmethod]
    #[pyo3(name = "from_object", signature = (path_or_filelike, load_tables=false))]
    fn py_from_object(
        _cls: &Bound<'_, PyType>,
        py: Python<'_>,
        path_or_filelike: Py<PyAny>,
        load_tables: bool,
    ) -> PyResult<Self> {
        Self::from_object(py, path_or_filelike, load_tables)
    }

    #[classmethod]
    #[pyo3(name = "from_filelike", signature = (filelike, load_tables=false))]
    fn py_from_filelike(
        _cls: &Bound<'_, PyType>,
        py: Python<'_>,
        filelike: Py<PyAny>,
        load_tables: bool,
    ) -> PyResult<Self> {
        py.detach(|| Self::from_filelike(filelike, load_tables))
    }

    #[classmethod]
    #[pyo3(name = "from_path", signature = (path, load_tables=false))]
    fn py_from_path(
        _cls: &Bound<'_, PyType>,
        py: Python<'_>,
        path: Py<PyAny>,
        load_tables: bool,
    ) -> PyResult<Self> {
        if let Ok(string_ref) = path.extract::<PathBuf>(py) {
            let path = string_ref.to_string_lossy().to_string();
            return py.detach(|| Self::from_path(&path, load_tables));
        }

        Err(PyTypeError::new_err(""))
    }

    #[pyo3(name = "get_sheet_by_name")]
    fn py_get_sheet_by_name(&mut self, py: Python<'_>, name: &str) -> PyResult<CalamineSheet> {
        py.detach(|| self.get_sheet_by_name(name))
    }

    #[pyo3(name = "get_sheet_by_index")]
    fn py_get_sheet_by_index(&mut self, py: Python<'_>, index: usize) -> PyResult<CalamineSheet> {
        py.detach(|| self.get_sheet_by_index(index))
    }

    #[pyo3(
        name = "stream_sheet_by_name",
        signature = (name, rich=false, min_row=0, max_row=None, min_col=0, max_col=None, r#ref=None)
    )]
    #[allow(clippy::too_many_arguments)] // the Python keyword signature
    fn py_stream_sheet_by_name(
        slf: &Bound<'_, Self>,
        name: &str,
        rich: bool,
        min_row: u32,
        max_row: Option<u32>,
        min_col: u32,
        max_col: Option<u32>,
        r#ref: Option<&str>,
    ) -> PyResult<CalamineSheetStream> {
        let (mut min_row, mut max_row, mut min_col, mut max_col) =
            (min_row, max_row, min_col, max_col);
        if let Some(r) = r#ref {
            // "B2:D10" or a single cell "B2" (openpyxl's ws["B2:D10"]).
            let mut parts = r.split(':');
            let bad = || pyo3::exceptions::PyValueError::new_err(format!("invalid ref: {r:?}"));
            let start = parse_a1(parts.next().unwrap_or_default()).ok_or_else(bad)?;
            let end = match parts.next() {
                Some(p) => parse_a1(p).ok_or_else(bad)?,
                None => start,
            };
            if parts.next().is_some() {
                return Err(bad());
            }
            (min_row, max_row) = (start.0.min(end.0), Some(start.0.max(end.0)));
            (min_col, max_col) = (start.1.min(end.1), Some(start.1.max(end.1)));
        }
        if max_row.is_some_and(|m| m < min_row) || max_col.is_some_and(|m| m < min_col) {
            return Err(pyo3::exceptions::PyValueError::new_err(
                "max_row/max_col must not be smaller than min_row/min_col",
            ));
        }
        let range = StreamRange {
            min_row,
            max_row,
            min_col,
            max_col,
        };
        let py = slf.py();
        let (id, info) = {
            let mut wb = slf.borrow_mut();
            let wb: &mut Self = &mut wb;
            py.detach(|| wb.start_stream(name, rich, range))?
        };
        CalamineSheetStream::new(py, name.to_owned(), slf.clone().unbind(), id, info)
    }

    #[pyo3(
        name = "stream_sheet_by_index",
        signature = (index, rich=false, min_row=0, max_row=None, min_col=0, max_col=None, r#ref=None)
    )]
    #[allow(clippy::too_many_arguments)] // the Python keyword signature
    fn py_stream_sheet_by_index(
        slf: &Bound<'_, Self>,
        index: usize,
        rich: bool,
        min_row: u32,
        max_row: Option<u32>,
        min_col: u32,
        max_col: Option<u32>,
        r#ref: Option<&str>,
    ) -> PyResult<CalamineSheetStream> {
        let name = slf.borrow().sheet_name_by_index(index)?;
        Self::py_stream_sheet_by_name(slf, &name, rich, min_row, max_row, min_col, max_col, r#ref)
    }

    /// Defined names: `{"name", "value", "sheet", "hidden", "comment"}` dicts.
    #[getter]
    fn defined_names<'py>(&self, py: Python<'py>) -> PyResult<Vec<Bound<'py, PyDict>>> {
        self.meta()?
            .names
            .iter()
            .map(|n| {
                let d = PyDict::new(py);
                d.set_item("name", &n.name)?;
                d.set_item("value", &n.value)?;
                d.set_item("sheet", &n.sheet)?;
                d.set_item("hidden", n.hidden)?;
                d.set_item("comment", &n.comment)?;
                Ok(d)
            })
            .collect()
    }

    /// Document properties (author, created, company, ...); dates as `datetime` (UTC).
    #[getter]
    fn properties<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        let d = PyDict::new(py);
        for (k, v) in &self.meta()?.properties {
            if matches!(k.as_str(), "created" | "modified" | "last_printed") {
                if let Some(dt) = parse_w3c_datetime(v) {
                    d.set_item(k, dt)?;
                    continue;
                }
            }
            d.set_item(k, v)?;
        }
        Ok(d)
    }

    /// Custom document properties (`docProps/custom.xml`), typed.
    #[getter]
    fn custom_properties<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        custom_properties_to_py(py, &self.meta()?.info.custom_properties)
    }

    /// Dates use the 1904 date system (openpyxl's `epoch == CALENDAR_MAC_1904`).
    #[getter]
    fn epoch_1904(&self) -> PyResult<bool> {
        Ok(self
            .meta()?
            .info
            .workbook_properties
            .iter()
            .any(|(k, v)| k == "date1904" && (v == "1" || v == "true")))
    }

    /// VBA code name of the workbook.
    #[getter]
    fn code_name(&self) -> PyResult<Option<String>> {
        Ok(self
            .meta()?
            .info
            .workbook_properties
            .iter()
            .find(|(k, _)| k == "codeName")
            .map(|(_, v)| v.clone()))
    }

    /// `<workbookPr>` attributes (`date1904`, `code_name`, `filter_privacy`, ...).
    #[getter]
    fn workbook_properties<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        attrs_to_py(py, &self.meta()?.info.workbook_properties)
    }

    /// Index of the active (selected) sheet (openpyxl's `wb.active`).
    #[getter]
    fn active_sheet_index(&self) -> PyResult<usize> {
        Ok(self
            .meta()?
            .info
            .views
            .first()
            .and_then(|v| v.iter().find(|(k, _)| k == "activeTab"))
            .and_then(|(_, v)| v.parse().ok())
            .unwrap_or(0))
    }

    /// Workbook window views (`active_tab`, `first_sheet`, `visibility`, ...).
    #[getter]
    fn workbook_views<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        let list = PyList::empty(py);
        for v in &self.meta()?.info.views {
            list.append(attrs_to_py(py, v)?)?;
        }
        Ok(list)
    }

    /// Calculation properties (`calc_id`, `full_calc_on_load`, ...), or `None`.
    #[getter]
    fn calculation<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyDict>>> {
        let meta = self.meta()?;
        meta.info
            .calculation
            .as_ref()
            .map(|a| attrs_to_py(py, a))
            .transpose()
    }

    /// Workbook protection (`lock_structure`, `lock_windows`, password hashes), or `None`.
    #[getter]
    fn workbook_protection<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyDict>>> {
        let meta = self.meta()?;
        meta.info
            .protection
            .as_ref()
            .map(|a| attrs_to_py(py, a))
            .transpose()
    }

    /// `<fileVersion>` attributes (`app_name`, `last_edited`, ...), or `None`.
    #[getter]
    fn file_version<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyDict>>> {
        let meta = self.meta()?;
        meta.info
            .file_version
            .as_ref()
            .map(|a| attrs_to_py(py, a))
            .transpose()
    }

    /// Theme color scheme: `{"dk1": "000000", "lt1": "FFFFFF", "accent1": ...}`.
    #[getter]
    fn theme_colors<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        let d = PyDict::new(py);
        for (k, v) in &self.meta()?.info.theme_colors {
            d.set_item(k, v)?;
        }
        Ok(d)
    }

    /// Raw theme XML (openpyxl's `loaded_theme`), or `None`.
    #[getter]
    fn theme_xml<'py>(&self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyBytes>>> {
        Ok(self
            .meta()?
            .info
            .theme_xml
            .as_ref()
            .map(|b| PyBytes::new(py, b)))
    }

    /// The file is a template (`.xltx` / `.xltm`; openpyxl's `wb.template`).
    #[getter]
    fn is_template(&self) -> PyResult<bool> {
        Ok(self.meta()?.info.is_template)
    }

    /// The file is macro-enabled (`.xlsm` / `.xltm`).
    #[getter]
    fn has_macros(&self) -> PyResult<bool> {
        Ok(self.meta()?.info.has_macros)
    }

    /// Named cell styles: `{"name", "builtin_id", "hidden"}` dicts.
    #[getter]
    fn named_styles<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyList>> {
        named_styles_to_py(py, &self.meta()?.stylesheet.named_styles)
    }

    #[getter]
    fn table_names(&self) -> PyResult<Vec<String>> {
        match &self.workbook_type {
            WorkbookType::Xlsx => match &self.table_names {
                Some(v) => Ok(v.clone()),
                None => Err(Error::TablesNotLoaded.into()),
            },
            _ => Err(Error::TablesNotSupported.into()),
        }
    }

    #[pyo3(name = "get_table_by_name")]
    fn py_get_table_by_name(&mut self, py: Python<'_>, name: &str) -> PyResult<CalamineTable> {
        py.detach(|| self.get_table_by_name(name))
    }

    fn close(&mut self) -> PyResult<()> {
        self.stream = None;
        match self.sheets {
            SheetsEnum::None => Err(Error::WorkbookClosed.into()),
            _ => {
                self.sheets = SheetsEnum::None;
                Ok(())
            }
        }
    }

    fn __enter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __exit__(
        &mut self,
        _exc_type: Py<PyAny>,
        _exc_value: Py<PyAny>,
        _traceback: Py<PyAny>,
    ) -> PyResult<()> {
        self.close()
    }
}

impl CalamineWorkbook {
    pub fn from_object(
        py: Python<'_>,
        path_or_filelike: Py<PyAny>,
        load_tables: bool,
    ) -> PyResult<Self> {
        if let Ok(string_ref) = path_or_filelike.extract::<PathBuf>(py) {
            let path = string_ref.to_string_lossy().to_string();
            return py.detach(|| Self::from_path(&path, load_tables));
        }

        py.detach(|| Self::from_filelike(path_or_filelike, load_tables))
    }

    pub fn from_filelike(filelike: Py<PyAny>, load_tables: bool) -> PyResult<Self> {
        let mut buf = vec![];
        PyFileLikeObject::with_requirements(filelike, true, false, true, false)?
            .read_to_end(&mut buf)?;
        let reader = Cursor::new(buf);
        let mut sheets =
            SheetsEnum::FileLike(open_workbook_auto_from_rs(reader).map_err(Error::Calamine)?);
        let sheet_names = sheets.sheet_names().to_owned();
        let sheets_metadata = sheets.sheets_metadata().to_owned();
        let meta = sheets.metadata().map_err(|e| PyErr::from(e).to_string());
        let stylesheet = meta.as_ref().ok().map(|m| Arc::clone(&m.stylesheet));

        let mut table_names: Option<Vec<String>> = None;
        if load_tables {
            sheets.load_tables()?;
            table_names = Some(sheets.table_names()?);
        }

        Ok(Self {
            path: None,
            workbook_type: WorkbookType::from(&sheets),
            stream: None,
            stream_seq: 0,
            stylesheet,
            rich_strings: None,
            meta,
            sheets,
            sheets_metadata,
            sheet_names,
            table_names,
        })
    }

    pub fn from_path(path: &str, load_tables: bool) -> PyResult<Self> {
        let mut sheets = SheetsEnum::File(open_workbook_auto(path).map_err(Error::Calamine)?);
        let sheet_names = sheets.sheet_names().to_owned();
        let sheets_metadata = sheets.sheets_metadata().to_owned();
        let meta = sheets.metadata().map_err(|e| PyErr::from(e).to_string());
        let stylesheet = meta.as_ref().ok().map(|m| Arc::clone(&m.stylesheet));

        let mut table_names: Option<Vec<String>> = None;
        if load_tables {
            sheets.load_tables()?;
            table_names = Some(sheets.table_names()?);
        }
        Ok(Self {
            path: Some(path.to_string()),
            workbook_type: WorkbookType::from(&sheets),
            stream: None,
            stream_seq: 0,
            stylesheet,
            rich_strings: None,
            meta,
            sheets,
            sheets_metadata,
            sheet_names,
            table_names,
        })
    }

    fn get_sheet_by_name(&mut self, name: &str) -> PyResult<CalamineSheet> {
        self.stream = None;
        let range = self.sheets.worksheet_range(name)?;
        let merge_cells_range = self.sheets.worksheet_merge_cells(name)?;
        Ok(CalamineSheet::new(
            name.to_owned(),
            range,
            merge_cells_range,
        ))
    }

    fn meta(&self) -> PyResult<&WorkbookMeta> {
        self.meta.as_ref().map_err(|e| {
            crate::CalamineError::new_err(format!("cannot read workbook metadata: {e}"))
        })
    }

    fn sheet_name_by_index(&self, index: usize) -> PyResult<String> {
        Ok(self
            .sheet_names
            .get(index)
            .ok_or_else(|| WorksheetNotFound::new_err(format!("Worksheet '{index}' not found")))?
            .to_string())
    }

    fn get_sheet_by_index(&mut self, index: usize) -> PyResult<CalamineSheet> {
        let name = self.sheet_name_by_index(index)?;
        self.get_sheet_by_name(&name)
    }

    /// Starts streaming `name`, replacing (invalidating) any active stream.
    fn start_stream(
        &mut self,
        name: &str,
        rich: bool,
        range: StreamRange,
    ) -> Result<(u64, Option<RichInfo>), Error> {
        self.stream = None;
        if !matches!(self.workbook_type, WorkbookType::Xlsx) {
            if let SheetsEnum::None = self.sheets {
                return Err(Error::WorkbookClosed);
            }
            return Err(Error::StreamingNotSupported(
                "streaming is only supported for xlsx / xlsm (and xltx / xltm) workbooks",
            ));
        }
        // Everything `rich` needs is read up front: the cell reader below keeps the
        // zip archive borrowed until the stream ends. Merged ranges, hyperlinks,
        // conditional formats, ... are stored after the cells in the sheet XML, so
        // they cost one extra pass over the sheet. Styles and rich text runs are
        // workbook-wide and loaded once.
        let info = if rich {
            let info = self.sheets.worksheet_info(name)?;
            let comments = self.sheets.worksheet_comments(name)?;
            let stylesheet = match &self.stylesheet {
                Some(s) => Arc::clone(s),
                None => {
                    let s = Arc::new(self.sheets.stylesheet()?);
                    self.stylesheet = Some(Arc::clone(&s));
                    s
                }
            };
            let rich_strings = match &self.rich_strings {
                Some(s) => Arc::clone(s),
                None => {
                    let s = Arc::new(self.sheets.rich_shared_strings()?);
                    self.rich_strings = Some(Arc::clone(&s));
                    s
                }
            };
            let sheet_name = |n: &&DefinedName| n.sheet.as_deref() == Some(name);
            let names = self
                .meta
                .as_ref()
                .map(|m| m.names.as_slice())
                .unwrap_or(&[]);
            let print_area = names
                .iter()
                .filter(sheet_name)
                .find(|n| n.name == "_xlnm.Print_Area")
                .map(|n| n.value.clone());
            let print_titles = names
                .iter()
                .filter(sheet_name)
                .find(|n| n.name == "_xlnm.Print_Titles")
                .map(|n| n.value.clone());
            Some(RichInfo {
                info,
                comments,
                stylesheet,
                rich_strings,
                print_area,
                print_titles,
            })
        } else {
            None
        };
        // SAFETY (both transmutes): the reader borrows the `Xlsx` inside `self.sheets`,
        // which lives in this Python object and never moves. `self.stream` is cleared
        // before `self.sheets` is used, replaced or dropped (see the `stream` field).
        let reader = match &mut self.sheets {
            SheetsEnum::None => return Err(Error::WorkbookClosed),
            SheetsEnum::File(Sheets::Xlsx(x)) => cells_reader(x, name)?.map(|r| {
                CellReader::File(unsafe {
                    std::mem::transmute::<
                        XlsxCellReader<'_, BufReader<File>>,
                        XlsxCellReader<'static, BufReader<File>>,
                    >(r)
                })
            }),
            SheetsEnum::FileLike(Sheets::Xlsx(x)) => cells_reader(x, name)?.map(|r| {
                CellReader::FileLike(unsafe {
                    std::mem::transmute::<
                        XlsxCellReader<'_, Cursor<Vec<u8>>>,
                        XlsxCellReader<'static, Cursor<Vec<u8>>>,
                    >(r)
                })
            }),
            _ => {
                return Err(Error::StreamingNotSupported(
                    "streaming is only supported for xlsx / xlsm (and xltx / xltm) workbooks",
                ))
            }
        };
        self.stream_seq += 1;
        self.stream = Some(StreamState::new(self.stream_seq, reader, rich, range));
        Ok((self.stream_seq, info))
    }

    /// Advances stream `id`. Returns `None` once it is exhausted.
    pub(crate) fn next_stream_row(&mut self, id: u64) -> PyResult<Option<StreamRow>> {
        let Some(stream) = self.stream.as_mut().filter(|s| s.id == id) else {
            return Err(match self.sheets {
                SheetsEnum::None => Error::WorkbookClosed,
                _ => Error::StreamInvalidated,
            }
            .into());
        };
        let row = stream.next_row();
        if !matches!(row, Ok(Some(_))) {
            self.stream = None;
        }
        Ok(row?)
    }

    pub(crate) fn end_stream(&mut self, id: u64) {
        if self.stream.as_ref().is_some_and(|s| s.id == id) {
            self.stream = None;
        }
    }

    fn get_table_by_name(&mut self, name: &str) -> PyResult<CalamineTable> {
        self.stream = None;
        match &self.workbook_type {
            WorkbookType::Xlsx => match &self.table_names {
                Some(_) => Ok(self.sheets.get_table_by_name(name)?),
                None => Err(Error::TablesNotLoaded.into()),
            },
            _ => Err(Error::TablesNotSupported.into()),
        }
    }
}
