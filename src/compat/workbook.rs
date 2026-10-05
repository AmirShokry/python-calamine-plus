//! `load_workbook(..., compatibility="openpyxl")`: an openpyxl-shaped, read-only workbook.

use std::fs::File;
use std::io::{BufReader, Cursor, Read};
use std::path::{Path, PathBuf};
use std::sync::Arc;

use calamine::{
    open_workbook_auto_from_rs, DefinedName as RawName, Reader, SheetType, SheetVisible, Sheets,
    StyleSheet, WorkbookInfo, Xlsx,
};
use pyo3::exceptions::{PyKeyError, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList};
use pyo3::{create_exception, PyTraverseError, PyVisit};
use pyo3_file::PyFileLikeObject;

use super::objects::{CustomProperty, DefinedName, DocumentProperties};
use super::reader::{with_main, xlsx_err, CompatOptions, MainXlsx, Shared, SharedBytes, Source};
use super::styles::CompatStyle;
use super::utils::{mac_epoch, windows_epoch};
use super::worksheet::{Chartsheet, SheetNames, Worksheet};
use crate::types::convert::{custom_properties_to_py, parse_w3c_datetime};
use crate::{CalamineError, Error};

create_exception!(
    python_calamine_plus,
    CompatibilityNotSupported,
    CalamineError,
    "The file or option cannot be read with `compatibility=\"openpyxl\"`."
);

/// Everything read from the workbook when it is opened.
struct Opened {
    main: MainXlsx,
    ctx: calamine::XlsxCellContext,
    sheets: Vec<(String, SheetType, SheetVisible, Option<String>)>,
    names: Vec<RawName>,
    props: Vec<(String, String)>,
    info: WorkbookInfo,
    stylesheet: StyleSheet,
}

fn open_main(source: &Source) -> Result<Opened, Error> {
    let mut main = match source {
        Source::Path(p) => {
            let f = File::open(p).map_err(|e| Error::Calamine(calamine::Error::Io(e)))?;
            MainXlsx::File(Xlsx::new(BufReader::new(f)).map_err(xlsx_err)?)
        }
        Source::Bytes(b) => MainXlsx::Bytes(Xlsx::new(Cursor::new(b.clone())).map_err(xlsx_err)?),
    };
    let (sheets, names, props, info, stylesheet, ctx) = with_main!(&mut main, x => {
        let sheets = x
            .sheets_metadata()
            .iter()
            .map(|s| (s.name.clone(), s.typ, s.visible, x.worksheet_part(&s.name).ok()))
            .collect::<Vec<_>>();
        let names = x.defined_names_scoped().map_err(xlsx_err)?;
        let props = x.document_properties().map_err(xlsx_err)?;
        let info = x.workbook_info().map_err(xlsx_err)?;
        let stylesheet = x.stylesheet().map_err(xlsx_err)?;
        let ctx = x.take_cell_context();
        (sheets, names, props, info, stylesheet, ctx)
    });
    Ok(Opened {
        main,
        ctx,
        sheets,
        names,
        props,
        info,
        stylesheet,
    })
}

/// A clear error for formats this profile does not read (openpyxl reads xlsx only).
fn unsupported_format(source: &Source, err: Error) -> PyErr {
    let other = match source {
        Source::Path(p) => {
            let ext = Path::new(p)
                .extension()
                .and_then(|e| e.to_str())
                .map(str::to_ascii_lowercase);
            matches!(ext.as_deref(), Some("xls" | "xlsb" | "ods" | "xla"))
                || calamine::open_workbook_auto(p)
                    .ok()
                    .is_some_and(|s| !matches!(s, Sheets::Xlsx(_)))
        }
        Source::Bytes(b) => open_workbook_auto_from_rs(Cursor::new(b.clone()))
            .ok()
            .is_some_and(|s| !matches!(s, Sheets::Xlsx(_))),
    };
    if other {
        CompatibilityNotSupported::new_err(
            "compatibility='openpyxl' reads .xlsx / .xlsm / .xltx / .xltm files only \
             (like openpyxl); use CalamineWorkbook for .xls / .xlsb / .ods",
        )
    } else {
        err.into()
    }
}

fn state_name(v: SheetVisible) -> &'static str {
    match v {
        SheetVisible::Visible => "visible",
        SheetVisible::Hidden => "hidden",
        SheetVisible::VeryHidden => "veryHidden",
    }
}

/// A workbook opened with `compatibility="openpyxl"`, like openpyxl's `Workbook`
/// (read-only: nothing can be modified or saved).
#[pyclass(module = "python_calamine_plus.openpyxl.workbook")]
pub struct Workbook {
    shared: Arc<Shared>,
    sheets: Vec<(String, SheetType, SheetVisible, Option<String>)>,
    objects: Vec<Option<Py<PyAny>>>,
    names: Vec<RawName>,
    props: Vec<(String, String)>,
    info: WorkbookInfo,
    named_styles: Vec<String>,
}

impl Workbook {
    pub(crate) fn load(
        py: Python<'_>,
        path_or_filelike: Py<PyAny>,
        read_only: bool,
        data_only: bool,
        opts: CompatOptions,
    ) -> PyResult<Self> {
        let source = if let Ok(p) = path_or_filelike.extract::<PathBuf>(py) {
            Source::Path(p)
        } else {
            let mut buf = Vec::new();
            PyFileLikeObject::with_requirements(path_or_filelike, true, false, true, false)?
                .read_to_end(&mut buf)?;
            Source::Bytes(SharedBytes(Arc::new(buf)))
        };
        let opened = match py.detach(|| open_main(&source)) {
            Ok(o) => o,
            Err(e) => return Err(py.detach(|| unsupported_format(&source, e))),
        };
        let epoch = if opened.ctx.is_1904() {
            mac_epoch()
        } else {
            windows_epoch()
        };
        let styles = if opened.stylesheet.cell_styles.is_empty() {
            vec![CompatStyle::openpyxl_default()]
        } else {
            opened
                .stylesheet
                .cell_styles
                .iter()
                .enumerate()
                .map(|(i, s)| CompatStyle::from_style(s, i))
                .collect()
        };
        let zero = match (
            &opened.stylesheet.zero_style,
            opened.stylesheet.cell_styles.is_empty(),
        ) {
            (Some(z), false) => CompatStyle::zero(z),
            _ => CompatStyle::openpyxl_default(),
        };
        let named_styles = opened
            .stylesheet
            .named_styles
            .iter()
            .map(|s| s.name.clone())
            .collect();
        let shared = Shared::new(
            source,
            opened.ctx,
            styles,
            zero,
            epoch,
            data_only,
            read_only,
            opts,
            opened.main,
        );
        Ok(Self {
            shared: Arc::new(shared),
            objects: opened.sheets.iter().map(|_| None).collect(),
            sheets: opened.sheets,
            names: opened.names,
            props: opened.props,
            info: opened.info,
            named_styles,
        })
    }

    fn sheet_names_for(&self, index: usize) -> SheetNames {
        let title = &self.sheets[index].0;
        let mut out = SheetNames::default();
        for n in self
            .names
            .iter()
            .filter(|n| n.sheet.as_deref() == Some(title.as_str()))
        {
            let dn = self.defined_name(n);
            match dn.reserved() {
                None => out.names.push(dn),
                Some("Print_Area") => out.print_area = Some(n.value.clone()),
                Some("Print_Titles") => out.print_titles = Some(n.value.clone()),
                Some(_) => {}
            }
        }
        out
    }

    fn defined_name(&self, n: &RawName) -> DefinedName {
        DefinedName {
            name: n.name.clone(),
            value: n.value.clone(),
            local_sheet_id: n
                .sheet
                .as_ref()
                .and_then(|s| self.sheets.iter().position(|(t, ..)| t == s)),
            hidden: n.hidden.then_some(true),
            comment: n.comment.clone(),
        }
    }

    /// The (cached) sheet object at `index`.
    fn sheet(slf: &Bound<'_, Self>, index: usize) -> PyResult<Py<PyAny>> {
        let py = slf.py();
        if let Some(obj) = &slf.borrow().objects[index] {
            return Ok(obj.clone_ref(py));
        }
        let wb = slf.borrow();
        let (title, typ, visible, part) = wb.sheets[index].clone();
        let obj = if typ == SheetType::ChartSheet {
            Py::new(
                py,
                Chartsheet {
                    wb: slf.clone().unbind(),
                    title,
                    sheet_state: state_name(visible),
                },
            )?
            .into_any()
        } else {
            Py::new(
                py,
                Worksheet::new(
                    slf.clone().unbind(),
                    Arc::clone(&wb.shared),
                    title,
                    part,
                    state_name(visible),
                    wb.sheet_names_for(index),
                ),
            )?
            .into_any()
        };
        drop(wb);
        slf.borrow_mut().objects[index] = Some(obj.clone_ref(py));
        Ok(obj)
    }

    fn index_of(&self, name: &str) -> Option<usize> {
        self.sheets.iter().position(|(n, ..)| n == name)
    }

    fn prop(&self, key: &str) -> Option<String> {
        self.props
            .iter()
            .find(|(k, _)| k == key)
            .map(|(_, v)| v.clone())
    }
}

#[pymethods]
impl Workbook {
    fn __traverse__(&self, visit: PyVisit<'_>) -> Result<(), PyTraverseError> {
        for obj in self.objects.iter().flatten() {
            visit.call(obj)?;
        }
        Ok(())
    }

    fn __clear__(&mut self) {
        for obj in self.objects.iter_mut() {
            *obj = None;
        }
    }

    fn __repr__(&self) -> String {
        format!(
            "<Workbook sheets={:?} compatibility='openpyxl'>",
            self.sheetnames()
        )
    }

    /// Titles of all sheets (worksheets and chartsheets), in workbook order.
    #[getter]
    fn sheetnames(&self) -> Vec<String> {
        self.sheets.iter().map(|(n, ..)| n.clone()).collect()
    }

    /// Worksheets (chartsheets excluded), in workbook order.
    #[getter]
    fn worksheets(slf: &Bound<'_, Self>) -> PyResult<Vec<Py<PyAny>>> {
        let indexes: Vec<usize> = (0..slf.borrow().sheets.len())
            .filter(|i| slf.borrow().sheets[*i].1 != SheetType::ChartSheet)
            .collect();
        indexes.into_iter().map(|i| Self::sheet(slf, i)).collect()
    }

    #[getter]
    fn chartsheets(slf: &Bound<'_, Self>) -> PyResult<Vec<Py<PyAny>>> {
        let indexes: Vec<usize> = (0..slf.borrow().sheets.len())
            .filter(|i| slf.borrow().sheets[*i].1 == SheetType::ChartSheet)
            .collect();
        indexes.into_iter().map(|i| Self::sheet(slf, i)).collect()
    }

    /// The active sheet (the workbook view's `activeTab`), or `None`.
    #[getter]
    fn active(slf: &Bound<'_, Self>) -> PyResult<Option<Py<PyAny>>> {
        let index = slf
            .borrow()
            .info
            .views
            .first()
            .and_then(|v| v.iter().find(|(k, _)| k == "activeTab"))
            .and_then(|(_, v)| v.parse::<usize>().ok())
            .unwrap_or(0);
        if index >= slf.borrow().sheets.len() {
            return Ok(None);
        }
        Self::sheet(slf, index).map(Some)
    }

    fn __getitem__(slf: &Bound<'_, Self>, key: &str) -> PyResult<Py<PyAny>> {
        let index = slf.borrow().index_of(key);
        match index {
            Some(i) => Self::sheet(slf, i),
            None => Err(PyKeyError::new_err(format!(
                "Worksheet {key} does not exist."
            ))),
        }
    }

    fn __contains__(&self, key: &str) -> bool {
        self.index_of(key).is_some()
    }

    fn __iter__(slf: &Bound<'_, Self>) -> PyResult<Py<PyAny>> {
        let sheets = Self::worksheets(slf)?;
        Ok(PyList::new(slf.py(), sheets)?
            .try_iter()?
            .into_any()
            .unbind())
    }

    /// Position of a sheet in the workbook.
    fn index(slf: &Bound<'_, Self>, worksheet: &Bound<'_, PyAny>) -> PyResult<usize> {
        let title: String = worksheet.getattr("title")?.extract()?;
        let index = slf.borrow().index_of(&title);
        index.ok_or_else(|| PyValueError::new_err(format!("{title} is not in list")))
    }

    #[getter]
    fn read_only(&self) -> bool {
        self.shared.read_only
    }

    #[getter]
    fn data_only(&self) -> bool {
        self.shared.data_only
    }

    /// `datetime(1899, 12, 30)` or, for the 1904 date system, `datetime(1904, 1, 1)`.
    #[getter]
    fn epoch(&self) -> chrono::NaiveDateTime {
        self.shared.epoch
    }

    #[getter]
    fn excel_base_date(&self) -> chrono::NaiveDateTime {
        self.shared.epoch
    }

    #[getter]
    fn template(&self) -> bool {
        self.info.is_template
    }

    #[getter]
    fn code_name(&self) -> Option<String> {
        self.info
            .workbook_properties
            .iter()
            .find(|(k, _)| k == "codeName")
            .map(|(_, v)| v.clone())
    }

    /// Always `None`: VBA projects are not loaded.
    #[getter]
    fn vba_archive(&self) -> Option<Py<PyAny>> {
        None
    }

    /// Raw theme XML, or `None`.
    #[getter]
    fn loaded_theme<'py>(&self, py: Python<'py>) -> Option<Bound<'py, PyBytes>> {
        self.info.theme_xml.as_ref().map(|b| PyBytes::new(py, b))
    }

    /// Core document properties.
    #[getter]
    fn properties(&self) -> DocumentProperties {
        let date = |k: &str| self.prop(k).and_then(|v| parse_w3c_datetime(&v));
        DocumentProperties {
            creator: self.prop("creator"),
            title: self.prop("title"),
            description: self.prop("description"),
            subject: self.prop("subject"),
            identifier: self.prop("identifier"),
            language: self.prop("language"),
            created: date("created"),
            modified: date("modified"),
            last_modified_by: self.prop("last_modified_by"),
            category: self.prop("category"),
            content_status: self.prop("content_status"),
            version: self.prop("version"),
            revision: self.prop("revision"),
            keywords: self.prop("keywords"),
            last_printed: date("last_printed"),
        }
    }

    /// Custom document properties, in file order.
    #[getter]
    fn custom_doc_props(&self, py: Python<'_>) -> PyResult<Vec<Py<CustomProperty>>> {
        let d = custom_properties_to_py(py, &self.info.custom_properties)?;
        d.iter()
            .map(|(k, v)| {
                Py::new(
                    py,
                    CustomProperty {
                        name: k.extract()?,
                        value: v.unbind(),
                    },
                )
            })
            .collect()
    }

    /// Workbook-scoped defined names, by name.
    #[getter]
    fn defined_names<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyDict>> {
        let d = PyDict::new(py);
        for n in self.names.iter().filter(|n| n.sheet.is_none()) {
            d.set_item(&n.name, Py::new(py, self.defined_name(n))?)?;
        }
        Ok(d)
    }

    /// Names of the named cell styles.
    #[getter]
    fn named_styles(&self) -> Vec<String> {
        self.named_styles.clone()
    }

    #[getter]
    fn style_names(&self) -> Vec<String> {
        self.named_styles.clone()
    }

    /// Close the file. Worksheets, cells already read and their values stay usable;
    /// reading more raises `WorkbookClosed`.
    fn close(&mut self, py: Python<'_>) {
        self.shared.close();
        for obj in self.objects.iter().flatten() {
            if let Ok(ws) = obj.bind(py).cast::<Worksheet>() {
                ws.get().release();
            }
        }
    }

    fn __enter__(slf: Py<Self>) -> Py<Self> {
        slf
    }

    fn __exit__(
        &mut self,
        py: Python<'_>,
        _exc_type: Py<PyAny>,
        _exc_value: Py<PyAny>,
        _traceback: Py<PyAny>,
    ) {
        self.close(py);
    }

    /// Diagnostics: whether the workbook is closed.
    #[getter]
    fn _closed(&self) -> bool {
        self.shared.is_closed()
    }

    fn __getattr__(&self, name: &str) -> PyResult<Py<PyAny>> {
        if matches!(
            name,
            "save" | "create_sheet" | "remove" | "copy_worksheet" | "move_sheet"
        ) {
            return Err(PyTypeError::new_err(format!(
                "Workbook.{name} is not available: compatibility='openpyxl' workbooks are read-only"
            )));
        }
        Err(pyo3::exceptions::PyAttributeError::new_err(format!(
            "'Workbook' object has no attribute '{name}'"
        )))
    }
}

pub(crate) fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Workbook>()?;
    m.add(
        "CompatibilityNotSupported",
        m.py().get_type::<CompatibilityNotSupported>(),
    )?;
    Ok(())
}
