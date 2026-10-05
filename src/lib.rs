use pyo3::exceptions::{PyTypeError, PyValueError};
use pyo3::prelude::*;

mod compat;
mod types;
use crate::types::{
    CalamineCell, CalamineError, CalamineSheet, CalamineSheetStream, CalamineTable,
    CalamineWorkbook, CellValue, Error, PasswordError, SheetMetadata, SheetTypeEnum,
    SheetVisibleEnum, StreamInvalidated, StreamingNotSupported, TableNotFound, TablesNotLoaded,
    TablesNotSupported, WorkbookClosed, WorksheetNotFound, XmlError, ZipError,
};

/// Opens a workbook. Without `compatibility` this is `CalamineWorkbook.from_object`;
/// `compatibility="openpyxl"` returns an openpyxl-shaped read-only workbook and
/// accepts openpyxl's `read_only` / `data_only` (other openpyxl options only affect
/// saving and are rejected unless left at their defaults).
#[pyfunction]
#[pyo3(signature = (
    path_or_filelike, load_tables=false, *, compatibility=None, read_only=None,
    data_only=None, keep_vba=None, keep_links=None, rich_text=None, formula_and_value=None,
    read_comments=None, read_hyperlinks=None, read_merged_cells=None
))]
#[allow(clippy::too_many_arguments)] // the Python keyword signature
fn load_workbook(
    py: Python,
    path_or_filelike: Py<PyAny>,
    load_tables: bool,
    compatibility: Option<&str>,
    read_only: Option<bool>,
    data_only: Option<bool>,
    keep_vba: Option<bool>,
    keep_links: Option<bool>,
    rich_text: Option<bool>,
    formula_and_value: Option<bool>,
    read_comments: Option<bool>,
    read_hyperlinks: Option<bool>,
    read_merged_cells: Option<bool>,
) -> PyResult<Py<PyAny>> {
    let openpyxl_options = [
        ("read_only", read_only.is_some()),
        ("data_only", data_only.is_some()),
        ("keep_vba", keep_vba.is_some()),
        ("keep_links", keep_links.is_some()),
        ("rich_text", rich_text.is_some()),
        ("formula_and_value", formula_and_value.is_some()),
        ("read_comments", read_comments.is_some()),
        ("read_hyperlinks", read_hyperlinks.is_some()),
        ("read_merged_cells", read_merged_cells.is_some()),
    ];
    match compatibility {
        None => {
            if let Some((name, _)) = openpyxl_options.iter().find(|(_, set)| *set) {
                return Err(PyTypeError::new_err(format!(
                    "load_workbook() option '{name}' requires compatibility=\"openpyxl\""
                )));
            }
            Ok(
                CalamineWorkbook::from_object(py, path_or_filelike, load_tables)?
                    .into_pyobject(py)?
                    .into_any()
                    .unbind(),
            )
        }
        Some("openpyxl") => {
            if load_tables {
                return Err(PyValueError::new_err(
                    "load_tables is not used with compatibility=\"openpyxl\"; tables are in ws.tables",
                ));
            }
            if keep_vba == Some(true) {
                return Err(PyValueError::new_err(
                    "keep_vba=True only matters when saving; compatibility=\"openpyxl\" workbooks are read-only",
                ));
            }
            if keep_links == Some(false) {
                return Err(PyValueError::new_err(
                    "keep_links=False only matters when saving; compatibility=\"openpyxl\" workbooks are read-only",
                ));
            }
            if rich_text == Some(true) {
                return Err(compat::CompatibilityNotSupported::new_err(
                    "rich_text=True (CellRichText values) is not supported with \
                     compatibility=\"openpyxl\"; rich text runs are available from \
                     CalamineWorkbook.stream_sheet_by_name(..., rich=True)",
                ));
            }
            let read_only = read_only.unwrap_or(false);
            // Normal mode always binds comments, hyperlinks and merged cells (openpyxl);
            // read-only mode only when asked (like the native rich stream).
            let metadata = [
                ("read_comments", read_comments),
                ("read_hyperlinks", read_hyperlinks),
                ("read_merged_cells", read_merged_cells),
            ];
            if !read_only {
                if let Some((name, _)) = metadata.iter().find(|(_, v)| *v == Some(false)) {
                    return Err(PyValueError::new_err(format!(
                        "{name}=False only applies to read_only=True; normal mode always reads                          comments, hyperlinks and merged cells, like openpyxl"
                    )));
                }
            }
            let opts = compat::CompatOptions {
                formula_and_value: formula_and_value.unwrap_or(false),
                comments: read_only && read_comments.unwrap_or(false),
                hyperlinks: read_only && read_hyperlinks.unwrap_or(false),
                merged: read_only && read_merged_cells.unwrap_or(false),
            };
            let wb = compat::Workbook::load(
                py,
                path_or_filelike,
                read_only,
                data_only.unwrap_or(false),
                opts,
            )?;
            Ok(Py::new(py, wb)?.into_any())
        }
        Some(other) => Err(PyValueError::new_err(format!(
            "unknown compatibility {other:?}; the only supported value is \"openpyxl\""
        ))),
    }
}

#[pymodule]
fn _python_calamine(py: Python, m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(load_workbook, m)?)?;
    m.add_class::<CalamineWorkbook>()?;
    m.add_class::<CalamineSheet>()?;
    m.add_class::<CalamineSheetStream>()?;
    m.add_class::<CalamineCell>()?;
    m.add_class::<SheetMetadata>()?;
    m.add_class::<SheetTypeEnum>()?;
    m.add_class::<SheetVisibleEnum>()?;
    m.add_class::<CalamineTable>()?;
    m.add("CalamineError", py.get_type::<CalamineError>())?;
    m.add("PasswordError", py.get_type::<PasswordError>())?;
    m.add("WorksheetNotFound", py.get_type::<WorksheetNotFound>())?;
    m.add("XmlError", py.get_type::<XmlError>())?;
    m.add("ZipError", py.get_type::<ZipError>())?;
    m.add("TablesNotSupported", py.get_type::<TablesNotSupported>())?;
    m.add("TablesNotLoaded", py.get_type::<TablesNotLoaded>())?;
    m.add("TableNotFound", py.get_type::<TableNotFound>())?;
    m.add("WorkbookClosed", py.get_type::<WorkbookClosed>())?;
    m.add(
        "StreamingNotSupported",
        py.get_type::<StreamingNotSupported>(),
    )?;
    m.add("StreamInvalidated", py.get_type::<StreamInvalidated>())?;
    compat::register(m)?;
    Ok(())
}
