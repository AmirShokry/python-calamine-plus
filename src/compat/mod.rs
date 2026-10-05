//! The openpyxl compatibility profile: `load_workbook(..., compatibility="openpyxl")`
//! returns an openpyxl-shaped, read-only workbook, and openpyxl's utilities
//! (`get_column_letter`, `range_boundaries`, `from_excel`, ...) are implemented here.
//! Nothing in this module changes the native `CalamineWorkbook` API.

mod objects;
pub(crate) mod reader;
mod styles;
pub(crate) mod utils;
mod workbook;
mod worksheet;

use pyo3::prelude::*;
pub(crate) use reader::CompatOptions;
pub(crate) use workbook::{CompatibilityNotSupported, Workbook};

pub(crate) fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    utils::register(m)?;
    styles::register(m)?;
    objects::register(m)?;
    worksheet::register(m)?;
    workbook::register(m)?;
    Ok(())
}
