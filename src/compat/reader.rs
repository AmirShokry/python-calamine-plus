//! Independent worksheet readers and openpyxl's cell value rules.
//!
//! Every iterator (and every worksheet's random-access cursor) owns its own handle on
//! the archive, so iterators never invalidate each other. They share the workbook's
//! shared strings, number formats and converted styles through `Shared`.

use std::fs::File;
use std::io::{BufReader, Cursor};
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};

use calamine::{
    DataRef, Error as CalamineCrateError, Xlsx, XlsxArchive, XlsxCellContext, XlsxCellReader,
    XlsxError,
};
use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};
use pyo3::prelude::*;
use pyo3::sync::PyOnceLock;

use super::objects::{ArrayFormula, DataTableFormula};
use super::styles::{CompatStyle, StylePy};
use super::utils::{from_excel, from_iso8601, ExcelDate, IsoValue};
use crate::Error;

/// Bytes of a workbook opened from a file-like object, shared by all readers.
#[derive(Clone)]
pub(crate) struct SharedBytes(pub(crate) Arc<Vec<u8>>);

impl AsRef<[u8]> for SharedBytes {
    fn as_ref(&self) -> &[u8] {
        self.0.as_slice()
    }
}

pub(crate) enum Source {
    Path(PathBuf),
    Bytes(SharedBytes),
}

/// The workbook's own reader, used for metadata (styles, comments, sheet info).
pub(crate) enum MainXlsx {
    File(Xlsx<BufReader<File>>),
    Bytes(Xlsx<Cursor<SharedBytes>>),
}

macro_rules! with_main {
    ($main:expr, $x:ident => $body:expr) => {
        match $main {
            $crate::compat::reader::MainXlsx::File($x) => $body,
            $crate::compat::reader::MainXlsx::Bytes($x) => $body,
        }
    };
}
pub(crate) use with_main;

/// Style id of cells openpyxl creates (not stored in the file) and of merged
/// placeholders: openpyxl's all-zero style array (see `StyleSheet::zero_style`).
pub(crate) const ZERO_STYLE: u32 = u32::MAX;

/// Opt-in extensions of the openpyxl profile (`load_workbook` keywords).
#[derive(Clone, Copy, Debug, Default)]
pub(crate) struct CompatOptions {
    /// `formula_and_value=True`: cells also carry `formula` and `cached_value`.
    pub(crate) formula_and_value: bool,
    /// `read_comments` / `read_hyperlinks` / `read_merged_cells` (read-only mode).
    pub(crate) comments: bool,
    pub(crate) hyperlinks: bool,
    pub(crate) merged: bool,
}

impl CompatOptions {
    /// Read-only cells need sheet metadata.
    pub(crate) fn any_metadata(&self) -> bool {
        self.comments || self.hyperlinks || self.merged
    }
}

/// Workbook-wide state shared by the workbook, its worksheets and all iterators.
pub(crate) struct Shared {
    pub(crate) source: Source,
    pub(crate) ctx: XlsxCellContext,
    pub(crate) styles: Vec<CompatStyle>,
    style_py: Vec<PyOnceLock<StylePy>>,
    zero: CompatStyle,
    zero_py: PyOnceLock<StylePy>,
    pub(crate) epoch: NaiveDateTime,
    pub(crate) data_only: bool,
    pub(crate) read_only: bool,
    pub(crate) opts: CompatOptions,
    closed: AtomicBool,
    pub(crate) main: Mutex<Option<MainXlsx>>,
}

pub(crate) fn xlsx_err(e: XlsxError) -> Error {
    Error::Calamine(CalamineCrateError::Xlsx(e))
}

impl Shared {
    #[allow(clippy::too_many_arguments)] // plain constructor
    pub(crate) fn new(
        source: Source,
        ctx: XlsxCellContext,
        styles: Vec<CompatStyle>,
        zero: CompatStyle,
        epoch: NaiveDateTime,
        data_only: bool,
        read_only: bool,
        opts: CompatOptions,
        main: MainXlsx,
    ) -> Self {
        let style_py = (0..styles.len()).map(|_| PyOnceLock::new()).collect();
        Self {
            source,
            ctx,
            styles,
            style_py,
            zero,
            zero_py: PyOnceLock::new(),
            epoch,
            data_only,
            read_only,
            opts,
            closed: AtomicBool::new(false),
            main: Mutex::new(Some(main)),
        }
    }

    pub(crate) fn is_closed(&self) -> bool {
        self.closed.load(Ordering::Acquire)
    }

    pub(crate) fn check_open(&self) -> Result<(), Error> {
        if self.is_closed() {
            Err(Error::WorkbookClosed)
        } else {
            Ok(())
        }
    }

    pub(crate) fn close(&self) -> bool {
        let was_open = !self.closed.swap(true, Ordering::AcqRel);
        *self.main.lock().expect("not poisoned") = None;
        was_open
    }

    /// The converted format of a cell's style id (`ZERO_STYLE` for cells the file does
    /// not store; the first format for unknown ids).
    pub(crate) fn style(&self, id: u32) -> &CompatStyle {
        if id == ZERO_STYLE {
            return &self.zero;
        }
        self.styles.get(id as usize).unwrap_or(&self.styles[0])
    }

    /// Python objects of a style id, built on first use.
    pub(crate) fn style_py(&self, py: Python<'_>, id: u32) -> PyResult<&StylePy> {
        if id == ZERO_STYLE {
            return self
                .zero_py
                .get_or_try_init(py, || StylePy::new(py, &self.zero));
        }
        let id = if (id as usize) < self.styles.len() {
            id as usize
        } else {
            0
        };
        self.style_py[id].get_or_try_init(py, || StylePy::new(py, &self.styles[id]))
    }

    /// Runs `f` with the workbook's metadata reader.
    pub(crate) fn with_main<T>(
        &self,
        f: impl FnOnce(&mut MainXlsx) -> Result<T, XlsxError>,
    ) -> Result<T, Error> {
        let mut guard = self.main.lock().expect("not poisoned");
        let main = guard.as_mut().ok_or(Error::WorkbookClosed)?;
        f(main).map_err(xlsx_err)
    }
}

/// A cell value as openpyxl reads it.
#[derive(Clone, Debug)]
pub(crate) enum CVal {
    None,
    Int(i64),
    Float(f64),
    Str(String),
    Bool(bool),
    DateTime(NaiveDateTime),
    Date(NaiveDate),
    Time(NaiveTime),
    Delta(Duration),
    /// Formula text with the leading `=`.
    Formula(String),
    Array {
        reference: Option<String>,
        text: String,
    },
    DataTable(Vec<(String, String)>),
}

impl CVal {
    pub(crate) fn is_none(&self) -> bool {
        matches!(self, CVal::None)
    }

    pub(crate) fn to_py(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        Ok(match self {
            CVal::None => py.None(),
            CVal::Int(v) => v.into_pyobject(py)?.into_any().unbind(),
            CVal::Float(v) => v.into_pyobject(py)?.into_any().unbind(),
            CVal::Str(v) | CVal::Formula(v) => v.into_pyobject(py)?.into_any().unbind(),
            CVal::Bool(v) => v.into_pyobject(py)?.to_owned().into_any().unbind(),
            CVal::DateTime(v) => v.into_pyobject(py)?.into_any().unbind(),
            CVal::Date(v) => v.into_pyobject(py)?.into_any().unbind(),
            CVal::Time(v) => v.into_pyobject(py)?.into_any().unbind(),
            CVal::Delta(v) => v.into_pyobject(py)?.into_any().unbind(),
            CVal::Array { reference, text } => Py::new(
                py,
                ArrayFormula {
                    reference: reference.clone(),
                    text: text.clone(),
                },
            )?
            .into_any(),
            CVal::DataTable(attrs) => Py::new(py, DataTableFormula::new(attrs.clone()))?.into_any(),
        })
    }
}

/// One stored cell (1-based column).
#[derive(Clone, Debug)]
pub(crate) struct CCell {
    pub(crate) col: u32,
    pub(crate) value: CVal,
    pub(crate) data_type: &'static str,
    pub(crate) style: u32,
    /// With `formula_and_value=True`: the formula and the saved result.
    pub(crate) extra: Option<Box<CExtra>>,
}

/// Formula and saved result of a cell, read in the same pass.
#[derive(Clone, Debug)]
pub(crate) struct CExtra {
    /// What openpyxl's formula view returns (`None` for cells without a formula).
    pub(crate) formula: CVal,
    /// What `data_only=True` returns.
    pub(crate) cached: CVal,
}

/// One stored row (1-based index), cells sorted by column.
#[derive(Clone, Debug, Default)]
pub(crate) struct CRow {
    pub(crate) index: u32,
    pub(crate) cells: Vec<CCell>,
    /// Column of the last cell in file order (openpyxl's read-only row width).
    pub(crate) last_col_in_file: u32,
}

impl CRow {
    pub(crate) fn get(&self, col: u32) -> Option<&CCell> {
        self.cells
            .binary_search_by_key(&col, |c| c.col)
            .ok()
            .map(|i| &self.cells[i])
    }
}

/// A raw cell from the XML before it is grouped into rows.
struct RawCell {
    row: u32,
    cell: CCell,
}

enum ReaderKind {
    File(XlsxCellReader<'static, BufReader<File>>),
    Bytes(XlsxCellReader<'static, Cursor<SharedBytes>>),
}

/// Keeps a reader's archive alive (the reader borrows it).
#[allow(dead_code)]
enum ArchiveBox {
    File(Box<XlsxArchive<BufReader<File>>>),
    Bytes(Box<XlsxArchive<Cursor<SharedBytes>>>),
}

macro_rules! each_reader {
    ($r:expr, $x:ident => $body:expr) => {
        match $r {
            ReaderKind::File($x) => $body,
            ReaderKind::Bytes($x) => $body,
        }
    };
}

/// openpyxl's `parse_cell`: value and data type of the cell last read by `r`.
macro_rules! convert_cell {
    ($r:expr, $value:expr, $formula:expr, $shared:expr) => {{
        let t = $r.last_cell_type();
        let style = $r.last_style_id();
        let formula_kind = $r.last_formula_kind().map(str::to_string);
        let attrs = $r.last_formula_attributes().to_vec();
        let int_literal = $r.last_value_is_int_literal();
        let has_is = $r.last_cell_has_inline_string();
        to_compat(
            $shared,
            t,
            style,
            $value,
            formula_kind,
            $formula,
            attrs,
            int_literal,
            has_is,
        )
    }};
}

#[allow(clippy::too_many_arguments)]
fn to_compat(
    shared: &Shared,
    t: &'static str,
    style: u32,
    value: DataRef<'_>,
    formula_kind: Option<String>,
    formula: Option<String>,
    attrs: Vec<(String, String)>,
    int_literal: bool,
    has_is: bool,
) -> Result<(CVal, &'static str, Option<Box<CExtra>>), Error> {
    let formula_value = formula_kind.map(|kind| {
        let text = format!("={}", formula.unwrap_or_default());
        match kind.as_str() {
            "array" => CVal::Array {
                reference: attrs
                    .iter()
                    .find(|(k, _)| k == "ref")
                    .map(|(_, v)| v.clone()),
                text,
            },
            "dataTable" => CVal::DataTable(attrs),
            _ => CVal::Formula(text),
        }
    });
    let with_extra = shared.opts.formula_and_value;
    if !shared.data_only && !with_extra {
        if let Some(f) = formula_value {
            return Ok((f, "f", None));
        }
    }
    let (cached, cached_type) = saved_value(shared, t, style, value, int_literal, has_is)?;
    let extra = with_extra.then(|| {
        Box::new(CExtra {
            formula: formula_value.clone().unwrap_or(CVal::None),
            cached: cached.clone(),
        })
    });
    match formula_value {
        Some(f) if !shared.data_only => Ok((f, "f", extra)),
        _ => Ok((cached, cached_type, extra)),
    }
}

/// openpyxl's value of a cell without its formula (the `data_only=True` view).
fn saved_value(
    shared: &Shared,
    t: &'static str,
    style: u32,
    value: DataRef<'_>,
    int_literal: bool,
    has_is: bool,
) -> Result<(CVal, &'static str), Error> {
    let number = |v: f64| -> CVal {
        if int_literal && v.fract() == 0.0 && v.abs() < 9.007_199_254_740_992e15 {
            CVal::Int(v as i64)
        } else {
            CVal::Float(v)
        }
    };
    Ok(match (t, value) {
        ("inlineStr", DataRef::String(s)) => (CVal::Str(s), "s"),
        // `<c t="inlineStr"/>` without `<is>` has no value (openpyxl keeps the type).
        ("inlineStr", _) if has_is => (CVal::Str(String::new()), "s"),
        ("inlineStr", _) => (CVal::None, "inlineStr"),
        (t, DataRef::Empty) => (CVal::None, t),
        ("n", v) => {
            let n = match v {
                DataRef::Int(i) => CVal::Int(i),
                DataRef::Float(f) => number(f),
                DataRef::DateTime(d) => number(d.as_f64()),
                DataRef::String(s) => return Ok((CVal::Str(s), "n")),
                DataRef::SharedString(s) => return Ok((CVal::Str(s.to_string()), "n")),
                other => return Ok((CVal::Str(format!("{other:?}")), "n")),
            };
            let s = shared.style(style);
            if !s.is_date {
                (n, "n")
            } else {
                let serial = match n {
                    CVal::Int(i) => i as f64,
                    CVal::Float(f) => f,
                    _ => unreachable!(),
                };
                match from_excel(serial, shared.epoch, s.is_timedelta) {
                    Some(ExcelDate::DateTime(dt)) => (CVal::DateTime(dt), "d"),
                    Some(ExcelDate::Time(t)) => (CVal::Time(t), "d"),
                    Some(ExcelDate::Delta(d)) => (CVal::Delta(d), "d"),
                    // openpyxl marks out-of-range date serials as errors.
                    None => (CVal::Str("#VALUE!".into()), "e"),
                }
            }
        }
        ("s", DataRef::SharedString(s)) => (CVal::Str(s.to_string()), "s"),
        ("s", DataRef::String(s)) => (CVal::Str(s), "s"),
        ("b", DataRef::Bool(b)) => (CVal::Bool(b), "b"),
        // openpyxl treats an empty `<v></v>` as no value and keeps the type.
        ("str", DataRef::String(s)) if s.is_empty() => (CVal::None, "str"),
        ("str", DataRef::String(s)) => (CVal::Str(s), "s"),
        ("e", DataRef::Error(e)) => (CVal::Str(e.to_string()), "e"),
        ("d", DataRef::DateTimeIso(s)) => match from_iso8601(&s) {
            Ok(None) => (CVal::None, "d"),
            Ok(Some(IsoValue::Date(d))) => (CVal::Date(d), "d"),
            Ok(Some(IsoValue::Time(t))) => (CVal::Time(t), "d"),
            Ok(Some(IsoValue::DateTime(dt))) => (CVal::DateTime(dt), "d"),
            Ok(Some(IsoValue::Delta(d))) => (CVal::Delta(d), "d"),
            Err(msg) => return Err(Error::Compat(msg)),
        },
        (t, DataRef::SharedString(s)) => (CVal::Str(s.to_string()), t),
        (t, DataRef::String(s)) => (CVal::Str(s), t),
        (t, DataRef::Bool(b)) => (CVal::Bool(b), t),
        (t, DataRef::Int(i)) => (CVal::Int(i), t),
        (t, DataRef::Float(f)) => (CVal::Float(f), t),
        (t, DataRef::Error(e)) => (CVal::Str(e.to_string()), t),
        (t, other) => (CVal::Str(format!("{other:?}")), t),
    })
}

/// A forward-only reader over one worksheet with its own archive handle.
pub(crate) struct SheetReader {
    // Declared before the archive and `shared`, so it is dropped first: it borrows both.
    reader: Option<ReaderKind>,
    _archive: Option<ArchiveBox>,
    shared: Arc<Shared>,
    pending: Option<RawCell>,
    eof: bool,
}

impl SheetReader {
    /// Opens worksheet `part` (`None`: not a worksheet; the reader is empty).
    pub(crate) fn open(shared: &Arc<Shared>, part: Option<&str>) -> Result<Self, Error> {
        shared.check_open()?;
        let mut out = Self {
            reader: None,
            _archive: None,
            shared: Arc::clone(shared),
            pending: None,
            eof: false,
        };
        let Some(part) = part else { return Ok(out) };
        // SAFETY: the reader borrows the boxed archive and `shared.ctx` (behind an
        // `Arc`); both live in `out` at stable heap addresses and outlive the reader,
        // which is declared (so dropped) first.
        let ctx: &'static XlsxCellContext =
            unsafe { &*(&out.shared.ctx as *const XlsxCellContext) };
        match &shared.source {
            Source::Path(p) => {
                let file = File::open(p).map_err(|e| Error::Calamine(CalamineCrateError::Io(e)))?;
                let mut archive =
                    Box::new(XlsxArchive::new(BufReader::new(file)).map_err(xlsx_err)?);
                let archive_ref: &'static mut XlsxArchive<BufReader<File>> =
                    unsafe { &mut *(archive.as_mut() as *mut _) };
                out.reader =
                    open_reader(ctx.cells_reader(archive_ref, part))?.map(ReaderKind::File);
                out._archive = Some(ArchiveBox::File(archive));
            }
            Source::Bytes(b) => {
                let mut archive =
                    Box::new(XlsxArchive::new(Cursor::new(b.clone())).map_err(xlsx_err)?);
                let archive_ref: &'static mut XlsxArchive<Cursor<SharedBytes>> =
                    unsafe { &mut *(archive.as_mut() as *mut _) };
                out.reader =
                    open_reader(ctx.cells_reader(archive_ref, part))?.map(ReaderKind::Bytes);
                out._archive = Some(ArchiveBox::Bytes(archive));
            }
        }
        Ok(out)
    }

    /// The `<dimension>` the sheet declares, as 1-based `(min_row, min_col, max_row, max_col)`.
    pub(crate) fn declared_dimension(&self) -> Option<(u32, u32, u32, u32)> {
        let r = self.reader.as_ref()?;
        let (has, d) = each_reader!(r, x => (x.has_dimension(), x.dimensions()));
        has.then(|| (d.start.0 + 1, d.start.1 + 1, d.end.0 + 1, d.end.1 + 1))
    }

    /// 1-based index of the last `<row>` element read (rows without cells included).
    pub(crate) fn last_row_element(&self) -> Option<u32> {
        let r = self.reader.as_ref()?;
        each_reader!(r, x => x.last_row_element()).map(|i| i + 1)
    }

    fn next_cell(&mut self) -> Result<Option<RawCell>, Error> {
        if let Some(c) = self.pending.take() {
            return Ok(Some(c));
        }
        if self.eof {
            return Ok(None);
        }
        let Some(reader) = self.reader.as_mut() else {
            self.eof = true;
            return Ok(None);
        };
        let shared = &*self.shared;
        let raw = if shared.data_only && !shared.opts.formula_and_value {
            each_reader!(reader, r => match r.next_cell().map_err(xlsx_err)? {
                Some(c) => {
                    let pos = c.get_position();
                    let value = c.get_value().clone();
                    Some((pos, convert_cell!(r, value, None, shared)?))
                }
                None => None,
            })
        } else {
            each_reader!(reader, r => match r.next_cell_with_formula().map_err(xlsx_err)? {
                Some(c) => {
                    let pos = c.pos;
                    Some((pos, convert_cell!(r, c.value, c.formula, shared)?))
                }
                None => None,
            })
        };
        Ok(match raw {
            None => {
                self.eof = true;
                None
            }
            Some(((row, col), (value, data_type, extra))) => {
                let style = each_reader!(reader, r => r.last_style_id());
                Some(RawCell {
                    row: row + 1,
                    cell: CCell {
                        col: col + 1,
                        value,
                        data_type,
                        style,
                        extra,
                    },
                })
            }
        })
    }

    /// The next stored row that has cells, or `None` at the end of the sheet.
    pub(crate) fn next_row(&mut self) -> Result<Option<CRow>, Error> {
        self.shared.check_open()?;
        let Some(first) = self.next_cell()? else {
            return Ok(None);
        };
        let mut row = CRow {
            index: first.row,
            last_col_in_file: first.cell.col,
            cells: vec![first.cell],
        };
        let mut sorted = true;
        while let Some(c) = self.next_cell()? {
            if c.row == row.index {
                sorted &= row.cells.last().is_some_and(|l| l.col < c.cell.col);
                row.last_col_in_file = c.cell.col;
                row.cells.push(c.cell);
            } else if c.row > row.index {
                self.pending = Some(c);
                break;
            } else {
                return Err(Error::Compat(
                    "rows are not stored in order; this sheet cannot be read with compatibility='openpyxl'"
                        .into(),
                ));
            }
        }
        if !sorted {
            // Later cells win, like openpyxl's cell dictionary.
            let mut cells: Vec<CCell> = Vec::with_capacity(row.cells.len());
            for c in row.cells.drain(..) {
                match cells.binary_search_by_key(&c.col, |x| x.col) {
                    Ok(i) => cells[i] = c,
                    Err(i) => cells.insert(i, c),
                }
            }
            row.cells = cells;
        }
        Ok(Some(row))
    }
}

fn open_reader<R>(r: Result<R, XlsxError>) -> Result<Option<R>, Error> {
    match r {
        Ok(r) => Ok(Some(r)),
        Err(XlsxError::NotAWorksheet(_)) => Ok(None),
        Err(e) => Err(xlsx_err(e)),
    }
}
