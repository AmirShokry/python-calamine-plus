//! openpyxl's coordinate, number format and date helpers, implemented in Rust.
//!
//! These follow `openpyxl.utils.cell`, `openpyxl.utils.datetime` and
//! `openpyxl.styles.numbers` (openpyxl 3.1) and are exposed to Python under the same
//! names, plus a few helpers openpyxl does not have (`coordinate_from_row_col`,
//! `serial_to_datetime`, `split_print_areas`).

use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime, Timelike};
use pyo3::create_exception;
use pyo3::exceptions::{PyException, PyTypeError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{
    PyDate, PyDateTime, PyDelta, PyDict, PyInt, PyIterator, PyList, PyTime, PyTuple,
};

create_exception!(
    python_calamine_plus,
    CellCoordinatesException,
    PyException,
    "Invalid cell coordinates (openpyxl's `CellCoordinatesException`)."
);

/// Largest column index openpyxl accepts (`ZZZ`).
pub(crate) const MAX_COLUMN: u32 = 18278;
const SECS_PER_DAY: f64 = 86400.0;

// ---- columns and coordinates -------------------------------------------------------

/// 1-based column index -> letters (`1` -> `A`, `28` -> `AB`).
pub(crate) fn column_letter(col: u32) -> String {
    let mut col = col;
    let mut out = Vec::with_capacity(3);
    while col > 0 {
        let rem = (col - 1) % 26;
        out.push(b'A' + rem as u8);
        col = (col - 1) / 26;
    }
    out.reverse();
    String::from_utf8(out).expect("ascii")
}

/// openpyxl's `column_index_from_string`: one to three letters, any case.
pub(crate) fn column_index(col: &str) -> Option<u32> {
    if col.chars().count() > 3 {
        return None;
    }
    let mut idx = 0u32;
    for (power, ch) in [1u32, 26, 676].iter().zip(col.chars().rev()) {
        let up = ch.to_ascii_uppercase();
        if !up.is_ascii_uppercase() {
            return None;
        }
        idx += (up as u32 - 'A' as u32 + 1) * power;
    }
    (1..=MAX_COLUMN).contains(&idx).then_some(idx)
}

fn column_error(col: &str) -> PyErr {
    PyValueError::new_err(format!(
        "'{col}' is not a valid column name. Column names are from A to ZZZ"
    ))
}

/// `A1` coordinate of a 1-based (row, column).
pub(crate) fn coordinate(row: u32, col: u32) -> String {
    format!("{}{}", column_letter(col), row)
}

/// The parts of openpyxl's `RANGE_EXPR` regex, matched at the start of `s`.
#[derive(Default, Debug, Clone, PartialEq)]
struct RangeExpr<'a> {
    min_col: Option<&'a str>,
    min_row: Option<&'a str>,
    sep: bool,
    max_col: Option<&'a str>,
    max_row: Option<&'a str>,
    /// Number of bytes matched.
    end: usize,
}

/// `[$]?([A-Za-z]{1,3})?[$]?(\d+)?(:[$]?([A-Za-z]{1,3})?[$]?(\d+)?)?`, greedy, as a
/// prefix of `s` (all parts are optional, so it always matches, possibly empty).
fn match_range_expr(s: &str) -> RangeExpr<'_> {
    let b = s.as_bytes();
    let mut i = 0;
    let part = |i: &mut usize| -> (Option<&str>, Option<&str>) {
        if b.get(*i) == Some(&b'$') {
            *i += 1;
        }
        let start = *i;
        while *i < b.len() && *i - start < 3 && b[*i].is_ascii_alphabetic() {
            *i += 1;
        }
        let col = (*i > start).then(|| &s[start..*i]);
        if b.get(*i) == Some(&b'$') {
            *i += 1;
        }
        let start = *i;
        while *i < b.len() && b[*i].is_ascii_digit() {
            *i += 1;
        }
        let row = (*i > start).then(|| &s[start..*i]);
        (col, row)
    };
    let (min_col, min_row) = part(&mut i);
    let mut out = RangeExpr {
        min_col,
        min_row,
        end: i,
        ..Default::default()
    };
    if b.get(i) == Some(&b':') {
        i += 1;
        let (max_col, max_row) = part(&mut i);
        out.sep = true;
        out.max_col = max_col;
        out.max_row = max_row;
        out.end = i;
    }
    out
}

/// Boundaries `(min_col, min_row, max_col, max_row)`, 1-based, `None` for open ends.
pub(crate) type Boundaries = (Option<u32>, Option<u32>, Option<u32>, Option<u32>);

/// openpyxl's `range_boundaries`.
pub(crate) fn range_boundaries(range: &str) -> PyResult<Boundaries> {
    let bad = || PyValueError::new_err(format!("{range} is not a valid coordinate or range"));
    let m = match_range_expr(range);
    if m.end != range.len() {
        return Err(bad());
    }
    if m.sep {
        let cols = [m.min_col.is_some(), m.max_col.is_some()];
        let rows = [m.min_row.is_some(), m.max_row.is_some()];
        let all = |v: &[bool]| v.iter().all(|x| *x);
        let any = |v: &[bool]| v.iter().any(|x| *x);
        let both = [cols[0], cols[1], rows[0], rows[1]];
        if !(all(&both) || all(&cols) && !any(&rows) || all(&rows) && !any(&cols)) {
            return Err(bad());
        }
    }
    let col = |c: Option<&str>| -> PyResult<Option<u32>> {
        c.map(|c| column_index(c).ok_or_else(|| column_error(c)))
            .transpose()
    };
    let row = |r: Option<&str>| -> PyResult<Option<u32>> {
        r.map(|r| r.parse::<u32>().map_err(|_| bad())).transpose()
    };
    let min_col = col(m.min_col)?;
    let min_row = row(m.min_row)?;
    let max_col = col(m.max_col)?.or(min_col);
    let max_row = row(m.max_row)?.or(min_row);
    Ok((min_col, min_row, max_col, max_row))
}

/// Splits a defined-name value like `'My Sheet'!$A$1:$B$2,Other!$C$3` into
/// `(sheet, range)` parts. Commas inside quoted sheet names do not split.
fn split_sheet_ranges(value: &str) -> Vec<&str> {
    let mut parts = Vec::new();
    let (mut start, mut quoted) = (0, false);
    for (i, ch) in value.char_indices() {
        match ch {
            '\'' => quoted = !quoted,
            ',' if !quoted => {
                parts.push(&value[start..i]);
                start = i + 1;
            }
            _ => {}
        }
    }
    parts.push(&value[start..]);
    parts
}

/// `(sheet title, range text)` of each comma-separated part of a defined name
/// value; parts that cannot be split are skipped.
pub(crate) fn sheet_ranges(value: &str) -> Vec<(Option<String>, String)> {
    split_sheet_ranges(value)
        .into_iter()
        .filter_map(|p| split_sheet_ref(p).map(|(t, r)| (t, r.trim().to_string())))
        .collect()
}

/// Sheet title (unquoted, `''` unescaped) and range text of `Sheet!A1:B2`.
fn split_sheet_ref(part: &str) -> Option<(Option<String>, &str)> {
    let part = part.trim();
    if let Some(rest) = part.strip_prefix('\'') {
        let mut title = String::new();
        let mut chars = rest.char_indices().peekable();
        while let Some((i, ch)) = chars.next() {
            if ch == '\'' {
                if matches!(chars.peek(), Some((_, '\''))) {
                    title.push('\'');
                    chars.next();
                    continue;
                }
                let after = &rest[i + 1..];
                return after.strip_prefix('!').map(|r| (Some(title), r));
            }
            title.push(ch);
        }
        None
    } else {
        match part.rfind('!') {
            Some(i) => Some((Some(part[..i].to_string()), &part[i + 1..])),
            None => Some((None, part)),
        }
    }
}

// ---- number formats ----------------------------------------------------------------

/// openpyxl's `BUILTIN_FORMATS`.
pub(crate) fn builtin_format(id: u32) -> Option<&'static str> {
    Some(match id {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        5 => "\"$\"#,##0_);(\"$\"#,##0)",
        6 => "\"$\"#,##0_);[Red](\"$\"#,##0)",
        7 => "\"$\"#,##0.00_);(\"$\"#,##0.00)",
        8 => "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "mm-dd-yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        37 => "#,##0_);(#,##0)",
        38 => "#,##0_);[Red](#,##0)",
        39 => "#,##0.00_);(#,##0.00)",
        40 => "#,##0.00_);[Red](#,##0.00)",
        41 => r#"_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)"#,
        42 => r#"_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)"#,
        43 => r#"_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)"#,
        44 => r#"_("$"* #,##0.00_)_("$"* \(#,##0.00\)_("$"* "-"??_)_(@_)"#,
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => return None,
    })
}

const BUILTIN_IDS: [u32; 36] = [
    0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 37, 38, 39,
    40, 41, 42, 43, 44, 45, 46, 47, 48, 49,
];

/// Removes openpyxl's `STRIP_RE` matches: quoted literals (`".*?"`) and bracketed
/// sections other than `[h]`, `[hh]`, `[m]`, `[mm]`, `[s]`, `[ss]`.
fn strip_format(fmt: &str) -> String {
    let b = fmt.as_bytes();
    let mut out = String::with_capacity(fmt.len());
    let mut i = 0;
    while i < b.len() {
        if b[i] == b'"' {
            if let Some(close) = fmt[i + 1..].find('"') {
                i += close + 2;
                continue;
            }
        } else if b[i] == b'[' {
            let rest = &fmt[i + 1..];
            let elapsed = ["h]", "hh]", "m]", "mm]", "s]", "ss]"]
                .iter()
                .any(|p| rest.starts_with(p));
            if !elapsed {
                if let Some(close) = rest.find(']') {
                    i += close + 2;
                    continue;
                }
            }
        }
        let ch = fmt[i..].chars().next().expect("in bounds");
        out.push(ch);
        i += ch.len_utf8();
    }
    out
}

/// openpyxl's `is_date_format`.
pub(crate) fn is_date_format(fmt: &str) -> bool {
    let first = fmt.split(';').next().unwrap_or("");
    let stripped = strip_format(first);
    let mut prev: Option<char> = None;
    for ch in stripped.chars() {
        if "dmhysDMHYS".contains(ch) && !matches!(prev, Some('_') | Some('\\')) {
            return true;
        }
        prev = Some(ch);
    }
    false
}

/// openpyxl's `is_timedelta_format`: an elapsed-time section (`[h]`, `[mm]`, ...).
pub(crate) fn is_timedelta_format(fmt: &str) -> bool {
    let first = fmt.split(';').next().unwrap_or("").to_ascii_lowercase();
    ["[h]", "[hh]", "[m]", "[mm]", "[s]", "[ss]"]
        .iter()
        .any(|p| first.contains(p))
}

// ---- dates -------------------------------------------------------------------------

pub(crate) fn windows_epoch() -> NaiveDateTime {
    NaiveDate::from_ymd_opt(1899, 12, 30)
        .expect("valid")
        .and_hms_opt(0, 0, 0)
        .expect("valid")
}

pub(crate) fn mac_epoch() -> NaiveDateTime {
    NaiveDate::from_ymd_opt(1904, 1, 1)
        .expect("valid")
        .and_hms_opt(0, 0, 0)
        .expect("valid")
}

/// Result of openpyxl's `from_excel`.
#[derive(Debug, Clone, PartialEq)]
pub(crate) enum ExcelDate {
    DateTime(NaiveDateTime),
    Time(NaiveTime),
    Delta(Duration),
}

/// Round half to even, like Python's `round()`.
fn round_half_even(x: f64) -> f64 {
    let r = x.round();
    if (x - x.trunc()).abs() == 0.5 {
        2.0 * (x / 2.0).round()
    } else {
        r
    }
}

/// CPython's `timedelta(days=value)` in microseconds.
fn days_to_micros(value: f64) -> Option<i64> {
    // Same steps as CPython's `accum()` for a float argument.
    let int_days = value.trunc();
    let frac = value - int_days;
    let whole = (int_days as i64).checked_mul(86_400_000_000)?;
    let us = frac * 86_400_000_000.0;
    let us_int = us.trunc();
    let leftover = us - us_int;
    whole
        .checked_add(us_int as i64)?
        .checked_add(round_half_even(leftover) as i64)
}

/// openpyxl's `from_excel`. `None` when the result is outside Python's date range.
pub(crate) fn from_excel(value: f64, epoch: NaiveDateTime, timedelta: bool) -> Option<ExcelDate> {
    if !value.is_finite() {
        return None;
    }
    if timedelta {
        let us = days_to_micros(value)?;
        let micros = us.rem_euclid(1_000_000);
        if micros == 0 {
            return Some(ExcelDate::Delta(Duration::microseconds(us)));
        }
        // timedelta(seconds=td.total_seconds() // 1, microseconds=round(td.microseconds, -3))
        let seconds = us.div_euclid(1_000_000);
        let ms = round_half_even(micros as f64 / 1000.0) as i64 * 1000;
        return Some(ExcelDate::Delta(Duration::microseconds(
            seconds.checked_mul(1_000_000)?.checked_add(ms)?,
        )));
    }
    let mut day = value.floor();
    let fraction = value - day;
    let ms = round_half_even(fraction * SECS_PER_DAY * 1000.0) as i64;
    if (0.0..1.0).contains(&value) && ms < 86_400_000 {
        let secs = (ms / 1000) as u32;
        return NaiveTime::from_hms_micro_opt(
            secs / 3600,
            secs / 60 % 60,
            secs % 60,
            (ms % 1000) as u32 * 1000,
        )
        .map(ExcelDate::Time);
    }
    if value > 0.0 && value < 60.0 && epoch == windows_epoch() {
        day += 1.0;
    }
    if day.abs() > 4_000_000.0 {
        return None;
    }
    let dt = epoch
        .checked_add_signed(Duration::days(day as i64))?
        .checked_add_signed(Duration::milliseconds(ms))?;
    (1..=9999)
        .contains(&chrono::Datelike::year(&dt))
        .then_some(ExcelDate::DateTime(dt))
}

/// Value of openpyxl's `from_ISO8601` (`t="d"` cells).
#[derive(Debug, Clone, PartialEq)]
pub(crate) enum IsoValue {
    Date(NaiveDate),
    Time(NaiveTime),
    DateTime(NaiveDateTime),
    Delta(Duration),
}

fn digits(b: &[u8], at: usize, n: usize) -> Option<u32> {
    let s = b.get(at..at + n)?;
    s.iter()
        .all(u8::is_ascii_digit)
        .then(|| std::str::from_utf8(s).ok()?.parse().ok())
        .flatten()
}

/// openpyxl's `from_ISO8601`. `Ok(None)` for an empty string, `Err` when invalid.
pub(crate) fn from_iso8601(s: &str) -> Result<Option<IsoValue>, String> {
    if s.is_empty() {
        return Ok(None);
    }
    let invalid = || format!("Invalid datetime value {s}");
    let b = s.as_bytes();
    let mut i = 0;
    let date = match (
        digits(b, 0, 4),
        b.get(4),
        digits(b, 5, 2),
        b.get(7),
        digits(b, 8, 2),
    ) {
        (Some(y), Some(b'-'), Some(m), Some(b'-'), Some(d)) => {
            i = 10;
            Some((y, m, d))
        }
        _ => None,
    };
    if b.get(i) == Some(&b'T') {
        i += 1;
    }
    let mut time = None;
    if let (Some(h), Some(b':'), Some(mi)) = (digits(b, i, 2), b.get(i + 2), digits(b, i + 3, 2)) {
        i += 5;
        let (mut sec, mut micro) = (0, 0);
        if let (Some(b':'), Some(sv)) = (b.get(i), digits(b, i + 1, 2)) {
            sec = sv;
            i += 3;
            if b.get(i) == Some(&b'.') {
                let n = b[i + 1..]
                    .iter()
                    .take(3)
                    .take_while(|c| c.is_ascii_digit())
                    .count();
                if n > 0 {
                    let frac: f64 = s[i..i + 1 + n].parse().map_err(|_| invalid())?;
                    micro = (frac * 1_000_000.0) as u32;
                    i += 1 + n;
                }
            }
        }
        time = Some((h, mi, sec, micro));
    }
    if date.is_some() || time.is_some() {
        let t = time.map(|(h, m, s, us)| NaiveTime::from_hms_micro_opt(h, m, s, us));
        return match (date, t) {
            (Some((y, m, d)), None) => NaiveDate::from_ymd_opt(y as i32, m, d)
                .map(IsoValue::Date)
                .map(Some)
                .ok_or_else(invalid),
            (None, Some(t)) => t.map(IsoValue::Time).map(Some).ok_or_else(invalid),
            (Some((y, m, d)), Some(t)) => NaiveDate::from_ymd_opt(y as i32, m, d)
                .zip(t)
                .map(|(d, t)| Some(IsoValue::DateTime(d.and_time(t))))
                .ok_or_else(invalid),
            (None, None) => unreachable!(),
        };
    }
    // PT((\d+)H)?((\d+)M)?((\d+(\.\d{1,3})?)S)?
    let _ = i;
    if let Some(rest) = s.strip_prefix("PT") {
        let rb = rest.as_bytes();
        let mut j = 0;
        let mut total = 0f64;
        let mut any = false;
        let number = |j: &mut usize, unit: u8, frac: bool| -> Option<f64> {
            let start = *j;
            let mut k = *j;
            while k < rb.len() && rb[k].is_ascii_digit() {
                k += 1;
            }
            if k == start {
                return None;
            }
            if frac && rb.get(k) == Some(&b'.') {
                let n = rb[k + 1..]
                    .iter()
                    .take(3)
                    .take_while(|c| c.is_ascii_digit())
                    .count();
                if n > 0 {
                    k += 1 + n;
                }
            }
            if rb.get(k) != Some(&unit) {
                return None;
            }
            let v = rest[start..k].parse().ok()?;
            *j = k + 1;
            Some(v)
        };
        if let Some(h) = number(&mut j, b'H', false) {
            total += h * 3600.0;
            any = true;
        }
        if let Some(m) = number(&mut j, b'M', false) {
            total += m * 60.0;
            any = true;
        }
        if let Some(sec) = number(&mut j, b'S', true) {
            total += sec;
            any = true;
        }
        if any {
            return Ok(Some(IsoValue::Delta(Duration::microseconds(
                round_half_even(total * 1_000_000.0) as i64,
            ))));
        }
    }
    Err(invalid())
}

// ---- Python wrappers ---------------------------------------------------------------

fn epoch_arg(epoch: Option<&Bound<'_, PyAny>>) -> PyResult<NaiveDateTime> {
    match epoch {
        None => Ok(windows_epoch()),
        Some(e) => e.extract::<NaiveDateTime>(),
    }
}

pub(crate) fn excel_date_to_py<'py>(py: Python<'py>, v: &ExcelDate) -> PyResult<Bound<'py, PyAny>> {
    Ok(match v {
        ExcelDate::DateTime(dt) => dt.into_pyobject(py)?.into_any(),
        ExcelDate::Time(t) => t.into_pyobject(py)?.into_any(),
        ExcelDate::Delta(d) => d.into_pyobject(py)?.into_any(),
    })
}

pub(crate) fn iso_to_py<'py>(py: Python<'py>, v: &IsoValue) -> PyResult<Bound<'py, PyAny>> {
    Ok(match v {
        IsoValue::Date(d) => d.into_pyobject(py)?.into_any(),
        IsoValue::Time(t) => t.into_pyobject(py)?.into_any(),
        IsoValue::DateTime(dt) => dt.into_pyobject(py)?.into_any(),
        IsoValue::Delta(d) => d.into_pyobject(py)?.into_any(),
    })
}

/// Excel column letters of a 1-based column index (`28` -> `"AB"`).
#[pyfunction]
pub(crate) fn get_column_letter(col_idx: i64) -> PyResult<String> {
    if !(1..=MAX_COLUMN as i64).contains(&col_idx) {
        return Err(PyValueError::new_err(format!(
            "Invalid column index {col_idx}"
        )));
    }
    Ok(column_letter(col_idx as u32))
}

/// 1-based column index of one to three column letters (`"AB"` -> `28`).
#[pyfunction]
pub(crate) fn column_index_from_string(col: &str) -> PyResult<u32> {
    column_index(col).ok_or_else(|| column_error(col))
}

/// `"B12"` -> `("B", 12)`.
#[pyfunction]
pub(crate) fn coordinate_from_string(coord_string: &str) -> PyResult<(String, u32)> {
    let m = match_range_expr(coord_string);
    let invalid =
        || CellCoordinatesException::new_err(format!("Invalid cell coordinates ({coord_string})"));
    match (m.min_col, m.min_row, m.sep, m.end == coord_string.len()) {
        (Some(col), Some(row), false, true) => {
            let row: u32 = row.parse().map_err(|_| invalid())?;
            if row == 0 {
                return Err(CellCoordinatesException::new_err(format!(
                    "There is no row 0 ({coord_string})"
                )));
            }
            Ok((col.to_string(), row))
        }
        _ => Err(invalid()),
    }
}

/// `"C7"` -> `(7, 3)` (1-based row, column).
#[pyfunction]
pub(crate) fn coordinate_to_tuple(py: Python<'_>, coordinate: &str) -> PyResult<(Py<PyAny>, u32)> {
    if coordinate.is_empty() {
        return Err(PyValueError::new_err("empty coordinate"));
    }
    let chars: Vec<(usize, char)> = coordinate.char_indices().collect();
    let idx = chars
        .iter()
        .find(|(_, c)| c.is_ascii_digit())
        .or(chars.last())
        .map(|(i, _)| *i)
        .expect("not empty");
    let (col, row) = coordinate.split_at(idx);
    // `int(row)` exactly like openpyxl (accepts surrounding whitespace, ...).
    let row = py.get_type::<PyInt>().call1((row,))?.unbind();
    Ok((row, column_index_from_string(col)?))
}

/// `A1` coordinate of a 1-based row and column (`(7, 3)` -> `"C7"`).
#[pyfunction]
pub(crate) fn coordinate_from_row_col(row: i64, column: i64) -> PyResult<String> {
    if !(1..=1_048_576).contains(&row) {
        return Err(PyValueError::new_err(format!(
            "Row numbers must be between 1 and 1048576. Row number supplied was {row}"
        )));
    }
    Ok(format!("{}{row}", get_column_letter(column)?))
}

/// Inverse of `coordinate_to_tuple`: `(7, 3)` -> `"C7"`.
#[pyfunction]
pub(crate) fn coordinate_from_tuple(row_column: (i64, i64)) -> PyResult<String> {
    coordinate_from_row_col(row_column.0, row_column.1)
}

/// `"B12"` -> `"$B$12"`, `"A1:B2"` -> `"$A$1:$B$2"`.
#[pyfunction]
pub(crate) fn absolute_coordinate(coord_string: &str) -> PyResult<String> {
    let m = match_range_expr(coord_string);
    if m.end != coord_string.len() {
        return Err(PyValueError::new_err(format!(
            "{coord_string} is not a valid coordinate range"
        )));
    }
    let abs = |p: Option<&str>| p.map(|v| format!("${v}")).unwrap_or_default();
    let mut out = format!("{}{}", abs(m.min_col), abs(m.min_row));
    if m.max_col.is_some() || m.max_row.is_some() {
        out.push_str(&format!(":{}{}", abs(m.max_col), abs(m.max_row)));
    }
    Ok(out)
}

/// `"A1:C3"` -> `(min_col, min_row, max_col, max_row)`, 1-based; `None` for open ends
/// (`"A:C"`, `"1:3"`).
#[pyfunction(name = "range_boundaries")]
pub(crate) fn py_range_boundaries(range_string: &str) -> PyResult<Boundaries> {
    range_boundaries(range_string)
}

/// `"'Sheet 1'!A1:B2"` -> `("Sheet 1", (1, 1, 2, 2))` (quotes inside the title stay doubled,
/// as in openpyxl).
#[pyfunction]
pub(crate) fn range_to_tuple(range_string: &str) -> PyResult<(String, Boundaries)> {
    let err = || PyValueError::new_err("Value must be of the form sheetname!A1:E4");
    let (title, rest) = if let Some(rest) = range_string.strip_prefix('\'') {
        let b = rest.as_bytes();
        let mut i = 0;
        loop {
            match b.get(i) {
                None => return Err(err()),
                Some(b'\'') if b.get(i + 1) == Some(&b'\'') => i += 2,
                Some(b'\'') => break,
                _ => i += 1,
            }
        }
        (&rest[..i], &rest[i + 1..])
    } else {
        let end = range_string
            .find(['\'', '^', ' ', '!'])
            .unwrap_or(range_string.len());
        (&range_string[..end], &range_string[end..])
    };
    let cells = rest.strip_prefix('!').ok_or_else(err)?;
    let m = match_range_expr(cells);
    Ok((title.to_string(), range_boundaries(&cells[..m.end])?))
}

/// Wraps a sheet title in quotes (`"Sheet 1"` -> `"'Sheet 1'"`), doubling inner quotes.
#[pyfunction]
pub(crate) fn quote_sheetname(sheetname: &str) -> String {
    format!("'{}'", sheetname.replace('\'', "''"))
}

fn full_boundaries(range: &str) -> PyResult<(u32, u32, u32, u32)> {
    match range_boundaries(range)? {
        (Some(a), Some(b), Some(c), Some(d)) => Ok((a, b, c, d)),
        _ => Err(PyValueError::new_err(format!(
            "{range} is not a valid coordinate or range"
        ))),
    }
}

/// Cell coordinates of a range, one tuple per row.
#[pyfunction]
pub(crate) fn rows_from_range<'py>(
    py: Python<'py>,
    range_string: &str,
) -> PyResult<Bound<'py, PyIterator>> {
    let (c0, r0, c1, r1) = full_boundaries(range_string)?;
    let rows = PyList::empty(py);
    for r in r0..=r1 {
        rows.append(PyTuple::new(
            py,
            (c0..=c1).map(|c| coordinate(r, c)).collect::<Vec<_>>(),
        )?)?;
    }
    rows.try_iter()
}

/// Cell coordinates of a range, one tuple per column.
#[pyfunction]
pub(crate) fn cols_from_range<'py>(
    py: Python<'py>,
    range_string: &str,
) -> PyResult<Bound<'py, PyIterator>> {
    let (c0, r0, c1, r1) = full_boundaries(range_string)?;
    let cols = PyList::empty(py);
    for c in c0..=c1 {
        cols.append(PyTuple::new(
            py,
            (r0..=r1).map(|r| coordinate(r, c)).collect::<Vec<_>>(),
        )?)?;
    }
    cols.try_iter()
}

fn column_arg(v: &Bound<'_, PyAny>) -> PyResult<u32> {
    if let Ok(s) = v.extract::<String>() {
        return column_index_from_string(&s);
    }
    let i: i64 = v.extract()?;
    u32::try_from(i).map_err(|_| PyValueError::new_err(format!("Invalid column index {i}")))
}

/// Column letters from `start` to `end` (letters or 1-based indexes), inclusive.
#[pyfunction]
pub(crate) fn get_column_interval(
    start: &Bound<'_, PyAny>,
    end: &Bound<'_, PyAny>,
) -> PyResult<Vec<String>> {
    let (start, end) = (column_arg(start)?, column_arg(end)?);
    (start..=end).map(|c| get_column_letter(c as i64)).collect()
}

/// Static A1 ranges of a print area defined name, e.g.
/// `"'Data'!$A$1:$C$10,'Data'!$E$1:$F$2"` -> `["A1:C10", "E1:F2"]`. Rows-only
/// (`"$1:$3"`) and columns-only (`"$A:$B"`) ranges are kept as such. Raises
/// `ValueError` for anything that is not a list of static ranges (formulas,
/// other defined names, ...).
#[pyfunction]
pub(crate) fn split_print_areas(value: &str) -> PyResult<Vec<String>> {
    let bad = || PyValueError::new_err(format!("{value:?} is not a static list of ranges"));
    let mut out = Vec::new();
    for part in split_sheet_ranges(value) {
        let (_, range) = split_sheet_ref(part).ok_or_else(bad)?;
        let range = range.trim();
        if range.is_empty() || range.contains('(') {
            return Err(bad());
        }
        let b = range_boundaries(range).map_err(|_| bad())?;
        let text = match b {
            (Some(c0), Some(r0), Some(c1), Some(r1))
                if (c0, r0) == (c1, r1) && !range.contains(':') =>
            {
                coordinate(r0, c0)
            }
            (Some(c0), Some(r0), Some(c1), Some(r1)) => {
                format!("{}:{}", coordinate(r0, c0), coordinate(r1, c1))
            }
            (None, Some(r0), None, Some(r1)) => format!("{r0}:{r1}"),
            (Some(c0), None, Some(c1), None) => {
                format!("{}:{}", column_letter(c0), column_letter(c1))
            }
            _ => return Err(bad()),
        };
        out.push(text);
    }
    Ok(out)
}

/// openpyxl's `is_date_format`.
#[pyfunction(name = "is_date_format")]
pub(crate) fn py_is_date_format(fmt: Option<&str>) -> bool {
    fmt.is_some_and(is_date_format)
}

/// openpyxl's `is_timedelta_format`.
#[pyfunction(name = "is_timedelta_format")]
pub(crate) fn py_is_timedelta_format(fmt: Option<&str>) -> bool {
    fmt.is_some_and(is_timedelta_format)
}

/// openpyxl's `BUILTIN_FORMATS` as a new dict.
#[pyfunction]
pub(crate) fn builtin_formats(py: Python<'_>) -> PyResult<Bound<'_, PyDict>> {
    let d = PyDict::new(py);
    for id in &BUILTIN_IDS {
        d.set_item(id, builtin_format(*id))?;
    }
    Ok(d)
}

/// Excel serial -> `datetime` (or `time` for `0 <= value < 1`, `timedelta` with
/// `timedelta=True`), like openpyxl's `from_excel`.
#[pyfunction(name = "from_excel", signature = (value, epoch=None, timedelta=false))]
pub(crate) fn py_from_excel<'py>(
    py: Python<'py>,
    value: Option<&Bound<'py, PyAny>>,
    epoch: Option<&Bound<'py, PyAny>>,
    timedelta: bool,
) -> PyResult<Option<Bound<'py, PyAny>>> {
    let Some(value) = value.filter(|v| !v.is_none()) else {
        return Ok(None);
    };
    let v: f64 = value.extract()?;
    let epoch = epoch_arg(epoch)?;
    match from_excel(v, epoch, timedelta) {
        Some(d) => excel_date_to_py(py, &d).map(Some),
        None => Err(pyo3::exceptions::PyOverflowError::new_err(format!(
            "date value out of range: {v}"
        ))),
    }
}

/// Excel serial -> `datetime`, always a `datetime` (fractions below 1 are times on
/// the epoch day). Uses the 1900 date system unless `epoch_1904=True`.
#[pyfunction(signature = (serial, epoch_1904=false))]
pub(crate) fn serial_to_datetime(serial: f64, epoch_1904: bool) -> PyResult<NaiveDateTime> {
    let epoch = if epoch_1904 {
        mac_epoch()
    } else {
        windows_epoch()
    };
    let out = if (0.0..1.0).contains(&serial) {
        let ms = round_half_even(serial * SECS_PER_DAY * 1000.0) as i64;
        epoch.checked_add_signed(Duration::milliseconds(ms))
    } else {
        match from_excel(serial, epoch, false) {
            Some(ExcelDate::DateTime(dt)) => Some(dt),
            Some(ExcelDate::Time(t)) => Some(epoch.date().and_time(t)),
            _ => None,
        }
    };
    out.ok_or_else(|| PyValueError::new_err(format!("date serial out of range: {serial}")))
}

fn time_to_days(t: &NaiveTime) -> f64 {
    (t.hour() as f64 * 3600.0
        + t.minute() as f64 * 60.0
        + t.second() as f64
        + (t.nanosecond() / 1000) as f64 / 1e6)
        / SECS_PER_DAY
}

/// Python `date` / `datetime` / `time` / `timedelta` -> Excel serial, like openpyxl's `to_excel`.
#[pyfunction(signature = (dt, epoch=None))]
pub(crate) fn to_excel(dt: &Bound<'_, PyAny>, epoch: Option<&Bound<'_, PyAny>>) -> PyResult<f64> {
    let epoch = epoch_arg(epoch)?;
    if dt.is_instance_of::<PyTime>() {
        return Ok(time_to_days(&dt.extract::<NaiveTime>()?));
    }
    if dt.is_instance_of::<PyDelta>() {
        let total: f64 = dt.call_method0("total_seconds")?.extract()?;
        return Ok(total / SECS_PER_DAY);
    }
    let dt: NaiveDateTime = if dt.is_instance_of::<PyDateTime>() {
        if !dt.getattr("tzinfo")?.is_none() {
            return Err(PyTypeError::new_err(
                "can't subtract offset-naive and offset-aware datetimes",
            ));
        }
        dt.extract()?
    } else if dt.is_instance_of::<PyDate>() {
        dt.extract::<NaiveDate>()?
            .and_hms_opt(0, 0, 0)
            .expect("midnight")
    } else {
        return Err(PyTypeError::new_err(
            "expected a date, datetime, time or timedelta",
        ));
    };
    let micros = (dt - epoch)
        .num_microseconds()
        .ok_or_else(|| PyValueError::new_err("date out of range"))?;
    let mut days = micros.div_euclid(86_400_000_000);
    if days > 0 && days <= 60 && epoch == windows_epoch() {
        days -= 1;
    }
    Ok(days as f64 + time_to_days(&dt.time()))
}

/// ISO 8601 text -> `date` / `time` / `datetime` / `timedelta`, like openpyxl's `from_ISO8601`.
#[pyfunction(name = "from_ISO8601")]
pub(crate) fn py_from_iso8601<'py>(
    py: Python<'py>,
    formatted_string: &str,
) -> PyResult<Option<Bound<'py, PyAny>>> {
    match from_iso8601(formatted_string) {
        Ok(Some(v)) => iso_to_py(py, &v).map(Some),
        Ok(None) => Ok(None),
        Err(msg) => Err(PyValueError::new_err(msg)),
    }
}

pub(crate) fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    let py = m.py();
    for f in [
        wrap_pyfunction!(get_column_letter, m)?,
        wrap_pyfunction!(column_index_from_string, m)?,
        wrap_pyfunction!(coordinate_from_string, m)?,
        wrap_pyfunction!(coordinate_to_tuple, m)?,
        wrap_pyfunction!(coordinate_from_row_col, m)?,
        wrap_pyfunction!(coordinate_from_tuple, m)?,
        wrap_pyfunction!(absolute_coordinate, m)?,
        wrap_pyfunction!(py_range_boundaries, m)?,
        wrap_pyfunction!(range_to_tuple, m)?,
        wrap_pyfunction!(quote_sheetname, m)?,
        wrap_pyfunction!(rows_from_range, m)?,
        wrap_pyfunction!(cols_from_range, m)?,
        wrap_pyfunction!(get_column_interval, m)?,
        wrap_pyfunction!(split_print_areas, m)?,
        wrap_pyfunction!(py_is_date_format, m)?,
        wrap_pyfunction!(py_is_timedelta_format, m)?,
        wrap_pyfunction!(builtin_formats, m)?,
        wrap_pyfunction!(serial_to_datetime, m)?,
        wrap_pyfunction!(to_excel, m)?,
        wrap_pyfunction!(py_from_iso8601, m)?,
    ] {
        m.add_function(f)?;
    }
    m.add_function(wrap_pyfunction!(py_from_excel, m)?)?;
    m.add("WINDOWS_EPOCH", windows_epoch().into_pyobject(py)?)?;
    m.add("MAC_EPOCH", mac_epoch().into_pyobject(py)?)?;
    m.add(
        "CellCoordinatesException",
        py.get_type::<CellCoordinatesException>(),
    )?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn letters() {
        for (i, s) in [
            (1, "A"),
            (26, "Z"),
            (27, "AA"),
            (52, "AZ"),
            (702, "ZZ"),
            (703, "AAA"),
            (18278, "ZZZ"),
        ] {
            assert_eq!(column_letter(i), s);
            assert_eq!(column_index(s), Some(i));
        }
        assert_eq!(column_index("zz"), Some(702));
        assert_eq!(column_index(""), None);
        assert_eq!(column_index("AAAA"), None);
    }

    #[test]
    fn date_formats() {
        assert!(is_date_format("yyyy-mm-dd"));
        assert!(is_date_format("[h]:mm:ss"));
        assert!(!is_date_format("0.00"));
        assert!(!is_date_format("\"days\" 0"));
        assert!(!is_date_format("[Red]0.00"));
        assert!(!is_date_format("_d0"));
        assert!(is_timedelta_format("[hh]:mm"));
        assert!(!is_timedelta_format("hh:mm"));
    }
}
