//! Conversion of the fork's rich workbook / sheet data to Python objects.

use calamine::{
    AutoFilter, CellStyle, ColumnInfo, ConditionalFormat, CustomProperty, DataValidation,
    DifferentialStyle, HeaderFooter, Hyperlink, NamedStyle, RowAttributes, Scenarios, SheetView,
    SortState, StyleAlignment, StyleBorder, StyleBorderSide, StyleColor, StyleFill, StyleFont,
    TableInfo, TextRun, WorksheetInfo,
};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyDict, PyList};

/// Attributes stored as `"1"`/`"0"` that are booleans (everything else that parses
/// as a number stays a number).
const BOOL_ATTRS: &[&str] = &[
    // sheetProtection / workbookProtection
    "sheet",
    "objects",
    "scenarios",
    "formatCells",
    "formatColumns",
    "formatRows",
    "insertColumns",
    "insertRows",
    "insertHyperlinks",
    "deleteColumns",
    "deleteRows",
    "selectLockedCells",
    "selectUnlockedCells",
    "sort",
    "autoFilter",
    "pivotTables",
    "lockStructure",
    "lockWindows",
    "lockRevision",
    // sheetView / customSheetView
    "showGridLines",
    "showRowColHeaders",
    "showZeros",
    "showFormulas",
    "rightToLeft",
    "tabSelected",
    "showRuler",
    "showOutlineSymbols",
    "defaultGridColor",
    "showWhiteSpace",
    "windowProtection",
    "showPageBreaks",
    "outlineSymbols",
    "zeroValues",
    "hiddenRows",
    "hiddenColumns",
    "filter",
    "showAutoFilter",
    "printArea",
    "filterUnique",
    "fitToPage",
    // printOptions / pageSetup / pageSetUpPr
    "horizontalCentered",
    "verticalCentered",
    "headings",
    "gridLines",
    "gridLinesSet",
    "usePrinterDefaults",
    "blackAndWhite",
    "draft",
    "useFirstPageNumber",
    "autoPageBreaks",
    // sheetPr / outlinePr / sheetFormatPr
    "filterMode",
    "published",
    "syncHorizontal",
    "syncVertical",
    "transitionEvaluation",
    "transitionEntry",
    "enableFormatConditionsCalculation",
    "summaryBelow",
    "summaryRight",
    "applyStyles",
    "customHeight",
    "zeroHeight",
    "thickTop",
    "thickBottom",
    // headerFooter
    "differentOddEven",
    "differentFirst",
    "scaleWithDoc",
    "alignWithMargins",
    // page breaks
    "man",
    "pt",
    // workbookPr / workbookView / calcPr
    "date1904",
    "dateCompatibility",
    "showBorderUnselectedTables",
    "promptedSolutions",
    "showInkAnnotation",
    "backupFile",
    "saveExternalLinkValues",
    "hidePivotFieldList",
    "showPivotChartFilter",
    "allowRefreshQuery",
    "publishItems",
    "checkCompatibility",
    "autoCompressPictures",
    "refreshAllConnections",
    "filterPrivacy",
    "minimized",
    "showHorizontalScroll",
    "showVerticalScroll",
    "showSheetTabs",
    "autoFilterDateGrouping",
    "fullCalcOnLoad",
    "calcCompleted",
    "calcOnSave",
    "concurrentCalc",
    "forceFullCalc",
    "fullPrecision",
    "iterate",
    // tables / auto filters / sorting
    "insertRow",
    "insertRowShift",
    "showFirstColumn",
    "showLastColumn",
    "showRowStripes",
    "showColumnStripes",
    "totalsRowShown",
    "hiddenButton",
    "showButton",
    "and",
    "top",
    "blank",
    "columnSort",
    "caseSensitive",
    "descending",
    // scenarios
    "locked",
    "hidden",
    "deleted",
    "undone",
    // conditional formatting extras
    "percent",
    "bottom",
    "aboveAverage",
    "equalAverage",
    "showValue",
    "reverse",
    "gradient",
];

/// `camelCase` -> `snake_case`.
pub(crate) fn snake(key: &str) -> String {
    let mut out = String::with_capacity(key.len() + 4);
    for (i, ch) in key.chars().enumerate() {
        if ch.is_ascii_uppercase() {
            if i > 0 {
                out.push('_');
            }
            out.push(ch.to_ascii_lowercase());
        } else {
            out.push(ch);
        }
    }
    out
}

/// Attribute value typed as bool / int / float / str.
fn typed<'py>(py: Python<'py>, key: &str, v: &str) -> PyResult<Bound<'py, PyAny>> {
    if v == "true" || v == "false" {
        return Ok(PyBool::new(py, v == "true").to_owned().into_any());
    }
    if BOOL_ATTRS.contains(&key) && (v == "1" || v == "0") {
        return Ok(PyBool::new(py, v == "1").to_owned().into_any());
    }
    // Keep leading zeros / signs that are not plain numbers (e.g. GUIDs, "007") as text.
    if let Ok(i) = v.parse::<i64>() {
        if i.to_string() == v {
            return Ok(i.into_pyobject(py)?.into_any());
        }
    }
    if let Ok(f) = v.parse::<f64>() {
        if f.is_finite() && !v.starts_with('0') || v.starts_with("0.") {
            return Ok(f.into_pyobject(py)?.into_any());
        }
    }
    Ok(v.into_pyobject(py)?.into_any())
}

/// Raw attributes as a dict with snake_case keys and typed values.
pub(crate) fn attrs_to_py<'py>(
    py: Python<'py>,
    attrs: &[(String, String)],
) -> PyResult<Bound<'py, PyDict>> {
    let d = PyDict::new(py);
    for (k, v) in attrs {
        d.set_item(snake(k), typed(py, k, v)?)?;
    }
    Ok(d)
}

fn opt_attrs_to_py<'py>(
    py: Python<'py>,
    attrs: &Option<Vec<(String, String)>>,
) -> PyResult<Option<Bound<'py, PyDict>>> {
    attrs.as_ref().map(|a| attrs_to_py(py, a)).transpose()
}

fn attrs_list_to_py<'py>(
    py: Python<'py>,
    list: &[Vec<(String, String)>],
) -> PyResult<Bound<'py, PyList>> {
    let out = PyList::empty(py);
    for a in list {
        out.append(attrs_to_py(py, a)?)?;
    }
    Ok(out)
}

pub(crate) fn color_to_py<'py>(
    py: Python<'py>,
    c: &Option<StyleColor>,
) -> PyResult<Option<Bound<'py, PyDict>>> {
    let Some(c) = c else { return Ok(None) };
    let d = PyDict::new(py);
    if let Some(v) = &c.rgb {
        d.set_item("rgb", v)?;
    }
    if let Some(v) = c.theme {
        d.set_item("theme", v)?;
    }
    if let Some(v) = c.tint {
        d.set_item("tint", v)?;
    }
    if let Some(v) = c.indexed {
        d.set_item("indexed", v)?;
    }
    if c.auto {
        d.set_item("auto", true)?;
    }
    Ok(Some(d))
}

pub(crate) fn font_to_py<'py>(py: Python<'py>, f: &StyleFont) -> PyResult<Bound<'py, PyDict>> {
    let d = PyDict::new(py);
    d.set_item("name", &f.name)?;
    d.set_item("size", f.size)?;
    d.set_item("bold", f.bold)?;
    d.set_item("italic", f.italic)?;
    d.set_item("underline", &f.underline)?;
    d.set_item("strike", f.strike)?;
    d.set_item("vert_align", &f.vert_align)?;
    d.set_item("color", color_to_py(py, &f.color)?)?;
    d.set_item("family", f.family)?;
    d.set_item("charset", f.charset)?;
    d.set_item("scheme", &f.scheme)?;
    d.set_item("outline", f.outline)?;
    d.set_item("shadow", f.shadow)?;
    d.set_item("condense", f.condense)?;
    d.set_item("extend", f.extend)?;
    Ok(d)
}

fn fill_to_py<'py>(py: Python<'py>, f: &StyleFill) -> PyResult<Bound<'py, PyDict>> {
    let d = PyDict::new(py);
    d.set_item("pattern", &f.pattern)?;
    d.set_item("fg_color", color_to_py(py, &f.fg_color)?)?;
    d.set_item("bg_color", color_to_py(py, &f.bg_color)?)?;
    let gradient = match &f.gradient {
        Some(g) => {
            let gd = PyDict::new(py);
            gd.set_item("type", &g.typ)?;
            gd.set_item("degree", g.degree)?;
            gd.set_item("left", g.left)?;
            gd.set_item("right", g.right)?;
            gd.set_item("top", g.top)?;
            gd.set_item("bottom", g.bottom)?;
            let stops = PyList::empty(py);
            for (pos, c) in &g.stops {
                let s = PyDict::new(py);
                s.set_item("position", pos)?;
                s.set_item("color", color_to_py(py, &Some(c.clone()))?)?;
                stops.append(s)?;
            }
            gd.set_item("stops", stops)?;
            Some(gd)
        }
        None => None,
    };
    d.set_item("gradient", gradient)?;
    Ok(d)
}

fn side_to_py<'py>(py: Python<'py>, side: &StyleBorderSide) -> PyResult<Bound<'py, PyDict>> {
    let s = PyDict::new(py);
    s.set_item("style", &side.style)?;
    s.set_item("color", color_to_py(py, &side.color)?)?;
    Ok(s)
}

fn border_to_py<'py>(py: Python<'py>, b: &StyleBorder) -> PyResult<Bound<'py, PyDict>> {
    let d = PyDict::new(py);
    for (name, side) in [
        ("left", &b.left),
        ("right", &b.right),
        ("top", &b.top),
        ("bottom", &b.bottom),
        ("diagonal", &b.diagonal),
        ("vertical", &b.vertical),
        ("horizontal", &b.horizontal),
    ] {
        d.set_item(name, side_to_py(py, side)?)?;
    }
    d.set_item("diagonal_up", b.diagonal_up)?;
    d.set_item("diagonal_down", b.diagonal_down)?;
    d.set_item("outline", b.outline)?;
    Ok(d)
}

fn alignment_to_py<'py>(py: Python<'py>, a: &StyleAlignment) -> PyResult<Bound<'py, PyDict>> {
    let d = PyDict::new(py);
    d.set_item("horizontal", &a.horizontal)?;
    d.set_item("vertical", &a.vertical)?;
    d.set_item("wrap_text", a.wrap_text)?;
    d.set_item("shrink_to_fit", a.shrink_to_fit)?;
    d.set_item("indent", a.indent)?;
    d.set_item("text_rotation", a.text_rotation)?;
    d.set_item("justify_last_line", a.justify_last_line)?;
    d.set_item("reading_order", a.reading_order)?;
    d.set_item("relative_indent", a.relative_indent)?;
    Ok(d)
}

pub(crate) fn style_to_py<'py>(py: Python<'py>, s: &CellStyle) -> PyResult<Bound<'py, PyDict>> {
    let protection = PyDict::new(py);
    protection.set_item("locked", s.locked)?;
    protection.set_item("hidden", s.hidden)?;

    let d = PyDict::new(py);
    d.set_item("number_format", &s.number_format)?;
    d.set_item("number_format_id", s.number_format_id)?;
    d.set_item("font", font_to_py(py, &s.font)?)?;
    d.set_item("fill", fill_to_py(py, &s.fill)?)?;
    d.set_item("border", border_to_py(py, &s.border)?)?;
    d.set_item("alignment", alignment_to_py(py, &s.alignment)?)?;
    d.set_item("protection", protection)?;
    d.set_item("named_style", &s.named_style)?;
    d.set_item("quote_prefix", s.quote_prefix)?;
    d.set_item("pivot_button", s.pivot_button)?;
    Ok(d)
}

/// A differential (conditional formatting) style; only the parts it sets.
pub(crate) fn dxf_to_py<'py>(
    py: Python<'py>,
    s: &DifferentialStyle,
) -> PyResult<Bound<'py, PyDict>> {
    let d = PyDict::new(py);
    if let Some(f) = &s.font {
        d.set_item("font", font_to_py(py, f)?)?;
    }
    if let Some(f) = &s.fill {
        d.set_item("fill", fill_to_py(py, f)?)?;
    }
    if let Some(b) = &s.border {
        d.set_item("border", border_to_py(py, b)?)?;
    }
    if let Some(n) = &s.number_format {
        d.set_item("number_format", n)?;
    }
    if let Some(a) = &s.alignment {
        d.set_item("alignment", alignment_to_py(py, a)?)?;
    }
    Ok(d)
}

pub(crate) fn named_styles_to_py<'py>(
    py: Python<'py>,
    styles: &[NamedStyle],
) -> PyResult<Bound<'py, PyList>> {
    let list = PyList::empty(py);
    for s in styles {
        let d = PyDict::new(py);
        d.set_item("name", &s.name)?;
        d.set_item("builtin_id", s.builtin_id)?;
        d.set_item("hidden", s.hidden)?;
        list.append(d)?;
    }
    Ok(list)
}

pub(crate) fn rich_text_to_py<'py>(
    py: Python<'py>,
    runs: &[TextRun],
) -> PyResult<Bound<'py, PyList>> {
    let list = PyList::empty(py);
    for run in runs {
        let d = PyDict::new(py);
        d.set_item("text", &run.text)?;
        d.set_item(
            "font",
            run.font.as_ref().map(|f| font_to_py(py, f)).transpose()?,
        )?;
        list.append(d)?;
    }
    Ok(list)
}

pub(crate) fn hyperlink_to_py<'py>(py: Python<'py>, h: &Hyperlink) -> PyResult<Bound<'py, PyDict>> {
    let d = PyDict::new(py);
    d.set_item("range", (h.range.start, h.range.end))?;
    d.set_item("target", &h.target)?;
    d.set_item("location", &h.location)?;
    d.set_item("display", &h.displayed_text)?;
    d.set_item("tooltip", &h.tooltip)?;
    Ok(d)
}

pub(crate) fn row_info_to_py<'py>(
    py: Python<'py>,
    row: u32,
    attrs: Option<&RowAttributes>,
) -> PyResult<Bound<'py, PyDict>> {
    let default = RowAttributes::default();
    let a = attrs.unwrap_or(&default);
    let d = PyDict::new(py);
    d.set_item("row", row)?;
    d.set_item("height", a.height)?;
    d.set_item("custom_height", a.custom_height)?;
    d.set_item("hidden", a.hidden)?;
    d.set_item("outline_level", a.outline_level)?;
    d.set_item("collapsed", a.collapsed)?;
    d.set_item("style_id", a.style)?;
    Ok(d)
}

/// Rows with non-default attributes, keyed by row index.
pub(crate) fn row_dimensions_to_py<'py>(
    py: Python<'py>,
    rows: &[RowAttributes],
) -> PyResult<Bound<'py, PyDict>> {
    let d = PyDict::new(py);
    for r in rows {
        d.set_item(r.index, row_info_to_py(py, r.index, Some(r))?)?;
    }
    Ok(d)
}

fn column_to_py<'py>(py: Python<'py>, c: &ColumnInfo) -> PyResult<Bound<'py, PyDict>> {
    let d = PyDict::new(py);
    d.set_item("min", c.min)?;
    d.set_item("max", c.max)?;
    d.set_item("width", c.width)?;
    d.set_item("custom_width", c.custom_width)?;
    d.set_item("hidden", c.hidden)?;
    d.set_item("outline_level", c.outline_level)?;
    d.set_item("collapsed", c.collapsed)?;
    d.set_item("style_id", c.style)?;
    d.set_item("best_fit", c.best_fit)?;
    Ok(d)
}

pub(crate) fn columns_to_py<'py>(
    py: Python<'py>,
    cols: &[ColumnInfo],
) -> PyResult<Bound<'py, PyList>> {
    let list = PyList::empty(py);
    for c in cols {
        list.append(column_to_py(py, c)?)?;
    }
    Ok(list)
}

pub(crate) fn conditional_formats_to_py<'py>(
    py: Python<'py>,
    cfs: &[ConditionalFormat],
    dxfs: &[DifferentialStyle],
) -> PyResult<Bound<'py, PyList>> {
    let list = PyList::empty(py);
    for cf in cfs {
        let rules = PyList::empty(py);
        for r in &cf.rules {
            let d = PyDict::new(py);
            d.set_item("type", &r.typ)?;
            d.set_item("priority", r.priority)?;
            d.set_item("operator", &r.operator)?;
            d.set_item("formulas", &r.formulas)?;
            d.set_item("stop_if_true", r.stop_if_true)?;
            d.set_item("text", &r.text)?;
            d.set_item("dxf_id", r.dxf_id)?;
            d.set_item(
                "style",
                r.dxf_id
                    .and_then(|i| dxfs.get(i as usize))
                    .map(|s| dxf_to_py(py, s))
                    .transpose()?,
            )?;
            let values = PyList::empty(py);
            for v in &r.values {
                let vd = PyDict::new(py);
                vd.set_item("type", &v.typ)?;
                vd.set_item("value", &v.value)?;
                values.append(vd)?;
            }
            d.set_item("values", values)?;
            let colors = PyList::empty(py);
            for c in &r.colors {
                colors.append(color_to_py(py, &Some(c.clone()))?)?;
            }
            d.set_item("colors", colors)?;
            d.set_item("icon_set", &r.icon_set)?;
            d.set_item("extra", attrs_to_py(py, &r.extra)?)?;
            rules.append(d)?;
        }
        let d = PyDict::new(py);
        d.set_item("ranges", &cf.ranges)?;
        d.set_item("pivot", cf.pivot)?;
        d.set_item("rules", rules)?;
        list.append(d)?;
    }
    Ok(list)
}

pub(crate) fn data_validations_to_py<'py>(
    py: Python<'py>,
    dvs: &[DataValidation],
) -> PyResult<Bound<'py, PyList>> {
    let list = PyList::empty(py);
    for v in dvs {
        let d = PyDict::new(py);
        d.set_item("ranges", &v.ranges)?;
        d.set_item("type", &v.typ)?;
        d.set_item("operator", &v.operator)?;
        d.set_item("formula1", &v.formula1)?;
        d.set_item("formula2", &v.formula2)?;
        d.set_item("allow_blank", v.allow_blank)?;
        d.set_item("in_cell_dropdown", v.in_cell_dropdown)?;
        d.set_item("show_input_message", v.show_input_message)?;
        d.set_item("show_error_message", v.show_error_message)?;
        d.set_item("error_style", &v.error_style)?;
        d.set_item("error_title", &v.error_title)?;
        d.set_item("error", &v.error)?;
        d.set_item("prompt_title", &v.prompt_title)?;
        d.set_item("prompt", &v.prompt)?;
        d.set_item("ime_mode", &v.ime_mode)?;
        list.append(d)?;
    }
    Ok(list)
}

fn sort_state_to_py<'py>(
    py: Python<'py>,
    s: &Option<SortState>,
) -> PyResult<Option<Bound<'py, PyDict>>> {
    let Some(s) = s else { return Ok(None) };
    let d = attrs_to_py(py, &s.attrs)?;
    d.set_item("conditions", attrs_list_to_py(py, &s.conditions)?)?;
    Ok(Some(d))
}

pub(crate) fn auto_filter_to_py<'py>(
    py: Python<'py>,
    af: &Option<AutoFilter>,
) -> PyResult<Option<Bound<'py, PyDict>>> {
    let Some(af) = af else { return Ok(None) };
    let columns = PyList::empty(py);
    for c in &af.columns {
        let d = attrs_to_py(py, &c.attrs)?;
        d.set_item("column", c.col_id)?;
        d.set_item("values", &c.values)?;
        d.set_item("blank", c.blank)?;
        d.set_item("date_groups", attrs_list_to_py(py, &c.date_groups)?)?;
        d.set_item("custom_and", c.custom_and)?;
        d.set_item("custom", attrs_list_to_py(py, &c.custom)?)?;
        d.set_item("top10", opt_attrs_to_py(py, &c.top10)?)?;
        d.set_item("dynamic", opt_attrs_to_py(py, &c.dynamic)?)?;
        d.set_item("color", opt_attrs_to_py(py, &c.color)?)?;
        d.set_item("icon", opt_attrs_to_py(py, &c.icon)?)?;
        columns.append(d)?;
    }
    let d = PyDict::new(py);
    d.set_item("ref", &af.range)?;
    d.set_item("columns", columns)?;
    d.set_item("sort", sort_state_to_py(py, &af.sort)?)?;
    Ok(Some(d))
}

fn views_to_py<'py>(py: Python<'py>, views: &[SheetView]) -> PyResult<Bound<'py, PyList>> {
    let list = PyList::empty(py);
    for v in views {
        let d = attrs_to_py(py, &v.attrs)?;
        d.set_item("pane", opt_attrs_to_py(py, &v.pane)?)?;
        d.set_item("selections", attrs_list_to_py(py, &v.selections)?)?;
        list.append(d)?;
    }
    Ok(list)
}

fn header_footer_to_py<'py>(
    py: Python<'py>,
    hf: &Option<HeaderFooter>,
) -> PyResult<Option<Bound<'py, PyDict>>> {
    let Some(hf) = hf else { return Ok(None) };
    let d = attrs_to_py(py, &hf.attrs)?;
    for (k, v) in &hf.parts {
        d.set_item(snake(k), v)?;
    }
    Ok(Some(d))
}

pub(crate) fn scenarios_to_py<'py>(
    py: Python<'py>,
    s: &Option<Scenarios>,
) -> PyResult<Option<Bound<'py, PyDict>>> {
    let Some(s) = s else { return Ok(None) };
    let list = PyList::empty(py);
    for sc in &s.scenarios {
        let d = attrs_to_py(py, &sc.attrs)?;
        d.set_item("input_cells", attrs_list_to_py(py, &sc.input_cells)?)?;
        list.append(d)?;
    }
    let d = attrs_to_py(py, &s.attrs)?;
    d.set_item("scenarios", list)?;
    Ok(Some(d))
}

pub(crate) fn breaks_to_py<'py>(
    py: Python<'py>,
    breaks: &[Vec<(String, String)>],
) -> PyResult<Bound<'py, PyList>> {
    attrs_list_to_py(py, breaks)
}

pub(crate) fn tables_to_py<'py>(
    py: Python<'py>,
    tables: &[TableInfo],
) -> PyResult<Bound<'py, PyList>> {
    let list = PyList::empty(py);
    for t in tables {
        let d = attrs_to_py(py, &t.attrs)?;
        let cols = PyList::empty(py);
        for c in &t.columns {
            let cd = attrs_to_py(py, &c.attrs)?;
            cd.set_item("calculated_formula", &c.calculated_formula)?;
            cd.set_item("totals_formula", &c.totals_formula)?;
            cols.append(cd)?;
        }
        d.set_item("columns", cols)?;
        d.set_item("style", opt_attrs_to_py(py, &t.style)?)?;
        d.set_item("auto_filter", auto_filter_to_py(py, &t.auto_filter)?)?;
        list.append(d)?;
    }
    Ok(list)
}

/// Sheet-level settings. `print_area` / `print_titles` come from the sheet's
/// `_xlnm.Print_Area` / `_xlnm.Print_Titles` defined names.
pub(crate) fn sheet_settings_to_py<'py>(
    py: Python<'py>,
    info: &WorksheetInfo,
    print_area: Option<&str>,
    print_titles: Option<&str>,
) -> PyResult<Bound<'py, PyDict>> {
    let d = PyDict::new(py);
    let pane = info.freeze_pane.as_ref();
    d.set_item("freeze_panes", pane.and_then(|p| p.top_left_cell.clone()))?;
    d.set_item("frozen_rows", pane.map_or(0, |p| p.rows as u32))?;
    d.set_item("frozen_columns", pane.map_or(0, |p| p.cols as u32))?;
    d.set_item("auto_filter", &info.auto_filter)?;
    d.set_item(
        "auto_filter_details",
        auto_filter_to_py(py, &info.auto_filter_detail)?,
    )?;
    d.set_item("sort_state", sort_state_to_py(py, &info.sort_state)?)?;
    d.set_item("protection", opt_attrs_to_py(py, &info.protection)?)?;
    d.set_item("page_setup", attrs_to_py(py, &info.page_setup)?)?;
    d.set_item("page_margins", attrs_to_py(py, &info.page_margins)?)?;
    d.set_item("print_options", attrs_to_py(py, &info.print_options)?)?;
    d.set_item("print_area", print_area)?;
    d.set_item("print_titles", print_titles)?;
    d.set_item(
        "header_footer",
        header_footer_to_py(py, &info.header_footer)?,
    )?;
    d.set_item("sheet_view", attrs_to_py(py, &info.sheet_view)?)?;
    d.set_item("views", views_to_py(py, &info.views)?)?;
    d.set_item("custom_views", attrs_list_to_py(py, &info.custom_views)?)?;
    d.set_item("sheet_properties", attrs_to_py(py, &info.sheet_properties)?)?;
    d.set_item(
        "outline_properties",
        attrs_to_py(py, &info.outline_properties)?,
    )?;
    d.set_item(
        "page_setup_properties",
        attrs_to_py(py, &info.page_setup_properties)?,
    )?;
    d.set_item("sheet_format", attrs_to_py(py, &info.sheet_format)?)?;
    d.set_item("tab_color", color_to_py(py, &info.tab_color)?)?;
    d.set_item("default_row_height", info.default_row_height)?;
    d.set_item("default_column_width", info.default_column_width)?;
    Ok(d)
}

/// Custom document properties typed by their value element.
pub(crate) fn custom_properties_to_py<'py>(
    py: Python<'py>,
    props: &[CustomProperty],
) -> PyResult<Bound<'py, PyDict>> {
    let d = PyDict::new(py);
    for p in props {
        let v = &p.value;
        let value: Bound<'py, PyAny> = match p.kind.as_str() {
            "i1" | "i2" | "i4" | "i8" | "int" | "ui1" | "ui2" | "ui4" | "ui8" | "uint" => {
                match v.parse::<i64>() {
                    Ok(i) => i.into_pyobject(py)?.into_any(),
                    Err(_) => v.into_pyobject(py)?.into_any(),
                }
            }
            "r4" | "r8" | "decimal" => match v.parse::<f64>() {
                Ok(f) => f.into_pyobject(py)?.into_any(),
                Err(_) => v.into_pyobject(py)?.into_any(),
            },
            "bool" => PyBool::new(py, v == "true" || v == "1")
                .to_owned()
                .into_any(),
            "filetime" | "date" => match parse_w3c_datetime(v) {
                Some(dt) => dt.into_pyobject(py)?.into_any(),
                None => v.into_pyobject(py)?.into_any(),
            },
            _ => v.into_pyobject(py)?.into_any(),
        };
        d.set_item(&p.name, value)?;
    }
    Ok(d)
}

/// `2024-01-31T10:30:00Z` (or with offset / without zone) as naive UTC.
pub(crate) fn parse_w3c_datetime(v: &str) -> Option<chrono::NaiveDateTime> {
    chrono::DateTime::parse_from_rfc3339(v)
        .map(|dt| dt.naive_utc())
        .or_else(|_| chrono::NaiveDateTime::parse_from_str(v, "%Y-%m-%dT%H:%M:%S%.f"))
        .or_else(|_| chrono::NaiveDateTime::parse_from_str(v, "%Y-%m-%dT%H:%M:%S"))
        .ok()
}

/// `"B12"` -> `(11, 1)` (0-based row, column).
pub(crate) fn parse_a1(cell: &str) -> Option<(u32, u32)> {
    let cell = cell.trim().replace('$', "");
    let split = cell.find(|c: char| c.is_ascii_digit())?;
    let (letters, digits) = cell.split_at(split);
    if letters.is_empty() || !letters.chars().all(|c| c.is_ascii_alphabetic()) {
        return None;
    }
    let mut col: u32 = 0;
    for c in letters.chars() {
        col = col
            .checked_mul(26)?
            .checked_add(c.to_ascii_uppercase() as u32 - 'A' as u32 + 1)?;
    }
    let row: u32 = digits.parse().ok()?;
    Some((row.checked_sub(1)?, col - 1))
}

/// 0-based column -> letters (`0` -> `A`, `27` -> `AB`).
pub(crate) fn column_letters(mut col: u32) -> String {
    let mut out = Vec::new();
    loop {
        out.push(b'A' + (col % 26) as u8);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    out.reverse();
    String::from_utf8(out).expect("ascii")
}
