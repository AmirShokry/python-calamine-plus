// Fork addition (python-calamine streaming): everything openpyxl can read that
// upstream calamine does not: full cell styles (and conditional-format styles),
// named styles, comments, scoped defined names, document and custom properties,
// workbook settings and theme, rich text runs, tables, and worksheet metadata
// (hyperlinks, conditional formatting, data validation, column/row dimensions,
// views, auto filter, protection, page setup, header/footer, breaks, ...).

use std::collections::HashMap;
use std::io::{Read, Seek};

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::QName;
use zip::result::ZipError;

use super::cells_reader::{parse_row, RowAttributes};
use super::{
    get_row_column, read_hyperlinks, read_merge_cells, read_sheet_hyperlink_rels,
    read_string_with_bufs, xml_reader, Hyperlink, XlReader, Xlsx, XlsxError,
};
use crate::attrs::{decode_attr, RawAttributes};
use crate::utils::{cached_zip_path, unescape_entity_to_buffer, unescape_xml};
use crate::Dimensions;

/// Raw attributes of an element as `(name, value)` pairs (namespace prefixes removed).
pub type Attributes = Vec<(String, String)>;

/// A color reference as stored in the file. Theme/indexed colors are not resolved.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct StyleColor {
    /// ARGB hex, e.g. `FFFF0000`.
    pub rgb: Option<String>,
    /// Theme color index.
    pub theme: Option<u32>,
    /// Legacy indexed palette color.
    pub indexed: Option<u32>,
    /// Tint applied to a theme color (-1.0..=1.0).
    pub tint: Option<f64>,
    /// Automatic (system) color.
    pub auto: bool,
}

/// Font of a cell format or text run.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct StyleFont {
    /// Font name.
    pub name: Option<String>,
    /// Size in points.
    pub size: Option<f64>,
    /// Bold.
    pub bold: bool,
    /// Italic.
    pub italic: bool,
    /// Underline style (`single`, `double`, ...), `None` when not underlined.
    pub underline: Option<String>,
    /// Strikethrough.
    pub strike: bool,
    /// `superscript` / `subscript`, `None` for baseline.
    pub vert_align: Option<String>,
    /// Font color.
    pub color: Option<StyleColor>,
    /// Font family (1 = roman, 2 = swiss, ...).
    pub family: Option<u32>,
    /// Character set.
    pub charset: Option<u32>,
    /// Theme font scheme (`major` / `minor`).
    pub scheme: Option<String>,
    /// Outline.
    pub outline: bool,
    /// Shadow.
    pub shadow: bool,
    /// Condense.
    pub condense: bool,
    /// Extend.
    pub extend: bool,
    /// Fork addition (openpyxl compatibility): `strike` as stored, `None` when absent.
    pub raw_strike: Option<bool>,
    /// Fork addition (openpyxl compatibility): `outline` as stored, `None` when absent.
    pub raw_outline: Option<bool>,
    /// Fork addition (openpyxl compatibility): `shadow` as stored, `None` when absent.
    pub raw_shadow: Option<bool>,
    /// Fork addition (openpyxl compatibility): `condense` as stored, `None` when absent.
    pub raw_condense: Option<bool>,
    /// Fork addition (openpyxl compatibility): `extend` as stored, `None` when absent.
    pub raw_extend: Option<bool>,
    /// Fork addition (openpyxl compatibility): `vertAlign` as stored (including `baseline`).
    pub raw_vert_align: Option<String>,
}

/// Gradient of a gradient fill.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct StyleGradient {
    /// `linear` or `path`.
    pub typ: Option<String>,
    /// Angle of a linear gradient.
    pub degree: Option<f64>,
    /// Path gradient bounds.
    pub left: Option<f64>,
    /// Path gradient bounds.
    pub right: Option<f64>,
    /// Path gradient bounds.
    pub top: Option<f64>,
    /// Path gradient bounds.
    pub bottom: Option<f64>,
    /// `(position, color)` stops.
    pub stops: Vec<(f64, StyleColor)>,
}

/// Fill of a cell format.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct StyleFill {
    /// Pattern type (`solid`, `gray125`, ...) or `gradient`; `None` when there is no fill.
    pub pattern: Option<String>,
    /// Foreground color (the visible color of a `solid` fill).
    pub fg_color: Option<StyleColor>,
    /// Background color.
    pub bg_color: Option<StyleColor>,
    /// Gradient details when `pattern` is `gradient`.
    pub gradient: Option<StyleGradient>,
}

/// One edge of a border.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct StyleBorderSide {
    /// Line style (`thin`, `medium`, `dashed`, ...).
    pub style: Option<String>,
    /// Line color.
    pub color: Option<StyleColor>,
}

/// Border of a cell format.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct StyleBorder {
    /// Left edge.
    pub left: StyleBorderSide,
    /// Right edge.
    pub right: StyleBorderSide,
    /// Top edge.
    pub top: StyleBorderSide,
    /// Bottom edge.
    pub bottom: StyleBorderSide,
    /// Diagonal line.
    pub diagonal: StyleBorderSide,
    /// Inner vertical edges (differential formats).
    pub vertical: StyleBorderSide,
    /// Inner horizontal edges (differential formats).
    pub horizontal: StyleBorderSide,
    /// Diagonal from bottom-left to top-right.
    pub diagonal_up: bool,
    /// Diagonal from top-left to bottom-right.
    pub diagonal_down: bool,
    /// Apply to the outline of a range.
    pub outline: bool,
    /// Fork addition (openpyxl compatibility): edges present in the file, as bits
    /// (left 1, right 2, top 4, bottom 8, diagonal 16, vertical 32, horizontal 64).
    pub raw_sides: u8,
    /// Fork addition (openpyxl compatibility): `outline` as stored, `None` when absent.
    pub raw_outline: Option<bool>,
}

/// Alignment of a cell format.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct StyleAlignment {
    /// Horizontal alignment.
    pub horizontal: Option<String>,
    /// Vertical alignment.
    pub vertical: Option<String>,
    /// Wrap text.
    pub wrap_text: bool,
    /// Shrink to fit.
    pub shrink_to_fit: bool,
    /// Indent level.
    pub indent: Option<u32>,
    /// Text rotation in degrees (255 = vertical text).
    pub text_rotation: Option<u32>,
    /// Justify the last line of distributed text.
    pub justify_last_line: bool,
    /// Reading order (0 = context, 1 = left-to-right, 2 = right-to-left).
    pub reading_order: Option<u32>,
    /// Relative indent.
    pub relative_indent: Option<i32>,
    /// Fork addition (openpyxl compatibility): `wrapText` as stored, `None` when absent.
    pub raw_wrap_text: Option<bool>,
    /// Fork addition (openpyxl compatibility): `shrinkToFit` as stored, `None` when absent.
    pub raw_shrink_to_fit: Option<bool>,
    /// Fork addition (openpyxl compatibility): `justifyLastLine` as stored, `None` when absent.
    pub raw_justify_last_line: Option<bool>,
}

/// A resolved cell format (`<xf>` in `cellXfs`). A cell's `s` attribute indexes these.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CellStyle {
    /// Number format id.
    pub number_format_id: u32,
    /// Number format code, e.g. `0.00%` (built-in ids are resolved to their code).
    pub number_format: String,
    /// Font.
    pub font: StyleFont,
    /// Fill.
    pub fill: StyleFill,
    /// Border.
    pub border: StyleBorder,
    /// Alignment.
    pub alignment: StyleAlignment,
    /// Locked (protection); Excel's default is `true`.
    pub locked: bool,
    /// Hidden formula (protection).
    pub hidden: bool,
    /// Named style the format is based on (`Normal`, `Heading 1`, ...).
    pub named_style: Option<String>,
    /// Text is prefixed with `'` in the formula bar.
    pub quote_prefix: bool,
    /// Pivot table button.
    pub pivot_button: bool,
    /// Fork addition (openpyxl compatibility): the number format is defined in the
    /// file's `<numFmts>` (not a built-in id).
    pub number_format_custom: bool,
}

/// A differential format (`<dxf>`), applied by conditional formatting rules.
/// Only the parts it sets are `Some`.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct DifferentialStyle {
    /// Font override.
    pub font: Option<StyleFont>,
    /// Fill override. Note: Excel stores a solid dxf fill's color in `bg_color`.
    pub fill: Option<StyleFill>,
    /// Border override.
    pub border: Option<StyleBorder>,
    /// Number format override.
    pub number_format: Option<String>,
    /// Alignment override.
    pub alignment: Option<StyleAlignment>,
}

/// A named cell style (`<cellStyle>`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct NamedStyle {
    /// Name, e.g. `Normal`.
    pub name: String,
    /// Built-in style id, if built in.
    pub builtin_id: Option<u32>,
    /// Hidden from the styles gallery.
    pub hidden: bool,
}

/// Cell formats, differential formats and named styles of a workbook.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct StyleSheet {
    /// Cell formats, indexed by a cell's style id.
    pub cell_styles: Vec<CellStyle>,
    /// Differential formats, indexed by a conditional formatting rule's `dxf_id`.
    pub differential_styles: Vec<DifferentialStyle>,
    /// Named styles.
    pub named_styles: Vec<NamedStyle>,
    /// Fork addition (openpyxl compatibility): the format of a cell without a style
    /// array (first font, fill and border, default alignment and protection), which
    /// openpyxl gives to cells it creates; `None` without a stylesheet.
    pub zero_style: Option<CellStyle>,
}

/// A cell comment (legacy note; threaded comments are exported by Excel as notes too).
#[derive(Debug, Clone, PartialEq)]
pub struct Comment {
    /// Cell position (row, column), 0-based.
    pub pos: (u32, u32),
    /// Author, if any.
    pub author: Option<String>,
    /// Plain text (rich text runs are concatenated).
    pub text: String,
}

/// A defined name with its scope (calamine's `defined_names` drops the scope).
#[derive(Debug, Clone, PartialEq)]
pub struct DefinedName {
    /// Name, e.g. `TaxRate` or `_xlnm.Print_Area`.
    pub name: String,
    /// Formula / reference, e.g. `Assumptions!$B$2`.
    pub value: String,
    /// Sheet the name is local to, `None` for workbook scope.
    pub sheet: Option<String>,
    /// Hidden name.
    pub hidden: bool,
    /// Comment.
    pub comment: Option<String>,
}

/// A formatted run of a rich text string.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct TextRun {
    /// Text of the run.
    pub text: String,
    /// Run font, `None` when the run uses the cell's font.
    pub font: Option<StyleFont>,
}

/// Width/visibility settings for a span of columns (`<col>`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ColumnInfo {
    /// First column (0-based).
    pub min: u32,
    /// Last column (0-based, inclusive).
    pub max: u32,
    /// Width in characters.
    pub width: Option<f64>,
    /// Width was set explicitly.
    pub custom_width: bool,
    /// Hidden.
    pub hidden: bool,
    /// Outline (grouping) level.
    pub outline_level: u8,
    /// Outline group collapsed.
    pub collapsed: bool,
    /// Default style id for the columns.
    pub style: Option<u32>,
    /// Width fits the widest cell.
    pub best_fit: bool,
}

/// A conditional formatting threshold (`<cfvo>`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ConditionalValue {
    /// `min`, `max`, `num`, `percent`, `percentile`, `formula`.
    pub typ: String,
    /// Threshold value or formula.
    pub value: Option<String>,
}

/// A conditional formatting rule (`<cfRule>`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ConditionalRule {
    /// `cellIs`, `expression`, `colorScale`, `dataBar`, `iconSet`, `top10`, ...
    pub typ: Option<String>,
    /// Priority (lower runs first).
    pub priority: Option<i64>,
    /// Operator for `cellIs` rules (`greaterThan`, `between`, ...).
    pub operator: Option<String>,
    /// Formulas (without `=`).
    pub formulas: Vec<String>,
    /// Index into `StyleSheet::differential_styles`.
    pub dxf_id: Option<u32>,
    /// Stop evaluating lower-priority rules when this one applies.
    pub stop_if_true: bool,
    /// Text for `containsText`-style rules.
    pub text: Option<String>,
    /// Thresholds of color scales, data bars and icon sets.
    pub values: Vec<ConditionalValue>,
    /// Colors of color scales and data bars.
    pub colors: Vec<StyleColor>,
    /// Icon set name (`3TrafficLights1`, ...).
    pub icon_set: Option<String>,
    /// Any other attributes of the rule and its color scale / data bar / icon set.
    pub extra: Attributes,
}

/// Conditional formatting applied to a set of ranges.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ConditionalFormat {
    /// Ranges, e.g. `["A1:A10", "C1"]`.
    pub ranges: Vec<String>,
    /// Pivot table conditional format.
    pub pivot: bool,
    /// Rules, in file order.
    pub rules: Vec<ConditionalRule>,
}

/// A data validation rule (`<dataValidation>`), including the extended form
/// Excel uses for lists that reference other sheets.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct DataValidation {
    /// Ranges it applies to.
    pub ranges: Vec<String>,
    /// `list`, `whole`, `decimal`, `date`, `time`, `textLength`, `custom`, or `None` (any value).
    pub typ: Option<String>,
    /// `between`, `greaterThan`, ...
    pub operator: Option<String>,
    /// First formula (list source, bound, ...), without `=`.
    pub formula1: Option<String>,
    /// Second formula (upper bound of `between`).
    pub formula2: Option<String>,
    /// Blank values are allowed.
    pub allow_blank: bool,
    /// The in-cell dropdown is shown (Excel stores the inverse as `showDropDown`).
    pub in_cell_dropdown: bool,
    /// Show the input message.
    pub show_input_message: bool,
    /// Show the error message.
    pub show_error_message: bool,
    /// `stop`, `warning` or `information`.
    pub error_style: Option<String>,
    /// Error title.
    pub error_title: Option<String>,
    /// Error message.
    pub error: Option<String>,
    /// Input prompt title.
    pub prompt_title: Option<String>,
    /// Input prompt.
    pub prompt: Option<String>,
    /// IME mode.
    pub ime_mode: Option<String>,
}

/// Frozen panes of a sheet view.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct FreezePane {
    /// First scrollable cell, e.g. `B2` (openpyxl's `freeze_panes`).
    pub top_left_cell: Option<String>,
    /// Frozen columns.
    pub cols: f64,
    /// Frozen rows.
    pub rows: f64,
}

/// A sheet view (`<sheetView>`) with its pane and selections.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct SheetView {
    /// View attributes (zoom, gridlines, ...).
    pub attrs: Attributes,
    /// Pane attributes (split / frozen panes).
    pub pane: Option<Attributes>,
    /// Selections (`activeCell`, `sqref`, `pane`).
    pub selections: Vec<Attributes>,
}

/// One column of an auto filter (`<filterColumn>`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct FilterColumn {
    /// Column index within the filter range (0-based).
    pub col_id: u32,
    /// Other attributes (`hiddenButton`, `showButton`).
    pub attrs: Attributes,
    /// Values filter: selected values.
    pub values: Vec<String>,
    /// Values filter: blanks are selected.
    pub blank: bool,
    /// Values filter: date group items.
    pub date_groups: Vec<Attributes>,
    /// Custom filters are combined with AND (otherwise OR).
    pub custom_and: bool,
    /// Custom filters (`operator`, `val`).
    pub custom: Vec<Attributes>,
    /// Top 10 filter.
    pub top10: Option<Attributes>,
    /// Dynamic filter (`aboveAverage`, `thisMonth`, ...).
    pub dynamic: Option<Attributes>,
    /// Color filter.
    pub color: Option<Attributes>,
    /// Icon filter.
    pub icon: Option<Attributes>,
}

/// Sort state of an auto filter or sheet.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct SortState {
    /// Attributes (`ref`, `caseSensitive`, ...).
    pub attrs: Attributes,
    /// Sort conditions (`ref`, `descending`, `sortBy`, ...).
    pub conditions: Vec<Attributes>,
}

/// An auto filter with its filter columns and sort state.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct AutoFilter {
    /// Filtered range.
    pub range: Option<String>,
    /// Filter columns.
    pub columns: Vec<FilterColumn>,
    /// Sort state.
    pub sort: Option<SortState>,
}

/// Header and footer texts (with Excel's `&`-codes) and options.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct HeaderFooter {
    /// Options (`differentOddEven`, `differentFirst`, `scaleWithDoc`, `alignWithMargins`).
    pub attrs: Attributes,
    /// `(element, text)` pairs: `oddHeader`, `oddFooter`, `evenHeader`, ...
    pub parts: Vec<(String, String)>,
}

/// A what-if scenario.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Scenario {
    /// Attributes (`name`, `locked`, `hidden`, `user`, `comment`).
    pub attrs: Attributes,
    /// Input cells (`r`, `val`, ...).
    pub input_cells: Vec<Attributes>,
}

/// Scenarios of a sheet.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Scenarios {
    /// Attributes (`current`, `show`, `sqref`).
    pub attrs: Attributes,
    /// Scenarios.
    pub scenarios: Vec<Scenario>,
}

/// A table (list object) column.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct TableColumn {
    /// Attributes (`id`, `name`, `totalsRowFunction`, `totalsRowLabel`, ...).
    pub attrs: Attributes,
    /// Calculated column formula.
    pub calculated_formula: Option<String>,
    /// Totals row formula.
    pub totals_formula: Option<String>,
}

/// A table (list object) of a worksheet.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct TableInfo {
    /// Attributes (`name`, `displayName`, `ref`, `headerRowCount`, `totalsRowCount`, ...).
    pub attrs: Attributes,
    /// Columns.
    pub columns: Vec<TableColumn>,
    /// Table style (`name`, `showRowStripes`, ...).
    pub style: Option<Attributes>,
    /// Auto filter of the table.
    pub auto_filter: Option<AutoFilter>,
}

/// Everything a worksheet stores outside its cell data.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct WorksheetInfo {
    /// `<dimension>` reference, e.g. `A1:H100` (as declared by the writer).
    pub dimension: Option<String>,
    /// Merged ranges.
    pub merged: Vec<Dimensions>,
    /// Hyperlinks, with external targets resolved.
    pub hyperlinks: Vec<Hyperlink>,
    /// Conditional formatting.
    pub conditional_formats: Vec<ConditionalFormat>,
    /// Data validation.
    pub data_validations: Vec<DataValidation>,
    /// Column widths / visibility / outline.
    pub columns: Vec<ColumnInfo>,
    /// Rows with non-default attributes (height, hidden, outline, style).
    pub rows: Vec<RowAttributes>,
    /// Frozen panes of the first view.
    pub freeze_pane: Option<FreezePane>,
    /// Auto filter range.
    pub auto_filter: Option<String>,
    /// Auto filter with columns and sort state.
    pub auto_filter_detail: Option<AutoFilter>,
    /// Sheet-level sort state (outside an auto filter).
    pub sort_state: Option<SortState>,
    /// Sheet protection attributes, as stored.
    pub protection: Option<Attributes>,
    /// Page setup attributes, as stored.
    pub page_setup: Attributes,
    /// Page margins attributes, as stored (inches).
    pub page_margins: Attributes,
    /// Print options attributes, as stored.
    pub print_options: Attributes,
    /// Header / footer.
    pub header_footer: Option<HeaderFooter>,
    /// Manual row page breaks (`id`, `min`, `max`, `man`, `pt`).
    pub row_breaks: Vec<Attributes>,
    /// Manual column page breaks.
    pub column_breaks: Vec<Attributes>,
    /// First sheet view attributes (zoom, gridlines, ...), as stored.
    pub sheet_view: Attributes,
    /// All sheet views.
    pub views: Vec<SheetView>,
    /// Custom sheet views (`guid`, `scale`, ...).
    pub custom_views: Vec<Attributes>,
    /// `<sheetPr>` attributes (`codeName`, `filterMode`, ...).
    pub sheet_properties: Attributes,
    /// `<outlinePr>` attributes (`summaryBelow`, `summaryRight`, ...).
    pub outline_properties: Attributes,
    /// `<pageSetUpPr>` attributes (`fitToPage`, `autoPageBreaks`).
    pub page_setup_properties: Attributes,
    /// `<sheetFormatPr>` attributes.
    pub sheet_format: Attributes,
    /// Sheet tab color.
    pub tab_color: Option<StyleColor>,
    /// Default row height (points).
    pub default_row_height: Option<f64>,
    /// Default column width (characters).
    pub default_column_width: Option<f64>,
    /// What-if scenarios.
    pub scenarios: Option<Scenarios>,
    /// Tables.
    pub tables: Vec<TableInfo>,
    /// Fork addition (openpyxl compatibility): bounding box (0-based) of all `<c>`
    /// elements, only filled by `Xlsx::worksheet_info_with_bounds`.
    pub cell_bounds: Option<Dimensions>,
}

/// A custom document property (`docProps/custom.xml`).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CustomProperty {
    /// Name.
    pub name: String,
    /// Value type element (`lpwstr`, `i4`, `r8`, `bool`, `filetime`, ...).
    pub kind: String,
    /// Value as stored.
    pub value: String,
}

/// Workbook-level settings openpyxl exposes.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct WorkbookInfo {
    /// `<workbookPr>` attributes (`date1904`, `codeName`, ...).
    pub workbook_properties: Attributes,
    /// `<workbookView>` attributes (`activeTab`, `firstSheet`, ...).
    pub views: Vec<Attributes>,
    /// `<calcPr>` attributes (`calcId`, `fullCalcOnLoad`, ...).
    pub calculation: Option<Attributes>,
    /// `<workbookProtection>` attributes.
    pub protection: Option<Attributes>,
    /// `<fileVersion>` attributes.
    pub file_version: Option<Attributes>,
    /// Custom document properties.
    pub custom_properties: Vec<CustomProperty>,
    /// Raw theme XML (openpyxl's `loaded_theme`).
    pub theme_xml: Option<Vec<u8>>,
    /// Theme color scheme as `(slot, rgb)`: `dk1`, `lt1`, `dk2`, `lt2`, `accent1`..`accent6`, `hlink`, `folHlink`.
    pub theme_colors: Vec<(String, String)>,
    /// The file is a template (`.xltx` / `.xltm`).
    pub is_template: bool,
    /// The file contains macros (`.xlsm` / `.xltm`).
    pub has_macros: bool,
}

/// Format codes of the built-in number formats (ECMA-376 18.8.30).
fn builtin_number_format(id: u32) -> &'static str {
    match id {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
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
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => "General",
    }
}

fn attr<RS: Read + Seek>(
    xml: &XlReader<'_, RS>,
    e: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>, XlsxError> {
    match e.raw_attr(name)? {
        Some(v) => Ok(Some(decode_attr(&xml.decoder(), v)?)),
        None => Ok(None),
    }
}

fn attr_num<RS: Read + Seek, T: std::str::FromStr>(
    xml: &XlReader<'_, RS>,
    e: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<T>, XlsxError> {
    Ok(attr(xml, e, name)?.and_then(|v| v.parse().ok()))
}

fn attr_bool<RS: Read + Seek>(
    xml: &XlReader<'_, RS>,
    e: &BytesStart<'_>,
    name: &[u8],
) -> Result<bool, XlsxError> {
    Ok(matches!(
        attr(xml, e, name)?.as_deref(),
        Some("1") | Some("true")
    ))
}

/// All attributes of an element, keys without namespace prefix.
fn all_attrs<RS: Read + Seek>(
    xml: &XlReader<'_, RS>,
    e: &BytesStart<'_>,
) -> Result<Attributes, XlsxError> {
    let mut out = Vec::new();
    for item in e.iter_raw_attrs() {
        let (k, v) = item?;
        if k.starts_with(b"xmlns") {
            continue;
        }
        let k = match k.iter().position(|&b| b == b':') {
            Some(i) => &k[i + 1..],
            None => k,
        };
        out.push((
            xml.decoder().decode(k)?.into_owned(),
            decode_attr(&xml.decoder(), v)?,
        ));
    }
    Ok(out)
}

/// Boolean element like `<b/>` or `<b val="0"/>`.
fn flag<RS: Read + Seek>(xml: &XlReader<'_, RS>, e: &BytesStart<'_>) -> Result<bool, XlsxError> {
    Ok(!matches!(
        attr(xml, e, b"val")?.as_deref(),
        Some("0") | Some("false")
    ))
}

fn color<RS: Read + Seek>(
    xml: &XlReader<'_, RS>,
    e: &BytesStart<'_>,
) -> Result<StyleColor, XlsxError> {
    Ok(StyleColor {
        rgb: attr(xml, e, b"rgb")?,
        theme: attr_num(xml, e, b"theme")?,
        indexed: attr_num(xml, e, b"indexed")?,
        tint: attr_num(xml, e, b"tint")?,
        auto: attr_bool(xml, e, b"auto")?,
    })
}

/// Reads the text content of an element (including nested elements) up to its
/// closing tag, without trimming.
fn read_plain_text<RS: Read + Seek>(
    xml: &mut XlReader<'_, RS>,
    closing: &[u8],
    buf: &mut Vec<u8>,
) -> Result<String, XlsxError> {
    let mut value = String::new();
    loop {
        buf.clear();
        match xml.read_event_into(buf)? {
            Event::Text(t) => value.push_str(&t.xml10_content()?),
            Event::CData(t) => value.push_str(&t.xml10_content()?),
            Event::GeneralRef(e) => unescape_entity_to_buffer(&e, &mut value)?,
            Event::End(e) if e.local_name().as_ref() == closing => break,
            Event::Eof => return Err(XlsxError::XmlEof("text")),
            _ => {}
        }
    }
    Ok(unescape_xml(&value).into_owned())
}

/// Reads the text of the element just opened, up to its closing tag.
fn element_text<RS: Read + Seek>(
    xml: &mut XlReader<'_, RS>,
    e: &BytesStart<'_>,
    buf: &mut Vec<u8>,
) -> Result<String, XlsxError> {
    let name = e.local_name().as_ref().to_vec();
    read_plain_text(xml, &name, buf)
}

/// Normalizes a relationship target relative to the folder of its source part.
fn resolve_target(base_folder: &str, target: &str) -> String {
    if let Some(abs) = target.strip_prefix('/') {
        return abs.to_string();
    }
    let mut parts: Vec<&str> = base_folder.split('/').filter(|p| !p.is_empty()).collect();
    for seg in target.split('/') {
        match seg {
            ".." => {
                parts.pop();
            }
            "." | "" => {}
            s => parts.push(s),
        }
    }
    parts.join("/")
}

fn split_ranges(s: &str) -> Vec<String> {
    s.split_whitespace().map(str::to_string).collect()
}

/// Applies a child element of `<font>` / `<rPr>`.
fn apply_font<RS: Read + Seek>(
    font: &mut StyleFont,
    xml: &XlReader<'_, RS>,
    e: &BytesStart<'_>,
    tag: &[u8],
) -> Result<(), XlsxError> {
    match tag {
        b"b" => font.bold = flag(xml, e)?,
        b"i" => font.italic = flag(xml, e)?,
        b"strike" => {
            font.strike = flag(xml, e)?;
            font.raw_strike = Some(font.strike);
        }
        b"outline" => {
            font.outline = flag(xml, e)?;
            font.raw_outline = Some(font.outline);
        }
        b"shadow" => {
            font.shadow = flag(xml, e)?;
            font.raw_shadow = Some(font.shadow);
        }
        b"condense" => {
            font.condense = flag(xml, e)?;
            font.raw_condense = Some(font.condense);
        }
        b"extend" => {
            font.extend = flag(xml, e)?;
            font.raw_extend = Some(font.extend);
        }
        b"u" => {
            let v = attr(xml, e, b"val")?.unwrap_or_else(|| "single".to_string());
            font.underline = (v != "none").then_some(v);
        }
        b"vertAlign" => {
            font.raw_vert_align = attr(xml, e, b"val")?;
            font.vert_align = font.raw_vert_align.clone().filter(|v| v != "baseline");
        }
        b"sz" => font.size = attr_num(xml, e, b"val")?,
        b"name" | b"rFont" => font.name = attr(xml, e, b"val")?,
        b"family" => font.family = attr_num(xml, e, b"val")?,
        b"charset" => font.charset = attr_num(xml, e, b"val")?,
        b"scheme" => font.scheme = attr(xml, e, b"val")?.filter(|v| v != "none"),
        b"color" => font.color = Some(color(xml, e)?),
        _ => {}
    }
    Ok(())
}

/// Applies a child element of `<fill>`.
fn apply_fill<RS: Read + Seek>(
    fill: &mut StyleFill,
    xml: &XlReader<'_, RS>,
    e: &BytesStart<'_>,
    tag: &[u8],
) -> Result<(), XlsxError> {
    match tag {
        b"patternFill" => {
            fill.pattern = attr(xml, e, b"patternType")?.filter(|p| p != "none");
            // A dxf patternFill without patternType means solid.
            if fill.pattern.is_none() && e.raw_attr(b"patternType")?.is_none() {
                fill.pattern = Some("solid".to_string());
            }
        }
        b"gradientFill" => {
            fill.pattern = Some("gradient".to_string());
            fill.gradient = Some(StyleGradient {
                typ: Some(attr(xml, e, b"type")?.unwrap_or_else(|| "linear".to_string())),
                degree: attr_num(xml, e, b"degree")?,
                left: attr_num(xml, e, b"left")?,
                right: attr_num(xml, e, b"right")?,
                top: attr_num(xml, e, b"top")?,
                bottom: attr_num(xml, e, b"bottom")?,
                stops: Vec::new(),
            });
        }
        b"stop" => {
            if let Some(g) = fill.gradient.as_mut() {
                let pos = attr_num(xml, e, b"position")?.unwrap_or(0.0);
                g.stops.push((pos, StyleColor::default()));
            }
        }
        b"color" => {
            if let Some(stop) = fill.gradient.as_mut().and_then(|g| g.stops.last_mut()) {
                stop.1 = color(xml, e)?;
            }
        }
        b"fgColor" => fill.fg_color = Some(color(xml, e)?),
        b"bgColor" => fill.bg_color = Some(color(xml, e)?),
        _ => {}
    }
    Ok(())
}

fn border_side<'a>(border: &'a mut StyleBorder, tag: &[u8]) -> Option<&'a mut StyleBorderSide> {
    match tag {
        b"left" | b"start" => Some(&mut border.left),
        b"right" | b"end" => Some(&mut border.right),
        b"top" => Some(&mut border.top),
        b"bottom" => Some(&mut border.bottom),
        b"diagonal" => Some(&mut border.diagonal),
        b"vertical" => Some(&mut border.vertical),
        b"horizontal" => Some(&mut border.horizontal),
        _ => None,
    }
}

fn new_border<RS: Read + Seek>(
    xml: &XlReader<'_, RS>,
    e: &BytesStart<'_>,
) -> Result<StyleBorder, XlsxError> {
    let outline = attr(xml, e, b"outline")?;
    Ok(StyleBorder {
        diagonal_up: attr_bool(xml, e, b"diagonalUp")?,
        diagonal_down: attr_bool(xml, e, b"diagonalDown")?,
        outline: matches!(outline.as_deref(), Some("1") | Some("true")),
        raw_outline: outline.map(|v| v == "1" || v == "true"),
        ..Default::default()
    })
}

/// Applies a child element of `<border>`; `side` tracks the edge being parsed.
fn apply_border<RS: Read + Seek>(
    border: &mut StyleBorder,
    side: &mut Option<Vec<u8>>,
    xml: &XlReader<'_, RS>,
    e: &BytesStart<'_>,
    tag: &[u8],
) -> Result<(), XlsxError> {
    if let Some(s) = border_side(border, tag) {
        s.style = attr(xml, e, b"style")?;
        *side = Some(tag.to_vec());
        border.raw_sides |= match tag {
            b"left" | b"start" => 1,
            b"right" | b"end" => 2,
            b"top" => 4,
            b"bottom" => 8,
            b"diagonal" => 16,
            b"vertical" => 32,
            _ => 64,
        };
    } else if tag == b"color" {
        if let Some(s) = side.as_deref().and_then(|t| border_side(border, t)) {
            s.color = Some(color(xml, e)?);
        }
    }
    Ok(())
}

fn alignment<RS: Read + Seek>(
    xml: &XlReader<'_, RS>,
    e: &BytesStart<'_>,
) -> Result<StyleAlignment, XlsxError> {
    let tristate = |name: &[u8]| -> Result<Option<bool>, XlsxError> {
        Ok(attr(xml, e, name)?.map(|v| v == "1" || v == "true"))
    };
    Ok(StyleAlignment {
        horizontal: attr(xml, e, b"horizontal")?,
        vertical: attr(xml, e, b"vertical")?,
        wrap_text: attr_bool(xml, e, b"wrapText")?,
        shrink_to_fit: attr_bool(xml, e, b"shrinkToFit")?,
        indent: attr_num(xml, e, b"indent")?,
        text_rotation: attr_num(xml, e, b"textRotation")?,
        justify_last_line: attr_bool(xml, e, b"justifyLastLine")?,
        reading_order: attr_num(xml, e, b"readingOrder")?,
        relative_indent: attr_num(xml, e, b"relativeIndent")?,
        raw_wrap_text: tristate(b"wrapText")?,
        raw_shrink_to_fit: tristate(b"shrinkToFit")?,
        raw_justify_last_line: tristate(b"justifyLastLine")?,
    })
}

/// Which `styles.xml` collection the parser is inside.
#[derive(Clone, Copy, PartialEq)]
enum Section {
    None,
    Fonts,
    Fills,
    Borders,
    CellXfs,
    Dxfs,
    CellStyles,
    Other,
}

/// Which part of a `<dxf>` the parser is inside.
#[derive(Clone, Copy, PartialEq)]
enum DxfPart {
    None,
    Font,
    Fill,
    Border,
}

#[derive(Default)]
struct RawXf {
    num_fmt_id: u32,
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    xf_id: Option<usize>,
    alignment: StyleAlignment,
    locked: bool,
    hidden: bool,
    quote_prefix: bool,
    pivot_button: bool,
}

/// Parses an `<autoFilter>` element's children up to its end.
fn read_auto_filter<RS: Read + Seek>(
    xml: &mut XlReader<'_, RS>,
    e: &BytesStart<'_>,
) -> Result<AutoFilter, XlsxError> {
    let mut af = AutoFilter {
        range: attr(xml, e, b"ref")?,
        ..Default::default()
    };
    let mut buf = Vec::new();
    let mut in_sort = false;
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"filterColumn" => {
                        let mut attrs = all_attrs(xml, &e)?;
                        let col_id = attrs
                            .iter()
                            .position(|(k, _)| k == "colId")
                            .map(|i| attrs.remove(i).1.parse().unwrap_or(0))
                            .unwrap_or(0);
                        af.columns.push(FilterColumn {
                            col_id,
                            attrs,
                            ..Default::default()
                        });
                    }
                    b"sortState" => {
                        af.sort = Some(SortState {
                            attrs: all_attrs(xml, &e)?,
                            conditions: Vec::new(),
                        });
                        in_sort = true;
                    }
                    b"sortCondition" if in_sort => {
                        let c = all_attrs(xml, &e)?;
                        if let Some(s) = af.sort.as_mut() {
                            s.conditions.push(c);
                        }
                    }
                    tag => {
                        let Some(col) = af.columns.last_mut() else {
                            continue;
                        };
                        match tag {
                            b"filters" => col.blank = attr_bool(xml, &e, b"blank")?,
                            b"filter" => {
                                if let Some(v) = attr(xml, &e, b"val")? {
                                    col.values.push(v);
                                }
                            }
                            b"dateGroupItem" => col.date_groups.push(all_attrs(xml, &e)?),
                            b"customFilters" => col.custom_and = attr_bool(xml, &e, b"and")?,
                            b"customFilter" => col.custom.push(all_attrs(xml, &e)?),
                            b"top10" => col.top10 = Some(all_attrs(xml, &e)?),
                            b"dynamicFilter" => col.dynamic = Some(all_attrs(xml, &e)?),
                            b"colorFilter" => col.color = Some(all_attrs(xml, &e)?),
                            b"iconFilter" => col.icon = Some(all_attrs(xml, &e)?),
                            _ => {}
                        }
                    }
                }
            }
            Event::End(e) => match e.local_name().as_ref() {
                b"sortState" => in_sort = false,
                b"autoFilter" => break,
                _ => {}
            },
            Event::Eof => return Err(XlsxError::XmlEof("autoFilter")),
            _ => {}
        }
    }
    Ok(af)
}

/// Reads a whole zip part into memory, `None` if it does not exist.
fn read_part<RS: Read + Seek>(
    zip: &mut zip::ZipArchive<RS>,
    cache: &HashMap<String, String>,
    path: &str,
) -> Result<Option<Vec<u8>>, XlsxError> {
    match zip.by_name(cached_zip_path(cache, path)) {
        Ok(mut f) => {
            let mut buf = Vec::with_capacity(f.size() as usize);
            f.read_to_end(&mut buf)?;
            Ok(Some(buf))
        }
        Err(ZipError::FileNotFound) => Ok(None),
        Err(e) => Err(e.into()),
    }
}

impl<RS: Read + Seek> Xlsx<RS> {
    /// Returns the cell formats (`cellXfs`), differential formats (`dxfs`) and named styles.
    pub fn stylesheet(&mut self) -> Result<StyleSheet, XlsxError> {
        let path = format!("{}styles.xml", self.xl_path);
        let mut xml = match xml_reader(&mut self.zip, &path, &self.zip_path_cache) {
            None => return Ok(StyleSheet::default()),
            Some(x) => x?,
        };

        let mut number_formats = HashMap::new();
        let mut fonts: Vec<StyleFont> = Vec::new();
        let mut fills: Vec<StyleFill> = Vec::new();
        let mut borders: Vec<StyleBorder> = Vec::new();
        let mut xfs: Vec<RawXf> = Vec::new();
        let mut dxfs: Vec<DifferentialStyle> = Vec::new();
        // (name, xfId into cellStyleXfs, builtinId, hidden)
        let mut cell_styles: Vec<(String, usize, Option<u32>, bool)> = Vec::new();

        let mut section = Section::None;
        let mut dxf_part = DxfPart::None;
        let mut side: Option<Vec<u8>> = None;
        let mut buf = Vec::with_capacity(1024);
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    let local = e.local_name();
                    let tag = local.as_ref();
                    match (section, tag) {
                        (Section::None, b"numFmt") => {
                            if let (Some(id), Some(code)) = (
                                attr_num::<RS, u32>(&xml, &e, b"numFmtId")?,
                                attr(&xml, &e, b"formatCode")?,
                            ) {
                                number_formats.insert(id, code);
                            }
                        }
                        (Section::None, b"fonts") => section = Section::Fonts,
                        (Section::None, b"fills") => section = Section::Fills,
                        (Section::None, b"borders") => section = Section::Borders,
                        (Section::None, b"cellXfs") => section = Section::CellXfs,
                        (Section::None, b"dxfs") => section = Section::Dxfs,
                        (Section::None, b"cellStyles") => section = Section::CellStyles,
                        // cellStyleXfs, ... contain the same element names; skip them.
                        (Section::None, b"cellStyleXfs" | b"colors" | b"tableStyles") => {
                            section = Section::Other
                        }

                        (Section::Fonts, b"font") => fonts.push(StyleFont::default()),
                        (Section::Fonts, _) => {
                            if let Some(font) = fonts.last_mut() {
                                apply_font(font, &xml, &e, tag)?;
                            }
                        }
                        (Section::Fills, b"fill") => fills.push(StyleFill::default()),
                        (Section::Fills, _) => {
                            if let Some(fill) = fills.last_mut() {
                                apply_fill(fill, &xml, &e, tag)?;
                                // Only dxf fills default to solid.
                                if tag == b"patternFill" && e.raw_attr(b"patternType")?.is_none() {
                                    fill.pattern = None;
                                }
                            }
                        }
                        (Section::Borders, b"border") => {
                            borders.push(new_border(&xml, &e)?);
                            side = None;
                        }
                        (Section::Borders, _) => {
                            if let Some(border) = borders.last_mut() {
                                apply_border(border, &mut side, &xml, &e, tag)?;
                            }
                        }

                        (Section::CellXfs, b"xf") => xfs.push(RawXf {
                            num_fmt_id: attr_num(&xml, &e, b"numFmtId")?.unwrap_or(0),
                            font_id: attr_num(&xml, &e, b"fontId")?.unwrap_or(0),
                            fill_id: attr_num(&xml, &e, b"fillId")?.unwrap_or(0),
                            border_id: attr_num(&xml, &e, b"borderId")?.unwrap_or(0),
                            xf_id: attr_num(&xml, &e, b"xfId")?,
                            quote_prefix: attr_bool(&xml, &e, b"quotePrefix")?,
                            pivot_button: attr_bool(&xml, &e, b"pivotButton")?,
                            locked: true,
                            ..Default::default()
                        }),
                        (Section::CellXfs, b"alignment") => {
                            if let Some(xf) = xfs.last_mut() {
                                xf.alignment = alignment(&xml, &e)?;
                            }
                        }
                        (Section::CellXfs, b"protection") => {
                            if let Some(xf) = xfs.last_mut() {
                                xf.locked = !matches!(
                                    attr(&xml, &e, b"locked")?.as_deref(),
                                    Some("0") | Some("false")
                                );
                                xf.hidden = attr_bool(&xml, &e, b"hidden")?;
                            }
                        }

                        (Section::CellStyles, b"cellStyle") => {
                            if let Some(name) = attr(&xml, &e, b"name")? {
                                cell_styles.push((
                                    name,
                                    attr_num(&xml, &e, b"xfId")?.unwrap_or(0),
                                    attr_num(&xml, &e, b"builtinId")?,
                                    attr_bool(&xml, &e, b"hidden")?,
                                ));
                            }
                        }

                        (Section::Dxfs, b"dxf") => {
                            dxfs.push(DifferentialStyle::default());
                            dxf_part = DxfPart::None;
                        }
                        (Section::Dxfs, _) => {
                            let Some(dxf) = dxfs.last_mut() else { continue };
                            match (dxf_part, tag) {
                                (_, b"font") => {
                                    dxf.font = Some(StyleFont::default());
                                    dxf_part = DxfPart::Font;
                                }
                                (_, b"fill") => {
                                    dxf.fill = Some(StyleFill::default());
                                    dxf_part = DxfPart::Fill;
                                }
                                (_, b"border") => {
                                    dxf.border = Some(new_border(&xml, &e)?);
                                    dxf_part = DxfPart::Border;
                                    side = None;
                                }
                                (_, b"numFmt") => {
                                    dxf.number_format = attr(&xml, &e, b"formatCode")?
                                }
                                (_, b"alignment") => dxf.alignment = Some(alignment(&xml, &e)?),
                                (DxfPart::Font, _) => {
                                    if let Some(f) = dxf.font.as_mut() {
                                        apply_font(f, &xml, &e, tag)?;
                                    }
                                }
                                (DxfPart::Fill, _) => {
                                    if let Some(f) = dxf.fill.as_mut() {
                                        apply_fill(f, &xml, &e, tag)?;
                                    }
                                }
                                (DxfPart::Border, _) => {
                                    if let Some(b) = dxf.border.as_mut() {
                                        apply_border(b, &mut side, &xml, &e, tag)?;
                                    }
                                }
                                _ => {}
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(e)) => match e.local_name().as_ref() {
                    b"fonts" | b"fills" | b"borders" | b"cellXfs" | b"cellStyleXfs" | b"dxfs"
                    | b"cellStyles" | b"colors" | b"tableStyles" => section = Section::None,
                    b"font" | b"fill" | b"border" => dxf_part = DxfPart::None,
                    b"left" | b"start" | b"right" | b"end" | b"top" | b"bottom" | b"diagonal"
                    | b"vertical" | b"horizontal" => side = None,
                    b"styleSheet" => break,
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => {}
            }
        }

        let style_name = |xf_id: usize| {
            cell_styles
                .iter()
                .find(|(_, id, _, _)| *id == xf_id)
                .map(|(n, _, _, _)| n.clone())
        };
        let cell_formats = xfs
            .into_iter()
            .map(|xf| CellStyle {
                number_format_id: xf.num_fmt_id,
                number_format: number_formats
                    .get(&xf.num_fmt_id)
                    .cloned()
                    .unwrap_or_else(|| builtin_number_format(xf.num_fmt_id).to_string()),
                font: fonts.get(xf.font_id).cloned().unwrap_or_default(),
                fill: fills.get(xf.fill_id).cloned().unwrap_or_default(),
                border: borders.get(xf.border_id).cloned().unwrap_or_default(),
                alignment: xf.alignment,
                locked: xf.locked,
                hidden: xf.hidden,
                named_style: style_name(xf.xf_id.unwrap_or(0)),
                quote_prefix: xf.quote_prefix,
                pivot_button: xf.pivot_button,
                number_format_custom: number_formats.contains_key(&xf.num_fmt_id),
            })
            .collect();
        let zero_style = Some(CellStyle {
            number_format_id: 0,
            number_format: "General".to_string(),
            font: fonts.first().cloned().unwrap_or_default(),
            fill: fills.first().cloned().unwrap_or_default(),
            border: borders.first().cloned().unwrap_or_default(),
            alignment: StyleAlignment::default(),
            locked: true,
            hidden: false,
            named_style: style_name(0),
            quote_prefix: false,
            pivot_button: false,
            number_format_custom: false,
        });
        Ok(StyleSheet {
            zero_style,
            cell_styles: cell_formats,
            differential_styles: dxfs,
            named_styles: cell_styles
                .into_iter()
                .map(|(name, _, builtin_id, hidden)| NamedStyle {
                    name,
                    builtin_id,
                    hidden,
                })
                .collect(),
        })
    }

    /// Returns every cell format of the workbook (`cellXfs`), indexed by a cell's style id.
    pub fn cell_styles(&mut self) -> Result<Vec<CellStyle>, XlsxError> {
        Ok(self.stylesheet()?.cell_styles)
    }

    fn sheet_path(&self, name: &str) -> Result<String, XlsxError> {
        self.sheets
            .iter()
            .find(|(n, _)| n == name)
            .map(|(_, p)| p.clone())
            .ok_or_else(|| XlsxError::WorksheetNotFound(name.into()))
    }

    /// Relationships of a part as `(id, type, resolved target, external)`.
    fn relationships(
        &mut self,
        part: &str,
    ) -> Result<Vec<(String, String, String, bool)>, XlsxError> {
        let (base_folder, file_name) = match part.rfind('/') {
            Some(i) => part.split_at(i),
            None => ("", part),
        };
        let file_name = file_name.trim_start_matches('/');
        let rel_path = if base_folder.is_empty() {
            format!("_rels/{file_name}.rels")
        } else {
            format!("{base_folder}/_rels/{file_name}.rels")
        };
        let mut xml = match xml_reader(&mut self.zip, &rel_path, &self.zip_path_cache) {
            None => return Ok(Vec::new()),
            Some(x) => x?,
        };
        let mut out = Vec::new();
        let mut buf = Vec::with_capacity(256);
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"Relationship" => {
                    let id = attr(&xml, &e, b"Id")?.unwrap_or_default();
                    let typ = attr(&xml, &e, b"Type")?.unwrap_or_default();
                    let target = attr(&xml, &e, b"Target")?.unwrap_or_default();
                    let external = attr(&xml, &e, b"TargetMode")?.as_deref() == Some("External");
                    let target = if external {
                        target
                    } else {
                        resolve_target(base_folder, &target)
                    };
                    out.push((id, typ, target, external));
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => {}
            }
        }
        Ok(out)
    }

    /// Returns the comments (notes) of a worksheet, in file order.
    pub fn worksheet_comments(&mut self, name: &str) -> Result<Vec<Comment>, XlsxError> {
        let sheet_path = self.sheet_path(name)?;
        let comments_path = self
            .relationships(&sheet_path)?
            .into_iter()
            .find(|(_, typ, _, ext)| typ.ends_with("/relationships/comments") && !ext)
            .map(|(_, _, target, _)| target);
        let Some(comments_path) = comments_path else {
            return Ok(Vec::new());
        };
        let mut xml = match xml_reader(&mut self.zip, &comments_path, &self.zip_path_cache) {
            None => return Ok(Vec::new()),
            Some(x) => x?,
        };

        let mut authors: Vec<String> = Vec::new();
        let mut comments = Vec::new();
        let mut current: Option<((u32, u32), Option<usize>)> = None;
        let (mut buf, mut xml_buf, mut text_buf) = (Vec::new(), Vec::new(), Vec::new());
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.local_name().as_ref() {
                    b"author" => authors.push(read_plain_text(&mut xml, b"author", &mut xml_buf)?),
                    b"comment" => {
                        let pos = match e.raw_attr(b"ref")? {
                            Some(r) => get_row_column(r)?,
                            None => continue,
                        };
                        current = Some((pos, attr_num(&xml, &e, b"authorId")?));
                    }
                    b"text" => {
                        let qname = e.name().as_ref().to_vec();
                        let text = read_string_with_bufs(
                            &mut xml,
                            QName(&qname),
                            &mut xml_buf,
                            &mut text_buf,
                        )?
                        .unwrap_or_default();
                        if let Some((pos, author_id)) = current.take() {
                            comments.push(Comment {
                                pos,
                                author: author_id.and_then(|i| authors.get(i).cloned()),
                                text,
                            });
                        }
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => {}
            }
        }
        Ok(comments)
    }

    /// Defined names with their sheet scope, hidden flag and comment.
    pub fn defined_names_scoped(&mut self) -> Result<Vec<DefinedName>, XlsxError> {
        let path = format!("{}workbook.xml", self.xl_path);
        let mut xml = match xml_reader(&mut self.zip, &path, &self.zip_path_cache) {
            None => return Ok(Vec::new()),
            Some(x) => x?,
        };
        let sheet_names: Vec<String> = self
            .metadata
            .sheets
            .iter()
            .map(|s| s.name.clone())
            .collect();
        let mut names = Vec::new();
        let (mut buf, mut text_buf) = (Vec::new(), Vec::new());
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"definedName" => {
                    let Some(name) = attr(&xml, &e, b"name")? else {
                        continue;
                    };
                    let sheet = attr_num::<RS, usize>(&xml, &e, b"localSheetId")?
                        .and_then(|i| sheet_names.get(i).cloned());
                    let hidden = attr_bool(&xml, &e, b"hidden")?;
                    let comment = attr(&xml, &e, b"comment")?;
                    let value = element_text(&mut xml, &e, &mut text_buf)?;
                    names.push(DefinedName {
                        name,
                        value,
                        sheet,
                        hidden,
                        comment,
                    });
                }
                Ok(Event::End(e)) if e.local_name().as_ref() == b"definedNames" => break,
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => {}
            }
        }
        Ok(names)
    }

    /// Document properties from `docProps/core.xml` and `docProps/app.xml`, as
    /// `(key, value)` pairs with snake_case keys (`creator`, `created`, `company`, ...).
    /// Dates are returned as stored (W3CDTF, e.g. `2024-01-31T10:00:00Z`).
    pub fn document_properties(&mut self) -> Result<Vec<(String, String)>, XlsxError> {
        const KEYS: &[(&[u8], &str)] = &[
            (b"title", "title"),
            (b"subject", "subject"),
            (b"creator", "creator"),
            (b"keywords", "keywords"),
            (b"description", "description"),
            (b"lastModifiedBy", "last_modified_by"),
            (b"lastPrinted", "last_printed"),
            (b"revision", "revision"),
            (b"created", "created"),
            (b"modified", "modified"),
            (b"category", "category"),
            (b"contentStatus", "content_status"),
            (b"identifier", "identifier"),
            (b"language", "language"),
            (b"version", "version"),
            (b"Application", "application"),
            (b"AppVersion", "app_version"),
            (b"Company", "company"),
            (b"Manager", "manager"),
            (b"HyperlinkBase", "hyperlink_base"),
        ];
        let mut props = Vec::new();
        for path in ["docProps/core.xml", "docProps/app.xml"] {
            let mut xml = match xml_reader(&mut self.zip, path, &self.zip_path_cache) {
                None => continue,
                Some(x) => x?,
            };
            let (mut buf, mut text_buf) = (Vec::new(), Vec::new());
            loop {
                buf.clear();
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(e)) => {
                        let local = e.local_name();
                        if let Some((_, key)) = KEYS.iter().find(|(k, _)| *k == local.as_ref()) {
                            let value = element_text(&mut xml, &e, &mut text_buf)?;
                            if !value.is_empty() {
                                props.push((key.to_string(), value));
                            }
                        }
                    }
                    Ok(Event::Eof) => break,
                    Err(e) => return Err(XlsxError::Xml(e)),
                    _ => {}
                }
            }
        }
        Ok(props)
    }

    /// Workbook settings, custom properties, theme and file kind.
    pub fn workbook_info(&mut self) -> Result<WorkbookInfo, XlsxError> {
        let mut info = WorkbookInfo::default();
        let (mut buf, mut text_buf) = (Vec::new(), Vec::new());

        let wb_path = format!("{}workbook.xml", self.xl_path);
        if let Some(xml) = xml_reader(&mut self.zip, &wb_path, &self.zip_path_cache) {
            let mut xml = xml?;
            loop {
                buf.clear();
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(e)) => match e.local_name().as_ref() {
                        b"workbookPr" => info.workbook_properties = all_attrs(&xml, &e)?,
                        b"workbookView" => info.views.push(all_attrs(&xml, &e)?),
                        b"calcPr" => info.calculation = Some(all_attrs(&xml, &e)?),
                        b"workbookProtection" => info.protection = Some(all_attrs(&xml, &e)?),
                        b"fileVersion" => info.file_version = Some(all_attrs(&xml, &e)?),
                        b"definedNames" | b"sheets" => {
                            text_buf.clear();
                            xml.read_to_end_into(e.name(), &mut text_buf)?;
                        }
                        _ => {}
                    },
                    Ok(Event::Eof) => break,
                    Err(e) => return Err(XlsxError::Xml(e)),
                    _ => {}
                }
            }
        }

        // Custom properties: <property name=".."><vt:lpwstr>value</vt:lpwstr></property>
        if let Some(xml) = xml_reader(&mut self.zip, "docProps/custom.xml", &self.zip_path_cache) {
            let mut xml = xml?;
            let mut name: Option<String> = None;
            loop {
                buf.clear();
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(e)) if e.local_name().as_ref() == b"property" => {
                        name = attr(&xml, &e, b"name")?;
                    }
                    Ok(Event::Start(e)) => {
                        if let Some(n) = name.take() {
                            let kind =
                                String::from_utf8_lossy(e.local_name().as_ref()).into_owned();
                            let value = element_text(&mut xml, &e, &mut text_buf)?;
                            info.custom_properties.push(CustomProperty {
                                name: n,
                                kind,
                                value,
                            });
                        }
                    }
                    Ok(Event::Eof) => break,
                    Err(e) => return Err(XlsxError::Xml(e)),
                    _ => {}
                }
            }
        }

        // File kind from the content type of the workbook part.
        if let Some(xml) = xml_reader(&mut self.zip, "[Content_Types].xml", &self.zip_path_cache) {
            let mut xml = xml?;
            let wb_part = format!("/{wb_path}");
            loop {
                buf.clear();
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(e)) if e.local_name().as_ref() == b"Override" => {
                        if attr(&xml, &e, b"PartName")?.as_deref() == Some(wb_part.as_str()) {
                            let ct = attr(&xml, &e, b"ContentType")?.unwrap_or_default();
                            info.is_template = ct.contains(".template");
                            info.has_macros = ct.contains("macroEnabled");
                        }
                    }
                    Ok(Event::Eof) => break,
                    Err(e) => return Err(XlsxError::Xml(e)),
                    _ => {}
                }
            }
        }

        // Theme.
        let theme_path = self
            .relationships(&wb_path)?
            .into_iter()
            .find(|(_, typ, _, ext)| typ.ends_with("/relationships/theme") && !ext)
            .map(|(_, _, target, _)| target);
        if let Some(path) = theme_path {
            if let Some(bytes) = read_part(&mut self.zip, &self.zip_path_cache, &path)? {
                info.theme_colors = theme_colors(&bytes);
                info.theme_xml = Some(bytes);
            }
        }
        Ok(info)
    }

    /// Rich text runs of the shared strings table, keyed by shared string index.
    /// Only strings with formatting runs are included.
    pub fn rich_shared_strings(&mut self) -> Result<HashMap<usize, Vec<TextRun>>, XlsxError> {
        let path = format!("{}sharedStrings.xml", self.xl_path);
        let mut xml = match xml_reader(&mut self.zip, &path, &self.zip_path_cache) {
            None => return Ok(HashMap::new()),
            Some(x) => x?,
        };
        let mut out = HashMap::new();
        let mut index: Option<usize> = None;
        let mut runs: Vec<TextRun> = Vec::new();
        let (mut in_run, mut in_rpr) = (false, false);
        let (mut buf, mut text_buf, mut skip_buf) = (Vec::new(), Vec::new(), Vec::new());
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"si" => {
                            index = Some(index.map_or(0, |i| i + 1));
                            runs.clear();
                        }
                        b"r" => {
                            runs.push(TextRun::default());
                            in_run = true;
                        }
                        b"rPr" if in_run => {
                            if let Some(run) = runs.last_mut() {
                                run.font = Some(StyleFont::default());
                            }
                            in_rpr = true;
                        }
                        b"t" if in_run => {
                            let text = read_plain_text(&mut xml, b"t", &mut text_buf)?;
                            if let Some(run) = runs.last_mut() {
                                run.text.push_str(&text);
                            }
                        }
                        // Phonetic guides are not part of the text.
                        b"rPh" | b"phoneticPr" => {
                            skip_buf.clear();
                            xml.read_to_end_into(e.name(), &mut skip_buf)?;
                        }
                        tag if in_rpr => {
                            if let Some(font) = runs.last_mut().and_then(|r| r.font.as_mut()) {
                                apply_font(font, &xml, &e, tag)?;
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(e)) => match e.local_name().as_ref() {
                    b"rPr" => in_rpr = false,
                    b"r" => in_run = false,
                    b"si" => {
                        if let (Some(i), false) = (index, runs.is_empty()) {
                            out.insert(i, std::mem::take(&mut runs));
                        }
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => {}
            }
        }
        Ok(out)
    }

    /// Everything a worksheet stores outside its cell values, in one pass over the
    /// sheet XML (cells are skipped; only `<row>` attributes are read), plus its tables.
    pub fn worksheet_info(&mut self, name: &str) -> Result<WorksheetInfo, XlsxError> {
        self.worksheet_info_impl(name, false)
    }

    /// Fork addition (openpyxl compatibility): `worksheet_info`, plus the bounding
    /// box of every stored cell (`WorksheetInfo::cell_bounds`), from the same pass.
    pub fn worksheet_info_with_bounds(&mut self, name: &str) -> Result<WorksheetInfo, XlsxError> {
        self.worksheet_info_impl(name, true)
    }

    fn worksheet_info_impl(
        &mut self,
        name: &str,
        track_cells: bool,
    ) -> Result<WorksheetInfo, XlsxError> {
        let path = self.sheet_path(name)?;
        let rels = read_sheet_hyperlink_rels(&mut self.zip, &path, &self.zip_path_cache)?;
        let mut info = WorksheetInfo::default();
        let mut table_ids: Vec<String> = Vec::new();
        {
            let mut xml = xml_reader(&mut self.zip, &path, &self.zip_path_cache)
                .ok_or_else(|| XlsxError::WorksheetNotFound(name.into()))??;
            read_worksheet_xml(&mut xml, &rels, &mut info, &mut table_ids, track_cells)?;
        }

        // Tables live in their own parts.
        if !table_ids.is_empty() {
            let table_paths: Vec<String> = self
                .relationships(&path)?
                .into_iter()
                .filter(|(id, typ, _, ext)| {
                    typ.ends_with("/relationships/table") && !ext && table_ids.contains(id)
                })
                .map(|(_, _, target, _)| target)
                .collect();
            for table_path in table_paths {
                if let Some(xml) = xml_reader(&mut self.zip, &table_path, &self.zip_path_cache) {
                    info.tables.push(read_table(&mut xml?)?);
                }
            }
        }
        Ok(info)
    }
}

/// Parses a table part (`xl/tables/tableN.xml`).
fn read_table<RS: Read + Seek>(xml: &mut XlReader<'_, RS>) -> Result<TableInfo, XlsxError> {
    let mut table = TableInfo::default();
    let (mut buf, mut text_buf) = (Vec::new(), Vec::new());
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"table" => table.attrs = all_attrs(xml, &e)?,
                b"autoFilter" => table.auto_filter = Some(read_auto_filter(xml, &e)?),
                b"tableColumn" => table.columns.push(TableColumn {
                    attrs: all_attrs(xml, &e)?,
                    ..Default::default()
                }),
                b"calculatedColumnFormula" => {
                    let f = element_text(xml, &e, &mut text_buf)?;
                    if let Some(c) = table.columns.last_mut() {
                        c.calculated_formula = Some(f);
                    }
                }
                b"totalsRowFormula" => {
                    let f = element_text(xml, &e, &mut text_buf)?;
                    if let Some(c) = table.columns.last_mut() {
                        c.totals_formula = Some(f);
                    }
                }
                b"tableStyleInfo" => table.style = Some(all_attrs(xml, &e)?),
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(table)
}

/// Theme color scheme of a theme part as `(slot, rgb)`.
fn theme_colors(bytes: &[u8]) -> Vec<(String, String)> {
    let mut reader = quick_xml::Reader::from_reader(bytes);
    let mut out = Vec::new();
    let mut slot: Option<String> = None;
    let mut in_scheme = false;
    let mut buf = Vec::new();
    loop {
        buf.clear();
        let event = match reader.read_event_into(&mut buf) {
            Ok(e) => e,
            Err(_) => break,
        };
        match event {
            Event::Start(e) | Event::Empty(e) => {
                let local = e.local_name();
                let tag = local.as_ref();
                if tag == b"clrScheme" {
                    in_scheme = true;
                } else if in_scheme && slot.is_none() {
                    slot = Some(String::from_utf8_lossy(tag).into_owned());
                } else if let Some(s) = slot.as_ref() {
                    let key: &[u8] = if tag == b"sysClr" { b"lastClr" } else { b"val" };
                    if tag == b"srgbClr" || tag == b"sysClr" {
                        if let Ok(Some(v)) = e.raw_attr(key) {
                            out.push((s.clone(), String::from_utf8_lossy(v).into_owned()));
                        }
                    }
                }
            }
            Event::End(e) => {
                let local = e.local_name();
                if local.as_ref() == b"clrScheme" {
                    break;
                }
                if slot.as_deref().map(str::as_bytes) == Some(local.as_ref()) {
                    slot = None;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    out
}

/// The single pass over a worksheet's XML behind `Xlsx::worksheet_info`.
fn read_worksheet_xml<RS: Read + Seek>(
    xml: &mut XlReader<'_, RS>,
    rels: &HashMap<String, String>,
    info: &mut WorksheetInfo,
    table_ids: &mut Vec<String>,
    track_cells: bool,
) -> Result<(), XlsxError> {
    // Open elements below <worksheet> (elements read by helpers are never pushed).
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut cf: Option<ConditionalFormat> = None;
    let mut dv: Option<DataValidation> = None;
    let (mut buf, mut text_buf, mut row_buf) = (Vec::new(), Vec::new(), Vec::new());
    loop {
        buf.clear();
        let event = xml.read_event_into(&mut buf)?;
        match event {
            Event::Start(e) => {
                let local = e.local_name();
                let tag = local.as_ref();
                let parent = stack.last().map(Vec::as_slice);
                let in_ext = stack.iter().any(|t| t == b"extLst");
                let top_level = stack.len() == 1; // directly below <worksheet>
                match tag {
                    b"sheetData" => {
                        let bounds = track_cells.then_some(&mut info.cell_bounds);
                        read_row_attributes(xml, &mut row_buf, &mut info.rows, bounds)?;
                        continue;
                    }
                    b"mergeCells" if top_level => {
                        info.merged = read_merge_cells(xml)?;
                        continue;
                    }
                    b"hyperlinks" if top_level => {
                        info.hyperlinks = read_hyperlinks(xml, rels)?;
                        continue;
                    }
                    b"autoFilter" if top_level => {
                        let af = read_auto_filter(xml, &e)?;
                        info.auto_filter = af.range.clone();
                        info.auto_filter_detail = Some(af);
                        continue;
                    }
                    b"oddHeader" | b"oddFooter" | b"evenHeader" | b"evenFooter"
                    | b"firstHeader" | b"firstFooter"
                        if parent == Some(b"headerFooter") =>
                    {
                        let key = String::from_utf8_lossy(tag).into_owned();
                        let text = element_text(xml, &e, &mut text_buf)?;
                        if let Some(hf) = info.header_footer.as_mut() {
                            hf.parts.push((key, text));
                        }
                        continue;
                    }
                    b"formula" if cf.is_some() => {
                        let f = element_text(xml, &e, &mut text_buf)?;
                        if let Some(rule) = cf.as_mut().and_then(|cf| cf.rules.last_mut()) {
                            rule.formulas.push(f);
                        }
                        continue;
                    }
                    b"formula1" | b"formula2" | b"sqref" if dv.is_some() => {
                        let text = element_text(xml, &e, &mut text_buf)?;
                        let v = dv.as_mut().expect("checked");
                        match tag {
                            b"formula1" => v.formula1 = Some(text),
                            b"formula2" => v.formula2 = Some(text),
                            _ => v.ranges = split_ranges(&text),
                        }
                        continue;
                    }
                    _ => {}
                }

                match tag {
                    b"dimension" if top_level => info.dimension = attr(xml, &e, b"ref")?,
                    b"sheetPr" if top_level => info.sheet_properties = all_attrs(xml, &e)?,
                    b"tabColor" if parent == Some(b"sheetPr") => {
                        info.tab_color = Some(color(xml, &e)?)
                    }
                    b"outlinePr" if parent == Some(b"sheetPr") => {
                        info.outline_properties = all_attrs(xml, &e)?
                    }
                    b"pageSetUpPr" if parent == Some(b"sheetPr") => {
                        info.page_setup_properties = all_attrs(xml, &e)?
                    }
                    b"sheetFormatPr" if top_level => {
                        info.default_row_height = attr_num(xml, &e, b"defaultRowHeight")?;
                        info.default_column_width = attr_num(xml, &e, b"defaultColWidth")?;
                        info.sheet_format = all_attrs(xml, &e)?;
                    }
                    b"sheetView" if parent == Some(b"sheetViews") => {
                        let attrs = all_attrs(xml, &e)?;
                        if info.views.is_empty() {
                            info.sheet_view = attrs.clone();
                        }
                        info.views.push(SheetView {
                            attrs,
                            ..Default::default()
                        });
                    }
                    b"pane" if parent == Some(b"sheetView") => {
                        let state = attr(xml, &e, b"state")?;
                        if info.freeze_pane.is_none()
                            && info.views.len() == 1
                            && matches!(state.as_deref(), Some("frozen") | Some("frozenSplit"))
                        {
                            info.freeze_pane = Some(FreezePane {
                                top_left_cell: attr(xml, &e, b"topLeftCell")?,
                                cols: attr_num(xml, &e, b"xSplit")?.unwrap_or(0.0),
                                rows: attr_num(xml, &e, b"ySplit")?.unwrap_or(0.0),
                            });
                        }
                        let attrs = all_attrs(xml, &e)?;
                        if let Some(v) = info.views.last_mut() {
                            v.pane = Some(attrs);
                        }
                    }
                    b"selection" if parent == Some(b"sheetView") => {
                        let attrs = all_attrs(xml, &e)?;
                        if let Some(v) = info.views.last_mut() {
                            v.selections.push(attrs);
                        }
                    }
                    b"customSheetView" if parent == Some(b"customSheetViews") => {
                        info.custom_views.push(all_attrs(xml, &e)?)
                    }
                    b"col" if parent == Some(b"cols") => {
                        let min: u32 = attr_num(xml, &e, b"min")?.unwrap_or(1);
                        let max: u32 = attr_num(xml, &e, b"max")?.unwrap_or(min);
                        info.columns.push(ColumnInfo {
                            min: min.saturating_sub(1),
                            max: max.saturating_sub(1),
                            width: attr_num(xml, &e, b"width")?,
                            custom_width: attr_bool(xml, &e, b"customWidth")?,
                            hidden: attr_bool(xml, &e, b"hidden")?,
                            outline_level: attr_num(xml, &e, b"outlineLevel")?.unwrap_or(0),
                            collapsed: attr_bool(xml, &e, b"collapsed")?,
                            style: attr_num(xml, &e, b"style")?,
                            best_fit: attr_bool(xml, &e, b"bestFit")?,
                        });
                    }
                    b"sortState" if top_level => {
                        info.sort_state = Some(SortState {
                            attrs: all_attrs(xml, &e)?,
                            conditions: Vec::new(),
                        })
                    }
                    b"sortCondition" if parent == Some(b"sortState") && stack.len() == 2 => {
                        let c = all_attrs(xml, &e)?;
                        if let Some(s) = info.sort_state.as_mut() {
                            s.conditions.push(c);
                        }
                    }
                    b"sheetProtection" if top_level => info.protection = Some(all_attrs(xml, &e)?),
                    b"printOptions" if top_level => info.print_options = all_attrs(xml, &e)?,
                    b"pageMargins" if top_level => info.page_margins = all_attrs(xml, &e)?,
                    b"pageSetup" if top_level => info.page_setup = all_attrs(xml, &e)?,
                    b"headerFooter" if top_level => {
                        info.header_footer = Some(HeaderFooter {
                            attrs: all_attrs(xml, &e)?,
                            parts: Vec::new(),
                        })
                    }
                    b"brk" if stack.len() == 2 && parent == Some(b"rowBreaks") => {
                        info.row_breaks.push(all_attrs(xml, &e)?)
                    }
                    b"brk" if stack.len() == 2 && parent == Some(b"colBreaks") => {
                        info.column_breaks.push(all_attrs(xml, &e)?)
                    }
                    b"scenarios" if top_level => {
                        info.scenarios = Some(Scenarios {
                            attrs: all_attrs(xml, &e)?,
                            scenarios: Vec::new(),
                        })
                    }
                    b"scenario" if parent == Some(b"scenarios") => {
                        let attrs = all_attrs(xml, &e)?;
                        if let Some(s) = info.scenarios.as_mut() {
                            s.scenarios.push(Scenario {
                                attrs,
                                input_cells: Vec::new(),
                            });
                        }
                    }
                    b"inputCells" if parent == Some(b"scenario") => {
                        let attrs = all_attrs(xml, &e)?;
                        if let Some(s) =
                            info.scenarios.as_mut().and_then(|s| s.scenarios.last_mut())
                        {
                            s.input_cells.push(attrs);
                        }
                    }
                    b"tablePart" if parent == Some(b"tableParts") => {
                        for (k, v) in all_attrs(xml, &e)? {
                            if k == "id" {
                                table_ids.push(v);
                            }
                        }
                    }

                    // Extended (x14) conditional formats duplicate/extend the
                    // regular ones (data bar details); only the regular ones are read.
                    b"conditionalFormatting" if top_level && !in_ext => {
                        cf = Some(ConditionalFormat {
                            ranges: split_ranges(&attr(xml, &e, b"sqref")?.unwrap_or_default()),
                            pivot: attr_bool(xml, &e, b"pivot")?,
                            rules: Vec::new(),
                        });
                    }
                    b"cfRule" if cf.is_some() => {
                        let mut rule = ConditionalRule::default();
                        for (k, v) in all_attrs(xml, &e)? {
                            match k.as_str() {
                                "type" => rule.typ = Some(v),
                                "priority" => rule.priority = v.parse().ok(),
                                "operator" => rule.operator = Some(v),
                                "dxfId" => rule.dxf_id = v.parse().ok(),
                                "stopIfTrue" => rule.stop_if_true = v == "1" || v == "true",
                                "text" => rule.text = Some(v),
                                _ => rule.extra.push((k, v)),
                            }
                        }
                        if let Some(cf) = cf.as_mut() {
                            cf.rules.push(rule);
                        }
                    }
                    b"cfvo" if cf.is_some() => {
                        let v = ConditionalValue {
                            typ: attr(xml, &e, b"type")?.unwrap_or_default(),
                            value: attr(xml, &e, b"val")?,
                        };
                        if let Some(rule) = cf.as_mut().and_then(|cf| cf.rules.last_mut()) {
                            rule.values.push(v);
                        }
                    }
                    b"color" if cf.is_some() => {
                        let c = color(xml, &e)?;
                        if let Some(rule) = cf.as_mut().and_then(|cf| cf.rules.last_mut()) {
                            rule.colors.push(c);
                        }
                    }
                    b"iconSet" | b"dataBar" | b"colorScale" if cf.is_some() => {
                        let attrs = all_attrs(xml, &e)?;
                        if let Some(rule) = cf.as_mut().and_then(|cf| cf.rules.last_mut()) {
                            for (k, v) in attrs {
                                if k == "iconSet" {
                                    rule.icon_set = Some(v);
                                } else {
                                    rule.extra.push((k, v));
                                }
                            }
                            if tag == b"iconSet" && rule.icon_set.is_none() {
                                rule.icon_set = Some("3TrafficLights1".to_string());
                            }
                        }
                    }

                    b"dataValidation" => {
                        let mut v = DataValidation {
                            in_cell_dropdown: true,
                            ..Default::default()
                        };
                        for (k, val) in all_attrs(xml, &e)? {
                            let truthy = val == "1" || val == "true";
                            match k.as_str() {
                                "sqref" => v.ranges = split_ranges(&val),
                                "type" => v.typ = Some(val),
                                "operator" => v.operator = Some(val),
                                "allowBlank" => v.allow_blank = truthy,
                                "showDropDown" => v.in_cell_dropdown = !truthy,
                                "showInputMessage" => v.show_input_message = truthy,
                                "showErrorMessage" => v.show_error_message = truthy,
                                "errorStyle" => v.error_style = Some(val),
                                "errorTitle" => v.error_title = Some(val),
                                "error" => v.error = Some(val),
                                "promptTitle" => v.prompt_title = Some(val),
                                "prompt" => v.prompt = Some(val),
                                "imeMode" => v.ime_mode = Some(val),
                                _ => {}
                            }
                        }
                        if v.typ.as_deref() == Some("none") {
                            v.typ = None;
                        }
                        dv = Some(v);
                    }
                    _ => {}
                }
                stack.push(tag.to_vec());
            }
            Event::End(e) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"conditionalFormatting" => {
                        if let Some(cf) = cf.take() {
                            info.conditional_formats.push(cf);
                        }
                    }
                    b"dataValidation" => {
                        if let Some(v) = dv.take() {
                            info.data_validations.push(v);
                        }
                    }
                    b"worksheet" => break,
                    _ => {}
                }
                stack.pop();
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(())
}

/// Skips the cells of `<sheetData>`, collecting non-default `<row>` attributes.
fn read_row_attributes<RS: Read + Seek>(
    xml: &mut XlReader<'_, RS>,
    buf: &mut Vec<u8>,
    rows: &mut Vec<RowAttributes>,
    mut bounds: Option<&mut Option<Dimensions>>,
) -> Result<(), XlsxError> {
    let mut row_index = 0u32;
    let mut col_index = 0u32;
    loop {
        buf.clear();
        match xml.read_event_into(buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"row" => {
                let (index, attrs) = parse_row(&e, row_index)?;
                row_index = index;
                col_index = 0;
                if let Some(a) = attrs {
                    rows.push(a);
                }
            }
            Event::Start(e) if bounds.is_some() && e.local_name().as_ref() == b"c" => {
                let pos = match e.raw_attr(b"r")? {
                    Some(r) => get_row_column(r)?,
                    None => (row_index, col_index),
                };
                col_index = pos.1 + 1;
                if let Some(b) = bounds.as_deref_mut() {
                    *b = Some(match *b {
                        None => Dimensions::new(pos, pos),
                        Some(d) => Dimensions::new(
                            (d.start.0.min(pos.0), d.start.1.min(pos.1)),
                            (d.end.0.max(pos.0), d.end.1.max(pos.1)),
                        ),
                    });
                }
            }
            Event::End(e) => match e.local_name().as_ref() {
                b"row" => row_index += 1,
                b"sheetData" => return Ok(()),
                _ => {}
            },
            Event::Eof => return Err(XlsxError::XmlEof("sheetData")),
            _ => {}
        }
    }
}

/// Reads an inline string (`<is>`) up to `closing`, returning its text (with the
/// same whitespace rules as calamine's string reader) and its formatting runs
/// (empty when the string has no runs).
pub(crate) fn read_inline_rich<RS: Read + Seek>(
    xml: &mut XlReader<'_, RS>,
    closing: &[u8],
    buf: &mut Vec<u8>,
    text_buf: &mut Vec<u8>,
) -> Result<(String, Vec<TextRun>), XlsxError> {
    let mut text = String::new();
    let mut runs: Vec<TextRun> = Vec::new();
    let (mut in_run, mut in_rpr) = (false, false);
    loop {
        buf.clear();
        match xml.read_event_into(buf)? {
            Event::Start(e) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"r" => {
                        runs.push(TextRun::default());
                        in_run = true;
                    }
                    b"rPr" if in_run => {
                        if let Some(run) = runs.last_mut() {
                            run.font = Some(StyleFont::default());
                        }
                        in_rpr = true;
                    }
                    b"t" => {
                        let preserve =
                            matches!(e.raw_attr(b"xml:space")?, Some(v) if v == b"preserve");
                        let raw = read_plain_text(xml, b"t", text_buf)?;
                        let value = if preserve {
                            raw
                        } else {
                            raw.trim_matches([' ', '\t', '\r', '\n']).to_string()
                        };
                        text.push_str(&value);
                        if in_run {
                            if let Some(run) = runs.last_mut() {
                                run.text.push_str(&value);
                            }
                        }
                    }
                    b"rPh" | b"phoneticPr" => {
                        text_buf.clear();
                        xml.read_to_end_into(e.name(), text_buf)?;
                    }
                    tag if in_rpr => {
                        if let Some(font) = runs.last_mut().and_then(|r| r.font.as_mut()) {
                            apply_font(font, xml, &e, tag)?;
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => match e.local_name().as_ref() {
                b"rPr" => in_rpr = false,
                b"r" => in_run = false,
                tag if tag == closing => break,
                _ => {}
            },
            Event::Eof => return Err(XlsxError::XmlEof("is")),
            _ => {}
        }
    }
    Ok((text, runs))
}
