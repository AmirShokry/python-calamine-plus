// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

use std::io::{Read, Seek};

use quick_xml::events::{BytesStart, Event};

use super::rich::{read_inline_rich, TextRun};
use super::{
    expand_shared_formula, get_dimension, get_row, get_row_column, read_string_with_bufs,
    Dimensions, XlReader,
};
use crate::attrs::RawAttributes;
use crate::datatype::DataRef;
use crate::formats::{format_excel_f64_ref, CellFormat};
use crate::utils::unescape_entity_to_buffer;
use crate::{Cell, XlsxError};

#[derive(Clone, Debug)]
struct SharedFormula {
    formula: String,
    range: Dimensions,
}

/// Workbook-level context used when reading cell values.
struct WorkbookContext<'a> {
    strings: &'a [String],
    formats: &'a [CellFormat],
    is_1904: bool,
}

/// Reusable scratch buffers for cell value parsing (avoid per-cell allocations).
struct ValueBufs {
    xml: Vec<u8>,
    value: String,
    str_inner: Vec<u8>,
    // Fork addition: shared string index of the last `t="s"` value read, and the
    // formatting runs of the last inline string (when `capture_runs` is set).
    last_sst: Option<usize>,
    capture_runs: bool,
    last_runs: Option<Vec<TextRun>>,
}

impl ValueBufs {
    fn new() -> Self {
        Self {
            xml: Vec::with_capacity(1024),
            value: String::with_capacity(64),
            str_inner: Vec::with_capacity(1024),
            last_sst: None,
            capture_runs: false,
            last_runs: None,
        }
    }
}

/// Fork addition: non-default attributes of a `<row>` element.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct RowAttributes {
    /// Row index (0-based).
    pub index: u32,
    /// Height in points.
    pub height: Option<f64>,
    /// Height was set explicitly.
    pub custom_height: bool,
    /// Hidden.
    pub hidden: bool,
    /// Outline (grouping) level.
    pub outline_level: u8,
    /// Outline group collapsed.
    pub collapsed: bool,
    /// Row default style id (only when the row has a custom format).
    pub style: Option<u32>,
}

/// Formula metadata attached to an XLSX cell record.
#[non_exhaustive]
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum XlsxFormulaMetadata {
    /// Ordinary, non-shared formula.
    Normal {
        /// Formula text.
        formula: String,
    },

    /// Shared formula anchor/template cell.
    Shared {
        /// Shared formula index (`si`). Shared formula indices are worksheet-local.
        shared_index: usize,
        /// Shared formula range, when present.
        range: Option<Dimensions>,
        /// Template formula text as stored on the anchor cell.
        formula: String,
    },

    /// Shared formula derived cell.
    SharedDerived {
        /// Shared formula index (`si`). Shared formula indices are worksheet-local.
        shared_index: usize,
    },
}

impl XlsxFormulaMetadata {
    /// Return the shared formula index when this record belongs to a shared group.
    pub fn shared_index(&self) -> Option<usize> {
        match self {
            Self::Normal { .. } => None,
            Self::Shared { shared_index, .. } | Self::SharedDerived { shared_index } => {
                Some(*shared_index)
            }
        }
    }
}

/// Internal formula metadata used while streaming a cell.
///
/// This mirrors the public [`XlsxFormulaMetadata`] but can also carry an
/// expanded formula string for shared-formula derived cells when callers ask
/// for the convenience text API. The public metadata API intentionally drops
/// that string to avoid per-derived-cell expansion and allocation.
#[derive(Clone, Debug, PartialEq, Eq)]
enum FormulaMetadata {
    Normal {
        formula: String,
    },

    Shared {
        shared_index: usize,
        range: Option<Dimensions>,
        formula: String,
    },

    SharedDerived {
        shared_index: usize,
        translated_formula: Option<String>,
    },
}

impl FormulaMetadata {
    fn formula_text(&self) -> Option<&str> {
        match self {
            Self::Normal { formula } | Self::Shared { formula, .. } => Some(formula),
            Self::SharedDerived {
                translated_formula, ..
            } => translated_formula.as_deref(),
        }
    }

    fn into_metadata(self) -> XlsxFormulaMetadata {
        match self {
            Self::Normal { formula } => XlsxFormulaMetadata::Normal { formula },
            Self::Shared {
                shared_index,
                range,
                formula,
            } => XlsxFormulaMetadata::Shared {
                shared_index,
                range,
                formula,
            },
            Self::SharedDerived { shared_index, .. } => {
                XlsxFormulaMetadata::SharedDerived { shared_index }
            }
        }
    }
}

/// A single XLSX cell structure containing both cached/literal value and expanded formula text.
#[derive(Clone, Debug, PartialEq)]
pub struct XlsxCellFormula<'a> {
    /// Zero-based `(row, column)` cell position.
    pub pos: (u32, u32),
    /// Literal or cached value associated with the cell.
    pub value: DataRef<'a>,
    /// Formula text, expanded for shared formulas when the shared-formula anchor
    /// has already been observed in stream order.
    pub formula: Option<String>,
}

/// A single XLSX cell record containing cached/literal value and formula metadata.
#[derive(Clone, Debug, PartialEq)]
pub struct XlsxCellFormulaMetadataRecord<'a> {
    /// Zero-based `(row, column)` cell position.
    pub pos: (u32, u32),
    /// Literal or cached value associated with the cell.
    pub value: DataRef<'a>,
    /// Formula metadata, when the cell contains a formula.
    pub formula: Option<XlsxFormulaMetadata>,
}

struct XlsxCellFormulaMetadataRecordInternal<'a> {
    pos: (u32, u32),
    value: DataRef<'a>,
    formula: Option<FormulaMetadata>,
}

/// An xlsx Cell Iterator.
///
/// The `next_*` methods all advance the same XML stream. Use one streaming
/// method per reader; mixing methods on a single reader will consume cells from
/// the current stream position.
pub struct XlsxCellReader<'a, RS>
where
    RS: Read + Seek,
{
    xml: XlReader<'a, RS>,
    strings: &'a [String],
    formats: &'a [CellFormat],
    is_1904: bool,
    dimensions: Dimensions,
    row_index: u32,
    col_index: u32,
    buf: Vec<u8>,
    cell_buf: Vec<u8>,
    value_bufs: ValueBufs,
    formulas: Vec<Option<SharedFormula>>,
    // Fork addition: style id (`s` attribute) of the cell last returned by a `next_cell*` method.
    last_style: u32,
    // Fork addition: attributes of the current row, and of rows that had no cells.
    row_attrs: Option<RowAttributes>,
    row_had_cells: bool,
    cellless_rows: Vec<RowAttributes>,
    // Fork addition: `t` of the last cell's formula and, for array / data table
    // formulas, all `<f>` attributes.
    last_formula_kind: Option<String>,
    last_formula_attrs: Vec<(String, String)>,
}

impl<'a, RS> XlsxCellReader<'a, RS>
where
    RS: Read + Seek,
{
    /// Create a new XLSX cell reader over a worksheet XML stream.
    pub fn new(
        mut xml: XlReader<'a, RS>,
        strings: &'a [String],
        formats: &'a [CellFormat],
        is_1904: bool,
    ) -> Result<Self, XlsxError> {
        let mut buf = Vec::with_capacity(1024);
        let mut dimensions = Dimensions::default();
        let mut sh_type = None;
        'xml: loop {
            buf.clear();
            match xml.read_event_into(&mut buf).map_err(XlsxError::Xml)? {
                Event::Start(e) => match e.local_name().as_ref() {
                    b"dimension" => {
                        if let Some(rdim) = e.raw_attr(b"ref")? {
                            dimensions = get_dimension(rdim)?;
                            continue 'xml;
                        }
                        return Err(XlsxError::UnexpectedNode("dimension"));
                    }
                    b"sheetData" => break,
                    typ => {
                        if sh_type.is_none() {
                            sh_type = Some(xml.decoder().decode(typ)?.to_string());
                        }
                    }
                },
                Event::Eof => {
                    if let Some(typ) = sh_type {
                        return Err(XlsxError::NotAWorksheet(typ));
                    } else {
                        return Err(XlsxError::XmlEof("worksheet"));
                    }
                }
                _ => (),
            }
        }
        Ok(Self {
            xml,
            strings,
            formats,
            is_1904,
            dimensions,
            row_index: 0,
            col_index: 0,
            buf: Vec::with_capacity(1024),
            cell_buf: Vec::with_capacity(1024),
            value_bufs: ValueBufs::new(),
            formulas: Vec::with_capacity(1024),
            last_style: 0,
            row_attrs: None,
            row_had_cells: false,
            cellless_rows: Vec::new(),
            last_formula_kind: None,
            last_formula_attrs: Vec::new(),
        })
    }

    /// Fork addition: style id (index into `Xlsx::cell_styles`) of the cell last
    /// returned by `next_cell` or `next_cell_with_formula*` (0, the default style,
    /// when the cell has none).
    pub fn last_style_id(&self) -> u32 {
        self.last_style
    }

    /// Fork addition: non-default attributes (height, hidden, outline) of the row
    /// containing the cell last returned by `next_cell` / `next_cell_with_formula*`.
    pub fn current_row_attributes(&self) -> Option<&RowAttributes> {
        self.row_attrs.as_ref()
    }

    /// Fork addition: takes the attributes of rows read so far that had no cells
    /// (e.g. hidden or resized empty rows).
    pub fn take_cellless_rows(&mut self) -> Vec<RowAttributes> {
        std::mem::take(&mut self.cellless_rows)
    }

    /// Fork addition: capture the formatting runs of inline rich strings
    /// (see `take_inline_runs`). Off by default.
    pub fn set_capture_inline_runs(&mut self, capture: bool) {
        self.value_bufs.capture_runs = capture;
    }

    /// Fork addition: kind of the formula of the cell last returned by
    /// `next_cell_with_formula*`: `normal`, `shared`, `array` or `dataTable`.
    pub fn last_formula_kind(&self) -> Option<&str> {
        self.last_formula_kind.as_deref()
    }

    /// Fork addition: `<f>` attributes (`ref`, `r1`, `dt2D`, ...) of the cell last
    /// returned by `next_cell_with_formula*`, for array and data table formulas.
    pub fn last_formula_attributes(&self) -> &[(String, String)] {
        &self.last_formula_attrs
    }

    /// Fork addition: formatting runs of the cell last returned by `next_cell` /
    /// `next_cell_with_formula*`, if it is an inline rich string.
    pub fn take_inline_runs(&mut self) -> Option<Vec<TextRun>> {
        self.value_bufs.last_runs.take()
    }

    /// Fork addition: shared string index of the cell last returned by
    /// `next_cell` / `next_cell_with_formula*`, if it is a shared string.
    pub fn last_shared_string_index(&self) -> Option<usize> {
        self.value_bufs.last_sst
    }

    fn end_row(&mut self) {
        if !self.row_had_cells {
            if let Some(attrs) = self.row_attrs.take() {
                self.cellless_rows.push(attrs);
            }
        }
        self.row_index += 1;
        self.col_index = 0;
    }

    /// Return the worksheet dimensions declared by the sheet XML.
    pub fn dimensions(&self) -> Dimensions {
        self.dimensions
    }

    /// Return the next cell value in XML stream order.
    pub fn next_cell(&mut self) -> Result<Option<Cell<DataRef<'a>>>, XlsxError> {
        loop {
            self.buf.clear();
            match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(row_element)) if row_element.local_name().as_ref() == b"row" => {
                    let (index, attrs) = parse_row(&row_element, self.row_index)?;
                    self.row_index = index;
                    self.row_attrs = attrs;
                    self.row_had_cells = false;
                }
                Ok(Event::End(row_element)) if row_element.local_name().as_ref() == b"row" => {
                    self.end_row();
                }
                Ok(Event::Start(c_element)) if c_element.local_name().as_ref() == b"c" => {
                    self.row_had_cells = true;
                    self.value_bufs.last_sst = None;
                    self.value_bufs.last_runs = None;
                    self.last_formula_kind = None;
                    self.last_formula_attrs.clear();
                    let (pos_attr, style_attr, type_attr) =
                        get_attrs!(c_element, b"r" => r, b"s" => s, b"t" => t)?;
                    let pos = if let Some(range) = pos_attr {
                        let (row, col) = get_row_column(range)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };
                    self.last_style = style_attr
                        .and_then(|s| atoi_simd::parse::<u32, true, false>(s).ok())
                        .unwrap_or(0);
                    let mut value = DataRef::Empty;
                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::Start(e)) => {
                                let ctx = WorkbookContext {
                                    strings: self.strings,
                                    formats: self.formats,
                                    is_1904: self.is_1904,
                                };
                                value = read_value(
                                    &ctx,
                                    &mut self.xml,
                                    &e,
                                    style_attr,
                                    type_attr,
                                    &mut self.value_bufs,
                                )?;
                            }
                            Ok(Event::End(e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;
                    return Ok(Some(Cell::new(pos, value)));
                }
                Ok(Event::End(e)) if e.local_name().as_ref() == b"sheetData" => {
                    return Ok(None);
                }
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
    }

    fn read_formula_record(
        xml: &mut XlReader<'_, RS>,
        formulas: &mut Vec<Option<SharedFormula>>,
        e: &BytesStart<'_>,
        pos: (u32, u32),
        expand_shared_derived: bool,
    ) -> Result<Option<FormulaMetadata>, XlsxError> {
        let formula = read_formula(xml, e)?;

        let (t_attr, si_attr, ref_attr) = get_attrs!(e, b"t" => t, b"si" => si, b"ref" => ref_)?;
        if t_attr == Some(b"shared".as_slice()) {
            let shared_index = match si_attr {
                Some(res) => match atoi_simd::parse::<usize, true, false>(res) {
                    Ok(res) => res,
                    Err(_) => return Err(XlsxError::Unexpected("si attribute must be a number")),
                },
                None => {
                    return Err(XlsxError::Unexpected(
                        "si attribute is mandatory if it is shared",
                    ));
                }
            };

            return match ref_attr {
                Some(res) => {
                    let range = get_dimension(res)?;
                    let formula = formula.unwrap_or_default();
                    if expand_shared_derived {
                        if formulas.len() <= shared_index {
                            formulas.resize(shared_index + 1, None);
                        }
                        formulas[shared_index] = Some(SharedFormula {
                            formula: formula.clone(),
                            range,
                        });
                    }
                    Ok(Some(FormulaMetadata::Shared {
                        shared_index,
                        range: Some(range),
                        formula,
                    }))
                }
                None => {
                    let translated_formula = if expand_shared_derived {
                        formulas
                            .get(shared_index)
                            .and_then(|template| template.as_ref())
                            .map(|template| {
                                expand_shared_formula(&template.formula, template.range.start, pos)
                            })
                            .transpose()?
                    } else {
                        None
                    };
                    Ok(Some(FormulaMetadata::SharedDerived {
                        shared_index,
                        translated_formula,
                    }))
                }
            };
        }

        Ok(formula.map(|formula| FormulaMetadata::Normal { formula }))
    }

    /// Return the next cell record, exposing cached/literal value plus expanded
    /// per-cell formula text. Shared-formula metadata is intentionally not
    /// exposed through this compatibility-oriented one-pass API; use
    /// [`Self::next_cell_with_formula_metadata`] when shared-formula metadata is needed.
    pub fn next_cell_with_formula(&mut self) -> Result<Option<XlsxCellFormula<'a>>, XlsxError> {
        Ok(self
            .next_cell_formula_record_impl(true)?
            .map(|record| XlsxCellFormula {
                pos: record.pos,
                value: record.value,
                formula: record
                    .formula
                    .and_then(|formula| formula.formula_text().map(str::to_string)),
            }))
    }

    /// Return the next cell record, exposing cached/literal value plus formula
    /// metadata. Shared formulas are reported semantically as anchors/derived
    /// placements instead of only as expanded text. Derived shared formulas carry
    /// only their shared index, avoiding per-cell formula expansion/allocation.
    pub fn next_cell_with_formula_metadata(
        &mut self,
    ) -> Result<Option<XlsxCellFormulaMetadataRecord<'a>>, XlsxError> {
        Ok(self
            .next_cell_formula_record_impl(false)?
            .map(|record| XlsxCellFormulaMetadataRecord {
                pos: record.pos,
                value: record.value,
                formula: record.formula.map(FormulaMetadata::into_metadata),
            }))
    }

    fn next_cell_formula_record_impl(
        &mut self,
        expand_shared_derived: bool,
    ) -> Result<Option<XlsxCellFormulaMetadataRecordInternal<'a>>, XlsxError> {
        loop {
            self.buf.clear();
            match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(row_element)) if row_element.local_name().as_ref() == b"row" => {
                    let (index, attrs) = parse_row(&row_element, self.row_index)?;
                    self.row_index = index;
                    self.row_attrs = attrs;
                    self.row_had_cells = false;
                }
                Ok(Event::End(row_element)) if row_element.local_name().as_ref() == b"row" => {
                    self.end_row();
                }
                Ok(Event::Start(c_element)) if c_element.local_name().as_ref() == b"c" => {
                    self.row_had_cells = true;
                    self.value_bufs.last_sst = None;
                    self.value_bufs.last_runs = None;
                    self.last_formula_kind = None;
                    self.last_formula_attrs.clear();
                    let (pos_attr, style_attr, type_attr) =
                        get_attrs!(c_element, b"r" => r, b"s" => s, b"t" => t)?;
                    let pos = if let Some(range) = pos_attr {
                        let (row, col) = get_row_column(range)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };
                    self.last_style = style_attr
                        .and_then(|s| atoi_simd::parse::<u32, true, false>(s).ok())
                        .unwrap_or(0);
                    let mut value = DataRef::Empty;
                    let mut formula = None;
                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::Start(e)) if e.local_name().as_ref() == b"f" => {
                                let kind = e.raw_attr(b"t")?;
                                if matches!(kind, Some(b"array") | Some(b"dataTable")) {
                                    for item in e.iter_raw_attrs() {
                                        let (k, v) = item?;
                                        self.last_formula_attrs.push((
                                            String::from_utf8_lossy(k).into_owned(),
                                            String::from_utf8_lossy(v).into_owned(),
                                        ));
                                    }
                                }
                                self.last_formula_kind = Some(kind.map_or("normal".into(), |t| {
                                    String::from_utf8_lossy(t).into_owned()
                                }));
                                formula = Self::read_formula_record(
                                    &mut self.xml,
                                    &mut self.formulas,
                                    &e,
                                    pos,
                                    expand_shared_derived,
                                )?;
                            }
                            Ok(Event::Start(e)) => {
                                let ctx = WorkbookContext {
                                    strings: self.strings,
                                    formats: self.formats,
                                    is_1904: self.is_1904,
                                };
                                value = read_value(
                                    &ctx,
                                    &mut self.xml,
                                    &e,
                                    style_attr,
                                    type_attr,
                                    &mut self.value_bufs,
                                )?;
                            }
                            Ok(Event::End(e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;
                    return Ok(Some(XlsxCellFormulaMetadataRecordInternal {
                        pos,
                        value,
                        formula,
                    }));
                }
                Ok(Event::End(e)) if e.local_name().as_ref() == b"sheetData" => {
                    return Ok(None);
                }
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
    }

    /// Return the next formula in XML stream order, expanding shared formulas.
    pub fn next_formula(&mut self) -> Result<Option<Cell<String>>, XlsxError> {
        loop {
            self.buf.clear();
            match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(row_element)) if row_element.local_name().as_ref() == b"row" => {
                    if let Some(r) = row_element.raw_attr(b"r")? {
                        self.row_index = get_row(r)?;
                    }
                }
                Ok(Event::End(row_element)) if row_element.local_name().as_ref() == b"row" => {
                    self.row_index += 1;
                    self.col_index = 0;
                }
                Ok(Event::Start(c_element)) if c_element.local_name().as_ref() == b"c" => {
                    let pos = if let Some(r) = c_element.raw_attr(b"r")? {
                        let (row, col) = get_row_column(r)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };
                    let mut value = None;
                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::Start(e)) => {
                                let formula_record = Self::read_formula_record(
                                    &mut self.xml,
                                    &mut self.formulas,
                                    &e,
                                    pos,
                                    true,
                                )?;
                                if let Some(formula_text) = formula_record
                                    .and_then(|record| record.formula_text().map(str::to_string))
                                {
                                    value = Some(formula_text);
                                }
                            }
                            Ok(Event::End(e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;
                    return Ok(Some(Cell::new(pos, value.unwrap_or_default())));
                }
                Ok(Event::End(e)) if e.local_name().as_ref() == b"sheetData" => {
                    return Ok(None);
                }
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
    }
}

/// Reads a cell value using pre-extracted `s` and `t` attributes
/// (avoids repeating attribute iteration on the `<c>` element).
fn read_value<'s, RS>(
    ctx: &WorkbookContext<'s>,
    xml: &mut XlReader<'_, RS>,
    e: &BytesStart<'_>,
    style_attr: Option<&[u8]>,
    type_attr: Option<&[u8]>,
    bufs: &mut ValueBufs,
) -> Result<DataRef<'s>, XlsxError>
where
    RS: Read + Seek,
{
    Ok(match e.local_name().as_ref() {
        b"is" if bufs.capture_runs => {
            // inlineStr, keeping its formatting runs (fork addition)
            let closing = e.local_name().as_ref().to_vec();
            let (text, runs) = read_inline_rich(xml, &closing, &mut bufs.xml, &mut bufs.str_inner)?;
            bufs.last_runs = (!runs.is_empty()).then_some(runs);
            DataRef::String(text)
        }
        b"is" => {
            // inlineStr
            read_string_with_bufs(xml, e.name(), &mut bufs.xml, &mut bufs.str_inner)?
                .map_or(DataRef::Empty, DataRef::String)
        }
        // Ignore <v> for inlineStr cells since it is redundant. The value is in
        // the <is> element, which is handled above.
        b"v" if matches!(type_attr, Some(b"inlineStr") | Some(b"is")) => {
            bufs.xml.clear();
            xml.read_to_end_into(e.name(), &mut bufs.xml)?;
            DataRef::Empty
        }
        b"v" => match type_attr {
            Some(b"n") | Some(b"s") | Some(b"b") | Some(b"e") | None => {
                // These types are always plain ASCII (no CR/LF or entities), so we can
                // parse directly from raw bytes, skipping `xml10_content()` + String
                bufs.xml.clear();
                let val = match xml.read_event_into(&mut bufs.xml)? {
                    Event::Text(t) => {
                        if type_attr == Some(b"s") {
                            bufs.last_sst = atoi_simd::parse::<usize, true, false>(&t).ok();
                        }
                        read_v(ctx, &t, style_attr, type_attr)?
                    }
                    Event::End(end) if end.name() == e.name() => return Ok(DataRef::Empty),
                    Event::Eof => return Err(XlsxError::XmlEof("v")),
                    _ => DataRef::Empty,
                };
                bufs.xml.clear();
                xml.read_to_end_into(e.name(), &mut bufs.xml)?;
                val
            }
            _ => {
                // Types that may contain entities, or need owned Strings (eg: "str", "d")
                bufs.value.clear();
                loop {
                    bufs.xml.clear();
                    match xml.read_event_into(&mut bufs.xml)? {
                        Event::Text(t) => bufs.value.push_str(&t.xml10_content()?),
                        Event::GeneralRef(e) => unescape_entity_to_buffer(&e, &mut bufs.value)?,
                        Event::End(end) if end.name() == e.name() => break,
                        Event::Eof => return Err(XlsxError::XmlEof("v")),
                        _ => (),
                    }
                }
                read_v(ctx, bufs.value.as_bytes(), style_attr, type_attr)?
            }
        },
        b"f" => {
            bufs.xml.clear();
            xml.read_to_end_into(e.name(), &mut bufs.xml)?;
            DataRef::Empty
        }
        _n => return Err(XlsxError::UnexpectedNode("v, f, or is")),
    })
}

/// Convert raw `<v>` bytes to a `&str`, returning an error on invalid UTF-8.
fn v_as_str(v: &[u8]) -> Result<&str, XlsxError> {
    std::str::from_utf8(v).map_err(|_| XlsxError::Unexpected("invalid UTF-8 in cell value"))
}

/// Parse a `<v>` cell value from raw bytes with pre-extracted
/// `s` (style) and `t` (type) attributes.
fn read_v<'s>(
    ctx: &WorkbookContext<'s>,
    v: &[u8],
    style_attr: Option<&[u8]>,
    type_attr: Option<&[u8]>,
) -> Result<DataRef<'s>, XlsxError> {
    let cell_format = match style_attr {
        Some(style) => {
            let id = atoi_simd::parse::<usize, true, false>(style).unwrap_or(0);
            ctx.formats.get(id)
        }
        None => Some(&CellFormat::Other),
    };
    match type_attr {
        Some(b"s") => {
            if v.is_empty() {
                return Ok(DataRef::Empty);
            }
            let idx = atoi_simd::parse::<usize, true, false>(v).unwrap_or(0);
            ctx.strings
                .get(idx)
                .map(|s| DataRef::SharedString(s))
                .ok_or(XlsxError::Unexpected(
                    "Cell string index not found in shared strings table",
                ))
        }
        Some(b"b") => Ok(DataRef::Bool(v != b"0")),
        Some(b"d") => Ok(DataRef::DateTimeIso(v_as_str(v)?.to_string())),
        Some(b"e") => Ok(DataRef::Error(v_as_str(v)?.parse()?)),
        Some(b"str") => Ok(DataRef::String(v_as_str(v)?.to_string())),
        Some(b"n") | None => {
            if v.is_empty() {
                return Ok(DataRef::Empty);
            }
            // Fork addition: whole numbers written without a decimal point or exponent
            // ("5", "-12") are integers, as in openpyxl. Dates/durations stay floats.
            let is_date = matches!(
                cell_format,
                Some(CellFormat::DateTime) | Some(CellFormat::TimeDelta)
            );
            let digits = v.strip_prefix(b"-").unwrap_or(v);
            if !is_date
                && !digits.is_empty()
                && digits.len() <= 18
                && digits.iter().all(u8::is_ascii_digit)
            {
                if let Ok(n) = v_as_str(v)?.parse::<i64>() {
                    return Ok(DataRef::Int(n));
                }
            }
            // If type is not known, we try to parse as Float for utility, but fall back to
            // String if this fails.
            fast_float2::parse::<f64, _>(v)
                .map(|n| format_excel_f64_ref(n, cell_format, ctx.is_1904))
                .or_else(|_| {
                    if type_attr.is_none() {
                        // No explicit type: fall back to String if not a valid float
                        Ok(DataRef::String(v_as_str(v)?.to_string()))
                    } else {
                        Err(XlsxError::ParseFloat(
                            v_as_str(v)?.parse::<f64>().unwrap_err(),
                        ))
                    }
                })
        }
        Some(t) => {
            let t = std::str::from_utf8(t).unwrap_or("<utf8 error>").to_string();
            Err(XlsxError::CellTAttribute(t))
        }
    }
}

/// Fork addition: row index and non-default attributes of a `<row>` element.
pub(crate) fn parse_row(
    row_element: &BytesStart<'_>,
    mut row_index: u32,
) -> Result<(u32, Option<RowAttributes>), XlsxError> {
    let (r, ht, custom_height, hidden, outline, collapsed, style, custom_format) = get_attrs!(
        row_element,
        b"r" => r,
        b"ht" => ht,
        b"customHeight" => ch,
        b"hidden" => hd,
        b"outlineLevel" => ol,
        b"collapsed" => co,
        b"s" => s,
        b"customFormat" => cf,
    )?;
    if let Some(r) = r {
        row_index = get_row(r)?;
    }
    let truthy = |v: Option<&[u8]>| matches!(v, Some(b"1") | Some(b"true"));
    let attrs = RowAttributes {
        index: row_index,
        height: ht.and_then(|v| std::str::from_utf8(v).ok()?.parse().ok()),
        custom_height: truthy(custom_height),
        hidden: truthy(hidden),
        outline_level: outline
            .and_then(|v| atoi_simd::parse::<u8, true, false>(v).ok())
            .unwrap_or(0),
        collapsed: truthy(collapsed),
        style: if truthy(custom_format) {
            style.and_then(|v| atoi_simd::parse::<u32, true, false>(v).ok())
        } else {
            None
        },
    };
    let is_default = attrs.height.is_none()
        && !attrs.hidden
        && attrs.outline_level == 0
        && !attrs.collapsed
        && attrs.style.is_none();
    Ok((row_index, (!is_default).then_some(attrs)))
}

fn read_formula<RS>(xml: &mut XlReader<RS>, e: &BytesStart) -> Result<Option<String>, XlsxError>
where
    RS: Read + Seek,
{
    match e.local_name().as_ref() {
        b"is" | b"v" => {
            xml.read_to_end_into(e.name(), &mut Vec::new())?;
            Ok(None)
        }
        b"f" => {
            let mut f_buf = Vec::with_capacity(512);
            let mut f = String::new();
            loop {
                match xml.read_event_into(&mut f_buf)? {
                    Event::Text(t) => f.push_str(&t.xml10_content()?),
                    Event::GeneralRef(e) => unescape_entity_to_buffer(&e, &mut f)?,
                    Event::End(end) if end.name() == e.name() => break,
                    Event::Eof => return Err(XlsxError::XmlEof("f")),
                    _ => (),
                }
                f_buf.clear();
            }
            Ok(Some(f))
        }
        _ => Err(XlsxError::UnexpectedNode("v, f, or is")),
    }
}
