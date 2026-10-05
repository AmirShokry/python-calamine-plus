//! Read-only style objects shaped like openpyxl's (`Font`, `PatternFill`, `Border`, ...).
//!
//! Attribute names, aliases, defaults and value types follow openpyxl 3.1 as it reads a
//! file. Differences: colors that are not RGB colors have `rgb = None` (openpyxl
//! returns a descriptor object there), and objects are immutable.

use std::collections::hash_map::DefaultHasher;
use std::fmt;
use std::hash::{Hash, Hasher};

use calamine::{
    CellStyle, StyleAlignment, StyleBorder, StyleBorderSide, StyleColor, StyleFill, StyleFont,
};
use pyo3::prelude::*;
use pyo3::sync::PyOnceLock;

use super::utils::{builtin_format, is_date_format, is_timedelta_format};

/// A Python object built on first access and then reused (style objects are shared by
/// every cell with the same format, so their children are too). Ignored by comparisons,
/// hashing and `repr`; clones start empty.
pub(crate) struct Lazy<T>(PyOnceLock<T>);

impl<T> Default for Lazy<T> {
    fn default() -> Self {
        Self(PyOnceLock::new())
    }
}

impl<T> Clone for Lazy<T> {
    fn clone(&self) -> Self {
        Self::default()
    }
}

impl<T> PartialEq for Lazy<T> {
    fn eq(&self, _: &Self) -> bool {
        true
    }
}

impl<T> fmt::Debug for Lazy<T> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("_")
    }
}

impl<T> Lazy<T> {
    fn get(&self, py: Python<'_>, init: impl FnOnce() -> PyResult<T>) -> PyResult<&T> {
        self.0.get_or_try_init(py, init)
    }
}

/// Cached `Option<Py<T>>` of an optional child value.
fn lazy_opt<T, V>(
    py: Python<'_>,
    cache: &Lazy<Option<Py<T>>>,
    value: &Option<V>,
) -> PyResult<Option<Py<T>>>
where
    T: pyo3::PyClass + Into<pyo3::PyClassInitializer<T>>,
    V: Clone + Into<T>,
{
    let cached = cache.get(py, || {
        value.clone().map(|v| Py::new(py, v.into())).transpose()
    })?;
    Ok(cached.as_ref().map(|o| o.clone_ref(py)))
}

/// Cached `Py<T>` of a child value.
fn lazy_one<T, V>(py: Python<'_>, cache: &Lazy<Py<T>>, value: &V) -> PyResult<Py<T>>
where
    T: pyo3::PyClass + Into<pyo3::PyClassInitializer<T>>,
    V: Clone + Into<T>,
{
    Ok(cache
        .get(py, || Py::new(py, value.clone().into()))?
        .clone_ref(py))
}

fn hash_of(s: &str) -> u64 {
    let mut h = DefaultHasher::new();
    s.hash(&mut h);
    h.finish()
}

/// One `#[pymethods]` block per class: shared `__eq__` / `__hash__` / `__repr__`
/// plus the class's own methods.
macro_rules! value_object {
    ($ty:ident { $($body:tt)* }) => {
        #[pymethods]
        impl $ty {
            $($body)*

            fn __eq__(&self, other: &Bound<'_, PyAny>) -> bool {
                other
                    .cast::<$ty>()
                    .map(|o| *o.get() == *self)
                    .unwrap_or(false)
            }

            fn __hash__(&self) -> u64 {
                hash_of(&format!("{self:?}"))
            }

            fn __repr__(&self) -> String {
                format!("<{} {:?}>", stringify!($ty), self)
            }
        }
    };
}

// ---- Color -------------------------------------------------------------------------

/// A color as openpyxl reads it: `type` is `rgb`, `theme`, `indexed` or `auto`, and
/// only the attribute of that type is set. Theme and indexed colors are not resolved.
#[pyclass(
    frozen,
    skip_from_py_object,
    module = "python_calamine_plus.openpyxl.styles"
)]
#[derive(Clone, Debug, PartialEq)]
pub struct Color {
    #[pyo3(get)]
    rgb: Option<String>,
    #[pyo3(get)]
    indexed: Option<u32>,
    #[pyo3(get)]
    theme: Option<u32>,
    #[pyo3(get)]
    auto: Option<bool>,
    #[pyo3(get)]
    tint: f64,
    #[pyo3(get, name = "type")]
    typ: &'static str,
}

impl Color {
    /// openpyxl's `Color()`: black RGB `00000000`.
    pub(crate) fn black() -> Self {
        Self {
            rgb: Some("00000000".into()),
            indexed: None,
            theme: None,
            auto: None,
            tint: 0.0,
            typ: "rgb",
        }
    }

    pub(crate) fn from_style(c: &StyleColor) -> Self {
        let mut out = Self {
            rgb: None,
            indexed: None,
            theme: None,
            auto: None,
            tint: c.tint.unwrap_or(0.0),
            typ: "rgb",
        };
        if let Some(i) = c.indexed {
            out.typ = "indexed";
            out.indexed = Some(i);
        } else if let Some(t) = c.theme {
            out.typ = "theme";
            out.theme = Some(t);
        } else if c.auto {
            out.typ = "auto";
            out.auto = Some(true);
        } else {
            let rgb = c.rgb.clone().unwrap_or_else(|| "00000000".into());
            out.rgb = Some(if rgb.len() == 6 {
                format!("00{rgb}")
            } else {
                rgb
            });
        }
        out
    }
}

value_object!(Color {
    /// The value of the color's type (`rgb` string, theme / indexed number, ...).
    #[getter]
    fn value(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        Ok(match self.typ {
            "indexed" => self.indexed.into_pyobject(py)?.into_any().unbind(),
            "theme" => self.theme.into_pyobject(py)?.into_any().unbind(),
            "auto" => self.auto.into_pyobject(py)?.into_any().unbind(),
            _ => self.rgb.clone().into_pyobject(py)?.into_any().unbind(),
        })
    }

    /// Legacy alias of `value`.
    #[getter]
    fn index(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        self.value(py)
    }
});

// ---- Font --------------------------------------------------------------------------

#[pyclass(
    frozen,
    skip_from_py_object,
    module = "python_calamine_plus.openpyxl.styles"
)]
#[derive(Clone, Debug, PartialEq)]
pub struct Font {
    #[pyo3(get)]
    name: Option<String>,
    #[pyo3(get)]
    sz: Option<f64>,
    #[pyo3(get)]
    b: bool,
    #[pyo3(get)]
    i: bool,
    #[pyo3(get)]
    u: Option<String>,
    #[pyo3(get)]
    strike: Option<bool>,
    #[pyo3(get, name = "vertAlign")]
    vert_align: Option<String>,
    color: Option<Color>,
    #[pyo3(get)]
    family: Option<f64>,
    #[pyo3(get)]
    charset: Option<u32>,
    #[pyo3(get)]
    scheme: Option<String>,
    #[pyo3(get)]
    outline: Option<bool>,
    #[pyo3(get)]
    shadow: Option<bool>,
    #[pyo3(get)]
    condense: Option<bool>,
    #[pyo3(get)]
    extend: Option<bool>,
    color_py: Lazy<Option<Py<Color>>>,
}

impl Font {
    pub(crate) fn from_style(f: &StyleFont) -> Self {
        Self {
            name: f.name.clone(),
            sz: f.size,
            b: f.bold,
            i: f.italic,
            u: f.underline.clone(),
            strike: f.raw_strike,
            vert_align: f.raw_vert_align.clone().filter(|v| v != "none"),
            color: f.color.as_ref().map(Color::from_style),
            family: f.family.map(f64::from),
            charset: f.charset,
            scheme: f.scheme.clone(),
            outline: f.raw_outline,
            shadow: f.raw_shadow,
            condense: f.raw_condense,
            extend: f.raw_extend,
            color_py: Lazy::default(),
        }
    }

    /// openpyxl's default font, used when a file has no stylesheet.
    pub(crate) fn openpyxl_default() -> Self {
        Self {
            name: Some("Calibri".into()),
            sz: Some(11.0),
            b: false,
            i: false,
            u: None,
            strike: None,
            vert_align: None,
            color: Some(Color {
                rgb: None,
                indexed: None,
                theme: Some(1),
                auto: None,
                tint: 0.0,
                typ: "theme",
            }),
            family: Some(2.0),
            charset: None,
            scheme: Some("minor".into()),
            outline: None,
            shadow: None,
            condense: None,
            extend: None,
            color_py: Lazy::default(),
        }
    }
}

value_object!(Font {
    #[getter]
    fn color(&self, py: Python<'_>) -> PyResult<Option<Py<Color>>> {
        lazy_opt(py, &self.color_py, &self.color)
    }
    #[getter]
    fn size(&self) -> Option<f64> {
        self.sz
    }
    #[getter]
    fn bold(&self) -> bool {
        self.b
    }
    #[getter]
    fn italic(&self) -> bool {
        self.i
    }
    #[getter]
    fn underline(&self) -> Option<String> {
        self.u.clone()
    }
    #[getter]
    fn strikethrough(&self) -> Option<bool> {
        self.strike
    }
});

// ---- Fills -------------------------------------------------------------------------

#[pyclass(
    frozen,
    skip_from_py_object,
    module = "python_calamine_plus.openpyxl.styles"
)]
#[derive(Clone, Debug, PartialEq)]
pub struct PatternFill {
    #[pyo3(get, name = "patternType")]
    pattern_type: Option<String>,
    fg: Color,
    bg: Color,
    fg_py: Lazy<Py<Color>>,
    bg_py: Lazy<Py<Color>>,
}

value_object!(PatternFill {
    #[getter]
    fn fill_type(&self) -> Option<String> {
        self.pattern_type.clone()
    }
    #[getter(fgColor)]
    fn fg_color(&self, py: Python<'_>) -> PyResult<Py<Color>> {
        lazy_one(py, &self.fg_py, &self.fg)
    }
    #[getter]
    fn start_color(&self, py: Python<'_>) -> PyResult<Py<Color>> {
        self.fg_color(py)
    }
    #[getter(bgColor)]
    fn bg_color(&self, py: Python<'_>) -> PyResult<Py<Color>> {
        lazy_one(py, &self.bg_py, &self.bg)
    }
    #[getter]
    fn end_color(&self, py: Python<'_>) -> PyResult<Py<Color>> {
        self.bg_color(py)
    }
    #[getter]
    fn tagname(&self) -> &'static str {
        "patternFill"
    }
});

#[pyclass(
    frozen,
    skip_from_py_object,
    module = "python_calamine_plus.openpyxl.styles"
)]
#[derive(Clone, Debug, PartialEq)]
pub struct Stop {
    color: Color,
    #[pyo3(get)]
    position: f64,
    color_py: Lazy<Py<Color>>,
}

value_object!(Stop {
    #[getter]
    fn color(&self, py: Python<'_>) -> PyResult<Py<Color>> {
        lazy_one(py, &self.color_py, &self.color)
    }
});

#[pyclass(
    frozen,
    skip_from_py_object,
    module = "python_calamine_plus.openpyxl.styles"
)]
#[derive(Clone, Debug, PartialEq)]
pub struct GradientFill {
    #[pyo3(get, name = "type")]
    typ: String,
    #[pyo3(get)]
    degree: f64,
    #[pyo3(get)]
    left: f64,
    #[pyo3(get)]
    right: f64,
    #[pyo3(get)]
    top: f64,
    #[pyo3(get)]
    bottom: f64,
    stops: Vec<Stop>,
    stops_py: Lazy<Vec<Py<Stop>>>,
}

value_object!(GradientFill {
    #[getter]
    fn fill_type(&self) -> String {
        self.typ.clone()
    }
    #[getter]
    fn stop(&self, py: Python<'_>) -> PyResult<Vec<Py<Stop>>> {
        let stops = self.stops_py.get(py, || {
            self.stops.iter().map(|s| Py::new(py, s.clone())).collect()
        })?;
        Ok(stops.iter().map(|s| s.clone_ref(py)).collect())
    }
    #[getter]
    fn tagname(&self) -> &'static str {
        "gradientFill"
    }
});

/// A fill as stored: pattern or gradient.
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum Fill {
    Pattern(PatternFill),
    Gradient(GradientFill),
}

impl Fill {
    pub(crate) fn from_style(f: &StyleFill) -> Self {
        if let Some(g) = &f.gradient {
            return Fill::Gradient(GradientFill {
                typ: g.typ.clone().unwrap_or_else(|| "linear".into()),
                degree: g.degree.unwrap_or(0.0),
                left: g.left.unwrap_or(0.0),
                right: g.right.unwrap_or(0.0),
                top: g.top.unwrap_or(0.0),
                bottom: g.bottom.unwrap_or(0.0),
                stops: g
                    .stops
                    .iter()
                    .map(|(pos, c)| Stop {
                        color: Color::from_style(c),
                        position: *pos,
                        color_py: Lazy::default(),
                    })
                    .collect(),
                stops_py: Lazy::default(),
            });
        }
        Fill::Pattern(PatternFill {
            pattern_type: f.pattern.clone(),
            fg: f
                .fg_color
                .as_ref()
                .map_or_else(Color::black, Color::from_style),
            bg: f
                .bg_color
                .as_ref()
                .map_or_else(Color::black, Color::from_style),
            fg_py: Lazy::default(),
            bg_py: Lazy::default(),
        })
    }

    pub(crate) fn to_py(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        Ok(match self {
            Fill::Pattern(p) => Py::new(py, p.clone())?.into_any(),
            Fill::Gradient(g) => Py::new(py, g.clone())?.into_any(),
        })
    }
}

// ---- Border ------------------------------------------------------------------------

#[pyclass(
    frozen,
    skip_from_py_object,
    module = "python_calamine_plus.openpyxl.styles"
)]
#[derive(Clone, Debug, PartialEq, Default)]
pub struct Side {
    #[pyo3(get)]
    pub(crate) style: Option<String>,
    pub(crate) color: Option<Color>,
    color_py: Lazy<Option<Py<Color>>>,
}

impl Side {
    fn from_style(s: &StyleBorderSide) -> Self {
        Self {
            style: s.style.clone(),
            color: s.color.as_ref().map(Color::from_style),
            color_py: Lazy::default(),
        }
    }

    /// openpyxl's `Side.__add__`.
    fn add(&self, other: &Side) -> Side {
        Side {
            style: self.style.clone().or_else(|| other.style.clone()),
            color: self.color.clone().or_else(|| other.color.clone()),
            color_py: Lazy::default(),
        }
    }
}

value_object!(Side {
    #[getter]
    fn color(&self, py: Python<'_>) -> PyResult<Option<Py<Color>>> {
        lazy_opt(py, &self.color_py, &self.color)
    }
    #[getter]
    fn border_style(&self) -> Option<String> {
        self.style.clone()
    }
});

/// Index of each edge in `Border::sides`.
pub(crate) mod edge {
    pub const LEFT: usize = 0;
    pub const RIGHT: usize = 1;
    pub const TOP: usize = 2;
    pub const BOTTOM: usize = 3;
    pub const DIAGONAL: usize = 4;
    pub const VERTICAL: usize = 5;
    pub const HORIZONTAL: usize = 6;
}

#[pyclass(
    frozen,
    skip_from_py_object,
    module = "python_calamine_plus.openpyxl.styles"
)]
#[derive(Clone, Debug, PartialEq)]
pub struct Border {
    /// left, right, top, bottom, diagonal, vertical, horizontal; `None` when absent.
    pub(crate) sides: [Option<Side>; 7],
    #[pyo3(get, name = "diagonalUp")]
    diagonal_up: bool,
    #[pyo3(get, name = "diagonalDown")]
    diagonal_down: bool,
    #[pyo3(get)]
    outline: bool,
    sides_py: [Lazy<Option<Py<Side>>>; 7],
}

impl Default for Border {
    fn default() -> Self {
        Self {
            sides: Default::default(),
            diagonal_up: false,
            diagonal_down: false,
            outline: true,
            sides_py: Default::default(),
        }
    }
}

impl Border {
    pub(crate) fn from_style(b: &StyleBorder) -> Self {
        let all = [
            &b.left,
            &b.right,
            &b.top,
            &b.bottom,
            &b.diagonal,
            &b.vertical,
            &b.horizontal,
        ];
        let mut sides: [Option<Side>; 7] = Default::default();
        for (i, s) in all.iter().enumerate() {
            if b.raw_sides & (1 << i) != 0 {
                sides[i] = Some(Side::from_style(s));
            }
        }
        Self {
            sides,
            diagonal_up: b.diagonal_up,
            diagonal_down: b.diagonal_down,
            outline: b.raw_outline.unwrap_or(true),
            sides_py: Default::default(),
        }
    }

    /// A border with only `side` set at `edge` (openpyxl's `Border(**{name: side})`).
    pub(crate) fn single(edge: usize, side: Option<Side>) -> Self {
        let mut b = Border::default();
        b.sides[edge] = side;
        b
    }

    /// openpyxl's `Border.__add__` (`Serialisable.__add__`).
    pub(crate) fn add(&self, other: &Border) -> Border {
        let mut sides: [Option<Side>; 7] = Default::default();
        for (i, slot) in sides.iter_mut().enumerate() {
            *slot = match (&self.sides[i], &other.sides[i]) {
                (Some(a), Some(b)) => Some(a.add(b)),
                (a, b) => a.clone().or_else(|| b.clone()),
            };
        }
        Border {
            sides,
            diagonal_up: self.diagonal_up || other.diagonal_up,
            diagonal_down: self.diagonal_down || other.diagonal_down,
            outline: self.outline || other.outline,
            sides_py: Default::default(),
        }
    }

    pub(crate) fn side(&self, edge: usize) -> Option<&Side> {
        self.sides[edge].as_ref()
    }
}

impl Border {
    fn side_obj(&self, py: Python<'_>, edge: usize) -> PyResult<Option<Py<Side>>> {
        lazy_opt(py, &self.sides_py[edge], &self.sides[edge])
    }
}

value_object!(Border {
    #[getter]
    fn left(&self, py: Python<'_>) -> PyResult<Option<Py<Side>>> {
        self.side_obj(py, edge::LEFT)
    }
    #[getter]
    fn right(&self, py: Python<'_>) -> PyResult<Option<Py<Side>>> {
        self.side_obj(py, edge::RIGHT)
    }
    #[getter]
    fn top(&self, py: Python<'_>) -> PyResult<Option<Py<Side>>> {
        self.side_obj(py, edge::TOP)
    }
    #[getter]
    fn bottom(&self, py: Python<'_>) -> PyResult<Option<Py<Side>>> {
        self.side_obj(py, edge::BOTTOM)
    }
    #[getter]
    fn diagonal(&self, py: Python<'_>) -> PyResult<Option<Py<Side>>> {
        self.side_obj(py, edge::DIAGONAL)
    }
    #[getter]
    fn vertical(&self, py: Python<'_>) -> PyResult<Option<Py<Side>>> {
        self.side_obj(py, edge::VERTICAL)
    }
    #[getter]
    fn horizontal(&self, py: Python<'_>) -> PyResult<Option<Py<Side>>> {
        self.side_obj(py, edge::HORIZONTAL)
    }
    /// `<start>` is read as `left` (always `None` here).
    #[getter]
    fn start(&self) -> Option<Py<Side>> {
        None
    }
    /// `<end>` is read as `right` (always `None` here).
    #[getter]
    fn end(&self) -> Option<Py<Side>> {
        None
    }
    #[getter]
    fn diagonal_direction(&self) -> Option<Py<PyAny>> {
        None
    }
});

// ---- Alignment / Protection --------------------------------------------------------

#[pyclass(
    frozen,
    skip_from_py_object,
    module = "python_calamine_plus.openpyxl.styles"
)]
#[derive(Clone, Debug, PartialEq)]
pub struct Alignment {
    #[pyo3(get)]
    horizontal: Option<String>,
    #[pyo3(get)]
    vertical: Option<String>,
    #[pyo3(get, name = "textRotation")]
    text_rotation: u32,
    #[pyo3(get, name = "wrapText")]
    wrap_text: Option<bool>,
    #[pyo3(get, name = "shrinkToFit")]
    shrink_to_fit: Option<bool>,
    #[pyo3(get)]
    indent: f64,
    #[pyo3(get, name = "relativeIndent")]
    relative_indent: f64,
    #[pyo3(get, name = "justifyLastLine")]
    justify_last_line: Option<bool>,
    #[pyo3(get, name = "readingOrder")]
    reading_order: f64,
}

impl Alignment {
    pub(crate) fn from_style(a: &StyleAlignment) -> Self {
        Self {
            horizontal: a.horizontal.clone(),
            vertical: a.vertical.clone(),
            text_rotation: a.text_rotation.unwrap_or(0),
            wrap_text: a.raw_wrap_text,
            shrink_to_fit: a.raw_shrink_to_fit,
            indent: a.indent.map_or(0.0, f64::from),
            relative_indent: a.relative_indent.map_or(0.0, f64::from),
            justify_last_line: a.raw_justify_last_line,
            reading_order: a.reading_order.map_or(0.0, f64::from),
        }
    }
}

value_object!(Alignment {
    #[getter]
    fn text_rotation(&self) -> u32 {
        self.text_rotation
    }
    #[getter]
    fn wrap_text(&self) -> Option<bool> {
        self.wrap_text
    }
    #[getter]
    fn shrink_to_fit(&self) -> Option<bool> {
        self.shrink_to_fit
    }
});

#[pyclass(
    frozen,
    skip_from_py_object,
    module = "python_calamine_plus.openpyxl.styles"
)]
#[derive(Clone, Debug, PartialEq)]
pub struct Protection {
    #[pyo3(get)]
    locked: bool,
    #[pyo3(get)]
    hidden: bool,
}
value_object!(Protection {});

// ---- per-format bundle -------------------------------------------------------------

/// Everything a cell format provides, converted once per format.
#[derive(Clone, Debug)]
pub(crate) struct CompatStyle {
    pub(crate) font: Font,
    pub(crate) fill: Fill,
    pub(crate) border: Border,
    pub(crate) alignment: Alignment,
    pub(crate) protection: Protection,
    pub(crate) number_format: String,
    pub(crate) is_date: bool,
    pub(crate) is_timedelta: bool,
    pub(crate) named_style: String,
    pub(crate) quote_prefix: bool,
    pub(crate) pivot_button: bool,
    /// openpyxl's `has_style`: the format is not the default one.
    pub(crate) has_style: bool,
}

impl CompatStyle {
    pub(crate) fn from_style(s: &CellStyle, index: usize) -> Self {
        // Custom formats keep their code; built-in ids use openpyxl's table
        // ("General" for ids openpyxl does not know).
        let number_format = if s.number_format_custom {
            s.number_format.clone()
        } else {
            builtin_format(s.number_format_id)
                .unwrap_or("General")
                .to_string()
        };
        Self {
            font: Font::from_style(&s.font),
            fill: Fill::from_style(&s.fill),
            border: Border::from_style(&s.border),
            alignment: Alignment::from_style(&s.alignment),
            protection: Protection {
                locked: s.locked,
                hidden: s.hidden,
            },
            is_date: is_date_format(&number_format),
            is_timedelta: is_timedelta_format(&number_format),
            number_format,
            named_style: s.named_style.clone().unwrap_or_else(|| "Normal".into()),
            quote_prefix: s.quote_prefix,
            pivot_button: s.pivot_button,
            has_style: index != 0,
        }
    }

    /// The format openpyxl gives cells it creates (an all-zero style array).
    pub(crate) fn zero(s: &CellStyle) -> Self {
        Self {
            has_style: false,
            ..Self::from_style(s, 0)
        }
    }

    /// openpyxl's defaults for a workbook without a stylesheet.
    pub(crate) fn openpyxl_default() -> Self {
        Self {
            font: Font::openpyxl_default(),
            fill: Fill::Pattern(PatternFill {
                pattern_type: None,
                fg: Color::black(),
                bg: Color::black(),
                fg_py: Lazy::default(),
                bg_py: Lazy::default(),
            }),
            border: Border::default(),
            alignment: Alignment::from_style(&StyleAlignment::default()),
            protection: Protection {
                locked: true,
                hidden: false,
            },
            number_format: "General".into(),
            is_date: false,
            is_timedelta: false,
            named_style: "Normal".into(),
            quote_prefix: false,
            pivot_button: false,
            has_style: false,
        }
    }

    pub(crate) fn protection(&self) -> &Protection {
        &self.protection
    }
}

/// Python objects of one format, built on first use and shared by its cells.
pub(crate) struct StylePy {
    pub(crate) font: Py<Font>,
    pub(crate) fill: Py<PyAny>,
    pub(crate) border: Py<Border>,
    pub(crate) alignment: Py<Alignment>,
    pub(crate) protection: Py<Protection>,
    pub(crate) number_format: Py<pyo3::types::PyString>,
    pub(crate) named_style: Py<pyo3::types::PyString>,
}

impl StylePy {
    pub(crate) fn new(py: Python<'_>, s: &CompatStyle) -> PyResult<Self> {
        Ok(Self {
            font: Py::new(py, s.font.clone())?,
            fill: s.fill.to_py(py)?,
            border: Py::new(py, s.border.clone())?,
            alignment: Py::new(py, s.alignment.clone())?,
            protection: Py::new(py, s.protection.clone())?,
            number_format: pyo3::types::PyString::new(py, &s.number_format).unbind(),
            named_style: pyo3::types::PyString::new(py, &s.named_style).unbind(),
        })
    }
}

pub(crate) fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Color>()?;
    m.add_class::<Font>()?;
    m.add_class::<PatternFill>()?;
    m.add_class::<GradientFill>()?;
    m.add_class::<Stop>()?;
    m.add_class::<Side>()?;
    m.add_class::<Border>()?;
    m.add_class::<Alignment>()?;
    m.add_class::<Protection>()?;
    Ok(())
}
