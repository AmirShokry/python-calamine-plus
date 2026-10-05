"""Builds small .xlsx files from raw sheet XML, for cases openpyxl cannot write
(shared formulas with cached values, inline strings, ISO dates, rows without cells,
sheets without a <dimension>, ...)."""

from __future__ import annotations

import zipfile
from pathlib import Path

NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
REL_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

# Cell formats (cellXfs) by index: 0 General, 1 yyyy-mm-dd, 2 [h]:mm:ss, 3 h:mm,
# 4 0.00, 5 "custom text"@, 6 bold, 7 m/d/yy h:mm (built-in 22), 8 [Red]0.00
STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet {NS}>
<numFmts count="4">
<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
<numFmt numFmtId="165" formatCode="[h]:mm:ss"/>
<numFmt numFmtId="166" formatCode="&quot;days&quot; 0"/>
<numFmt numFmtId="167" formatCode="[Red]0.00"/>
</numFmts>
<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><b/><sz val="12"/><color rgb="FFFF0000"/><name val="Arial"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color indexed="64"/></left><right style="thick"/><top style="thin"/><bottom style="double"><color theme="4" tint="0.5"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="9">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="20" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1">
<alignment horizontal="center" wrapText="1" indent="2"/><protection locked="0"/></xf>
<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="167" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def sheet(rows_xml: str, dimension: str | None = "A1", extra: str = "") -> str:
    dim = f'<dimension ref="{dimension}"/>' if dimension else ""
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<worksheet {NS} {REL_NS}>{dim}<sheetData>{rows_xml}</sheetData>{extra}</worksheet>"
    )


def build(
    path: Path,
    sheets: dict[str, str],
    shared_strings: list[str] = (),  # type: ignore[assignment]
    date1904: bool = False,
    defined_names: str = "",
) -> Path:
    names = list(sheets)
    wb_sheets = "".join(
        f'<sheet name="{n}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
        for i, n in enumerate(names)
    )
    pr = '<workbookPr date1904="1"/>' if date1904 else "<workbookPr/>"
    dn = f"<definedNames>{defined_names}</definedNames>" if defined_names else ""
    workbook = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<workbook {NS} {REL_NS}>{pr}<sheets>{wb_sheets}</sheets>{dn}</workbook>"
    )
    rels = "".join(
        f'<Relationship Id="rId{i + 1}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        f'Target="worksheets/sheet{i + 1}.xml"/>'
        for i in range(len(names))
    )
    n = len(names)
    rels += (
        f'<Relationship Id="rId{n + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
        f'<Relationship Id="rId{n + 2}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
        'Target="sharedStrings.xml"/>'
    )
    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{rels}</Relationships>"
    )
    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(n)
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        f"{overrides}</Types>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/></Relationships>'
    )
    sst = "".join(f"<si><t>{s}</t></si>" for s in shared_strings)
    shared = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<sst {NS} count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">{sst}</sst>'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", STYLES)
        z.writestr("xl/sharedStrings.xml", shared)
        for i, name in enumerate(names):
            z.writestr(f"xl/worksheets/sheet{i + 1}.xml", sheets[name])
    return path
