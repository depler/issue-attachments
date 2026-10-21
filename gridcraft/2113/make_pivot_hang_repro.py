"""Build a synthetic XLSX that reproduces the "PivotTable Fields" hang in GridCraft.

Fully synthetic: a support-ticket log and one small PivotTable over it. The point is the
shape - a long, wide source range plus a PivotTable whose report holds the active cell -
which makes `pivot.fields` re-scan the whole source for every UI frame.

The package mirrors what GridCraft's own writer emits (and what Excel accepts): axis
pivotFields carry their item lists, the cache carries per-field sharedItems, and the cache
is saved with saveData="0" refreshOnLoad="1", so a reader rebuilds the report from the
source list - exactly like the workbook from the bug report.

    python make_pivot_hang_repro.py [rows] [out.xlsx]
"""

import hashlib
import io
import random
import sys
import zipfile
from datetime import date, timedelta

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "application/vnd.openxmlformats-officedocument.spreadsheetml"

ROWS = int(sys.argv[1]) if len(sys.argv) > 1 else 250_000
OUT = sys.argv[2] if len(sys.argv) > 2 else "pivot-fields-hang.xlsx"

EPOCH = date(1899, 12, 30)

# name, kind, values (for "str" fields)
COLS = [
    ("TICKET_ID", "int", None),
    ("OPENED_ON", "date", None),
    ("REGION", "str", ["Central", "East", "North", "South", "West"]),
    ("QUEUE", "str", ["Billing", "Integrations", "Mobile", "Onboarding", "Security"]),
    ("PRIORITY", "str", ["P1", "P2", "P3", "P4"]),
    ("TEAM", "str", ["Cygnus", "Draco", "Falcon", "Lyra", "Orion", "Vega"]),
    ("AGENT", "str", [f"Agent {i:02d}" for i in range(1, 25)]),
    ("STATUS", "str", ["Closed", "Open", "Pending", "Resolved"]),
    ("CHANNEL", "str", ["Chat", "Email", "Phone", "Portal"]),
    ("ESCALATED", "bool", None),
    ("REOPENED", "bool", None),
    ("FIRST_REPLY_MIN", "int", None),
    ("RESOLVE_MIN", "int", None),
    ("CSAT", "float", None),
    ("BACKLOG_DAYS", "int", None),
    ("SLA_BREACH", "bool", None),
    ("ATTACHMENTS", "int", None),
    ("UPDATES", "int", None),
    ("LAST_UPDATE", "date", None),
    ("OWNER", "str", ["Desk A", "Desk B", "Desk C"]),
]
IDX = {c[0]: i for i, c in enumerate(COLS)}
REGION, PRIORITY = IDX["REGION"], IDX["PRIORITY"]
RESOLVE_MIN, CSAT = IDX["RESOLVE_MIN"], IDX["CSAT"]

COL_LETTERS = []
for _i in range(len(COLS)):
    _n, _s = _i + 1, ""
    while _n:
        _n, _r = divmod(_n - 1, 26)
        _s = chr(65 + _r) + _s
    COL_LETTERS.append(_s)

strings: dict[str, int] = {}
order: list[str] = []


def sid(text: str) -> int:
    if text not in strings:
        strings[text] = len(order)
        order.append(text)
    return strings[text]


def esc(t: str) -> str:
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def num(v: float) -> str:
    return f"{v:.0f}" if float(v).is_integer() else repr(float(v))


def main() -> None:
    rnd = random.Random(20261021)
    base = date(2024, 1, 1)

    # per column: [min, max, any non-integer]
    stats: list[list] = [[float("inf"), float("-inf"), False] for _ in COLS]

    header = "".join(f'<c r="{COL_LETTERS[i]}1" t="s"><v>{sid(name)}</v></c>' for i, (name, _k, _v) in enumerate(COLS))
    rows = io.StringIO()
    rows.write(f'<row r="1" spans="1:{len(COLS)}">{header}</row>')

    for r in range(2, ROWS + 2):
        cells = []
        opened = (base + timedelta(days=rnd.randrange(0, 900))).toordinal() - EPOCH.toordinal()
        for i, (name, kind, values) in enumerate(COLS):
            ref = f"{COL_LETTERS[i]}{r}"
            if kind == "str":
                cells.append(f'<c r="{ref}" t="s"><v>{sid(rnd.choice(values))}</v></c>')
                continue
            if kind == "date":
                v = opened
                cells.append(f'<c r="{ref}" s="1"><v>{v}</v></c>')
            elif kind == "bool":
                v = rnd.randrange(0, 2)
                cells.append(f'<c r="{ref}"><v>{v}</v></c>')
            elif kind == "int":
                v = 900_000 + r if name == "TICKET_ID" else (rnd.randrange(5, 5000) if name == "RESOLVE_MIN" else rnd.randrange(0, 200))
                cells.append(f'<c r="{ref}"><v>{v}</v></c>')
            else:  # float
                v = rnd.randrange(10, 51) / 10
                cells.append(f'<c r="{ref}"><v>{v:.1f}</v></c>')
            st = stats[i]
            st[0] = min(st[0], v)
            st[1] = max(st[1], v)
            if float(v) != int(v):
                st[2] = True
        rows.write(f'<row r="{r}" spans="1:{len(COLS)}">{"".join(cells)}</row>')

    last_row = ROWS + 1
    last_col = COL_LETTERS[-1]

    def sheet(body: str, dimension: str, view: str = "") -> str:
        return (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<worksheet xmlns="{NS}" xmlns:r="{R_NS}"><dimension ref="{dimension}"/>{view}'
            f"<sheetData>{body}</sheetData></worksheet>"
        )

    data_sheet = sheet(rows.getvalue(), f"A1:{last_col}{last_row}")

    # Report sheet: captions around the report area, no computed cells - a workbook saved with
    # "save source data with file" off, which Excel rebuilds on open.
    captions = [("A1", "Support tickets - pivot test"), ("A3", "Row Labels"), ("B3", "Priority")]
    by_row: dict[int, list[str]] = {}
    for ref, text in captions:
        by_row.setdefault(int(ref[1:]), []).append(f'<c r="{ref}" t="s"><v>{sid(text)}</v></c>')
    body = "".join(f'<row r="{r}">{"".join(c)}</row>' for r, c in sorted(by_row.items()))
    view = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="C6" sqref="C6"/></sheetView></sheetViews>'
    report_sheet = sheet(body, "A1:M20", view)

    workbook = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{NS}" xmlns:r="{R_NS}"><workbookPr/>'
        f'<bookViews><workbookView activeTab="1"/></bookViews>'
        f'<sheets><sheet name="Tickets" sheetId="1" r:id="rId1"/><sheet name="Report" sheetId="2" r:id="rId2"/></sheets>'
        f'<calcPr calcId="191029"/>'
        f'<pivotCaches><pivotCache cacheId="1" r:id="rId5"/></pivotCaches>'
        f"</workbook>"
    )

    wb_rels = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PKG}">'
        f'<Relationship Id="rId1" Type="{R_NS}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{R_NS}/worksheet" Target="worksheets/sheet2.xml"/>'
        f'<Relationship Id="rId3" Type="{R_NS}/styles" Target="styles.xml"/>'
        f'<Relationship Id="rId4" Type="{R_NS}/sharedStrings" Target="sharedStrings.xml"/>'
        f'<Relationship Id="rId5" Type="{R_NS}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>'
        f"</Relationships>"
    )

    root_rels = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PKG}">'
        f'<Relationship Id="rId1" Type="{R_NS}/officeDocument" Target="xl/workbook.xml"/>'
        f"</Relationships>"
    )

    sheet2_rels = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PKG}">'
        f'<Relationship Id="rId1" Type="{R_NS}/pivotTable" Target="../pivotTables/pivotTable1.xml"/>'
        f"</Relationships>"
    )

    table_rels = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PKG}">'
        f'<Relationship Id="rId1" Type="{R_NS}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/>'
        f"</Relationships>"
    )

    # Cache fields, with the sharedItems Excel expects (and a reader uses for its item labels).
    cache_fields = []
    for i, (name, kind, values) in enumerate(COLS):
        if kind == "str":
            items = "".join(f'<s v="{esc(v)}"/>' for v in sorted(values))
            cache_fields.append(f'<cacheField name="{name}" numFmtId="0"><sharedItems count="{len(values)}">{items}</sharedItems></cacheField>')
        else:
            lo, hi, frac = stats[i]
            integer = "" if frac else ' containsInteger="1"'
            cache_fields.append(
                f'<cacheField name="{name}" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0"'
                f' containsNumber="1"{integer} minValue="{num(lo)}" maxValue="{num(hi)}"/></cacheField>'
            )

    cache_def = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<pivotCacheDefinition xmlns="{NS}" xmlns:r="{R_NS}" saveData="0" refreshOnLoad="1" refreshedBy="GridCraft"'
        f' createdVersion="6" refreshedVersion="6" minRefreshableVersion="3">'
        f'<cacheSource type="worksheet"><worksheetSource ref="A1:{last_col}{last_row}" sheet="Tickets"/></cacheSource>'
        f'<cacheFields count="{len(COLS)}">{"".join(cache_fields)}</cacheFields>'
        f"</pivotCacheDefinition>"
    )

    # Axis fields carry their item list and the default item, as Excel and GridCraft write them.
    pivot_fields = []
    for i in range(len(COLS)):
        if i in (REGION, PRIORITY):
            axis = "axisRow" if i == REGION else "axisCol"
            names = sorted(COLS[i][2])
            items = "".join(f'<item x="{k}"/>' for k in range(len(names))) + '<item t="default"/>'
            pivot_fields.append(
                f'<pivotField axis="{axis}" showAll="0" sortType="ascending">'
                f'<items count="{len(names) + 1}">{items}</items></pivotField>'
            )
        elif i in (RESOLVE_MIN, CSAT):
            pivot_fields.append('<pivotField dataField="1" showAll="0"/>')
        else:
            pivot_fields.append('<pivotField showAll="0"/>')

    pivot = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<pivotTableDefinition xmlns="{NS}" name="TicketPivot" cacheId="1" applyNumberFormats="0"'
        f' applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0"'
        f' applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3"'
        f' useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1"'
        f' multipleFieldFilters="0">'
        f'<location ref="A3:M20" firstHeaderRow="1" firstDataRow="3" firstDataCol="1"/>'
        f'<pivotFields count="{len(COLS)}">{"".join(pivot_fields)}</pivotFields>'
        f'<rowFields count="1"><field x="{REGION}"/></rowFields>'
        f'<colFields count="2"><field x="{PRIORITY}"/><field x="-2"/></colFields>'
        f'<dataFields count="2">'
        f'<dataField name="Sum of RESOLVE_MIN" fld="{RESOLVE_MIN}" baseField="0" baseItem="0"/>'
        f'<dataField name="Average of CSAT" fld="{CSAT}" subtotal="average" baseField="0" baseItem="0"/>'
        f"</dataFields>"
        f'<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0"'
        f' showColStripes="0" showLastColumn="1"/>'
        f"</pivotTableDefinition>"
    )

    styles = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="{NS}">'
        f'<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>'
        f'<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        f'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        f'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        f'<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'
        f'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        f"</styleSheet>"
    )

    shared = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<sst xmlns="{NS}" count="{len(order)}" uniqueCount="{len(order)}">'
        + "".join(f"<si><t>{esc(t)}</t></si>" for t in order)
        + "</sst>"
    )

    content_types = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        f'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        f'<Default Extension="xml" ContentType="application/xml"/>'
        f'<Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/>'
        f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="{CT}.worksheet+xml"/>'
        f'<Override PartName="/xl/worksheets/sheet2.xml" ContentType="{CT}.worksheet+xml"/>'
        f'<Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/>'
        f'<Override PartName="/xl/sharedStrings.xml" ContentType="{CT}.sharedStrings+xml"/>'
        f'<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="{CT}.pivotTable+xml"/>'
        f'<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="{CT}.pivotCacheDefinition+xml"/>'
        f"</Types>"
    )

    parts = {
        "[Content_Types].xml": content_types,
        "_rels/.rels": root_rels,
        "xl/workbook.xml": workbook,
        "xl/_rels/workbook.xml.rels": wb_rels,
        "xl/styles.xml": styles,
        "xl/sharedStrings.xml": shared,
        "xl/worksheets/sheet1.xml": data_sheet,
        "xl/worksheets/sheet2.xml": report_sheet,
        "xl/worksheets/_rels/sheet2.xml.rels": sheet2_rels,
        "xl/pivotTables/pivotTable1.xml": pivot,
        "xl/pivotTables/_rels/pivotTable1.xml.rels": table_rels,
        "xl/pivotCache/pivotCacheDefinition1.xml": cache_def,
    }

    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for name, text in parts.items():
            z.writestr(name, text.encode("utf-8"))

    digest = hashlib.sha256(open(OUT, "rb").read()).hexdigest()
    print(f"{OUT}: {ROWS:,} data rows x {len(COLS)} columns, sha256 {digest[:16]}")


if __name__ == "__main__":
    main()
