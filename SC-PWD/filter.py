#!/usr/bin/env python3
"""Filter SC/PWD 10% discount invoices in the January 2026 folder and export them to Excel.

This version is intentionally limited to the January folder first, matching the user request.
It writes to a dedicated output directory to avoid permission issues when Excel or another app
keeps the target .xlsx file open.
"""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape

ROOT_DIR = Path(r"F:\PROJECTS\SC-PWD FILTRATION\SC-PWD")
YEAR_DIR = ROOT_DIR / "2026"
MONTHS = [
    "JANUARY",
    "FEBRUARY",
    "MARCH",
    "APRIL",
    "MAY",
    "JUNE",
    "JULY",
    "AUGUST",
]
OUTPUT_DIR = ROOT_DIR / "output"

DISCOUNT_RE = re.compile(r"Less:\s*(SC|PWD)\s+10%\s+PROMO\s+([0-9,]+\.?\d{0,2})", re.IGNORECASE)
TOTAL_DUE_RE = re.compile(r"TOTAL DUE\s+([0-9,]+\.?\d{0,2})", re.IGNORECASE)
DATE_RE = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")
PLACEHOLDER_RE = re.compile(r"^(?:[_.\s]+|\.{1,})$")
ID_RE = re.compile(r"\d{3,}")


def normalize_text(value: str | None) -> str:
    if value is None:
        return ""
    text = value.strip()
    if not text:
        return ""
    text = text.replace("Cust Name:", "").replace("Address:", "")
    text = text.replace("_", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip(" .")


def is_placeholder(value: str) -> bool:
    if not value:
        return True
    return bool(PLACEHOLDER_RE.fullmatch(value.strip()))


def extract_date(lines: list[str], index: int) -> str:
    for i in range(index, -1, -1):
        match = DATE_RE.search(lines[i])
        if match:
            return match.group(0)
    return ""


def extract_customer_name(lines: list[str], start_index: int) -> str:
    for i in range(start_index, min(start_index + 150, len(lines))):
        line = lines[i].strip()
        if not line.startswith("Cust Name:"):
            continue
        value = normalize_text(line.split(":", 1)[1] if ":" in line else line)
        if value and not is_placeholder(value):
            return value
    return ""


def extract_address(lines: list[str], start_index: int) -> str:
    for i in range(start_index, min(start_index + 150, len(lines))):
        line = lines[i].strip()
        if not line.startswith("Address:"):
            continue
        value = normalize_text(line.split(":", 1)[1] if ":" in line else line)
        if value and not is_placeholder(value):
            return value
    return ""


def extract_id_number(name: str, address: str) -> str:
    for value in (name, address):
        if not value:
            continue
        matches = ID_RE.findall(value)
        if matches:
            return matches[-1]
    return ""


def parse_month_records(file_path: Path) -> list[dict[str, str]]:
    text = file_path.read_text(encoding="utf-8", errors="ignore")
    lines = text.splitlines()
    records: list[dict[str, str]] = []

    for idx, line in enumerate(lines):
        match = DISCOUNT_RE.search(line)
        if not match:
            continue

        amount = match.group(2).replace(",", "")
        date = extract_date(lines, idx)
        name = extract_customer_name(lines, idx)
        address = extract_address(lines, idx)
        if not name:
            continue

        customer_id = extract_id_number(name, address)
        if not customer_id:
            continue

        total_due = ""
        for j in range(idx, min(len(lines), idx + 80)):
            total_match = TOTAL_DUE_RE.search(lines[j])
            if total_match:
                total_due = total_match.group(1).replace(",", "")
                break

        records.append(
            {
                "Date": date,
                "Customer name": name,
                "ID Number": customer_id,
                "Amount Discount": amount,
                "Total Amount Due": total_due,
            }
        )

    return records


def col_letter(column_index: int) -> str:
    letters = ""
    while column_index:
        column_index, remainder = divmod(column_index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def make_cell(row_num: int, col_num: int, value: object, style_id: int = 0) -> str:
    ref = f"{col_letter(col_num)}{row_num}"
    if isinstance(value, (int, float)):
        return f'<c r="{ref}" s="{style_id}"><v>{value}</v></c>'
    text = str(value)
    return f'<c r="{ref}" s="{style_id}" t="inlineStr"><is><t>{escape(text)}</t></is></c>'


def choose_output_path(base_path: Path) -> Path:
    base_path.parent.mkdir(parents=True, exist_ok=True)
    if not base_path.exists():
        return base_path

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return base_path.with_name(f"{base_path.stem}_{timestamp}{base_path.suffix}")


def write_xlsx_manual(rows: list[list[object]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    worksheet_rows = []
    for row_num, row in enumerate(rows, start=1):
        style_id = 3
        if row_num == 1:
            style_id = 1
        elif row_num == 2:
            style_id = 2

        cells = "".join(
            make_cell(row_num, col_idx, value, style_id=style_id)
            for col_idx, value in enumerate(row, start=1)
        )
        worksheet_rows.append(f"<row r=\"{row_num}\">{cells}</row>")

    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <mergeCells count="1">
    <mergeCell ref="A1:E1"/>
  </mergeCells>
  <sheetData>
    {''.join(worksheet_rows)}
  </sheetData>
</worksheet>
'''

    workbook_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="SC_PWD" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
'''

    workbook_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'''

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
'''

    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'''

    core_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                  xmlns:dc="http://purl.org/dc/elements/1.1/"
                  xmlns:dcterms="http://purl.org/dc/terms/"
                  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Copilot</dc:creator>
  <cp:lastModifiedBy>Copilot</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-09-17T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-09-17T00:00:00Z</dcterms:modified>
</cp:coreProperties>
'''

    app_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
</Properties>
'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="0"/>
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="12"/><name val="Calibri"/><family val="2"/><color rgb="FFFFFFFF"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="0" applyFill="0" applyBorder="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
'''

    with ZipFile(output_path, "w", compression=ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("docProps/core.xml", core_xml)
        zf.writestr("docProps/app.xml", app_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/styles.xml", styles_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)

    return None


def export_records(output_path: Path, records: list[dict[str, str]], month_name: str) -> int:
    rows: list[list[object]] = [
        [f"{month_name} 2026 SC/PWD 10% Discount Records"] + [""] * 4,
        ["Date", "Customer name", "ID Number", "Amount Discount", "Total Amount Due"],
    ]
    for record in records:
        rows.append([
            record["Date"],
            record["Customer name"],
            record["ID Number"],
            record["Amount Discount"],
            record["Total Amount Due"],
        ])

    write_xlsx_manual(rows, output_path)
    return len(records)


def main() -> None:
    output_dir = OUTPUT_DIR
    output_dir.mkdir(parents=True, exist_ok=True)

    for month_name in MONTHS:
        month_dir = YEAR_DIR / month_name
        if not month_dir.exists():
            continue

        records: list[dict[str, str]] = []
        for file_path in sorted(month_dir.glob("*.txt")):
            records.extend(parse_month_records(file_path))

        if not records:
            print(f"No SC/PWD 10% records were found in {month_dir}; skipping.")
            continue

        output_path = choose_output_path(output_dir / f"{month_name} 2026 SC & PWD RECORDS.xlsx")
        count = export_records(output_path, records, month_name)
        print(f"Created: {output_path}")
        print(f"Total filtered records for {month_name}: {count}")


if __name__ == "__main__":
    main()
