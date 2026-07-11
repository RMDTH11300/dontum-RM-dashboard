from __future__ import annotations

import argparse
import json
import re
import shutil
import zipfile
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path
import xml.etree.ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

MONTHS_TH = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
FISCAL_MONTH_ORDER = [10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8, 9]


def col_to_num(col: str) -> int:
    n = 0
    for c in col:
        n = n * 26 + (ord(c.upper()) - 64)
    return n - 1


def excel_serial_to_iso(value):
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)) and 1 <= value <= 100000:
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(value))).date().isoformat()
        except Exception:
            return value
    return value


def read_xlsx_sheet(path: Path, sheet_index: int = 0):
    """Read cell values with Python standard library only.

    This is deliberately small and deterministic for static dashboard data conversion.
    It supports shared strings, inline strings, booleans, and numeric values.
    """
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("a:si", NS):
                shared.append("".join((t.text or "") for t in si.iter(f"{{{NS['a']}}}t")))

        wbroot = ET.fromstring(z.read("xl/workbook.xml"))
        sheets = [(s.attrib["name"], s.attrib[f"{{{NS['r']}}}id"]) for s in wbroot.find("a:sheets", NS)]
        relroot = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rels = {r.attrib["Id"]: r.attrib["Target"] for r in relroot}
        sheet_name, rid = sheets[sheet_index]
        target = rels[rid]
        if not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/")

        root = ET.fromstring(z.read(target))
        rows = []
        max_cols = 0
        for row in root.findall(".//a:sheetData/a:row", NS):
            vals = {}
            for c in row.findall("a:c", NS):
                ref = c.attrib.get("r", "")
                match = re.match(r"([A-Z]+)", ref)
                if not match:
                    continue
                idx = col_to_num(match.group(1))
                cell_type = c.attrib.get("t")
                v = c.find("a:v", NS)
                value = None if v is None else v.text
                if cell_type == "s" and value is not None:
                    value = shared[int(value)]
                elif cell_type == "inlineStr":
                    inline = c.find("a:is", NS)
                    value = "".join((t.text or "") for t in inline.iter(f"{{{NS['a']}}}t")) if inline is not None else ""
                elif cell_type == "b" and value is not None:
                    value = value == "1"
                elif value is not None:
                    try:
                        number = float(value)
                        value = int(number) if number.is_integer() else number
                    except ValueError:
                        pass
                vals[idx] = value
                max_cols = max(max_cols, idx + 1)
            rows.append(vals)

        matrix = []
        for vals in rows:
            arr = [None] * max_cols
            for idx, value in vals.items():
                arr[idx] = value
            matrix.append(arr)
        return sheet_name, matrix, [name for name, _ in sheets]


def clean_value(value):
    if isinstance(value, str):
        value = value.replace("\r\n", "\n").replace("\r", "\n").strip()
        return value if value else None
    return value


def risk_code(title) -> str:
    if title is None:
        return "ไม่ระบุ"
    text = str(title).strip()
    match = re.match(r"([A-Za-z]{2,5}\d{2,4})", text)
    return match.group(1).upper() if match else (text.split(":", 1)[0][:30] or "ไม่ระบุ")


def risk_type(title) -> str:
    code = risk_code(title)
    if code.startswith("C"):
        return "Clinical"
    if code.startswith("G"):
        return "Non-clinical"
    return "Other"


def severity_key(value) -> str:
    if value is None or value == "":
        return "ไม่ระบุ"
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip().upper()


def is_high_severity(value) -> bool:
    key = severity_key(value)
    if key in list("EFGHI"):
        return True
    try:
        return float(key) >= 3
    except ValueError:
        return False


def normalize_incident_file(path: Path, fiscal_year: int):
    _, rows, _ = read_xlsx_sheet(path, 0)
    if not rows:
        raise ValueError(f"ไม่พบข้อมูลใน {path.name}")
    headers = [str(v).strip() if v is not None else f"คอลัมน์_{i+1}" for i, v in enumerate(rows[0])]
    date_indices = {i for i, h in enumerate(headers) if h.startswith("วันที่")}
    records = []
    for raw in rows[1:]:
        raw = list(raw[: len(headers)]) + [None] * max(0, len(headers) - len(raw))
        row = []
        for i, value in enumerate(raw[: len(headers)]):
            value = clean_value(value)
            if i in date_indices:
                value = excel_serial_to_iso(value)
            row.append(value)
        if any(v not in (None, "") for v in row):
            records.append(row)
    return headers, records


def find_header_index(headers, name):
    try:
        return headers.index(name)
    except ValueError:
        for i, header in enumerate(headers):
            if name in header:
                return i
    raise KeyError(name)


def summarize_year(fiscal_year: int, headers, rows):
    idx_title = find_header_index(headers, "รหัส: เรื่องอุบัติการณ์")
    idx_severity = find_header_index(headers, "ความรุนแรง")
    idx_unit = find_header_index(headers, "หน่วยงานที่บันทึกรายงาน")
    idx_status = find_header_index(headers, "สถานะ")
    idx_date = find_header_index(headers, "วันที่เกิดอุบัติการณ์")

    type_counts = Counter()
    severity_counts = Counter()
    unit_counts = Counter()
    risk_counts = Counter()
    status_counts = Counter()
    month_counts = Counter()
    high_count = 0

    for row in rows:
        title = row[idx_title] if idx_title < len(row) else None
        sev = row[idx_severity] if idx_severity < len(row) else None
        unit = row[idx_unit] if idx_unit < len(row) else None
        status = row[idx_status] if idx_status < len(row) else None
        date_value = row[idx_date] if idx_date < len(row) else None
        type_counts[risk_type(title)] += 1
        severity_counts[severity_key(sev)] += 1
        unit_counts[str(unit or "ไม่ระบุ")] += 1
        risk_counts[risk_code(title)] += 1
        status_counts[str(status or "ไม่ระบุ")] += 1
        if is_high_severity(sev):
            high_count += 1
        if isinstance(date_value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", date_value):
            month_counts[int(date_value[5:7])] += 1

    monthly = [{"month": m, "label": MONTHS_TH[m-1], "count": month_counts[m]} for m in FISCAL_MONTH_ORDER]
    return {
        "fiscalYear": fiscal_year,
        "total": len(rows),
        "clinical": type_counts["Clinical"],
        "nonClinical": type_counts["Non-clinical"],
        "other": type_counts["Other"],
        "highSeverity": high_count,
        "uniqueUnits": len([k for k in unit_counts if k != "ไม่ระบุ"]),
        "severity": [{"label": k, "count": v} for k, v in sorted(severity_counts.items(), key=lambda kv: (-kv[1], kv[0]))],
        "monthly": monthly,
        "topRisks": [{"label": k, "count": v} for k, v in risk_counts.most_common(10)],
        "topUnits": [{"label": k, "count": v} for k, v in unit_counts.most_common(10)],
        "topStatuses": [{"label": k, "count": v} for k, v in status_counts.most_common(8)],
    }


def extract_profiles(path: Path, fiscal_year: int):
    _, rows, _ = read_xlsx_sheet(path, 0)
    result = []
    section = None
    for row in rows:
        first = clean_value(row[0]) if row else None
        if isinstance(first, str):
            if "Clinical risk" in first and "Non-clinical" not in first:
                section = "Clinical"
                continue
            if "Non-clinical risk" in first:
                section = "Non-clinical"
                continue
        if section and isinstance(first, (int, float)) and len(row) >= 7:
            result.append({
                "fiscalYear": fiscal_year,
                "type": section,
                "rank": int(first),
                "risk": clean_value(row[1]),
                "likelihood": clean_value(row[2]),
                "impact": clean_value(row[3]),
                "score": clean_value(row[4]),
                "level": clean_value(row[5]),
                "control": clean_value(row[6]),
            })
    return result


def extract_register(path: Path, fiscal_year: int):
    _, rows, _ = read_xlsx_sheet(path, 0)
    result = []
    for row in rows[2:]:
        row = list(row) + [None] * max(0, 12 - len(row))
        if not clean_value(row[0]):
            if result:
                break
            continue
        result.append({
            "fiscalYear": fiscal_year,
            "riskId": clean_value(row[0]),
            "source": clean_value(row[1]),
            "dateAdded": clean_value(row[2]),
            "title": clean_value(row[3]),
            "description": clean_value(row[4]),
            "quarter": clean_value(row[5]),
            "likelihood": clean_value(row[6]),
            "impact": clean_value(row[7]),
            "level": clean_value(row[8]),
            "prevention": clean_value(row[9]),
            "monitor": clean_value(row[10]),
            "mitigation": clean_value(row[11]),
        })
    return result


def write_json(path: Path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")


def discover_files(source: Path):
    all_xlsx = list(source.glob("*.xlsx"))
    incidents = {}
    profiles = {}
    registers = {}
    for p in all_xlsx:
        name = p.name
        m = re.match(r"(66|67|68|69)_.*\.xlsx$", name, re.I)
        if m:
            incidents[2500 + int(m.group(1))] = p
            continue
        if name.startswith("Risk_Profile"):
            m = re.search(r"(2566|2567|2568|2569)", name)
            if m:
                profiles[int(m.group(1))] = p
        elif name.startswith("Risk_Register"):
            m = re.search(r"(2566|2567|2568|2569)", name)
            if m:
                registers[int(m.group(1))] = p
    return incidents, profiles, registers


def build(source: Path, output: Path):
    incidents, profiles, registers = discover_files(source)
    years = sorted(set(incidents) & set(profiles) & set(registers))
    if not years:
        raise SystemExit("ไม่พบชุดไฟล์ครบถ้วน กรุณาตรวจชื่อไฟล์ปี 66-69, Risk_Profile และ Risk_Register")

    output.mkdir(parents=True, exist_ok=True)
    (output / "data").mkdir(exist_ok=True)
    headers_reference = None
    summaries = []
    profile_data = {}
    register_data = {}
    validation = []

    for year in years:
        headers, rows = normalize_incident_file(incidents[year], year)
        if headers_reference is None:
            headers_reference = headers
        elif headers != headers_reference:
            raise ValueError(f"หัวคอลัมน์ของปี {year} ไม่ตรงกับปีแรก")
        write_json(output / "data" / f"incidents_{year}.json", {
            "fiscalYear": year,
            "sourceFile": incidents[year].name,
            "rows": rows,
        })
        summary = summarize_year(year, headers, rows)
        summaries.append(summary)
        profile_data[str(year)] = extract_profiles(profiles[year], year)
        register_data[str(year)] = extract_register(registers[year], year)
        validation.append({
            "year": year,
            "incidentRows": len(rows),
            "profileRows": len(profile_data[str(year)]),
            "registerRows": len(register_data[str(year)]),
            "sourceIncident": incidents[year].name,
        })

    columns = {header: i for i, header in enumerate(headers_reference)}
    meta = {
        "hospital": "โรงพยาบาลดอนตูม",
        "systemName": "Dashboard ระบบบริหารความเสี่ยง",
        "years": years,
        "defaultYear": max(years),
        "generatedAt": datetime.now().isoformat(timespec="seconds"),
        "fiscalYearRule": "ปีงบประมาณถูกกำหนดจากชุดไฟล์ต้นทางของแต่ละปีโดยตรง ไม่ได้คำนวณจากปี พ.ศ. ของวันที่เกิดอุบัติการณ์",
        "fiscalPeriod": "1 ตุลาคม – 30 กันยายน",
        "headers": headers_reference,
        "columns": columns,
        "summaries": summaries,
        "validation": validation,
    }
    write_json(output / "data" / "meta.json", meta)
    write_json(output / "data" / "profiles.json", profile_data)
    write_json(output / "data" / "registers.json", register_data)
    return validation


def main():
    parser = argparse.ArgumentParser(description="สร้าง JSON สำหรับ Dashboard ความเสี่ยงโรงพยาบาลดอนตูม")
    parser.add_argument("--source", type=Path, default=Path("source"), help="โฟลเดอร์ไฟล์ Excel")
    parser.add_argument("--output", type=Path, default=Path("."), help="โฟลเดอร์เว็บไซต์")
    args = parser.parse_args()
    validation = build(args.source.resolve(), args.output.resolve())
    print(json.dumps(validation, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
