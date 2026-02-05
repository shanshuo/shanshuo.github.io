#!/usr/bin/env python3
"""
Convert mrt/Singapore MRT gates.xlsx to mrt/mrt-gates.json and embed into mrt/index.html.
Uses only Python stdlib (zipfile + xml). Run from repo root:
  python scripts/excel_to_json.py
"""

import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REPO_ROOT = Path(__file__).resolve().parent.parent
MRT_DIR = REPO_ROOT / "mrt"
XLSX_PATH = MRT_DIR / "Singapore MRT gates.xlsx"
OUT_PATH = MRT_DIR / "mrt-gates.json"
HTML_PATH = MRT_DIR / "index.html"
EMBED_PLACEHOLDER = "__EMBED_MRT_DATA__"

# Standard output keys; Excel columns vary by sheet (header row defines mapping)
OUTPUT_KEYS = ["Name", "Code", "Destination", "Escalator(s)", "TWD LINE", "Stair(s)", "Lift(s)"]


def col_letter_to_index(ref):
    """Convert cell ref (e.g. 'D2') to column index 0-6 for A-G."""
    match = re.match(r"^([A-G])\d+$", ref, re.IGNORECASE)
    if not match:
        return None
    return ord(match.group(1).upper()) - ord("A")


def load_shared_strings(zipf):
    """Load shared strings from xl/sharedStrings.xml."""
    with zipf.open("xl/sharedStrings.xml") as f:
        root = ET.parse(f).getroot()
    strings = []
    for si in root.findall("main:si", NS):
        t_el = si.find("main:t", NS)
        if t_el is not None and t_el.text is not None:
            strings.append(t_el.text)
        else:
            strings.append("".join(n.text or "" for n in si.iter() if n.text))
    return strings


def get_sheet_names(zipf):
    """Return list of sheet names in order (EW, NS, NE, CC, DT, TE)."""
    with zipf.open("xl/workbook.xml") as f:
        root = ET.parse(f).getroot()
    return [s.get("name") for s in root.findall(".//main:sheet", NS)]


def parse_sheet_rows(zipf, sheet_path, shared_strings):
    """
    Parse a worksheet: row 1 = header (column names), rest = data.
    Returns list of dicts, each keyed by header name (e.g. Name, Code, Lift(s)).
    Column order differs per sheet (EW has Name then Code; NS/NE/CC/DT/TE have Code then Name, etc.).
    """
    with zipf.open(sheet_path) as f:
        root = ET.parse(f).getroot()
    all_rows = []
    for row_elem in root.findall(".//main:row", NS):
        row_dict = {}
        for c in row_elem.findall("main:c", NS):
            ref = c.get("r")
            if not ref:
                continue
            col_idx = col_letter_to_index(ref)
            if col_idx is None or col_idx > 6:
                continue
            t = c.get("t")
            v_el = c.find("main:v", NS)
            v = v_el.text if v_el is not None and v_el.text else ""
            if t == "s":
                try:
                    row_dict[col_idx] = shared_strings[int(v)].strip()
                except (ValueError, IndexError):
                    row_dict[col_idx] = ""
            else:
                row_dict[col_idx] = v.strip() if v else ""
        # Build list of 7 values in order A..G
        row_list = [row_dict.get(i, "") for i in range(7)]
        all_rows.append(row_list)
    if not all_rows:
        return []
    header_row = all_rows[0]
    data_rows = all_rows[1:]
    # Build list of dicts keyed by header name (so column order per sheet is correct)
    result = []
    for row_list in data_rows:
        record = {}
        for col_idx, header_name in enumerate(header_row):
            if col_idx < len(row_list):
                record[header_name] = row_list[col_idx] or ""
        result.append(record)
    return result


def gate_str_to_list(s):
    """Convert cell value like '7 18' or '13' or 'EW: 13\\nTE: 19' to list of strings."""
    if not s or not str(s).strip():
        return []
    # Split by any whitespace (space, newline) and filter empty
    return [x.strip() for x in re.split(r"[\s\n]+", str(s)) if x.strip()]


def sheet_rows_to_records(row_dicts):
    """Convert list of row dicts (keyed by header name) to list of records with carried Name/Code."""
    records = []
    last_name, last_code = "", ""
    for row in row_dicts:
        name = (row.get("Name") or "").strip()
        code = (row.get("Code") or "").strip()
        dest = (row.get("Destination") or "").strip()
        esc = (row.get("Escalator(s)") or "").strip()
        twd = (row.get("TWD LINE") or "").strip()
        stair = (row.get("Stair(s)") or "").strip()
        lift = (row.get("Lift(s)") or "").strip()
        if name:
            last_name = name
        if code:
            last_code = code
        # Skip header-like rows (e.g. "Name" / "Code" in first columns)
        if last_name in ("Name", "Code") and last_code in ("Name", "Code"):
            continue
        if not dest and not esc and not twd and not stair and not lift:
            continue
        records.append({
            "stationName": last_name,
            "stationCode": last_code,
            "towards": dest,
            "escalators": gate_str_to_list(esc),
            "twdLine": gate_str_to_list(twd),
            "stairs": gate_str_to_list(stair),
            "lifts": gate_str_to_list(lift),
        })
    return records


def main():
    if not XLSX_PATH.exists():
        raise SystemExit(f"Excel file not found: {XLSX_PATH}")
    with zipfile.ZipFile(XLSX_PATH, "r") as z:
        shared_strings = load_shared_strings(z)
        sheet_names = get_sheet_names(z)
        out = {"lines": {}}
        for i, name in enumerate(sheet_names):
            # sheet1.xml = first sheet, sheet2.xml = second, etc.
            sheet_path = f"xl/worksheets/sheet{i + 1}.xml"
            try:
                row_dicts = parse_sheet_rows(z, sheet_path, shared_strings)
            except KeyError:
                continue
            out["lines"][name] = sheet_rows_to_records(row_dicts)
    json_str = json.dumps(out, ensure_ascii=False)
    OUT_PATH.write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Wrote {OUT_PATH}")

    # Embed JSON into HTML so opening mrt-gates.html via file:// works (fetch is blocked by CORS).
    if HTML_PATH.exists():
        html = HTML_PATH.read_text(encoding="utf-8")
        # Escape </script> in JSON so it does not close the script tag in HTML
        embed_str = json_str.replace("</script>", "</scr\" + \"ipt>")
        marker = "window.__MRT_GATES_DATA__ = "
        if EMBED_PLACEHOLDER in html:
            html = html.replace(EMBED_PLACEHOLDER, embed_str)
            HTML_PATH.write_text(html, encoding="utf-8")
            print(f"Embedded data into {HTML_PATH}")
        elif marker in html:
            # Replace previously embedded JSON (find balanced { } then replace)
            start = html.index(marker) + len(marker)
            depth = 0
            end = start
            for i, c in enumerate(html[start:], start):
                if c == "{":
                    depth += 1
                elif c == "}":
                    depth -= 1
                    if depth == 0:
                        end = i + 1
                        break
            html = html[:start] + embed_str + html[end:]
            HTML_PATH.write_text(html, encoding="utf-8")
            print(f"Updated embedded data in {HTML_PATH}")
        else:
            print(f"Note: embed marker not found in HTML; skip embed. Open via HTTP or run from repo root.")
    else:
        print(f"Note: {HTML_PATH} not found; skip embed.")


if __name__ == "__main__":
    main()
