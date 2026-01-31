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

# Column letters to our keys (A=Name, B=Code, C=Destination, D=Escalator, E=TWD LINE, F=Stair, G=Lift)
COL_KEYS = ["Name", "Code", "Destination", "Escalator(s)", "TWD LINE", "Stair(s)", "Lift(s)"]


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
    """Parse a worksheet into list of dicts, one per row, keys = COL_KEYS."""
    with zipf.open(sheet_path) as f:
        root = ET.parse(f).getroot()
    rows = []
    for row_elem in root.findall(".//main:row", NS):
        row_idx = int(row_elem.get("r", 0))
        row_dict = {}
        for c in row_elem.findall("main:c", NS):
            ref = c.get("r")
            if not ref:
                continue
            col_idx = col_letter_to_index(ref)
            if col_idx is None or col_idx >= len(COL_KEYS):
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
        # Build full row list (0..6), empty string for missing cells
        row_list = [row_dict.get(i, "") for i in range(len(COL_KEYS))]
        rows.append(row_list)
    return rows


def gate_str_to_list(s):
    """Convert cell value like '7 18' or '13' to list of strings ['7','18'] or ['13']."""
    if not s or not str(s).strip():
        return []
    return [x.strip() for x in str(s).split() if x.strip()]


def sheet_rows_to_records(rows):
    """Convert parsed rows (list of 7-tuples) to list of dicts with carried Name/Code."""
    if not rows:
        return []
    # Skip header row
    data_rows = rows[1:] if rows[0][0] == "Name" else rows
    records = []
    last_name, last_code = "", ""
    for row in data_rows:
        name, code, dest, esc, twd, stair, lift = row
        if name:
            last_name = name
        if code:
            last_code = code
        # Skip fully empty rows
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
                rows = parse_sheet_rows(z, sheet_path, shared_strings)
            except KeyError:
                continue
            out["lines"][name] = sheet_rows_to_records(rows)
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
