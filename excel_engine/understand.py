"""Streaming workbook analysis: what each sheet is, how columns are computed, and how sheets feed each other.

Reads the xlsx zip directly (lxml iterparse per sheet) so 300 MB+ unpacked planning workbooks are analysed in
seconds without a full openpyxl load. The resulting "workbook map" drives the LLM fill plan and the in-place patch.
"""
from __future__ import annotations

import io
import posixpath
import re
import zipfile
from datetime import date, timedelta

from lxml import etree

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"

PROFILE_ROWS = 5000          # per-column type profile uses the first N data rows
KEY_CAP = 20000              # item keys kept per sheet
DISTINCT_CAP = 40            # low-cardinality text columns keep their distinct values
HEADER_SCAN = 40             # header row is searched in the first N rows

ITEM_KEY = re.compile(r"^[A-Z]{1,5}-?\d{3,7}[A-Z]?$")
_MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}
_MONTH_RE = re.compile(
    r"^(?:sum\s+of\s+|total\s+|avg\s+of\s+)?(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?"
    r"(?:[\s\-/'.]*(\d{2,4}))?$", re.I)
_SHEET_REF = re.compile(r"'((?:[^']|'')+)'!|(?<![\w.'\]])([A-Za-z_][\w.]*)!")
_TABLE_REF = re.compile(r"\b([A-Za-z_][\w.]*)\[")
_BUILTIN_DATE_FMTS = set(range(14, 23)) | {45, 46, 47}


def col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def split_ref(ref: str):
    m = re.match(r"^\$?([A-Z]+)\$?(\d+)$", ref or "")
    return (m.group(1), int(m.group(2))) if m else (None, None)


def excel_date(serial) -> date | None:
    try:
        v = float(serial)
    except (TypeError, ValueError):
        return None
    if not 20000 < v < 80000:
        return None
    return date(1899, 12, 30) + timedelta(days=int(v))


def month_of_header(text, serial_date: date | None = None):
    """'Sum of Jul' -> (7, None); 'Jul-25' -> (7, 2025); 'Sept-26' -> (9, 2026); date serial -> (m, y)."""
    if serial_date:
        return serial_date.month, serial_date.year
    t = str(text or "").strip()
    m = _MONTH_RE.match(t)
    if not m:
        return None
    month = _MONTHS[m.group(1).lower()[:3]]
    year = m.group(2)
    if year:
        y = int(year)
        year = y + 2000 if y < 100 else y
    return month, year or None


class Package:
    """Resolved parts of an xlsx package: sheets, rels, shared strings, styles, tables, pivots."""

    def __init__(self, data: bytes):
        self.data = data
        self.zf = zipfile.ZipFile(io.BytesIO(data))
        self.names = set(self.zf.namelist())
        self.workbook_xml = self.zf.read("xl/workbook.xml")
        self.sheets = self._sheets()
        self.shared = self._shared_strings()
        self.date_styles = self._date_styles()

    def read(self, name):
        return self.zf.read(name)

    def rels(self, part):
        base = posixpath.dirname(part)
        rel_path = posixpath.join(base, "_rels", posixpath.basename(part) + ".rels")
        if rel_path not in self.names:
            return {}
        out = {}
        root = etree.fromstring(self.zf.read(rel_path))
        for el in root.iter(PKG + "Relationship"):
            target = el.get("Target") or ""
            if el.get("TargetMode") == "External":
                full = target
            elif target.startswith("/"):
                full = target.lstrip("/")
            else:
                full = posixpath.normpath(posixpath.join(base, target))
            out[el.get("Id")] = {"type": (el.get("Type") or "").rsplit("/", 1)[-1], "target": full}
        return out

    def _sheets(self):
        root = etree.fromstring(self.workbook_xml)
        rels = self.rels("xl/workbook.xml")
        out = []
        for i, el in enumerate(root.iter(NS + "sheet")):
            rid = el.get(RNS + "id")
            part = rels.get(rid, {}).get("target")
            out.append({"index": i, "name": el.get("name"), "sheet_id": el.get("sheetId"),
                        "state": el.get("state") or "visible", "rid": rid, "part": part})
        return out

    def _shared_strings(self):
        rels = self.rels("xl/workbook.xml")
        part = next((r["target"] for r in rels.values() if r["type"] == "sharedStrings"), None)
        if not part or part not in self.names:
            return []
        out = []
        with self.zf.open(part) as f:
            for _ev, el in etree.iterparse(f, events=("end",), tag=NS + "si"):
                out.append("".join(t.text or "" for t in el.iter(NS + "t")))
                el.clear()
        return out

    def _date_styles(self):
        if "xl/styles.xml" not in self.names:
            return set()
        root = etree.fromstring(self.zf.read("xl/styles.xml"))
        custom = {}
        for nf in root.iter(NS + "numFmt"):
            code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", nf.get("formatCode") or "").lower()
            custom[int(nf.get("numFmtId"))] = bool(re.search(r"[dmy]", code)) and "0" not in code[:1]
        xfs = root.find(NS + "cellXfs")
        out = set()
        if xfs is not None:
            for i, xf in enumerate(xfs.findall(NS + "xf")):
                fid = int(xf.get("numFmtId") or 0)
                if fid in _BUILTIN_DATE_FMTS or custom.get(fid):
                    out.add(str(i))
        return out

    def tables_for(self, sheet_part):
        out = []
        for rel in self.rels(sheet_part).values():
            if rel["type"] != "table" or rel["target"] not in self.names:
                continue
            root = etree.fromstring(self.zf.read(rel["target"]))
            cols = []
            for tc in root.iter(NS + "tableColumn"):
                calc = tc.find(NS + "calculatedColumnFormula")
                cols.append({"name": tc.get("name"), "calc": calc.text if calc is not None else None})
            out.append({"part": rel["target"], "name": root.get("displayName") or root.get("name"),
                        "ref": root.get("ref"), "columns": cols})
        return out

    def pivots_for(self, sheet_part):
        out = []
        caches = self.pivot_caches()
        for rel in self.rels(sheet_part).values():
            if rel["type"] != "pivotTable" or rel["target"] not in self.names:
                continue
            root = etree.fromstring(self.zf.read(rel["target"]))
            loc = root.find(NS + "location")
            cache = None
            for prel in self.rels(rel["target"]).values():
                if prel["type"] == "pivotCacheDefinition":
                    cache = caches.get(prel["target"])
            fields = (cache or {}).get("fields") or []

            def names(tag):
                parent = root.find(NS + tag)
                if parent is None:
                    return []
                idx = [int(f.get("x") or f.get("fld") or -1) for f in parent]
                return [fields[i] for i in idx if 0 <= i < len(fields)]

            out.append({"part": rel["target"], "name": root.get("name"),
                        "location": loc.get("ref") if loc is not None else None,
                        "cache": (cache or {}).get("part"),
                        "source_sheet": (cache or {}).get("source_sheet"),
                        "source_ref": (cache or {}).get("source_ref"),
                        "rows": names("rowFields"), "cols": names("colFields"), "pages": names("pageFields"),
                        "data": [d.get("name") for d in root.iter(NS + "dataField")]})
        return out

    def pivot_caches(self):
        if hasattr(self, "_caches"):
            return self._caches
        out = {}
        for name in self.names:
            if not re.match(r"xl/pivotCache/pivotCacheDefinition\d+\.xml$", name):
                continue
            head = self.zf.read(name)
            root = etree.fromstring(head)
            src = root.find(f"{NS}cacheSource/{NS}worksheetSource")
            out[name] = {"part": name,
                         "source_sheet": src.get("sheet") if src is not None else None,
                         "source_ref": src.get("ref") if src is not None else None,
                         "source_name": src.get("name") if src is not None else None,
                         "fields": [f.get("name") for f in root.iter(NS + "cacheField")]}
        self._caches = out
        return out


def _cell_value(c, pkg):
    """Return (kind, value, formula) for a <c> element."""
    t = c.get("t")
    fe = c.find(NS + "f")
    formula = None
    if fe is not None:
        formula = fe.text or ""
    v = c.find(NS + "v")
    raw = v.text if v is not None else None
    if t == "s" and raw is not None:
        try:
            return "text", pkg.shared[int(raw)], formula
        except (IndexError, ValueError):
            return "text", "", formula
    if t == "inlineStr":
        return "text", "".join(x.text or "" for x in c.iter(NS + "t")), formula
    if t in ("str", "e"):
        return "text", raw or "", formula
    if t == "b":
        return "bool", raw == "1", formula
    if raw is None:
        return "blank", None, formula
    try:
        num = float(raw)
    except ValueError:
        return "text", raw, formula
    if c.get("s") in pkg.date_styles:
        return "date", num, formula
    return "number", num, formula


def _refs(formula: str):
    sheets = set()
    for a, b in _SHEET_REF.findall(formula or ""):
        name = (a or b).replace("''", "'")
        if name:
            sheets.add(name)
    tables = set(_TABLE_REF.findall(formula or ""))
    return sheets, tables


def _in_range(col, row, ranges):
    return any(c1 <= col <= c2 and r1 <= row <= r2 for c1, r1, c2, r2 in ranges)


def _parse_range(ref):
    if not ref:
        return None
    parts = ref.replace("$", "").split(":")
    a = re.match(r"([A-Z]+)(\d+)", parts[0])
    b = re.match(r"([A-Z]+)(\d+)", parts[-1])
    if not a or not b:
        return None
    return col_index(a.group(1)), int(a.group(2)), col_index(b.group(1)), int(b.group(2))


def analyse_sheet(pkg: Package, sheet: dict, max_rows: int | None = None) -> dict:
    part = sheet["part"]
    tables = pkg.tables_for(part) if part else []
    pivots = pkg.pivots_for(part) if part else []
    pivot_ranges = [r for r in (_parse_range(p["location"]) for p in pivots) if r]
    info = {"name": sheet["name"], "index": sheet["index"], "state": sheet["state"], "part": part,
            "tables": [{"name": t["name"], "ref": t["ref"], "columns": [c["name"] for c in t["columns"]],
                        "calculated": {c["name"]: c["calc"] for c in t["columns"] if c["calc"]}} for t in tables],
            "pivots": pivots, "rows": 0, "last_row": 0, "header_row": None, "columns": [],
            "key_col": None, "key_values": [], "key_prefixes": [], "depends_on": [], "month_cols": []}
    if not part or part not in pkg.names:
        return info

    head_rows = {}
    profile = {}
    shared_master = {}
    last_row = 0
    rows = 0
    header_row = None
    table_header = None
    if tables:
        rng = _parse_range(tables[0]["ref"])
        if rng:
            table_header = rng[1]
    deps, table_deps = set(), set()

    with pkg.zf.open(part) as f:
        for _ev, row in etree.iterparse(f, events=("end",), tag=NS + "row"):
            r = int(row.get("r") or (last_row + 1))
            rows += 1
            last_row = r
            cells = row.findall(NS + "c")
            if r <= HEADER_SCAN:
                head_rows[r] = []
            if header_row is None and r > HEADER_SCAN:
                header_row = _pick_header(head_rows, table_header, pkg)
            for c in cells:
                letter, _ = split_ref(c.get("r") or "")
                if not letter:
                    continue
                ci = col_index(letter)
                kind, value, formula = _cell_value(c, pkg)
                fe = c.find(NS + "f")
                if fe is not None and fe.get("t") == "shared":
                    si = fe.get("si")
                    if formula:
                        shared_master[si] = formula
                    else:
                        formula = shared_master.get(si, "")
                if r <= HEADER_SCAN:
                    head_rows[r].append((ci, kind, value, formula, c.get("s")))
                if formula and (header_row is None or r - header_row <= PROFILE_ROWS):
                    s, t = _refs(formula)
                    deps |= s
                    table_deps |= t
                if header_row is not None and r > header_row and not _in_range(ci, r, pivot_ranges):
                    _profile(profile, ci, kind, value, formula, r - header_row, c.get("s"))
            row.clear()
            while row.getprevious() is not None:
                del row.getparent()[0]
            if max_rows and rows >= max_rows:
                info["truncated"] = True
                break
    if header_row is None:
        header_row = _pick_header(head_rows, table_header, pkg)
        # profile rows that were inside the header scan window
        if header_row:
            for r, cells in head_rows.items():
                if r > header_row:
                    for ci, kind, value, formula, style in cells:
                        _profile(profile, ci, kind, value, formula, r - header_row, style)
    elif head_rows:
        # rows after the header but inside the scan window were seen before header_row was known
        for r, cells in head_rows.items():
            if r > header_row:
                for ci, kind, value, formula, style in cells:
                    if not _in_range(ci, r, pivot_ranges):
                        _profile(profile, ci, kind, value, formula, r - header_row, style)

    info["rows"] = rows
    info["last_row"] = last_row
    info["header_row"] = header_row
    headers = {}
    if header_row and header_row in head_rows:
        for ci, kind, value, formula, _s in head_rows[header_row]:
            if _in_range(ci, header_row, pivot_ranges):
                continue
            label = None
            if kind == "text" and str(value).strip() and not formula:
                label = str(value).strip()
            elif kind in ("number", "date") and excel_date(value):
                label = excel_date(value).isoformat()
            if label:
                headers[ci] = (label, kind == "date" or (kind == "number" and bool(excel_date(value))))
    if tables:
        rng = _parse_range(tables[0]["ref"])
        if rng and rng[1] == header_row:
            for i, col in enumerate(tables[0]["columns"]):
                headers[rng[0] + i] = (col["name"] if isinstance(col, dict) else col, False)

    columns = []
    for ci in sorted(headers):
        label, is_date_header = headers[ci]
        p = profile.get(ci, {})
        n = max(p.get("n", 0), 1)
        kind = "blank"
        if p.get("formula", 0) >= 0.5 * n:
            kind = "formula"
        elif p.get("n", 0):
            counts = {k: p.get(k, 0) for k in ("number", "text", "date", "bool")}
            kind = max(counts, key=counts.get) if any(counts.values()) else "blank"
        month = month_of_header(label, date.fromisoformat(label) if is_date_header else None)
        col = {"col": col_letter(ci), "index": ci, "header": label, "kind": kind,
               "filled": p.get("n", 0) - p.get("blank", 0), "formulas": p.get("formula", 0),
               "formula_sample": (p.get("formula_sample") or "")[:300] or None,
               "distinct": sorted(p["distinct"])[:DISTINCT_CAP] if p.get("distinct") is not None else None,
               "style": p.get("style"), "keyish": p.get("keyish", 0)}
        if month:
            col["month"] = {"month": month[0], "year": month[1]}
            info["month_cols"].append(col["col"])
        columns.append(col)
    info["columns"] = columns
    info["data_first"] = (header_row + 1) if header_row else None

    key = _pick_key(columns)
    if key:
        info["key_col"] = key["col"]
    info["depends_on"] = sorted(d for d in deps if d != sheet["name"])
    info["table_refs"] = sorted(table_deps)
    return info


def _pick_header(head_rows, table_header, pkg):
    if table_header and table_header in head_rows:
        return table_header
    best, score = None, 1
    for r in sorted(head_rows):
        cells = head_rows[r]
        texts = sum(1 for _c, k, v, f, _s in cells if k == "text" and not f and str(v).strip())
        dates = sum(1 for _c, k, v, f, _s in cells if k in ("number", "date") and not f and excel_date(v))
        sc = texts + (dates if texts >= 1 else 0)
        below = [rr for rr in head_rows if rr > r]
        if sc > score and below:
            best, score = r, sc
    return best


def _profile(profile, ci, kind, value, formula, offset, style):
    p = profile.setdefault(ci, {"n": 0, "blank": 0, "formula": 0, "number": 0, "text": 0, "date": 0,
                                "bool": 0, "distinct": set(), "keyish": 0})
    if offset > PROFILE_ROWS:
        # past the profile window: only track distinct low-cardinality text (entry types, prefixes)
        if kind == "text" and not formula and p.get("distinct") is not None:
            _distinct(p, value)
        return
    p["n"] += 1
    if p.get("style") is None and style is not None and not formula:
        p["style"] = style
    if formula:
        p["formula"] += 1
        if not p.get("formula_sample"):
            p["formula_sample"] = "=" + formula
        return
    if kind == "blank" or (kind == "text" and not str(value).strip()):
        p["blank"] += 1
        return
    p[kind] = p.get(kind, 0) + 1
    if kind == "text":
        if ITEM_KEY.match(str(value).strip()):
            p["keyish"] += 1
        if p.get("distinct") is not None:
            _distinct(p, value)


def _distinct(p, value):
    s = str(value).strip()
    if not s:
        return
    p["distinct"].add(s[:60])
    if len(p["distinct"]) > DISTINCT_CAP:
        p["distinct"] = None


def _pick_key(columns):
    best = None
    for col in columns:
        filled = max(col["filled"], 1)
        if col["kind"] == "formula" or col["keyish"] < 0.6 * filled or col["filled"] < 1:
            continue
        if best is None or col["keyish"] > best["keyish"]:
            best = col
    return best


def collect_keys(pkg: Package, sheet: dict, key_col: str, header_row: int, cap=KEY_CAP):
    """Item keys in the key column (template's own row list)."""
    out = []
    ci = col_index(key_col)
    with pkg.zf.open(sheet["part"]) as f:
        for _ev, row in etree.iterparse(f, events=("end",), tag=NS + "row"):
            r = int(row.get("r") or 0)
            if r > header_row:
                for c in row.findall(NS + "c"):
                    letter, _ = split_ref(c.get("r") or "")
                    if letter and col_index(letter) == ci:
                        kind, value, _f = _cell_value(c, pkg)
                        if kind == "text" and str(value).strip():
                            out.append(str(value).strip())
                        break
            row.clear()
            while row.getprevious() is not None:
                del row.getparent()[0]
            if len(out) >= cap:
                break
    return out


def workbook_map(data: bytes, name: str = "workbook.xlsx", on_progress=None, fast: bool = False) -> dict:
    """fast=True caps each sheet at a few thousand rows (chat context); the rebuild job uses the full map."""
    pkg = Package(data)
    max_rows = 3000 if fast else None
    sheets = []
    table_sheet = {}
    for i, sheet in enumerate(pkg.sheets):
        if on_progress:
            try:
                on_progress(i + 1, len(pkg.sheets), sheet["name"])
            except Exception:
                pass
        try:
            info = analyse_sheet(pkg, sheet, max_rows=max_rows)
        except Exception as e:  # one unreadable sheet must not sink the map
            info = {"name": sheet["name"], "index": sheet["index"], "state": sheet["state"], "part": sheet["part"],
                    "error": f"{type(e).__name__}: {e}"[:200], "columns": [], "tables": [], "pivots": [],
                    "rows": 0, "header_row": None, "key_col": None, "depends_on": [], "month_cols": []}
        if info.get("key_col") and info.get("header_row"):
            keys = collect_keys(pkg, sheet, info["key_col"], info["header_row"], cap=200 if fast else KEY_CAP)
            info["key_values"] = keys
            info["key_count"] = len(keys)
            info["key_prefixes"] = sorted({re.split(r"[-\d]", k, maxsplit=1)[0] + "-" for k in keys if "-" in k})[:20]
        for t in info.get("tables") or []:
            table_sheet[t["name"]] = sheet["name"]
        sheets.append(info)
    # Structured references (Table1[...]) count as dependencies on the table's sheet
    for info in sheets:
        extra = {table_sheet[t] for t in info.get("table_refs") or [] if t in table_sheet}
        info["depends_on"] = sorted((set(info.get("depends_on") or []) | extra) - {info["name"]})
    feeds = {}
    for info in sheets:
        for d in info.get("depends_on") or []:
            feeds.setdefault(d, []).append(info["name"])
        for p in info.get("pivots") or []:
            if p.get("source_sheet"):
                feeds.setdefault(p["source_sheet"], []).append(f"{info['name']} (pivot {p['name']})")
    for info in sheets:
        info["feeds"] = sorted(set(feeds.get(info["name"], [])))
    root = etree.fromstring(pkg.workbook_xml)
    names = [{"name": d.get("name"), "ref": (d.text or "")[:200]} for d in root.iter(NS + "definedName")
             if not (d.get("name") or "").startswith("_xlnm")]
    flags = []
    if any(n.startswith("xl/pivotCache") for n in pkg.names):
        flags.append("Pivot tables present — set to refresh when the file is opened in Excel")
    if "xl/vbaProject.bin" in pkg.names:
        flags.append("VBA macros present — preserved, not executed")
    if any("externalLinks" in n for n in pkg.names):
        flags.append("External workbook links present")
    return {"name": name, "sheets": sheets, "defined_names": names[:100], "unsupported": flags,
            "parts": sorted(pkg.names)}


def compact(wmap: dict, key_sample=8) -> dict:
    """LLM-sized view of the map: structure and meaning, no bulk values."""
    out = []
    for s in wmap["sheets"]:
        out.append({
            "sheet": s["name"], "state": s.get("state"), "rows": s.get("rows"), "header_row": s.get("header_row"),
            "tables": [{"name": t["name"], "ref": t["ref"], "calculated": list(t.get("calculated", {}))[:20]}
                       for t in s.get("tables") or []],
            "pivots": [{"name": p["name"], "location": p["location"], "source": p.get("source_sheet"),
                        "rows": p.get("rows"), "cols": p.get("cols"), "data": p.get("data")}
                       for p in s.get("pivots") or []],
            "key_col": s.get("key_col"), "key_count": s.get("key_count"),
            "key_sample": (s.get("key_values") or [])[:key_sample], "key_prefixes": s.get("key_prefixes"),
            "depends_on": s.get("depends_on"), "feeds": s.get("feeds"),
            "columns": [{"col": c["col"], "header": c["header"][:60], "kind": c["kind"], "filled": c["filled"],
                         **({"formula": c["formula_sample"][:200]} if c.get("formula_sample") else {}),
                         **({"values": c["distinct"][:12]} if c.get("distinct") and c["kind"] == "text"
                            and len(c["distinct"]) <= 15 else {}),
                         **({"month": c["month"]} if c.get("month") else {})}
                        for c in (s.get("columns") or [])[:80]],
        })
    return {"workbook": wmap["name"], "sheets": out, "defined_names": wmap.get("defined_names", [])[:30]}


def flow_text(wmap: dict, limit=12000) -> str:
    """Readable workbook flow for chat context."""
    lines = [f"Workbook: {wmap['name']} — {len(wmap['sheets'])} sheets"]
    for s in wmap["sheets"]:
        bits = [f"Sheet '{s['name']}'" + (" (hidden)" if s.get("state") != "visible" else ""),
                f"{s.get('rows', 0):,} rows"]
        if s.get("tables"):
            bits.append("table " + ", ".join(f"{t['name'][:30]} {t['ref']}" for t in s["tables"]))
        if s.get("key_col"):
            bits.append(f"keyed by column {s['key_col']} ({s.get('key_count', 0)} items, "
                        f"prefixes {', '.join(s.get('key_prefixes') or [])})")
        if s.get("truncated"):
            bits[1] = f"{s.get('rows', 0):,}+ rows"
        lines.append(" · ".join(bits))
        cols = s.get("columns") or []
        if cols:
            lines.append("  Columns: " + "; ".join(
                f"{c['col']} {c['header'][:30]} [{c['kind']}]" for c in cols[:30])
                + (f"; …{len(cols) - 30} more" if len(cols) > 30 else ""))
        shown = [c for c in cols if c.get("formula_sample")][:8]
        for c in shown:
            lines.append(f"  {c['col']} {c['header'][:30]} = {c['formula_sample'][:140]}")
        for p in s.get("pivots") or []:
            lines.append(f"  Pivot {p['name']} at {p['location']} from '{p.get('source_sheet')}' "
                         f"rows={p.get('rows')} cols={p.get('cols')} data={p.get('data')}")
        if s.get("depends_on"):
            lines.append("  Reads from: " + ", ".join(s["depends_on"]))
        if s.get("feeds"):
            lines.append("  Feeds: " + ", ".join(s["feeds"]))
    if wmap.get("unsupported"):
        lines.append("Notes: " + "; ".join(wmap["unsupported"]))
    text = "\n".join(lines)
    return text[:limit]
