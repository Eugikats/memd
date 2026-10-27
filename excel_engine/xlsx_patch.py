"""Patch the original xlsx in place: write system values into planned input cells, keep everything else.

Every zip part is copied as-is except the sheets the plan touches. Formula cells are never overwritten; formula
columns of replaced ledger tables are re-applied to each new row. Excel recalculates formulas and refreshes pivots
when the file is opened (fullCalcOnLoad, refreshOnLoad), so calculated sheets follow the refreshed sources.
"""
from __future__ import annotations

import io
import re
import shutil
import tempfile
import zipfile
from datetime import date, datetime
from xml.sax.saxutils import escape

from excel_engine.understand import NS, Package, col_index, col_letter, split_ref

_ROW = rb"<%srow\b[^>]*?(?:/>|>.*?</%srow>)"
_CELL = rb"<%sc\b[^>]*?(?:/>|>.*?</%sc>)"
_ATTR_R = re.compile(rb'\br="([A-Z]+)(\d+)"')
_ROW_R = re.compile(rb'^<[\w:]*row\b[^>]*?\br="(\d+)"')
_STYLE = re.compile(rb'\bs="(\d+)"')
_SPANS = re.compile(rb'\s+spans="[^"]*"')
MISSING_SHEET = "Not in template"
MISSING_CAP = 5000
_EPOCH = date(1899, 12, 30)


def _num(v):
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, int):
        return str(v)
    f = float(v)
    if f.is_integer() and abs(f) < 1e15:
        return str(int(f))
    return repr(round(f, 9))


def _cell(ref, value, style=None, formula=None):
    s = f' s="{style.decode() if isinstance(style, bytes) else style}"' if style else ""
    if formula is not None:
        return f'<c r="{ref}"{s}><f>{escape(formula)}</f></c>'.encode()
    if value is None or value == "":
        return f'<c r="{ref}"{s}/>'.encode()
    if isinstance(value, bool):
        return f'<c r="{ref}"{s} t="b"><v>{"1" if value else "0"}</v></c>'.encode()
    if isinstance(value, (datetime, date)):
        d = value.date() if isinstance(value, datetime) else value
        return f'<c r="{ref}"{s}><v>{(d - _EPOCH).days}</v></c>'.encode()
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{s}><v>{_num(value)}</v></c>'.encode()
    text = str(value)
    space = ' xml:space="preserve"' if text != text.strip() else ""
    return f'<c r="{ref}"{s} t="inlineStr"><is><t{space}>{escape(text)}</t></is></c>'.encode()


class SheetXml:
    """Byte-level view of a worksheet part: prefix, rows (as raw bytes), suffix."""

    def __init__(self, xml: bytes):
        m = re.search(rb"<([\w]+:)?sheetData\b[^>]*?(/?)>", xml)
        if not m:
            raise ValueError("worksheet has no sheetData")
        self.ns = m.group(1) or b""
        if m.group(2) == b"/":
            self.prefix = xml[:m.start()] + b"<" + self.ns + b"sheetData>"
            self.body = b""
            self.suffix = b"</" + self.ns + b"sheetData>" + xml[m.end():]
        else:
            end = xml.rfind(b"</" + self.ns + b"sheetData>")
            self.prefix = xml[:m.end()]
            self.body = xml[m.end():end]
            self.suffix = xml[end:]
        self.row_re = re.compile(_ROW % (self.ns, self.ns), re.S)
        self.cell_re = re.compile(_CELL % (self.ns, self.ns), re.S)

    def rows(self):
        for m in self.row_re.finditer(self.body):
            raw = m.group(0)
            rm = _ROW_R.match(raw)
            yield (int(rm.group(1)) if rm else None), raw

    def cells(self, row_raw):
        out = {}
        for m in self.cell_re.finditer(row_raw):
            c = m.group(0)
            a = _ATTR_R.search(c[:80])
            if a:
                out[col_index(a.group(1).decode())] = c
        return out

    def row_open(self, row_raw):
        end = row_raw.find(b">")
        head = row_raw[:end + 1]
        if head.endswith(b"/>"):
            head = head[:-2] + b">"
        return _SPANS.sub(b"", head)

    def row_close(self):
        return b"</" + self.ns + b"row>"

    def build_row(self, open_tag, cells):
        return open_tag + b"".join(cells[k] for k in sorted(cells)) + self.row_close()

    def new_row_open(self, r):
        return b"<" + self.ns + b'row r="%d">' % r


def _cell_text(cell, shared):
    """Plain value of a raw <c> (string or number) for key lookups."""
    t = re.search(rb'\bt="(\w+)"', cell[:120])
    t = t.group(1) if t else None
    if t == b"inlineStr":
        m = re.search(rb"<(?:\w+:)?t\b[^>]*>(.*?)</(?:\w+:)?t>", cell, re.S)
        return _unescape(m.group(1).decode("utf-8", "replace")) if m else ""
    v = re.search(rb"<(?:\w+:)?v>(.*?)</(?:\w+:)?v>", cell, re.S)
    if not v:
        return ""
    raw = v.group(1).decode("utf-8", "replace")
    if t == b"s":
        try:
            return shared[int(raw)]
        except (ValueError, IndexError):
            return ""
    return _unescape(raw)


def _unescape(s):
    return s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&")


def _is_formula_cell(cell):
    return b"<f" in cell or b":f>" in cell or b":f " in cell


def _strip_cached(cell):
    """Formula cell without its cached value — Excel recalculates on open."""
    return re.sub(rb"<(\w+:)?v>.*?</(\w+:)?v>", b"", cell, flags=re.S)


def _style(cell):
    m = _STYLE.search(cell[:120])
    return m.group(1) if m else None


_A1 = re.compile(r"(?<![A-Za-z_\[\]!\$'\"#.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(\[])")


def _has_relative_rows(formula):
    return any(not m.group(3) for m in _A1.finditer(_strip_strings(formula)))


def _strip_strings(f):
    return re.sub(r'"[^"]*"', '""', f)


def _translate(formula, origin_ref, target_ref):
    from openpyxl.formula.translate import Translator
    try:
        return Translator("=" + formula, origin=origin_ref).translate_formula(target_ref)[1:]
    except Exception:
        return formula


class Patcher:
    def __init__(self, data: bytes, wmap: dict, plan: dict, sources, on_progress=None):
        self.data = data
        self.pkg = Package(data)
        self.wmap = wmap
        self.plan = plan
        self.src = sources
        self.on_progress = on_progress
        self.parts = {}            # part name -> new bytes or temp file path
        self.results = []
        self.table_refs = {}       # table part -> new ref
        self.missing = []
        self.drop = set()

    # ---------------------------------------------------------------- keyed update
    def keyed_update(self, entry, sinfo):
        sx = SheetXml(self.pkg.read(sinfo["part"]))
        header = sinfo["header_row"]
        key_ci = col_index(entry["key_col"])
        cols = [c for c in entry["columns"] if c["action"] in ("fill", "fill_month", "fill_location")]
        targets = {col_index(c["col"]): c for c in cols}
        col_styles = {c["index"]: c.get("style") for c in sinfo.get("columns") or []}
        shared = self.pkg.shared
        out = io.BytesIO()
        out.write(sx.prefix)
        keys_seen, matched, unknown, written, protected = [], 0, [], 0, 0
        for r, raw in sx.rows():
            if r is None or r <= header:
                out.write(raw)
                continue
            cells = sx.cells(raw)
            key_cell = cells.get(key_ci)
            key = _cell_text(key_cell, shared).strip() if key_cell else ""
            if not key:
                out.write(raw)
                continue
            keys_seen.append(key)
            if not self.src.known(key):
                unknown.append(key)
            else:
                matched += 1
            changed = False
            for ci, c in targets.items():
                old = cells.get(ci)
                if old is not None and _is_formula_cell(old):
                    protected += 1
                    continue
                value = self._item_target(c, key)
                style = _style(old) if old is not None else (col_styles.get(ci) or None)
                cells[ci] = _cell(f"{col_letter(ci)}{r}", value, style)
                changed = True
                if value is not None:
                    written += 1
            if changed:
                for ci, cell in list(cells.items()):
                    if ci not in targets and _is_formula_cell(cell):
                        cells[ci] = _strip_cached(cell)
                out.write(sx.build_row(sx.row_open(raw), cells))
            else:
                out.write(raw)
        out.write(sx.suffix)
        self.parts[sinfo["part"]] = out.getvalue()
        if entry.get("append_missing"):
            self._collect_missing(entry, sinfo, set(keys_seen))
        return {"rows": len(keys_seen), "matched": matched, "unknown": len(unknown),
                "unknown_sample": unknown[:10], "cells_written": written, "formula_cells_protected": protected}

    def _item_target(self, c, key):
        if c["action"] == "fill":
            return self.src.item_value(c["field"], key)
        if c["action"] == "fill_location":
            return self.src.location_value(c["location"], key)
        if c["action"] == "fill_month":
            return self.src.month_value(c["measure"], c["month"], key)
        return None

    def _collect_missing(self, entry, sinfo, seen):
        prefixes = tuple(entry.get("item_prefixes") or sinfo.get("key_prefixes") or [])
        if not prefixes:
            return
        counts = {}
        for k in seen:
            for p in prefixes:
                if k.startswith(p):
                    counts[p] = counts.get(p, 0) + 1
        major = tuple(p for p in prefixes if counts.get(p, 0) >= max(5, 0.1 * len(seen))) or prefixes
        for item_no, row in self.src.measures.items():
            if item_no in seen or not str(item_no).startswith(major):
                continue
            if not (row.get("stock_qty") or row.get("usage_qty") or row.get("incoming_qty")):
                continue
            self.missing.append((sinfo["name"], item_no, row))
            if len(self.missing) >= MISSING_CAP:
                return

    # ---------------------------------------------------------------- ledger replace
    def replace_rows(self, entry, sinfo):
        sx = SheetXml(self.pkg.read(sinfo["part"]))
        header = sinfo["header_row"]
        plan_cols = {col_index(c["col"]): c for c in entry["columns"]}
        table = (sinfo.get("tables") or [None])[0]
        span = None
        if table:
            rng = re.findall(r"[A-Z]+", table["ref"].replace("$", ""))
            span = (col_index(rng[0]), col_index(rng[-1]))
        else:
            idx = [c["index"] for c in sinfo.get("columns") or []]
            span = (min(idx), max(idx)) if idx else (1, 1)
        rows = self.src.ledger(entry.get("entry_types"), entry.get("item_prefixes"))
        n = len(rows)
        if not n:
            return {"rows": 0, "cells_written": 0, "reason": "no system ledger rows for these entry types / items "
                                                             "(12-month history not loaded yet)"}
        first_data = header + 1
        new_last = header + n

        # Template row: styles and formulas per column from the first data row
        tmpl_styles, tmpl_formulas = {}, {}
        for r, raw in sx.rows():
            if r == first_data:
                for ci, cell in sx.cells(raw).items():
                    if span[0] <= ci <= span[1]:
                        tmpl_styles[ci] = _style(cell)
                        fm = re.search(rb"<(?:\w+:)?f\b[^>]*>(.*?)</(?:\w+:)?f>", cell, re.S)
                        if fm:
                            tmpl_formulas[ci] = _unescape(fm.group(1).decode("utf-8", "replace"))
                break
        calc = (table or {}).get("calculated") or {}
        if table:
            for i, name in enumerate(table.get("columns") or []):
                ci = span[0] + i
                if name in calc and ci not in tmpl_formulas:
                    tmpl_formulas[ci] = calc[name]
        for ci, c in plan_cols.items():
            if c["action"] == "keep_formula" and ci not in tmpl_formulas:
                c["action"] = "leave"
        relative = {ci: _has_relative_rows(f) for ci, f in tmpl_formulas.items()}
        fill_cols = {ci: c for ci, c in plan_cols.items() if c["action"] == "fill" and span[0] <= ci <= span[1]}
        formula_cols = {ci: tmpl_formulas[ci] for ci, c in plan_cols.items()
                        if c["action"] == "keep_formula" and span[0] <= ci <= span[1] and ci in tmpl_formulas}

        def generated_cells(r):
            led = rows[r - first_data]
            cells = {}
            for ci, c in fill_cols.items():
                cells[ci] = _cell(f"{col_letter(ci)}{r}", self.src.ledger_value(c["field"], led), tmpl_styles.get(ci))
            for ci, f in formula_cols.items():
                if relative.get(ci) and r != first_data:
                    f = _translate(f, f"{col_letter(ci)}{first_data}", f"{col_letter(ci)}{r}")
                cells[ci] = _cell(f"{col_letter(ci)}{r}", None, tmpl_styles.get(ci), formula=f)
            return cells

        tmp = tempfile.TemporaryFile()
        tmp.write(_dimension(sx.prefix, span, max(new_last, header)))
        prev = header
        for r, raw in sx.rows():
            if r is None or r <= header:
                tmp.write(raw)
                prev = r or prev
                continue
            # rows missing from the original between prev and r
            for gr in range(prev + 1, min(r, new_last + 1)):
                tmp.write(sx.build_row(sx.new_row_open(gr), generated_cells(gr)))
            cells = sx.cells(raw)
            keep = {ci: c for ci, c in cells.items() if not (span[0] <= ci <= span[1])}
            if r <= new_last:
                keep.update(generated_cells(r))
            if keep:
                tmp.write(sx.build_row(sx.row_open(raw), keep))
            prev = r
        for gr in range(prev + 1, new_last + 1):
            tmp.write(sx.build_row(sx.new_row_open(gr), generated_cells(gr)))
        tmp.write(sx.suffix)
        self.parts[sinfo["part"]] = tmp
        if table:
            top = re.match(r"\$?([A-Z]+)\$?(\d+)", table["ref"]).groups()
            bottom = re.findall(r"([A-Z]+)\$?\d+$", table["ref"].replace("$", ""))[0]
            self.table_refs[self._table_part(sinfo, table["name"])] = \
                f"{top[0]}{top[1]}:{bottom}{max(new_last, int(top[1]) + 1)}"
        self._retarget_pivots(sinfo, header, sinfo.get("last_row") or header, new_last)
        return {"rows": n, "cells_written": n * len(fill_cols), "formula_columns": len(formula_cols),
                "entry_types": entry.get("entry_types") or [], "item_prefixes": entry.get("item_prefixes") or [],
                "previous_rows": max((sinfo.get("last_row") or header) - header, 0),
                "first_date": str(rows[0].get("Posting_Date"))[:10] if rows else None,
                "last_date": str(rows[-1].get("Posting_Date"))[:10] if rows else None}

    def _table_part(self, sinfo, name):
        for t in self.pkg.tables_for(sinfo["part"]):
            if t["name"] == name:
                return t["part"]
        return None

    def _retarget_pivots(self, sinfo, header, old_last, new_last):
        """Pivot caches that read an explicit range ending at the old last row follow the new row count."""
        for part, cache in self.pkg.pivot_caches().items():
            if cache.get("source_sheet") != sinfo["name"] or not cache.get("source_ref"):
                continue
            m = re.match(r"^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$", cache["source_ref"])
            if not m or int(m.group(4)) in (1048576,) or int(m.group(4)) != old_last:
                continue
            ref = f"{m.group(1)}{m.group(2)}:{m.group(3)}{max(new_last, int(m.group(2)) + 1)}"
            self._pivot_ref[part] = (cache["source_ref"], ref)

    # ---------------------------------------------------------------- package level
    def run(self):
        self._pivot_ref = {}
        entries = [e for e in self.plan["sheets"] if e["strategy"] != "keep"]
        by_name = {s["name"]: s for s in self.wmap["sheets"]}
        for i, entry in enumerate(entries):
            sinfo = by_name.get(entry["sheet"])
            if self.on_progress:
                try:
                    self.on_progress(i + 1, len(entries), entry["sheet"])
                except Exception:
                    pass
            if not sinfo or not sinfo.get("part") or not sinfo.get("header_row"):
                self.results.append({"sheet": entry["sheet"], "strategy": entry["strategy"], "status": "skipped",
                                     "reason": "sheet layout could not be read"})
                continue
            try:
                if entry["strategy"] == "keyed_update":
                    res = self.keyed_update(entry, sinfo)
                else:
                    res = self.replace_rows(entry, sinfo)
                written = res.get("cells_written", 0)
                status = "refreshed" if written else "not refreshed"
                self.results.append({"sheet": entry["sheet"], "strategy": entry["strategy"], "status": status, **res})
            except Exception as e:
                self.results.append({"sheet": entry["sheet"], "strategy": entry["strategy"], "status": "failed",
                                     "reason": f"{type(e).__name__}: {e}"[:200]})
        return self._write()

    def _write(self):
        zin = self.pkg.zf
        workbook = self._workbook_xml()
        rels = self.pkg.read("xl/_rels/workbook.xml.rels")
        ctypes = self.pkg.read("[Content_Types].xml")
        calc_parts = [r["target"] for r in self.pkg.rels("xl/workbook.xml").values() if r["type"] == "calcChain"]
        for part in calc_parts:
            # calcChain lists formula cells; it is rebuilt by Excel and would be stale after the patch.
            self.drop.add(part)
            rels = re.sub(rb'<Relationship\b[^>]*calcChain[^>]*/>', b"", rels)
            ctypes = re.sub(rb'<Override\b[^>]*PartName="/%s"[^>]*/>' % re.escape(part.encode()), b"", ctypes)
        new_sheet = None
        if self.missing:
            new_sheet, workbook, rels, ctypes = self._add_missing_sheet(workbook, rels, ctypes)
        buf = tempfile.TemporaryFile()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zout:
            for info in zin.infolist():
                name = info.filename
                if name in self.drop:
                    continue
                if name == "xl/workbook.xml":
                    zout.writestr(_info(info), workbook)
                elif name == "xl/_rels/workbook.xml.rels":
                    zout.writestr(_info(info), rels)
                elif name == "[Content_Types].xml":
                    zout.writestr(_info(info), ctypes)
                elif name in self.parts:
                    part = self.parts[name]
                    if isinstance(part, bytes):
                        zout.writestr(_info(info), part)
                    else:
                        part.seek(0)
                        with zout.open(_info(info), "w") as dst:
                            shutil.copyfileobj(part, dst, 1 << 20)
                        part.close()
                elif name in self.table_refs:
                    zout.writestr(_info(info), _retarget_table(zin.read(name), self.table_refs[name]))
                elif re.match(r"xl/pivotCache/pivotCacheDefinition\d+\.xml$", name):
                    zout.writestr(_info(info), _refresh_on_load(zin.read(name), self._pivot_ref.get(name)))
                else:
                    with zin.open(info) as src, zout.open(_info(info), "w") as dst:
                        shutil.copyfileobj(src, dst, 1 << 20)
            if new_sheet:
                zout.writestr(new_sheet[0], new_sheet[1])
        buf.seek(0)
        out = buf.read()
        buf.close()
        return {"bytes": out, "sheets": self.results, "missing_items": len(self.missing),
                "missing_sheet": MISSING_SHEET if new_sheet else None}

    def _workbook_xml(self):
        wb = self.pkg.workbook_xml
        if re.search(rb"<(\w+:)?calcPr\b", wb):
            def fix(m):
                tag = m.group(0)
                tag = re.sub(rb'\s+fullCalcOnLoad="[^"]*"', b"", tag)
                return tag[:-2].rstrip() + b' fullCalcOnLoad="1"/>' if tag.endswith(b"/>") else \
                    tag[:-1].rstrip() + b' fullCalcOnLoad="1">'
            wb = re.sub(rb"<(\w+:)?calcPr\b[^>]*>", fix, wb, count=1)
        else:
            m = re.search(rb"<(\w+:)?workbook\b", wb)
            p = (m.group(1) or b"") if m else b""
            closing = b"</" + p + b"workbook>"
            for anchor in (b"<" + p + b"oleSize", b"<" + p + b"customWorkbookViews", b"<" + p + b"pivotCaches",
                           b"<" + p + b"smartTagPr", b"<" + p + b"webPublishing", b"<" + p + b"fileRecoveryPr",
                           b"<" + p + b"webPublishObjects", b"<" + p + b"extLst", closing):
                i = wb.find(anchor)
                if i >= 0:
                    wb = wb[:i] + b"<" + p + b'calcPr fullCalcOnLoad="1"/>' + wb[i:]
                    break
        return wb

    def _add_missing_sheet(self, workbook, rels, ctypes):
        existing = {n for n in self.pkg.names if n.startswith("xl/worksheets/sheet")}
        k = 1
        while f"xl/worksheets/sheet{k}.xml" in existing:
            k += 1
        part = f"xl/worksheets/sheet{k}.xml"
        rid_nums = [int(x) for x in re.findall(rb'Id="rId(\d+)"', rels)] or [0]
        rid = f"rId{max(rid_nums) + 1}"
        sheet_ids = [int(x) for x in re.findall(rb'sheetId="(\d+)"', workbook)] or [0]
        names = {s["name"].lower() for s in self.pkg.sheets}
        title = MISSING_SHEET
        j = 2
        while title.lower() in names:
            title = f"{MISSING_SHEET} {j}"
            j += 1
        root_tag = re.search(rb"<(\w+:)?workbook\b[^>]*>", workbook).group(0)
        rprefix = re.search(rb'xmlns:(\w+)="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
                            root_tag)
        rp = rprefix.group(1).decode() if rprefix else "r"
        p = (re.search(rb"<(\w+:)?workbook\b", workbook).group(1) or b"").decode()
        decl = "" if rprefix else ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        tag = f'<{p}sheet name="{escape(title)}" sheetId="{max(sheet_ids) + 1}" {rp}:id="{rid}"{decl}/>'.encode()
        workbook = workbook.replace(f"</{p}sheets>".encode(), tag + f"</{p}sheets>".encode(), 1)
        rels = rels.replace(b"</Relationships>", (
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
            f'worksheet" Target="worksheets/sheet{k}.xml"/></Relationships>').encode(), 1)
        ctypes = ctypes.replace(b"</Types>", (
            f'<Override PartName="/{part}" ContentType="application/vnd.openxmlformats-officedocument.'
            f'spreadsheetml.worksheet+xml"/></Types>').encode(), 1)
        return (part, self._missing_xml()), workbook, rels, ctypes

    def _missing_xml(self):
        head = ["Template sheet", "Item No.", "Description", "Class", "UoM", "Stock", "Incoming (open PO)",
                "Avg daily usage", "Cover days", "Cover status"]
        rows = [b"".join(_cell(f"{col_letter(i + 1)}1", h) for i, h in enumerate(head))]
        for n, (sheet, item_no, m) in enumerate(self.missing, start=2):
            vals = [sheet, item_no, m.get("description"), m.get("item_class"), m.get("uom"), m.get("stock_qty"),
                    m.get("incoming_qty"), m.get("daily_usage"), m.get("cover_days"), m.get("cover_status")]
            vals = [None if (isinstance(v, float) and v != v) else v for v in vals]
            rows.append(b"".join(_cell(f"{col_letter(i + 1)}{n}", v) for i, v in enumerate(vals)))
        body = b"".join(b'<row r="%d">' % (i + 1) + r + b"</row>" for i, r in enumerate(rows))
        return (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                b'<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" '
                b'state="frozen"/></sheetView></sheetViews><cols><col min="1" max="1" width="22" customWidth="1"/>'
                b'<col min="2" max="2" width="14" customWidth="1"/><col min="3" max="3" width="48" customWidth="1"/>'
                b'<col min="4" max="10" width="14" customWidth="1"/></cols>'
                b"<sheetData>" + body + b"</sheetData></worksheet>")


def _info(info):
    z = zipfile.ZipInfo(info.filename, date_time=info.date_time)
    z.compress_type = zipfile.ZIP_DEFLATED
    z.external_attr = info.external_attr
    return z


def _dimension(prefix, span, last_row):
    def fix(m):
        ref = m.group(2).decode()
        parts = ref.split(":")
        start = parts[0]
        end_col = re.match(r"([A-Z]+)", parts[-1]).group(1) if len(parts) > 1 else re.match(r"([A-Z]+)", start).group(1)
        end_row = int(re.search(r"(\d+)$", parts[-1]).group(1)) if re.search(r"(\d+)$", parts[-1]) else 1
        end_col = col_letter(max(col_index(end_col), span[1]))
        return m.group(1) + f'{start}:{end_col}{max(end_row, last_row)}"'.encode()
    return re.sub(rb'(<(?:\w+:)?dimension\b[^>]*\bref=")([^"]+)"', fix, prefix, count=1)


def _retarget_table(xml, ref):
    xml = re.sub(rb'(<(?:\w+:)?table\b[^>]*\bref=")[^"]+"', lambda m: m.group(1) + ref.encode() + b'"', xml, count=1)
    return re.sub(rb'(<(?:\w+:)?autoFilter\b[^>]*\bref=")[^"]+"', lambda m: m.group(1) + ref.encode() + b'"', xml,
                  count=1)


def _refresh_on_load(xml, retarget=None):
    m = re.search(rb"<(\w+:)?pivotCacheDefinition\b[^>]*>", xml)
    if not m:
        return xml
    tag = m.group(0)
    new = re.sub(rb'\s+refreshOnLoad="[^"]*"', b"", tag)
    new = new[:-1].rstrip() + b' refreshOnLoad="1">' if not new.endswith(b"/>") else new
    xml = xml[:m.start()] + new + xml[m.end():]
    if retarget:
        old, ref = retarget
        xml = re.sub(rb'(<(?:\w+:)?worksheetSource\b[^>]*\bref=")' + re.escape(old.encode()) + b'"',
                     lambda mm: mm.group(1) + ref.encode() + b'"', xml, count=1)
    return xml


def patch(data: bytes, wmap: dict, plan: dict, sources, on_progress=None) -> dict:
    return Patcher(data, wmap, plan, sources, on_progress).run()
