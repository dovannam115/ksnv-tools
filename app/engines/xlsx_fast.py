"""Fast read-only xlsx reader (stdlib only): yields rows of selected columns as dicts.

Built for big PIAS exports (300k rows): streams the sheet XML with iterparse and keeps only the
columns asked for, so it does not need pandas/openpyxl.
"""
import datetime
import re
import zipfile
import xml.etree.ElementTree as ET

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_COL = re.compile(r"([A-Z]+)")


def _colnum(ref):
    n = 0
    for ch in _COL.match(ref).group(1):
        n = n * 26 + ord(ch) - 64
    return n - 1


def _shared(z):
    try:
        data = z.open("xl/sharedStrings.xml")
    except KeyError:
        return []
    out = []
    for ev, el in ET.iterparse(data):
        if el.tag == NS + "si":
            out.append("".join(t.text or "" for t in el.iter(NS + "t")))
            el.clear()
    return out


def sheet_names(path):
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        return [s.get("name") for s in wb.iter(NS + "sheet")]


def _sheet_path(z, name):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid = None
    sheets = list(wb.iter(NS + "sheet"))
    for s in sheets:
        if name is None or s.get("name") == name:
            rid = s.get(RNS + "id")
            break
    if rid is None:
        raise KeyError(name)
    for r in rels:
        if r.get("Id") == rid:
            t = r.get("Target").lstrip("/")
            return t if t.startswith("xl/") else "xl/" + t
    raise KeyError(name)


def excel_date(v):
    if isinstance(v, (int, float)):
        return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(v))
    if isinstance(v, str) and v:
        for f in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%Y %H:%M:%S"):
            try:
                return datetime.datetime.strptime(v[:19], f)
            except ValueError:
                pass
    return None


def iter_rows(path, sheet=None, header_row=1, columns=None):
    """Yield dict per data row. columns: header names to keep (None = all).
    header_row=0: no header, yield {column index: value, "_r": row number} for every row."""
    with zipfile.ZipFile(path) as z:
        ss = _shared(z)
        sp = _sheet_path(z, sheet)
        want, hdr = None, None
        cur, rnum = {}, 0
        for ev, el in ET.iterparse(z.open(sp)):
            tag = el.tag
            if tag == NS + "c":
                ref = el.get("r")
                if ref is None:
                    el.clear(); continue
                ci = _colnum(ref)
                if want is not None and ci not in want:
                    el.clear(); continue
                t = el.get("t")
                v = el.find(NS + "v")
                if t == "s" and v is not None:
                    val = ss[int(v.text)]
                elif t == "inlineStr":
                    val = "".join(x.text or "" for x in el.iter(NS + "t"))
                elif v is not None and v.text is not None:
                    if t in ("str", "e"):
                        val = v.text
                    elif t == "b":
                        val = v.text == "1"
                    else:
                        try:
                            f = float(v.text); val = int(f) if f.is_integer() else f
                        except ValueError:
                            val = v.text
                else:
                    val = None
                cur[ci] = val
                el.clear()
            elif tag == NS + "row":
                rnum = int(el.get("r") or rnum + 1)
                if header_row == 0:
                    if cur:
                        cur["_r"] = rnum
                        yield cur
                elif rnum == header_row:
                    hdr = {ci: str(v).strip() for ci, v in cur.items() if v is not None}
                    if columns is not None:
                        want = {ci for ci, h in hdr.items() if h in columns}
                elif hdr is not None and rnum > header_row and cur:
                    yield {hdr[ci]: v for ci, v in cur.items() if ci in hdr}
                cur = {}
                el.clear()
