"""KSNV data folder: cumulative policy history + the monthly reference files.

Layout (sync-friendly: one file per line per month, a re-import replaces that month's file):
  <store>/policies/<line>/<YYYY-MM>.csv   one row per đơn/SĐBS x product x type (PIAS rows aggregated; phi = phí phần PVI)
  <store>/cgtkt/CG&TKT_<YYYY-MM>.xlsx     revenue statistics (CG&TKT) for the period ending that month
  <store>/ref/<file>                      last-year KSNV workbook (doanh thu lũy kế năm trước)
  <store>/import_log.csv                  what was imported, when, by whom
"""
import csv
import datetime
import getpass
import glob
import os
import re
import shutil

import xlsx_fast as X

POLICY_COLS = {"Số đơn BH/SĐBS", "Loại đơn", "Ngày cấp đơn", "Ngày cuối", "Mã sp",
               "Phí bảo hiểm(VNĐ)", "Mã sđbs", "Nguồn cấp đơn", "T.lệ Đồng chính", "T.lệ Đồng phụ"}


def pvi_share(r):
    """PVI's share of a co-insured policy: leader -> T.lệ Đồng chính, follower -> T.lệ Đồng phụ."""
    def num(k):
        try:
            return float(r.get(k))
        except (TypeError, ValueError):
            return None
    dc, dp = num("T.lệ Đồng chính"), num("T.lệ Đồng phụ")
    if dc is None and dp is None:
        return 1.0
    if dc and dc > 0:
        return dc / 100
    return (dp or 0) / 100
FIELDS = ["so", "sp", "loai", "nguon", "ngay_cap", "ngay_cuoi", "goc", "phi", "so_dong"]
REV_SHEET = re.compile(r"^\s*Chi tiet thang n\S*m (20\d{2}) NV\s*$", re.I)   # BC doanh thu thực hiện (P.KD)
XCG_SHEET = re.compile(r"^\s*Rà soát XCG Tháng (\d{1,2})\s*$", re.I)           # file rà soát XCG các P.KD đã giải trình


class StoreError(Exception):
    pass


class Policy:
    __slots__ = ("so", "loai", "day", "end", "sp", "fee", "goc", "nguon")


def _s(v):
    return "" if v is None else str(v).strip()


def _d(v):
    d = X.excel_date(v)
    return d.strftime("%Y-%m-%d") if d else ""


# ---------------------------------------------------------------- classify an uploaded file
def classify(path):
    """('policy', sheet) | ('cgtkt', (year, month)) | ('prevyear', year) | ('skip', reason)"""
    name = os.path.basename(path)
    if not name.lower().endswith((".xlsx", ".xlsm")):
        return "skip", "không phải file Excel"
    try:
        sheets = X.sheet_names(path)
    except Exception as e:
        return "skip", "không đọc được (%s)" % e
    if "CG&TKT" in sheets:
        for r in X.iter_rows(path, "CG&TKT", 0):
            txt = " ".join(str(v) for k, v in r.items() if k != "_r" and v)
            m = re.search(r"Đến ngày\s*:\s*(\d{1,2})/(\d{1,2})/(\d{4})", txt)
            if m:
                return "cgtkt", (int(m.group(3)), int(m.group(2)))
            if r["_r"] > 6:
                break
        return "skip", "file CG&TKT không có dòng 'Đến ngày'"
    for s in sheets:
        m = REV_SHEET.match(s)
        if m:
            return "revenue", int(m.group(1))
    for s in sheets:
        m = XCG_SHEET.match(s)
        if m:
            for r in X.iter_rows(path, s, 0):
                txt = " ".join(str(v) for k, v in r.items() if k != "_r" and v)
                d = re.search(r"Từ\s*:\s*\d{1,2}/(\d{1,2})/(\d{4})", txt)
                if d:
                    return "xcg", (int(d.group(2)), int(d.group(1)))
                if r["_r"] > 6:
                    break
            return "skip", "file rà soát XCG không có dòng 'Cấp Từ: dd/mm/yyyy'"
    if "THỰC HIỆN" in sheets and "CẤP ĐƠN" in sheets:
        m = re.search(r"(20\d{2})", name)
        return "prevyear", int(m.group(1)) if m else None
    sh = "Sheet" if "Sheet" in sheets else sheets[0]
    for r in X.iter_rows(path, sh, 0):
        if any(v == "Số đơn BH/SĐBS" for k, v in r.items() if k != "_r"):
            return "policy", sh
        if r["_r"] >= 1:
            break
    return "skip", "không nhận ra loại file"


# ---------------------------------------------------------------- import
def read_policy_agg(path, sheet, log=print):
    """Aggregate a PIAS export: key (so, sp, loai, nguon, ngay_cap, ngay_cuoi, goc) -> [fee VND, rows]."""
    agg, n = {}, 0
    for r in X.iter_rows(path, sheet, 1, POLICY_COLS):
        so = _s(r.get("Số đơn BH/SĐBS"))
        sp = _s(r.get("Mã sp"))
        if not so or not sp:
            continue
        n += 1
        if n % 50000 == 0:
            log("  %s rows ..." % f"{n:,}")
        try:
            fee = float(r.get("Phí bảo hiểm(VNĐ)") or 0) * pvi_share(r)    # doanh thu phần PVI
        except (TypeError, ValueError):
            fee = 0.0
        key = (so, sp, _s(r.get("Loại đơn")), _s(r.get("Nguồn cấp đơn")), _d(r.get("Ngày cấp đơn")),
               _d(r.get("Ngày cuối")), "1" if _s(r.get("Mã sđbs")) == "" else "0")
        a = agg.get(key)
        if a:
            a[0] += fee
            a[1] += 1
        else:
            agg[key] = [fee, 1]
    return n, agg


def write_months(store, agg):
    """Write one CSV per (line, month) found in agg, replacing what was there."""
    by = {}
    for k, (fee, cnt) in agg.items():
        if k[4]:
            by.setdefault((k[1][:2], k[4][:7]), []).append(list(k) + [round(fee), cnt])
    written = []
    for (line, ym), rows in sorted(by.items()):
        d = os.path.join(store, "policies", line)
        os.makedirs(d, exist_ok=True)
        f = os.path.join(d, ym + ".csv")
        with open(f + ".tmp", "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(FIELDS)
            w.writerows(sorted(rows))
        os.replace(f + ".tmp", f)
        written.append((line, ym, len(rows)))
    return written


def import_files(paths, store, log=print):
    """Import a batch. Policy files are merged first (a key seen in two files counts once), so two
    exports that share a month do not overwrite each other; then each (line, month) file is replaced."""
    os.makedirs(store, exist_ok=True)
    done, skipped, batch = [], [], {}
    for p in paths:
        name = os.path.basename(p)
        kind, info = classify(p)
        if kind == "skip":
            skipped.append((name, info))
            log("Skip %s: %s" % (name, info))
            continue
        log("Importing %s ..." % name)
        if kind == "policy":
            n, agg = read_policy_agg(p, info, log)
            for k, v in agg.items():
                batch.setdefault(k, v)
            months = sorted({k[4][:7] for k in agg if k[4]})
            lines = sorted({k[1][:2] for k in agg})
            done.append({"file": name, "kind": "Danh sách đơn", "rows": n, "lines": ", ".join(lines),
                         "months": "%s → %s" % (months[0], months[-1]) if months else "-"})
        elif kind == "cgtkt":
            y, m = info
            d = os.path.join(store, "cgtkt")
            os.makedirs(d, exist_ok=True)
            shutil.copy2(p, os.path.join(d, "CG&TKT_%04d-%02d.xlsx" % (y, m)))
            done.append({"file": name, "kind": "CG&TKT", "rows": "", "lines": "", "months": "%04d-%02d" % (y, m)})
        elif kind == "revenue":
            d = os.path.join(store, "revenue")
            os.makedirs(d, exist_ok=True)
            shutil.copy2(p, os.path.join(d, "Revenue_%d.xlsx" % info))
            done.append({"file": name, "kind": "BC doanh thu %d" % info, "rows": "", "lines": "", "months": ""})
        elif kind == "xcg":
            y, m = info
            d = os.path.join(store, "xcg")
            os.makedirs(d, exist_ok=True)
            shutil.copy2(p, os.path.join(d, "XCG_Review_%04d-%02d.xlsx" % (y, m)))
            done.append({"file": name, "kind": "Rà soát XCG", "rows": "", "lines": "", "months": "%04d-%02d" % (y, m)})
        elif kind == "prevyear":
            d = os.path.join(store, "ref")
            os.makedirs(d, exist_ok=True)
            for old in glob.glob(os.path.join(d, "*.xlsx")):
                if info and str(info) in os.path.basename(old):
                    os.remove(old)
            shutil.copy2(p, os.path.join(d, name))
            done.append({"file": name, "kind": "BC KSNV năm %s" % (info or "?"), "rows": "", "lines": "", "months": ""})
    written = write_months(store, batch) if batch else []
    _log_import(store, done)
    return done, skipped, written


def _log_import(store, done):
    f = os.path.join(store, "import_log.csv")
    new = not os.path.exists(f)
    with open(f, "a", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        if new:
            w.writerow(["time", "user", "file", "kind", "rows", "lines", "months"])
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        for d in done:
            w.writerow([now, getpass.getuser(), d["file"], d["kind"], d["rows"], d["lines"], d["months"]])


# ---------------------------------------------------------------- read
def _read_csv(f):
    out = []
    with open(f, newline="", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            p = Policy()
            p.so, p.sp, p.loai, p.nguon = r["so"], r["sp"], r["loai"], r["nguon"]
            p.day = int(r["ngay_cap"][8:10]) if r["ngay_cap"] else None
            p.end = datetime.datetime.strptime(r["ngay_cuoi"], "%Y-%m-%d") if r["ngay_cuoi"] else None
            p.goc = r["goc"] == "1"
            p.fee = float(r["phi"] or 0) / 1e6
            out.append(p)
    return out


def month_policies(store, year, month):
    """{line: [Policy]} issued in that month."""
    ym = "%04d-%02d" % (year, month)
    out = {}
    for f in glob.glob(os.path.join(store, "policies", "*", ym + ".csv")):
        out[os.path.basename(os.path.dirname(f))] = _read_csv(f)
    return out


def expiring(store, year, month, lines):
    """{line: [Policy]} of every stored month whose Ngày cuối falls in year/month."""
    out = {}
    for ln in lines:
        rows = []
        for f in sorted(glob.glob(os.path.join(store, "policies", ln, "*.csv"))):
            rows += [p for p in _read_csv(f) if p.end and p.end.year == year and p.end.month == month]
        if rows:
            out[ln] = rows
    return out


def coverage(store):
    """{line: [YYYY-MM, ...]} months present in the store."""
    out = {}
    for d in sorted(glob.glob(os.path.join(store, "policies", "*"))):
        out[os.path.basename(d)] = sorted(os.path.basename(f)[:7] for f in glob.glob(os.path.join(d, "*.csv")))
    return out


def cgtkt_file(store, year, month):
    f = os.path.join(store, "cgtkt", "CG&TKT_%04d-%02d.xlsx" % (year, month))
    return f if os.path.exists(f) else None


def revenue_file(store, year):
    f = os.path.join(store, "revenue", "Revenue_%d.xlsx" % year)
    return f if os.path.exists(f) else None


def xcg_file(store, year, month):
    f = os.path.join(store, "xcg", "XCG_Review_%04d-%02d.xlsx" % (year, month))
    return f if os.path.exists(f) else None


def xcg_key(so, an_chi, noi_dung):
    """Match key between the review file and tab RÀ SOÁT XCG: số đơn + số ấn chỉ + nội dung cần rà soát."""
    def n(v):
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        s = _s(v)
        return s[:-2] if re.fullmatch(r"\d+\.0", s) else s
    return n(so), n(an_chi), n(noi_dung)


def read_xcg_review(path):
    """{key: kết quả hậu kiểm} from a 'Rà soát XCG Tháng m' file (data from row 15, cột C, D, P, Q)."""
    sh = next(s for s in X.sheet_names(path) if XCG_SHEET.match(s))
    out = {}
    for r in X.iter_rows(path, sh, 0):
        if r["_r"] < 15 or not _s(r.get(2)):
            continue
        q = r.get(16)
        out[xcg_key(r.get(2), r.get(3), r.get(15))] = "" if q in (None, 0, "0") else _s(q)
    return out


def revenue_sheet(path):
    return next(s for s in X.sheet_names(path) if REV_SHEET.match(s))


def prevyear_file(store, year):
    fs = glob.glob(os.path.join(store, "ref", "*.xlsx"))
    hit = [f for f in fs if str(year - 1) in os.path.basename(f)]
    return (hit or fs or [None])[0]
