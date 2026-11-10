"""KSNV monthly report engine: read PIAS exports in a month folder, compute values for
Google Sheet tabs CẤP ĐƠN, THỰC HIỆN and LƯU 2026 (SLTT row).

Pure computation, no Google calls, stdlib only (xlsx_fast). The caller passes previous-month
values read from the Sheet (N and M of each leaf line) and gets back plain dicts.
"""
import calendar
import glob
import json
import math
import os
import re

import xlsx_fast as X

HERE = os.path.dirname(os.path.abspath(__file__))
CFG_PATH = os.path.join(HERE, "ksnv_rules.json")

LINES = ["VCX Ô tô", "TNDS Ô tô và khác", "TNDS Xe máy", "VCX Xe máy và khác",
         "CSSK", "GVHS", "CN khác", "MyHealth",
         "Thân tàu", "Hàng hóa", "TNDS/P&I",
         "Cháy - Tài sản", "Kỹ thuật", "Trách nhiệm", "Khác", "Nhận TBH"]
POLICY_COLS = {"Số đơn BH/SĐBS", "Loại đơn", "Ngày cấp đơn", "Ngày cuối", "Mã sp",
               "Phí bảo hiểm(VNĐ)", "Mã sđbs", "Nguồn cấp đơn"}


def load_cfg(path=None):
    with open(path or CFG_PATH, encoding="utf-8") as f:
        return json.load(f)


def _say(log, msg):
    if log:
        log(msg)


# ---------------------------------------------------------------- product -> report line
class LineMap:
    def __init__(self, cfg):
        self.base = cfg["product_line"]
        self.ov = sorted(cfg["line_override"].items(), key=lambda kv: len(kv[0]), reverse=True)

    def __call__(self, sp):
        sp = str(sp or "").strip()
        for k, v in self.ov:
            if sp.startswith(k):
                return v
        g = self.base.get(sp)
        return g if g in LINES else "Khác"


# ---------------------------------------------------------------- input files
class Policy:
    __slots__ = ("so", "loai", "day", "end", "sp", "fee", "goc", "nguon")


def _s(v):
    return "" if v is None else str(v).strip()


def read_policy_file(path):
    names = X.sheet_names(path)
    sh = "Sheet" if "Sheet" in names else names[0]
    out = []
    for r in X.iter_rows(path, sh, 1, POLICY_COLS):
        so = _s(r.get("Số đơn BH/SĐBS"))
        if not so:
            continue
        p = Policy()
        p.so, p.loai, p.sp, p.nguon = so, _s(r.get("Loại đơn")), _s(r.get("Mã sp")), _s(r.get("Nguồn cấp đơn"))
        d = X.excel_date(r.get("Ngày cấp đơn"))
        p.day = d.day if d else None
        p.end = X.excel_date(r.get("Ngày cuối"))
        try:
            p.fee = float(r.get("Phí bảo hiểm(VNĐ)") or 0) / 1e6
        except (TypeError, ValueError):
            p.fee = 0.0
        p.goc = _s(r.get("Mã sđbs")) == ""
        out.append(p)
    return out


def policy_files(folder):
    """Map line code -> path: 'So don/01.xlsx', '01_CN.xlsx'... in the folder or its 'So don' subfolder."""
    found = {}
    for pat in (os.path.join(folder, "So don", "*.xlsx"), os.path.join(folder, "*.xlsx")):
        for f in sorted(glob.glob(pat)):
            b = os.path.basename(f)
            m = re.match(r"(\d{2})(?:[_\s.\-]|\.xlsx)", b)
            if not m or b.startswith("~$") or b.startswith("45"):
                continue
            found.setdefault(m.group(1), f)
    return found


def read_policies(folder, log=None):
    out = {}
    for code, f in sorted(policy_files(folder).items()):
        _say(log, "Reading %s ..." % os.path.relpath(f, folder))
        rows = read_policy_file(f)
        if rows:
            out[code] = rows
    return out


def find_cgtkt(folder, year):
    fs = [f for f in glob.glob(os.path.join(folder, "*.xlsx")) if "CG&TKT" in os.path.basename(f).upper()]
    cur = [f for f in fs if str(year) in os.path.basename(f)]
    cur = cur or [f for f in fs if str(year - 1) not in os.path.basename(f)]
    return cur[0] if cur else None


def find_prev_year(folder):
    for f in glob.glob(os.path.join(folder, "*.xlsx")):
        n = os.path.basename(f).upper()
        if "LUY KE" in n or "LŨY KẾ" in n:
            return f
    return None


def read_cgtkt(path, linemap):
    out = []
    for r in X.iter_rows(path, "CG&TKT", 0):
        if r["_r"] < 10 or r.get(2) in (None, ""):
            continue
        sp = _s(r.get(2))
        if not re.fullmatch(r"\d{6}", sp):
            continue
        dt = float(r.get(8) or 0) / 1e6
        lk = float(r.get(9) or 0) / 1e6
        out.append({"sp": sp, "ten": _s(r.get(3)), "dt": dt, "lk": lk, "line": linemap(sp)})
    return out


class RevenueError(Exception):
    pass


def _key(v):
    return re.sub(r"\s+", "", _s(v))


def read_revenue_report(path, sheet, month, cfg):
    """BC doanh thu thực hiện (file P.KD, sheet 'Chi tiet thang năm YYYY NV'): {line: [DT tháng, lũy kế]} in triệu đồng.
    Returns None when the month column is still empty (report not updated yet)."""
    rm = cfg["thuc_hien"]["report_map"]
    by_stt, by_code = rm["stt"], {_key(k): v for k, v in rm["code"].items()}
    mcol = lkcol = None
    out = {ln: [0.0, 0.0] for ln in LINES}
    total = None
    for r in X.iter_rows(path, sheet, 0):
        cells = {k: _s(v) for k, v in r.items() if k != "_r"}
        if mcol is None:
            for k, v in cells.items():
                if v.upper() == "THÁNG %d" % month:
                    mcol = k
                if re.match(r"^l[uũ][yỹ] kế", v, re.I):
                    lkcol = k
            continue
        stt, code = cells.get(0, ""), cells.get(2, "")
        num = lambda c: float(r.get(c) or 0) if isinstance(r.get(c), (int, float)) else 0.0
        if stt.upper().startswith("TỔNG"):
            total = (num(mcol), num(lkcol))
            break
        ln = by_code.get(_key(code)) if _key(code) in by_code else by_stt.get(stt)
        if ln:
            out[ln][0] += num(mcol)
            out[ln][1] += num(lkcol)
    if mcol is None or lkcol is None or total is None:
        raise RevenueError("File báo cáo doanh thu không có cột THÁNG %d / Lũy kế / dòng TỔNG CỘNG" % month)
    got = (sum(v[0] for v in out.values()), sum(v[1] for v in out.values()))
    if abs(got[0] - total[0]) > 1 or abs(got[1] - total[1]) > 1:
        raise RevenueError("File báo cáo doanh thu đổi cấu trúc: cộng các nghiệp vụ không ra dòng TỔNG CỘNG")
    if not total[0]:
        return None
    return {ln: [v[0] / 1e6, v[1] / 1e6] for ln, v in out.items()}


def read_prev_year(path, month):
    """THỰC HIỆN tab of last year's KSNV workbook: column M = DT lũy kế năm trước for that month."""
    out = {}
    for r in X.iter_rows(path, "THỰC HIỆN", 0):
        try:
            m = int(float(r.get(1)))
        except (TypeError, ValueError):
            continue
        ln, v = r.get(4), r.get(12)
        if m == month and ln in LINES and isinstance(v, (int, float)):
            out[ln] = float(v)
    return out


# ---------------------------------------------------------------- 1. CẤP ĐƠN
def build_cap_don(pol, cfg, year, month):
    nd = calendar.monthrange(year, month)[1]
    res, warn = {}, []
    for grp, rule in cfg["cap_don"].items():
        tot = [0] * nd
        if not any(part["file"] in pol for part in rule):     # a line with no policy that month is normal (07)
            warn.append("CẤP ĐƠN %s: chưa nạp danh sách đơn tháng %02d/%d" % (grp, month, year))
        for part in rule:
            rows = pol.get(part["file"])
            if rows is None:
                continue
            excl = set(part.get("exclude_sp") or [])
            per_day = {}
            for p in rows:
                if part.get("pias_only") and p.nguon != "PIAS":
                    continue
                if p.sp in excl or not p.day:
                    continue
                per_day.setdefault(p.day, set()).add(p.so)
            for d, s in per_day.items():
                if 1 <= d <= nd:
                    tot[d - 1] += len(s)
        res[grp] = tot
    return res, warn


# ---------------------------------------------------------------- 2. THỰC HIỆN
def round_to_total(vals):
    """Round each value so the parts add up to the total cut to whole triệu (how the reports read
    148,218,504,980 đ as 148.218), distributing by largest remainder."""
    base = {k: math.floor(v) for k, v in vals.items()}
    need = math.floor(sum(vals.values()) + 1e-9) - sum(base.values())
    for k in sorted(vals, key=lambda k: vals[k] - base[k], reverse=True)[:max(0, need)]:
        base[k] += 1
    return base


def renewal_fee(pol, linemap, cfg):
    fee = {}
    for code in cfg["thuc_hien"]["renew_fee_files"]:
        for p in pol.get(code, []):
            if p.loai == "Tái tục" and p.goc:
                ln = linemap(p.sp)
                fee[ln] = fee.get(ln, 0.0) + p.fee
    return fee


def build_thuc_hien(rev, pol, prev, q_prev, linemap, cfg):
    """rev: {line: [DT tháng, lũy kế]} triệu đồng, from the revenue report (or CG&TKT).
    I and N are taken straight from it (Nam 10/11/2026: lũy kế lấy theo số báo cáo, không cộng dồn);
    M = tái tục tháng trước + DT đơn tái tục trong list; L = N - M."""
    th = cfg["thuc_hien"]
    fee = renewal_fee(pol, linemap, cfg)
    Ir = round_to_total({ln: rev.get(ln, [0, 0])[0] for ln in LINES})
    Nr = round_to_total({ln: rev.get(ln, [0, 0])[1] for ln in LINES})
    rows, warn = {}, []
    for ln in LINES:
        I, N = Ir[ln], Nr[ln]
        m_prev = float(prev.get(ln, {}).get("M") or 0)
        if ln in th["all_renewal_lines"]:          # CSSK: toàn bộ là tái tục
            tt, M = I, N
        else:
            tt = round(fee.get(ln, 0.0)) if ln in th["renewal_lines"] else 0
            M = round(m_prev + tt)
        rows[ln] = dict(I=I, J=0, N=N, O=0, M_prev=round(m_prev), TT=tt, M=M, L=N - M,
                        Q=round(q_prev[ln]) if ln in q_prev else None)
    return rows, fee, warn


def cgtkt_revenue(cg):
    rev = {}
    for r in cg:
        v = rev.setdefault(r["line"], [0.0, 0.0])
        v[0] += r["dt"]
        v[1] += r["lk"]
    return rev


# ---------------------------------------------------------------- 3. Tái tục (SLTT)
def build_tai_tuc(pol, due, cfg, year, month):
    res, warn = {}, []
    for grp, rule in cfg["tai_tuc"].items():
        done, need = set(), set()
        for part in rule:
            sps = set(part.get("sp") or [])
            excl = set(part.get("exclude_sp") or [])
            for src, bag in ((pol, done), (due, need)):
                for p in src.get(part["file"], []):
                    if (sps and p.sp not in sps) or p.sp in excl:
                        continue
                    if part.get("pias_only") and p.nguon != "PIAS":
                        continue
                    if not p.goc:
                        continue
                    if bag is done:
                        if p.loai == "Tái tục":
                            bag.add(p.so)
                    elif p.end and p.end.year == year and p.end.month == month:
                        bag.add(p.so)
        res[grp] = dict(can=len(need) if due else None, da=len(done))
    return res, warn


# ---------------------------------------------------------------- run
def run(store, year, month, prev, cfg=None, log=None):
    """store: the KSNV data folder (ksnv_store layout)."""
    import ksnv_store as ST
    cfg = cfg or load_cfg()
    lm = LineMap(cfg)
    warn = []
    _say(log, "Reading the policy lists of %02d/%d ..." % (month, year))
    pol = ST.month_policies(store, year, month)
    need = sorted({p["file"] for r in cfg["tai_tuc"].values() for p in r})
    _say(log, "Looking for policies ending in %02d/%d ..." % (month, year))
    due = ST.expiring(store, year, month, need)
    cov = ST.coverage(store)

    cap, w = build_cap_don(pol, cfg, year, month); warn += w

    th = fee = rev = None
    cg_path, py_path = ST.cgtkt_file(store, year, month), ST.prevyear_file(store, year)
    rv_path = ST.revenue_file(store, year)
    src = None
    if rv_path:
        _say(log, "Reading %s ..." % os.path.basename(rv_path))
        rev = read_revenue_report(rv_path, ST.revenue_sheet(rv_path), month, cfg)
        src = rv_path if rev else None
    if rev is None and cg_path:
        _say(log, "Reading %s ..." % os.path.basename(cg_path))
        rev, src = cgtkt_revenue(read_cgtkt(cg_path, lm)), cg_path
    if rev is not None:
        q = {}
        if py_path:
            _say(log, "Reading %s ..." % os.path.basename(py_path))
            q = read_prev_year(py_path, month)
        else:
            warn.append("THỰC HIỆN: folder chưa có file BC KSNV năm %d, cột Q để trống" % (year - 1))
        th, fee, w = build_thuc_hien(rev, pol, prev, q, lm, cfg); warn += w
    else:
        warn.append("THỰC HIỆN: chưa nạp báo cáo doanh thu có số tháng %02d/%d" % (month, year))

    tt, w = build_tai_tuc(pol, due, cfg, year, month); warn += w
    return dict(cap_don=cap, thuc_hien=th, renew_fee=fee, tai_tuc=tt, warnings=warn,
                files=sorted(pol), cgtkt=cg_path, revenue=src, prev_year=py_path, coverage=cov)
