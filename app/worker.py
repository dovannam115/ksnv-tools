# One run of one tool, started by server.py as a child process:
#   python -u worker.py <tool> <job dir>
# Reads <job dir>\args.json, writes its files into <job dir>, ends with <job dir>\result.json.
# Lines it prints go to the run panel in the window. Markers the server reads:
#   [[P]] done|total|label     progress bar
#   [[ERROR]] text             the message the user sees when the run fails
import os, sys, json, time, shutil, traceback, threading, datetime, decimal
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common as C

try: sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
except Exception: pass

DIR = ""


def say(*a):
    print(" ".join(str(x) for x in a), flush=True)


def prog(done, total, label=""):
    print("[[P]] %d|%d|%s" % (done, total, str(label).replace("\n", " ")[:120]), flush=True)


class UserError(Exception):
    """A failure the user can fix (wrong file, missing column...): shown without a traceback."""


def jsonable(v):
    if v is None or isinstance(v, (bool, int, str)): return v
    if isinstance(v, float): return None if v != v else v
    if isinstance(v, decimal.Decimal): return float(v)
    if isinstance(v, (datetime.datetime, datetime.date)): return v.strftime("%d/%m/%Y")
    if isinstance(v, (bytes, bytearray)): return v.hex()
    return str(v)


def table(title, cols, rows, views=None, links=(), main=None, limit=2000):
    """A table for the result panel. rows = list of lists (same order as cols) or dicts.
    views = [(label, predicate on the row dict)]; links = columns holding a URL."""
    recs = [r if isinstance(r, dict) else dict(zip(cols, r)) for r in rows]
    out = {"title": title, "cols": cols, "total": len(recs), "links": list(links),
           "rows": [[jsonable(r.get(c)) for c in cols] for r in recs[:limit]]}
    if views:
        out["views"] = [{"label": lb, "rows": [i for i, r in enumerate(recs[:limit]) if f(r)]}
                        for lb, f in views]
    if main: out["main"] = [cols.index(c) for c in main if c in cols]
    return out


def stat(label, value, tone=""):
    return {"label": label, "value": value, "tone": tone}


def money(x):
    try: return "{:,.0f}".format(float(x)).replace(",", ".")
    except Exception: return str(x)



# ------------------------------------------------------------------ KSNV Report (Google Sheet)
KSNV_PARTS = {"CD": "CẤP ĐƠN", "TH": "THỰC HIỆN", "TT": "LƯU 2026", "RX": "RÀ SOÁT XCG"}


def ksnv_store_dir(a):
    d = (a.get("store") or "").strip().strip('"')
    if not d:
        raise UserError("Choose the KSNV data folder first.")
    if not os.path.isdir(os.path.dirname(d.rstrip("\\/")) or d):
        raise UserError("Folder not found: %s" % d)
    os.makedirs(d, exist_ok=True)
    return d


def ksnv_sheet():
    import ksnv_engine as E, ksnv_sheet as S, gsheet_client as G
    cfg = E.load_cfg(C.KSNV_RULES)
    auth = C.google_auth()
    if not auth.signed_in():
        raise UserError("Not signed in to Google yet. Click 'Sign in to Google' at the top of the page.")
    try:
        auth.access_token()
    except G.AuthError as e:
        raise UserError(str(e))
    return E, S, G, cfg, auth, G.Sheet(auth, cfg["sheet_id"])


def ksnv_coverage_table(store):
    import ksnv_store as ST
    cov = ST.coverage(store)
    rows = []
    for ln, months in cov.items():
        rows.append({"Line": ln, "Từ": months[0] if months else "", "Đến": months[-1] if months else "",
                     "Số tháng": len(months)})
    import glob
    cg = sorted(os.path.basename(f)[7:14] for f in glob.glob(os.path.join(store, "cgtkt", "*.xlsx")))
    ref = sorted(os.path.basename(f) for f in glob.glob(os.path.join(store, "ref", "*.xlsx")))
    rv = sorted(os.path.basename(f)[8:12] for f in glob.glob(os.path.join(store, "revenue", "*.xlsx")))
    xc = sorted(os.path.basename(f)[11:18] for f in glob.glob(os.path.join(store, "xcg", "*.xlsx")))
    notes = ["BC doanh thu: %s" % (", ".join(rv) if rv else "chưa có"), "Rà soát XCG: %s" % (", ".join(xc) if xc else "chưa có"), "CG&TKT: %s" % (", ".join(cg) if cg else "chưa có"),
             "BC KSNV năm trước: %s" % (", ".join(ref) if ref else "chưa có")]
    return table("Data folder · danh sách đơn theo line", ["Line", "Từ", "Đến", "Số tháng"], rows), notes


def run_ksnv(a):
    mode = a.get("mode") or "preview"
    store = ksnv_store_dir(a)

    if mode == "import":
        import ksnv_store as ST
        files = a.get("files") or []
        if not files:
            raise UserError("Add the files to import.")
        say("Data folder: %s" % store)
        done, skipped, written = ST.import_files(files, store, log=say)
        tb, notes = ksnv_coverage_table(store)
        imp = table("Imported", ["file", "kind", "rows", "lines", "months"], done)
        warns = ["%s: %s" % s for s in skipped]
        return {"stats": [stat("Files imported", len(done), "good"), stat("Skipped", len(skipped), "warn" if skipped else ""),
                          stat("Month files written", len(written))],
                "warns": warns, "notes": notes, "tables": [imp, tb]}

    E, S, G, cfg, auth, sh = ksnv_sheet()
    say("Google account: %s" % auth.email)
    if mode == "restore":
        S.restore(sh, a["backup"])
        return {"stats": [stat("Restored", os.path.basename(a["backup"]), "good")]}
    if mode == "write":
        plan = json.load(open(os.path.join(a["plan_dir"], "plan.json"), encoding="utf-8"))
        parts = [p for p in (a.get("parts") or []) if p in KSNV_PARTS]
        keep = tuple("'%s'!" % KSNV_PARTS[p] for p in parts)
        data = [(rg, v) for rg, v in plan["data"] if rg.startswith(keep)]
        if not data:
            raise UserError("Nothing to write. Tick at least one part.")
        say("Saving the current cells before writing ...")
        bk = S.backup(sh, data, DIR)
        say("Writing %d ranges to '%s' ..." % (len(data), plan["title"]))
        sh.update(data)
        n_cells = sum(len(v) * max(len(r) for r in v) for _, v in data)
        return {"stats": [stat("Month", "T%s/%s" % (plan["month"], plan["year"])),
                          stat("Parts written", ", ".join(KSNV_PARTS[p] for p in parts), "good"),
                          stat("Cells", n_cells)],
                "notes": ["Undo puts back the cells as they were before this write."],
                "files": [bk], "backup": bk}

    # preview: data folder + Sheet -> what would be written
    import calendar
    y, m = int(a["year"]), int(a["month"])
    title = sh.meta()["properties"]["title"]
    say("Reading the Sheet '%s' ..." % title)
    lay = S.read_layout(sh, m)
    if not lay["rows_th"]:
        raise UserError("Month %d not found in tab THỰC HIỆN." % m)
    res = E.run(store, y, m, lay["prev"], cfg, log=say)
    if not res["files"]:
        raise UserError("The data folder has no policy list for %02d/%d. Import the month's files first." % (m, y))
    nd = calendar.monthrange(y, m)[1]
    data, missing = S.build_updates(res, lay, m, nd, cfg)
    import ksnv_store as ST
    xcg = None
    xf = ST.xcg_file(store, y, m)
    if xf:
        say("Reading %s ..." % os.path.basename(xf))
        xd, xcg = S.build_xcg_update(sh, ST.read_xcg_review(xf), m, ST.xcg_key)
        data += xd
        if xcg["rows"] > xcg["filled"]:
            missing.append("RÀ SOÁT XCG: %d đơn chưa có kết quả hậu kiểm" % (xcg["rows"] - xcg["filled"]))
    json.dump({"year": y, "month": m, "title": title, "account": auth.email, "store": store, "data": data},
              open(os.path.join(DIR, "plan.json"), "w", encoding="utf-8"), ensure_ascii=False)

    tables = []
    days = ["%d" % d for d in range(1, nd + 1)]
    cd_rows = [dict({"Nghiệp vụ": g, "Tổng": sum(v)}, **{d: v[i] for i, d in enumerate(days)})
               for g, v in res["cap_don"].items()]
    tables.append(table("CẤP ĐƠN · tổng đơn + SĐBS theo ngày (hậu kiểm = tổng, sai = 0)",
                        ["Nghiệp vụ", "Tổng"] + days, cd_rows))
    th = res["thuc_hien"] or {}
    if th:
        cols = ["Nghiệp vụ", "DT tháng (I)", "Lũy kế (N)", "Tái tục trước", "Tái tục tháng", "Lũy kế tái tục (M)",
                "Cấp mới (L)", "LK năm trước (Q)"]
        recs = [{"Nghiệp vụ": ln, "DT tháng (I)": r["I"], "Lũy kế (N)": r["N"], "Tái tục trước": r["M_prev"],
                 "Tái tục tháng": r["TT"], "Lũy kế tái tục (M)": r["M"], "Cấp mới (L)": r["L"],
                 "LK năm trước (Q)": r["Q"]}
                for ln, r in th.items()]
        recs.append({"Nghiệp vụ": "Tổng", **{c: sum((x[c] or 0) for x in recs) for c in cols[1:]}})
        tables.append(table("THỰC HIỆN · triệu đồng", cols, recs))
    tt = res["tai_tuc"]
    tables.append(table("Tái tục · LƯU 2026 dòng SLTT", ["Nghiệp vụ", "Cần tái tục", "Đã tái tục"],
                        [[g, x["can"] if x["can"] is not None else "", x["da"]] for g, x in tt.items()]))
    if xcg:
        tables.append(table("RÀ SOÁT XCG · cột Q Công tác hậu kiểm - khắc phục",
                            ["Đơn trên Sheet", "Ghép được với file", "Có kết quả sau khi ghi", "Dòng trong file"],
                            [[xcg["rows"], xcg["matched"], xcg["filled"], xcg["file_rows"]]]))
    cov_tb, cov_notes = ksnv_coverage_table(store)
    tables.append(cov_tb)
    tot_i = sum(r["I"] for r in th.values()) if th else 0
    by_tab = Counter(rg.split("!")[0].strip("'") for rg, _ in data)
    return {"stats": [stat("Month", "T%d/%d" % (m, y)),
                      stat("Đơn + SĐBS", sum(sum(v) for v in res["cap_don"].values())),
                      stat("DT tháng (tr.đ)", tot_i),
                      stat("Ranges to write", len(data), "good" if not missing else "warn")],
            "warns": res["warnings"] + missing,
            "notes": ["Sheet: %s" % title, "Google: %s" % auth.email] + cov_notes +
                     [", ".join("%s %d" % kv for kv in by_tab.items())],
            "tables": tables, "plan_dir": DIR,
            "parts": [p for p, tab in KSNV_PARTS.items() if by_tab.get(tab)]}


def run_gsignin(a):
    auth = C.google_auth()
    say("A browser page opens. Sign in with the Google account that can edit the KSNV Sheet.")
    say("If Google says the app is not verified: Advanced, then Go to KSNV Report.")
    email = auth.sign_in(timeout=300)
    say("Signed in as %s." % email)
    return {"stats": [stat("Signed in", email, "good")]}



RUN = {"ksnv": run_ksnv, "ksnvw": run_ksnv, "gsignin": run_gsignin}


def main():
    global DIR
    tool, DIR = sys.argv[1], sys.argv[2]
    a = json.load(open(os.path.join(DIR, "args.json"), encoding="utf-8"))
    res = {"ok": False}
    try:
        res = RUN[tool](a) or {}
        res["ok"] = True
        rc = 0
    except UserError as e:
        print("[[ERROR]] " + str(e), flush=True); res["error"] = str(e); rc = 2
    except Exception as e:
        # errors the engines raise for the user (wrong file, no access...) are shown without a traceback
        user = type(e).__name__ in ("ValueError", "AuthError", "RevenueError", "StoreError")
        tb = traceback.format_exc()
        if not user: say(tb)
        msg = str(e) if user else "Unexpected error: %s" % e
        print("[[ERROR]] " + msg.replace("\n", " "), flush=True)
        res["error"] = msg; res["trace"] = "" if user else tb; rc = 2
    res["files"] = [{"name": os.path.basename(p), "path": p} for p in res.get("files", []) if os.path.exists(p)]
    json.dump(res, open(os.path.join(DIR, "result.json"), "w", encoding="utf-8"), ensure_ascii=False, default=jsonable)
    sys.exit(rc)


if __name__ == "__main__":
    main()
