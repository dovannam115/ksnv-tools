"""Locate target cells on the KSNV Google Sheet and turn engine results into batch updates."""
import json
import os
import time

from ksnv_engine import LINES

SPREADSHEET_ID = "1yiZFxK5shtLesEiwSB5edJW5Zrr2f3nYeKHRdCnpjX8"
TAB_TH, TAB_CD, TAB_LUU, TAB_XCG = "THỰC HIỆN", "CẤP ĐƠN", "LƯU 2026", "RÀ SOÁT XCG"
CD_LABELS = {"tong": "Tổng đơn đã cấp và đơn SĐBS", "hk": "Tổng số đơn/SĐBS đã hậu kiểm",
             "sai": "Số đơn cấp sai/ lập SĐBS sai"}
CD_HAS_HK = {"BH HÀNG HẢI", "BH PHI HÀNG HẢI"}
DAY_COLS = [chr(c) for c in range(ord("I"), ord("Z") + 1)] + ["A" + chr(c) for c in range(ord("A"), ord("M") + 1)]  # I..AM
TT_GROUPS = ["VCX Ô TÔ", "BH CON NGƯỜI", "BH HÀNG HẢI", "BH PHI HÀNG HẢI"]


def q(tab):
    return "'" + tab + "'"


def _num(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def read_layout(sh, month):
    th, cd, luu = sh.get([f"{q(TAB_TH)}!B1:Q400", f"{q(TAB_CD)}!C1:C250", f"{q(TAB_LUU)}!A1:L150"])
    rows_th, prev = {}, {}
    prev_rows = {}
    for i, r in enumerate(th, 1):
        r = r + [""] * (16 - len(r))
        m, line = _num(r[0]), r[3]
        if line in LINES and m == month:
            rows_th[line] = i
        if line in LINES and m == month - 1:
            prev[line] = {"N": r[12] if r[12] != "" else 0, "M": r[11] if r[11] != "" else 0}
            prev_rows[line] = i
    rows_cd = {}
    for i, r in enumerate(cd, 1):
        if r:
            rows_cd[str(r[0])] = i
    sltt = dgtt = None
    dgtt_cur = []
    for i, r in enumerate(luu, 1):
        r = r + [""] * (12 - len(r))
        if str(r[2]).strip().upper() == f"THÁNG {month}":
            if r[0] == "SLTT":
                sltt = i
            elif r[0] == "ĐGTT":
                dgtt, dgtt_cur = i, [r[3], r[5], r[7], r[9]]
    return dict(rows_th=rows_th, prev=prev, prev_rows=prev_rows, rows_cd=rows_cd,
                sltt=sltt, dgtt=dgtt, dgtt_cur=dgtt_cur)


def build_updates(res, lay, month, ndays, cfg):
    data, missing = [], []
    # ---- CẤP ĐƠN
    for grp, vals in res["cap_don"].items():
        for key, label in CD_LABELS.items():
            if key == "hk" and grp not in CD_HAS_HK:
                continue
            row = lay["rows_cd"].get(f"{month}{grp}{label}")
            if not row:
                missing.append(f"CẤP ĐƠN: không thấy dòng {month}/{grp}/{label}")
                continue
            v = vals if key in ("tong", "hk") else [0] * ndays
            data.append((f"{q(TAB_CD)}!{DAY_COLS[0]}{row}:{DAY_COLS[ndays - 1]}{row}", [v]))
    # ---- THỰC HIỆN: I J K L M N O P Q
    th = res.get("thuc_hien") or {}
    allr = set(cfg["thuc_hien"]["all_renewal_lines"])
    for line, r in th.items():
        row, prow = lay["rows_th"].get(line), lay["prev_rows"].get(line)
        if not row:
            missing.append(f"THỰC HIỆN: không thấy dòng tháng {month} {line}")
            continue
        N = r["N"]                                    # lũy kế lấy thẳng từ báo cáo doanh thu
        if line in allr:
            M = f"=N{row}"                            # CSSK: toàn bộ là tái tục
        elif prow:
            M = f"=M{prow}+{r['TT']}"                 # tái tục trước + tái tục tháng
        else:
            M = r["M"]
        Q = r["Q"] if r["Q"] is not None else ""
        data.append((f"{q(TAB_TH)}!I{row}:Q{row}",
                     [[r["I"], 0, f"=I{row}+J{row}", f"=N{row}-M{row}", M, N, 0, f"=N{row}+O{row}", Q]]))
    # ---- LƯU 2026: SLTT (+ ĐGTT draft when empty)
    tt = res.get("tai_tuc") or {}
    if lay["sltt"]:
        vals = []
        for g in TT_GROUPS:
            x = tt.get(g, {})
            vals += ["" if x.get("can") is None else x["can"], x.get("da", "")]
        if all(tt.get(g, {}).get("can") is None for g in TT_GROUPS):
            for k, g in enumerate(TT_GROUPS):     # only "Đã TT" cells: E, G, I, K
                col = "EGIK"[k]
                data.append((f"{q(TAB_LUU)}!{col}{lay['sltt']}", [[tt.get(g, {}).get("da", 0)]]))
        else:
            data.append((f"{q(TAB_LUU)}!D{lay['sltt']}:K{lay['sltt']}", [vals]))
    else:
        missing.append(f"LƯU 2026: không thấy dòng SLTT tháng {month}")
    if lay["dgtt"] and not any(str(x).strip() for x in lay["dgtt_cur"]):
        names = ["BH VCX Ô tô", "BH Con người", "BH Hàng hải", "BH Phi Hàng hải"]
        for k, g in enumerate(TT_GROUPS):
            x = tt.get(g, {})
            da, can = x.get("da", 0), x.get("can")
            if not da and not can:
                txt = f"- {names[k]}: Trong tháng không có đơn tái tục"
            elif can:
                txt = f"- {names[k]}: Tái tục thành công {da}/{can} đơn đến kỳ"
            else:
                txt = f"- {names[k]}: Tái tục thành công {da} đơn"
            data.append((f"{q(TAB_LUU)}!{'DFHJ'[k]}{lay['dgtt']}", [[txt]]))
    return data, missing


def build_xcg_update(sh, review, month, key):
    """Column Q of tab RÀ SOÁT XCG from the review file. Rows come from a FILTER formula, so they are
    matched by key, not by position. A row the file leaves blank keeps what the Sheet already has.
    Returns (data, stats)."""
    v = sh.get([f"{q(TAB_XCG)}!A15:Q6000"])[0]
    out, matched, filled, other_month = [], 0, 0, 0
    for r in v:
        r = r + [""] * (17 - len(r))
        if not str(r[2]).strip():
            break
        if _num(r[0]) != month:
            other_month += 1
        txt = review.get(key(r[2], r[3], r[15]))
        if txt is not None:
            matched += 1
        new = txt or r[16]
        filled += bool(str(new).strip())
        out.append([new])
    st = dict(rows=len(out), matched=matched, filled=filled, unmatched=len(out) - matched,
              file_rows=len(review), other_month=other_month)
    if not out:
        return [], st
    return [(f"{q(TAB_XCG)}!Q15:Q{14 + len(out)}", out)], st


def backup(sh, data, folder):
    os.makedirs(folder, exist_ok=True)
    old = sh.get([rg for rg, _ in data], render="FORMULA")
    path = os.path.join(folder, f"backup_{time.strftime('%Y%m%d_%H%M%S')}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump([{"range": rg, "values": v} for (rg, _), v in zip(data, old)], f, ensure_ascii=False, indent=1)
    return path


def restore(sh, path):
    with open(path, encoding="utf-8") as f:
        items = json.load(f)
    sh.clear([x["range"] for x in items])
    full = [(x["range"], x["values"]) for x in items if x["values"]]
    if full:
        sh.update(full)
