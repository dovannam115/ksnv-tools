# Build dist\KSNV_Report.zip to hand over. Run: app\pyembed\python.exe app\make_package.py
# Ships code + pyembed + app\data (ksnv_rules.json), with EMPTY output\ and app\dump\.
# Left out: app\dump\* (this PC's Google sign-in, token, logs, uploads), output\ (Sheet cell backups),
# app\_backup_*, __pycache__, this script, and the GitHub token in update.json (blanked).
import os, json, zipfile
from common import ROOT, APP

TOP   = "KSNV-Report"
OUT   = os.path.join(ROOT, "dist", "KSNV_Report.zip")
FILES = ["Install.bat", "KSNV Report.bat", "VERSION", "RELEASE-NOTES.txt"]
EMPTY = ["output/", "app/dump/"]


def skip_app(rel):
    parts = rel.replace("\\", "/").split("/")
    if "__pycache__" in parts: return True
    if parts[0] in ("dump",) or parts[0].startswith("_backup_"): return True
    if rel in ("make_package.py", "update.json"): return True
    if parts[-1].startswith("~$"): return True
    return False


def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    n = 0
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for f in FILES:
            z.write(os.path.join(ROOT, f), f"{TOP}/{f}"); n += 1
        for d in EMPTY:
            z.writestr(f"{TOP}/{d}", "")
        # repo + branch travel with the zip so updates work; the token goes too only when the repo is private
        try: c = json.load(open(os.path.join(APP, "update.json"), encoding="utf-8"))
        except Exception: c = {"repo": "", "branch": "main", "token": ""}
        z.writestr(f"{TOP}/app/update.json", json.dumps(c, indent=2)); n += 1
        for root, dirs, files in os.walk(APP):
            for f in files:
                full = os.path.join(root, f); rel = os.path.relpath(full, APP)
                if skip_app(rel): continue
                z.write(full, f"{TOP}/app/{rel.replace(os.sep, '/')}"); n += 1
    must = ["app/pyembed/python.exe", "app/pyembed/pythonw.exe", "app/pyembed/python312._pth",
            "app/pyembed/Lib/site-packages/requests/__init__.py",
            "app/server.py", "app/worker.py", "app/common.py", "app/ui/index.html", "app/ui/app.ico",
            "app/engines/ksnv_engine.py", "app/engines/ksnv_google.py", "app/data/ksnv_rules.json", "VERSION"]
    names = set(zipfile.ZipFile(OUT).namelist())
    miss = [m for m in must if f"{TOP}/{m}" not in names]
    if any("google_token" in x or "/dump/ui_token" in x for x in names): miss.append("LEAK: app/dump")
    if any(x.startswith(f"{TOP}/output/") and x != f"{TOP}/output/" for x in names): miss.append("LEAK: output")
    print(f"{n} files -> {OUT} ({os.path.getsize(OUT)/1e6:.1f} MB)")
    print("MISSING: " + ", ".join(miss) if miss else "Check OK: all required files inside, nothing private.")


if __name__ == "__main__":
    main()
