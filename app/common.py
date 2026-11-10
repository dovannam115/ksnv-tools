# Paths and setup shared by server.py (the window) and worker.py (one run of a tool).
import os, sys

APP     = os.path.dirname(os.path.abspath(__file__))
ROOT    = os.path.dirname(APP)
ENGINES = os.path.join(APP, "engines")
DATA    = os.path.join(APP, "data")           # ksnv_rules.json
DUMP    = os.path.join(APP, "dump")           # token, log, Google sign-in
OUTPUT  = os.path.join(ROOT, "output")        # one folder per run: output\<Tool>\<date time>\

KSNV_RULES = os.path.join(DATA, "ksnv_rules.json")
GTOKEN     = os.path.join(DUMP, "google_token.json")        # Google sign-in of this PC

if ENGINES not in sys.path: sys.path.insert(0, ENGINES)

# folder name of each tool under output\
TOOLS = {"ksnv": "KSNV_Report", "ksnvw": "KSNV_Report", "gsignin": "_gsignin", "update": "_update"}


def google_auth():
    import gsheet_client, ksnv_google
    return gsheet_client.GoogleAuth(ksnv_google.client(), GTOKEN)


# Where "Update" looks for new versions; app\update.json wins when it names a repo.
REPO = {"repo": "dovannam115/ksnv-tools", "branch": "main", "token": ""}


def update_cfg():
    import json
    out = dict(REPO)
    try:
        j = json.load(open(os.path.join(APP, "update.json"), encoding="utf-8"))
        if j.get("repo"): out.update({k: v for k, v in j.items() if v or k == "token"})
    except Exception:
        pass
    return out


def http_get(url, token="", timeout=25):
    import urllib.request
    req = urllib.request.Request(url, headers={"User-Agent": "KSNVTools", "Cache-Control": "no-cache"})
    if token: req.add_header("Authorization", "Bearer " + token)
    with urllib.request.urlopen(req, timeout=timeout) as r: return r.read()
