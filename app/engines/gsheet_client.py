"""Minimal Google Sheets client: OAuth desktop (loopback + PKCE) and values read/write.

Only needs `requests`. Token (with refresh_token) is cached in a JSON file so later runs are silent.
"""
import base64
import hashlib
import http.server
import json
import os
import secrets
import socket
import subprocess
import threading
import time
import urllib.parse
import webbrowser

import requests

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "openid", "email"]
API = "https://sheets.googleapis.com/v4/spreadsheets"


class AuthError(Exception):
    pass


def _open_browser(url):
    """Prefer Chrome (Nam's default for OAuth), fall back to the system browser."""
    for p in (r"C:\Program Files\Google\Chrome\Application\chrome.exe",
              r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
              os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe")):
        if os.path.exists(p):
            try:
                subprocess.Popen([p, url])
                return
            except OSError:
                pass
    webbrowser.open(url)


class GoogleAuth:
    def __init__(self, client_file, token_file):
        """client_file: path to the OAuth client JSON, or the parsed dict."""
        if isinstance(client_file, dict):
            c = client_file
        else:
            with open(client_file, encoding="utf-8") as f:
                c = json.load(f)
        c = c.get("installed") or c.get("web")
        self.cid, self.secret = c["client_id"], c["client_secret"]
        self.auth_uri, self.token_uri = c["auth_uri"], c["token_uri"]
        self.token_file = token_file
        self.tok = None
        if os.path.exists(token_file):
            try:
                with open(token_file, encoding="utf-8") as f:
                    self.tok = json.load(f)
            except (OSError, ValueError):
                self.tok = None

    # ---- state
    @property
    def email(self):
        return (self.tok or {}).get("email")

    def signed_in(self):
        return bool(self.tok and self.tok.get("refresh_token"))

    def sign_out(self):
        self.tok = None
        if os.path.exists(self.token_file):
            os.remove(self.token_file)

    def _save(self):
        os.makedirs(os.path.dirname(self.token_file) or ".", exist_ok=True)
        with open(self.token_file, "w", encoding="utf-8") as f:
            json.dump(self.tok, f)

    # ---- interactive sign-in
    def sign_in(self, timeout=300, on_url=None):
        verifier = secrets.token_urlsafe(64)
        challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).decode().rstrip("=")
        state = secrets.token_urlsafe(16)
        s = socket.socket(); s.bind(("127.0.0.1", 0)); port = s.getsockname()[1]; s.close()
        redirect = f"http://127.0.0.1:{port}"
        got = {}

        class H(http.server.BaseHTTPRequestHandler):
            def do_GET(self):
                q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
                if "code" in q or "error" in q:
                    got.update({k: v[0] for k, v in q.items()})
                    ok = "code" in q
                    msg = "Signed in. You can close this tab." if ok else "Sign-in failed: " + q["error"][0]
                    body = f"<html><body style='font-family:sans-serif;padding:40px'><h3>{msg}</h3></body></html>"
                    self.send_response(200); self.send_header("Content-Type", "text/html; charset=utf-8"); self.end_headers()
                    self.wfile.write(body.encode())
                else:
                    self.send_response(204); self.end_headers()

            def log_message(self, *a):
                pass

        srv = http.server.HTTPServer(("127.0.0.1", port), H)
        th = threading.Thread(target=srv.serve_forever, daemon=True); th.start()
        url = self.auth_uri + "?" + urllib.parse.urlencode(dict(
            client_id=self.cid, redirect_uri=redirect, response_type="code", scope=" ".join(SCOPES),
            code_challenge=challenge, code_challenge_method="S256", state=state,
            access_type="offline", prompt="consent select_account"))
        if on_url:
            on_url(url)
        _open_browser(url)
        t0 = time.time()
        while not got and time.time() - t0 < timeout:
            time.sleep(0.3)
        srv.shutdown()
        if not got:
            raise AuthError("Sign-in timed out")
        if "error" in got or got.get("state") != state:
            raise AuthError("Sign-in failed: " + got.get("error", "state mismatch"))
        r = requests.post(self.token_uri, data=dict(
            code=got["code"], client_id=self.cid, client_secret=self.secret, redirect_uri=redirect,
            grant_type="authorization_code", code_verifier=verifier), timeout=60)
        if r.status_code != 200:
            raise AuthError("Token exchange failed: " + r.text[:300])
        t = r.json()
        t["expires_at"] = time.time() + t.get("expires_in", 3600) - 60
        t["email"] = self._email_of(t["access_token"])
        self.tok = t
        self._save()
        return t["email"]

    def _email_of(self, access):
        try:
            r = requests.get("https://openidconnect.googleapis.com/v1/userinfo",
                             headers={"Authorization": "Bearer " + access}, timeout=30)
            return r.json().get("email")
        except Exception:
            return None

    def access_token(self):
        if not self.signed_in():
            raise AuthError("Not signed in to Google")
        if time.time() < self.tok.get("expires_at", 0):
            return self.tok["access_token"]
        r = requests.post(self.token_uri, data=dict(
            client_id=self.cid, client_secret=self.secret, refresh_token=self.tok["refresh_token"],
            grant_type="refresh_token"), timeout=60)
        if r.status_code != 200:
            raise AuthError("Google session expired, please sign in again (" + r.text[:200] + ")")
        t = r.json()
        self.tok["access_token"] = t["access_token"]
        self.tok["expires_at"] = time.time() + t.get("expires_in", 3600) - 60
        self._save()
        return self.tok["access_token"]


class Sheet:
    def __init__(self, auth, spreadsheet_id):
        self.auth, self.sid = auth, spreadsheet_id

    def _h(self):
        return {"Authorization": "Bearer " + self.auth.access_token()}

    def _check(self, r):
        if r.status_code == 403:
            raise AuthError(f"Tài khoản {self.auth.email} không có quyền với Sheet này (403)")
        if r.status_code >= 400:
            raise RuntimeError(f"Google Sheets API {r.status_code}: {r.text[:400]}")
        return r.json()

    def meta(self):
        return self._check(requests.get(f"{API}/{self.sid}", headers=self._h(),
                                         params={"fields": "properties.title,sheets.properties"}, timeout=60))

    def get(self, ranges, render="UNFORMATTED_VALUE"):
        r = requests.get(f"{API}/{self.sid}/values:batchGet", headers=self._h(), timeout=120,
                         params=[("ranges", x) for x in ranges] + [("valueRenderOption", render)])
        return [v.get("values", []) for v in self._check(r)["valueRanges"]]

    def update(self, data, raw=False):
        """data: list of (range, 2D values). USER_ENTERED so '=...' strings become formulas."""
        body = {"valueInputOption": "RAW" if raw else "USER_ENTERED",
                "data": [{"range": rg, "values": vals} for rg, vals in data]}
        return self._check(requests.post(f"{API}/{self.sid}/values:batchUpdate", headers=self._h(),
                                         json=body, timeout=120))

    def clear(self, ranges):
        return self._check(requests.post(f"{API}/{self.sid}/values:batchClear", headers=self._h(),
                                         json={"ranges": list(ranges)}, timeout=120))
