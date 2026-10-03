"""Reads the board-meeting sheets and writes the dashboard's encrypted data files.

Local files (testing):
  python publish.py --index "sheets/Takshashila Board Meetings Index.xlsx" --out site
Google Sheets (GitHub Action):
  python publish.py --google-index <index sheet ID> --out site
  env: DASHBOARD_PASSWORD, GOOGLE_SERVICE_ACCOUNT_JSON (service-account key JSON)
Also: --embed template.html writes site/index.html with every meeting embedded (works offline and as a preview).
"""
import argparse, base64, io, json, os, re, sys
from PIL import Image
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
import sheetfmt

SALT = b"TILN-board-dashboard-v2-salt-01"   # public; must match the page
ITER = 250000

def key(pw):
    return PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=SALT, iterations=ITER).derive(pw.encode())

def enc(k, obj):
    # IV derived from the content (HMAC), so unchanged data gives unchanged files and the nightly sync commits nothing.
    import hmac, hashlib
    pt = json.dumps(obj, ensure_ascii=False, sort_keys=True).encode()
    iv = hmac.new(k, pt, hashlib.sha256).digest()[:12]
    return {"v": 1, "iv": base64.b64encode(iv).decode(), "ct": base64.b64encode(AESGCM(k).encrypt(iv, pt, None)).decode()}

# ---------------------------------------------------------------- Google access (only used with --google-index)
class Google:
    def __init__(self):
        from google.oauth2 import service_account
        from google.auth.transport.requests import AuthorizedSession
        info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
        cred = service_account.Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/drive.readonly"])
        self.s = AuthorizedSession(cred)
    def sheet_xlsx(self, file_id):
        r = self.s.get(f"https://www.googleapis.com/drive/v3/files/{file_id}/export",
                       params={"mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})
        r.raise_for_status(); return io.BytesIO(r.content)
    def find_picture(self, name):
        """Finds a picture by file name, in PICTURES_FOLDER_ID if set, else anywhere the service account can see."""
        q = "name = '" + name.replace("'", "\\'") + "' and trashed = false and mimeType contains 'image/'"
        folder = os.environ.get("PICTURES_FOLDER_ID")
        if folder: q += f" and '{folder}' in parents"
        r = self.s.get("https://www.googleapis.com/drive/v3/files", params={"q": q, "fields": "files(id,name)", "pageSize": 5,
                       "supportsAllDrives": "true", "includeItemsFromAllDrives": "true"})
        r.raise_for_status(); f = r.json().get("files", [])
        return f[0]["id"] if f else None
    def file(self, file_id):
        r = self.s.get(f"https://www.googleapis.com/drive/v3/files/{file_id}", params={"alt": "media", "supportsAllDrives": "true"})
        r.raise_for_status(); return r.content

def drive_id(s):
    m = re.search(r"/d/([A-Za-z0-9_-]{20,})", s) or re.search(r"[?&]id=([A-Za-z0-9_-]{20,})", s)
    if m: return m.group(1)
    return s if re.fullmatch(r"[A-Za-z0-9_-]{25,}", s) else None

# ---------------------------------------------------------------- pictures
def picture(ref, base_dir, g, warn, where):
    if not ref:
        warn.append(f"{where}: a Picture row has no picture yet. Insert the picture into column D (Insert → Image → Image in cell) and press Sync."); return None
    try:
        fid = drive_id(ref) if ref.startswith("http") or len(ref) > 24 else None
        name = os.path.basename(ref)
        local = [p for p in (os.path.join(base_dir, ref), os.path.join(base_dir, "images", name)) if os.path.exists(p)]
        if fid and g: raw = g.file(fid)
        elif fid: warn.append(f"{where}: Drive picture needs the Google sync; skipped."); return None
        elif local: raw = open(local[0], "rb").read()
        elif g:
            fid = g.find_picture(name)
            if not fid: warn.append(f"{where}: no picture named '{name}' in the Pictures folder; skipped."); return None
            raw = g.file(fid)
        else: warn.append(f"{where}: picture '{ref}' not found; skipped."); return None
        im = Image.open(io.BytesIO(raw)).convert("RGB"); im.thumbnail((1400, 1400))
        b = io.BytesIO(); im.save(b, "JPEG", quality=78, optimize=True)
        return "data:image/jpeg;base64," + base64.b64encode(b.getvalue()).decode()
    except Exception as e:
        warn.append(f"{where}: picture could not be read ({e}); skipped."); return None

def build(index_rows, open_sheet, base_dir, g):
    meetings, warn = [], []
    for r in index_rows:
        if not r["on"]: continue
        c, w = sheetfmt.read_meeting(open_sheet(r["sheet"]))
        w = [f"[{r['id']}] {x}" for x in w]
        if c["meeting"]["id"] != r["id"]: w.append(f"[{r['id']}] Meeting tab says ID {c['meeting']['id']}; the index ID is used.")
        c["meeting"]["id"] = r["id"]
        if r["sheet"].startswith("http"): c["meeting"]["sheet_url"] = r["sheet"]
        for t in c["tabs"].values():
            for p in t["panels"]:
                if p["type"] == "carousel":
                    keep = []
                    for s in p["slides"]:
                        d = picture(s["img"], base_dir, g, w, f"[{r['id']}] {t['name']} {p['code']}")
                        if d: keep.append({"img": d, "caption": s["caption"]})
                    p["slides"] = keep
        c["warnings"] = [x.split("] ", 1)[-1] for x in w]
        meetings.append(c); warn += w
    meetings.sort(key=lambda c: c["meeting"]["date"])
    return meetings, warn

def main():
    a = argparse.ArgumentParser()
    a.add_argument("--index"); a.add_argument("--google-index"); a.add_argument("--out", default="site")
    a.add_argument("--embed", help="page template to write out/index.html with all meetings embedded")
    a.add_argument("--light", action="store_true", help="embed no meetings; the page loads them from data/ (for GitHub Pages)")
    a.add_argument("--strict", action="store_true", help="fail if there are warnings")
    o = a.parse_args()
    pw = os.environ.get("DASHBOARD_PASSWORD")
    if not pw: sys.exit("Set DASHBOARD_PASSWORD.")
    g = None
    if o.google_index:
        g = Google(); rows = sheetfmt.read_index(g.sheet_xlsx(o.google_index)); base = "."
        open_sheet = lambda ref: g.sheet_xlsx(drive_id(ref) or ref)
    else:
        rows = sheetfmt.read_index(o.index); base = os.path.dirname(os.path.abspath(o.index))
        open_sheet = lambda ref: os.path.join(base, ref)
    meetings, warn = build(rows, open_sheet, base, g)
    k = key(pw)
    os.makedirs(os.path.join(o.out, "data"), exist_ok=True)
    for c in meetings:
        d = os.path.join(o.out, "data", c["meeting"]["id"]); os.makedirs(d, exist_ok=True)
        json.dump(enc(k, c), open(os.path.join(d, "meeting.json"), "w"))
    import hashlib
    version = hashlib.sha256(json.dumps(meetings, sort_keys=True, ensure_ascii=False).encode()).hexdigest()[:12]
    json.dump({"version": version, "meetings": [{k: c["meeting"][k] for k in ("id", "date", "window_start", "window_end")} for c in meetings]},
              open(os.path.join(o.out, "data", "manifest.json"), "w"), indent=1)
    if o.embed:
        t = open(o.embed).read().replace("/*SEED*/null", json.dumps(enc(k, {"meetings": [] if o.light else meetings,
                                                      "config": {"sync_url": os.environ.get("SYNC_URL", ""), "sync_key": os.environ.get("SYNC_KEY", "")}})))
        t = t.replace("/*SALT*/", base64.b64encode(SALT).decode()).replace("/*ITER*/", str(ITER))
        logo = os.path.join(os.path.dirname(os.path.abspath(o.embed)), "logo_mark.png")
        t = t.replace("/*LOGO*/", "data:image/png;base64," + base64.b64encode(open(logo, "rb").read()).decode())
        open(os.path.join(o.out, "index.html"), "w").write(t)
    print(f"Published {len(meetings)} meeting(s): " + ", ".join(c["meeting"]["id"] for c in meetings))
    for w in warn: print("WARNING", w)
    if o.strict and warn: sys.exit(1)

if __name__ == "__main__":
    main()
