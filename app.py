#!/usr/bin/env python3
"""Crafty Korner India: web server + admin panel.
Serves the site, and exposes /admin/ where the owner adds, edits and deletes products.
Every change is saved to products.json and the site is rebuilt at once (build.py)."""
import os, re, sys, json, time, shutil, secrets, subprocess, threading, uuid, io
from copy import copy
from openpyxl import load_workbook
from datetime import datetime
from flask import Flask, request, session, jsonify, send_from_directory, abort, redirect, make_response, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from PIL import Image, ImageOps
from urllib.parse import urlparse
from werkzeug.middleware.proxy_fix import ProxyFix
import notify
Image.MAX_IMAGE_PIXELS = 60_000_000

ROOT = os.path.dirname(os.path.abspath(__file__))
P = lambda *a: os.path.join(ROOT, *a)
if os.path.exists(P(".env")):  # simple .env loader (settings and secrets live here, never in the website files)
    for _l in open(P(".env"), encoding="utf-8"):
        _l = _l.strip()
        if _l and not _l.startswith("#") and "=" in _l:
            _k, _v = _l.split("=", 1); os.environ.setdefault(_k.strip(), _v.strip().strip('"').strip("'"))
DATA, CONF, MEDIA, XLSX, STAMP = P("products.json"), P("admin_config.json"), P("media"), P("inventory.xlsx"), P(".xlsx_stamp")
EXCEL_OK = True
CATS = ["Furniture", "Handicrafts", "Antiques", "Home Décor"]
STATUS = ["Available", "Sold", "Hidden"]
GENERATED = ["collection", "blog", "antiques", "furniture", "handicrafts", "home-decor", "custom-orders", "about", "visit", "index.html", "404.html", "sitemap.xml", "robots.txt"]
SERVE_EXT = {".html", ".css", ".js", ".png", ".jpg", ".jpeg", ".webp", ".ico", ".svg", ".xml", ".txt", ".webmanifest"}
NEVER = {"requirements.txt", "README.txt"}
lock = threading.Lock()
slug = lambda s: re.sub(r"[^a-z0-9]+", "-", str(s).lower()).strip("-")

def conf():
    if not os.path.exists(CONF):
        json.dump({"secret": secrets.token_hex(32), "hash": None}, open(CONF, "w"))
    c = json.load(open(CONF))
    if os.environ.get("ADMIN_PASSWORD") and not c["hash"]:
        c["hash"] = generate_password_hash(os.environ["ADMIN_PASSWORD"]); json.dump(c, open(CONF, "w"))
    return c

app = Flask(__name__, static_folder=None)
app.secret_key = conf()["secret"]
if os.environ.get("TRUST_PROXY") == "1": app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
app.config.update(MAX_CONTENT_LENGTH=80 * 1024 * 1024, SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax",
                  SESSION_COOKIE_SECURE=os.environ.get("COOKIE_SECURE") == "1", PERMANENT_SESSION_LIFETIME=12 * 3600)

# ---------- security headers ----------
CSP = ("default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; font-src https://fonts.gstatic.com; "
       "img-src 'self' data: blob:; connect-src 'self'; frame-ancestors 'none'; base-uri 'self'; form-action 'self'; object-src 'none'")
@app.before_request
def _https():
    if os.environ.get("FORCE_HTTPS") == "1" and request.scheme != "https": return redirect(request.url.replace("http://", "https://", 1), 301)

@app.after_request
def _headers(r):
    r.headers.update({"Content-Security-Policy": CSP, "X-Content-Type-Options": "nosniff", "X-Frame-Options": "DENY", "Referrer-Policy": "strict-origin-when-cross-origin",
                      "Permissions-Policy": "camera=(), microphone=(), geolocation=(), payment=()", "Cross-Origin-Opener-Policy": "same-origin"})
    if os.environ.get("COOKIE_SECURE") == "1": r.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
    if request.path.startswith("/api/"): r.headers["Cache-Control"] = "no-store"
    return r

# ---------- data + rebuild ----------
def load(): return json.load(open(DATA, encoding="utf-8"))

def rebuild():
    """Build into a scratch folder, then swap the finished pages in, so visitors never see a half-built site."""
    w = P("_build"); shutil.rmtree(w, ignore_errors=True); os.makedirs(w)
    for f in ("build.py", "products.json", "inventory.xlsx"): shutil.copy(P(f), w)
    r = subprocess.run([sys.executable, "build.py"], cwd=w, capture_output=True, text=True, timeout=180, env={**{k: v for k, v in os.environ.items() if k not in ("BASE_PATH", "SITE_ORIGIN", "STATIC_ENQUIRY", "OUT_DIR")}, "MEDIA_DIR": MEDIA})
    if r.returncode: raise RuntimeError((r.stderr or r.stdout)[-600:])
    for g in GENERATED:
        src, dst = os.path.join(w, g), P(g)
        if not os.path.exists(src): continue
        if os.path.isdir(dst): shutil.rmtree(dst)
        os.replace(src, dst)
    shutil.rmtree(w, ignore_errors=True)

def save(products):
    """Backup, write, rebuild. If the rebuild fails the previous data is restored."""
    os.makedirs(P("backups"), exist_ok=True)
    old = open(DATA, encoding="utf-8").read()
    open(P("backups", datetime.now().strftime("products-%Y%m%d-%H%M%S-%f.json")), "w", encoding="utf-8").write(old)
    for b in sorted(os.listdir(P("backups")))[:-80]: os.remove(P("backups", b))
    if os.path.exists(XLSX): shutil.copy(XLSX, P("backups", "inventory-" + datetime.now().strftime("%Y%m%d-%H%M%S-%f") + ".xlsx"))
    open(DATA, "w", encoding="utf-8").write(json.dumps(products, ensure_ascii=False, indent=1))
    export_xlsx(products)
    try: rebuild()
    except Exception:
        open(DATA, "w", encoding="utf-8").write(old); raise

# ---------- Excel (inventory.xlsx) stays in step with the admin panel ----------
def export_xlsx(ps):
    """Write all products into the Inventory sheet (the Blog sheet and formatting are kept)."""
    global EXCEL_OK
    try:
        wb = load_workbook(XLSX); ws = wb["Inventory"]
        if ws.cell(1, 12).value != "ID":
            ws.cell(1, 12).value = "ID"; ws.cell(1, 12)._style = copy(ws.cell(1, 11)._style); ws.column_dimensions["L"].width = 14
        for row in ws.iter_rows(min_row=2, max_row=max(ws.max_row, len(ps) + 1), max_col=12):
            for c in row: c.value = None
        for i, p in enumerate(ps, 2):
            vals = [p["name"], p["category"], p["status"], p["price"], "Yes" if p["show_price"] else "No", p["material"], p["era"], p["size"], p["description"], p["photo"], ", ".join(p["more"]), p["id"]]
            for j, v in enumerate(vals, 1):
                c = ws.cell(i, j); c.value = v or None
                if j == 12: c._style = copy(ws.cell(i, 11)._style)
        wb.save(XLSX); open(STAMP, "w").write(str(os.stat(XLSX).st_mtime_ns)); EXCEL_OK = True
    except Exception: EXCEL_OK = False
    return EXCEL_OK

def read_xlsx(path):
    wb = load_workbook(path, data_only=True)
    if "Inventory" not in wb.sheetnames: raise ValueError("This file has no “Inventory” sheet.")
    ws = wb["Inventory"]; h = [str(c.value).strip() if c.value else "" for c in ws[1]]
    old_all = load(); by_id = {p["id"]: p for p in old_all}; by_name = {p["name"].strip().lower(): p for p in old_all}; out, used = [], set()
    for n, r in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not r or not r[0]: continue
        d = dict(zip(h, r)); g = lambda k: str(d.get(k) or "").strip()
        old = by_id.get(g("ID")) or by_name.get(g("Name").lower())
        if old and old["id"] in used: old = None
        pr = d.get("Price (INR)"); pr = re.sub(r"[^\d.]", "", pr) if isinstance(pr, str) else pr
        sp = g("Show price").lower()
        try:
            p = clean({"name": g("Name"), "category": g("Category"), "status": g("Status"), "price": pr, "show_price": sp == "yes" or (sp == "" and bool(pr)),
                       "material": g("Material"), "era": g("Era / Age"), "size": g("Size"), "description": g("Description"), "photo": g("Photo file"),
                       "more": [x.strip() for x in g("More photos").split(",") if x.strip()]}, old)
        except ValueError as e: raise ValueError(f"Excel row {n} ({g('Name') or 'no name'}): {e}")
        used.add(p["id"]); out.append(p)
    return out

def excel_changed():
    return not os.path.exists(STAMP) or open(STAMP).read().strip() != str(os.stat(XLSX).st_mtime_ns)

def sync_from_excel():
    """If inventory.xlsx was edited outside the panel, pull its products in and republish. Returns (changed, error)."""
    if not os.path.exists(STAMP): export_xlsx(load()); return False, None
    if not excel_changed(): return False, None
    try: ps = read_xlsx(XLSX)
    except Exception as e: return False, str(e)
    if ps == load(): open(STAMP, "w").write(str(os.stat(XLSX).st_mtime_ns)); return False, None
    save(ps); return True, None

def urls(products):
    seen, out = set(), {}
    for p in products:
        if p["status"] == "Hidden": continue
        base = s = slug(p["name"]); n = 2
        while s in seen: s = f"{base}-{n}"; n += 1
        seen.add(s); out[p["id"]] = f"/collection/{s}/"
    return out

def trash_orphans(files, products):
    used = {p["photo"] for p in products} | {m for p in products for m in p["more"]}
    os.makedirs(P("_trash"), exist_ok=True)
    for f in files:
        if f and f not in used and os.path.exists(os.path.join(MEDIA, f)) and f.startswith("u-"):
            shutil.move(os.path.join(MEDIA, f), P("_trash", f))

def clean(d, old=None):
    t = lambda k, n: str(d.get(k) or "").strip()[:n]
    name = t("name", 120)
    if not name: raise ValueError("Please enter a product name.")
    d = {**d, "category": next((c for c in CATS if c.lower() == str(d.get("category") or "").strip().lower()), None)}
    if d["category"] not in CATS: raise ValueError("Please choose a category (Furniture, Handicrafts, Antiques or Home Décor).")
    status = str(d.get("status") or "Available").strip().title()
    if status not in STATUS: raise ValueError("Invalid status.")
    price = d.get("price")
    if isinstance(price, str): price = re.sub(r"[^\d.]", "", price)
    try: price = int(float(price)) if price not in (None, "") else None
    except (TypeError, ValueError): raise ValueError("Price must be a number.")
    if price is not None and not 0 <= price < 10**9: raise ValueError("Price is out of range.")
    ok = lambda f: bool(re.fullmatch(r"[A-Za-z0-9._-]+\.(jpg|jpeg|png|webp)", f or "")) and os.path.exists(os.path.join(MEDIA, f))
    photos = [f for f in ([d.get("photo")] + list(d.get("more") or [])) if f]
    if not photos: raise ValueError("Please add at least one photo.")
    for f in photos:
        if not ok(f): raise ValueError(f"Photo file “{f}” was not found in the media folder.")
    return dict(id=(old or {}).get("id") or "p" + uuid.uuid4().hex[:8], name=name, category=d["category"], status=status, price=price,
                show_price=bool(d.get("show_price")) and price is not None, material=t("material", 160), era=t("era", 120), size=t("size", 120),
                description=str(d.get("description") or "").strip()[:4000], photo=photos[0], more=photos[1:12])

# ---------- auth ----------
fails = {}
def authed(): return session.get("ok") is True
def guard():
    if not authed(): abort(401)
    if request.method != "GET" and not secrets.compare_digest(request.headers.get("X-CSRF", ""), session.get("csrf", "x")): abort(403)

@app.errorhandler(401)
def _401(e): return jsonify(error="Please log in."), 401
@app.errorhandler(403)
def _403(e): return jsonify(error="Session expired. Refresh the page."), 403
@app.errorhandler(413)
def _413(e): return jsonify(error="That upload is too large."), 413

@app.get("/api/session")
def api_session():
    c = conf()
    if authed() and "csrf" not in session: session["csrf"] = secrets.token_hex(16)
    return jsonify(authed=authed(), setup=c["hash"] is None, csrf=session.get("csrf"), cats=CATS, status=STATUS)

@app.post("/api/login")
def api_login():
    c, ip, now = conf(), request.remote_addr, time.time()
    n, t0 = fails.get(ip, (0, now))
    if now - t0 > 600: n, t0 = 0, now
    if n >= 6: return jsonify(error="Too many attempts. Try again in a few minutes."), 429
    pw = (request.get_json(silent=True) or {}).get("password", "")
    if c["hash"] is None:  # first run: owner chooses the password
        if len(pw) < 10: return jsonify(error="Use at least 10 characters."), 400
        c["hash"] = generate_password_hash(pw); json.dump(c, open(CONF, "w"))
    elif not check_password_hash(c["hash"], pw):
        fails[ip] = (n + 1, t0); return jsonify(error="That password is not right."), 401
    fails.pop(ip, None); session.clear(); session.permanent = True; session["ok"] = True; session["csrf"] = secrets.token_hex(16)
    return jsonify(ok=True, csrf=session["csrf"])

@app.post("/api/logout")
def api_logout(): session.clear(); return jsonify(ok=True)

@app.post("/api/password")
def api_password():
    guard(); d = request.get_json(silent=True) or {}; c = conf()
    if not check_password_hash(c["hash"], d.get("old", "")): return jsonify(error="Current password is not right."), 400
    if len(d.get("new", "")) < 10: return jsonify(error="Use at least 10 characters."), 400
    c["hash"] = generate_password_hash(d["new"]); json.dump(c, open(CONF, "w")); return jsonify(ok=True)

# ---------- products ----------
@app.get("/api/products")
def api_list():
    guard(); ps = load(); u = urls(ps)
    return jsonify([{**p, "url": u.get(p["id"])} for p in ps])

@app.post("/api/products")
def api_add():
    guard()
    with lock:
        try: ps = load(); p = clean(request.get_json(force=True)); ps.insert(0, p); save(ps)
        except ValueError as e: return jsonify(error=str(e)), 400
        except Exception as e: return jsonify(error="Could not publish: " + str(e)), 500
    return jsonify({**p, "url": urls(ps).get(p["id"]), "excel_ok": EXCEL_OK})

@app.put("/api/products/<pid>")
def api_edit(pid):
    guard()
    with lock:
        try:
            ps = load(); i = next((k for k, x in enumerate(ps) if x["id"] == pid), None)
            if i is None: return jsonify(error="Product not found."), 404
            before = [ps[i]["photo"]] + ps[i]["more"]; ps[i] = clean(request.get_json(force=True), ps[i]); save(ps); trash_orphans(before, ps)
        except ValueError as e: return jsonify(error=str(e)), 400
        except Exception as e: return jsonify(error="Could not publish: " + str(e)), 500
    return jsonify({**ps[i], "url": urls(ps).get(pid), "excel_ok": EXCEL_OK})

@app.delete("/api/products/<pid>")
def api_del(pid):
    guard()
    with lock:
        try:
            ps = load(); gone = next((x for x in ps if x["id"] == pid), None)
            if not gone: return jsonify(error="Product not found."), 404
            ps = [x for x in ps if x["id"] != pid]; save(ps); trash_orphans([gone["photo"]] + gone["more"], ps)
        except Exception as e: return jsonify(error="Could not delete: " + str(e)), 500
    return jsonify(ok=True, excel_ok=EXCEL_OK)

@app.post("/api/sync")
def api_sync():
    guard()
    with lock:
        try: changed, err = sync_from_excel()
        except Exception as e: changed, err = False, "Could not publish the Excel changes: " + str(e)
    return jsonify(changed=changed, error=err)

@app.get("/api/excel")
def api_excel_get():
    guard()
    with lock: export_xlsx(load())
    return send_file(XLSX, as_attachment=True, download_name="crafty-korner-inventory.xlsx")

@app.post("/api/excel")
def api_excel_post():
    guard(); f = request.files.get("file"); tmp = P("_upload.xlsx")
    if not f: return jsonify(error="Choose an .xlsx file."), 400
    with lock:
        try:
            f.save(tmp); ps = read_xlsx(tmp)
            os.makedirs(P("backups"), exist_ok=True); shutil.copy(XLSX, P("backups", "inventory-before-upload-" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx"))
            shutil.copy(tmp, XLSX); save(ps)
        except ValueError as e: return jsonify(error=str(e)), 400
        except Exception as e: return jsonify(error="Could not read that file: " + str(e)), 400
        finally:
            if os.path.exists(tmp): os.remove(tmp)
    return jsonify(ok=True, count=len(ps), excel_ok=EXCEL_OK)

@app.post("/api/upload")
def api_upload():
    guard(); saved = []
    for f in request.files.getlist("files")[:12]:
        try:
            im = ImageOps.exif_transpose(Image.open(f.stream))
            if im.mode in ("RGBA", "LA", "P"):
                im = im.convert("RGBA"); bg = Image.new("RGB", im.size, (255, 255, 255)); bg.paste(im, mask=im.getchannel("A")); im = bg
            im = im.convert("RGB"); im.thumbnail((1600, 1600), Image.LANCZOS)
            name = f"u-{slug(os.path.splitext(f.filename or 'photo')[0])[:30] or 'photo'}-{secrets.token_hex(3)}.jpg"
            im.save(os.path.join(MEDIA, name), "JPEG", quality=85, optimize=True, progressive=True); saved.append(name)
        except Exception: return jsonify(error=f"“{f.filename}” is not a valid image (use JPG, PNG or WebP)."), 400
    return jsonify(files=saved)

# ---------- customer enquiries: saved, then sent to WhatsApp and Gmail ----------
ENQ, ENQ_PHOTOS = P("enquiries.json"), P("enquiries", "photos")
enq_lock, enq_hits = threading.Lock(), {}
line = lambda s, n: re.sub(r"\s+", " ", re.sub(r"[\x00-\x1f\x7f]", " ", str(s or ""))).strip()[:n]
block = lambda s, n: re.sub(r"[\x00-\x08\x0b-\x1f\x7f]", "", str(s or "").replace("\r\n", "\n")).strip()[:n]
def enq_load(): return json.load(open(ENQ, encoding="utf-8")) if os.path.exists(ENQ) else []
def enq_save(L):
    fd = os.open(ENQ, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)   # customer details: owner-only file
    with os.fdopen(fd, "w", encoding="utf-8") as fh: fh.write(json.dumps(L, ensure_ascii=False, indent=1))
    os.chmod(ENQ, 0o600)

def deliver(eid):
    with enq_lock: e = next((x for x in enq_load() if x["id"] == eid), None)
    if not e: return
    files = [os.path.join(ENQ_PHOTOS, f) for f in e["photos"]]; res = {}
    for k, fn in (("whatsapp", notify.send_whatsapp), ("email", notify.send_email)):
        try: res[k] = fn(e, files)
        except Exception as ex: res[k] = "failed: " + str(ex)[:200]
    with enq_lock:
        L = enq_load()
        for x in L:
            if x["id"] == eid: x["delivery"] = res
        enq_save(L)

@app.post("/api/enquiry")
def api_enquiry():
    if (request.content_length or 0) > 30 * 1024 * 1024: return jsonify(error="Photos are too large. Please send fewer or smaller photos."), 413
    o = request.headers.get("Origin")
    if o and urlparse(o).netloc != request.host: return jsonify(error="Not allowed."), 403
    f, now = request.form, time.time()
    if f.get("website"): return jsonify(ok=True)  # honeypot: bots fill this hidden field
    try: age = now - int(f.get("ts", 0)) / 1000
    except ValueError: age = 0
    if 0 <= age < 3: return jsonify(error="Please take a moment to fill the form and try again."), 429
    hits = [t for t in enq_hits.get(request.remote_addr, []) if now - t < 3600]
    if len(hits) >= 6 or sum(len(v) for v in enq_hits.values()) > 400: return jsonify(error="Too many messages from this device. Please WhatsApp or call us instead."), 429
    name, phone, email = line(f.get("name"), 80), line(f.get("phone"), 20), line(f.get("email"), 120)
    msg, extra = block(f.get("message"), 2000), line(f.get("extra"), 200)
    digits = re.sub(r"\D", "", phone)
    if len(name) < 2: return jsonify(error="Please enter your name."), 400
    if not 8 <= len(digits) <= 15: return jsonify(error="Please enter a valid phone number."), 400
    if email and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email): return jsonify(error="That email address does not look right."), 400
    if len(msg) < 3: return jsonify(error="Please write a short message."), 400
    kind = f.get("kind") if f.get("kind") in ("custom order", "product enquiry", "general query") else "general query"
    if extra: msg += "\nWood and budget: " + extra
    eid, os_ = "e" + uuid.uuid4().hex[:10], []
    os.makedirs(ENQ_PHOTOS, exist_ok=True)
    for n, up in enumerate(request.files.getlist("photos")[:4]):
        if not up.filename: continue
        try:
            im = ImageOps.exif_transpose(Image.open(up.stream)).convert("RGB"); im.thumbnail((1600, 1600), Image.LANCZOS)
            fn = f"{eid}-{n + 1}.jpg"; im.save(os.path.join(ENQ_PHOTOS, fn), "JPEG", quality=85); os_.append(fn)  # re-encoded: strips metadata and anything hidden in the file
        except Exception:
            for fn in os_:
                try: os.remove(os.path.join(ENQ_PHOTOS, fn))
                except OSError: pass
            return jsonify(error=f"“{line(up.filename, 40)}” is not a valid photo (use JPG, PNG or WebP)."), 400
    rec = dict(id=eid, time=datetime.now().strftime("%d %b %Y, %I:%M %p"), kind=kind, name=name, phone=phone, email=email, message=msg,
               product=line(f.get("product"), 120), product_url=f.get("product_url") if re.fullmatch(r"/collection/[a-z0-9-]+/", f.get("product_url") or "") else "", photos=os_, delivery={})
    with enq_lock: L = enq_load(); L.insert(0, rec); enq_save(L[:1000])
    enq_hits.setdefault(request.remote_addr, []).append(now); enq_hits[request.remote_addr] = hits + [now]
    threading.Thread(target=deliver, args=(eid,), daemon=True).start()
    return jsonify(ok=True)

@app.get("/api/enquiries")
def api_enq_list(): guard(); return jsonify(enq_load())

@app.get("/api/enquiry-photo/<name>")
def api_enq_photo(name):
    guard()
    if not re.fullmatch(r"e[0-9a-f]{10}-[1-4]\.jpg", name): abort(404)
    return send_from_directory(ENQ_PHOTOS, name)

@app.post("/api/enquiries/<eid>/resend")
def api_enq_resend(eid): guard(); deliver(eid); return jsonify(ok=True, enquiries=enq_load())

@app.delete("/api/enquiries/<eid>")
def api_enq_del(eid):
    guard()
    with enq_lock:
        L = enq_load(); gone = next((x for x in L if x["id"] == eid), None)
        if gone:
            for f in gone["photos"]:
                try: os.remove(os.path.join(ENQ_PHOTOS, f))
                except OSError: pass
            enq_save([x for x in L if x["id"] != eid])
    return jsonify(ok=True)

# ---------- serving the site ----------
def send(path, **kw):
    r = make_response(send_from_directory(ROOT, path, **kw))
    if path.endswith((".html", ".xml")) or path.startswith("admin") or path in ("style.css", "script.js"): r.headers["Cache-Control"] = "no-cache"
    elif path.startswith("media/"): r.headers["Cache-Control"] = "public, max-age=86400"
    if path.startswith("admin"): r.headers["X-Robots-Tag"] = "noindex"
    return r

@app.get("/", defaults={"p": ""})
@app.get("/<path:p>")
def site(p):
    full = os.path.realpath(P(p))
    if not full.startswith(ROOT + os.sep) and full != ROOT: abort(404)
    rel = os.path.relpath(full, ROOT).replace(os.sep, "/"); rel = "" if rel == "." else rel
    if os.path.isdir(full):
        if p and not p.endswith("/"): return redirect("/" + p + "/", 301)
        rel = (rel + "/index.html").lstrip("/")
    ext = os.path.splitext(rel)[1].lower()
    top = rel.split("/")[0]
    if rel in NEVER or ext not in SERVE_EXT or top.startswith(("_", ".", "backups", "enquiries")) or not os.path.isfile(P(rel)):
        return (send("404.html"), 404)
    return send(rel)

if __name__ == "__main__":
    conf(); port = int(os.environ.get("PORT", 8000))
    try: print("Excel sync:", sync_from_excel())
    except Exception as e: print("Excel sync skipped:", e)
    print(f"Crafty Korner India running on http://localhost:{port}   (admin: /admin/)")
    try:
        from waitress import serve; serve(app, host=os.environ.get("HOST", "0.0.0.0"), port=port, max_request_body_size=80 * 1024 * 1024)
    except ImportError: app.run(host="0.0.0.0", port=port)
