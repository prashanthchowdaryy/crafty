#!/usr/bin/env python3
"""Reads products.json (managed from the /admin panel) and inventory.xlsx (Blog sheet) and rebuilds the whole website (pages, product pages, blog, sitemap).
Run with:  python build.py                    (local / backend hosting, site at "/", writes into this folder; or double-click update.bat)
           python build.py --github-pages     (static GitHub Pages version at /crafty/, written to the docs/ folder; or double-click build-github-pages.bat)
Settings can also come from environment variables: BASE_PATH, SITE_ORIGIN, STATIC_ENQUIRY, OUT_DIR (see the GITHUB PAGES block below)."""
import os, re, sys, html, json, shutil
from datetime import date, datetime
from urllib.parse import quote
from openpyxl import load_workbook

# ---------------- EDIT THESE ----------------
SITE = "https://www.virasat.in"          # your real domain, no slash at the end (used by the normal/local build)
NAME = "Crafty Korner India"
WA = "918439136097"                      # country code + number, no + or spaces
PHONE = "+91 84391 36097"
PHONE2 = "+91 83743 35934"
EMAIL = "prashanthchowdaryy083@gmail.com"
STREET = "Shilparamam Arts & Crafts Village, Madhapur"
PIN = "500081"
HOURS = "10:30 am to 8:00 pm, every day"
INSTAGRAM = ""   # paste your full Instagram link; leave empty to hide the icon
FACEBOOK = ""    # full Facebook link; empty = hidden
YOUTUBE = ""     # full YouTube link; empty = hidden
# --------------------------------------------
# ---------------- GITHUB PAGES / DEPLOYMENT SETTINGS ----------------
# A normal build (python build.py) is unchanged: site at "/", enquiry forms post to the Flask API, files are written into this folder.
# The GitHub Pages build (python build.py --github-pages) is a preset for the four settings below.
GH_PRESET = dict(BASE_PATH="/crafty/", SITE_ORIGIN="https://prashanthchowdaryy.github.io", STATIC_ENQUIRY="1", OUT_DIR="docs")
if "--github-pages" in sys.argv:
    for _k, _v in GH_PRESET.items(): os.environ.setdefault(_k, _v)   # an environment variable you set yourself still wins
BASE = "/" + os.environ.get("BASE_PATH", "/").strip().strip("/") + "/"    # always starts and ends with "/": "/" or "/crafty/"
BASE = "/" if BASE == "//" else BASE
SITE = (os.environ.get("SITE_ORIGIN", "").strip() or SITE).rstrip("/") + BASE.rstrip("/")  # canonical / Open Graph / sitemap base, e.g. https://prashanthchowdaryy.github.io/crafty
STATIC_ENQUIRY = os.environ.get("STATIC_ENQUIRY", "0").strip() == "1"   # True on GitHub Pages: forms open WhatsApp instead of calling /api/enquiry
OUT = os.environ.get("OUT_DIR", ".").strip() or "."               # where the finished site is written ("." = this folder)
MEDIA_DIR = os.environ.get("MEDIA_DIR", "media")
# ---------------------------------------------------------------------
GLOBAL_KW = ("Crafty Korner India, handicrafts Hyderabad, wooden handicrafts, Shilparamam handicrafts, handmade wooden furniture Hyderabad, "
  "carved wooden furniture, Indian handicrafts online, wooden home decor, antique shop Hyderabad, antique furniture Hyderabad, "
  "inlay work furniture, wooden serving trays, wooden stools, wooden wall clocks, console table, wooden gifts, custom wooden furniture India, "
  "Charminar showpiece, Indian home decor, handmade gifts India, handicraft shop Madhapur, Shilparamam Hyderabad")

e = html.escape
slug = lambda s: re.sub(r"[^a-z0-9]+", "-", str(s).lower()).strip("-")
wa = lambda m: f"https://wa.me/{WA}?text=" + quote(m)
trunc = lambda s, n=155: s if len(s) <= n else s[:n - 1].rsplit(" ", 1)[0] + "…"
def inr(n):
    s = str(int(n)); h, t = s[:-3], s[-3:]
    return "₹" + (re.sub(r"\B(?=(\d{2})+$)", ",", h) + "," + t if h else t)

CATS = {
 "Furniture": ("furniture", "Handmade wooden furniture in Hyderabad: stools, side tables and more",
   "Handmade wooden stools, side tables and tables with carved and inlay tops, from our shop at Shilparamam. Every piece can also be made in your size."),
 "Handicrafts": ("handicrafts", "Indian handicrafts and wooden gifts in Hyderabad",
   "Wooden trays, lidded boxes, goblets, carved panels and gifts, made by hand and chosen one by one at our handicraft shop in Shilparamam."),
 "Antiques": ("antiques", "Antique and heritage furniture in Hyderabad",
   "Heritage-style furniture such as console tables, settees and a canopy jhoola, from our shop at Shilparamam. Ask us about the age and history of each piece."),
 "Home Décor": ("home-decor", "Home décor and wall decor in Hyderabad",
   "Wall clocks, mirror frames, wall brackets, mosaic lamps and showpieces, including a Charminar glass model, to bring Indian craft into your home."),
}
CAT_ORDER = list(CATS)

# ---------- read the Excel file ----------
wb = load_workbook("inventory.xlsx", data_only=True)
def rows(name):
    ws = wb[name]; h = [str(c.value).strip() if c.value else "" for c in ws[1]]
    return [dict(zip(h, r)) for r in ws.iter_rows(min_row=2, values_only=True) if r and r[0]]

items, seen = [], set()
for r in json.load(open("products.json", encoding="utf-8")):
    st = str(r.get("status") or "Available").strip().title()
    if st == "Hidden": continue
    base = s = slug(r["name"]); n = 2
    while s in seen: s = f"{base}-{n}"; n += 1
    seen.add(s)
    price = r.get("price") if r.get("show_price") and r.get("price") else None
    it = dict(name=str(r["name"]).strip(), slug=s, cat=r.get("category") or "Antiques", sold=(st == "Sold"),
              price=price, mat=str(r.get("material") or ""), era=str(r.get("era") or ""),
              size=str(r.get("size") or ""), desc=str(r.get("description") or ""), img=str(r.get("photo") or ""), more=[x for x in (r.get("more") or []) if x])
    if not os.path.exists(os.path.join(MEDIA_DIR, it["img"])): print("WARNING: photo not found in media/:", it["img"], "for", it["name"])
    items.append(it)
items.sort(key=lambda i: i["sold"])

posts = []
for r in rows("Blog"):
    d = r.get("Date"); d = d.date() if isinstance(d, datetime) else (d if isinstance(d, date) else date.today())
    posts.append(dict(title=str(r["Title"]).strip(), slug=slug(r["Title"]), date=d,
                      summary=str(r.get("Summary") or ""), body=str(r.get("Body") or ""),
                      img=str(r.get("Image") or "").strip(), tag=str(r.get("Tag") or "").strip(),
                      mins=max(1, -(-len(str(r.get("Body") or "").split()) // 200))))
posts.sort(key=lambda p: p["date"], reverse=True)

def md(t):
    out = []
    for b in re.split(r"\n\s*\n", t.strip()):
        b = b.strip()
        out.append(f"<h2>{e(b[3:])}</h2>" if b.startswith("## ")
                   else "<p>" + re.sub(r"\[(.+?)\]\((.+?)\)", r'<a href="\2">\1</a>', e(b)) + "</p>")
    return "\n".join(out)

# ---------- page shell ----------
URLS = []
def ico(d, cls="ico"): return f'<svg class="{cls}" viewBox="0 0 24 24" aria-hidden="true" focusable="false">{d}</svg>'
I = dict(
 pin='<path d="M12 21s-7-6.200-7-11a7 7 0 1 1 14 0c0 4.800-7 11-7 11z"/><circle cx="12" cy="10" r="2.500"/>',
 truck='<path d="M2 6.500h11v9.500H2zM13 9.500h4l3 3.200V16h-7"/><circle cx="6.500" cy="17.500" r="1.800"/><circle cx="16.500" cy="17.500" r="1.800"/>',
 tools='<path d="M14.500 5.500l4 4M3.500 20.500l9-9M12 7l5 5-2 2-5-5zM15.500 4l2-1.500 3.500 3.500L19.500 8"/><path d="M5 5l4 4M3 8l5-5"/>',
 insta='<rect x="3.500" y="3.500" width="17" height="17" rx="5"/><circle cx="12" cy="12" r="4"/><circle cx="17" cy="7" r=".8"/>',
 fb='<path d="M14 8.500h2.500V5H14c-2 0-3.500 1.500-3.500 3.500V11H8v3.500h2.500V21H14v-6.500h2.500L17 11h-3V8.800c0-.2.100-.3.300-.3z"/>',
 yt='<rect x="2.500" y="5.500" width="19" height="13" rx="4"/><path d="M10.500 9.500v5l4.500-2.500z"/>',
 search='<circle cx="10.500" cy="10.500" r="6.500"/><path d="M15.500 15.500L21 21"/>',
 chat='<path d="M4 5h16v11H9.500L5 20v-4H4z"/><path d="M8 9.500h8M8 12.500h5"/>',
 hand='<path d="M3 14.500h3.500l3 1.500h4l5-2.500a1.400 1.400 0 0 1 1.500 2.300L14 19.500H6.500L3 18z"/><path d="M12 12.500c-3-2-4-3.500-4-5a2.500 2.500 0 0 1 4-1.500 2.500 2.500 0 0 1 4 1.500c0 1.500-1 3-4 5z"/>',
 leaf='<path d="M5 19C5 10 10 5 20 4c0 10-5 15-14 15z"/><path d="M5 19c3-5 6-8 10-10"/>',
 arrow='<path d="M4 12h15M13 6l6 6-6 6"/>',
 cal='<rect x="4" y="5.500" width="16" height="14" rx="2"/><path d="M4 10h16M8.500 3.500v4M15.500 3.500v4"/>',
 clock='<circle cx="12" cy="12" r="8"/><path d="M12 7.500V12l3 2"/>',
 down='<path d="M6 9.500l6 6 6-6"/>',
)
SOCIAL = "".join(f'<a href="{e(u)}" target="_blank" rel="noopener" aria-label="{n}">{ico(I[k])}</a>' for u, n, k in ((INSTAGRAM, "Instagram", "insta"), (FACEBOOK, "Facebook", "fb"), (YOUTUBE, "YouTube", "yt")) if u)
def nav_html(path):
    cur = lambda ok: ' aria-current="page"' if ok else ""
    cs_all = [v[0] for v in CATS.values()]
    coll_on = path.startswith("/collection/") or any(path.startswith(f"/{c}/") for c in cs_all)
    menu = '<a href="/collection/">All collections</a>' + "".join(f'<a href="/{v[0]}/">{e(k)}</a>' for k, v in CATS.items())
    return (f'<a href="/"{cur(path == "/")}>Home</a>'
            f'<div class="dd"><a href="/collection/"{cur(coll_on)}>Collections</a><button type="button" class="dd-btn" aria-label="Show collections" aria-expanded="false">{ico(I["down"])}</button><div class="dd-menu">{menu}</div></div>'
            f'<a href="/custom-orders/"{cur(path.startswith("/custom-orders/"))}>Custom Orders</a><a href="/about/"{cur(path.startswith("/about/"))}>Our Story</a>'
            f'<a href="/blog/"{cur(path.startswith("/blog/"))}>Blog</a><a href="/visit/"{cur(path.startswith("/visit/"))}>Visit Us</a>')
ROOT_ATTR = re.compile(r'\b(href|src|action|data-src)="/(?!/)')
def with_base(content):
    """Every page leaves through here: root-relative links and assets ("/style.css", "/collection/") get the base path ("/crafty/").
    External links (https://, tel:, mailto:, wa.me) never start with a single "/", so they are left alone. With BASE "/" nothing changes."""
    return content if BASE == "/" else ROOT_ATTR.sub(lambda m: f'{m.group(1)}="{BASE}', content)
def out(path, content):
    d = os.path.join(OUT, path.strip("/")); os.makedirs(d, exist_ok=True)
    open(os.path.join(d, "index.html"), "w", encoding="utf-8").write(with_base(content))

IMGS = {}
def fit_title(t, n=68):
    """Google cuts titles near 60 characters: drop middle parts first, then shorten the brand part."""
    parts = t.split(" | ")
    while len(parts) > 2 and len(" | ".join(parts)) > n: parts.pop(1)
    if len(" | ".join(parts)) > n and len(parts) == 2: parts[1] = NAME
    return " | ".join(parts)

def page(path, title, desc, body, schema=(), img="/media/og-virasat.jpg", og="website", index=True, kw="", images=()):
    url = SITE + path
    title = fit_title(title)
    if index: URLS.append(path)
    if images: IMGS[path] = list(images)
    kws = e(", ".join(x for x in (kw, GLOBAL_KW) if x))
    ld = "".join(f'<script type="application/ld+json">{json.dumps(s, ensure_ascii=False).replace('</', '<\\/')}</script>\n' for s in schema)
    nav = nav_html(path)
    robots = "index,follow,max-image-preview:large" if index else "noindex"
    doc = f'''<!DOCTYPE html>
<html lang="en-IN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{e(title)}</title>
<meta name="description" content="{e(trunc(desc))}">
<meta name="robots" content="{robots}">
<meta name="keywords" content="{kws}">
<meta name="author" content="{NAME}">
<meta name="geo.region" content="IN-TG"><meta name="geo.placename" content="Hyderabad"><meta name="theme-color" content="#4A2512">
<link rel="icon" href="/favicon.ico" sizes="any"><link rel="icon" type="image/png" sizes="32x32" href="/favicon-32.png"><link rel="icon" type="image/png" sizes="192x192" href="/icon-192.png">
<link rel="apple-touch-icon" href="/apple-touch-icon.png"><link rel="manifest" href="/site.webmanifest">
<link rel="alternate" hreflang="en-IN" href="{url}">
<link rel="canonical" href="{url}">
<meta property="og:type" content="{og}"><meta property="og:site_name" content="{NAME}">
<meta property="og:title" content="{e(title)}"><meta property="og:description" content="{e(trunc(desc))}">
<meta property="og:url" content="{url}"><meta property="og:image" content="{SITE}{img}"><meta property="og:locale" content="en_IN">
<meta property="og:image:alt" content="{NAME} handmade wooden handicrafts and antiques, Hyderabad"><meta name="twitter:card" content="summary_large_image">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Playfair+Display:wght@500;600;700&family=Cormorant+Garamond:ital,wght@0,600;0,700;1,500;1,600&display=swap" rel="stylesheet">
<link rel="stylesheet" href="/style.css">
{ld}</head>
<body>
<div class="topbar"><div class="topbar-in"><ul><li>{ico(I["pin"])}Shilparamam, Hyderabad</li><li>{ico(I["truck"])}Pan India Delivery</li><li>{ico(I["tools"])}Custom Orders Available</li></ul><div class="social">{SOCIAL}</div></div></div>
<header class="top">
  <a class="brand" href="/" aria-label="{NAME} home"><img class="logo" src="/logo-mark.png" alt="{NAME} logo" width="67" height="58"><span class="brand-text"><span class="brand-en">{NAME}</span><span class="brand-tag">Handcrafts · Furniture · Antiques</span></span></a>
  <span class="vsep" aria-hidden="true"></span>
  <nav id="nav" aria-label="Main">{nav}</nav>
  <div class="hicons"><button type="button" class="search-btn" aria-label="Search" aria-expanded="false">{ico(I["search"])}</button><a class="chat-btn" href="{wa("Hello Crafty Korner India, I found your website and would like to know more.")}" aria-label="Chat on WhatsApp">{ico(I["chat"])}</a></div>
  <button class="menu-btn" aria-label="Menu" aria-expanded="false"><span></span><span></span><span></span></button>
</header>
<div class="searchbar" hidden><form action="/collection/" method="get" role="search"><input type="search" name="q" placeholder="Search furniture, trays, clocks, lamps…" aria-label="Search the collection" autocomplete="off"><button class="btn gold2" type="submit">Search</button></form></div>
<main>
{body}
</main>
<footer>
  <div class="border-band" aria-hidden="true"></div>
  <img class="flogo" src="/logo-mark.png" alt="{NAME} logo" width="98" height="84">
  <nav class="fnav" aria-label="Footer">{"".join(f'<a href="/{v[0]}/">{e(k)}</a>' for k, v in CATS.items())}<a href="/custom-orders/">Custom orders</a><a href="/blog/">Blog</a><a href="/visit/">Visit us</a></nav>
  <p>{NAME}, {e(STREET)}, Hyderabad {PIN}.</p>
  <p><a href="tel:{PHONE.replace(" ","")}">{e(PHONE)}</a> · <a href="tel:{PHONE2.replace(" ","")}">{e(PHONE2)}</a> · <a href="mailto:{EMAIL}">{EMAIL}</a></p>
  <p>© {date.today().year} {NAME}. Handcrafts, furniture and antiques.</p>
</footer>
<a class="wa" href="{wa("Hello Crafty Korner India, I found your website and would like to know more.")}" aria-label="Chat on WhatsApp">WhatsApp</a>
<script src="/script.js" data-wa="{WA}" data-base="{BASE}" data-enquiry="{"whatsapp" if STATIC_ENQUIRY else "api"}" data-site="{SITE}" defer></script>
</body>
</html>'''
    out(path, doc)

def enq_form(kind, product="", url="", msg="", extra=False, fid="", ph="Tell us what you are looking for", btn="Send enquiry"):
    ex = '<label>Wood and budget <small>(optional)</small><input name="extra" maxlength="200" placeholder="e.g. teak, around ₹25,000"></label>' if extra else ""
    idattr = f' id="{fid}"' if fid else ""
    return (f'<form class="enq"{idattr} data-kind="{e(kind)}" data-product="{e(product)}" data-url="{e(url)}">'
            '<label>Your name<input name="name" required minlength="2" maxlength="80" autocomplete="name"></label>'
            '<label>Phone / WhatsApp number<input name="phone" type="tel" required maxlength="20" inputmode="tel" autocomplete="tel" placeholder="+91 98765 43210"></label>'
            '<label>Email <small>(optional)</small><input name="email" type="email" maxlength="120" autocomplete="email"></label>'
            f'<label>Your message<textarea name="message" rows="4" required minlength="3" maxlength="2000" placeholder="{e(ph)}">{e(msg)}</textarea></label>{ex}'
            '<label>Photos <small>(optional, up to 4)</small><input name="photos" type="file" accept="image/*" multiple></label>'
            '<div class="hp" aria-hidden="true"><label>Website<input name="website" tabindex="-1" autocomplete="off"></label></div>'
            f'<button class="btn gold" type="submit">{btn}</button><p class="enq-msg" role="status" aria-live="polite"></p>'
            f'<p class="note">{"Sending opens WhatsApp with your details ready to send. " if STATIC_ENQUIRY else ""}We reply on the number you give. Prefer chat? <a href="{wa("Hello Crafty Korner India, I would like to know more.")}" rel="noopener">Message us on WhatsApp</a>.</p></form>')

def crumbs(trail):
    vis = " / ".join(f'<a href="{u}">{e(t)}</a>' for u, t in trail[:-1]) + (" / " if len(trail) > 1 else "") + e(trail[-1][1])
    ld = {"@context": "https://schema.org", "@type": "BreadcrumbList", "itemListElement":
          [{"@type": "ListItem", "position": k + 1, "name": t, "item": SITE + u} for k, (u, t) in enumerate(trail)]}
    return f'<p class="crumbs">{vis}</p>', ld

def head(trail, h1, intro="", cls="light"):
    c, ld = crumbs(trail)
    return f'<section class="pagehead {cls}">{c}<h1>{e(h1)}</h1>{"<p>" + e(intro) + "</p>" if intro else ""}</section>', ld

def alt(i): return f'{i["name"]}' + (f' – {i["mat"]}' if i["mat"] else "") + f' | {NAME}, Hyderabad'
def card(i):
    pr = "Sold" if i["sold"] else (inr(i["price"]) if i["price"] else "Price on request")
    tag = '<span class="tag">Sold</span>' if i["sold"] else ""
    return f'''<article class="piece" data-cat="{e(i["cat"])}" data-s="{e((i["name"] + " " + i["mat"] + " " + i["cat"] + " " + i["desc"]).lower())}"><a href="/collection/{i["slug"]}/"><div class="frame"><img src="/media/{e(i["img"])}" alt="{e(alt(i))}" width="800" height="1000" loading="lazy">{tag}</div><h3>{e(i["name"])}</h3></a><p class="mat">{e(i["mat"] or i["era"])}</p><p class="price">{pr}<a class="vd" href="/collection/{i["slug"]}/">View details</a></p></article>'''
grid = lambda L: '<div class="grid">' + "".join(card(i) for i in L) + "</div>"

def postcard(p):
    u = f'/blog/{p["slug"]}/'
    img = f'<img src="/media/{e(p["img"])}" alt="{e(p["title"])}" width="900" height="420" loading="lazy">' if p["img"] else ""
    tag = f'<span class="ptag">{e(p["tag"])}</span>' if p["tag"] else ""
    return (f'<article class="pcard"><a class="pimg" href="{u}" tabindex="-1" aria-hidden="true">{img}{tag}</a><h3><a href="{u}">{e(p["title"])}</a></h3>'
            f'<p class="psum">{e(p["summary"])}</p><div class="pmeta"><span>{ico(I["cal"])}<time datetime="{p["date"]}">{p["date"].strftime("%b %d, %Y")}</time></span><span>{ico(I["clock"])}{p["mins"]} min read</span></div></article>')

# wipe generated folders so deleted items disappear
if OUT != ".":   # separate output folder (GitHub Pages build): start from an empty folder, but never wipe anything outside this project
    _o = os.path.abspath(OUT)
    if not _o.startswith(os.getcwd() + os.sep): sys.exit("OUT_DIR must be a folder inside the project: " + OUT)
    shutil.rmtree(_o, ignore_errors=True); os.makedirs(_o)
else:
    for d in ["collection", "blog", "antiques", "handmade-wood", "furniture", "handicrafts", "home-decor", "custom-orders", "about", "visit"]:
        shutil.rmtree(d, ignore_errors=True)

STORE = {"@context": "https://schema.org", "@type": ["Store", "LocalBusiness"], "name": NAME, "logo": SITE + "/logo.png", "areaServed": "India", "keywords": GLOBAL_KW, "url": SITE + "/", "telephone": [PHONE, PHONE2], "email": EMAIL, "sameAs": [u for u in (INSTAGRAM, FACEBOOK, YOUTUBE) if u],
  "image": SITE + "/media/og-virasat.jpg", "description": "Handmade wooden handicrafts, carved furniture, home decor, gifts and antiques at Shilparamam, Madhapur, Hyderabad. Custom orders and delivery across India.",
  "address": {"@type": "PostalAddress", "streetAddress": STREET, "addressLocality": "Hyderabad", "addressRegion": "Telangana", "postalCode": PIN, "addressCountry": "IN"},
  "openingHoursSpecification": [{"@type": "OpeningHoursSpecification", "dayOfWeek": ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"], "opens": "10:30", "closes": "20:00"}]}

# ---------- home ----------
avail = [i for i in items if not i["sold"]] or items
FAQ = [
 ("Where is Crafty Korner India located?", f"We are at {STREET}, Hyderabad {PIN}. The shop is open {HOURS}."),
 ("Do you deliver handmade wooden furniture and handicrafts across India?", "Yes. We deliver across India, and we send photos and a short video on WhatsApp before dispatch so you can see the exact piece."),
 ("Can I order wooden furniture or decor in my own size?", "Yes. Send a reference photo, the size and your budget on WhatsApp or through our custom order page, and we reply with wood options, a price and a time."),
 ("How do I get the price of a piece?", f"Prices are shared on request. Message us on WhatsApp at {PHONE} or {PHONE2} with the name of the piece and we will reply with the price."),
 ("How can I contact Crafty Korner India?", f"Call or WhatsApp {PHONE} or {PHONE2}, or email {EMAIL}."),
]
faq_html = "".join(f"<details><summary>{e(q)}</summary><p>{e(a)}</p></details>" for q, a in FAQ)
faq_ld = {"@context": "https://schema.org", "@type": "FAQPage", "mainEntity": [{"@type": "Question", "name": q, "acceptedAnswer": {"@type": "Answer", "text": a}} for q, a in FAQ]}
TILE_ALT = {"Furniture": "Handmade inlay wooden stools and furniture", "Handicrafts": "Nested wooden trays and handicrafts", "Antiques": "Cane-back settee and heritage furniture", "Home Décor": "Carved wooden wall clocks and home decor"}
tiles = "".join(f'<a class="tile" href="/{v[0]}/"><img src="/media/tile-{v[0]}.jpg" alt="{e(TILE_ALT[k])} | {NAME}, Hyderabad" width="1000" height="420" loading="lazy"><span class="tile-l">{e(k)} {ico(I["arrow"])}</span></a>' for k, v in CATS.items())
TRUST = [("hand", "Handcrafted<br>Products"), ("leaf", "Authentic<br>Indian Art"), ("truck", "Pan India<br>Delivery"), ("tools", "Custom<br>Creations")]
trust = "".join(f'<div class="ti-item"><span class="ti">{ico(I[k])}</span><b>{t}</b></div>' for k, t in TRUST)
body = f'''<section class="hero2">
  <img class="hero-bg" src="/media/hero.jpg" alt="Handmade inlay wooden stools at {NAME}, Shilparamam, Hyderabad" width="1600" height="640" fetchpriority="high">
  <div class="hero-shade"></div>
  <div class="hero2-in"><p class="eyebrow2">Handcrafted in India</p><h1>Timeless Pieces for <em>Modern Homes</em></h1><span class="orn-l" aria-hidden="true"></span>
  <p>A curated collection of handmade wooden furniture, traditional handicrafts and rare antiques, crafted by skilled Indian artisans.</p>
  <a class="btn gold2" href="/collection/">Explore Collection {ico(I["arrow"])}</a></div>
</section>
<section class="trust" aria-label="Why choose us"><div class="trust-in">{trust}</div></section>
<div class="scallop" aria-hidden="true"></div>
<section class="coll"><h2 class="orn-h"><span>Explore Our Collections</span></h2><p class="sub">Furniture, Handicrafts and Antiques that Bring Indian Heritage Home</p><div class="tiles">{tiles}</div></section>
<section class="journal"><div class="jin"><div class="jl"><p class="eyebrow3">From our journal</p><h2>The Crafty Korner Blog</h2><p>Stories, care tips and inspiration from the world of Indian handicrafts and wooden furniture.</p><a class="btn dark2" href="/blog/">Explore All Blogs {ico(I["arrow"])}</a></div><span class="jsep" aria-hidden="true"></span><div class="jcards">{"".join(postcard(p) for p in posts[:3])}</div></div></section>
<section class="more"><h2 class="orn-h"><span>New in the Shop</span></h2><p class="sub">Handmade wooden handicrafts, furniture and home decor from Shilparamam, Hyderabad</p>{grid(avail[:8])}<p class="center"><a class="btn dark2" href="/collection/">View all pieces {ico(I["arrow"])}</a></p></section>
<section class="more faqs"><h2 class="orn-h"><span>Questions We Are Often Asked</span></h2><div class="faq">{faq_html}</div></section>'''
page("/", "Handicrafts & Handmade Wooden Furniture in Hyderabad | Crafty Korner India, Shilparamam",
     "Crafty Korner India at Shilparamam, Madhapur, Hyderabad: handmade wooden handicrafts, carved furniture, inlay trays, stools, wall clocks, home decor, gifts and antiques. Custom orders and delivery across India.",
     body, [STORE, {"@context": "https://schema.org", "@type": "WebSite", "name": NAME, "url": SITE + "/"}, faq_ld],
     images=["/media/hero.jpg"] + ["/media/tile-" + v[0] + ".jpg" for v in CATS.values()] + ["/media/" + i["img"] for i in avail[:8]])

# ---------- collection and categories ----------
h, ld = head([("/", "Home"), ("/collection/", "Collection")], "Handmade wooden handicrafts and antiques",
             "Carved furniture, inlay trays, stools, wall clocks, lamps and gifts from our shop at Shilparamam, Hyderabad.")
cats_present = [c for c in CAT_ORDER if any(i["cat"] == c for i in items)]
chips = f'<button type="button" class="on" data-filter="all">All <span>{len(items)}</span></button>' + "".join(
    f'<button type="button" data-filter="{e(c)}">{e(c)} <span>{sum(1 for i in items if i["cat"] == c)}</span></button>' for c in cats_present)
page("/collection/", "Collection: Handmade Wooden Handicrafts, Furniture & Antiques | Crafty Korner India Hyderabad",
     "Browse handmade wooden handicrafts, carved furniture, inlay trays, stools, wall clocks, lamps, gifts and antiques from Crafty Korner India at Shilparamam, Hyderabad. Delivery across India.",
     h + f'<section class="wrap"><div class="filters" role="group" aria-label="Filter the collection">{chips}</div><p class="qnote" hidden></p>{grid(items)}<p class="noresult" hidden>No pieces match your search. <a href="/collection/">See the full collection</a> or <a href="/custom-orders/">ask for a custom piece</a>.</p></section>', [ld],
     images=["/media/" + i["img"] for i in items])
for cat, (cs, h1, intro) in CATS.items():
    L = [i for i in items if i["cat"] == cat]
    h, ld = head([("/", "Home"), ("/collection/", "Collection"), (f"/{cs}/", cat)], h1, intro)
    page(f"/{cs}/", f"{h1} | {NAME}", intro, h + f'<section class="wrap">{grid(L) if L else "<p>New pieces are arriving soon.</p>"}</section>', [ld],
         kw=f"{cat} Hyderabad, buy {cat.lower()} online India", images=["/media/" + i["img"] for i in L])

# ---------- product pages ----------
for i in items:
    p = f"/collection/{i['slug']}/"; img = "/media/" + i["img"]
    cs = CATS.get(i["cat"], ("collection", "", ""))[0]
    h, ld = head([("/", "Home"), ("/collection/", "Collection"), (p, i["name"])], i["name"])
    price = inr(i["price"]) if i["price"] else "Price on request"
    spec = "".join(f"<dt>{k}</dt><dd>{e(v)}</dd>" for k, v in (("Category", i["cat"]), ("Material", i["mat"]), ("Age", i["era"]), ("Size", i["size"])) if v)
    if i["sold"]:
        cta = f'<p class="soldnote">This piece has been sold. Antiques are single pieces, but we can often find or make something similar.</p><details class="enqbox"><summary class="btn">Ask for something similar</summary>{enq_form("product enquiry", i["name"], p, "Do you have something similar to the " + i["name"] + "?")}</details>'
    else:
        cta = f'<details class="enqbox"><summary class="btn">Enquire about this piece</summary>{enq_form("product enquiry", i["name"], p, "I am interested in the " + i["name"] + ". Is it available?")}</details> <a class="btn ghost-dark" href="/custom-orders/">Need a different size?</a>'
    rel = [x for x in items if x["cat"] == i["cat"] and x is not i and not x["sold"]][:3]
    allp = [i["img"]] + i["more"]
    thumbs = ('<div class="thumbs">' + "".join(f'<button type="button" data-src="/media/{e(x)}" aria-label="Photo {n+1} of {len(allp)}"><img src="/media/{e(x)}" alt="{e(i["name"])} photo {n+1} | {NAME}" width="80" height="80" loading="lazy"></button>' for n, x in enumerate(allp)) + '</div>') if len(allp) > 1 else ""
    paras = "".join(f"<p>{e(x)}</p>" for x in i["desc"].split("\n") if x.strip())
    body = h + f'''<section class="wrap prod"><div class="gal"><div class="frame"><img id="mainimg" src="{img}" alt="{e(alt(i))}" width="800" height="1000">{'<span class="tag">Sold</span>' if i["sold"] else ""}</div>{thumbs}</div>
<div><p class="big-price">{"Sold" if i["sold"] else price}</p><dl class="spec">{spec}</dl>{paras}<p class="cta">{cta}</p></div></section>
{'<section class="wrap"><h2>More like this</h2>' + grid(rel) + '</section>' if rel else ''}'''
    prod = {"@context": "https://schema.org", "@type": "Product", "name": i["name"], "keywords": f'{i["name"]}, {i["mat"]}, Hyderabad', "description": i["desc"], "image": [SITE + "/media/" + x for x in [i["img"]] + i["more"]],
            "category": i["cat"], "brand": {"@type": "Brand", "name": NAME}}
    if i["mat"]: prod["material"] = i["mat"]
    if i["price"]:
        prod["offers"] = {"@type": "Offer", "url": SITE + p, "priceCurrency": "INR", "price": int(i["price"]),
            "availability": "https://schema.org/SoldOut" if i["sold"] else "https://schema.org/InStock",
            "itemCondition": "https://schema.org/UsedCondition" if i["cat"] == "Antiques" else "https://schema.org/NewCondition"}
    nm = i["name"].lower()
    page(p, f'{i["name"]} | Buy {i["cat"]} in Hyderabad | {NAME}',
         f'{i["name"]} ({i["mat"]}) at {NAME}, Shilparamam, Hyderabad. {i["desc"]}', body, [prod, ld], img=img, og="product",
         kw=f'{i["name"]}, {i["mat"]}, buy {nm} online India, {nm} Hyderabad, {i["cat"]} Hyderabad',
         images=["/media/" + x for x in [i["img"]] + i["more"]])

# ---------- custom orders, about, visit ----------
h, ld = head([("/", "Home"), ("/custom-orders/", "Custom orders")], "Custom handmade wooden furniture and decor in Hyderabad",
             "Tell us what you need. We make it in our workshop and deliver across India.")
page("/custom-orders/", "Custom Wooden Furniture & Handicrafts Made to Order in Hyderabad | Crafty Korner India",
     "Order handmade teak, rosewood or sheesham furniture, mandirs, doors and gifts made to your size by Crafty Korner India, Hyderabad. Delivery across India.",
     h + '''<section class="custom"><div class="wrap two"><div><h2>How it works</h2><p>Tell us what you have in mind. Send a reference photo with the form and we will reply with wood options, a price and a time.</p>
<ul class="steps"><li>You share the idea, size and budget.</li><li>We suggest wood and design, and quote in writing.</li><li>We make it in our workshop, send photos, and deliver across India.</li></ul><p>Not sure which wood? Read our <a href="/blog/teak-rosewood-or-sheesham-which-wood-for-handmade-furniture/">guide to teak, rosewood and sheesham</a>.</p></div>
''' + enq_form("custom order", extra=True, fid="orderForm", ph="e.g. a teak mandir, 3 ft tall", btn="Send custom order") + '''</div></section>''', [ld])

h, ld = head([("/", "Home"), ("/about/", "Our story")], "Why Crafty Korner India: antiques and handmade woodcraft at Shilparamam")
page("/about/", "About Crafty Korner India | Antique & Handmade Wood Shop, Shilparamam Hyderabad",
     "Learn how we source antiques from old homes and temples and make handmade teak and rosewood pieces in Hyderabad.",
     h + f'''<section class="wrap story"><div class="narrow"><p>{NAME} started with a simple belief: a good piece of wood or brass should outlive the person who buys it.</p>
<p>Our antiques come from old homes and temples across Telangana, Tamil Nadu and Rajasthan. Our new work is made by carpenters who learned the craft from their fathers, using seasoned teak, rosewood and sheesham, with no veneer and no shortcuts.</p>
<p>Come and hold a piece before you decide. If you cannot visit, we will send photos and a video on WhatsApp. See the <a href="/collection/">collection</a> or read the <a href="/blog/">blog</a>.</p></div></section>''', [ld])

h, ld = head([("/", "Home"), ("/visit/", "Visit us")], "Visit our shop at Shilparamam, Madhapur, Hyderabad")
page("/visit/", "Visit Crafty Korner India | Antique Shop at Shilparamam, Madhapur, Hyderabad",
     f"Find Crafty Korner India at {STREET}, Hyderabad {PIN}. Open {HOURS}. Call or WhatsApp {PHONE} or {PHONE2}.",
     h + f'''<section class="visit"><div class="wrap two"><address><p><strong>{e(STREET)}</strong><br>Hyderabad, Telangana {PIN}</p><p>Open {HOURS}</p>
<p><a class="btn gold" href="https://www.google.com/maps/search/?api=1&amp;query=Shilparamam+Hyderabad" target="_blank" rel="noopener">Get directions</a></p></address>
<div><p>Call or WhatsApp<br><a class="phone" href="tel:{PHONE.replace(" ", "")}">{e(PHONE)}</a><br><a class="phone" href="tel:{PHONE2.replace(" ", "")}">{e(PHONE2)}</a></p><p>Email<br><a href="mailto:{EMAIL}">{EMAIL}</a></p><h2 class="vh">Send us a message</h2>{enq_form("general query", ph="How can we help?", btn="Send message")}</div></div></section>''', [ld, STORE])

# ---------- blog ----------
h, ld = head([("/", "Home"), ("/blog/", "Blog")], "The Crafty Korner Blog", "Stories, care tips and inspiration from the world of Indian handicrafts and wooden furniture.")
page("/blog/", "Blog: Wooden Furniture, Handicraft Care & Carving Guides | Crafty Korner India Hyderabad",
     "Care tips, styling ideas and guides on wood carving, choosing teak, rosewood or sheesham, and ordering custom furniture in India.",
     h + f'<section class="wrap"><div class="posts">{"".join(postcard(p) for p in posts)}</div></section>', [ld])
for p in posts:
    path = f"/blog/{p['slug']}/"
    h, ld = head([("/", "Home"), ("/blog/", "Blog"), (path, p["title"])], p["title"], p["summary"])
    art = {"@context": "https://schema.org", "@type": "Article", "headline": p["title"], "description": p["summary"],
           "datePublished": p["date"].isoformat(), "dateModified": p["date"].isoformat(), "mainEntityOfPage": SITE + path,
           "image": SITE + "/media/" + (p["img"] or "og-virasat.jpg"), "author": {"@type": "Organization", "name": NAME},
           "publisher": {"@type": "Organization", "name": NAME}}
    page(path, f'{p["title"]} | {NAME}', p["summary"], h + f'<article class="wrap article"><p class="ameta"><time datetime="{p["date"]}">{p["date"].strftime("%d %B %Y")}</time> · {p["mins"]} min read</p>{f'<img class="aimg" src="/media/{e(p["img"])}" alt="{e(p["title"])}" width="900" height="420">' if p["img"] else ""}{md(p["body"])}</article>', [art, ld], og="article", img="/media/" + (p["img"] or "og-virasat.jpg"), images=["/media/" + p["img"]] if p["img"] else ())

# ---------- 404, sitemap, robots ----------
page("/404-page/", "Page not found | " + NAME, "Page not found.", '<section class="wrap article"><h1>This page is not here</h1><p>Try the <a href="/collection/">collection</a> or go back to the <a href="/">home page</a>.</p></section>', index=False)
_404 = os.path.join(OUT, "404.html")
os.replace(os.path.join(OUT, "404-page", "index.html"), _404); shutil.rmtree(os.path.join(OUT, "404-page"))
_h = open(_404, encoding="utf-8").read()   # the error page has no address of its own: no canonical / hreflang / og:url
_h = re.sub(r'<link rel="(?:canonical|alternate)"[^>]*404-page/">\n?', "", _h); _h = re.sub(r'<meta property="og:url" content="[^"]*404-page/">', "", _h)
open(_404, "w", encoding="utf-8").write(_h)
today = date.today().isoformat()
def _u(u):
    ims = "".join(f"<image:image><image:loc>{SITE}{e(x)}</image:loc></image:image>" for x in IMGS.get(u, []))
    return f"<url><loc>{SITE}{u}</loc><lastmod>{today}</lastmod>{ims}</url>\n"
open(os.path.join(OUT, "sitemap.xml"), "w", encoding="utf-8").write('<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9" xmlns:image="http://www.google.com/schemas/sitemap-image/1.1">\n' +
    "".join(_u(u) for u in URLS) + "</urlset>\n")
open(os.path.join(OUT, "robots.txt"), "w").write(f"User-agent: *\nAllow: /\nDisallow: {BASE}admin/\nDisallow: {BASE}api/\n\nSitemap: {SITE}/sitemap.xml\n")
open(os.path.join(OUT, "site.webmanifest"), "w", encoding="utf-8").write(json.dumps({
    "name": NAME, "short_name": "Crafty Korner", "description": "Handmade wooden handicrafts, furniture and antiques from Shilparamam, Hyderabad.",
    "start_url": BASE, "scope": BASE, "display": "standalone", "background_color": "#F6EEE2", "theme_color": "#4A2512",
    "icons": [{"src": BASE + "icon-192.png", "sizes": "192x192", "type": "image/png"}, {"src": BASE + "icon-512.png", "sizes": "512x512", "type": "image/png"}]}, indent=2) + "\n")

# separate output folder: copy the public static files next to the pages so the folder can be published as it is
if OUT != ".":
    for f in ("style.css", "script.js", "favicon.ico", "favicon-32.png", "icon-192.png", "icon-512.png", "apple-touch-icon.png", "logo.png", "logo-mark.png"):
        shutil.copy(f, os.path.join(OUT, f))
    shutil.copytree(MEDIA_DIR, os.path.join(OUT, "media"))
    open(os.path.join(OUT, ".nojekyll"), "w").close()   # tell GitHub Pages to publish the files as they are (no Jekyll processing)
print(f"Done: {len(items)} items, {len(posts)} blog posts, {len(URLS)} pages in sitemap.  Base path: {BASE}  Site: {SITE}/  Written to: {os.path.abspath(OUT)}")
