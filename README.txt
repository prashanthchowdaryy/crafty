CRAFTY KORNER INDIA - WEBSITE + ADMIN PANEL

RUN ON YOUR COMPUTER: double-click run.bat (needs Python 3.10+).
  Website: http://localhost:8000    Admin: http://localhost:8000/admin/
  The first visit to /admin/ asks you to choose a password (10+ characters). Do this straight away, before anyone else can.

ADMIN PANEL
  Products tab: add, edit, delete, add prices. Every change is live on the website within a second and is written to inventory.xlsx.
  Enquiries tab: every message from the website (with customer photos), delivery status for WhatsApp and Gmail, "Send again" and "Delete".
  Excel: inventory.xlsx and the admin stay in step (Download Excel / Upload Excel in the admin). Keep the ID column (L).

ORDERS AND ENQUIRIES TO YOUR WHATSAPP AND GMAIL
  Customers use the forms on the website (product pages and Custom Orders). They never leave the page.
  Every message is saved first (you can always read it in the Enquiries tab), then sent to your Gmail (photos attached) and your WhatsApp (text + photos).
  Setup: copy .env.example to .env and fill it in. Never share or upload .env.
   - Gmail: turn on 2-Step Verification, create an "App password", put it in SMTP_PASSWORD.
   - WhatsApp: WhatsApp Business Cloud API (developers.facebook.com > create an app > WhatsApp). Put the token and Phone number ID in .env,
     and your own number in WHATSAPP_TO. WhatsApp only delivers plain text if you messaged the business number in the last 24 hours, so also create
     one approved message template with a single {{1}} variable and put its name in WHATSAPP_TEMPLATE. Gmail always works without this.
  Until WhatsApp or Gmail is set up, enquiries still arrive in the Enquiries tab.

GOING ONLINE (needs a host that runs Python and keeps files: a VPS, or Render/Railway with a disk)
  pip install -r requirements.txt   then   python app.py
  Put it behind https and set in .env:  COOKIE_SECURE=1  FORCE_HTTPS=1  TRUST_PROXY=1  SITE_URL=https://yourdomain
  Static-only hosts (Netlify, GitHub Pages) cannot run the admin or the enquiry forms.

SECURITY BUILT IN
  Login lockout after repeated wrong passwords, signed cookies, CSRF protection, security headers (CSP, no framing, no sniffing),
  spam traps and rate limits on the forms, photos re-saved (strips hidden data and location), private files never served
  (.env, products.json, enquiries, backups, code). Backups of every product change are kept in /backups.
  Blog posts: edit the Blog sheet in inventory.xlsx, then run update.bat.

GITHUB PAGES (static copy of the public website)
  Build it:   double-click build-github-pages.bat   (or:  python build.py --github-pages)
  This writes a ready-to-publish copy of the site into the docs/ folder with every link and file path starting with /crafty/.
  The normal site in this folder (localhost:8000 / your own server) is not touched and keeps working at "/".
  Publish:    git add -A, git commit, git push. In the GitHub repo: Settings > Pages > Deploy from a branch > main > /docs.
  Run it again after every change to products, photos or the Blog sheet (products are edited locally in /admin, then rebuild).
  What works on GitHub Pages: all pages, search, filters, photo galleries, WhatsApp buttons, and the enquiry forms in a WhatsApp version:
  the form opens WhatsApp with the customer's details filled in, and the customer must press Send there. Nothing is saved on the site,
  and customer photos cannot be attached (the page tells them to add photos in the chat).
  What does NOT work on GitHub Pages: the admin panel, the Enquiries tab, Excel sync, photo uploads, Gmail alerts. Keep using app.py for those.
  Other settings (environment variables, all optional): BASE_PATH (default /), SITE_ORIGIN, STATIC_ENQUIRY=1, OUT_DIR (default . ; docs for GitHub Pages).
  Moving to a custom domain later: build with  set BASE_PATH=/ && set SITE_ORIGIN=https://www.yourdomain.in && set STATIC_ENQUIRY=1 && set OUT_DIR=docs && python build.py
