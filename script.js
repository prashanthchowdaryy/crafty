const btn = document.querySelector(".menu-btn"), nav = document.getElementById("nav");
btn.addEventListener("click", () => btn.setAttribute("aria-expanded", nav.classList.toggle("open")));
nav.addEventListener("click", e => { if (e.target.closest("a")) { nav.classList.remove("open"); btn.setAttribute("aria-expanded", false); } });
const dd = document.querySelector(".dd-btn");
if (dd) dd.addEventListener("click", () => dd.setAttribute("aria-expanded", dd.parentElement.classList.toggle("open")));

// search icon opens the search bar
const sb = document.querySelector(".search-btn"), bar = document.querySelector(".searchbar");
if (sb && bar) sb.addEventListener("click", () => {
  bar.hidden = !bar.hidden; sb.setAttribute("aria-expanded", String(!bar.hidden));
  if (!bar.hidden) bar.querySelector("input").focus();
});

// enquiry / custom order forms: sent straight to the shop (WhatsApp + email) without leaving the page
const CFG = (document.querySelector("script[data-wa]") || {dataset: {}}).dataset;
const WA = CFG.wa || "";
const BASE = CFG.base || "/";                          // "/" locally, "/crafty/" on GitHub Pages (set by build.py)
const STATIC_MODE = CFG.enquiry === "whatsapp";        // GitHub Pages has no server: forms open WhatsApp instead of calling /api/enquiry
const SITE = (CFG.site || location.origin + BASE.replace(/\/$/, "")).replace(/\/$/, "");
const T0 = Date.now();

// GitHub Pages version: build a WhatsApp message from the form and open it. Nothing is stored anywhere; the enquiry is only
// delivered when the customer presses Send inside WhatsApp, and the page says so.
function waText(form) {
  const v = n => ((form.elements[n] || {}).value || "").trim();
  const d = form.dataset, L = [];
  L.push(`Hello Crafty Korner India, ${d.kind ? "this is a " + d.kind + " from your website." : "I have a question from your website."}`);
  if (d.product) L.push(`Product: ${d.product}`);
  if (d.url) L.push(`Link: ${SITE}${d.url}`);
  L.push("", `Name: ${v("name")}`, `Phone: ${v("phone")}`);
  if (v("email")) L.push(`Email: ${v("email")}`);
  if (v("extra")) L.push(`Wood and budget: ${v("extra")}`);
  L.push("", `Message: ${v("message")}`);
  const n = form.elements.photos && form.elements.photos.files ? form.elements.photos.files.length : 0;
  if (n) L.push("", `I have ${n} photo${n > 1 ? "s" : ""} to share here in this chat.`);
  const t = L.join("\n");
  return t.length > 1800 ? t.slice(0, 1797) + "..." : t;     // keeps the wa.me link a safe length
}
function sendViaWhatsApp(form) {
  const link = `https://wa.me/${WA}?text=${encodeURIComponent(waText(form))}`;
  const a = document.createElement("a"); a.href = link; a.target = "_blank"; a.rel = "noopener";
  document.body.append(a); a.click(); a.remove();             // done inside the click handler, so pop-up blockers allow it
  const msg = form.querySelector(".enq-msg"); msg.className = "enq-msg"; msg.textContent = "";
  msg.append("WhatsApp should now open with your details filled in. Tap Send in WhatsApp to deliver your enquiry; it is not sent until you do. ");
  const again = document.createElement("a"); again.href = link; again.target = "_blank"; again.rel = "noopener"; again.textContent = "Open WhatsApp again";
  msg.append(again);
  if (form.elements.photos && form.elements.photos.files.length) msg.append(" Photos cannot be attached from this page, so please add them in the WhatsApp chat.");
}
async function shrink(file) {  // phone photos are big: resize in the browser before uploading
  try {
    const bmp = await createImageBitmap(file), k = Math.min(1, 1600 / Math.max(bmp.width, bmp.height));
    const c = document.createElement("canvas"); c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
    c.getContext("2d").drawImage(bmp, 0, 0, c.width, c.height);
    return await new Promise(r => c.toBlob(r, "image/jpeg", .85));
  } catch (err) { return file.size < 6e6 ? file : null; }
}
document.querySelectorAll("form.enq").forEach(form => form.addEventListener("submit", async ev => {
  ev.preventDefault();
  if (!form.reportValidity()) return;
  if (STATIC_MODE) { if (!(form.elements.website && form.elements.website.value)) sendViaWhatsApp(form); return; }   // hidden spam-trap field filled in: do nothing
  const msg = form.querySelector(".enq-msg"), btn = form.querySelector("button[type=submit]"), label = btn.textContent;
  const f = new FormData(form); f.delete("photos");
  f.set("kind", form.dataset.kind); f.set("product", form.dataset.product); f.set("product_url", form.dataset.url); f.set("ts", T0);
  msg.className = "enq-msg"; msg.textContent = ""; btn.disabled = true; btn.textContent = "Sending…";
  try {
    const files = [...form.elements.photos.files].slice(0, 4);
    for (const [n, file] of files.entries()) { const b = await shrink(file); if (b) f.append("photos", b, `photo${n + 1}.jpg`); }
    const r = await fetch(BASE + "api/enquiry", {method: "POST", body: f});
    const d = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(d.error || "Something went wrong.");
    form.innerHTML = `<div class="enq-ok" role="status"><h3>Thank you, we have your message</h3><p>We will get back to you on the number you gave. For anything urgent, call us or <a href="https://wa.me/${WA}" rel="noopener">message us on WhatsApp</a>.</p></div>`;
  } catch (err) {
    msg.className = "enq-msg bad"; msg.textContent = (err.message || "Could not send.") + " You can also message us on WhatsApp.";
    btn.disabled = false; btn.textContent = label;
  }
}));

// product page photo gallery
const mainImg = document.getElementById("mainimg");
if (mainImg) document.querySelectorAll(".thumbs button").forEach(b => b.addEventListener("click", () => {
  mainImg.src = b.dataset.src;
  mainImg.classList.toggle("full", b !== document.querySelector(".thumbs button"));
  document.querySelectorAll(".thumbs button").forEach(x => x.classList.toggle("on", x === b));
}));

// collection page: category buttons + ?q= search
const chips = document.querySelectorAll(".filters button");
const pieces = document.querySelectorAll(".grid .piece");
const qnote = document.querySelector(".qnote"), nores = document.querySelector(".noresult");
let cat = "all", q = (new URLSearchParams(location.search).get("q") || "").trim().toLowerCase();
function apply() {
  let n = 0;
  pieces.forEach(p => {
    const ok = (cat === "all" || p.dataset.cat === cat) && (!q || q.split(/\s+/).every(w => p.dataset.s.includes(w)));
    p.hidden = !ok; if (ok) n++;
  });
  if (nores) nores.hidden = n > 0;
}
chips.forEach(b => b.addEventListener("click", () => {
  chips.forEach(x => x.classList.toggle("on", x === b)); cat = b.dataset.filter; apply();
}));
if (pieces.length && q) {
  if (qnote) { qnote.hidden = false; qnote.textContent = `Showing results for "${q}". `; const a = document.createElement("a"); a.href = BASE + "collection/"; a.textContent = "Clear search"; qnote.append(a); }
  const inp = document.querySelector(".searchbar input"); if (inp) inp.value = q;
}
if (pieces.length) apply();
