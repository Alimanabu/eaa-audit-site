"""Free one-page snapshot: scan 3-4 pages of a shop with axe-core and write a 1-page A4 PDF to attach to an email.

One shop:   python quick.py https://www.shop.nl            -> out/quick/shop.nl/EAA_free_snapshot_shop.nl.pdf
Many shops: python quick.py --list shops.xlsx [--column site]   (.csv or .xlsx; one shop per row)

Pages: home, one category, one product, the cart (the contact page replaces a cart that cannot be found).
Rules of the scan are those of snapshot.py: robots.txt is obeyed, 2 s pause between pages, cookies are only ever
rejected, nothing is ordered or submitted. Before the picture of an element is taken, leftover cookie dialogs and every
fixed or sticky bar that does not hold the element are hidden, so the picture shows the element and nothing on top of it.
The batch mode runs each shop in its own process with a time limit; one broken address never stops the rest.
Exit code of a single run: 0 PDF written, 2 error (the reason is in summary.json and on the last line).
"""
import argparse
import csv
import io
import json
import re
import subprocess
import sys
import time
import urllib.parse
from datetime import datetime, timezone
from pathlib import Path

from PIL import Image, ImageDraw
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeout
from playwright.sync_api import sync_playwright

from report import summarise
from report_pdf import FONTS, safe  # importing report_pdf registers the Vera fonts
from snapshot import (AXE_JS, COVERAGE, RUN_AXE, VIEWPORT, inject_axe, robots_for, settle, slim_rule)
from wcag import AXE_TAGS, CRITERIA, IMPACT_ORDER, PLAIN

HERE = Path(__file__).parent
SITE = "https://eaa-audit.net"
AUDIT_PRICE = "€390"
MAX_FINDINGS = 3
MIN_PAGES = 3
MAX_PAGES = 5
PAUSE = 3   # seconds between two page loads of one shop: at most 1 page in 3 s
BLOCK_TEXT = re.compile(r"captcha|verify (that )?you are (a )?human|just a moment|checking your browser|access denied|"
                        r"are you a robot|vous .tes un robot|v.rification de s.curit.|attention required", re.I)
SHOP_TIMEOUT_S = 420
DEVICE_SCALE = 2
CROP_MIN = (380, 170)   # css px
CROP_MAX = (700, 300)
SIZE_MIN_PX = 8

# --- cookie dialogs: reject only, in the languages of the markets we write to ---------------------------------------
REJECT_ANY = re.compile(
    r"^\s*(reject|decline|deny|refuse|only necessary|necessary only|essential only|(allow|use|accept) (only )?"
    r"(essential|necessary)|continue without"                                    # English
    r"|(alles |alle )?(weiger|afwijzen)|alleen (de )?(noodzakelijke|essenti.le|functionele)|niet akkoord"
    r"|doorgaan zonder|nee,? bedankt|enkel (de )?(noodzakelijke|essenti.le)"      # Dutch
    r"|(tout |tous )?(refuser|rejeter)|continuer sans (accepter|consentir)|uniquement (les )?(cookies )?"
    r"(n.cessaires|essentiels)|seulement (les )?(cookies )?(n.cessaires|essentiels)|non merci"  # French
    r"|(alle )?ablehnen|nur (notwendige|erforderliche|essenzielle)|ohne zustimmung)", re.I)
ACCEPT_ANY = re.compile(r"^\s*(accept|allow|agree|accepteer|alles accepteren|akkoord|accepter|tout accepter|j.accepte"
                        r"|alle akzeptieren)\b", re.I)

# hides what is left of a cookie dialog: fixed or sticky blocks that look like one; returns how many were hidden
HIDE_COOKIE_LEFTOVERS = r"""() => {
  const named = /cookie|consent|gdpr|onetrust|cybot|usercentrics|didomi|truste|cmp|toestemming|axeptio|iubenda|osano|borlabs|privacy-?(banner|notice)/i;
  const words = /cookie|consent|toestemming|vie priv|donn.es personnelles|datenschutz/i;
  let n = 0;
  for (const el of document.querySelectorAll('body *')) {
    const s = getComputedStyle(el);
    if (s.display === 'none' || s.visibility === 'hidden') continue;
    const fixed = s.position === 'fixed' || s.position === 'sticky';
    const id = (el.id || '') + ' ' + (typeof el.className === 'string' ? el.className : '');
    const dialog = ['dialog', 'alertdialog'].includes(el.getAttribute('role')) || el.tagName === 'DIALOG';
    const r = el.getBoundingClientRect();
    if (!(r.width * r.height > 6000)) continue;
    if ((fixed && named.test(id)) || (dialog && words.test(el.innerText || '') && named.test(id + (el.innerText||'').slice(0, 200)))
        || (el.id && /^(usercentrics-root|CybotCookiebotDialog|onetrust-consent-sdk)$/.test(el.id))) {
      el.style.setProperty('display', 'none', 'important'); n++;
    }
  }
  document.documentElement.style.setProperty('overflow', 'auto', 'important');
  if (document.body) document.body.style.setProperty('overflow', 'auto', 'important');
  return n;
}"""

LINKS_JS = r"""() => [...document.querySelectorAll('a[href]')].map(a => {
  const r = a.getBoundingClientRect();
  const text = (a.innerText || a.getAttribute('aria-label') || a.title || '').trim().replace(/\s+/g, ' ').slice(0, 60);
  return {href: a.href, text, nav: !!a.closest('nav, header, [role=navigation]'), footer: !!a.closest('footer, [role=contentinfo]'),
          img: !!a.querySelector('img, picture'), main: !!a.closest('main, [role=main]'), vis: r.width > 0 && r.height > 0};
})"""

PREP_ELEMENT = r"""(sel) => {
  const el = document.querySelector(sel);
  if (!el) return null;
  const s = getComputedStyle(el);
  if (s.display === 'none' || s.visibility === 'hidden') return null;
  el.scrollIntoView({block: 'center', inline: 'center', behavior: 'instant'});
  return true;
}"""

ELEMENT_RECT = r"""(sel) => {
  const r = document.querySelector(sel).getBoundingClientRect();
  return {x: r.x, y: r.y, w: r.width, h: r.height, vw: innerWidth, vh: innerHeight};
}"""

HIDE_OVERLAYS = r"""(sel) => {
  const target = document.querySelector(sel), hidden = [];
  for (const el of document.querySelectorAll('body *')) {
    const s = getComputedStyle(el);
    if ((s.position === 'fixed' || s.position === 'sticky') && s.visibility !== 'hidden'
        && !el.contains(target) && !target.contains(el)) {
      el.style.setProperty('visibility', 'hidden', 'important'); hidden.push(el);
    }
  }
  window.__eaaHidden = hidden;
  return hidden.length;
}"""

SHOW_OVERLAYS = """() => { for (const el of window.__eaaHidden || []) el.style.removeProperty('visibility'); window.__eaaHidden = []; }"""

NOT_PRODUCT = re.compile(
    r"login|account|mijn|compte|signin|sign-in|register|cart|basket|/bag|winkelwagen|winkelmand|panier|checkout|contact|"
    r"help|faq|about|over-ons|a-propos|blog|store|winkels|magasin|wishlist|verlang|favor|privacy|terms|voorwaarden|cgv|"
    r"mentions|cookie|retour|return|delivery|verzend|livraison|shipping|gift|cadeau|newsletter|nieuwsbrief|"
    r"klantenservice|service|press|career|vacature|jobs|sitemap|accessib|toegankelijk|search|zoek|recherche|login|"
    r"/pages?/|/blogs?/|/news|\.(pdf|jpe?g|png|gif|svg|webp|zip)$", re.I)
CATEGORY_HINT = re.compile(r"sale|solde|outlet|new|nouveau|nieuw|women|woman|femme|dames|vrouw|\bmen\b|\bman\b|homme|"
                           r"heren|kids|kinder|enfant|shop|collection|categor|produit|product|alle|boutique|mode|"
                           r"vetement|kleding|schoen|chaussure|maison|wonen|home|jeans|dress|robe|jurk", re.I)
PRODUCT_HINT = re.compile(r"/(products?|produits?|producten|artikel|item|p|pd|dp)/[^/]+|-p-?\d+|/\d{5,}|\.html?$", re.I)
CART_HINT = re.compile(r"cart|basket|winkelwagen|winkelmand|panier|warenkorb|/bag", re.I)
CONTACT_HINT = re.compile(r"contact|klantenservice|service-client|aide|customer-?service|help", re.I)
INFO_HINT = re.compile(r"return|retour|delivery|verzend|bezorg|livraison|shipping|faq", re.I)
CART_GUESSES = ("/cart", "/winkelwagen", "/winkelmand", "/panier", "/basket")


STUDIO_HINT = re.compile(r"service|work|portfolio|case|about|team|solution|expertise|agence|agency|r.alisation|"
                         r"projet|savoir|offre|prestation|shopify|ecommerce|e-commerce", re.I)


class StoreError(Exception):
    pass


def normalise(raw):
    raw = raw.strip()
    if not re.match(r"^https?://", raw, re.I):
        raw = "https://" + raw
    parts = urllib.parse.urlparse(raw)
    if not parts.netloc or "." not in parts.netloc:
        raise StoreError(f"not a web address: {raw!r}")
    return f"{parts.scheme}://{parts.netloc}", re.sub(r"^www\.", "", parts.netloc.lower())


def host_key(url):
    return re.sub(r"^www\.", "", urllib.parse.urlparse(url).netloc.lower())


def local_links(links, root):
    """Same-shop links without fragment and query, as (url, record)."""
    seen, out = set(), []
    for link in links:
        parts = urllib.parse.urlparse(link["href"])
        if parts.scheme not in ("http", "https") or host_key(link["href"]) != host_key(root):
            continue
        url = urllib.parse.urlunparse((parts.scheme, parts.netloc, parts.path, "", "", ""))
        if parts.path in ("", "/") or url in seen:
            continue
        seen.add(url)
        out.append((url, link))
    return out


def depth(url):
    return len([p for p in urllib.parse.urlparse(url).path.split("/") if p])


STRONG_CATEGORY = re.compile(r"/(collections?|categor(y|ies|ie)|categorie|c|cat|shop|rayon|boutique)/[^/]+", re.I)


def category_candidates(links, root):
    """Menu-like links: short text, shallow path, not a service page. Visible ones and typical category paths first."""
    found = [(u, l) for u, l in local_links(links, root)
             if 1 <= depth(u) <= 3 and 3 <= len(l["text"]) <= 30 and not NOT_PRODUCT.search(u)
             and not NOT_PRODUCT.search(l["text"]) and not l["footer"] and not l["img"]]
    found.sort(key=lambda t: (not STRONG_CATEGORY.search(t[0]), not t[1]["vis"], not t[1]["nav"],
                              not CATEGORY_HINT.search(t[0] + " " + t[1]["text"])))
    return [u for u, _ in found][:4]


def product_candidates(links, root, category_url):
    found = [(u, l) for u, l in local_links(links, root)
             if (l["img"] or PRODUCT_HINT.search(u)) and u != category_url and not NOT_PRODUCT.search(u) and not l["footer"]
             and (PRODUCT_HINT.search(u) or (depth(u) >= 2 and not STRONG_CATEGORY.search(u)))]
    found.sort(key=lambda t: (not PRODUCT_HINT.search(t[0]), not t[1]["main"]))
    urls = [u for u, _ in found]
    return urls[2:3] + urls[:2] + urls[3:6] if len(urls) > 2 else urls  # a mid-list product rather than the first


def pick_link(links, root, pattern, footer=None):
    for url, link in local_links(links, root):
        if pattern.search(url) or pattern.search(link["text"]):
            if footer is None or link["footer"] == footer:
                return url
    return None


class Scanner:
    def __init__(self, root, slug, out, studio=False):
        self.studio = studio
        self.extras = []
        self.root, self.slug, self.out = root, slug, out
        self.source = AXE_JS.read_text(encoding="utf-8")
        self.rp = robots_for(root + "/")
        self.pages = []      # scan records, in order
        self.tabs = {}       # n -> open page, kept until the pictures are taken
        self.banner = {"clicked": None, "hiddenAfter": 0}
        self.last_visit = 0.0

    def allowed(self, url):
        return self.rp.can_fetch("*", url)

    def close_cookies(self, page, first):
        if first:
            for _ in range(6):  # dialogs often appear a few seconds after the load
                if self.find_reject(page) or page.evaluate("() => !!document.querySelector('[id*=cookie i],[id*=consent i],[class*=cookie i],[id*=onetrust i],[id*=cybot i]')"):
                    break
                page.wait_for_timeout(1000)
        label = self.find_reject(page, click=True)
        if label:
            self.banner["clicked"] = self.banner["clicked"] or label
            page.wait_for_timeout(1200)
        hidden = page.evaluate(HIDE_COOKIE_LEFTOVERS)
        self.banner["hiddenAfter"] += hidden
        return hidden

    @staticmethod
    def find_reject(page, click=False):
        for frame in page.frames:
            try:
                buttons = frame.get_by_role("button", name=REJECT_ANY)
                for i in range(min(buttons.count(), 6)):
                    b = buttons.nth(i)
                    if b.is_visible():
                        label = b.inner_text().strip()[:60] or "(button)"
                        if len(label) > 50 or ACCEPT_ANY.search(label):
                            continue
                        if click:
                            b.click(timeout=3000)
                        return label
            except PlaywrightError:
                continue
        return None

    def visit(self, ctx, kind, url, want_links=True):
        """Open one page in a new tab and scan it. Raises StoreError and closes the tab when the page is not usable."""
        if not self.allowed(url):
            raise StoreError(f"{kind}: disallowed by robots.txt")
        wait = PAUSE - (time.time() - self.last_visit)
        if self.last_visit and wait > 0:
            time.sleep(wait)
        page = ctx.new_page()
        n = len(self.pages) + 1
        rec = {"n": n, "type": kind, "url": url}
        try:
            try:
                resp = page.goto(url, wait_until="domcontentloaded", timeout=45000)
            except PlaywrightTimeout:
                raise StoreError(f"{kind}: no answer in 45 s")
            self.last_visit = time.time()
            rec["status"] = resp.status if resp else None
            if rec["status"] == 429:  # the shop asks us to slow down: one retry after a long pause, nothing more
                time.sleep(30)
                resp = page.goto(url, wait_until="domcontentloaded", timeout=45000)
                self.last_visit = time.time()
                rec["status"] = resp.status if resp else None
            if rec["status"] and rec["status"] >= 400:
                raise StoreError(f"{kind}: HTTP {rec['status']}")
            settle(page)
            seen = page.evaluate("() => (document.title + ' ' + document.body.innerText).slice(0, 1500)")
            if BLOCK_TEXT.search(seen) and len(seen) < 1500:
                raise StoreError(f"{kind}: blocked by a bot check or captcha, not bypassed")
            self.close_cookies(page, first=(n == 1))
            rec["finalUrl"] = page.url
            rec["title"] = page.title()
            links = page.evaluate(LINKS_JS) if want_links else []
            rec["framesNotInjected"] = inject_axe(page, self.source)
            if n == 1:
                self.axe_version = page.evaluate("axe.version")
                self.coverage = page.evaluate(COVERAGE, AXE_TAGS)
            axe = page.evaluate(RUN_AXE, AXE_TAGS)
            rec["violations"] = [slim_rule(r) for r in axe["violations"]]
            rec["incomplete"] = [slim_rule(r) for r in axe["incomplete"]]
        except (PlaywrightError, PlaywrightTimeout) as e:
            page.close()
            raise StoreError(f"{kind}: {str(e).splitlines()[0][:160]}")
        except StoreError:
            page.close()
            raise
        self.pages.append(rec)
        self.tabs[n] = page
        print(f"  [{n}] {kind}: {len(rec['violations'])} rules failed  {url}", flush=True)
        return rec, links

    def run_studio(self, ctx, home_links):
        """A studio has no products or cart: its own home, up to two inner pages and the contact page."""
        links = [(u, l) for u, l in local_links(home_links, self.root)
                 if depth(u) <= 2 and 3 <= len(l["text"]) <= 40 and not re.search(r"privacy|cookie|legal|terms|mentions|cgv|\.(pdf|jpe?g|png)$|/blogs?/|episode|podcast", u, re.I)]
        links.sort(key=lambda t: (not STUDIO_HINT.search(t[0] + " " + t[1]["text"]), not t[1]["vis"]))
        contact = pick_link(home_links, self.root, CONTACT_HINT)
        for url in [u for u, _ in links if u != contact][:4]:
            if len([p for p in self.pages if p["type"] == "page"]) == 2:
                break
            try:
                self.visit(ctx, "page", url, want_links=False)
            except StoreError as e:
                print(f"  skip {url}: {e}", flush=True)
        if contact:
            try:
                self.visit(ctx, "contact", contact, want_links=False)
            except StoreError as e:
                print(f"  skip {contact}: {e}", flush=True)
        if len(self.pages) < MIN_PAGES:
            raise StoreError(f"only {len(self.pages)} page(s) could be scanned, {MIN_PAGES} needed")

    def run(self, ctx):
        home, home_links = self.visit(ctx, "home", self.root + "/")
        final = urllib.parse.urlparse(home["finalUrl"])
        if host_key(home["finalUrl"]) != host_key(self.root):  # the shop lives on another address: follow it
            self.root = f"{final.scheme}://{final.netloc}"
            self.rp = robots_for(self.root + "/")
        if self.studio:
            return self.run_studio(ctx, home_links)
        category_url = None
        product_url = None
        for cand in category_candidates(home_links, self.root):
            try:
                rec, links = self.visit(ctx, "category", cand)
            except StoreError as e:
                print(f"  skip {cand}: {e}", flush=True)
                continue
            category_url = cand
            for prod in product_candidates(links, self.root, cand):
                try:
                    self.visit(ctx, "product", prod, want_links=False)
                    product_url = prod
                    break
                except StoreError as e:
                    print(f"  skip {prod}: {e}", flush=True)
            if product_url:
                break
            # no usable product here: this category page is dropped from the scan (the first one stays), the next
            # candidate is tried, so that at most one category page is reported
            if len([p for p in self.pages if p["type"] == "category"]) > 1:
                self.pages.pop()
                self.tabs.pop(rec["n"]).close()
        cart = pick_link(home_links, self.root, CART_HINT)
        for url in ([cart] if cart else []) + [self.root + p for p in CART_GUESSES]:
            try:
                self.visit(ctx, "cart", url, want_links=False)
                break
            except StoreError as e:
                print(f"  skip {url}: {e}", flush=True)
                if cart and url == cart:
                    cart = None
        # contact and delivery pages: they replace a missing cart or add a page when the picture list is short
        self.extras = [("contact", u) for u in [pick_link(home_links, self.root, CONTACT_HINT, footer=True)
                                                or pick_link(home_links, self.root, CONTACT_HINT)] if u]
        self.extras += [("info", u) for u in [pick_link(home_links, self.root, INFO_HINT, footer=True)] if u]
        if not any(p["type"] == "cart" for p in self.pages):
            self.add_extra(ctx)
        if len(self.pages) < MIN_PAGES:
            self.add_extra(ctx)
        if len(self.pages) < MIN_PAGES:
            raise StoreError(f"only {len(self.pages)} page(s) could be scanned, {MIN_PAGES} needed")

    def add_extra(self, ctx):
        """Scan the next extra page; False when none is left or the limit of pages is reached."""
        while self.extras and len(self.pages) < MAX_PAGES:
            kind, url = self.extras.pop(0)
            if any(p["url"] == url for p in self.pages):
                continue
            try:
                self.visit(ctx, kind, url, want_links=False)
                return True
            except StoreError as e:
                print(f"  skip {url}: {e}", flush=True)
        return False


# --- the three main findings -----------------------------------------------------------------------------------------

def crop_box(r):
    w = min(max(r["w"] + 140, CROP_MIN[0]), CROP_MAX[0])
    h = min(max(r["h"] + 140, CROP_MIN[1]), CROP_MAX[1])
    x0 = min(max(r["x"] + r["w"] / 2 - w / 2, 0), max(r["vw"] - w, 0))
    y0 = min(max(r["y"] + r["h"] / 2 - h / 2, 0), max(r["vh"] - h, 0))
    return x0, y0, w, h


def capture(page, selector, strict):
    """Picture of one element with a red frame, or None when it cannot be shown cleanly."""
    try:
        if not page.evaluate(PREP_ELEMENT, selector):
            return None
        page.wait_for_timeout(250)
        r = page.evaluate(ELEMENT_RECT, selector)
        if r["w"] < SIZE_MIN_PX or r["h"] < SIZE_MIN_PX:
            return None
        inside = r["x"] >= -1 and r["y"] >= -1 and r["x"] + r["w"] <= r["vw"] + 1 and r["y"] + r["h"] <= r["vh"] + 1
        fits = r["w"] <= CROP_MAX[0] - 40 and r["h"] <= CROP_MAX[1] - 40
        if not inside or (strict and not fits):
            return None
        x0, y0, w, h = crop_box(r)
        page.evaluate(HIDE_OVERLAYS, selector)
        try:
            page.wait_for_timeout(150)
            raw = page.screenshot(clip={"x": x0, "y": y0, "width": w, "height": h}, type="png")
        finally:
            page.evaluate(SHOW_OVERLAYS)
    except PlaywrightError:
        return None
    img = Image.open(io.BytesIO(raw)).convert("RGB")
    k = img.width / w
    box = [(r["x"] - x0) * k, (r["y"] - y0) * k, (r["x"] + r["w"] - x0) * k, (r["y"] + r["h"] - y0) * k]
    pad = 3 * k
    frame = [max(box[0] - pad, 1), max(box[1] - pad, 1), min(box[2] + pad, img.width - 2), min(box[3] + pad, img.height - 2)]
    draw = ImageDraw.Draw(img)
    for i in range(3 * DEVICE_SCALE):  # 3 css px thick, so that it stays visible when the page is scaled to the PDF
        draw.rectangle([frame[0] - i, frame[1] - i, frame[2] + i, frame[3] + i], outline=(225, 29, 72))
    return img, {"x": round(r["x"]), "y": round(r["y"]), "w": round(r["w"]), "h": round(r["h"]),
                 "crop": [round(x0), round(y0), round(w), round(h)]}


def choose_findings(scanner, summary, pictures):
    """Walk the rules by impact and number of elements; keep the first three with a clean picture of an element."""
    ranked = sorted(summary["failed"], key=lambda r: (r["id"] not in PLAIN, -IMPACT_ORDER[r["impact"]],
                                                       -len(r["unique"]), r["id"]))
    findings, used, skipped = [], set(), []
    for rule in ranked:
        if len(findings) == MAX_FINDINGS:
            break
        nodes = []   # (page number, selector) of single-selector targets: iframe and shadow targets are skipped
        for page in scanner.pages:
            for v in page["violations"]:
                if v["id"] == rule["id"]:
                    nodes += [(page["n"], v_node["target"][0], page["type"]) for v_node in v["nodes"]
                              if len(v_node["target"]) == 1 and isinstance(v_node["target"][0], str)
                              and (page["n"], v_node["target"][0]) not in used]  # one element, one picture
        shot = None
        for strict in (True, False):
            for n, selector, kind in nodes[:12]:
                shot = capture(scanner.tabs[n], selector, strict)
                if shot:
                    break
            if shot:
                break
        if not shot:
            skipped.append({"rule": rule["id"], "reason": "its elements are already shown, inside a frame, hidden or "
                            f"smaller than {SIZE_MIN_PX} px" if nodes else "its only element is already shown or inside a frame"})
            continue
        img, where = shot
        used.add((n, selector))
        path = pictures / f"{len(findings) + 1}_{rule['id']}.jpg"
        img.save(path, "JPEG", quality=88)
        findings.append({"rule": rule["id"], "impact": rule["impact"], "level": rule["level"],
                         "criteria": rule["criteria"], "elements": len(rule["unique"]), "instances": rule["instances"],
                         "pages": len(rule["pages"]), "pageType": kind, "pageNumber": n, "selector": selector,
                         "picture": path.name, "where": where, "documented": rule["id"] in PLAIN,
                         "help": rule["help"]})
        findings[-1]["what"], findings[-1]["who"] = finding_texts(findings[-1], len(scanner.pages))
    # a rule whose element cannot be shown still is a problem of the shop: it is listed without a picture
    for item in skipped[:MAX_FINDINGS - len(findings)]:
        rule = next(r for r in ranked if r["id"] == item["rule"])
        f = {"rule": rule["id"], "impact": rule["impact"], "level": rule["level"], "criteria": rule["criteria"],
             "elements": len(rule["unique"]), "instances": rule["instances"], "pages": len(rule["pages"]),
             "pageType": None, "pageNumber": None, "selector": None, "picture": None, "where": None,
             "documented": rule["id"] in PLAIN, "help": rule["help"]}
        f["what"], f["who"] = finding_texts(f, len(scanner.pages))
        findings.append(f)
    return findings, skipped


# --- the PDF ---------------------------------------------------------------------------------------------------------

def finding_texts(f, total_pages):
    """(what is wrong, who it affects) in plain language, one line each."""
    if f["documented"]:
        title, who, _ = PLAIN[f["rule"]]
    else:
        title = f["help"].rstrip(".")
        who = "People who use a screen reader, a keyboard or voice control may be unable to use this part of the page."
    elements = f"{f['elements']} element{'s' if f['elements'] != 1 else ''}"
    return f"{title}. Found on {elements}, on {f['pages']} of {total_pages} pages.", who


PAGE_LABEL = {"home": "home page", "category": "category page", "product": "product page", "cart": "cart",
              "contact": "contact page", "info": "help page", "page": "inner page"}


def build_pdf(summary, findings, pictures, path):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import Image as RLImage
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    ACCENT, GREY, LINE = colors.HexColor("#1D4ED8"), colors.HexColor("#F3F4F6"), colors.HexColor("#9CA3AF")
    st = lambda name, **kw: ParagraphStyle(name, fontName=kw.pop("font", "Vera"), **kw)
    small = st("small", fontSize=8, leading=10.5, textColor=colors.HexColor("#4B5563"))
    body = st("body", fontSize=9.5, leading=12.5)
    brand = st("brand", font="Vera-Bold", fontSize=9, leading=11, textColor=ACCENT)
    title = st("title", font="Vera-Bold", fontSize=20, leading=24)
    big = st("big", font="Vera-Bold", fontSize=22, leading=26, alignment=TA_CENTER, textColor=ACCENT)
    cap = st("cap", fontSize=8, leading=10.5, alignment=TA_CENTER, textColor=colors.HexColor("#374151"))
    h2 = st("h2", font="Vera-Bold", fontSize=12, leading=15, textColor=ACCENT)
    lab = st("lab", fontSize=9.5, leading=13)

    date = datetime.fromisoformat(summary["scannedAt"]).strftime("%d %B %Y").lstrip("0")
    story = [Paragraph("EAA AUDIT · FREE ACCESSIBILITY SNAPSHOT", brand), Spacer(1, 3),
             Paragraph(safe(summary.get("display") or summary["shop"]), title),
             Paragraph(safe(f"Scanned {date} · desktop view 1366 × 900 · WCAG 2.1 A and AA · "
                            f"axe-core {summary['axeVersion']}" + (f" · {summary['shop']}" if summary.get("display") else "")), small),
             Spacer(1, 8)]

    n_pages = summary["pages"]
    tiles = [(str(n_pages), "pages checked"), (str(summary["violations"]), f"violations found on {summary['distinctElements']} distinct elements"),
             (f"{summary['criteriaFailed']} of {summary['criteriaTotal']}", "WCAG 2.1 A/AA criteria not passed")]
    cells = [[Paragraph(v, big) for v, _ in tiles], [Paragraph(safe(t), cap) for _, t in tiles]]
    tile_table = Table(cells, colWidths=[5.8 * cm] * 3)
    tile_table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), GREY), ("BOX", (0, 0), (0, -1), 0.5, LINE),
                                    ("BOX", (1, 0), (1, -1), 0.5, LINE), ("BOX", (2, 0), (2, -1), 0.5, LINE),
                                    ("TOPPADDING", (0, 0), (-1, 0), 7), ("BOTTOMPADDING", (0, 1), (-1, 1), 7),
                                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    story += [tile_table, Spacer(1, 4),
              Paragraph(safe("Pages: " + ", ".join(PAGE_LABEL[p["type"]] for p in summary["pageList"]) + "."), small),
              Spacer(1, 8)]

    if findings:
        story.append(Paragraph(f"The {len(findings)} biggest problem{'s' if len(findings) != 1 else ''}", h2))
        story.append(Spacer(1, 4))
        for i, f in enumerate(findings, 1):
            what, who = finding_texts(f, n_pages)
            if f["picture"]:
                img = PILImage_size(pictures / f["picture"])
                w = 9.2 * cm
                h = w * img[1] / img[0]
                if h > 4.6 * cm:
                    h, w = 4.6 * cm, 4.6 * cm * img[0] / img[1]
                shot = RLImage(str(pictures / f["picture"]), width=w, height=h)
            else:
                shot = Paragraph("No picture: the element is hidden or too small to be framed on screen.", small)
            level = f"WCAG {', '.join(f['criteria'])} · level {f['level']} · {f['impact']} impact"
            text = [Paragraph(f"<b>{i}. What is wrong:</b> {safe(what)}", lab), Spacer(1, 3),
                    Paragraph(f"<b>Who it affects:</b> {safe(who)}", lab), Spacer(1, 3),
                    Paragraph(safe(level), small)]
            row = Table([[shot, text]], colWidths=[9.6 * cm, 8.0 * cm])
            row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                     ("RIGHTPADDING", (0, 0), (0, 0), 8), ("TOPPADDING", (0, 0), (-1, -1), 2),
                                     ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                                     ("LINEBELOW", (0, 0), (-1, 0), 0.4, LINE)]))
            story += [row, Spacer(1, 5)]
    else:
        story.append(Paragraph("The automated rules found no failures on the pages checked. This does not show that "
                               "the shop is accessible: the rules cover only part of what a person meets.", body))

    story += [Spacer(1, 2),
              Paragraph("This snapshot is an automated scan of the pages named above. Automated rules find only part "
                        "of the accessibility problems, so it is not proof that the shop conforms to WCAG, EN 301 549 "
                        "or the European Accessibility Act, and not legal advice. A red frame marks the element in "
                        "each picture.", small), Spacer(1, 8)]
    link = f'<link href="{SITE}" color="#1D4ED8">{SITE}</link>'
    cta = Table([[Paragraph(f"<b>Want every key page covered?</b> The full automated audit ({AUDIT_PRICE}) checks up to 10 "
                            f"pages you agree with us and lists every problem with a “Fix first” list and "
                            f"the manual checks still needed. Details: {link}", body)]], colWidths=[17.6 * cm])
    cta.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, ACCENT), ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#EFF6FF")),
                             ("LEFTPADDING", (0, 0), (-1, -1), 9), ("RIGHTPADDING", (0, 0), (-1, -1), 9),
                             ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
    story.append(cta)

    doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=1.7 * cm, rightMargin=1.7 * cm, topMargin=1.5 * cm,
                            bottomMargin=1.3 * cm, title=f"Free accessibility snapshot: {summary['shop']}",
                            author="EAA Audit", subject="Automated WCAG 2.1 A/AA snapshot")
    doc.build(story)


def PILImage_size(path):
    with Image.open(path) as im:
        return im.size


def pdf_pages(path):
    return len(re.findall(rb"/Type\s*/Page[^s]", Path(path).read_bytes()))


# --- one shop --------------------------------------------------------------------------------------------------------

def run_one(raw, out_root, display=None, studio=False):
    started = time.time()
    root, slug = normalise(raw)
    out = Path(out_root) / slug
    pictures = out / "pictures"
    pictures.mkdir(parents=True, exist_ok=True)
    for old in list(pictures.glob("*")) + [out / "summary.json"]:
        old.unlink(missing_ok=True)
    summary = {"shop": slug, "display": display, "url": root + "/", "error": None}
    scanner = Scanner(root, slug, out, studio)
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            try:
                ctx = browser.new_context(viewport=VIEWPORT, locale="en-IE", device_scale_factor=DEVICE_SCALE)
                scanner.run(ctx)
                scanned_at = datetime.now(timezone.utc).isoformat(timespec="seconds")
                while True:
                    result = {"shop": slug, "scannedAt": scanned_at, "viewport": VIEWPORT, "tags": AXE_TAGS,
                              "axeVersion": scanner.axe_version, "chromium": browser.version, "pauseSeconds": PAUSE,
                              "cookieBanner": scanner.banner, "automatedCoverage": scanner.coverage,
                              "pages": scanner.pages}
                    s = summarise(result)
                    for old in pictures.glob("*"):
                        old.unlink()
                    findings, skipped = choose_findings(scanner, s, pictures)
                    if len(findings) >= MAX_FINDINGS or not scanner.add_extra(ctx):
                        break
                (out / "results.json").write_text(json.dumps(result, indent=1, ensure_ascii=False), encoding="utf-8")
            finally:
                browser.close()
        summary.update({
            "scannedAt": scanned_at, "axeVersion": scanner.axe_version, "pages": len(s["pages"]),
            "pageList": [{"type": p["type"], "url": p["url"]} for p in s["pages"]],
            "violations": s["levels"]["total"]["instances"], "distinctElements": s["levels"]["total"]["uniqueElements"],
            "criteriaFailed": s["failedCriteria"], "criteriaTotal": len(CRITERIA),
            "levels": s["levels"],
            "findings": findings, "skippedRules": skipped, "cookieBanner": scanner.banner})
        pdf = out / f"EAA_free_snapshot_{slug}.pdf"
        build_pdf(summary, findings, pictures, pdf)
        n = pdf_pages(pdf)
        if n != 1:
            raise StoreError(f"the PDF has {n} pages, it must have 1")
        summary["pdf"] = str(pdf)
    except StoreError as e:
        summary["error"] = str(e)
    except Exception as e:  # a crash in one shop must still leave a reason behind
        summary["error"] = f"{type(e).__name__}: {str(e).splitlines()[0][:200] if str(e) else ''}"
    summary["seconds"] = round(time.time() - started, 1)
    summary.setdefault("pages", len(scanner.pages))
    (out / "summary.json").write_text(json.dumps(summary, indent=1, ensure_ascii=False), encoding="utf-8")
    return summary


# --- many shops ------------------------------------------------------------------------------------------------------

def read_sites(path, column=None):
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        rows = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
        wb.close()
    else:
        text = path.read_text(encoding="utf-8-sig")
        dialect = csv.Sniffer().sniff(text.splitlines()[0], delimiters=",;\t") if text.strip() else csv.excel
        rows = list(csv.reader(io.StringIO(text), dialect))
    rows = [r for r in rows if any(c not in (None, "") for c in r)]
    if not rows:
        return []
    head = [str(c or "").strip().lower() for c in rows[0]]
    names = [column.lower()] if column else ["site", "website", "url", "domain", "сайт", "адрес"]
    idx = next((i for i, h in enumerate(head) if h in names), None)
    if idx is None:
        idx = next((i for i, h in enumerate(head) if any(n in h for n in names)), None)
    if idx is None:
        raise SystemExit(f"no column {names} in {path.name}; header row: {rows[0]}")
    return [str(r[idx]).strip() for r in rows[1:] if idx < len(r) and r[idx] not in (None, "")]


def kill_tree(proc):
    subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True)


def run_batch(sites, out_root, timeout):
    out_root = Path(out_root)
    out_root.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = out_root / f"summary_{stamp}.csv"
    rows, seen = [], set()
    for n, raw in enumerate(sites, 1):
        try:
            slug = normalise(raw)[1]
        except StoreError as e:
            rows.append({"store": raw, "pages": "", "violations": "", "criteria": "", "error": str(e), "pdf": "", "seconds": ""})
            print(f"[{n}/{len(sites)}] {raw}: {e}", flush=True)
            continue
        if slug in seen:
            continue
        seen.add(slug)
        print(f"[{n}/{len(sites)}] {slug}", flush=True)
        t0 = time.time()
        folder = out_root / slug
        folder.mkdir(parents=True, exist_ok=True)
        (folder / "summary.json").unlink(missing_ok=True)
        log = open(folder / "run.log", "w", encoding="utf-8")
        proc = subprocess.Popen([sys.executable, str(Path(__file__).resolve()), raw, "--out", str(out_root)],
                                stdout=log, stderr=subprocess.STDOUT)
        try:
            proc.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            kill_tree(proc)
            proc.wait()
            (folder / "summary.json").write_text(json.dumps({"error": f"timeout after {timeout} s"}), encoding="utf-8")
        finally:
            log.close()
        try:
            s = json.loads((folder / "summary.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            s = {"error": f"no result, exit code {proc.returncode}"}
        rows.append({"store": slug, "pages": s.get("pages", ""), "violations": s.get("violations", ""),
                     "criteria": s.get("criteriaFailed", ""), "error": s.get("error") or "", "pdf": s.get("pdf", ""),
                     "seconds": s.get("seconds", round(time.time() - t0, 1))})
        print(f"    -> {rows[-1]['error'] or 'ok'}  pages {rows[-1]['pages']}, violations {rows[-1]['violations']}, "
              f"criteria {rows[-1]['criteria']}, {rows[-1]['seconds']} s", flush=True)
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)
    print(f"summary: {csv_path}  ({sum(not r['error'] for r in rows)} ok, {sum(bool(r['error']) for r in rows)} with errors)")
    return csv_path


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("shop", nargs="?", help="shop address, for example https://www.shop.nl")
    ap.add_argument("--list", help=".csv or .xlsx with one shop per row")
    ap.add_argument("--column", help="name of the column with the address (default: site, website, url, domain)")
    ap.add_argument("--out", default=str(HERE / "out" / "quick"))
    ap.add_argument("--name", help="shop name for the title of the PDF")
    ap.add_argument("--studio", action="store_true", help="a web studio: home, inner pages and contact, no product or cart")
    ap.add_argument("--timeout", type=int, default=SHOP_TIMEOUT_S, help="seconds per shop in batch mode")
    args = ap.parse_args()
    if args.list:
        run_batch(read_sites(args.list, args.column), args.out, args.timeout)
        return
    if not args.shop:
        ap.error("give a shop address or --list")
    summary = run_one(args.shop, args.out, args.name, args.studio)
    if summary["error"]:
        print(f"ERROR: {summary['error']}", flush=True)
        raise SystemExit(2)
    print(f"PDF: {summary['pdf']}\n{summary['pages']} pages, {summary['violations']} violations on "
          f"{summary['distinctElements']} elements, {summary['criteriaFailed']} of {summary['criteriaTotal']} criteria, "
          f"{summary['seconds']} s")


if __name__ == "__main__":
    main()
