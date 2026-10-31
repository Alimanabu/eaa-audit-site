"""Scan up to 10 pages with axe-core, one pass, and save results.json with screenshots.

If out/<shop>/notes.json has brandTerms, a masked copy of each page goes to screenshots_anon/.

python snapshot.py pages.json --shop example.ie --keyboard-page product
"""
import argparse
import io
import json
import re
import time
import urllib.parse
import urllib.request
import urllib.robotparser
from datetime import datetime, timezone
from pathlib import Path

from PIL import Image, ImageChops
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeout
from playwright.sync_api import sync_playwright

from wcag import AXE_TAGS, criteria_from_tags, level_from_tags

HERE = Path(__file__).parent
AXE_JS = HERE / "vendor" / "axe.min.js"
MAX_PAGES = 10
PAUSE_S = 2
MAX_TABS = 400
FRAME_STOPS = 30
ANON_CROP_PX = 1500
MAX_NODES = 200
VIEWPORT = {"width": 1366, "height": 900}
ADD_TO_BAG = re.compile(r"add to (bag|cart|basket)", re.I)
REJECT = re.compile(r"^\s*(reject|decline|deny|refuse|only necessary|necessary only|essential only"
                    r"|(allow|use|accept) (only )?(essential|necessary)|continue without)", re.I)
ACCEPT = re.compile(r"^\s*(accept|allow|agree)\b.*(all|cookies)", re.I)
SKIP_NAME = re.compile(r"^\s*(skip|jump) to\b", re.I)

RUN_AXE = """async (tags) => {
  axe.configure({allowedOrigins: ['<unsafe_all_origins>']});
  const run = axe.run(document, {
    runOnly: {type: 'tag', values: tags},
    resultTypes: ['violations', 'incomplete'],
    iframes: true, frameWaitTime: 15000, pingWaitTime: 1000,
  });
  const limit = new Promise((_, no) => setTimeout(() => no(new Error('axe timeout')), 120000));
  return await Promise.race([run, limit]);
}"""

COVERAGE = """(tags) => axe.getRules(tags).map(r => ({id: r.ruleId, tags: r.tags}))"""

# describes the focused element; the second read after blur shows whether focus changes anything visible
FOCUS_INFO = """() => {
  const el = document.activeElement;
  if (!el || el === document.body || el === document.documentElement) return null;
  const look = e => { const s = getComputedStyle(e);
    return [s.outlineStyle, s.outlineWidth, s.outlineColor, s.boxShadow, s.borderTopColor, s.borderBottomColor,
            s.backgroundColor, s.color, s.textDecorationLine].join('|'); };
  const focused = look(el);
  const r = el.getBoundingClientRect();
  const name = (el.getAttribute('aria-label') || el.innerText || el.value || el.title || '').trim().replace(/\\s+/g, ' ');
  const hiddenByAria = !!el.closest('[aria-hidden="true"]');
  const twin = el.cloneNode(false);
  const noFocusStyle = (() => {
    const prev = el.style.transition; el.style.transition = 'none';
    el.blur(); const v = look(el); el.focus({preventScroll: true}); el.style.transition = prev; return v; })();
  return {tag: el.tagName.toLowerCase(), role: el.getAttribute('role'), id: el.id || null,
    name: name.slice(0, 80), href: el.getAttribute('href'),
    x: Math.round(r.x), y: Math.round(r.y), w: Math.round(r.width), h: Math.round(r.height),
    inViewport: r.width > 0 && r.height > 0 && r.bottom > 0 && r.right > 0
      && r.top < innerHeight && r.left < innerWidth,
    hiddenByAria, focusVisibleChange: focused !== noFocusStyle,
    inMain: !!el.closest('main, [role=main]'), html: twin.outerHTML.slice(0, 160)};
}"""

# masks the live page before a picture and returns it to normal afterwards; no request leaves the browser
MASK_ON = r"""({terms, blurText}) => {
  const mark = (el, kind) => el.setAttribute('data-eaa-mask', kind);
  for (const el of document.querySelectorAll('body *')) {
    const bg = getComputedStyle(el).backgroundImage;
    if (bg && bg !== 'none' && bg.includes('url(')) mark(el, 'strong');
  }
  document.querySelectorAll('a[href="/"] *, [class*="logo" i], [id*="logo" i]').forEach(e => mark(e, 'strong'));
  document.querySelectorAll('[class*="vendor" i], [class*="brand" i]').forEach(e => mark(e, 'soft'));
  document.querySelectorAll('footer, [id*="footer" i], [class*="footer" i]').forEach(e => {
    if (e.getBoundingClientRect().width > innerWidth * 0.8) mark(e, 'soft'); });
  if (blurText) document.querySelectorAll(
    'main p, main li, main td, main address, main h1, main h2, main h3, main a').forEach(e => mark(e, 'soft'));
  const style = document.createElement('style');
  style.id = 'eaa-mask';
  style.textContent = `img, picture, video, canvas, iframe, object, embed, svg, [data-eaa-mask="strong"]
      {filter: blur(14px) !important}
    [data-eaa-mask="soft"] {filter: blur(6px) !important}`;
  document.head.appendChild(style);
  const pattern = new RegExp(terms.map(t => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|'), 'gi');
  window.__eaaText = [];
  const walk = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  for (let node = walk.nextNode(); node; node = walk.nextNode()) {
    const text = node.nodeValue.replace(pattern, 'Example Store');
    if (text !== node.nodeValue) { window.__eaaText.push([node, node.nodeValue]); node.nodeValue = text; }
  }
}"""

MASK_OFF = """() => {
  document.getElementById('eaa-mask').remove();
  document.querySelectorAll('[data-eaa-mask]').forEach(e => e.removeAttribute('data-eaa-mask'));
  for (const [node, text] of window.__eaaText) node.nodeValue = text;
}"""

FOCUS_BY_NAME = """(wanted) => {
  const clean = t => (t || '').trim().replace(/\\s+/g, ' ');
  for (const el of document.querySelectorAll('a, button, [role=button]')) {
    const name = clean(el.getAttribute('aria-label') || el.innerText || el.textContent);
    if (name.slice(0, 80) === wanted) {
      el.focus();
      return {tag: el.tagName.toLowerCase(), href: el.getAttribute('href')};
    }
  }
  return null;
}"""

WHERE_NOW = """() => {
  const el = document.activeElement;
  const body = !el || el === document.body || el === document.documentElement;
  const r = body ? {x: 0, y: 0, width: 0, height: 0, bottom: 0, right: 0, top: 0, left: 0} : el.getBoundingClientRect();
  const label = body ? '' : (el.getAttribute('aria-label') || el.innerText || el.value || el.title || '');
  return {focusOnBody: body, tag: body ? 'body' : el.tagName.toLowerCase(), id: body ? null : el.id || null,
    role: body ? null : el.getAttribute('role'), name: label.trim().replace(/\\s+/g, ' ').slice(0, 80),
    x: Math.round(r.x), y: Math.round(r.y), scrollY: Math.round(scrollY), hash: location.hash,
    inMain: !body && !!el.closest('main, [role=main]'), inNav: !body && !!el.closest('nav, header, [role=navigation]'),
    inFooter: !body && !!el.closest('footer, [role=contentinfo]'),
    inViewport: r.width > 0 && r.height > 0 && r.bottom > 0 && r.right > 0 && r.top < innerHeight && r.left < innerWidth};
}"""


def robots_for(url):
    root = "{0.scheme}://{0.netloc}".format(urllib.parse.urlparse(url))
    rp = urllib.robotparser.RobotFileParser()
    try:
        req = urllib.request.Request(root + "/robots.txt", headers={"User-Agent": "EAA-Audit-Snapshot/0.1"})
        with urllib.request.urlopen(req, timeout=20) as r:
            rp.parse(r.read().decode("utf-8", "replace").splitlines())
    except OSError:
        rp.parse([])
    return rp


def slim_rule(rule):
    tags = rule["tags"]
    return {
        "id": rule["id"], "impact": rule.get("impact"), "help": rule["help"], "description": rule["description"],
        "helpUrl": rule["helpUrl"], "criteria": criteria_from_tags(tags), "level": level_from_tags(tags),
        "nodeCount": len(rule["nodes"]),
        "nodes": [{"target": n["target"], "html": n["html"][:300], "impact": n.get("impact"),
                   "summary": (n.get("failureSummary") or "")[:400]}
                  for n in rule["nodes"][:MAX_NODES]],
    }


def inject_axe(page, source):
    """Evaluate in every frame: unlike a script tag this is not stopped by the shop's CSP."""
    skipped = []
    for frame in page.frames:
        try:
            frame.evaluate(source + "\n;void 0")
        except PlaywrightError as e:
            skipped.append({"url": frame.url[:120], "error": str(e)[:120]})
    return skipped


def settle(page):
    try:
        page.wait_for_load_state("load", timeout=20000)
    except PlaywrightTimeout:
        pass
    page.wait_for_timeout(2500)
    # scroll once so that lazy images and sections exist for axe and the screenshot
    height = page.evaluate("document.documentElement.scrollHeight")
    for y in range(0, min(height, 12000), 700):
        page.evaluate("y => window.scrollTo(0, y)", y)
        page.wait_for_timeout(120)
    page.evaluate("window.scrollTo(0, 0)")
    page.wait_for_timeout(800)


def reject_cookies(page):
    """Never accepts tracking: only a reject or necessary-only button is pressed."""
    for frame in page.frames:
        try:
            buttons = frame.get_by_role("button", name=REJECT)
            if buttons.count() and buttons.first.is_visible():
                label = buttons.first.inner_text().strip()
                buttons.first.click(timeout=3000)
                page.wait_for_timeout(1000)
                return label
        except PlaywrightError:
            continue
    return None


def banner_still_open(page):
    for frame in page.frames:
        try:
            if frame.get_by_role("button", name=ACCEPT).first.is_visible():
                return True
        except PlaywrightError:
            continue
    return False


def focus_shows_in_pixels(page, stop):
    """Style comparison misses rings drawn by parents or pseudo-elements, so compare the picture too."""
    pad = 10
    clip = {"x": max(stop["x"] - pad, 0), "y": max(stop["y"] - pad, 0),
            "width": min(stop["w"] + 2 * pad, VIEWPORT["width"]), "height": min(stop["h"] + 2 * pad, VIEWPORT["height"])}
    focused = Image.open(io.BytesIO(page.screenshot(clip=clip))).convert("RGB")
    page.evaluate("window.__eaaEl = document.activeElement; window.__eaaEl.blur()")
    page.wait_for_timeout(250)
    plain = Image.open(io.BytesIO(page.screenshot(clip=clip))).convert("RGB")
    page.evaluate("window.__eaaEl.focus({preventScroll: true})")
    page.wait_for_timeout(120)
    return ImageChops.difference(focused, plain).getbbox() is not None


def keyboard_pass(page, shots):
    """Real Tab key events through one page. Nothing is activated here, so no order or form is touched."""
    page.evaluate("window.scrollTo(0, 0)")
    page.mouse.move(0, 0)
    stops = []
    pictures = {}
    body_hits = 0
    finished = False
    ended = f"limit of {MAX_TABS} Tab presses"
    for i in range(1, MAX_TABS + 1):
        page.keyboard.press("Tab")
        page.wait_for_timeout(120)
        info = page.evaluate(FOCUS_INFO)
        if info is None:
            body_hits += 1
            continue
        info["tab"] = i
        info["pixelChange"] = None
        if info["inViewport"] and not info["focusVisibleChange"]:
            info["pixelChange"] = focus_shows_in_pixels(page, info)
        key = [info["tag"], info["name"], info["href"], info["x"], info["y"]]
        if stops and key == stops[0]["key"] and body_hits:
            finished = True  # focus left the page and came back to the first stop: the circle is complete
            ended = "end of page"
            break
        info["repeat"] = any(s["key"] == key for s in stops)
        info["key"] = key
        stops.append(info)
        invisible = not info["focusVisibleChange"] and info["pixelChange"] is False
        for label, hit in (("first_stop", len(stops) == 1), ("no_focus_ring", invisible),
                           ("offscreen", not info["inViewport"]), ("sample_stop", len(stops) == 15)):
            if hit and label not in pictures:
                shot = shots / f"keyboard_{label}.jpg"
                page.screenshot(path=str(shot), type="jpeg", quality=70)
                pictures[label] = {"file": f"screenshots/{shot.name}", "tab": i, "name": info["name"],
                                   "rect": [info["x"], info["y"], info["w"], info["h"]]}
        # inside an embedded frame every stop looks the same from the outside, and with a frame last on the page
        # headless Chromium keeps Tab inside it even without a trap, so this is reported apart from a trap
        in_frame = info["tag"] == "iframe"
        window = FRAME_STOPS if in_frame else 8
        if len(stops) >= window and all(s["repeat"] for s in stops[-(window - 1):]):
            ended = "focus stayed in an embedded frame" if in_frame else "repeating stops"
            break
    no_ring = [s for s in stops if not s["focusVisibleChange"] and s["pixelChange"] is False]
    summary = {
        "tabsPressed": stops[-1]["tab"] if stops else 0,
        "stops": len(stops),
        "skipLinkFirst": bool(stops and re.search(r"skip|jump to|main content", stops[0]["name"], re.I)),
        "tabsToMain": next((s["tab"] for s in stops if s["inMain"]), None),
        "tabsToAddToBag": next((s["tab"] for s in stops if ADD_TO_BAG.search(s["name"])), None),
        "offscreenStops": sum(not s["inViewport"] for s in stops),
        "ariaHiddenStops": sum(s["hiddenByAria"] for s in stops),
        "noRingStops": len(no_ring),
        "repeatStops": sum(s["repeat"] for s in stops),
        "completedCircle": finished,
        "trapSuspected": ended == "repeating stops",
        "reachedTabLimit": ended.startswith("limit"),
        "endedBy": ended,
    }
    return {"summary": summary, "stops": stops, "pictures": pictures}


def try_skip_links(page, stops):
    """Press Enter on each skip link and record where focus lands. Only same-page links are pressed."""
    links = []
    for name in dict.fromkeys(s["name"] for s in stops if SKIP_NAME.match(s["name"])):
        page.evaluate("window.scrollTo(0, 0)")
        page.wait_for_timeout(150)
        link = {"name": name}
        links.append(link)
        found = page.evaluate(FOCUS_BY_NAME, name)
        if not found or (found["tag"] == "a" and not (found["href"] or "").startswith("#")):
            link["pressed"] = False
            link["why"] = "not found again" if not found else f"points to {found['href']}"
            continue
        page.wait_for_timeout(150)
        link["pressed"] = True
        page.keyboard.press("Enter")
        page.wait_for_timeout(600)
        link["after"] = page.evaluate(WHERE_NOW)
        page.keyboard.press("Tab")
        page.wait_for_timeout(200)
        link["nextTab"] = page.evaluate(WHERE_NOW)
    return links


def anon_picture(page, out, name, terms, blur_text):
    folder = out / "screenshots_anon"
    folder.mkdir(exist_ok=True)
    page.evaluate(MASK_ON, {"terms": terms, "blurText": blur_text})
    page.wait_for_timeout(400)
    raw = page.screenshot(full_page=True, type="png")
    page.evaluate(MASK_OFF)
    img = Image.open(io.BytesIO(raw)).convert("RGB")
    img = img.crop((0, 0, img.width, min(img.height, ANON_CROP_PX)))
    path = folder / name
    img.save(path, "JPEG", quality=72)  # saved afresh, so no metadata from the page travels with it
    return f"screenshots_anon/{name}"


def merge_banner_only(after, before):
    """Problems seen only while the cookie dialog was open stay in the results."""
    added = 0
    by_id = {r["id"]: r for r in after}
    for rule in before:
        target = by_id.get(rule["id"])
        if target is None:
            target = {**rule, "nodes": [], "nodeCount": 0}
            after.append(target)
            by_id[rule["id"]] = target
        known = {json.dumps(n["target"]) for n in target["nodes"]}
        for node in rule["nodes"]:
            if json.dumps(node["target"]) not in known:
                target["nodes"].append(node)
                target["nodeCount"] += 1
                added += 1
    return added


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pages")
    ap.add_argument("--shop", required=True)
    ap.add_argument("--out", default=str(HERE / "out"))
    ap.add_argument("--keyboard-page", default="product")
    ap.add_argument("--no-robots", action="store_true", help="only for local test pages")
    args = ap.parse_args()

    plan = json.loads(Path(args.pages).read_text(encoding="utf-8"))
    if len(plan) > MAX_PAGES:
        raise SystemExit(f"{len(plan)} pages, the limit is {MAX_PAGES}")
    out = Path(args.out) / args.shop
    shots = out / "screenshots"
    shots.mkdir(parents=True, exist_ok=True)
    source = AXE_JS.read_text(encoding="utf-8")
    notes_file = out / "notes.json"
    terms = json.loads(notes_file.read_text(encoding="utf-8")).get("brandTerms") if notes_file.exists() else None

    rp = None if args.no_robots else robots_for(plan[0]["url"])
    started = time.time()
    result = {"shop": args.shop, "scannedAt": datetime.now(timezone.utc).isoformat(timespec="seconds"),
              "viewport": VIEWPORT, "tags": AXE_TAGS, "pauseSeconds": PAUSE_S, "pages": [],
              "cookieBanner": None, "keyboard": None}

    with sync_playwright() as p:
        browser = p.chromium.launch()
        result["chromium"] = browser.version
        context = browser.new_context(viewport=VIEWPORT, locale="en-IE")
        page = context.new_page()
        for n, item in enumerate(plan, 1):
            if n > 1:
                time.sleep(PAUSE_S)
            rec = {"n": n, "type": item["type"], "url": item["url"]}
            result["pages"].append(rec)
            if rp and not rp.can_fetch("*", item["url"]):
                rec["error"] = "disallowed by robots.txt"
                print(f"[{n}] skipped by robots.txt: {item['url']}")
                continue
            t0 = time.time()
            try:
                resp = page.goto(item["url"], wait_until="domcontentloaded", timeout=45000)
                rec["status"] = resp.status if resp else None
                settle(page)
                rec["finalUrl"] = page.url
                rec["title"] = page.title()
                rec["pageHeight"] = page.evaluate("document.documentElement.scrollHeight")
                rec["frames"] = len(page.frames)
                rec["framesNotInjected"] = inject_axe(page, source)
                if n == 1:
                    result["axeVersion"] = page.evaluate("axe.version")
                    result["automatedCoverage"] = page.evaluate(COVERAGE, AXE_TAGS)
                axe = page.evaluate(RUN_AXE, AXE_TAGS)
                rec["violations"] = [slim_rule(r) for r in axe["violations"]]
                rec["incomplete"] = [slim_rule(r) for r in axe["incomplete"]]
                if n == 1 or banner_still_open(page):
                    # the dialog is scanned as a visitor meets it, then closed and the page scanned again
                    result["cookieBanner"] = reject_cookies(page)
                    if banner_still_open(page):
                        print("STOP: cookie banner still open, nothing saved, do not repeat", flush=True)
                        raise SystemExit(3)
                    if result["cookieBanner"]:
                        before = rec["violations"]
                        rec["incompleteWithBanner"] = sum(r["nodeCount"] for r in rec["incomplete"])
                        axe = page.evaluate(RUN_AXE, AXE_TAGS)
                        rec["violations"] = [slim_rule(r) for r in axe["violations"]]
                        rec["incomplete"] = [slim_rule(r) for r in axe["incomplete"]]
                        rec["bannerOnlyElements"] = merge_banner_only(rec["violations"], before)
                rec["bannerOpen"] = banner_still_open(page)
                shot = shots / f"{n:02d}_{item['type']}.jpg"
                page.screenshot(path=str(shot), full_page=True, type="jpeg", quality=70)
                rec["screenshot"] = str(shot.relative_to(out)).replace("\\", "/")
                if terms:
                    rec["screenshotAnon"] = anon_picture(page, out, shot.name, terms,
                                                         item["type"] in ("contact", "returns"))
                if item["type"] == args.keyboard_page and result["keyboard"] is None:
                    t1 = time.time()
                    result["keyboard"] = keyboard_pass(page, shots)
                    result["keyboard"]["page"] = item["url"]
                    result["keyboard"]["skipLinks"] = try_skip_links(page, result["keyboard"]["stops"])
                    result["keyboard"]["seconds"] = round(time.time() - t1, 1)
            except (PlaywrightError, PlaywrightTimeout) as e:
                rec["error"] = str(e)[:300]
            rec["seconds"] = round(time.time() - t0, 1)
            print(f"[{n}] {item['type']}: status {rec.get('status')}, {len(rec.get('violations', []))} rules failed, "
                  f"{rec['seconds']} s, {rec.get('error', 'ok')}", flush=True)
        browser.close()

    result["crawlSeconds"] = round(time.time() - started, 1)
    (out / "results.json").write_text(json.dumps(result, indent=1, ensure_ascii=False), encoding="utf-8")
    print(f"saved {out / 'results.json'}, {result['crawlSeconds']} s")


if __name__ == "__main__":
    main()
