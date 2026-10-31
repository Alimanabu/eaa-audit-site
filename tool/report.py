"""Build the English client report from results.json and notes.json, as .docx and .pdf.

python report.py out/sample-001 [--anon-only]
Writes the internal version; if notes.json has an "anon" section, also the version without the shop's identity.
"""
import argparse
import io
import json
import re
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont

from report_pdf import FONTS, write_pdf
from wcag import BY_ID, CRITERIA, IMPACT_ORDER, PLAIN, criteria_from_tags

ACCENT = RGBColor(0x1D, 0x4E, 0xD8)
GREY = "F3F4F6"
FILL = {"Fail": "FEE2E2", "Needs review": "FEF3C7", "No automated failure": "F3F4F6", "Manual only": "FFFFFF"}
TOP_CROP_PX = 1100
SHOT_CM = 12.5  # two page pictures fit on one A4 page
CELL_PX = 6
INK_BELOW = 225
PHOTO_MIN_PX = 120
SCRAP_AREA_PX = 900
RING_PX = 6
PLACEHOLDER = (209, 213, 219)
PLAQUE = (229, 231, 235)
LABEL = (75, 85, 99)
ATTRIBUTES_KEPT = {"role", "href", "src", "alt", "type"}  # plus aria-*: the rest is class names and vendor hooks
AUTOMATION_NOTE = (
    "Automated testing checks only about 30–40 % of WCAG criteria. This report therefore does not show that "
    "the shop conforms to WCAG 2.1 AA or to the European Accessibility Act. A clean automated result on a "
    "criterion is not a pass: everything marked “Manual check required” below has not been tested."
)
SCREEN_READER = [
    "Browse a category and open a product with NVDA or VoiceOver",
    "Choose a size or colour and hear the selected value announced",
    "Add a product to the bag and hear the confirmation (4.1.3 Status Messages)",
    "Open and close the menu, filters, mini-bag and pop-ups; check where focus goes",
    "Complete the checkout up to the payment step, including error messages",
    "Check that headings and landmarks allow fast navigation of each page type",
]
KEYBOARD = [
    "Operate every control with Enter, Space, arrow keys and Esc (the script pressed Enter on skip links only)",
    "Open the main menu, filters, size selector and image gallery by keyboard",
    "Complete the purchase path: product, bag, delivery, payment",
    "Check focus is kept inside open dialogs and returns to the trigger when they close",
    "Repeat on all remaining page types",
    "Confirm visually that focus order follows the visual order",
]
DYNAMIC_CONTRAST = [
    "Text over photos and video banners, including carousels at each slide",
    "Hover, focus, active and disabled states of buttons and links",
    "Error and success messages, sale prices and badges",
    "Non-text contrast of icons, form borders and focus rings (1.4.11)",
    "Cookie banner, pop-ups, drop-downs and the mobile menu",
]


def count(n, word):
    return f"{n} {word}" + ("" if n == 1 else "s")


def elements_key(node):
    # the selector alone: shared templates repeat it, while html can differ (an animated placeholder, say)
    return json.dumps(node["target"])


def aggregate(pages, field):
    rules = {}
    for page in pages:
        for v in page.get(field, []):
            r = rules.setdefault(v["id"], {**{k: v[k] for k in ("id", "impact", "help", "description", "helpUrl",
                                                                "criteria", "level")},
                                           "pages": [], "instances": 0, "unique": {}})
            if IMPACT_ORDER[v["impact"]] > IMPACT_ORDER[r["impact"]]:
                r["impact"] = v["impact"]
            r["pages"].append(page["type"])
            r["instances"] += v["nodeCount"]
            for node in v["nodes"]:
                r["unique"].setdefault(elements_key(node), node)
    return sorted(rules.values(), key=lambda r: (-IMPACT_ORDER[r["impact"]], -len(r["unique"]), r["id"]))


def criteria_rows(failed, review, coverage):
    has_rule = {sc for rule in coverage for sc in criteria_from_tags(rule["tags"])}
    rows = []
    for sc, level, name, _ in CRITERIA:
        f = [r for r in failed if sc in r["criteria"]]
        i = [r for r in review if sc in r["criteria"]]
        if f:
            status = "Fail"
        elif i:
            status = "Needs review"
        elif sc in has_rule:
            status = "No automated failure"
        else:
            status = "Manual only"
        rows.append({"sc": sc, "level": level, "name": name, "status": status,
                     "rules": ", ".join(r["id"] for r in f or i),
                     "elements": sum(len(r["unique"]) for r in f or i)})
    return rows, len(has_rule)


def summarise(result):
    pages = [p for p in result["pages"] if "violations" in p]
    failed = aggregate(pages, "violations")
    review = aggregate(pages, "incomplete")
    rows, covered = criteria_rows(failed, review, result.get("automatedCoverage", []))
    levels = {}
    for lvl in ("A", "AA"):
        rs = [r for r in failed if r["level"] == lvl]
        levels[lvl] = {"problemTypes": len(rs), "instances": sum(r["instances"] for r in rs),
                       "uniqueElements": len({k for r in rs for k in r["unique"]})}
    # one element can break several rules, so the total is a union, not a sum
    levels["total"] = {"problemTypes": len(failed), "instances": sum(r["instances"] for r in failed),
                       "uniqueElements": len({k for r in failed for k in r["unique"]})}
    return {"pages": pages, "failed": failed, "review": review, "rows": rows, "levels": levels,
            "criteriaWithRule": covered, "failedCriteria": sum(r["status"] == "Fail" for r in rows)}


def cropped_jpeg(path):
    img = Image.open(path).convert("RGB")
    if img.height > TOP_CROP_PX:
        img = img.crop((0, 0, img.width, TOP_CROP_PX))
    img.thumbnail((1000, 1400))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=72)
    return buf.getvalue()


def blocks_on_page(img):
    """Boxes of everything that is not white background, found on a coarse grid, largest first.
    A cell counts when most of it is ink, so the soft edge of a blurred photo does not bridge the gap to the next."""
    ink = img.convert("L").point(lambda v: 255 if v < INK_BELOW else 0)
    size = (img.width // CELL_PX, img.height // CELL_PX)
    cells = ink.resize(size, Image.BOX).point(lambda v: 255 if v >= 128 else 0).load()
    seen, boxes = set(), []
    for start in ((x, y) for y in range(size[1]) for x in range(size[0])):
        if start in seen or not cells[start]:
            continue
        seen.add(start)
        stack, xs, ys = [start], [], []
        while stack:
            x, y = stack.pop()
            xs.append(x)
            ys.append(y)
            for near in ((x - 1, y), (x + 1, y), (x, y - 1), (x, y + 1)):
                if 0 <= near[0] < size[0] and 0 <= near[1] < size[1] and near not in seen and cells[near]:
                    seen.add(near)
                    stack.append(near)
        boxes.append((min(xs) * CELL_PX, min(ys) * CELL_PX, (max(xs) + 1) * CELL_PX, (max(ys) + 1) * CELL_PX))
    return sorted(boxes, key=lambda b: (b[0] - b[2]) * (b[3] - b[1]))


def wireframe_jpeg(path):
    """The page reduced to grey blocks: no pixel of the shop remains, only where its parts stood.
    Blur alone leaves words in the menus and the shapes of the models in the photos."""
    img = Image.open(path).convert("RGB")
    img = img.crop((0, 0, img.width, min(img.height, TOP_CROP_PX)))
    plan = Image.new("RGB", img.size, "white")
    draw = ImageDraw.Draw(plan)
    font = ImageFont.truetype(f"{FONTS}/Vera.ttf", 22)
    kept = []
    for x0, y0, x1, y1 in blocks_on_page(img):
        inside = any(a <= x0 and b <= y0 and c >= x1 and d >= y1 for a, b, c, d in kept)
        if inside or (x1 - x0) * (y1 - y0) < SCRAP_AREA_PX:  # letter-sized scraps read as text
            continue
        kept.append((x0, y0, x1, y1))
        photo = x1 - x0 >= PHOTO_MIN_PX and y1 - y0 >= PHOTO_MIN_PX
        draw.rectangle((x0, y0, x1, y1), fill=PLACEHOLDER if photo else PLAQUE)
        if photo:
            draw.text(((x0 + x1) / 2, (y0 + y1) / 2), "image", fill=LABEL, font=font, anchor="mm")
    plan.thumbnail((1000, 1400))
    buf = io.BytesIO()
    plan.save(buf, "JPEG", quality=85)
    return buf.getvalue()


def fragment_jpeg(path, rect, zoom=2):
    """A tight crop around one focused element. Its inside is covered and only the edge stays:
    the focus ring, or the lack of one, is what the picture is there to show."""
    x, y, w, h = rect
    img = Image.open(path).convert("RGB").crop((x, y, x + w, y + h))
    ImageDraw.Draw(img).rectangle((RING_PX, RING_PX, w - RING_PX, h - RING_PX), fill=PLAQUE)
    img = img.resize((img.width * zoom, img.height * zoom), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=85)
    return buf.getvalue(), img.width * 13.5 / 1366


def neutral_element(selector, html, labels):
    """An element for the shareable report: class names and vendor attributes can be searched for across the web,
    so the selector becomes a plain description and the markup keeps only what the finding is about."""
    label = next((text for pattern, text in labels if re.search(pattern, selector)), None)
    if not label:  # no description written: the tag path is all that stays
        label = re.sub(r"[.#][\w-]+|\[[^\]]*\]|:[\w()-]+", "", selector).strip(" >") or "page element"

    def keep(attribute):
        name = attribute.group(1).lower()
        return attribute.group(0) if name in ATTRIBUTES_KEPT or name.startswith("aria-") else ""

    def clean(tag):
        return re.sub(r'\s([\w:-]+)(?:="[^"]*")?', keep, tag.group(0))
    return label, re.sub(r"<[^>]+>", clean, html)[:200]


def make_scrub(anon, domain):
    neutral = anon.get("neutral", [])
    terms = sorted(anon.get("terms", []), key=len, reverse=True)
    pattern = re.compile("|".join(re.escape(t) for t in terms), re.I) if terms else None
    address = re.compile(rf"https?://\S+|\S*(?:{re.escape(domain)}|shopify\.com)\S*")

    def scrub(text):
        for vendor_wording, plain in neutral:
            text = text.replace(vendor_wording, plain)
        text = address.sub("[address withheld]", text)
        text = re.sub(r'\b(href|src)([$^*]?=)\\?"[^"]*"', r'\1\2"..."', text)
        text = re.sub(r"#?shopify-section\S*", "[section]", text)
        text = re.sub(r"\d{6,}", "[id]", text)
        return pattern.sub(anon["short"], text) if pattern else text
    return scrub


def blocks_for(result, notes, s, out, anon):
    """The whole report as plain blocks, so that .docx and .pdf cannot drift apart."""
    a = notes["anon"] if anon else {}
    shop = a.get("name") or notes.get("shopName", result["shop"])
    date = datetime.fromisoformat(result["scannedAt"]).strftime("%d %B %Y")
    pages = s["pages"]
    B = []

    def h(level, text):
        B.append({"k": "h", "level": level, "text": text})

    def para(text, lead=None, keep=False):
        B.append({"k": "p", "text": text, "lead": lead, "keep": keep})

    def bullets(items):
        if items:
            B.append({"k": "ul", "items": list(items)})

    def table(headers, rows, widths, size=9, fills=None):
        B.append({"k": "table", "headers": headers, "rows": [[str(c) for c in r] for r in rows],
                  "widths": widths, "size": size, "fills": fills})

    def picture(data, width_cm, alt):
        B.append({"k": "img", "data": data, "width": width_cm, "alt": alt})

    def element_lines(node):
        selector, html = " > ".join(str(x) for x in node["target"]), " ".join(node["html"].split())
        return neutral_element(selector, html, a.get("elementLabels", [])) if anon else (selector[:140], html[:200])

    def keyboard_pictures():
        for key, pic in (result.get("keyboard") or {}).get("pictures", {}).items():
            if anon and key not in a.get("keyboardFragments", []):
                continue
            h(3, f"Keyboard pass, {key.replace('_', ' ')}: Tab press {pic['tab']}, "
                 f"{pic['name'] or 'element without a name'}")
            if anon:
                data, width = fragment_jpeg(out / pic["file"], pic["rect"])
                picture(data, width, f"Focused element at Tab press {pic['tab']}")
            else:
                picture(cropped_jpeg(out / pic["file"]), SHOT_CM,
                        f"Viewport during the keyboard pass at Tab press {pic['tab']}")

    def fix_first():
        by_id = {r["id"]: r for r in s["failed"]}
        h(2, "Fix first")
        for n, item in enumerate(notes["fixFirst"], 1):
            rules = [by_id[i] for i in item["rules"]]
            worst = max(rules, key=lambda r: IMPACT_ORDER[r["impact"]])
            elements = len({k for r in rules for k in r["unique"]})
            spread = len({p for r in rules for p in r["pages"]})
            para(f" {item['text']} ({worst['impact']}, level {worst['level']}; "
                 f"{count(elements, 'element')}, {spread} of {len(pages)} pages)", lead=f"{n}.")

    B.append({"k": "title", "title": "Accessibility Snapshot Report", "shop": shop,
              "subtitle": "WCAG 2.1 Level A and AA · European Accessibility Act",
              "lines": [f"Scan date: {date}", f"Pages checked: {len(pages)}",
                        f"Tools: axe-core {result.get('axeVersion')}, Chromium {result.get('chromium')}, "
                        "scripted keyboard pass on one page", "Prepared by: EAA Audit"],
              "note": a.get("note") or notes.get("sampleNote", "Sample report."),
              "marker": None if anon else "INTERNAL: names the shop. Do not share outside the team.",
              "legal": "This is a technical assessment, not legal advice."})

    h(1, "1. Summary")
    if notes.get("fixFirst"):
        fix_first()
    para(AUTOMATION_NOTE, lead="How far this report goes. ")
    if s["criteriaWithRule"]:
        pct = round(100 * s["criteriaWithRule"] / len(CRITERIA))
        para(f"In this scan the automated rules touched {s['criteriaWithRule']} of the {len(CRITERIA)} WCAG 2.1 A and "
             f"AA criteria ({pct} %), and most of those only in part: the rules can find a missing image description, "
             "but not a wrong one.")
    la, laa, total = s["levels"]["A"], s["levels"]["AA"], s["levels"]["total"]
    para(f"The scan of {len(pages)} pages found {count(total['problemTypes'], 'type')} of problem on "
         f"{count(total['uniqueElements'], 'distinct element')} ({count(total['instances'], 'occurrence')} across "
         f"pages) and failed {s['failedCriteria']} WCAG criteria. Elements that repeat across pages usually come "
         "from a shared template, so one fix can solve several pages.")
    table(["Level", "Problem types", "Distinct elements", "Occurrences on all pages"],
          [["A", la["problemTypes"], la["uniqueElements"], la["instances"]],
           ["AA", laa["problemTypes"], laa["uniqueElements"], laa["instances"]],
           ["Total", total["problemTypes"], total["uniqueElements"], total["instances"]]], [3, 4, 4, 6], size=10)
    if s["review"]:
        para(f"A further {count(len(s['review']), 'rule type')} on {sum(len(r['unique']) for r in s['review'])} "
             "elements could not be decided automatically (see section 3). They are not counted above.")

    h(2, "The five most important problems")
    top = s["failed"][:5]
    for n, r in enumerate(top, 1):
        ptitle, who, fix = PLAIN.get(r["id"], (r["help"], r["description"], "See the rule guidance: " + r["helpUrl"]))
        para(f" ({r['impact']}, level {r['level']}, WCAG {', '.join(r['criteria'])}; "
             f"{count(len(r['unique']), 'element')} on {len(r['pages'])} of {len(pages)} pages)",
             lead=f"{n}. {ptitle}")
        para("Who is affected: " + who)
        para("How to fix: " + fix)
        extra = notes.get("ruleNotes", {}).get(r["id"])
        if extra:
            para("In this shop: " + extra)
    # fewer than five automated problem types: the rest comes from the keyboard pass
    for n, item in enumerate(notes.get("extraTop", [])[:5 - len(top)], len(top) + 1):
        para(f" (manual finding, {item['scope']})", lead=f"{n}. {item['title']}")
        para("Who is affected: " + item["who"])
        para("How to fix: " + item["fix"])
        para("In this shop: " + item["evidence"])

    h(1, "2. Results by WCAG criterion")
    para("Fail: at least one automated rule failed. Needs review: the automated rule could not decide. "
         "No automated failure: rules ran and found nothing, but they cover only part of the criterion, "
         "so a manual check is still required. Manual only: no automated rule exists.")
    table(["WCAG", "Level", "Criterion", "Result", "Rules", "Elements"],
          [[r["sc"], r["level"], r["name"], r["status"], r["rules"], r["elements"] or ""] for r in s["rows"]],
          [1.3, 1.1, 5.2, 3.0, 4.7, 1.6], size=8, fills=FILL)

    h(2, "Findings in detail")
    for r in s["failed"]:
        h(3, PLAIN.get(r["id"], (r["help"],))[0])
        para(f"Rule {r['id']} · {r['impact']} · level {r['level']} · WCAG {', '.join(r['criteria'])} "
             f"· {count(len(r['unique']), 'distinct element')}, {count(r['instances'], 'occurrence')} · "
             f"pages: {', '.join(dict.fromkeys(r['pages']))}")
        B.append({"k": "code", "items": [element_lines(node) for node in list(r["unique"].values())[:3]]})

    h(1, "3. Results by page")
    rows = []
    for p in pages:
        counts = {lvl: [v for v in p["violations"] if v["level"] == lvl] for lvl in ("A", "AA")}
        row = [p["n"], p["type"], p.get("status"), len(counts["A"]), len(counts["AA"]),
               sum(v["nodeCount"] for v in p["violations"]), sum(v["nodeCount"] for v in p["incomplete"])]
        if not anon:
            row.insert(2, p["finalUrl"].split("//", 1)[-1][:58])
        rows.append(row)
    if anon:
        table(["#", "Page", "HTTP", "Rules failed A", "Rules failed AA", "Elements", "To review"],
              rows, [1, 4, 1.6, 2.8, 2.8, 2.4, 2.4], size=9)
    else:
        table(["#", "Page", "Address", "HTTP", "Rules failed A", "Rules failed AA", "Elements", "To review"],
              rows, [0.7, 1.8, 6.4, 1.1, 1.6, 1.7, 1.6, 1.6], size=8)
    for p in result["pages"]:
        if p.get("error"):
            para(f"Not checked: {p['type']} ({p['url']}): {p['error'].splitlines()[0]}")
    for line in notes.get("pageNotes", []):
        para(line)
    if s["review"]:
        h(2, "Could not be decided automatically")
        table(["Rule", "Level", "WCAG", "Elements", "Why it needs a person"],
              [[r["id"], r["level"], ", ".join(r["criteria"]), len(r["unique"]), r["help"]] for r in s["review"]],
              [4, 1.2, 1.8, 1.6, 8.4], size=8)

    h(1, "4. What must be checked by hand")
    para("The items below were not done in this snapshot. Each is marked “Manual check required”. "
         "Until they are done, the shop’s conformance is unknown.")
    kb = result.get("keyboard")
    h(2, "Keyboard")
    if kb:
        k = kb["summary"]
        where = "the product page" if anon else kb["page"].split("//", 1)[-1]
        para(f"Done: one scripted pass on {where} with real Tab key presses and nothing bought or submitted. "
             f"{k['stops']} focus stops; Tab presses to reach the main content: {k['tabsToMain'] or 'not reached'}; "
             f"to reach Add to bag: {k['tabsToAddToBag'] or 'not reached'}; stops outside the visible area: "
             f"{k['offscreenStops']}; stops inside content hidden from screen readers: {k['ariaHiddenStops']}. "
             f"The pass ended by: {k['endedBy']}. This pass did not involve a person with a keyboard.")
        for link in kb.get("skipLinks", []):
            if not link.get("pressed"):
                para(f"Skip link “{link['name']}”: not pressed ({link.get('why')}).")
                continue
            land, nxt = link["after"], link["nextTab"]
            area = "the main content" if land["inMain"] else "the footer" if land["inFooter"] else "another area"
            para(f"Skip link “{link['name']}” pressed with Enter: focus moved to a {land['tag']} element in "
                 f"{area}; the next Tab reached “{nxt['name'][:40]}”.")
        bullets(notes.get("keyboardNotes", []))
        if anon:
            keyboard_pictures()
    para("Manual check required:")
    bullets(KEYBOARD)
    h(2, "Screen reader")
    para("Not done. Manual check required:")
    bullets(SCREEN_READER)
    h(2, "Contrast in changing states")
    para("Automated contrast rules read one static state of the page. Manual check required:")
    bullets(DYNAMIC_CONTRAST)
    h(2, "All other criteria that automation cannot settle")
    table(["WCAG", "Criterion", "Manual check required"],
          [[r["sc"], r["name"], BY_ID[r["sc"]][3]] for r in s["rows"] if r["status"] != "Fail"],
          [1.5, 6, 9.5], size=8)
    h(2, "Not covered by this snapshot")
    bullets(notes.get("limits", []) + [
        "The filled bag, checkout and payment steps",
        "Mobile view and the accessibility statement required by the European Accessibility Act",
    ])

    h(1, "Appendix: screenshots")
    if anon:
        para("Desktop view reduced to its layout. Text and photos are replaced by grey blocks, so nothing here "
             "identifies the shop. Only the top of each page is shown.")
    else:
        para(f"Desktop view, {result['viewport']['width']} by {result['viewport']['height']} pixels. Long pages are "
             "cropped to their top part; full-length images are kept with the audit data.")
    for p in pages:
        source = p.get("screenshotAnon") if anon else p.get("screenshot")
        if not source:
            continue
        h(2, f"{p['n']}. {p['type'].capitalize()}")
        if not anon:
            para(p["finalUrl"], keep=True)
        data = wireframe_jpeg(out / source) if anon else cropped_jpeg(out / source)
        what = "Layout" if anon else "Screenshot"
        picture(data, SHOT_CM, f"{what} of the {p['type']} page of {shop}, top part")
    if not anon:
        keyboard_pictures()

    if anon:
        scrub = make_scrub(a, result["shop"].split(".")[0])
        for b in B:
            for key in ("text", "lead", "alt", "shop", "note"):
                if b.get(key):
                    b[key] = scrub(b[key])
            if b["k"] == "ul":
                b["items"] = [scrub(x) for x in b["items"]]
            if b["k"] == "code":
                b["items"] = [(scrub(x), scrub(y)) for x, y in b["items"]]
            if b["k"] == "table":
                b["rows"] = [[scrub(c) for c in r] for r in b["rows"]]
    return B


def shade(cell, color):
    fill = OxmlElement("w:shd")
    fill.set(qn("w:val"), "clear")
    fill.set(qn("w:color"), "auto")
    fill.set(qn("w:fill"), color)
    cell._tc.get_or_add_tcPr().append(fill)


def docx_table(doc, b):
    t = doc.add_table(rows=1, cols=len(b["headers"]))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    t.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))  # repeats on each page, marks the header
    for row in [b["headers"], *b["rows"]]:
        cells = t.rows[0].cells if row is b["headers"] else t.add_row().cells
        for i, value in enumerate(row):
            cell = cells[i]
            cell.width = Cm(b["widths"][i])
            run = cell.paragraphs[0].add_run(value)
            run.font.size = Pt(b["size"])
            cell.paragraphs[0].paragraph_format.space_after = Pt(0)
            if row is b["headers"]:
                run.bold = True
                shade(cell, GREY)
            elif b["fills"] and b["fills"].get(value):
                shade(cell, b["fills"][value])
    doc.add_paragraph()


def field(paragraph, code):
    for kind, text in (("begin", None), (None, code), ("end", None)):
        run = paragraph.add_run()
        if kind:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        else:
            el = OxmlElement("w:instrText")
            el.text = text
        run._r.append(el)


def write_docx(blocks, path, title, footer):
    doc = Document()
    base = doc.styles["Normal"]
    base.font.name = "Calibri"
    base.font.size = Pt(11)
    for lang in doc.styles.element.iter(qn("w:lang")):
        lang.set(qn("w:val"), "en-IE")
    for name, size in (("Heading 1", 18), ("Heading 2", 13)):
        doc.styles[name].font.name = "Calibri"
        doc.styles[name].font.size = Pt(size)
        doc.styles[name].font.color.rgb = ACCENT
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.left_margin = section.right_margin = Cm(2)
    section.top_margin = section.bottom_margin = Cm(2)
    doc.core_properties.title = title
    doc.core_properties.author = "EAA Audit"
    doc.core_properties.language = "en-IE"
    foot = section.footer.paragraphs[0]
    foot.add_run(footer + " · page ")
    field(foot, "PAGE")
    for run in foot.runs:
        run.font.size = Pt(9)

    for b in blocks:
        kind = b["k"]
        if kind == "title":
            for _ in range(5):
                doc.add_paragraph()
            run = doc.add_paragraph().add_run(b["title"])
            run.bold, run.font.size, run.font.color.rgb = True, Pt(30), ACCENT
            doc.add_paragraph().add_run(b["shop"]).font.size = Pt(18)
            doc.add_paragraph(b["subtitle"])
            doc.add_paragraph()
            for line in b["lines"]:
                doc.add_paragraph(line)
            doc.add_paragraph()
            if b["marker"]:
                doc.add_paragraph().add_run(b["marker"]).bold = True
            doc.add_paragraph().add_run(b["note"]).italic = True
            doc.add_paragraph(b["legal"])
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        elif kind == "h":
            doc.add_heading(b["text"], b["level"])
        elif kind == "p":
            p = doc.add_paragraph()
            if b["lead"]:
                p.add_run(b["lead"]).bold = True
            p.add_run(b["text"])
            p.paragraph_format.keep_with_next = b["keep"]
        elif kind == "ul":
            for item in b["items"]:
                doc.add_paragraph(item, style="List Bullet")
        elif kind == "code":
            for selector, html in b["items"]:
                run = doc.add_paragraph(style="List Bullet").add_run(selector + "\n" + html)
                run.font.name = "Consolas"
                run.font.size = Pt(8)
        elif kind == "table":
            docx_table(doc, b)
        elif kind == "img":
            doc.add_picture(io.BytesIO(b["data"]), width=Cm(b["width"]))
            doc.paragraphs[-1]._p.xpath(".//wp:docPr")[0].set("descr", b["alt"])
    doc.save(path)


def build(out, anon_only=False):
    result = json.loads((out / "results.json").read_text(encoding="utf-8"))
    notes_file = out / "notes.json"
    notes = json.loads(notes_file.read_text(encoding="utf-8")) if notes_file.exists() else {}
    s = summarise(result)

    versions = [("internal", False, f"EAA_snapshot_{out.name}_internal", notes.get("shopName", result["shop"]),
                 "EAA Audit · INTERNAL, do not share")]
    if notes.get("anon"):
        versions.append(("anon", True, "EAA_snapshot_" + notes["anon"]["slug"], notes["anon"]["name"],
                         "EAA Audit · sample report"))
    for _, anon, stem, shop, footer in versions:
        if anon_only and not anon:
            continue
        blocks = blocks_for(result, notes, s, out, anon)
        title = f"Accessibility snapshot: {shop}"
        write_docx(blocks, out / f"{stem}.docx", title, footer)
        write_pdf(blocks, out / f"{stem}.pdf", title, footer)
        print(f"saved {stem}: docx {(out / f'{stem}.docx').stat().st_size // 1024} KB, "
              f"pdf {(out / f'{stem}.pdf').stat().st_size // 1024} KB")

    (out / "summary.json").write_text(json.dumps({
        "levels": s["levels"], "failedCriteria": s["failedCriteria"], "criteriaWithRule": s["criteriaWithRule"],
        "pagesChecked": len(s["pages"]), "needsReviewRules": len(s["review"]),
        "needsReviewElements": sum(len(r["unique"]) for r in s["review"]),
        "rules": [{"id": r["id"], "impact": r["impact"], "level": r["level"], "criteria": r["criteria"],
                   "pages": len(r["pages"]), "instances": r["instances"], "uniqueElements": len(r["unique"]),
                   "documented": r["id"] in PLAIN} for r in s["failed"]],
    }, indent=1), encoding="utf-8")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--anon-only", action="store_true", help="leave the internal version as it is")
    args = ap.parse_args()
    build(Path(args.folder), args.anon_only)
