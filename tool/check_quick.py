"""Acceptance check for the one-page snapshot made by quick.py.

python check_quick.py out/quick/shop.nl [out/quick/other.com ...]

For each shop folder it checks:
  1. the PDF has exactly 1 page;
  2. the numbers printed in the PDF (pages, violations, distinct elements, criteria, and per problem elements and pages)
     equal the numbers counted again from the raw axe results in results.json, not from summary.json;
  3. no word "fine", "penalty" or "penalties" and no euro sign except the price of the audit (EUR 390, once);
  4. the PDF carries https://eaa-audit.net and one picture per problem;
  5. in each picture there is a red frame, the frame's element is one the axe results name, and OCR reads no
     cookie-dialog wording;
Exit code 1 when anything fails. Needs Poppler (pdftotext, pdfinfo, pdfimages) and Tesseract.
"""
import json
import re
import subprocess
import sys
from pathlib import Path

from PIL import Image

OCR = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
FRAME = (225, 29, 72)
COOKIE_WORDS = re.compile(r"cookie|consent|accept|weiger|refus|akkoord|toestemming|privacy|rejeter|alles", re.I)
FORBIDDEN = re.compile(r"\bfines?\b|penalt", re.I)
PRICE = "\u20ac390"


def run(*cmd):
    return subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace").stdout


def raw_numbers(results):
    """Counted straight from the axe rules saved in results.json."""
    pages = [p for p in results["pages"] if "violations" in p]
    per_rule = {}
    elements, criteria, total = set(), set(), 0
    for p in pages:
        for rule in p["violations"]:
            total += rule["nodeCount"]
            criteria.update(rule["criteria"])
            r = per_rule.setdefault(rule["id"], {"elements": set(), "pages": set(), "selectors": set()})
            r["pages"].add(p["n"])
            for node in rule["nodes"]:
                key = json.dumps(node["target"])
                elements.add(key)
                r["elements"].add(key)
                r["selectors"].add(key)
    return {"pages": len(pages), "violations": total, "elements": len(elements), "criteria": len(criteria),
            "rules": per_rule}


def text_problems(text, expected):
    """Pure text checks; returns a list of what is wrong."""
    bad = []
    flat = re.sub(r"\s+", " ", text)
    for label, pattern, want in (
            ("pages", r"(\d+) pages checked", expected["pages"]),
            ("violations", r"(\d+) violations found on", expected["violations"]),
            ("distinct elements", r"violations found on (\d+) distinct elements", expected["elements"]),
            ("criteria", r"(\d+) of 50 WCAG 2\.1 A/AA criteria not passed", expected["criteria"])):
        m = re.search(pattern, flat)
        if not m:
            bad.append(f"{label}: not found in the PDF text")
        elif int(m.group(1)) != want:
            bad.append(f"{label}: PDF says {m.group(1)}, axe results say {want}")
    if FORBIDDEN.search(text):
        bad.append(f"forbidden word: {FORBIDDEN.search(text).group(0)}")
    if text.count("\u20ac") != 1 or text.count(PRICE) != 1:
        bad.append(f"euro sign must appear once, in {PRICE}: found {text.count(chr(0x20ac))} sign(s), {text.count(PRICE)} price(s)")
    if "https://eaa-audit.net" not in flat:
        bad.append("link https://eaa-audit.net missing")
    return bad


def check_folder(folder):
    folder = Path(folder)
    bad = []
    pdfs = list(folder.glob("*.pdf"))
    if len(pdfs) != 1:
        return [f"expected 1 PDF in {folder}, found {len(pdfs)}"]
    pdf = pdfs[0]
    results = json.loads((folder / "results.json").read_text(encoding="utf-8"))
    summary = json.loads((folder / "summary.json").read_text(encoding="utf-8"))
    expected = raw_numbers(results)

    pages = re.search(r"Pages:\s+(\d+)", run("pdfinfo", str(pdf)))
    if not pages or int(pages.group(1)) != 1:
        bad.append(f"PDF has {pages.group(1) if pages else '?'} pages, must be 1")
    text = run("pdftotext", "-enc", "UTF-8", str(pdf), "-")
    bad += text_problems(text, expected)
    if not 3 <= expected["pages"] <= 5:
        bad.append(f"{expected['pages']} pages scanned, must be 3-5")

    flat = re.sub(r"\s+", " ", text)
    findings = summary["findings"]
    for f in findings:
        rule = expected["rules"].get(f["rule"])
        if not rule:
            bad.append(f"problem {f['rule']} is not in the axe results")
            continue
        want = f"Found on {len(rule['elements'])} element{'s' if len(rule['elements']) != 1 else ''}, on {len(rule['pages'])} of {expected['pages']} pages"
        if want not in flat:
            bad.append(f"{f['rule']}: PDF text lacks '{want}'")
        if json.dumps([f["selector"]]) not in rule["selectors"]:
            bad.append(f"{f['rule']}: picture selector {f['selector']!r} is not an element of this rule in the axe results")
        picture = folder / "pictures" / f["picture"]
        img = Image.open(picture).convert("RGB")
        # JPEG blurs the exact colour a little: count pixels within 60 of the frame colour in each channel
        red = sum(1 for px in img.getdata() if all(abs(a - b) < 60 for a, b in zip(px, FRAME)))
        if red < 300:
            bad.append(f"{f['picture']}: no red frame ({red} px)")
        ocr = run(OCR, str(picture), "stdout", "--psm", "11")
        if COOKIE_WORDS.search(ocr):
            bad.append(f"{f['picture']}: cookie wording readable: {COOKIE_WORDS.search(ocr).group(0)!r}")
    shown = len(re.findall(r"^\s*\d+\s+\d+\s+image", run("pdfimages", "-list", str(pdf)), re.M))
    if shown != len(findings):
        bad.append(f"{shown} pictures in the PDF, {len(findings)} problems")
    if len(findings) < 3:  # not a failure: a shop with few visible problems cannot show three honest pictures
        print(f"     note: {len(findings)} problems with a clean picture; skipped: {summary.get('skippedRules')}")
    return bad


def main():
    folders = sys.argv[1:]
    if not folders:
        raise SystemExit(__doc__)
    failed = 0
    for folder in folders:
        problems = check_folder(folder)
        print(f"{'OK  ' if not problems else 'FAIL'} {folder}")
        for p in problems:
            print("     -", p)
        failed += bool(problems)
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
