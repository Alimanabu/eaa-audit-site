"""Check the shareable report for traces of the shop: text, metadata, alt text and text inside pictures.
Besides the listed terms, any word readable in a picture is a trace unless it is general interface wording.
Names of third-party services and the shop's own CSS classes count too: both can be searched for across the web.

python check_anon.py out/sample-001
Exit code 1 means a trace was found. Needs Poppler (pdftotext, pdfimages, pdfinfo) and Tesseract.
"""
import argparse
import json
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from PIL import Image

OCR = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


INTERFACE_WORDS = {"search", "menu", "cart", "image"}  # image: the caption of our own placeholders
VENDORS = ["klarna", "clearpay", "afterpay", "paypal", "stripe", r"klevu\w*", "algolia", "doofinder", "searchanise",
           "gorgias", "tidio", "zendesk", "klaviyo", r"yotpo\w*", "trustpilot", "cookiebot", "onetrust", "userway",
           "accessibe", "shopify", "woocommerce", "magento", "bigcommerce"]
VENDOR_TRACE = re.compile(r"\b(?:" + "|".join(VENDORS) + r")\b", re.I)
VENDOR_CLASS = re.compile(r"\bku(?:[A-Z]\w*|[-_]\w+)")  # search widget classes; case matters: plain "ku" is no trace
SHOP_CLASS = re.compile(r"__|[a-z][A-Z]|\w-\w+-\w")  # a built name, not an ordinary word such as "tiles" or "mini-bag"


def tool(*cmd):
    return subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace").stdout


def class_names(result):
    """CSS classes of the elements the scan reported: the report must not repeat them."""
    names = set()
    for page in result["pages"]:
        for field in ("violations", "incomplete"):
            for rule in page.get(field, []):
                for node in rule["nodes"]:
                    for value in re.findall(r'class="([^"]*)"', node["html"]):
                        names.update(value.split())
                    for part in node["target"]:
                        names.update(re.findall(r"\.([A-Za-z_][\w-]*)", str(part)))
    return sorted(n for n in names if SHOP_CLASS.search(n))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    out = Path(ap.parse_args().folder)
    notes = json.loads((out / "notes.json").read_text(encoding="utf-8"))
    result = json.loads((out / "results.json").read_text(encoding="utf-8"))
    anon = notes["anon"]
    stem = out / f"EAA_snapshot_{anon['slug']}"
    words = [t for t in notes["brandTerms"]] + anon["traces"] + ["shopify.com", result["shop"].split(".")[0]]
    words += re.findall(r"data-id=\"(\d{6,})\"", json.dumps(result).replace('\\"', '"'))[:40]
    for page in result["pages"]:  # the last part of each address: product and collection names
        tail = page["url"].rstrip("/").rsplit("/", 1)[-1].split("?")[0]
        if len(tail) > 8:
            words.append(tail)
    words += class_names(result)
    trace = re.compile("|".join(re.escape(w) for w in dict.fromkeys(words)), re.I)
    found = []
    read = {}

    def scan(where, text):
        for pattern in (trace, VENDOR_TRACE, VENDOR_CLASS):
            for m in pattern.finditer(text):
                found.append((where, text[max(m.start() - 30, 0):m.end() + 30].replace("\n", " ")))

    with zipfile.ZipFile(f"{stem}.docx") as z:
        for name in z.namelist():
            if name.endswith((".xml", ".rels")):
                scan(f"docx:{name}", z.read(name).decode("utf-8", "replace"))
    layer = tool("pdftotext", "-enc", "UTF-8", f"{stem}.pdf", "-")
    scan("pdf:text", layer)
    scan("pdf:wrapped", layer.replace("\n", ""))  # a long class name can break across two lines
    scan("pdf:info", tool("pdfinfo", f"{stem}.pdf"))

    pictures = 0
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        tool("pdfimages", "-j", f"{stem}.pdf", str(tmp / "pdf"))
        with zipfile.ZipFile(f"{stem}.docx") as z:
            for name in z.namelist():
                if "media/" in name:
                    (tmp / ("docx_" + Path(name).name)).write_bytes(z.read(name))
        big = tmp / "big.png"
        for image in sorted(tmp.iterdir()):
            pictures += 1
            source = Image.open(image).convert("RGB")
            source.resize((source.width * 3, source.height * 3), Image.LANCZOS).save(big)
            for file in (image, big):  # small words come out only when enlarged
                for psm in ("11", "3"):
                    text = tool(OCR, str(file), "-", "--psm", psm)
                    scan(f"ocr:{image.name}", text)
                    for word in re.findall(r"[A-Za-z]{3,}", text):
                        read.setdefault(word.lower(), set()).add(image.name)
            if "docx_" in image.name:
                exif = image.read_bytes()[:400]
                if b"Exif" in exif:
                    found.append((f"exif:{image.name}", "metadata present"))

    for word in sorted(read.keys() - INTERFACE_WORDS):
        found.append((f"ocr-word:{sorted(read[word])[0]}", word))
    print(f"searched for {len(dict.fromkeys(words))} terms, pictures read by OCR: {pictures}")
    print("words read in pictures:", {w: len(read[w]) for w in sorted(read)})
    for where, text in found:
        print("TRACE", where, repr(text))
    print("traces:", len(found))
    sys.exit(1 if found else 0)


if __name__ == "__main__":
    main()
