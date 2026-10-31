"""Propose a page list from robots.txt and sitemap.xml. Plain requests only, no rendering.

python discover.py https://www.example.ie --out pages.json
Review pages.json before running snapshot.py.
"""
import argparse
import json
import re
import sys
import time
import urllib.parse
import urllib.request
import urllib.robotparser

UA = "EAA-Audit-Snapshot/0.1"
PAUSE_S = 2


def fetch(url):
    time.sleep(PAUSE_S)
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=20) as r:
        return r.read().decode("utf-8", "replace")


def locs(xml):
    return [u.replace("&amp;", "&") for u in re.findall(r"<loc>\s*([^<\s]+)\s*</loc>", xml)]


def pick(urls, words):
    for word in words:
        for u in urls:
            if word in urllib.parse.urlparse(u).path.lower():
                return u
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("site")
    ap.add_argument("--out", default="pages.json")
    ap.add_argument("--search-term", default="shirt")
    args = ap.parse_args()
    root = args.site.rstrip("/")

    robots_txt = fetch(root + "/robots.txt")
    rp = urllib.robotparser.RobotFileParser()
    rp.parse(robots_txt.splitlines())

    index = locs(fetch(root + "/sitemap.xml"))
    children = {k: next((u for u in index if k in u), None) for k in ("collections", "products", "pages")}
    found = {k: locs(fetch(u)) if u else [] for k, u in children.items()}

    # a mid-list product and category are more typical than the first, often legacy, entries
    products = [u for u in found["products"] if "/products/" in u]
    collections = [u for u in found["collections"] if "/collections/" in u and "/all" not in u]
    plan = [
        ("home", root + "/"),
        ("category", pick(collections, ["women", "men", "new", "sale"]) or (collections[0] if collections else None)),
        ("product", products[len(products) // 2] if products else None),
        ("cart", root + "/cart"),
        ("search", root + "/search?q=" + urllib.parse.quote(args.search_term)),
        ("contact", pick(found["pages"], ["contact"])),
        ("returns", pick(found["pages"], ["return", "delivery", "shipping"])),
        ("account", root + "/account/login"),
    ]
    pages = []
    for kind, url in plan:
        if not url:
            print(f"not found: {kind}", file=sys.stderr)
        elif not rp.can_fetch("*", url):
            print(f"robots.txt disallows {kind}: {url}", file=sys.stderr)
        else:
            pages.append({"type": kind, "url": url})
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(pages, f, indent=2)
    print(f"{len(pages)} pages, sitemap: {len(products)} products, {len(collections)} collections, "
          f"{len(found['pages'])} pages -> {args.out}")


if __name__ == "__main__":
    main()
