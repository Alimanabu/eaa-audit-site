# Accessibility snapshot tool

Lives on the `tool` branch and is never merged into `main`: GitHub Pages publishes `main`.

Needs Playwright with Chromium, python-docx, Pillow and reportlab. Run with the Python that has them:
`%LOCALAPPDATA%\Programs\Python\Python314\python.exe`. `vendor/axe.min.js` is axe-core 4.14.0 (MPL-2.0).

1. `python discover.py https://www.shop.ie --out pages.json --search-term jeans`: proposes pages from robots.txt and the sitemap.
2. Review `pages.json`: at most 10 pages, a product with size options.
3. `python snapshot.py pages.json --shop shop.ie`: one pass, 2 s pause. Exit code 3 and `STOP: cookie banner still open` mean nothing was saved: do not repeat, look at the button text first. Put `brandTerms` in notes.json beforehand to get masked pictures in `screenshots_anon/`.
4. Rename `out/shop.ie` to a neutral `out/sample-NNN`, so that no file name carries the shop, and write `notes.json` there after reading the results (see `out/sample-001`: `tool/out/` is in `.gitignore` because it holds unmasked pictures, addresses and the shop's name; only the shareable `EAA_snapshot_example-store.pdf` and `.docx` are committed, as the reference sample, the rest stays on this computer).
5. `python report.py out/sample-001`: writes internal .docx and .pdf, summary.json and, if notes.json has `anon`, the versions without the shop's identity. All formats come from one list of blocks. In the shareable version every picture is a plan of grey blocks, each element is a plain description instead of a selector (`anon.elementLabels`), and vendor wording in notes is replaced from `anon.neutral`. `--anon-only` leaves the internal version as it is.
6. Run `check_anon` with `Python314\python.exe`: the `python` from Git Bash has no Pillow. `python check_anon.py out/sample-001`: searches the shareable files for traces, with OCR on the pictures at normal size and enlarged 3 times. Any word readable in a picture counts as a trace, except Search, Menu, Cart and the `image` caption. Names from `VENDORS`, `ku…` widget classes and the CSS classes reported in `results.json` count as traces in the text too; a new vendor goes into `VENDORS`. Exit code 1 means a trace was found.

The keyboard pass presses Tab, and Enter on same-page skip links only. The tool never places orders or submits forms. PDFs are untagged: reportlab cannot tag them.

## Free one-page snapshot for a cold email (`quick.py`)

One shop, one command, one PDF: `python quick.py https://www.shop.nl` writes `out/quick/shop.nl/EAA_free_snapshot_shop.nl.pdf` (1 page A4: pages checked, violations, WCAG criteria not passed, the 3 biggest problems each with a framed picture of the element, what is wrong, who it affects, a line about the full audit and https://eaa-audit.net). The same folder holds `results.json` (raw axe results, same format as `snapshot.py`), `summary.json` and `pictures/`. It needs about 1 to 1.5 minutes per shop.

Pages: home, one category, one product, the cart; a contact or delivery page is added when the cart cannot be found or when fewer than 3 problems have a clean picture (at most 5 pages). The scan rules are those of `snapshot.py`: robots.txt, 2 s pause, cookies only rejected (in English, Dutch, French and German), nothing ordered or submitted. Before a picture is taken, what is left of a cookie dialog and every fixed or sticky bar that does not hold the element are hidden, so the picture shows the element and nothing on top of it.

Many shops: `python quick.py --list shops.xlsx` (`.csv` or `.xlsx`, the column named site, website, url, domain or сайт, or give `--column NAME`). Each shop runs in its own process with a limit of 420 s (`--timeout`); a broken address or a shop that blocks the scan (HTTP 403 is common on large chains) is written to the summary and the run goes on. The summary is `out/quick/summary_<time>.csv` with the columns store, pages, violations, criteria, error, pdf, seconds.

`python check_quick.py out/quick/shop.nl ...` is the acceptance check: 1 page; the numbers in the PDF equal the numbers counted again from `results.json`; no "fine", "penalty" or euro sign except the price of the audit; a red frame and no cookie wording in each picture. `python -m unittest test_check_quick` holds the negative tests.

## Quote and invoice templates (`templates/`)

`python templates/build_templates.py` writes `EAA_quote_template` and `EAA_invoice_template` as .docx and .pdf. Everything unknown until the first client is a yellow placeholder in square brackets ([IP NAME], [BIN], [IBAN], [SWIFT], [GOVERNING LAW] ...); `--check` re-reads all four files and compares the placeholders. Replace the brackets in the .docx, save as PDF, send.
