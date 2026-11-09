"""Quote and invoice templates for the automated audit (EUR 390), as .docx and .pdf, in English.

python build_templates.py          writes the four files next to this script
python build_templates.py --check  re-reads them and checks the placeholders

Every detail that is not known until the first client is a highlighted placeholder in square brackets: [IP NAME],
[BIN], [IBAN], [SWIFT] and so on. Both formats come from one list of blocks, so they cannot differ. The wording of the
service follows the terms of service on eaa-audit.net (what is included, what is not, 3 business days, refund,
liability limit). Wording that nobody has confirmed (VAT, governing law, payment before work) is marked by a
placeholder or named in the report to the lead.
"""
import re
import sys
from pathlib import Path
from xml.sax.saxutils import escape

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

HERE = Path(__file__).parent
PLACEHOLDER = re.compile(r"\[[^\[\]]+\]")
PRICE = "EUR 390.00"
SITE = "https://eaa-audit.net"

PAYMENT_ROWS = [
    ("Beneficiary", "[IP NAME]"),
    ("BIN", "[BIN]"),
    ("Beneficiary address", "[IP ADDRESS]"),
    ("Bank", "[BANK NAME]"),
    ("Bank address", "[BANK ADDRESS]"),
    ("SWIFT / BIC", "[SWIFT]"),
    ("IBAN (EUR account)", "[IBAN]"),
    ("Intermediary bank, if the bank requires one", "[INTERMEDIARY BANK AND SWIFT, OR \"NONE\"]"),
    ("Currency", "EUR"),
]
BANK_CHARGES = ("Pay by international bank transfer (SWIFT) in EUR. The payer bears all bank charges, including those "
                "of intermediary banks (SWIFT charge option OUR), so that the full amount reaches the beneficiary.")


def quote_blocks():
    return [
        ("title", "Quote: automated accessibility audit"),
        ("kv", [("Quote no.", "[QUOTE NO.]"), ("Date", "[QUOTE DATE]"), ("Valid until", "[QUOTE DATE + 14 DAYS]"),
                ("Prepared for", "[CLIENT COMPANY NAME]"), ("Client address", "[CLIENT ADDRESS]"),
                ("Client contact", "[CLIENT CONTACT NAME AND EMAIL]"), ("Website to be scanned", "[WEBSITE URL]"),
                ("Supplier", "[IP NAME], BIN [BIN], [IP ADDRESS]"), ("Supplier contact", "hello@eaa-audit.net")]),
        ("h", "1. Service"),
        ("p", "An automated check of the website [WEBSITE URL] against WCAG 2.1 levels A and AA, with a written report. "
              "The report is a technical assessment. It does not show that the website conforms to WCAG, EN 301 549 "
              "or the European Accessibility Act, and it is not legal advice."),
        ("h", "2. What is included"),
        ("bullets", [
            "Automated WCAG 2.1 A and AA scan (axe-core rules, Chromium, desktop view) of up to 10 pages agreed with "
            "you, for example home, category, product, cart, search, contact, returns and account sign-in. We confirm "
            "the page list by email before work starts. A page that your robots.txt disallows is skipped and named in "
            "the report.",
            "A scripted keyboard pass on the product page: Tab order, visible focus, suspected traps, skip links.",
            "A report in English as PDF and Word: the five most important problems in plain language (who is "
            "affected, how to fix); every problem found with its level, impact, pages and up to three example "
            "elements; elements repeated across pages counted once; a short “Fix first” list; a result for each "
            "of the 50 WCAG 2.1 A and AA criteria and the checks that still need a person."]),
        ("h", "3. What is not included"),
        ("p", "Testing with a screen reader, a hand-run keyboard test, the filled bag, checkout and payment (we never "
              "place orders or submit forms), the mobile view, WCAG 2.2, an accessibility statement, a re-check after "
              "your fixes, monitoring, and fixing your website."),
        ("h", "4. Delivery"),
        ("p", "The report is delivered by email within 3 business days of receiving your payment. Delivery date: "
              "[DELIVERY DATE]. Your order is confirmed when you accept this quote in writing and we have confirmed the "
              "page list. Work starts when the payment has arrived."),
        ("h", "5. Price"),
        ("table", ["Item", "Price"], [["Automated audit, up to 10 pages (introductory price)", PRICE],
                                      ["Total", PRICE]], [12.0, 4.6]),
        ("p", "Prices are in EUR and exclude VAT where applicable. We confirm the price by this quote before work "
              "starts."),
        ("pagebreak", None),
        ("h", "6. Payment"),
        ("p", "An invoice follows your acceptance. It is payable in full within [PAYMENT TERM IN DAYS] days of the "
              "invoice date, and work starts when the payment has arrived. " + BANK_CHARGES),
        ("kv", PAYMENT_ROWS),
        ("h", "7. Terms, legal details and governing law"),
        ("p", "Supplier: [IP NAME], individual entrepreneur, BIN [BIN], [IP ADDRESS]. The terms of service at "
              + SITE + "/terms.html apply, including the refund if we cannot deliver the report and the limit of our "
              "liability to the price paid. This quote and any contract formed by its acceptance are governed by "
              "[GOVERNING LAW]. Disputes are settled by [COURT OR ARBITRATION AND PLACE]."),
        ("h", "8. How to accept"),
        ("p", "Reply to this email with: “I accept quote [QUOTE NO.] for [WEBSITE URL], and I confirm that I may "
              "allow a scan of this website.” We then send the invoice."),
        ("p", "[IP NAME] · EAA Audit · hello@eaa-audit.net · " + SITE),
    ]


def invoice_blocks():
    return [
        ("title", "Invoice"),
        ("kv", [("Invoice no.", "[INVOICE NO.]"), ("Invoice date", "[INVOICE DATE]"),
                ("Payment due", "[DUE DATE]"), ("Quote no.", "[QUOTE NO.]")]),
        ("h", "Supplier"),
        ("kv", [("Name", "[IP NAME]"), ("BIN", "[BIN]"), ("Address", "[IP ADDRESS]"), ("Email", "hello@eaa-audit.net")]),
        ("h", "Customer"),
        ("kv", [("Company", "[CLIENT COMPANY NAME]"), ("Address", "[CLIENT ADDRESS]"),
                ("VAT or tax ID", "[CLIENT VAT OR TAX ID]"), ("Contact", "[CLIENT CONTACT NAME AND EMAIL]")]),
        ("h", "Services"),
        ("table", ["Description", "Qty", "Unit price", "Amount"],
         [["Automated accessibility audit against WCAG 2.1 A and AA, up to 10 pages, with a report in English "
           "as PDF and Word. Website: [WEBSITE URL]. As per quote [QUOTE NO.].", "1", PRICE, PRICE],
          ["Total due", "", "", PRICE]], [9.0, 1.4, 3.1, 3.1]),
        ("p", "Currency: EUR. VAT: not charged by the supplier; the customer accounts for VAT where it applies in the "
              "customer’s country. [CONFIRM THE VAT WORDING WITH AN ACCOUNTANT BEFORE THE FIRST INVOICE]"),
        ("h", "Payment"),
        ("p", BANK_CHARGES + " Payment reference: invoice [INVOICE NO.]. Payment is due by [DUE DATE]."),
        ("kv", PAYMENT_ROWS),
        ("p", "Terms of service: " + SITE + "/terms.html. Governing law: [GOVERNING LAW]. Disputes: "
              "[COURT OR ARBITRATION AND PLACE]."),
        ("p", "[IP NAME] · EAA Audit · hello@eaa-audit.net · " + SITE),
    ]


DOCS = {"EAA_quote_template": quote_blocks, "EAA_invoice_template": invoice_blocks}


# --- docx ------------------------------------------------------------------------------------------------------------

def add_runs(paragraph, text, bold=False, size=None):
    """Plain text, with each [placeholder] on a yellow highlight so it cannot be missed."""
    pos = 0
    for m in PLACEHOLDER.finditer(text):
        for chunk, mark in ((text[pos:m.start()], False), (m.group(0), True)):
            if chunk:
                run = paragraph.add_run(chunk)
                run.bold = bold or mark
                if size:
                    run.font.size = Pt(size)
                if mark:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        pos = m.end()
    if text[pos:]:
        run = paragraph.add_run(text[pos:])
        run.bold = bold
        if size:
            run.font.size = Pt(size)


def shade(cell, color):
    props = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    props.append(shd)


def write_docx(blocks, path, title):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2)
    sec.top_margin = sec.bottom_margin = Cm(1.8)
    base = doc.styles["Normal"]
    base.font.name = "Arial"
    base.font.size = Pt(10)
    base.paragraph_format.space_after = Pt(5)
    doc.core_properties.title = title
    doc.core_properties.author = "EAA Audit"
    for kind, data, *rest in blocks:
        if kind == "title":
            p = doc.add_paragraph()
            add_runs(p, data, bold=True, size=20)
        elif kind == "h":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.keep_with_next = True
            add_runs(p, data, bold=True, size=11.5)
        elif kind == "pagebreak":
            doc.add_page_break()
        elif kind == "p":
            add_runs(doc.add_paragraph(), data)
        elif kind == "bullets":
            for item in data:
                add_runs(doc.add_paragraph(style="List Bullet"), item)
        elif kind == "kv":
            t = doc.add_table(rows=0, cols=2)
            t.style = "Table Grid"
            for label, value in data:
                row = t.add_row().cells
                row[0].width, row[1].width = Cm(5.6), Cm(11.4)
                add_runs(row[0].paragraphs[0], label, bold=True, size=9.5)
                add_runs(row[1].paragraphs[0], value, size=9.5)
                shade(row[0], "F3F4F6")
            doc.add_paragraph().paragraph_format.space_after = Pt(2)
        elif kind == "table":
            heads, rows, widths = data, rest[0], rest[1]
            t = doc.add_table(rows=1, cols=len(heads))
            t.style = "Table Grid"
            for i, h in enumerate(heads):
                t.rows[0].cells[i].width = Cm(widths[i])
                add_runs(t.rows[0].cells[i].paragraphs[0], h, bold=True, size=9.5)
                shade(t.rows[0].cells[i], "F3F4F6")
            for r in rows:
                cells = t.add_row().cells
                for i, v in enumerate(r):
                    cells[i].width = Cm(widths[i])
                    add_runs(cells[i].paragraphs[0], v, bold=r[0].startswith("Total"), size=9.5)
            doc.add_paragraph().paragraph_format.space_after = Pt(2)
    doc.save(path)


# --- pdf -------------------------------------------------------------------------------------------------------------

def write_pdf(blocks, path, title):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import ListFlowable, ListItem, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    from report_pdf import safe  # registers the Vera fonts

    body = ParagraphStyle("body", fontName="Vera", fontSize=9.2, leading=12.4, spaceAfter=5)
    cell = ParagraphStyle("cell", parent=body, fontSize=8.8, leading=11.4, spaceAfter=0)
    h = ParagraphStyle("h", fontName="Vera-Bold", fontSize=11, leading=14, spaceBefore=8, spaceAfter=4, keepWithNext=1)
    big = ParagraphStyle("big", fontName="Vera-Bold", fontSize=19, leading=24, spaceAfter=8)
    grey = colors.HexColor("#F3F4F6")

    def marked(text, bold=False):
        out, pos = "", 0
        for m in PLACEHOLDER.finditer(text):
            out += safe(text[pos:m.start()]) + f'<font name="Vera-Bold" backColor="#FFF200">{safe(m.group(0))}</font>'
            pos = m.end()
        out += safe(text[pos:])
        return f"<b>{out}</b>" if bold else out

    def grid(rows, widths, head=False):
        t = Table([[Paragraph(marked(c, bold=(head and y == 0) or (not head and x == 0)), cell) for x, c in enumerate(r)]
                   for y, r in enumerate(rows)], colWidths=[w * cm for w in widths])
        style = [("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#9CA3AF")), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                 ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]
        style.append(("BACKGROUND", (0, 0), (-1, 0), grey) if head else ("BACKGROUND", (0, 0), (0, -1), grey))
        t.setStyle(TableStyle(style))
        return t

    story = []
    for kind, data, *rest in blocks:
        if kind == "title":
            story.append(Paragraph(marked(data), big))
        elif kind == "h":
            story.append(Paragraph(marked(data), h))
        elif kind == "pagebreak":
            story.append(PageBreak())
        elif kind == "p":
            story.append(Paragraph(marked(data), body))
        elif kind == "bullets":
            story.append(ListFlowable([ListItem(Paragraph(marked(i), body), leftIndent=12) for i in data],
                                      bulletType="bullet", start="•", leftIndent=12, bulletFontName="Vera"))
        elif kind == "kv":
            story += [grid([list(r) for r in data], [5.6, 11.4]), Spacer(1, 4)]
        elif kind == "table":
            widths = rest[1]
            story += [grid([data] + rest[0], widths, head=True), Spacer(1, 4)]
    SimpleDocTemplate(str(path), pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm, topMargin=1.8 * cm,
                      bottomMargin=1.6 * cm, title=title, author="EAA Audit").build(story)


# --- build and check -------------------------------------------------------------------------------------------------

def build():
    for stem, make in DOCS.items():
        blocks = make()
        title = stem.replace("_", " ")
        write_docx(blocks, HERE / f"{stem}.docx", title)
        write_pdf(blocks, HERE / f"{stem}.pdf", title)
        print(f"saved {stem}.docx and {stem}.pdf")


def docx_text(path):
    doc = Document(path)
    parts = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for row in t.rows:
            parts += [c.text for c in row.cells]
    return "\n".join(parts)


def pdf_text(path):
    import subprocess
    return subprocess.run(["pdftotext", "-enc", "UTF-8", str(path), "-"], capture_output=True, text=True,
                          encoding="utf-8", errors="replace").stdout


NEEDED = {"EAA_quote_template": ["[IP NAME]", "[BIN]", "[IBAN]", "[SWIFT]", "[QUOTE NO.]", "[QUOTE DATE]",
                                 "[GOVERNING LAW]", "[CLIENT COMPANY NAME]", "[WEBSITE URL]"],
          "EAA_invoice_template": ["[IP NAME]", "[BIN]", "[IBAN]", "[SWIFT]", "[INVOICE NO.]", "[INVOICE DATE]",
                                   "[DUE DATE]", "[GOVERNING LAW]", "[CLIENT COMPANY NAME]", "[WEBSITE URL]"]}
LEFTOVERS = re.compile(r"\{\{|\}\}|_{3,}|\bX{3,}\b|TODO|TBD|\.{4,}|<[A-Z ]+>|\bLorem\b")


def check():
    failed = 0
    for stem in DOCS:
        texts = {"docx": docx_text(HERE / f"{stem}.docx"), "pdf": pdf_text(HERE / f"{stem}.pdf")}
        sets = {}
        for kind, text in texts.items():
            flat = re.sub(r"\s+", " ", text)
            sets[kind] = set(PLACEHOLDER.findall(flat))
            problems = [f"missing {p}" for p in NEEDED[stem] if p not in flat]
            problems += [f"leftover {m.group(0)!r}" for m in LEFTOVERS.finditer(text)]
            # an open bracket without its close would be a placeholder that cannot be told from text
            problems += ["unbalanced square brackets"] if flat.count("[") != flat.count("]") else []
            print(f"{'OK  ' if not problems else 'FAIL'} {stem}.{kind}: {len(sets[kind])} placeholders")
            for p in problems:
                print("     -", p)
            failed += bool(problems)
        if sets["docx"] != sets["pdf"]:
            print(f"FAIL {stem}: docx and pdf differ: {sorted(sets['docx'] ^ sets['pdf'])}")
            failed += 1
        print("     ", ", ".join(sorted(sets["docx"])))
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    sys.path.insert(0, str(HERE.parent))
    check() if "--check" in sys.argv else build()
