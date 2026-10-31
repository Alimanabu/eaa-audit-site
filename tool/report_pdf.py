"""Render the report blocks from report.py as a PDF with reportlab.

Limit: reportlab makes untagged PDFs, so images carry no alt text. Headings do become PDF bookmarks.
"""
import io
import os
from xml.sax.saxutils import escape

import reportlab
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (CondPageBreak, Image, LongTable, PageBreak, Paragraph, SimpleDocTemplate, Spacer,
                                TableStyle)

ACCENT = colors.HexColor("#1D4ED8")
GREY = colors.HexColor("#F3F4F6")
LINE = colors.HexColor("#9CA3AF")
FONTS = os.path.join(os.path.dirname(reportlab.__file__), "fonts")
for name, file in (("Vera", "Vera.ttf"), ("Vera-Bold", "VeraBd.ttf"), ("Vera-Italic", "VeraIt.ttf"),
                   ("Vera-BoldItalic", "VeraBI.ttf")):
    pdfmetrics.registerFont(TTFont(name, os.path.join(FONTS, file)))
pdfmetrics.registerFontFamily("Vera", normal="Vera", bold="Vera-Bold", italic="Vera-Italic",
                              boldItalic="Vera-BoldItalic")
GLYPHS = pdfmetrics.getFont("Vera").face.charToGlyph


def safe(text):
    """Escape for reportlab markup; a character the font lacks would otherwise print as a black box."""
    return escape("".join(c if ord(c) in GLYPHS or c == "\n" else "?" for c in str(text)))


BODY = ParagraphStyle("body", fontName="Vera", fontSize=9.5, leading=13.5, spaceAfter=6, alignment=TA_LEFT)
BODY_KEEP = ParagraphStyle("body_keep", parent=BODY, keepWithNext=1)
CODE = ParagraphStyle("code", parent=BODY, fontSize=7, leading=9.5, leftIndent=12, spaceAfter=4,
                      textColor=colors.HexColor("#374151"))
BULLET = ParagraphStyle("bullet", parent=BODY, leftIndent=14, bulletIndent=2, spaceAfter=3, bulletFontName="Vera")
HEADINGS = {
    1: ParagraphStyle("h1", fontName="Vera-Bold", fontSize=16, leading=20, textColor=ACCENT, spaceBefore=12,
                      spaceAfter=8, keepWithNext=1),
    2: ParagraphStyle("h2", fontName="Vera-Bold", fontSize=11.5, leading=15, textColor=ACCENT, spaceBefore=10,
                      spaceAfter=5, keepWithNext=1),
    3: ParagraphStyle("h3", fontName="Vera-Bold", fontSize=9.5, leading=13, spaceBefore=8, spaceAfter=3,
                      keepWithNext=1),
}
# keepWithNext would carry a whole table to the next page and leave a hole; this lets it start here and split
H2_OVER_TABLE = ParagraphStyle("h2", parent=HEADINGS[2], keepWithNext=0)
ROOM_FOR_TABLE = 4 * cm


class ReportDoc(SimpleDocTemplate):
    def afterFlowable(self, flowable):
        if isinstance(flowable, Paragraph) and flowable.style.name in ("h1", "h2"):
            key = f"h{id(flowable)}"
            self.canv.bookmarkPage(key)
            self.canv.addOutlineEntry(flowable.getPlainText(), key, int(flowable.style.name[1]) - 1, closed=False)


def pdf_table(b):
    style = ParagraphStyle("cell", fontName="Vera", fontSize=b["size"] - 1, leading=b["size"] + 1.5)
    head = ParagraphStyle("head", parent=style, fontName="Vera-Bold")
    rows = [[Paragraph(safe(c), head) for c in b["headers"]]]
    rows += [[Paragraph(safe(c), style) for c in row] for row in b["rows"]]
    commands = [("GRID", (0, 0), (-1, -1), 0.4, LINE), ("BACKGROUND", (0, 0), (-1, 0), GREY),
                ("VALIGN", (0, 0), (-1, -1), "TOP"), ("TOPPADDING", (0, 0), (-1, -1), 1.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5), ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3)]
    for y, row in enumerate(b["rows"], 1):
        for x, value in enumerate(row):
            fill = (b["fills"] or {}).get(value)
            if fill and fill != "FFFFFF":
                commands.append(("BACKGROUND", (x, y), (x, y), colors.HexColor("#" + fill)))
    table = LongTable(rows, colWidths=[w * cm for w in b["widths"]], repeatRows=1)
    table.setStyle(TableStyle(commands))
    return table


def pdf_title(b):
    big = ParagraphStyle("title", fontName="Vera-Bold", fontSize=26, leading=32, textColor=ACCENT, spaceAfter=10)
    shop = ParagraphStyle("shop", fontName="Vera", fontSize=15, leading=20, spaceAfter=8)
    story = [Spacer(1, 5 * cm), Paragraph(safe(b["title"]), big), Paragraph(safe(b["shop"]), shop),
             Paragraph(safe(b["subtitle"]), BODY), Spacer(1, 0.6 * cm)]
    story += [Paragraph(safe(line), BODY) for line in b["lines"]]
    story.append(Spacer(1, 0.6 * cm))
    if b["marker"]:
        story.append(Paragraph(f"<b>{safe(b['marker'])}</b>", BODY))
    story += [Paragraph(f"<i>{safe(b['note'])}</i>", BODY), Paragraph(safe(b["legal"]), BODY), PageBreak()]
    return story


def pdf_image(b):
    w, h = PILImage.open(io.BytesIO(b["data"])).size
    width = b["width"] * cm
    return Image(io.BytesIO(b["data"]), width=width, height=width * h / w)


def write_pdf(blocks, path, title, footer):
    story = []
    for n, b in enumerate(blocks):
        kind = b["k"]
        if kind == "title":
            story += pdf_title(b)
        elif kind == "h":
            if b["level"] == 2 and n + 1 < len(blocks) and blocks[n + 1]["k"] == "table":
                story += [CondPageBreak(ROOM_FOR_TABLE), Paragraph(safe(b["text"]), H2_OVER_TABLE)]
            else:
                story.append(Paragraph(safe(b["text"]), HEADINGS[b["level"]]))
        elif kind == "p":
            lead = f"<b>{safe(b['lead'])}</b>" if b["lead"] else ""
            story.append(Paragraph(lead + safe(b["text"]), BODY_KEEP if b["keep"] else BODY))
        elif kind == "ul":
            story += [Paragraph(safe(item), BULLET, bulletText="•") for item in b["items"]]
        elif kind == "code":
            for selector, html in b["items"]:
                story.append(Paragraph(safe(selector) + "<br/>" + safe(html), CODE))
        elif kind == "table":
            story += [pdf_table(b), Spacer(1, 0.35 * cm)]
        elif kind == "img":
            story += [pdf_image(b), Spacer(1, 0.3 * cm)]

    def page_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Vera", 8)
        canvas.setFillColor(colors.HexColor("#4B5563"))
        canvas.drawString(2 * cm, 1.1 * cm, f"{footer} · page {doc.page}")
        canvas.restoreState()

    doc = ReportDoc(str(path), pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm, topMargin=2 * cm,
                    bottomMargin=2 * cm, title=title, author="EAA Audit", subject="WCAG 2.1 AA snapshot",
                    creator="EAA Audit", lang="en-IE")
    doc.build(story, onFirstPage=page_footer, onLaterPages=page_footer)
