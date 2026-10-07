"""Build Linear_Algebra_Week2_Lecture_Notes.pdf: one entry per slide (slide image + notes).

Run after the deck PDF exists:
    pdftoppm -jpeg -r 100 Linear_Algebra_Week2_Slides.pdf notes_img/slide
    python3 build_notes_pdf.py
"""
import glob
import os

import matplotlib
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle)

from lecture_notes import NOTES

FONT_DIR = os.path.join(os.path.dirname(matplotlib.__file__), "mpl-data", "fonts", "ttf")
pdfmetrics.registerFont(TTFont("DejaVu", os.path.join(FONT_DIR, "DejaVuSans.ttf")))
pdfmetrics.registerFont(TTFont("DejaVu-Bold", os.path.join(FONT_DIR, "DejaVuSans-Bold.ttf")))
pdfmetrics.registerFont(TTFont("DejaVu-Italic", os.path.join(FONT_DIR, "DejaVuSans-Oblique.ttf")))
pdfmetrics.registerFont(TTFont("DejaVuSerif-Bold", os.path.join(FONT_DIR, "DejaVuSerif-Bold.ttf")))
from reportlab.pdfbase.pdfmetrics import registerFontFamily
registerFontFamily("DejaVu", normal="DejaVu", bold="DejaVu-Bold", italic="DejaVu-Italic", boldItalic="DejaVu-Bold")

NAVY = colors.HexColor("#14213D")
TEAL = colors.HexColor("#1F6F8B")
GREY = colors.HexColor("#5B6675")

H1 = ParagraphStyle("h1", fontName="DejaVuSerif-Bold", fontSize=22, leading=27, textColor=NAVY, spaceAfter=6)
H2 = ParagraphStyle("h2", fontName="DejaVuSerif-Bold", fontSize=12.5, leading=16, textColor=NAVY, spaceAfter=4)
EYE = ParagraphStyle("eye", fontName="DejaVu-Bold", fontSize=8, leading=10, textColor=TEAL)
BODY = ParagraphStyle("body", fontName="DejaVu", fontSize=9.6, leading=13.8, textColor=colors.HexColor("#222933"),
                      alignment=TA_LEFT, spaceAfter=6)
SMALL = ParagraphStyle("small", fontName="DejaVu", fontSize=9, leading=13, textColor=GREY)

OUT = "Linear_Algebra_Week2_Lecture_Notes.pdf"
IMGS = sorted(glob.glob("notes_img/slide-*.jpg"))
assert len(IMGS) == len(NOTES), (len(IMGS), len(NOTES))

PAGE_W, PAGE_H = letter
MARGIN = 0.75 * inch
CONTENT_W = PAGE_W - 2 * MARGIN
IMG_W = 2.9 * inch
IMG_H = IMG_W * 9 / 16
GAP = 0.25 * inch


def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("DejaVu", 8)
    canvas.setFillColor(GREY)
    canvas.drawString(MARGIN, 0.5 * inch, "Math for Data Science  ·  Linear Algebra, Week 2  ·  Lecture notes")
    canvas.drawRightString(PAGE_W - MARGIN, 0.5 * inch, str(doc.page))
    canvas.restoreState()


doc = SimpleDocTemplate(OUT, pagesize=letter, leftMargin=MARGIN, rightMargin=MARGIN, topMargin=0.8 * inch,
                        bottomMargin=0.85 * inch, title="Linear Algebra Week 2 – Lecture Notes",
                        author="Math for Data Science")
story = []

# Title block
story.append(Paragraph("MATH FOR DATA SCIENCE · LINEAR ALGEBRA, WEEK 2", EYE))
story.append(Spacer(1, 4))
story.append(Paragraph("Lecture Notes: Solving Systems of Linear Equations", H1))
story.append(Paragraph(
    "One entry per slide of the Week 2 deck. Each entry shows the slide, what to say, and the mathematics worked out in full "
    "(every row operation and every arithmetic step), so the notes can be used to teach from, or handed to students as a study guide. "
    "The deck has six parts: solving systems, matrix row reduction, rank, row echelon form, reduced row echelon form, and Gaussian "
    "elimination, with a machine-learning motivation up front and three in-class quizzes.", BODY))
story.append(Spacer(1, 10))

parts = {3: "Motivation", 6: "Part 1 · Solving systems of linear equations", 22: "Part 2 · Matrix row reduction",
         27: "Part 3 · Rank of a matrix", 37: "Part 4 · Row echelon form", 45: "Part 5 · Reduced row echelon form",
         50: "Part 6 · The Gaussian elimination algorithm", 57: "Recap"}

for n in sorted(NOTES):
    title, paras = NOTES[n]
    head = []
    if n in parts:
        head = [Spacer(1, 6), Paragraph(parts[n].upper(), EYE), Spacer(1, 4)]
    img = Image(IMGS[n - 1], width=IMG_W, height=IMG_H)
    img.hAlign = "LEFT"
    text_cell = [Paragraph(f"Slide {n} · {title}", H2)] + [Paragraph(p, BODY) for p in paras]
    tbl = Table([[img, text_cell]], colWidths=[IMG_W, CONTENT_W - IMG_W - GAP], hAlign="LEFT")
    tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (0, 0), GAP),
        ("RIGHTPADDING", (1, 0), (1, 0), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("LINEBELOW", (0, 0), (-1, 0), 0.4, colors.HexColor("#D5DAE2")),
    ]))
    story.append(KeepTogether(head + [tbl, Spacer(1, 12)]))

doc.build(story, onFirstPage=footer, onLaterPages=footer)
print("wrote", OUT)
