"""Build the UM AI presentation as a PPTX — visual edition with architecture diagrams."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ---- Palette ----
CVS_RED    = RGBColor(0xCC, 0x00, 0x00)
CVS_RED_D  = RGBColor(0x99, 0x00, 0x00)
DARK       = RGBColor(0x14, 0x18, 0x22)
INK        = RGBColor(0x22, 0x28, 0x33)
GRAY       = RGBColor(0x60, 0x6A, 0x77)
LIGHT      = RGBColor(0xF2, 0xF4, 0xF7)
SOFT       = RGBColor(0xE8, 0xEC, 0xF2)
ACCENT     = RGBColor(0x00, 0x55, 0x99)
ACCENT_D   = RGBColor(0x00, 0x3A, 0x6B)
CYAN       = RGBColor(0x00, 0xA8, 0xC8)
GREEN      = RGBColor(0x2E, 0x8B, 0x57)
AMBER      = RGBColor(0xF4, 0xA3, 0x00)
PURPLE     = RGBColor(0x6B, 0x4B, 0xA8)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
MIST       = RGBColor(0xDD, 0xE2, 0xEA)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SLIDE_W, SLIDE_H = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


# ===== Helpers =====
def add_rect(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    shp = slide.shapes.add_shape(shape, x, y, w, h)
    shp.shadow.inherit = False
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line; shp.line.width = Pt(0.75)
    shp.text_frame.text = ""
    return shp


def add_text(slide, x, y, w, h, text, size=14, bold=False, color=INK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font="Calibri",
             italic=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left=Emu(0); tf.margin_right=Emu(0); tf.margin_top=Emu(0); tf.margin_bottom=Emu(0)
    tf.vertical_anchor = anchor
    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run(); run.text = line
        run.font.size = Pt(size); run.font.bold = bold; run.font.italic = italic
        run.font.color.rgb = color; run.font.name = font
    return tb


def add_bullets(slide, x, y, w, h, items, size=14, color=INK, font="Calibri",
                bullet="•", line_spacing=1.15):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left=Emu(0); tf.margin_right=Emu(0); tf.margin_top=Emu(0); tf.margin_bottom=Emu(0)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(3)
        p.line_spacing = line_spacing
        if isinstance(item, tuple):
            label, body = item
            r0 = p.add_run(); r0.text = f"{bullet}  "
            r0.font.size=Pt(size); r0.font.color.rgb=color; r0.font.name=font
            r1 = p.add_run(); r1.text = label
            r1.font.size=Pt(size); r1.font.bold=True; r1.font.color.rgb=color; r1.font.name=font
            r2 = p.add_run(); r2.text = body
            r2.font.size=Pt(size); r2.font.color.rgb=color; r2.font.name=font
        else:
            r = p.add_run(); r.text = f"{bullet}  {item}"
            r.font.size=Pt(size); r.font.color.rgb=color; r.font.name=font
    return tb


def add_pill(slide, x, y, w, h, text, fill, txt_color=WHITE, size=11, bold=True):
    shp = add_rect(slide, x, y, w, h, fill=fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    tf = shp.text_frame; tf.word_wrap = True
    tf.margin_left=Inches(0.05); tf.margin_right=Inches(0.05)
    tf.margin_top=Inches(0.02); tf.margin_bottom=Inches(0.02)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=txt_color; r.font.name="Calibri"
    return shp


def add_node(slide, x, y, w, h, title, subtitle=None, fill=WHITE, border=ACCENT,
             title_color=INK, sub_color=GRAY, title_size=12, sub_size=9):
    """Box with title and optional subtitle."""
    shp = add_rect(slide, x, y, w, h, fill=fill, line=border, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    shp.line.width = Pt(1.5)
    tf = shp.text_frame; tf.word_wrap=True
    tf.margin_left=Inches(0.08); tf.margin_right=Inches(0.08)
    tf.margin_top=Inches(0.06); tf.margin_bottom=Inches(0.06)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = title
    r.font.size=Pt(title_size); r.font.bold=True; r.font.color.rgb=title_color; r.font.name="Calibri"
    if subtitle:
        p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
        p2.space_before = Pt(2)
        r2 = p2.add_run(); r2.text = subtitle
        r2.font.size=Pt(sub_size); r2.font.color.rgb=sub_color; r2.font.name="Calibri"
    return shp


def add_arrow(slide, x1, y1, x2, y2, color=ACCENT, weight=2.0):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    c.line.color.rgb = color
    c.line.width = Pt(weight)
    # Add arrowhead via XML
    ln = c.line._get_or_add_ln()
    tail = etree.SubElement(ln, qn('a:tailEnd'))
    tail.set('type', 'triangle')
    tail.set('w', 'med'); tail.set('len', 'med')
    return c


def add_header(slide, kicker, title, accent=CVS_RED):
    # Left red bar (vertical accent)
    add_rect(slide, 0, 0, Inches(0.18), SLIDE_H, fill=accent)
    # Top thin line
    add_rect(slide, Inches(0.18), 0, SLIDE_W - Inches(0.18), Inches(0.06), fill=DARK)
    add_text(slide, Inches(0.55), Inches(0.30), Inches(12), Inches(0.3),
             kicker, size=11, bold=True, color=accent)
    add_text(slide, Inches(0.55), Inches(0.55), Inches(12), Inches(0.6),
             title, size=28, bold=True, color=DARK)


def add_footer(slide, page_num, total=15):
    add_text(slide, Inches(0.55), Inches(7.15), Inches(6), Inches(0.3),
             "UM × AI · Internal · Draft", size=9, color=GRAY)
    add_text(slide, Inches(11.5), Inches(7.15), Inches(1.3), Inches(0.3),
             f"{page_num} / {total}", size=9, color=GRAY, align=PP_ALIGN.RIGHT)


def add_table(slide, x, y, w, h, data, header_fill=DARK, header_color=WHITE,
              body_size=11, header_size=11, col_widths=None, alt=True):
    rows, cols = len(data), len(data[0])
    tbl_shape = slide.shapes.add_table(rows, cols, x, y, w, h)
    tbl = tbl_shape.table
    if col_widths:
        total = sum(col_widths)
        for i, cw in enumerate(col_widths):
            tbl.columns[i].width = int(w * cw / total)
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.margin_left=Inches(0.1); cell.margin_right=Inches(0.1)
            cell.margin_top=Inches(0.05); cell.margin_bottom=Inches(0.05)
            tf = cell.text_frame; tf.word_wrap=True; tf.text=""
            p = tf.paragraphs[0]
            run = p.add_run(); run.text = str(val); run.font.name="Calibri"
            if r == 0:
                cell.fill.solid(); cell.fill.fore_color.rgb = header_fill
                run.font.bold=True; run.font.color.rgb=header_color; run.font.size=Pt(header_size)
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if (not alt or r % 2 == 1) else LIGHT
                run.font.color.rgb=INK; run.font.size=Pt(body_size)
    return tbl


def add_stat(slide, x, y, w, h, number, label, color=CVS_RED, bg=LIGHT):
    add_rect(slide, x, y, w, h, fill=bg, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # accent bar
    add_rect(slide, x, y, Inches(0.08), h, fill=color)
    add_text(slide, x + Inches(0.2), y + Inches(0.18), w - Inches(0.3), Inches(0.7),
             number, size=28, bold=True, color=color)
    add_text(slide, x + Inches(0.2), y + Inches(0.95), w - Inches(0.3), Inches(0.5),
             label, size=10, color=INK)


# =============== SLIDE 1 — Title (dramatic) ===============
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, SLIDE_W, SLIDE_H, fill=DARK)
# Diagonal accent bands
for i, (yy, col) in enumerate([(Inches(2.6), CVS_RED), (Inches(2.72), CYAN), (Inches(2.78), ACCENT)]):
    add_rect(s, 0, yy, SLIDE_W, Inches(0.04 if i else 0.08), fill=col)

add_text(s, Inches(0.7), Inches(0.6), Inches(12), Inches(0.4),
         "■ CVS HEALTH · UM × AI", size=12, bold=True, color=CVS_RED)
add_text(s, Inches(0.7), Inches(1.7), Inches(12), Inches(1),
         "UTILIZATION MANAGEMENT × AI", size=14, bold=True, color=CYAN)
add_text(s, Inches(0.7), Inches(2.0), Inches(12), Inches(2),
         "From unstructured\nclinical documents\nto AI-assisted decisions.",
         size=52, bold=True, color=WHITE)
add_text(s, Inches(0.7), Inches(5.4), Inches(12), Inches(0.5),
         "The technology is solvable. The human interface is the real work.",
         size=18, color=RGBColor(0xCF, 0xD6, 0xDF), italic=True)
add_text(s, Inches(0.7), Inches(6.9), Inches(12), Inches(0.4),
         "CDST · Clinical Decision Support Tool   ·   Inpatient · SNF · Spine · Pre-cert",
         size=11, color=RGBColor(0x90, 0x9A, 0xA8))


# =============== SLIDE 2 — Problem stats ===============
s = prs.slides.add_slide(BLANK)
add_header(s, "01 · THE PROBLEM", "Why UM, Why Now")
add_text(s, Inches(0.55), Inches(1.25), Inches(12.3), Inches(0.4),
         "UM ensures the right care, at the right time, in the right setting — using evidence-based criteria.",
         size=14, color=GRAY, italic=True)

add_stat(s, Inches(0.55), Inches(2.0), Inches(2.95), Inches(1.5), "$250M", "Annual admin budget", CVS_RED)
add_stat(s, Inches(3.65), Inches(2.0), Inches(2.95), Inches(1.5), "2.26M", "Cases reviewed / year", CVS_RED)
add_stat(s, Inches(6.75), Inches(2.0), Inches(2.95), Inches(1.5), "20 – 60 min", "Per case review", ACCENT)
add_stat(s, Inches(9.85), Inches(2.0), Inches(2.95), Inches(1.5), "3–50 × 10–200", "Docs × pages / case", ACCENT)

# Pressure box
add_rect(s, Inches(0.55), Inches(3.85), Inches(12.25), Inches(2.85), fill=LIGHT,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_rect(s, Inches(0.55), Inches(3.85), Inches(0.12), Inches(2.85), fill=CVS_RED)
add_text(s, Inches(0.85), Inches(4.0), Inches(12), Inches(0.4),
         "PRESSURE STACKING UP", size=12, bold=True, color=CVS_RED)
add_bullets(s, Inches(0.85), Inches(4.45), Inches(12), Inches(2.3), [
    ("Faxed PDFs, often not searchable — ", "no copy/paste, no cross-doc search natively"),
    ("CMS 2026 Final Rule — ", "Two-Midnight alignment, FHIR PA APIs, 72hr urgent / 7-day standard decisions"),
    ("Rising Medicare cost trend — ", "volume up while admin budget is already strained"),
    ("Nurses working overtime — ", "majority of work is concurrent review; burnout is real"),
], size=13)
add_footer(s, 2)


# =============== SLIDE 3 — Workflow chevrons ===============
s = prs.slides.add_slide(BLANK)
add_header(s, "02 · WHERE TIME GOES", "The Clinical Reality")
add_text(s, Inches(0.55), Inches(1.3), Inches(12.3), Inches(0.5),
         "Open MedCompass → open 3–50 docs one by one → scan 10–200 pages each → transcribe into a case note.",
         size=14, color=GRAY, italic=True)

labels = [("OPEN", "MedCompass"), ("DOWNLOAD", "Docs one-by-one"),
          ("SCAN", "Often unsearchable"), ("TRANSCRIBE", "Into case note")]
colors = [ACCENT, ACCENT_D, CVS_RED, CVS_RED_D]
x0 = Inches(0.55); cw = Inches(3.0); gap = Inches(0.13); y0 = Inches(2.2); ch = Inches(1.5)
for i, ((k, v), color) in enumerate(zip(labels, colors)):
    shp = s.shapes.add_shape(MSO_SHAPE.CHEVRON, x0 + (cw + gap) * i, y0, cw, ch)
    shp.fill.solid(); shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    tf = shp.text_frame; tf.word_wrap=True; tf.text=""
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p1 = tf.paragraphs[0]; p1.alignment=PP_ALIGN.CENTER
    r1 = p1.add_run(); r1.text = k
    r1.font.size=Pt(11); r1.font.bold=True; r1.font.color.rgb=RGBColor(0xFF,0xCC,0xCC) if i>=2 else RGBColor(0xC9,0xDF,0xF0)
    p2 = tf.add_paragraph(); p2.alignment=PP_ALIGN.CENTER
    r2 = p2.add_run(); r2.text = v
    r2.font.size=Pt(16); r2.font.bold=True; r2.font.color.rgb=WHITE

# Big punch line
add_rect(s, Inches(0.55), Inches(4.2), Inches(12.25), Inches(1.1),
         fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, Inches(0.55), Inches(4.3), Inches(12.25), Inches(0.9),
         "The hardest part isn't \"the AI\".",
         size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text(s, Inches(0.55), Inches(4.75), Inches(12.25), Inches(0.5),
         "It's the workflow, the interface, the trust.",
         size=18, color=CYAN, align=PP_ALIGN.CENTER, italic=True)

add_bullets(s, Inches(0.55), Inches(5.5), Inches(12.25), Inches(1.5), [
    "Documents have identical names; must be downloaded individually",
    "Many PDFs aren't searchable — no copy/paste",
    "Nurses re-read the same content across docs to assemble one patient picture",
], size=13)
add_footer(s, 3)


# =============== SLIDE 4 — CDST Co-Pilot (with mock UI) ===============
s = prs.slides.add_slide(BLANK)
add_header(s, "03 · OUR SOLUTION", "CDST — A Co-Pilot for Nurses")
add_text(s, Inches(0.55), Inches(1.3), Inches(12.3), Inches(0.4),
         "Turn unstructured clinical documents into an AI summary the nurse validates instead of transcribes.",
         size=13, color=GRAY, italic=True)

# Mock app frame
app_x, app_y, app_w, app_h = Inches(0.55), Inches(1.85), Inches(12.25), Inches(4.6)
add_rect(s, app_x, app_y, app_w, app_h, fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
# title bar
add_rect(s, app_x + Inches(0.1), app_y + Inches(0.1), app_w - Inches(0.2), Inches(0.35),
         fill=RGBColor(0x2A,0x32,0x3F), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
# Traffic lights
for i, col in enumerate([RGBColor(0xFF,0x5F,0x57), RGBColor(0xFE,0xBC,0x2E), RGBColor(0x28,0xC8,0x40)]):
    add_rect(s, app_x + Inches(0.22) + Inches(0.22)*i, app_y + Inches(0.2), Inches(0.15), Inches(0.15),
             fill=col, shape=MSO_SHAPE.OVAL)
add_text(s, app_x + Inches(1.2), app_y + Inches(0.15), Inches(8), Inches(0.3),
         "CDST · Case #251027207666 · Smith, J. · Acute IP",
         size=10, bold=True, color=RGBColor(0xCF,0xD6,0xDF))

# Document panel (left)
docx = app_x + Inches(0.15); docy = app_y + Inches(0.55)
docw = Inches(5.6); doch = app_h - Inches(0.7)
add_rect(s, docx, docy, docw, doch, fill=WHITE)
add_text(s, docx + Inches(0.15), docy + Inches(0.1), docw - Inches(0.3), Inches(0.3),
         "📄  Document Viewer", size=11, bold=True, color=ACCENT)
# Doc list items
for i, doc in enumerate(["ED Notes · 12 pages", "H&P · 8 pages", "Imaging Report · 3 pages",
                          "Labs · 22 pages", "Cardiology Consult · 4 pages"]):
    yy = docy + Inches(0.5) + Inches(0.32) * i
    add_rect(s, docx + Inches(0.15), yy, docw - Inches(0.3), Inches(0.28),
             fill=LIGHT if i % 2 else SOFT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, docx + Inches(0.3), yy + Inches(0.04), docw - Inches(0.5), Inches(0.22),
             "■  " + doc, size=10, color=INK)
# Search bar
sby = docy + doch - Inches(0.5)
add_rect(s, docx + Inches(0.15), sby, docw - Inches(0.3), Inches(0.32),
         fill=WHITE, line=MIST, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, docx + Inches(0.3), sby + Inches(0.05), docw - Inches(0.5), Inches(0.25),
         "🔍  Search across all documents...", size=10, color=GRAY, italic=True)

# AI summary panel (right)
aix = docx + docw + Inches(0.15); aiy = docy
aiw = app_w - (aix - app_x) - Inches(0.15); aih = doch
add_rect(s, aix, aiy, aiw, aih, fill=WHITE)
add_text(s, aix + Inches(0.15), aiy + Inches(0.1), aiw - Inches(0.3), Inches(0.3),
         "🤖  AI Clinical Summary", size=11, bold=True, color=CVS_RED)
# Field cards
fields = [("HPI / CC", "65yo M, SOB x3 days, fever to 101.4°F..."),
          ("ED Summary", "Vitals: BP 142/88, HR 105, SpO2 88% on RA..."),
          ("Pertinent PMH", "HTN, T2DM, COPD on home O2"),
          ("Clinical Summary", "Pneumonia, RLL infiltrate on CXR. Sepsis criteria met...")]
for i, (label, body) in enumerate(fields):
    yy = aiy + Inches(0.5) + Inches(0.78) * i
    add_rect(s, aix + Inches(0.15), yy, aiw - Inches(0.3), Inches(0.7),
             fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, aix + Inches(0.15), yy, Inches(0.06), Inches(0.7), fill=CVS_RED)
    add_text(s, aix + Inches(0.3), yy + Inches(0.05), aiw - Inches(0.6), Inches(0.25),
             label, size=10, bold=True, color=CVS_RED)
    add_text(s, aix + Inches(0.3), yy + Inches(0.28), aiw - Inches(0.6), Inches(0.4),
             body, size=9, color=INK)
    # feedback icons
    add_text(s, aix + aiw - Inches(0.55), yy + Inches(0.05), Inches(0.4), Inches(0.25),
             "👍  👎", size=10, color=GRAY, align=PP_ALIGN.RIGHT)

# Bottom banner
add_text(s, Inches(0.55), Inches(6.65), Inches(12.25), Inches(0.4),
         "The nurse owns the final answer. AI accelerates; it does not decide.",
         size=13, bold=True, color=DARK, align=PP_ALIGN.CENTER, italic=True)
add_footer(s, 4)


# =============== SLIDE 5 — Three Pillars ===============
s = prs.slides.add_slide(BLANK)
add_header(s, "04 · SCOPE", "Three Pillars, One Co-Pilot Pattern")

pillars = [
    ("PRE-CERT / PA", "350K+ / yr", "By procedure group",
     ["Spine CDST · live", "Pre-cert AI layer · in build"], ACCENT),
    ("CONCURRENT REVIEW", "1.2M+ / yr", "Largest volume · Inpatient Acute",
     ["Medicare · scaled May '25", "Commercial · launched Aug '25"], CVS_RED),
    ("POST-ACUTE", "SNF · LTAC · Rehab", "Across all post-acute",
     ["SNF CDA+ · live", "SNF CDST + post-acute summarization · next"], PURPLE),
]
x0 = Inches(0.55); cw = Inches(4.13); gap = Inches(0.05); y0 = Inches(1.6); ch = Inches(4.8)
for i, (k, v, sub, items, color) in enumerate(pillars):
    x = x0 + (cw + gap) * i
    add_rect(s, x, y0, cw, ch, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, x, y0, cw, Inches(0.8), fill=color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # Mask bottom of header to remove bottom rounding overlap
    add_rect(s, x, y0 + Inches(0.6), cw, Inches(0.2), fill=color)
    add_text(s, x, y0 + Inches(0.18), cw, Inches(0.4),
             k, size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, x, y0 + Inches(0.45), cw, Inches(0.4),
             v, size=11, color=RGBColor(0xFF,0xFF,0xFF), align=PP_ALIGN.CENTER, italic=True)
    add_text(s, x + Inches(0.3), y0 + Inches(1.05), cw - Inches(0.6), Inches(0.5),
             sub, size=12, bold=True, color=INK)
    for j, it in enumerate(items):
        yy = y0 + Inches(1.65) + Inches(0.55) * j
        add_rect(s, x + Inches(0.3), yy, cw - Inches(0.6), Inches(0.45),
                 fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_rect(s, x + Inches(0.3), yy, Inches(0.06), Inches(0.45), fill=color)
        add_text(s, x + Inches(0.45), yy + Inches(0.1), cw - Inches(0.75), Inches(0.3),
                 it, size=11, color=INK)
    # Footer note
    add_text(s, x + Inches(0.3), y0 + ch - Inches(0.5), cw - Inches(0.6), Inches(0.4),
             "Lines of Business: Medicare · Commercial · Medicaid",
             size=9, color=GRAY, italic=True, align=PP_ALIGN.CENTER)

add_text(s, Inches(0.55), Inches(6.7), Inches(12.25), Inches(0.4),
         "Same co-pilot principle, expanding scope.",
         size=14, bold=True, color=DARK, align=PP_ALIGN.CENTER)
add_footer(s, 5)


# =============== SLIDE 6 — Results ===============
s = prs.slides.add_slide(BLANK)
add_header(s, "05 · RESULTS TO DATE", "Why It's Working")

# Big numbers
add_stat(s, Inches(0.55), Inches(1.45), Inches(2.95), Inches(1.4), "95%", "Adoption (Acute IP)", CVS_RED)
add_stat(s, Inches(3.65), Inches(1.45), Inches(2.95), Inches(1.4), "658K", "Reviews since Mar '25", ACCENT)
add_stat(s, Inches(6.75), Inches(1.45), Inches(2.95), Inches(1.4), "15–20%", "Time-per-review ↓", ACCENT)
add_stat(s, Inches(9.85), Inches(1.45), Inches(2.95), Inches(1.4), "$16M", "SAI booked · 2025", GREEN)

add_stat(s, Inches(0.55), Inches(3.0), Inches(2.95), Inches(1.4), "5.5 → 8–10", "Min saved (now → EOY)", CVS_RED)
add_stat(s, Inches(3.65), Inches(3.0), Inches(2.95), Inches(1.4), "70K+", "Reviews / month", ACCENT)
add_stat(s, Inches(6.75), Inches(3.0), Inches(2.95), Inches(1.4), "2.2K", "Non-MN cases flagged / yr", ACCENT)
add_stat(s, Inches(9.85), Inches(3.0), Inches(2.95), Inches(1.4), "0.1%", "In-app negative feedback", GREEN)

# Quotes
add_rect(s, Inches(0.55), Inches(4.7), Inches(6.05), Inches(1.95),
         fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, Inches(0.8), Inches(4.85), Inches(0.5), Inches(0.6),
         "“", size=48, bold=True, color=CYAN)
add_text(s, Inches(1.3), Inches(4.9), Inches(5.1), Inches(1.2),
         "30 attachments on 1 case. Thank goodness for this tool.",
         size=13, color=WHITE, italic=True)
add_text(s, Inches(1.3), Inches(6.25), Inches(5.1), Inches(0.3),
         "— Nurse · IP Medicare", size=10, color=CYAN)

add_rect(s, Inches(6.75), Inches(4.7), Inches(6.05), Inches(1.95),
         fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, Inches(7.0), Inches(4.85), Inches(0.5), Inches(0.6),
         "“", size=48, bold=True, color=CYAN)
add_text(s, Inches(7.5), Inches(4.9), Inches(5.1), Inches(1.3),
         "Did a great job pulling clinical from the doc, even though the fax was not clear. Terrible fax — pretty well.",
         size=12, color=WHITE, italic=True)
add_text(s, Inches(7.5), Inches(6.25), Inches(5.1), Inches(0.3),
         "— Nurse · IP Commercial", size=10, color=CYAN)
add_footer(s, 6)


# =============== SLIDE 7 — Human challenge ===============
s = prs.slides.add_slide(BLANK)
add_header(s, "06 · THE REAL CHALLENGE", "It Isn't Technical. It's Human.")
add_text(s, Inches(0.55), Inches(1.3), Inches(12.25), Inches(0.4),
         "Adoption + trust beat raw model accuracy. AI is recall-biased — nurses do the final edit. Edit patterns are the signal.",
         size=13, color=GRAY, italic=True)

# Two cards — Care Days vs ED
add_rect(s, Inches(0.55), Inches(1.95), Inches(6.05), Inches(3.0),
         fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_rect(s, Inches(0.55), Inches(1.95), Inches(6.05), Inches(0.55),
         fill=ACCENT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_rect(s, Inches(0.55), Inches(2.3), Inches(6.05), Inches(0.2), fill=ACCENT)
add_text(s, Inches(0.55), Inches(2.05), Inches(6.05), Inches(0.4),
         "CARE DAYS · NARRATIVE over-generation",
         size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
# Mini bar chart for deletions
bars = [("Subjective", 17.2), ("Diet", 11.5), ("Exams", 10.3), ("Respiratory", 9.2), ("EKG", 8.9)]
chart_y = Inches(2.7); bar_h = Inches(0.32); gap = Inches(0.05)
for i, (lbl, pct) in enumerate(bars):
    yy = chart_y + (bar_h + gap) * i
    add_text(s, Inches(0.8), yy, Inches(1.6), bar_h, lbl, size=10, color=INK,
             anchor=MSO_ANCHOR.MIDDLE)
    bw = Inches(3.5) * (pct / 20.0)
    add_rect(s, Inches(2.5), yy + Inches(0.06), bw, Inches(0.2), fill=ACCENT,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, Inches(2.5) + bw + Inches(0.05), yy, Inches(0.7), bar_h,
             f"{pct}%", size=10, bold=True, color=ACCENT, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Inches(0.8), Inches(4.55), Inches(5.6), Inches(0.35),
         "Tune for ASSERTIVENESS — strong-evidence sections only",
         size=10, bold=True, color=ACCENT, italic=True)

add_rect(s, Inches(6.75), Inches(1.95), Inches(6.05), Inches(3.0),
         fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_rect(s, Inches(6.75), Inches(1.95), Inches(6.05), Inches(0.55),
         fill=CVS_RED, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_rect(s, Inches(6.75), Inches(2.3), Inches(6.05), Inches(0.2), fill=CVS_RED)
add_text(s, Inches(6.75), Inches(2.05), Inches(6.05), Inches(0.4),
         "ED SUMMARY · OBJECTIVE BULK over-generation",
         size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
bars2 = [("Vitals", 17.1), ("Labs", 16.9), ("Imaging", 16.3), ("Therapy", 14.0), ("Exams", 13.7)]
for i, (lbl, pct) in enumerate(bars2):
    yy = chart_y + (bar_h + gap) * i
    add_text(s, Inches(7.0), yy, Inches(1.6), bar_h, lbl, size=10, color=INK,
             anchor=MSO_ANCHOR.MIDDLE)
    bw = Inches(3.5) * (pct / 20.0)
    add_rect(s, Inches(8.7), yy + Inches(0.06), bw, Inches(0.2), fill=CVS_RED,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, Inches(8.7) + bw + Inches(0.05), yy, Inches(0.7), bar_h,
             f"{pct}%", size=10, bold=True, color=CVS_RED, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Inches(7.0), Inches(4.55), Inches(5.6), Inches(0.35),
         "Tune for SELECTIVITY — only abnormal / critical values",
         size=10, bold=True, color=CVS_RED, italic=True)

# 4 pillar row
pillars4 = [
    ("1", "Page Classification & Dedup", "Clean docs before AI"),
    ("2", "Missing Vitals / Lab Resolution", "Structured gap-filling"),
    ("3", "User Behavior Deep Dive", "Chairside observation"),
    ("4", "UX Enablement", "Surface insights, improve review"),
]
y = Inches(5.2); cw = Inches(2.96); gap = Inches(0.12); x0 = Inches(0.55)
for i, (n, h, b) in enumerate(pillars4):
    x = x0 + (cw + gap) * i
    add_rect(s, x, y, cw, Inches(1.5), fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, x + Inches(0.2), y + Inches(0.1), Inches(0.6), Inches(0.6),
             n, size=32, bold=True, color=CYAN)
    add_text(s, x + Inches(0.8), y + Inches(0.18), cw - Inches(1.0), Inches(0.45),
             h, size=11, bold=True, color=WHITE)
    add_text(s, x + Inches(0.8), y + Inches(0.7), cw - Inches(1.0), Inches(0.6),
             b, size=10, color=RGBColor(0xCF,0xD6,0xDF))
add_footer(s, 7)


# =============== SLIDE 8 — GPT-4o → Gemini Vision (PIPELINE DIAGRAM) ===============
s = prs.slides.add_slide(BLANK)
add_header(s, "07 · TECHNICAL ARCHITECTURE", "GPT-4o → Gemini Vision · Page-by-Page")
add_text(s, Inches(0.55), Inches(1.3), Inches(12.25), Inches(0.4),
         "Switch LLM backbone from GPT-4o (OCR text in) to Gemini-1.5-pro-002 Vision (image in). Preserve the rest of the pipeline.",
         size=12, color=GRAY, italic=True)

# CURRENT pipeline (top)
y_cur = Inches(2.0)
add_text(s, Inches(0.55), y_cur, Inches(3), Inches(0.3),
         "CURRENT PIPELINE", size=11, bold=True, color=GRAY)
add_rect(s, Inches(0.55), y_cur + Inches(0.32), Inches(12.25), Inches(1.05),
         fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
# nodes
nodes_cur = [("📄  Document", "Faxed PDF"), ("🔤  OCR", "Layout-based text"),
             ("🧠  GPT-4o", "Azure OpenAI"), ("📋  Summary", "QuickBase fields")]
nx = Inches(0.85); nw = Inches(2.6); ny = y_cur + Inches(0.48); nh = Inches(0.75)
for i, (t, sub) in enumerate(nodes_cur):
    add_node(s, nx + (nw + Inches(0.45)) * i, ny, nw, nh, t, sub,
             fill=WHITE, border=GRAY, title_color=INK)
    if i < len(nodes_cur) - 1:
        ax1 = nx + (nw + Inches(0.45)) * i + nw
        add_arrow(s, ax1, ny + nh/2, ax1 + Inches(0.45), ny + nh/2, color=GRAY, weight=2.0)

# UPDATED pipeline (bottom)
y_new = Inches(3.7)
add_text(s, Inches(0.55), y_new, Inches(5), Inches(0.3),
         "UPDATED PIPELINE · Gemini Vision",
         size=11, bold=True, color=CVS_RED)
add_rect(s, Inches(0.55), y_new + Inches(0.32), Inches(12.25), Inches(1.05),
         fill=RGBColor(0xFD,0xF1,0xF1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_rect(s, Inches(0.55), y_new + Inches(0.32), Inches(0.1), Inches(1.05), fill=CVS_RED)
nodes_new = [("📄  Document", "Faxed PDF"),
             ("🖼  Page Image", "Per-page render"),
             ("👁  Gemini Vision", "1.5-pro-002 · multimodal"),
             ("📋  Summary", "QuickBase + citations")]
for i, (t, sub) in enumerate(nodes_new):
    add_node(s, nx + (nw + Inches(0.45)) * i, y_new + Inches(0.48), nw, nh, t, sub,
             fill=WHITE, border=CVS_RED, title_color=CVS_RED)
    if i < len(nodes_new) - 1:
        ax1 = nx + (nw + Inches(0.45)) * i + nw
        add_arrow(s, ax1, y_new + Inches(0.48) + nh/2,
                  ax1 + Inches(0.45), y_new + Inches(0.48) + nh/2, color=CVS_RED, weight=2.2)

# Bottom 3-column highlight cards
cy = Inches(5.45); cw3 = Inches(4.0); cg = Inches(0.13); cx0 = Inches(0.55)
cards = [
    ("WHY VISION", ACCENT,
     ["Word position + layout context preserved",
      "Faxes that break OCR still parse cleanly",
      "Native PDF page handling"]),
    ("ECONOMICS", GREEN,
     ["GPT-4o image: $0.001 / img",
      "Gemini Pro image: $0.0003 / img  (≈3× cheaper)",
      "Reduced OCR dependency"]),
    ("TARGETS", CVS_RED,
     ["95%+ summarization accuracy",
      "<1% pipeline failure rate",
      "100K+ requests / day at scale"]),
]
for i, (k, color, items) in enumerate(cards):
    x = cx0 + (cw3 + cg) * i
    add_rect(s, x, cy, cw3, Inches(1.6), fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, x, cy, cw3, Inches(0.35), fill=color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, x, cy + Inches(0.2), cw3, Inches(0.15), fill=color)
    add_text(s, x, cy + Inches(0.05), cw3, Inches(0.3),
             k, size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_bullets(s, x + Inches(0.2), cy + Inches(0.45), cw3 - Inches(0.4), Inches(1.1),
                items, size=10)

add_text(s, Inches(0.55), Inches(7.1), Inches(12.25), Inches(0.3),
         "Evaluation: 3-way bake-off (GPT-4o OCR · GPT-4o image · Gemini Vision image) on 20 MD-reviewed cases.  UI/UX: no change.",
         size=9, color=GRAY, align=PP_ALIGN.CENTER, italic=True)
add_footer(s, 8)


# =============== SLIDE 9 — Page-by-page internal architecture ===============
s = prs.slides.add_slide(BLANK)
add_header(s, "07b · TECHNICAL ARCHITECTURE", "Inside the Pipeline · Page-by-Page Extraction")
add_text(s, Inches(0.55), Inches(1.3), Inches(12.25), Inches(0.4),
         "Each PDF page → Vision call → aggregated into one mini-document per source doc → mapped to fields with page citations.",
         size=12, color=GRAY, italic=True)

# Document fan-out
src_x = Inches(0.55); src_y = Inches(2.1)
add_node(s, src_x, src_y, Inches(2.0), Inches(1.1),
         "📄  Source PDF", "ED Notes · 12 pages",
         fill=DARK, border=DARK, title_color=WHITE, sub_color=CYAN)

# Page images column
pg_x = Inches(3.1); pg_w = Inches(1.4); pg_h = Inches(0.55)
page_centers = []
for i in range(4):
    yy = src_y - Inches(0.05) + Inches(0.4) * i
    add_node(s, pg_x, yy, pg_w, pg_h,
             f"🖼  Page {i+1}", None,
             fill=WHITE, border=ACCENT, title_color=ACCENT, title_size=10)
    page_centers.append(yy + pg_h/2)
add_text(s, pg_x, src_y + Inches(1.65), pg_w, Inches(0.3),
         "+ N more...", size=9, color=GRAY, align=PP_ALIGN.CENTER, italic=True)
# arrow from PDF to pages
add_arrow(s, src_x + Inches(2.0), src_y + Inches(0.55),
          pg_x, src_y + Inches(0.55), color=DARK, weight=1.8)

# Vision LLM stack
vx = Inches(5.0); vy = src_y + Inches(0.1)
add_rect(s, vx, vy, Inches(2.4), Inches(1.6),
         fill=RGBColor(0xFD,0xF1,0xF1), line=CVS_RED, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, vx, vy + Inches(0.12), Inches(2.4), Inches(0.4),
         "👁  GEMINI VISION", size=12, bold=True, color=CVS_RED, align=PP_ALIGN.CENTER)
add_text(s, vx, vy + Inches(0.5), Inches(2.4), Inches(0.3),
         "Per-page call · 2 extraction Qs",
         size=10, color=INK, align=PP_ALIGN.CENTER, italic=True)
add_rect(s, vx + Inches(0.2), vy + Inches(0.9), Inches(2.0), Inches(0.28),
         fill=WHITE, line=CVS_RED, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, vx + Inches(0.2), vy + Inches(0.92), Inches(2.0), Inches(0.24),
         "Q1: clinical content",
         size=9, color=INK, align=PP_ALIGN.CENTER)
add_rect(s, vx + Inches(0.2), vy + Inches(1.22), Inches(2.0), Inches(0.28),
         fill=WHITE, line=CVS_RED, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, vx + Inches(0.2), vy + Inches(1.24), Inches(2.0), Inches(0.24),
         "Q2: dates + references",
         size=9, color=INK, align=PP_ALIGN.CENTER)

# arrows from pages to LLM
for c in page_centers:
    add_arrow(s, pg_x + pg_w, c, vx, vy + Inches(0.8), color=ACCENT, weight=1.2)

# Aggregation
ax = Inches(8.0); ay = vy
add_node(s, ax, ay, Inches(2.2), Inches(1.6),
         "🧩  AGGREGATION", "1 mini-document\nper source doc",
         fill=DARK, border=DARK, title_color=WHITE, sub_color=CYAN, title_size=12, sub_size=10)
add_arrow(s, vx + Inches(2.4), vy + Inches(0.8), ax, ay + Inches(0.8),
          color=CVS_RED, weight=2.0)

# Mapping + parser
mx = Inches(10.8); my = ay
add_node(s, mx, my, Inches(2.0), Inches(1.6),
         "🗺  MAPPING +\nPARSER",
         "Field assignment\nPage citations",
         fill=WHITE, border=ACCENT, title_color=ACCENT, sub_color=GRAY,
         title_size=11, sub_size=9)
add_arrow(s, ax + Inches(2.2), ay + Inches(0.8), mx, my + Inches(0.8),
          color=DARK, weight=2.0)

# Output card
out_y = Inches(4.5)
add_rect(s, Inches(0.55), out_y, Inches(12.25), Inches(1.6),
         fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_rect(s, Inches(0.55), out_y, Inches(0.12), Inches(1.6), fill=GREEN)
add_text(s, Inches(0.85), out_y + Inches(0.12), Inches(8), Inches(0.4),
         "📋  STRUCTURED OUTPUT  →  QuickBase template",
         size=12, bold=True, color=GREEN)
fields = [("HPI / CC", "1 field · cited page 1"),
          ("ED Summary", "12 fields · cited pages 1–4"),
          ("Clinical Summary", "8 fields · cited pages 5–10"),
          ("PMH", "5 fields · cited pages 2, 7")]
fw = Inches(2.9); fg = Inches(0.05); fx0 = Inches(0.85); fy = out_y + Inches(0.7)
for i, (l, b) in enumerate(fields):
    x = fx0 + (fw + fg) * i
    add_rect(s, x, fy, fw, Inches(0.75), fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, x, fy, Inches(0.06), Inches(0.75), fill=GREEN)
    add_text(s, x + Inches(0.18), fy + Inches(0.08), fw - Inches(0.3), Inches(0.3),
             l, size=11, bold=True, color=INK)
    add_text(s, x + Inches(0.18), fy + Inches(0.4), fw - Inches(0.3), Inches(0.3),
             b, size=9, color=GRAY, italic=True)

# Engineering scope strip
en_y = Inches(6.35)
add_text(s, Inches(0.55), en_y, Inches(12.25), Inches(0.3),
         "ENGINEERING CHANGE SCOPE", size=10, bold=True, color=DARK)
add_pill(s, Inches(0.55), en_y + Inches(0.35), Inches(2.5), Inches(0.45),
         "Extraction Qs update", ACCENT)
add_pill(s, Inches(3.15), en_y + Inches(0.35), Inches(2.5), Inches(0.45),
         "Aggregation + mapping", ACCENT)
add_pill(s, Inches(5.75), en_y + Inches(0.35), Inches(2.5), Inches(0.45),
         "Prompt + parser", ACCENT)
add_pill(s, Inches(8.35), en_y + Inches(0.35), Inches(2.5), Inches(0.45),
         "Non-LLM header/footer", GRAY)
add_pill(s, Inches(10.95), en_y + Inches(0.35), Inches(1.85), Inches(0.45),
         "UI: no change", GREEN)
add_footer(s, 9)


# =============== SLIDE 10 — Phased Roadmap ===============
s = prs.slides.add_slide(BLANK)
add_header(s, "08 · ROADMAP", "Phased Accuracy Improvement · May → Oct 2026")
add_text(s, Inches(0.55), Inches(1.3), Inches(12.25), Inches(0.4),
         "Each phase compounds the last. Baseline (Jan–Apr '26): 460 👍 / 439 👎 from 268 unique cases.",
         size=12, color=GRAY, italic=True)

# Timeline track
ty = Inches(2.0)
add_rect(s, Inches(0.55), ty + Inches(0.42), Inches(12.25), Inches(0.06), fill=MIST)

phases = [
    ("MAY '26", "Doc-Reference Recovery", "Backend fuzzy-match",
     "Recovers 1,714 / 1,762 historical failures (~97%). No prompt/model change.", "~31%", CVS_RED),
    ("JUN '26", "Prompt + Function", "Prompt engineering",
     "Date logic (planned vs actual), administered vs ordered meds, patient anchoring.", "~44%", ACCENT),
    ("AUG '26", "Page Classification", "ML pre-processing",
     "Filters duplicate / irrelevant pages before the LLM sees content.", "~30%", PURPLE),
    ("OCT '26", "Model Upgrade", "Foundation model refresh",
     "Gemini 3.1 — long-context, lower hallucination, compounding lift.", "95%+", GREEN),
]
x0 = Inches(0.55); cw = Inches(2.96); gap = Inches(0.13); y = ty
for i, (when, what, type_, body, lift, color) in enumerate(phases):
    x = x0 + (cw + gap) * i
    # circle node on timeline
    add_rect(s, x + cw/2 - Inches(0.18), ty + Inches(0.30), Inches(0.36), Inches(0.36),
             fill=color, shape=MSO_SHAPE.OVAL)
    add_text(s, x + cw/2 - Inches(0.18), ty + Inches(0.32), Inches(0.36), Inches(0.32),
             str(i+1), size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
    # card below
    cy = ty + Inches(0.95)
    add_rect(s, x, cy, cw, Inches(3.6), fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, x, cy, cw, Inches(0.45), fill=color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, x, cy + Inches(0.25), cw, Inches(0.2), fill=color)
    add_text(s, x, cy + Inches(0.08), cw, Inches(0.3),
             when, size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.2), cy + Inches(0.6), cw - Inches(0.4), Inches(0.5),
             what, size=13, bold=True, color=INK)
    add_text(s, x + Inches(0.2), cy + Inches(1.0), cw - Inches(0.4), Inches(0.3),
             type_, size=10, color=GRAY, italic=True)
    add_text(s, x + Inches(0.2), cy + Inches(1.4), cw - Inches(0.4), Inches(1.4),
             body, size=10, color=INK)
    # lift badge
    add_rect(s, x + Inches(0.2), cy + Inches(2.85), cw - Inches(0.4), Inches(0.6),
             fill=WHITE, line=color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, x, cy + Inches(2.92), cw, Inches(0.45),
             f"LIFT  {lift}", size=14, bold=True, color=color, align=PP_ALIGN.CENTER)

# Goal bar
gy = Inches(6.85)
add_text(s, Inches(0.55), gy, Inches(2), Inches(0.3),
         "CUMULATIVE", size=10, bold=True, color=GRAY)
add_rect(s, Inches(2.4), gy + Inches(0.04), Inches(8.5), Inches(0.22),
         fill=MIST, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
# progress segments
xs = Inches(2.4); ws = [Inches(2.0), Inches(3.0), Inches(2.2), Inches(1.3)]
cols = [CVS_RED, ACCENT, PURPLE, GREEN]
for w, col in zip(ws, cols):
    add_rect(s, xs, gy + Inches(0.04), w, Inches(0.22), fill=col,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    xs = xs + w
add_text(s, Inches(11.1), gy, Inches(2), Inches(0.3),
         "→ 95%+", size=12, bold=True, color=GREEN)
add_footer(s, 10)


# =============== SLIDE 11 — Agentic AI Architecture (BIG DIAGRAM) ===============
s = prs.slides.add_slide(BLANK)
add_header(s, "09 · TECHNICAL ARCHITECTURE", "Agentic AI · Next-Gen Modular Pipeline")
add_text(s, Inches(0.55), Inches(1.3), Inches(12.25), Inches(0.4),
         "Root orchestrator + modular agents.  Same pattern for IP, SNF, Spine — different components per pillar.",
         size=12, color=GRAY, italic=True)

# Inputs (left column)
ix = Inches(0.55); iy = Inches(2.0); iw = Inches(2.1); ih = Inches(0.8)
add_text(s, ix, iy - Inches(0.3), iw, Inches(0.3),
         "INPUTS", size=10, bold=True, color=GRAY)
add_node(s, ix, iy, iw, ih, "📄  Clinical Docs", "PDFs · faxes · EMR",
         fill=DARK, border=DARK, title_color=WHITE, sub_color=CYAN)
add_node(s, ix, iy + Inches(1.0), iw, ih, "📑  MCG Criteria", "Unstructured",
         fill=DARK, border=DARK, title_color=WHITE, sub_color=CYAN)

# Root orchestrator (center-top)
ox = Inches(5.2); oy = Inches(1.85); ow = Inches(3.0); oh = Inches(0.75)
add_text(s, ox, oy - Inches(0.3), ow, Inches(0.3),
         "ORCHESTRATION", size=10, bold=True, color=GRAY, align=PP_ALIGN.CENTER)
add_node(s, ox, oy, ow, oh, "🎯  ROOT ORCHESTRATOR", "Agent coordination",
         fill=CVS_RED, border=CVS_RED, title_color=WHITE, sub_color=RGBColor(0xFF,0xCC,0xCC))

# RETRIEVE band
band_y = Inches(3.0); band_h = Inches(1.55)
add_rect(s, Inches(3.05), band_y, Inches(9.78), band_h,
         fill=SOFT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, Inches(3.15), band_y + Inches(0.08), Inches(4), Inches(0.3),
         "1 · RETRIEVE CLINICAL INFORMATION",
         size=10, bold=True, color=ACCENT_D)
agents1 = [
    ("📚  Document\nOrganizer", "Gemini 2.5 Flash"),
    ("🔍  Clinical\nExtraction", "Gemini 2.5 Flash"),
    ("🧭  RAG Context\nRetrieval", "Gemini 2.0 Flash"),
]
nx = Inches(3.2); nw = Inches(3.15); nh = Inches(1.0); ny = band_y + Inches(0.42)
for i, (t, m) in enumerate(agents1):
    add_node(s, nx + (nw + Inches(0.1)) * i, ny, nw, nh, t, m,
             fill=WHITE, border=ACCENT, title_color=ACCENT, sub_color=GRAY,
             title_size=11, sub_size=9)

# DECISION band
band2_y = Inches(4.7); band2_h = Inches(1.55)
add_rect(s, Inches(3.05), band2_y, Inches(9.78), band2_h,
         fill=RGBColor(0xFD,0xF1,0xF1), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, Inches(3.15), band2_y + Inches(0.08), Inches(6), Inches(0.3),
         "2 · DECISION MAKING + RESULTS SUMMARIZATION",
         size=10, bold=True, color=CVS_RED_D)
agents2 = [
    ("📊  Criteria\nSummarizer", "Gemini 2.5 Flash"),
    ("⚖️  Decision\nAgent", "Gemini 2.0 Flash"),
    ("🧮  Decision\nAggregator", "Rule-based"),
    ("📝  Rationale\nAgent", "Gemini 2.0 Flash"),
]
nx2 = Inches(3.2); nw2 = Inches(2.32); nh2 = Inches(1.0); ny2 = band2_y + Inches(0.42)
for i, (t, m) in enumerate(agents2):
    add_node(s, nx2 + (nw2 + Inches(0.1)) * i, ny2, nw2, nh2, t, m,
             fill=WHITE, border=CVS_RED, title_color=CVS_RED, sub_color=GRAY,
             title_size=11, sub_size=9)

# CRITERIA STRUCTURE — left side under inputs
cdx = ix; cdy = Inches(4.2); cdw = iw; cdh = Inches(2.0)
add_text(s, cdx, cdy - Inches(0.3), cdw, Inches(0.3),
         "STRUCTURE CRITERIA", size=10, bold=True, color=GRAY)
add_node(s, cdx, cdy, cdw, Inches(1.0),
         "🔧  Criteria\nDecomposer", "Gemini 2.0 Flash",
         fill=WHITE, border=PURPLE, title_color=PURPLE, sub_color=GRAY)

# Arrows: inputs → orchestrator
add_arrow(s, ix + iw, iy + Inches(0.4), ox, oy + Inches(0.4), color=DARK)
add_arrow(s, cdx + cdw, cdy + Inches(0.5), nx, ny + Inches(0.5), color=PURPLE)
# orchestrator → band1
add_arrow(s, ox + ow/2, oy + oh, ox + ow/2, band_y, color=CVS_RED, weight=2.0)
# band1 → band2 (one arrow)
add_arrow(s, Inches(8), band_y + band_h, Inches(8), band2_y, color=DARK, weight=2.0)

# Output card on right
oxx = Inches(11.45); oyy = Inches(2.6); oww = Inches(1.35); ohh = Inches(3.5)
# Actually overflows — instead put output under bands
out_y = Inches(6.4)
add_rect(s, Inches(0.55), out_y, Inches(12.25), Inches(0.65),
         fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, Inches(0.85), out_y + Inches(0.12), Inches(12), Inches(0.4),
         "OUTPUT  →  ⚡ Atomic decision  ·  📌 Cited evidence  ·  📝 Rationale  ·  ✅ Final case outcome",
         size=12, bold=True, color=WHITE)
add_footer(s, 11)


# =============== SLIDE 12 — IP vs SNF comparison ===============
s = prs.slides.add_slide(BLANK)
add_header(s, "09b · TECHNICAL ARCHITECTURE", "Same Pattern, Different Shape · IP vs SNF")
add_text(s, Inches(0.55), Inches(1.3), Inches(12.25), Inches(0.4),
         "Both pipelines follow the same agentic pattern. Differences drive open decisions on framework, cost, and scope.",
         size=12, color=GRAY, italic=True)

# Two-column compare
left_x = Inches(0.55); right_x = Inches(6.95); col_w = Inches(6.0); col_y = Inches(1.95)
# Headers
add_rect(s, left_x, col_y, col_w, Inches(0.6), fill=ACCENT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, left_x, col_y + Inches(0.12), col_w, Inches(0.4),
         "INPATIENT ACUTE", size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_rect(s, right_x, col_y, col_w, Inches(0.6), fill=PURPLE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, right_x, col_y + Inches(0.12), col_w, Inches(0.4),
         "SNF (Post-Acute)", size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Rows
rows = [
    ("EXTRACTION", "Existing Gemini summary + RAG vector DB", "Built from scratch · document structurer"),
    ("CRITERIA", "Decomposer for 20+ diagnoses (MCG)", "Fixed 1 criteria set (simpler)"),
    ("FRONT-END", "Dynamic UI by diagnosis · rationale display", "Static UI from back-end data"),
    ("FRAMEWORK", "LangChain", "Google ADK"),
    ("MODEL", "Gemini 2.0 Flash", "Gemini 2.5 Flash + thinking budget"),
    ("LATENCY", "~3–5 min", "~10 min"),
]
ry = col_y + Inches(0.75); rh = Inches(0.65)
for i, (k, a, b) in enumerate(rows):
    yy = ry + (rh + Inches(0.06)) * i
    # row label centered
    add_rect(s, Inches(6.0), yy, Inches(0.9), rh, fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, Inches(6.0), yy + Inches(0.16), Inches(0.9), Inches(0.4),
             k, size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # left value
    add_rect(s, left_x, yy, Inches(5.35), rh, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, left_x + Inches(0.25), yy + Inches(0.16), Inches(5.0), Inches(0.4),
             a, size=11, color=INK)
    # right value
    add_rect(s, Inches(6.95), yy, Inches(5.85), rh, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, Inches(6.95) + Inches(0.25), yy + Inches(0.16), Inches(5.5), Inches(0.4),
             b, size=11, color=INK)

# Open decisions strip
od_y = Inches(6.45)
add_rect(s, Inches(0.55), od_y, Inches(12.25), Inches(0.75),
         fill=AMBER, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, Inches(0.85), od_y + Inches(0.08), Inches(11), Inches(0.3),
         "OPEN DECISIONS", size=10, bold=True, color=WHITE)
add_text(s, Inches(0.85), od_y + Inches(0.36), Inches(11), Inches(0.35),
         "Framework alignment (LangChain vs Google ADK)  ·  Cost + latency at scale  ·  Single generalized pipeline vs two",
         size=11, color=WHITE)
add_footer(s, 12)


# =============== SLIDE 13 — Phase Arc ===============
s = prs.slides.add_slide(BLANK)
add_header(s, "10 · THE ARC", "Each Phase Was Possible Because of the Last")

phases = [
    ("PHASE 1 · 2024", "Clinical Document\nAssistant",
     ["Unified per-case viewer", "Cross-document search", "→ DATA foundation"], ACCENT),
    ("PHASE 2 · TODAY", "CDST", 
     ["AI summarization", "Field auto-population", "In-app feedback loop", "→ TRUST foundation"], CVS_RED),
    ("PHASE 3 · Q3/Q4 '26", "Agentic AI",
     ["Criteria-aware agents", "Cited evidence", "Decision recommendations", "→ Nurse = DECISION-VALIDATOR"], GREEN),
]
x0 = Inches(0.55); cw = Inches(4.13); gap = Inches(0.05); y0 = Inches(1.7); ch = Inches(4.7)
for i, (k, t, items, color) in enumerate(phases):
    x = x0 + (cw + gap) * i
    add_rect(s, x, y0, cw, ch, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, x, y0, cw, Inches(0.5), fill=color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, x, y0 + Inches(0.3), cw, Inches(0.2), fill=color)
    add_text(s, x, y0 + Inches(0.12), cw, Inches(0.3),
             k, size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.3), y0 + Inches(0.7), cw - Inches(0.6), Inches(1.2),
             t, size=24, bold=True, color=INK)
    add_bullets(s, x + Inches(0.3), y0 + Inches(2.2), cw - Inches(0.6), Inches(2.3),
                items, size=12)

# Connecting arrows between phases
for i in range(2):
    x1 = x0 + (cw + gap) * i + cw
    add_arrow(s, x1 - Inches(0.05), y0 + ch/2, x1 + Inches(0.1), y0 + ch/2,
              color=DARK, weight=2.5)

add_rect(s, Inches(0.55), Inches(6.65), Inches(12.25), Inches(0.55),
         fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_text(s, Inches(0.55), Inches(6.75), Inches(12.25), Inches(0.4),
         "Same co-pilot principle, expanding scope. Human clinical judgment always retained.",
         size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_footer(s, 13)


# =============== SLIDE 14 — How I Use AI Personally ===============
s = prs.slides.add_slide(BLANK)
add_header(s, "11 · HOW I USE AI PERSONALLY",
           "It's Not Just the Product · It's How I Work Every Day")
add_text(s, Inches(0.55), Inches(1.3), Inches(12.25), Inches(0.4),
         "Same principle as CDST — AI accelerates, I decide. Here's my personal stack.",
         size=13, color=GRAY, italic=True)

uses = [
    ("📊", "ANALYSIS", "Analysis",
     "BigQuery pulls, feedback bucketing, edit-pattern detection.\nAI does the heavy lifting; I bring the hypothesis & spot-audit.",
     CVS_RED),
    ("🎯", "POWERPOINT", "PowerPoint",
     "From CSV to stakeholder-ready deck in one prompt.\nNo designer, no template wrestling — just the story.",
     ACCENT),
    ("🛠", "APPS", "Create Apps",
     "Streamlit dashboards, internal tools, prototypes.\nShip a feedback dashboard in a day instead of a sprint.",
     PURPLE),
    ("📦", "ARTIFACTS", "Build Artifacts",
     "Excel trackers, Word reports, FRDs, mock UIs, diagrams.\nThe plumbing every PM owes their team — without the wait.",
     AMBER),
    ("🤖", "HERMES AGENT", "Hermes Agent",
     "Daily standup in my inbox. Pipeline status, feedback delta,\nwhat shipped, what broke — auto-summarized while I sleep.",
     GREEN),
]
x0 = Inches(0.55); cw = Inches(2.42); gap = Inches(0.05); y0 = Inches(1.9); ch = Inches(4.0)
for i, (icon, kicker, title, body, color) in enumerate(uses):
    x = x0 + (cw + gap) * i
    # card
    add_rect(s, x, y0, cw, ch, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # top accent
    add_rect(s, x, y0, cw, Inches(0.12), fill=color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_rect(s, x, y0 + Inches(0.06), cw, Inches(0.08), fill=color)
    # icon
    add_text(s, x, y0 + Inches(0.3), cw, Inches(0.7),
             icon, size=36, bold=True, color=DARK, align=PP_ALIGN.CENTER)
    # kicker
    add_text(s, x, y0 + Inches(1.05), cw, Inches(0.3),
             f"USE 0{i+1}", size=9, bold=True, color=color, align=PP_ALIGN.CENTER)
    # title
    add_text(s, x, y0 + Inches(1.35), cw, Inches(0.5),
             title, size=18, bold=True, color=INK, align=PP_ALIGN.CENTER)
    # divider line
    add_rect(s, x + Inches(0.8), y0 + Inches(1.95), cw - Inches(1.6), Inches(0.02),
             fill=color)
    # body
    add_text(s, x + Inches(0.2), y0 + Inches(2.1), cw - Inches(0.4), Inches(1.8),
             body, size=10, color=INK, align=PP_ALIGN.CENTER)

# Punch line
py = Inches(6.1)
add_rect(s, Inches(0.55), py, Inches(12.25), Inches(0.85),
         fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
add_rect(s, Inches(0.55), py, Inches(0.12), Inches(0.85), fill=CYAN)
add_text(s, Inches(0.85), py + Inches(0.13), Inches(11.5), Inches(0.35),
         "Zero tickets filed. Zero meetings about meetings.",
         size=14, bold=True, color=CYAN)
add_text(s, Inches(0.85), py + Inches(0.45), Inches(11.5), Inches(0.4),
         "I run my own data, draft my own artifacts, and ship my own analysis — so DS & Eng can build, not export.",
         size=12, color=WHITE, italic=True)
add_footer(s, 14, total=15)


# =============== SLIDE 15 — Takeaways ===============
s = prs.slides.add_slide(BLANK)
add_header(s, "12 · TAKEAWAYS", "What to Remember")

takeaways = [
    ("01", "The healthcare problem is hard.",
     "2.26M cases · $250M admin · faxed PDFs · CMS pressure.", CVS_RED),
    ("02", "AI works when it serves the workflow.",
     "CDST is a co-pilot · 95% adoption · 15–20% time saved · $16M SAI.", ACCENT),
    ("03", "The technical path is clear.",
     "Page-by-page Gemini Vision · ML pre-processing · prompts · model upgrades.", PURPLE),
    ("04", "The hard part is human.",
     "UX · trust · edit patterns · chairside observation · per-form strategy.", AMBER),
    ("05", "Agentic is next.",
     "Modular components · decision agents · only possible because we earned trust first.", GREEN),
]
y = Inches(1.55)
for i, (n, head, body, color) in enumerate(takeaways):
    yy = y + Inches(1.05) * i
    # number badge
    add_rect(s, Inches(0.55), yy, Inches(0.85), Inches(0.9),
             fill=color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, Inches(0.55), yy + Inches(0.18), Inches(0.85), Inches(0.6),
             n, size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # card
    add_rect(s, Inches(1.55), yy, Inches(11.25), Inches(0.9),
             fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_text(s, Inches(1.8), yy + Inches(0.13), Inches(11), Inches(0.4),
             head, size=15, bold=True, color=INK)
    add_text(s, Inches(1.8), yy + Inches(0.5), Inches(11), Inches(0.4),
             body, size=12, color=GRAY, italic=True)
add_footer(s, 15, total=15)


# Save
out = "/Users/a965598/Downloads/UM_AI_Presentation/UM_AI_Presentation.pptx"
prs.save(out)
print(f"Saved: {out}")
