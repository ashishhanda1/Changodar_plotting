"""Build the marketing deck output/Changodar_Plots_Marketing.pptx (run build_images.py first)."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

A = "assets/"
OUT = "output/Changodar_Plots_Marketing.pptx"
DARK, MID, LIGHT, GOLD, CREAM, WHITE, GREY = (RGBColor.from_string(c) for c in
    ("0F3D2E", "2E7D5B", "E8F3EC", "D4A017", "FBF7EE", "FFFFFF", "5F6B66"))
FONT = "Arial"

prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
BL = prs.slide_layouts[6]

def rect(s, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE):
    r = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    r.fill.solid(); r.fill.fore_color.rgb = fill; r.line.fill.background(); r.shadow.inherit = False
    return r

def text(s, x, y, w, h, t, size=18, color=DARK, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = tb.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05)
    lines = t if isinstance(t, list) else [t]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if isinstance(ln, tuple): ln, sz, col, b = ln
        else: sz, col, b = size, color, bold
        p.alignment = align
        r = p.add_run(); r.text = ln; f = r.font
        f.name = FONT; f.size = Pt(sz); f.color.rgb = col; f.bold = b
        p.space_after = Pt(6)
    return tb

def bg(s, c=CREAM): rect(s, 0, 0, 13.333, 7.5, c)

def heading(s, t, sub=None):
    rect(s, 0.6, 0.55, 0.12, 0.75, GOLD)
    text(s, 0.85, 0.42, 11.5, 0.7, t, 34, DARK, True)
    if sub: text(s, 0.87, 1.05, 11.5, 0.5, sub, 16, GREY)

def footer(s, n):
    text(s, 0.6, 7.0, 8, 0.35, "Changodar Plots  |  150 sq yd residential plots", 10, GREY)
    text(s, 12.0, 7.0, 0.8, 0.35, str(n), 10, GREY, align=PP_ALIGN.RIGHT)

def card(s, x, y, w, h, big, small, fill=WHITE, bigc=DARK, smallc=GREY, bs=30):
    rect(s, x, y, w, h, fill, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.08
    text(s, x + 0.2, y + 0.2, w - 0.4, h * 0.5, big, bs, bigc, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(s, x + 0.2, y + h * 0.55, w - 0.4, h * 0.4, small, 14, smallc, False, PP_ALIGN.CENTER)

# 1 Title
s = prs.slides.add_slide(BL)
s.shapes.add_picture(A + "site_layout.png", 0, 0, prs.slide_width, prs.slide_height)
rect(s, 0, 0, 13.333, 7.5, DARK).fill.fore_color.brightness = 0
ov = s.shapes[-1]; ov.fill.solid(); ov.fill.fore_color.rgb = DARK
from pptx.oxml.ns import qn
sf = ov.fill._xPr.find(qn("a:solidFill")); clr = sf[0]
alpha = clr.makeelement(qn("a:alpha"), {"val": "72000"}); clr.append(alpha)
text(s, 0.9, 1.9, 11.5, 0.6, "CHANGODAR, AHMEDABAD", 20, GOLD, True)
text(s, 0.9, 2.5, 11.5, 1.6, "Own Your Plot at Changodar", 56, WHITE, True)
text(s, 0.9, 3.75, 11.5, 0.8, "166 residential plots of 150 sq yd each  |  Easy DP and 4-month EMI options", 22, WHITE)
rect(s, 0.9, 4.9, 3.6, 0.85, GOLD, MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 0.9, 4.9, 3.6, 0.85, "From Rs 19,500 / sq yd", 22, DARK, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)

# 2 At a glance
s = prs.slides.add_slide(BL); bg(s)
heading(s, "The Project at a Glance", "A planned plotting project with clear titles and simple payment options")
for i, (b, sm) in enumerate([("25,000", "sq yd total land"), ("166", "plots available"),
                             ("150 sq yd", "each plot (30 ft x 45 ft)"), ("Rs 30 lakh", "regular price per plot")]):
    card(s, 0.7 + i * 3.05, 2.0, 2.8, 2.0, b, sm, WHITE, DARK if i % 2 == 0 else MID)
text(s, 0.7, 4.5, 12, 2.2, [
    ("Why invest in a plot?", 22, DARK, True),
    ("•  Land you own outright, ready to build your home on your own timeline", 18, DARK, False),
    ("•  Ideal size for a family home: 1,350 sq ft of land", 18, DARK, False),
    ("•  Pay in full and save, or spread the payment over 4 months", 18, DARK, False)])
footer(s, 2)

# 3 Layout and plot size
s = prs.slides.add_slide(BL); bg(s)
heading(s, "Layout and Plot Size", "Illustrative layout. The final layout follows the approved plan.")
s.shapes.add_picture(A + "site_layout.png", Inches(0.7), Inches(1.75), width=Inches(7.6))
s.shapes.add_picture(A + "single_plot.png", Inches(8.7), Inches(1.6), height=Inches(4.3))
text(s, 8.6, 5.95, 4.3, 1.0, [("Every plot: 150 sq yd", 20, DARK, True),
                              ("30 ft wide x 45 ft deep, road facing", 15, GREY, False)], align=PP_ALIGN.CENTER)
text(s, 0.7, 6.15, 7.6, 0.6, "Internal roads  •  Green park  •  Tree-lined boundary", 16, MID, True, PP_ALIGN.CENTER)
footer(s, 3)

# 4 Pricing
s = prs.slides.add_slide(BL); bg(s)
heading(s, "Simple, Transparent Pricing", "Choose the option that suits you")
rect(s, 0.8, 1.9, 5.6, 4.4, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.06
text(s, 1.0, 2.1, 5.2, 0.5, "REGULAR PRICE", 16, GREY, True, PP_ALIGN.CENTER)
text(s, 1.0, 2.6, 5.2, 1.0, "Rs 20,000", 48, DARK, True, PP_ALIGN.CENTER)
text(s, 1.0, 3.55, 5.2, 0.5, "per sq yd", 18, GREY, False, PP_ALIGN.CENTER)
text(s, 1.0, 4.3, 5.2, 1.6, [("Rs 30,00,000 per plot", 24, DARK, True),
                             ("Pay over 4 months with EMI", 16, GREY, False)], align=PP_ALIGN.CENTER)
rect(s, 6.9, 1.9, 5.6, 4.4, DARK, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.06
rect(s, 10.3, 1.55, 2.4, 0.5, GOLD, MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 10.3, 1.55, 2.4, 0.5, "BEST VALUE", 15, DARK, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
text(s, 7.1, 2.15, 5.2, 0.5, "DP PRICE: PAY IN 2 WEEKS", 16, GOLD, True, PP_ALIGN.CENTER)
text(s, 7.1, 2.6, 5.2, 1.0, "Rs 19,500", 48, WHITE, True, PP_ALIGN.CENTER)
text(s, 7.1, 3.55, 5.2, 0.5, "per sq yd", 18, LIGHT, False, PP_ALIGN.CENTER)
text(s, 7.1, 4.3, 5.2, 1.6, [("Rs 29,25,000 per plot", 24, WHITE, True),
                             ("You save Rs 75,000", 20, GOLD, True)], align=PP_ALIGN.CENTER)
text(s, 0.8, 6.45, 12, 0.4, "Prices exclude stamp duty, registration and other government charges.", 12, GREY)
footer(s, 4)

# 5 EMI
s = prs.slides.add_slide(BL); bg(s)
heading(s, "Easy 4-Month EMI", "Book now and pay in 4 equal monthly instalments")
pic = s.shapes.add_picture(A + "emi_timeline.png", 0, Inches(1.65), height=Inches(3.0))
pic.left = int((prs.slide_width - pic.width) / 2)
for i, (b, sm) in enumerate([("Rs 30,00,000", "plot price"), ("15% p.a.", "interest, for 4 months"),
                             ("Rs 7,87,500", "monthly EMI x 4"), ("Rs 31,50,000", "total payable")]):
    card(s, 0.7 + i * 3.05, 4.85, 2.8, 1.75, b, sm, GOLD if i == 2 else WHITE, DARK, DARK if i == 2 else GREY, 26)
footer(s, 5)

# 6 Compare
s = prs.slides.add_slide(BL); bg(s)
heading(s, "DP or EMI: Compare at a Glance", "Per 150 sq yd plot")
rows = [("", "DP (pay in 2 weeks)", "4-month EMI"), ("Price per sq yd", "Rs 19,500", "Rs 20,000"),
        ("Plot price", "Rs 29,25,000", "Rs 30,00,000"), ("Interest", "None", "Rs 1,50,000 (15% p.a. for 4 months)"),
        ("Payment", "Full amount within 14 days", "Rs 7,87,500 every month x 4"),
        ("Total you pay", "Rs 29,25,000", "Rs 31,50,000"), ("Effective price per sq yd", "Rs 19,500", "Rs 21,000")]
tbl = s.shapes.add_table(len(rows), 3, Inches(0.8), Inches(1.8), Inches(11.7), Inches(4.4)).table
tbl.columns[0].width = Inches(3.5); tbl.columns[1].width = Inches(3.8); tbl.columns[2].width = Inches(4.4)
for r, row in enumerate(rows):
    for c, v in enumerate(row):
        cell = tbl.cell(r, c); cell.text = v
        cell.fill.solid()
        cell.fill.fore_color.rgb = DARK if r == 0 else (LIGHT if r == 5 else WHITE)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
        f = p.runs[0].font if p.runs else p.font
        f.name = FONT; f.size = Pt(17); f.bold = r in (0, 5) or c == 0
        f.color.rgb = WHITE if r == 0 else DARK
text(s, 0.8, 6.35, 11.7, 0.5, "Pay in full and save Rs 2,25,000 compared with the EMI option.", 18, MID, True)
footer(s, 6)

# 7 Sales team ladder
s = prs.slides.add_slide(BL); bg(s, DARK)
text(s, 0.8, 0.45, 11.5, 0.8, "Join Our Sales Team: Ladder Income", 34, WHITE, True)
text(s, 0.82, 1.1, 11.5, 0.5, "30 sales partners. The more plots you sell, the higher your rate.", 17, LIGHT)
rect(s, 0.8, 1.8, 7.4, 4.9, CREAM, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.04
s.shapes.add_picture(A + "ladder.png", Inches(1.0), Inches(2.0), width=Inches(7.0))
text(s, 8.7, 2.0, 4.0, 4.8, [
    ("How it works", 22, GOLD, True),
    ("•  Commission on the plot sale value", 17, WHITE, False),
    ("•  Your rate applies to ALL plots you sell, not just the extra ones", 17, WHITE, False),
    ("•  Sell 6 or more plots to reach the top 8% rate", 17, WHITE, False),
    ("•  Earn Rs 1.5 lakh on your very first plot", 17, WHITE, False)])
text(s, 0.6, 7.0, 8, 0.35, "Changodar Plots  |  150 sq yd residential plots", 10, LIGHT)

# 8 Earnings
s = prs.slides.add_slide(BL); bg(s)
heading(s, "What You Can Earn", "Commission at the regular price of Rs 30 lakh per plot")
s.shapes.add_picture(A + "earnings.png", Inches(0.6), Inches(1.65), height=Inches(5.0))
for i, (b, sm) in enumerate([("Rs 1.5 lakh", "sell 1 plot (5%)"), ("Rs 11.25 lakh", "sell 5 plots (7.5%)"),
                             ("Rs 24 lakh", "sell 10 plots (8%)")]):
    card(s, 9.3, 1.75 + i * 1.65, 3.4, 1.45, b, sm, GOLD if i == 2 else WHITE, DARK, DARK if i == 2 else GREY, 26)
footer(s, 8)

# 9 How to book
s = prs.slides.add_slide(BL); bg(s)
heading(s, "How to Book Your Plot", "Four simple steps")
steps = [("1", "Visit the site", "See the layout and choose your plot"),
         ("2", "Pick a payment plan", "DP at Rs 19,500 / sq yd or 4-month EMI"),
         ("3", "Pay", "Full amount within 14 days, or EMI 1 in month 1"),
         ("4", "Get registered", "Sale deed registered in your name")]
for i, (n, t, d) in enumerate(steps):
    x = 0.8 + i * 3.05
    rect(s, x, 2.0, 2.8, 3.8, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE).adjustments[0] = 0.06
    c = rect(s, x + 0.95, 2.3, 0.9, 0.9, GOLD if i == 3 else DARK, MSO_SHAPE.OVAL)
    text(s, x + 0.95, 2.3, 0.9, 0.9, n, 28, WHITE, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(s, x + 0.15, 3.45, 2.5, 0.6, t, 20, DARK, True, PP_ALIGN.CENTER)
    text(s, x + 0.15, 4.15, 2.5, 1.4, d, 15, GREY, False, PP_ALIGN.CENTER)
footer(s, 9)

# 10 CTA
s = prs.slides.add_slide(BL)
s.shapes.add_picture(A + "site_layout.png", 0, 0, prs.slide_width, prs.slide_height)
ov = rect(s, 0, 0, 13.333, 7.5, DARK)
clr = ov.fill._xPr.find(qn("a:solidFill"))[0]; clr.append(clr.makeelement(qn("a:alpha"), {"val": "78000"}))
text(s, 0.9, 2.0, 11.5, 1.2, "Book Your Plot Today", 54, WHITE, True, PP_ALIGN.CENTER)
text(s, 0.9, 3.2, 11.5, 0.8, "Only 166 plots  |  DP from Rs 29.25 lakh  |  EMI Rs 7,87,500 x 4", 22, GOLD, True, PP_ALIGN.CENTER)
rect(s, 3.9, 4.4, 5.5, 1.3, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 3.9, 4.4, 5.5, 1.3, [("Call: [Your phone number]", 22, DARK, True), ("Site office: Changodar, Ahmedabad", 16, GREY, False)],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

prs.save(OUT); print("deck saved")
