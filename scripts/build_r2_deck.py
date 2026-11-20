"""R2 sales-partner deck: output/R2_Sales_Partner_Plan.pptx (run build_images.py and build_r2_images.py first)."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

A = "assets/"; OUT = "output/R2_Sales_Partner_Plan.pptx"
DARK, MID, LIGHT, GOLD, CREAM, WHITE, GREY = (RGBColor.from_string(c) for c in
    ("0F3D2E", "2E7D5B", "E8F3EC", "D4A017", "FBF7EE", "FFFFFF", "5F6B66"))
FONT = "Arial"
prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5); BL = prs.slide_layouts[6]

def rect(s, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE):
    r = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    r.fill.solid(); r.fill.fore_color.rgb = fill; r.line.fill.background(); r.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE: r.adjustments[0] = 0.07
    return r
def text(s, x, y, w, h, t, size=18, color=DARK, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = tb.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor; tf.margin_left = tf.margin_right = Inches(0.05)
    for i, ln in enumerate(t if isinstance(t, list) else [t]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        ln, sz, col, b = ln if isinstance(ln, tuple) else (ln, size, color, bold)
        p.alignment = align; p.space_after = Pt(6)
        r = p.add_run(); r.text = ln; f = r.font; f.name = FONT; f.size = Pt(sz); f.color.rgb = col; f.bold = b
    return tb
def bg(s, c=CREAM): rect(s, 0, 0, 13.333, 7.5, c)
def overlay(s, alpha="74000"):
    s.shapes.add_picture(A + "site_layout.png", 0, 0, prs.slide_width, prs.slide_height)
    ov = rect(s, 0, 0, 13.333, 7.5, DARK); clr = ov.fill._xPr.find(qn("a:solidFill"))[0]
    clr.append(clr.makeelement(qn("a:alpha"), {"val": alpha}))
def heading(s, t, sub=None, dark=False):
    rect(s, 0.6, 0.55, 0.12, 0.75, GOLD)
    text(s, 0.85, 0.42, 11.8, 0.7, t, 34, WHITE if dark else DARK, True)
    if sub: text(s, 0.87, 1.05, 11.8, 0.5, sub, 16, LIGHT if dark else GREY)
def footer(s, n, dark=False):
    c = LIGHT if dark else GREY
    text(s, 0.6, 7.0, 8, 0.35, "R2 Sales Partner Plan  |  Changodar Plots", 10, c)
    text(s, 12.0, 7.0, 0.8, 0.35, str(n), 10, c, align=PP_ALIGN.RIGHT)
def card(s, x, y, w, h, big, small, fill=WHITE, bigc=DARK, smallc=GREY, bs=30):
    rect(s, x, y, w, h, fill, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x + 0.15, y + 0.15, w - 0.3, h * 0.5, big, bs, bigc, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(s, x + 0.15, y + h * 0.56, w - 0.3, h * 0.4, small, 14, smallc, False, PP_ALIGN.CENTER)

# 1 Title
s = prs.slides.add_slide(BL); overlay(s)
text(s, 0.9, 1.7, 11.5, 0.6, "FOR SALES PARTNERS  |  CHANGODAR PLOTS", 18, GOLD, True)
text(s, 0.9, 2.3, 11.5, 1.2, "R2 Sales Partner Plan", 58, WHITE, True)
text(s, 0.9, 3.55, 11.5, 0.8, "Earn up to Rs 2,000 for every sq yd you sell", 28, WHITE)
rect(s, 0.9, 4.75, 4.6, 0.9, GOLD, MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 0.9, 4.75, 4.6, 0.9, "Rs 3 lakh per plot at the top slab", 20, DARK, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)

# 2 Opportunity
s = prs.slides.add_slide(BL); bg(s)
heading(s, "The Opportunity", "A large, ready inventory and a small, focused sales team")
for i, (b, sm) in enumerate([("166", "plots to sell"), ("150 sq yd", "every plot (30 ft x 45 ft)"),
                             ("Rs 30 lakh", "value of each plot"), ("30", "sales partners only")]):
    card(s, 0.7 + i * 3.05, 1.95, 2.8, 2.0, b, sm, GOLD if i == 3 else WHITE, DARK, DARK if i == 3 else GREY, 32)
text(s, 0.7, 4.45, 12, 2.3, [
    ("Why this works for you", 22, DARK, True),
    ("•  About 5 to 6 plots per partner on average if all 166 plots are sold", 18, DARK, False),
    ("•  A fixed rupee amount per sq yd, so you always know exactly what you earn", 18, DARK, False),
    ("•  Buyers get two easy options (DP or 4-month EMI), which helps you close faster", 18, DARK, False)])
footer(s, 2)

# 3 Slabs
s = prs.slides.add_slide(BL); bg(s, DARK)
heading(s, "Your R2 Income Slabs", "Fixed rupees for every sq yd you sell", dark=True)
rect(s, 0.7, 1.8, 7.4, 4.95, CREAM, MSO_SHAPE.ROUNDED_RECTANGLE)
s.shapes.add_picture(A + "r2_ladder.png", Inches(0.95), Inches(2.0), width=Inches(6.9))
text(s, 8.6, 1.9, 4.2, 4.9, [
    ("How it works", 22, GOLD, True),
    ("•  1 plot: Rs 1,700 per sq yd", 18, WHITE, False),
    ("•  2 to 4 plots: Rs 1,800 per sq yd", 18, WHITE, False),
    ("•  5 plots and above: Rs 2,000 per sq yd", 18, WHITE, False),
    ("•  Your slab rate applies to ALL the plots you sell, not just the extra ones", 18, WHITE, False),
    ("•  Same payout whether the buyer picks DP or EMI", 18, WHITE, False)])
footer(s, 3, dark=True)

# 4 Earnings chart
s = prs.slides.add_slide(BL); bg(s)
heading(s, "What You Can Earn", "Each plot is 150 sq yd")
s.shapes.add_picture(A + "r2_earnings.png", Inches(0.6), Inches(1.6), height=Inches(5.1))
for i, (b, sm) in enumerate([("Rs 2.55 lakh", "sell 1 plot"), ("Rs 15 lakh", "sell 5 plots"), ("Rs 30 lakh", "sell 10 plots")]):
    card(s, 9.3, 1.75 + i * 1.65, 3.4, 1.45, b, sm, GOLD if i == 2 else WHITE, DARK, DARK if i == 2 else GREY, 28)
footer(s, 4)

# 5 Table
s = prs.slides.add_slide(BL); bg(s)
heading(s, "Earnings Table", "Your total earning by the number of plots you sell")
rate = lambda k: 1700 if k == 1 else 1800 if k <= 4 else 2000
rows = [("Plots sold", "Rate per sq yd", "Earning per plot", "Total earning")]
for k in (1, 2, 3, 4, 5, 6, 8, 10):
    r = rate(k); rows.append((str(k), f"Rs {r:,}", f"Rs {r*150/1e5:g} lakh", f"Rs {k*r*150/1e5:g} lakh"))
tb = s.shapes.add_table(len(rows), 4, Inches(1.4), Inches(1.75), Inches(10.5), Inches(4.8)).table
for r, row in enumerate(rows):
    for c, v in enumerate(row):
        cell = tb.cell(r, c); cell.text = v; cell.fill.solid(); cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        top = r >= 5
        cell.fill.fore_color.rgb = DARK if r == 0 else (RGBColor.from_string("FBEFC8") if top else WHITE)
        p = cell.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        f = p.runs[0].font; f.name = FONT; f.size = Pt(18); f.bold = r == 0 or c == 3
        f.color.rgb = WHITE if r == 0 else DARK
text(s, 1.4, 6.6, 10.5, 0.4, "Highlighted rows are the top slab: Rs 2,000 per sq yd on every plot.", 14, GREY)
footer(s, 5)

# 6 Climb
s = prs.slides.add_slide(BL); bg(s)
heading(s, "Push for Your 5th Plot", "Hitting the top slab lifts the rate on every plot you have sold")
card(s, 0.9, 2.0, 3.6, 2.6, "Rs 10.8 lakh", "4 plots at Rs 1,800 / sq yd", WHITE, DARK, GREY, 34)
arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(4.8), Inches(2.85), Inches(1.4), Inches(0.9))
arrow.fill.solid(); arrow.fill.fore_color.rgb = GOLD; arrow.line.fill.background()
card(s, 6.5, 2.0, 3.6, 2.6, "Rs 15 lakh", "5 plots at Rs 2,000 / sq yd", DARK, WHITE, LIGHT, 34)
card(s, 10.4, 2.0, 2.3, 2.6, "+Rs 4.2 lakh", "from one more sale", GOLD, DARK, DARK, 24)
text(s, 0.9, 5.0, 11.8, 1.6, [
    ("One extra plot adds Rs 4.2 lakh: Rs 3 lakh for that plot, plus Rs 1.2 lakh more on your first 4.", 20, DARK, True),
    ("In the same way, a 2nd plot takes you from Rs 2.55 lakh to Rs 5.4 lakh.", 18, GREY, False)])
footer(s, 6)

# 7 Pitch kit
s = prs.slides.add_slide(BL); bg(s)
heading(s, "Your Pitch Kit: What You Are Selling", "The key facts to share with buyers")
s.shapes.add_picture(A + "single_plot.png", Inches(0.7), Inches(1.65), height=Inches(4.9))
facts = [("Plot", "150 sq yd (30 ft x 45 ft), road facing"),
         ("Regular price", "Rs 20,000 / sq yd = Rs 30 lakh"),
         ("DP price", "Rs 19,500 / sq yd = Rs 29.25 lakh, full payment in 2 weeks (buyer saves Rs 75,000)"),
         ("4-month EMI", "Rs 7,87,500 x 4 = Rs 31.5 lakh (15% p.a. flat interest)")]
for i, (k, v) in enumerate(facts):
    y = 1.75 + i * 1.2
    rect(s, 5.6, y, 7.1, 1.05, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, 5.6, y, 0.12, 1.05, GOLD)
    text(s, 5.9, y + 0.08, 6.7, 0.4, k, 15, MID, True)
    text(s, 5.9, y + 0.42, 6.7, 0.6, v, 16, DARK)
text(s, 5.6, 6.6, 7.1, 0.4, "Your R2 payout is the same whichever option the buyer picks.", 14, MID, True)
footer(s, 7)

# 8 Process
s = prs.slides.add_slide(BL); bg(s)
heading(s, "How It Works", "From joining to getting paid")
steps = [("1", "Join the team", "Become one of 30 sales partners"),
         ("2", "Bring buyers", "Arrange site visits and share the layout"),
         ("3", "Close the sale", "Buyer books with DP or 4-month EMI"),
         ("4", "Get paid", "Rs 1,700 to Rs 2,000 per sq yd as per your slab")]
for i, (n, t, d) in enumerate(steps):
    x = 0.8 + i * 3.05
    rect(s, x, 2.0, 2.8, 3.8, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, x + 0.95, 2.3, 0.9, 0.9, GOLD if i == 3 else DARK, MSO_SHAPE.OVAL)
    text(s, x + 0.95, 2.3, 0.9, 0.9, n, 28, WHITE, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(s, x + 0.15, 3.45, 2.5, 0.6, t, 20, DARK, True, PP_ALIGN.CENTER)
    text(s, x + 0.15, 4.15, 2.5, 1.4, d, 15, GREY, False, PP_ALIGN.CENTER)
text(s, 0.8, 6.15, 11.8, 0.5, "Payout timing and terms are as per the sales partner agreement with the developer.", 13, GREY)
footer(s, 8)

# 9 CTA
s = prs.slides.add_slide(BL); overlay(s, "78000")
text(s, 0.9, 1.9, 11.5, 1.2, "Join the R2 Sales Team", 54, WHITE, True, PP_ALIGN.CENTER)
text(s, 0.9, 3.1, 11.5, 0.8, "Only 30 seats  |  Up to Rs 2,000 per sq yd  |  166 plots to sell", 22, GOLD, True, PP_ALIGN.CENTER)
rect(s, 3.9, 4.3, 5.5, 1.3, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE)
text(s, 3.9, 4.3, 5.5, 1.3, [("Call: [Your phone number]", 22, DARK, True), ("Changodar, Ahmedabad", 16, GREY, False)],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
prs.save(OUT); print("deck saved")
