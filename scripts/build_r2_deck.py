"""R2 sales-partner deck. Usage: python scripts/build_r2_deck.py [en|hi|gu] (run build_images.py <lang> first)."""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from deckkit import *
from pptx.util import Inches
from pptx.dml.color import RGBColor

D = Deck(sys.argv[1] if len(sys.argv) > 1 else "en"); T, R = D.T, D.rs
FOOT = T("R2 Sales Partner Plan  |  Changodar Plots")

s = D.slide(); D.overlay(s)
D.text(s, 0.9, 1.7, 11.5, 0.6, T("FOR SALES PARTNERS  |  CHANGODAR PLOTS"), 18, GOLD, True)
D.text(s, 0.9, 2.3, 11.5, 1.2, T("R2 Sales Partner Plan"), 56, WHITE, True)
D.text(s, 0.9, 3.55, 11.5, 0.8, R(T("Earn up to Rs 2,000 for every sq yd you sell")), 28, WHITE)
D.rect(s, 0.9, 4.75, 5.0, 0.9, GOLD, MSO_SHAPE.ROUNDED_RECTANGLE)
D.text(s, 0.9, 4.75, 5.0, 0.9, R(T("Rs 3 lakh per plot at the top slab")), 20, DARK, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)

s = D.slide(); D.bg(s)
D.heading(s, T("The Opportunity"), T("A large, ready inventory and a small, focused sales team"))
for i, (b, sm) in enumerate([("166", "plots to sell"), ("150 sq yd", "every plot (30 ft x 45 ft)"),
                             ("Rs 30 lakh", "value of each plot"), ("30", "sales partners only")]):
    D.card(s, 0.7 + i * 3.05, 1.95, 2.8, 2.0, R(T(b)), T(sm), GOLD if i == 3 else WHITE, DARK, DARK if i == 3 else GREY, 32)
D.text(s, 0.7, 4.45, 12, 2.3, [(T("Why this works for you"), 22, DARK, True)] + [("•  " + T(x), 18, DARK, False) for x in (
    "About 5 to 6 plots per partner on average if all 166 plots are sold",
    "A fixed rupee amount per sq yd, so you always know exactly what you earn",
    "Buyers get two easy options (DP or 4-month EMI), which helps you close faster")])
D.footer(s, 2, FOOT)

s = D.slide(); D.bg(s, DARK)
D.heading(s, T("Your R2 Income Slabs"), T("Fixed rupees for every sq yd you sell"), dark=True)
D.rect(s, 0.7, 1.8, 7.4, 4.95, CREAM, MSO_SHAPE.ROUNDED_RECTANGLE)
s.shapes.add_picture(D.A + "r2_ladder.png", Inches(0.95), Inches(2.0), width=Inches(6.9))
D.text(s, 8.6, 1.9, 4.2, 4.9, [(T("How it works"), 22, GOLD, True)] + [("•  " + R(T(x)), 18, WHITE, False) for x in (
    "1 plot: Rs 1,700 per sq yd", "2 to 4 plots: Rs 1,800 per sq yd", "5 plots and above: Rs 2,000 per sq yd",
    "Your slab rate applies to ALL the plots you sell, not just the extra ones", "Same payout whether the buyer picks DP or EMI")])
D.footer(s, 3, FOOT, dark=True)

s = D.slide(); D.bg(s)
D.heading(s, T("What You Can Earn"), T("Each plot is 150 sq yd"))
s.shapes.add_picture(D.A + "r2_earnings.png", Inches(0.6), Inches(1.6), height=Inches(4.85))
D.text(s, 0.6, 6.45, 8.4, 0.4, R(T("Bars: plots you sell (1 to 10) and what you earn in Rs lakh")), 13, GREY, False, PP_ALIGN.CENTER)
for i, (b, sm) in enumerate([("Rs 2.55 lakh", "sell 1 plot"), ("Rs 15 lakh", "sell 5 plots"), ("Rs 30 lakh", "sell 10 plots")]):
    D.card(s, 9.3, 1.75 + i * 1.65, 3.4, 1.45, R(T(b)), T(sm), GOLD if i == 2 else WHITE, DARK, DARK if i == 2 else GREY, 28)
D.footer(s, 4, FOOT)

s = D.slide(); D.bg(s)
D.heading(s, T("Earnings Table"), T("Your total earning by the number of plots you sell"))
rate = lambda k: 1700 if k == 1 else 1800 if k <= 4 else 2000
lakh = lambda v: R(T("Rs {v} lakh").format(v=f"{v:g}"))
rows = [(T("Plots sold"), T("Rate per sq yd"), T("Earning per plot"), T("Total earning"))]
for k in (1, 2, 3, 4, 5, 6, 8, 10):
    r = rate(k); rows.append((str(k), R(f"Rs {r:,}"), lakh(r * 150 / 1e5), lakh(k * r * 150 / 1e5)))
D.table(s, rows, 1.4, 1.75, 10.5, 4.8, None, 18, lambda r: r >= 5, RGBColor.from_string("FBEFC8"), boldcol=3)
D.text(s, 1.4, 6.6, 10.5, 0.4, R(T("Highlighted rows are the top slab: Rs 2,000 per sq yd on every plot.")), 14, GREY)
D.footer(s, 5, FOOT)

s = D.slide(); D.bg(s)
D.heading(s, T("Push for Your 5th Plot"), T("Hitting the top slab lifts the rate on every plot you have sold"))
D.card(s, 0.9, 2.0, 3.6, 2.6, R(T("Rs 10.8 lakh")), R(T("4 plots at Rs 1,800 / sq yd")), WHITE, DARK, GREY, 34)
arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(4.8), Inches(2.85), Inches(1.4), Inches(0.9))
arrow.fill.solid(); arrow.fill.fore_color.rgb = GOLD; arrow.line.fill.background()
D.card(s, 6.5, 2.0, 3.6, 2.6, R(T("Rs 15 lakh")), R(T("5 plots at Rs 2,000 / sq yd")), DARK, WHITE, LIGHT, 34)
D.card(s, 10.4, 2.0, 2.3, 2.6, R(T("+Rs 4.2 lakh")), T("from one more sale"), GOLD, DARK, DARK, 24)
D.text(s, 0.9, 5.0, 11.8, 1.6, [
    (R(T("One extra plot adds Rs 4.2 lakh: Rs 3 lakh for that plot, plus Rs 1.2 lakh more on your first 4.")), 20, DARK, True),
    (R(T("In the same way, a 2nd plot takes you from Rs 2.55 lakh to Rs 5.4 lakh.")), 18, GREY, False)])
D.footer(s, 6, FOOT)

s = D.slide(); D.bg(s)
D.heading(s, T("Your Pitch Kit: What You Are Selling"), T("The key facts to share with buyers"))
pic = s.shapes.add_picture(D.A + "single_plot.png", Inches(0.7), Inches(1.65), height=Inches(4.9))
pic.left = Inches(3.0) - int(pic.width / 2)
facts = [("Plot", "150 sq yd (30 ft x 45 ft), road facing"), ("Regular price", "Rs 20,000 / sq yd = Rs 30 lakh"),
         ("DP price", "Rs 19,500 / sq yd = Rs 29.25 lakh, full payment in 2 weeks (buyer saves Rs 75,000)"),
         ("4-month EMI", "Rs 7,87,500 x 4 = Rs 31.5 lakh (15% p.a. flat interest)")]
for i, (k, v) in enumerate(facts):
    y = 1.75 + i * 1.2
    D.rect(s, 5.6, y, 7.1, 1.05, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE); D.rect(s, 5.6, y, 0.12, 1.05, GOLD)
    D.text(s, 5.9, y + 0.06, 6.7, 0.4, T(k), 15, MID, True)
    D.text(s, 5.9, y + 0.42, 6.7, 0.6, R(T(v)), 15, DARK)
D.text(s, 5.6, 6.6, 7.1, 0.4, T("Your R2 payout is the same whichever option the buyer picks."), 14, MID, True)
D.footer(s, 7, FOOT)

s = D.slide(); D.bg(s)
D.heading(s, T("How It Works"), T("From joining to getting paid"))
steps = [("Join the team", "Become one of 30 sales partners"), ("Bring buyers", "Arrange site visits and share the layout"),
         ("Close the sale", "Buyer books with DP or 4-month EMI"), ("Get paid", "Rs 1,700 to Rs 2,000 per sq yd as per your slab")]
for i, (t, d) in enumerate(steps):
    x = 0.8 + i * 3.05
    D.rect(s, x, 2.0, 2.8, 3.8, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE)
    D.rect(s, x + 0.95, 2.3, 0.9, 0.9, GOLD if i == 3 else DARK, MSO_SHAPE.OVAL)
    D.text(s, x + 0.95, 2.3, 0.9, 0.9, str(i + 1), 28, WHITE, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    D.text(s, x + 0.15, 3.45, 2.5, 0.6, T(t), 20, DARK, True, PP_ALIGN.CENTER)
    D.text(s, x + 0.15, 4.15, 2.5, 1.4, R(T(d)), 15, GREY, False, PP_ALIGN.CENTER)
D.text(s, 0.8, 6.15, 11.8, 0.5, T("Payout timing and terms are as per the sales partner agreement with the developer."), 13, GREY)
D.footer(s, 8, FOOT)

s = D.slide(); D.overlay(s, "78000")
D.text(s, 0.9, 1.5, 11.5, 1.2, T("Join the R2 Sales Team"), 54, WHITE, True, PP_ALIGN.CENTER)
D.text(s, 0.9, 2.75, 11.5, 0.8, R(T("Only 30 seats  |  Up to Rs 2,000 per sq yd  |  166 plots to sell")), 22, GOLD, True, PP_ALIGN.CENTER)
D.contact(s, 3.4, 4.0, 6.5, 2.0, "Changodar, Ahmedabad")
D.save("R2_Sales_Partner_Plan")
