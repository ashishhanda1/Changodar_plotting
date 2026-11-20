"""Buyer marketing deck. Usage: python scripts/build_deck.py [en|hi|gu] (run build_images.py <lang> first)."""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from deckkit import *
from pptx.util import Inches

D = Deck(sys.argv[1] if len(sys.argv) > 1 else "en"); T, R = D.T, D.rs
FOOT = T("Changodar Plots  |  150 sq yd residential plots")

s = D.slide(); D.overlay(s, "72000")
D.text(s, 0.9, 1.9, 11.5, 0.6, T("CHANGODAR, AHMEDABAD"), 20, GOLD, True)
D.text(s, 0.9, 2.5, 11.5, 1.6, T("Own Your Plot at Changodar"), 54, WHITE, True)
D.text(s, 0.9, 3.75, 11.5, 0.8, T("166 residential plots of 150 sq yd each  |  Easy DP and 4-month EMI options"), 22, WHITE)
D.rect(s, 0.9, 4.9, 4.2, 0.85, GOLD, MSO_SHAPE.ROUNDED_RECTANGLE)
D.text(s, 0.9, 4.9, 4.2, 0.85, R(T("From Rs 19,500 / sq yd")), 22, DARK, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)

s = D.slide(); D.bg(s)
D.heading(s, T("The Project at a Glance"), T("A planned plotting project with clear titles and simple payment options"))
for i, (b, sm) in enumerate([("25,000", "sq yd total land"), ("166", "plots available"),
                             ("150 sq yd", "each plot (30 ft x 45 ft)"), ("Rs 30 lakh", "regular price per plot")]):
    D.card(s, 0.7 + i * 3.05, 2.0, 2.8, 2.0, R(T(b)), T(sm), WHITE, DARK if i % 2 == 0 else MID)
D.text(s, 0.7, 4.5, 12, 2.2, [(T("Why invest in a plot?"), 22, DARK, True)] +
       [("•  " + T(x), 18, DARK, False) for x in ("Land you own outright, ready to build your home on your own timeline",
        "Ideal size for a family home: 1,350 sq ft of land", "Pay in full and save, or spread the payment over 4 months")])
D.footer(s, 2, FOOT)

s = D.slide(); D.bg(s)
D.heading(s, T("Layout and Plot Size"), T("Illustrative layout. The final layout follows the approved plan."))
s.shapes.add_picture(D.A + "site_layout.png", Inches(0.7), Inches(1.75), width=Inches(7.6))
pic = s.shapes.add_picture(D.A + "single_plot.png", Inches(8.6), Inches(1.6), height=Inches(4.3))
pic.left = Inches(10.75) - int(pic.width / 2)
D.text(s, 8.6, 5.95, 4.3, 1.0, [(T("Every plot: 150 sq yd"), 20, DARK, True), (T("30 ft wide x 45 ft deep, road facing"), 15, GREY, False)], align=PP_ALIGN.CENTER)
D.text(s, 0.7, 6.15, 7.6, 0.6, T("Internal roads  •  Green park  •  Tree-lined boundary"), 16, MID, True, PP_ALIGN.CENTER)
D.footer(s, 3, FOOT)

s = D.slide(); D.bg(s)
D.heading(s, T("Simple, Transparent Pricing"), T("Choose the option that suits you"))
D.rect(s, 0.8, 1.9, 5.6, 4.4, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE)
D.text(s, 1.0, 2.15, 5.2, 0.5, T("REGULAR PRICE"), 16, GREY, True, PP_ALIGN.CENTER)
D.text(s, 1.0, 2.6, 5.2, 1.0, R("Rs 20,000"), 48, DARK, True, PP_ALIGN.CENTER)
D.text(s, 1.0, 3.55, 5.2, 0.5, T("per sq yd"), 18, GREY, False, PP_ALIGN.CENTER)
D.text(s, 1.0, 4.3, 5.2, 1.6, [(R(T("Rs 30,00,000 per plot")), 24, DARK, True), (T("Pay over 4 months with EMI"), 16, GREY, False)], align=PP_ALIGN.CENTER)
D.rect(s, 6.9, 1.9, 5.6, 4.4, DARK, MSO_SHAPE.ROUNDED_RECTANGLE)
D.rect(s, 10.3, 1.55, 2.4, 0.5, GOLD, MSO_SHAPE.ROUNDED_RECTANGLE)
D.text(s, 10.3, 1.55, 2.4, 0.5, T("BEST VALUE"), 15, DARK, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
D.text(s, 7.1, 2.15, 5.2, 0.5, T("DP PRICE: PAY IN 2 WEEKS"), 16, GOLD, True, PP_ALIGN.CENTER)
D.text(s, 7.1, 2.6, 5.2, 1.0, R("Rs 19,500"), 48, WHITE, True, PP_ALIGN.CENTER)
D.text(s, 7.1, 3.55, 5.2, 0.5, T("per sq yd"), 18, LIGHT, False, PP_ALIGN.CENTER)
D.text(s, 7.1, 4.3, 5.2, 1.6, [(R(T("Rs 29,25,000 per plot")), 24, WHITE, True), (R(T("You save Rs 75,000")), 20, GOLD, True)], align=PP_ALIGN.CENTER)
D.text(s, 0.8, 6.45, 12, 0.4, T("Prices exclude stamp duty, registration and other government charges."), 12, GREY)
D.footer(s, 4, FOOT)

s = D.slide(); D.bg(s)
D.heading(s, T("Easy 4-Month EMI"), T("Book now and pay in 4 equal monthly instalments"))
pic = s.shapes.add_picture(D.A + "emi_timeline.png", 0, Inches(1.75), height=Inches(2.7))
pic.left = int((D.prs.slide_width - pic.width) / 2)
for i, (b, sm) in enumerate([("Rs 30,00,000", "plot price"), ("15% p.a.", "interest, for 4 months"),
                             ("Rs 7,87,500", "monthly EMI x 4"), ("Rs 31,50,000", "total payable")]):
    D.card(s, 0.7 + i * 3.05, 4.85, 2.8, 1.75, R(T(b)), T(sm), GOLD if i == 2 else WHITE, DARK, DARK if i == 2 else GREY, 26)
D.footer(s, 5, FOOT)

s = D.slide(); D.bg(s)
D.heading(s, T("DP or EMI: Compare at a Glance"), T("Per 150 sq yd plot"))
rows = [("", T("DP (pay in 2 weeks)"), T("4-month EMI")), (T("Price per sq yd"), R("Rs 19,500"), R("Rs 20,000")),
        (T("Plot price"), R("Rs 29,25,000"), R("Rs 30,00,000")), (T("Interest"), T("None"), R(T("Rs 1,50,000 (15% p.a. for 4 months)"))),
        (T("Payment"), T("Full amount within 14 days"), R(T("Rs 7,87,500 every month x 4"))),
        (T("Total you pay"), R("Rs 29,25,000"), R("Rs 31,50,000")), (T("Effective price per sq yd"), R("Rs 19,500"), R("Rs 21,000"))]
D.table(s, rows, 0.8, 1.8, 11.7, 4.4, [3.5, 3.8, 4.4], 17, lambda r: r == 5, LIGHT, left_first=True)
D.text(s, 0.8, 6.35, 11.7, 0.5, R(T("Pay in full and save Rs 2,25,000 compared with the EMI option.")), 18, MID, True)
D.footer(s, 6, FOOT)

s = D.slide(); D.bg(s, DARK)
D.text(s, 0.8, 0.45, 11.5, 0.8, T("Join Our Sales Team: Ladder Income"), 34, WHITE, True)
D.text(s, 0.82, 1.1, 11.5, 0.5, T("30 sales partners. The more plots you sell, the higher your rate."), 17, LIGHT)
D.rect(s, 0.8, 1.8, 7.4, 4.9, CREAM, MSO_SHAPE.ROUNDED_RECTANGLE)
s.shapes.add_picture(D.A + "ladder.png", Inches(1.0), Inches(2.0), width=Inches(7.0))
D.text(s, 8.7, 2.0, 4.0, 4.8, [(T("How it works"), 22, GOLD, True)] + [("•  " + R(T(x)), 17, WHITE, False) for x in (
    "Commission on the plot sale value", "Your rate applies to ALL plots you sell, not just the extra ones",
    "Sell 6 or more plots to reach the top 8% rate", "Earn Rs 1.5 lakh on your very first plot")])
D.footer(s, 7, FOOT, dark=True)

s = D.slide(); D.bg(s)
D.heading(s, T("What You Can Earn"), R(T("Commission at the regular price of Rs 30 lakh per plot")))
s.shapes.add_picture(D.A + "earnings.png", Inches(0.6), Inches(1.65), height=Inches(4.8))
D.text(s, 0.6, 6.45, 8.4, 0.4, R(T("Bars: plots you sell (1 to 10) and what you earn in Rs lakh")), 13, GREY, False, PP_ALIGN.CENTER)
for i, (b, sm) in enumerate([("Rs 1.5 lakh", "sell 1 plot (5%)"), ("Rs 11.25 lakh", "sell 5 plots (7.5%)"), ("Rs 24 lakh", "sell 10 plots (8%)")]):
    D.card(s, 9.3, 1.75 + i * 1.65, 3.4, 1.45, R(T(b)), T(sm), GOLD if i == 2 else WHITE, DARK, DARK if i == 2 else GREY, 26)
D.footer(s, 8, FOOT)

s = D.slide(); D.bg(s)
D.heading(s, T("How to Book Your Plot"), T("Four simple steps"))
steps = [("Visit the site", "See the layout and choose your plot"), ("Pick a payment plan", "DP at Rs 19,500 / sq yd or 4-month EMI"),
         ("Pay", "Full amount within 14 days, or EMI 1 in month 1"), ("Get registered", "Sale deed registered in your name")]
for i, (t, d) in enumerate(steps):
    x = 0.8 + i * 3.05
    D.rect(s, x, 2.0, 2.8, 3.8, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE)
    D.rect(s, x + 0.95, 2.3, 0.9, 0.9, GOLD if i == 3 else DARK, MSO_SHAPE.OVAL)
    D.text(s, x + 0.95, 2.3, 0.9, 0.9, str(i + 1), 28, WHITE, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    D.text(s, x + 0.15, 3.45, 2.5, 0.6, T(t), 20, DARK, True, PP_ALIGN.CENTER)
    D.text(s, x + 0.15, 4.15, 2.5, 1.4, R(T(d)), 15, GREY, False, PP_ALIGN.CENTER)
D.footer(s, 9, FOOT)

s = D.slide(); D.overlay(s, "78000")
D.text(s, 0.9, 1.6, 11.5, 1.2, T("Book Your Plot Today"), 54, WHITE, True, PP_ALIGN.CENTER)
D.text(s, 0.9, 2.85, 11.5, 0.8, R(T("Only 166 plots  |  DP from Rs 29.25 lakh  |  EMI Rs 7,87,500 x 4")), 22, GOLD, True, PP_ALIGN.CENTER)
D.contact(s, 3.4, 4.0, 6.5, 2.0, "Site office: Changodar, Ahmedabad")
D.save("Changodar_Plots_Marketing")
