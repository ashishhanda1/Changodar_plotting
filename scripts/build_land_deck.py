"""Landowner deck (for the two partners). Reads the recalculated output/Land_Deal_Payment_Plan.xlsx."""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from deckkit import *
from pptx.util import Inches
from openpyxl import load_workbook
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

wb = load_workbook("output/Land_Deal_Payment_Plan.xlsx", data_only=True)
IN, PS, DC, PT = wb["Inputs"], wb["Payment Schedule"], wb["Developer Cash Check"], wb["Partner Summary"]
cr = lambda v: f"Rs {v/1e7:,.2f} Cr"
crs = lambda v: f"{v/1e7:,.2f}"
lk = lambda v: f"Rs {v/1e5:,.2f} lakh"
P1, P2, SH = IN["B8"].value, IN["B10"].value, IN["B9"].value
TOTAL = PS["B3"].value
months = [PS.cell(6 + m, 1).value for m in range(9)]
due = [PS.cell(6 + m, 5).value for m in range(9)]
pct = [PS.cell(6 + m, 4).value for m in range(9)]
dates = [PS.cell(6 + m, 2).value for m in range(9)]
miles = [PS.cell(6 + m, 3).value for m in range(9)]
cum = [PS.cell(6 + m, 6).value for m in range(9)]
cum_cash = [DC.cell(12, 2 + m).value for m in range(12)]
cum_land = [DC.cell(14, 2 + m).value for m in range(12)]
peak, peak_m, cover_m = DC["B17"].value, DC["B18"].value, DC["B19"].value
cash_by8 = DC["B20"].value
G_DARK, G_MID, C_GOLD = "#0F3D2E", "#2E7D5B", "#D4A017"
os.makedirs("assets/en", exist_ok=True)

# Chart 1: schedule
fig, ax = plt.subplots(figsize=(11, 5.6), dpi=150); fig.patch.set_alpha(0); ax.set_facecolor("none")
b = ax.bar(months, [d / 1e7 for d in due], color=[C_GOLD if m == 0 else G_MID for m in months], width=0.65, label="Paid that month")
for r, d in zip(b, due): ax.text(r.get_x() + r.get_width()/2, d/1e7 + 0.35, f"{d/1e7:.2f}", ha="center", fontsize=12, fontweight="bold", color=G_DARK)
ax2 = ax.twinx(); ax2.plot(months, [c / 1e7 for c in cum], color=G_DARK, marker="o", lw=3, label="Total received so far")
ax2.set_ylim(0, 40); ax.set_ylim(0, 8)
ax.set_xticks(months); ax.set_xticklabels(["Signing"] + [f"M{m}" for m in months[1:]])
ax.set_ylabel("Paid in the month (Rs Cr)", fontsize=12, color=G_DARK); ax2.set_ylabel("Total received (Rs Cr)", fontsize=12, color=G_DARK)
for a in (ax, ax2):
    a.tick_params(labelsize=12, colors=G_DARK); a.spines["top"].set_visible(False)
h1, l1 = ax.get_legend_handles_labels(); h2, l2 = ax2.get_legend_handles_labels()
ax.legend(h1 + h2, l1 + l2, loc="upper left", frameon=False, fontsize=12)
fig.savefig("assets/en/land_schedule.png", bbox_inches="tight", transparent=True); plt.close(fig)

# Chart 2: developer cash vs our payments
x = list(range(12))
fig, ax = plt.subplots(figsize=(11, 5.6), dpi=150); fig.patch.set_alpha(0); ax.set_facecolor("none")
cc = [v / 1e7 for v in cum_cash]; cl = [v / 1e7 for v in cum_land]
ax.plot(x, cc, color=G_MID, lw=3.5, marker="o", label="Developer's cash from plot sales (cumulative)")
ax.plot(x, cl, color=C_GOLD, lw=3.5, marker="s", label="Owed to us (cumulative)")
ax.fill_between(x, cc, cl, where=[l > c for c, l in zip(cc, cl)], color="#E57373", alpha=0.3, interpolate=True, label="Developer uses own funds")
ax.set_xticks(x); ax.set_xticklabels(["Sign"] + [f"M{m}" for m in x[1:]])
ax.set_ylabel("Rs Cr", fontsize=12, color=G_DARK); ax.tick_params(labelsize=12, colors=G_DARK)
for s in ("top", "right"): ax.spines[s].set_visible(False)
ax.legend(loc="upper left", frameon=False, fontsize=12)
fig.savefig("assets/en/land_devcash.png", bbox_inches="tight", transparent=True); plt.close(fig)

D = Deck("en"); T = D.T
FOOT = f"Private: land deal plan for {P1} and {P2}"

s = D.slide(); D.overlay(s, "78000")
D.text(s, 0.9, 1.7, 11.5, 0.6, "PRIVATE  |  FOR THE LANDOWNERS", 18, GOLD, True)
D.text(s, 0.9, 2.3, 11.5, 1.2, "Our Land Deal: Changodar", 56, WHITE, True)
D.text(s, 0.9, 3.55, 11.5, 0.8, f"25,000 sq yd to the developer at Rs 15,000 / sq yd. {cr(TOTAL)}, fully paid in 8 months.", 24, WHITE)
D.rect(s, 0.9, 4.75, 5.2, 0.9, GOLD, MSO_SHAPE.ROUNDED_RECTANGLE)
D.text(s, 0.9, 4.75, 5.2, 0.9, f"{P1} and {P2}: {SH:.0%} / {1-SH:.0%}", 22, DARK, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)

s = D.slide(); D.bg(s)
D.heading(s, "The Deal at a Glance", "We sell the whole land at a flat rate; the developer takes the selling risk")
for i, (b_, sm) in enumerate([("25,000 sq yd", "land sold"), ("Rs 15,000", "per sq yd, flat"), (cr(TOTAL), "total land value"), ("8 months", "to full payment")]):
    D.card(s, 0.7 + i * 3.05, 1.95, 2.8, 2.0, b_, sm, GOLD if i == 2 else WHITE, DARK, DARK if i == 2 else GREY, 28)
g1, g2 = PT["B6"].value, PT["C6"].value
D.text(s, 0.7, 4.45, 12, 2.3, [("What it means for us", 22, DARK, True),
    (f"•  {P1} receives {cr(g1)} and {P2} receives {cr(g2)} (before 1% TDS, which is claimed back)", 18, DARK, False),
    (f"•  {cr(due[0])} comes at signing as a token; the rest arrives in 8 monthly instalments", 18, DARK, False),
    ("•  No selling, commission or development cost on our side", 18, DARK, False)])
D.footer(s, 2, FOOT)

s = D.slide(); D.bg(s)
D.heading(s, "How the Deal Works", "Four stages from signing to final sale deed")
steps = [("Sign", f"Registered agreement to sell. We receive the {pct[0]:.0%} token: {cr(due[0])}"),
         ("Developer sells", "166 plots at Rs 20,000 / sq yd through 30 salespeople"),
         ("Monthly payments", "Developer pays us every month from sales, backed by post-dated cheques"),
         ("Month 8: done", "Final payment received, sale deed for the remaining land")]
for i, (t, d) in enumerate(steps):
    x0 = 0.8 + i * 3.05
    D.rect(s, x0, 2.0, 2.8, 4.0, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE)
    D.rect(s, x0 + 0.95, 2.3, 0.9, 0.9, GOLD if i == 3 else DARK, MSO_SHAPE.OVAL)
    D.text(s, x0 + 0.95, 2.3, 0.9, 0.9, str(i + 1), 28, WHITE, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    D.text(s, x0 + 0.15, 3.45, 2.5, 0.6, t, 20, DARK, True, PP_ALIGN.CENTER)
    D.text(s, x0 + 0.15, 4.15, 2.5, 1.7, d, 15, GREY, False, PP_ALIGN.CENTER)
D.footer(s, 3, FOOT)

s = D.slide(); D.bg(s)
D.heading(s, "When the Money Comes In", "Rs crore received each month, and the running total")
s.shapes.add_picture("assets/en/land_schedule.png", Inches(0.6), Inches(1.75), width=Inches(9.1))
D.card(s, 10.1, 1.8, 2.7, 1.5, cr(due[0]), "at signing", GOLD, DARK, DARK, 22)
D.card(s, 10.1, 3.5, 2.7, 1.5, cr(cum[4]), "by month 4 (50%)", WHITE, DARK, GREY, 22)
D.card(s, 10.1, 5.2, 2.7, 1.5, cr(cum[8]), "by month 8 (100%)", DARK, WHITE, LIGHT, 22)
D.footer(s, 4, FOOT)

s = D.slide(); D.bg(s)
D.heading(s, "Payment Schedule", f"Signing on {dates[0]:%d %b %Y}; all amounts before 1% TDS")
rows = [("Month", "Due date", "Milestone", "% of total", "Amount", "Total so far")]
for m in range(9):
    rows.append((str(m), f"{dates[m]:%d %b %Y}", miles[m], f"{pct[m]:.0%}", cr(due[m]), cr(cum[m])))
D.table(s, rows, 0.8, 1.65, 11.7, 5.2, [1.1, 1.9, 3.3, 1.5, 1.95, 1.95], 15, lambda r: r in (1, 9), RGBColor.from_string("FBEFC8"))
D.footer(s, 5, FOOT)

s = D.slide(); D.bg(s)
D.heading(s, "Partner Split", f"{SH:.0%} / {1-SH:.0%}, change the share in the spreadsheet")
for i, (name, colL) in enumerate(((P1, "B"), (P2, "C"))):
    x0 = 0.8 + i * 6.1
    D.rect(s, x0, 1.8, 5.7, 4.8, WHITE if i else DARK, MSO_SHAPE.ROUNDED_RECTANGLE)
    fg, sub = (DARK, GREY) if i else (WHITE, LIGHT)
    D.text(s, x0 + 0.3, 1.95, 5.1, 0.6, f"{name}  ({PT[colL + '5'].value:.0%})", 24, GOLD if not i else MID, True)
    D.text(s, x0 + 0.3, 2.6, 5.1, 1.0, cr(PT[colL + "6"].value), 44, fg, True)
    D.text(s, x0 + 0.3, 3.55, 5.1, 0.5, "gross share of the land value", 15, sub)
    D.text(s, x0 + 0.3, 4.2, 5.1, 2.3, [
        (f"TDS (claimed back): {lk(PT[colL + '7'].value)}", 17, fg, False),
        (f"Net cash by month 8: {cr(PT[colL + '8'].value)}", 17, fg, True),
        (f"At signing: {cr(PT[colL + '10'].value)}", 17, fg, False),
        (f"By month 4: {cr(PT[colL + '11'].value)}", 17, fg, False)])
D.footer(s, 6, FOOT)

s = D.slide(); D.bg(s)
D.heading(s, "Can the Developer Pay on Time?", "Developer's cash from plot sales vs what is owed to us (Rs crore, cumulative)")
s.shapes.add_picture("assets/en/land_devcash.png", Inches(0.6), Inches(1.75), width=Inches(9.1))
D.text(s, 9.9, 1.8, 3.0, 5.0, [
    ("What this shows", 20, DARK, True),
    (f"•  Sales cannot cover the token and early instalments, so the developer needs about {cr(peak)} of own money (peak in month {peak_m})", 15, DARK, False),
    (f"•  From {cover_m} onward, plot sales alone cover everything owed to us", 15, DARK, False),
    (f"•  By month 8 the developer has collected {cr(cash_by8)} from sales, more than the {cr(TOTAL)} owed", 15, DARK, False),
    ("•  Assumes 166 plots sold in 8 months; check the developer's funds before signing", 15, MID, True)])
D.footer(s, 7, FOOT)

s = D.slide(); D.bg(s, DARK)
D.heading(s, "Safeguards Before We Sign", "Discuss each point with our lawyer", dark=True)
sg = ["Registered agreement to sell (banakhat) with the full schedule", "At least 10% token at signing, forfeited if the developer walks away",
      "Post-dated cheques or a bank guarantee for every instalment", "All payments by bank transfer, never cash",
      "Title stays with us until the final payment", "Plot registries only for land already paid for",
      "18% p.a. late interest, cure period of 15 to 30 days", "We can cancel after 2 missed instalments and keep the token"]
for i, t in enumerate(sg):
    x0 = 0.8 + (i % 2) * 6.1; y0 = 1.85 + (i // 2) * 1.2
    D.rect(s, x0, y0, 5.8, 1.0, RGBColor.from_string("1C5240"), MSO_SHAPE.ROUNDED_RECTANGLE)
    D.rect(s, x0 + 0.2, y0 + 0.25, 0.5, 0.5, GOLD, MSO_SHAPE.OVAL)
    D.text(s, x0 + 0.2, y0 + 0.25, 0.5, 0.5, "✓", 16, DARK, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    D.text(s, x0 + 0.9, y0, 4.8, 1.0, t, 16, WHITE, False, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
D.footer(s, 8, FOOT, dark=True)

s = D.slide(); D.bg(s)
D.heading(s, "Tax and Paperwork", "Points to plan with our CA")
D.text(s, 0.8, 1.8, 11.8, 4.8, [
    ("TDS", 20, MID, True),
    (f"The developer deducts 1% TDS on each payment ({lk(PS['H15'].value)} in total) and deposits it against our PAN. We collect Form 16B for every payment and claim it in our tax returns.", 17, DARK, False),
    ("Capital gains", 20, MID, True),
    ("The land sale is taxed as capital gains. The amount depends on how long we have held the land and its purchase cost. Plan this with a CA before the first payment.", 17, DARK, False),
    ("Costs on the developer", 20, MID, True),
    ("Stamp duty, registration, layout approvals and development costs should all be paid by the developer, as written in the agreement.", 17, DARK, False),
    ("Between the two of us", 20, MID, True),
    ("A short co-owner agreement recording our shares, who signs, and where the money is received.", 17, DARK, False)])
D.footer(s, 9, FOOT)

s = D.slide(); D.overlay(s, "80000")
D.text(s, 0.9, 1.2, 11.5, 1.0, "Next Steps", 50, WHITE, True, PP_ALIGN.CENTER)
nxt = ["Agree the 8-month schedule and token with the developer", "Lawyer drafts the registered agreement to sell with all safeguards",
       "Collect the token, post-dated cheques and Form 16B process at signing", "Track every instalment monthly in the spreadsheet"]
for i, t in enumerate(nxt):
    y0 = 2.5 + i * 1.0
    D.rect(s, 2.4, y0, 8.5, 0.8, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE)
    D.text(s, 2.6, y0, 0.6, 0.8, str(i + 1), 24, GOLD, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    D.text(s, 3.3, y0, 7.5, 0.8, t, 18, DARK, True, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
D.save("Land_Deal_Plan_for_Landowners")
