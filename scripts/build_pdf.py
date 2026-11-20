from openpyxl import load_workbook
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm

wb = load_workbook("output/Plotting_Sales_Plan.xlsx", data_only=True)
PR, PP, PS, ST, CL = (wb[n] for n in ["Plot & Revenue", "Payment Plans", "Profit Summary", "Salesperson Tracker", "Commission Ladder"])

def inr(x, dec=0):
    neg = x < 0; x = abs(round(x))
    s = str(int(x))
    if len(s) > 3:
        head, tail = s[:-3], s[-3:]
        parts = []
        while len(head) > 2: parts.insert(0, head[-2:]); head = head[:-2]
        if head: parts.insert(0, head)
        s = ",".join(parts) + "," + tail
    return ("-" if neg else "") + "Rs " + s
def cr(x): return f"Rs {x/1e7:,.2f} Cr"
def lk(x): return f"Rs {x/1e5:,.2f} lakh"
def pct(x): return f"{x*100:.1f}%".replace(".0%", "%")

NAVY = colors.HexColor("#1F3864"); LIGHT = colors.HexColor("#D9E1F2"); GREEN = colors.HexColor("#E2EFDA")
ss = getSampleStyleSheet()
H1 = ParagraphStyle("H1", parent=ss["Title"], textColor=NAVY, fontSize=20, spaceAfter=4)
H2 = ParagraphStyle("H2", parent=ss["Heading2"], textColor=NAVY, fontSize=13, spaceBefore=10, spaceAfter=5)
B = ParagraphStyle("B", parent=ss["BodyText"], fontSize=9.5, leading=13)
SM = ParagraphStyle("SM", parent=B, fontSize=8.5, leading=11.5, textColor=colors.HexColor("#404040"))
SUBT = ParagraphStyle("S", parent=B, alignment=1, textColor=colors.HexColor("#595959"))

def tbl(data, widths, header=True, highlight=(), right_from=1, font=9):
    t = Table(data, colWidths=[w*mm for w in widths], hAlign="LEFT")
    st = [("FONT", (0,0), (-1,-1), "Helvetica", font),
          ("GRID", (0,0), (-1,-1), 0.4, colors.HexColor("#BFBFBF")),
          ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
          ("ALIGN", (right_from,0), (-1,-1), "RIGHT"),
          ("TOPPADDING", (0,0), (-1,-1), 3), ("BOTTOMPADDING", (0,0), (-1,-1), 3)]
    if header:
        st += [("BACKGROUND", (0,0), (-1,0), NAVY), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
               ("FONT", (0,0), (-1,0), "Helvetica-Bold", font), ("ALIGN", (0,0), (-1,0), "CENTER")]
    for r in highlight:
        st += [("BACKGROUND", (0,r), (-1,r), GREEN), ("FONT", (0,r), (-1,r), "Helvetica-Bold", font)]
    t.setStyle(TableStyle(st)); return t

v = lambda ws, ref: ws[ref].value
plots, area_plots, leftover = v(PR,"B9"), v(PR,"B10"), v(PR,"B11")
land_paid = v(PS,"B6"); sales = v(PS,"B8"); emi_int = v(PS,"B9"); comm = v(PS,"B19")
dev_profit = v(PS,"B14")

story = []
story += [Paragraph("Changodar Plotting Project", H1),
          Paragraph("Plot sales plan: land deal, sales team ladder commission and payment options", SUBT),
          Spacer(1, 8)]

story.append(Paragraph("At a glance", H2))
glance = [["Item", "Figure"],
    ["Total land", "25,000 sq yd"],
    ["Plot size / number of plots", f"150 sq yd / {plots} plots ({leftover} sq yd left over)"],
    ["Landowner receives (flat Rs 15,000 per sq yd)", cr(land_paid)],
    ["Regular plot price (Rs 20,000 per sq yd)", f"{lk(v(PR,'B13'))} per plot"],
    ["DP plot price (Rs 19,500 per sq yd, pay in 2 weeks)", f"{lk(v(PR,'B15'))} per plot"],
    ["4-month EMI (15% p.a., flat)", f"{inr(v(PP,'C11'))} per month x 4 = {lk(v(PP,'C12'))}"],
    ["Sales team commission (sample plan, 30 people)", f"{cr(comm)} ({pct(v(PS,'B23'))} of sales)"],
    ["Developer net profit (sample plan, before costs and tax)", cr(dev_profit)],
]
story.append(tbl(glance, [95, 80], highlight=(3, 8), right_from=9))
story.append(Spacer(1, 6))
story.append(Paragraph(
    "The landowner sells the whole 25,000 sq yd to the developer at a flat Rs 15,000 per sq yd and is paid "
    f"<b>{cr(land_paid)}</b>. The developer cuts the land into {plots} plots of 150 sq yd and sells them at Rs 20,000 per sq yd "
    "through 30 salespeople on a ladder commission of 5% to 8%. After commission the developer keeps about Rs 18,400 to Rs 19,000 "
    "per sq yd, which matches the \"25,000 x 19,000\" figure in the brief (19,000 is the net at the lowest 5% slab).", B))

story.append(Paragraph("1. Plots and revenue", H2))
rev = [["Item", "Value"],
    ["Saleable land (common area set to 0%)", f"{v(PR,'B7'):,} sq yd"],
    ["Number of 150 sq yd plots", f"{plots}"],
    ["Area in plots", f"{area_plots:,} sq yd"],
    ["Gross revenue, all plots at regular price", cr(v(PR,"B17"))],
    ["Gross revenue, all plots on DP", cr(v(PR,"B18"))],
    ["Full land at Rs 20,000 (25,000 x 20,000)", cr(v(PR,"B19"))],
    ["Full land at Rs 19,000 net (25,000 x 19,000)", cr(v(PR,"B20"))],
    ["Paid to landowner (25,000 x 15,000)", cr(v(PR,"B21"))]]
story.append(tbl(rev, [95, 80], highlight=(2,)))
story.append(Spacer(1, 4))
story.append(Paragraph(
    "<b>Note on common area:</b> a real plotting layout usually gives 30% to 40% of the land to roads, parks and amenities. "
    "At 35% the project would have about 108 plots instead of 166. The Excel has an input for this, set to 0% as asked.", SM))

story.append(PageBreak())
story.append(Paragraph("2. Ladder commission for the 30 salespeople", H2))
story.append(Paragraph(
    "The rate depends on how many plots a salesperson sells, and that rate applies to <b>all</b> their plots. "
    "Commission is paid on the plot sale value.", B))
story.append(Spacer(1, 4))
lad = [["Plots sold", "Rate", "Sale value", "Commission earned", "Per plot", "Developer keeps / sq yd"]]
for r in range(14, 24):
    a,b,c,d,e,f = (CL.cell(r,k).value for k in range(1,7))
    lad.append([str(a), pct(b), lk(c), inr(d), inr(e), inr(f)])
story.append(tbl(lad, [20, 16, 30, 35, 28, 40], font=8.5))
story.append(Paragraph("Slabs: 1 plot 5%, 2 plots 6%, 3 to 4 plots 6.5%, 5 plots 7.5%, 6 or more plots 8% (maximum). "
                       "The 4-plot and 6+ slabs were assumed, since the brief jumps from 3 to 5 and caps at 8%.", SM))

story.append(Paragraph("3. Payment options per plot", H2))
pay = [["", "DP (pay in 2 weeks)", "EMI, flat 15%", "EMI, reducing 15%"],
    ["Price per sq yd", inr(v(PP,"B5")), inr(v(PP,"C5")), inr(v(PP,"D5"))],
    ["Plot price / principal", inr(v(PP,"B7")), inr(v(PP,"C7")), inr(v(PP,"D7"))],
    ["Interest for 4 months", "-", inr(v(PP,"C10")), inr(v(PP,"D10"))],
    ["Monthly EMI (x4)", "-", inr(v(PP,"C11")), inr(v(PP,"D11"))],
    ["Total paid by buyer", inr(v(PP,"B12")), inr(v(PP,"C12")), inr(v(PP,"D12"))],
    ["Effective price per sq yd", inr(v(PP,"B13")), inr(v(PP,"C13")), inr(v(PP,"D13"))],
    ["vs regular plot price", inr(v(PP,"B14")), "+" + inr(v(PP,"C14")), "+" + inr(v(PP,"D14"))]]
story.append(tbl(pay, [48, 42, 42, 42], highlight=(5,)))
story.append(Spacer(1, 4))
story.append(Paragraph(
    "<b>How the EMI is worked out (flat, as asked):</b> Rs 30,00,000 x 15% x 4/12 = Rs 1,50,000 interest. "
    "Rs 31,50,000 / 4 = <b>Rs 7,87,500 per month</b>. A bank-style reducing-balance EMI would be lower at "
    f"{inr(v(PP,'D11'))} per month. A DP buyer saves Rs 75,000 against the regular price, and Rs 2,25,000 against the flat EMI.", SM))

story.append(PageBreak())
story.append(Paragraph("4. Sample sales plan: 166 plots across 30 salespeople", H2))
story.append(Paragraph("A sample split so every plot is sold, with about half of each person's plots on DP and the rest on EMI. "
                       "Replace with real figures in the Salesperson Tracker sheet.", SM))
story.append(Spacer(1, 3))
rows = [["#", "Plots", "DP", "EMI", "Sale value", "Rate", "Commission"]]
for r in range(5, 35):
    c = [ST.cell(r,k).value for k in range(1,9)]
    rows.append([str(c[0]), str(c[2]), str(c[3]), str(c[4]), lk(c[5]), pct(c[6]), inr(c[7])])
c = [ST.cell(35,k).value for k in range(1,9)]
rows.append(["Total", str(c[2]), str(c[3]), str(c[4]), cr(c[5]), pct(c[6]), inr(c[7])])
half = 16
left = rows[:half]; right = [rows[0]] + rows[half:]
lt = tbl(left, [8, 11, 9, 9, 22, 11, 22], font=7.5)
rt = tbl(right, [10, 11, 9, 9, 22, 11, 22], font=7.5, highlight=(len(right)-1,))
story.append(Table([[lt, rt]], colWidths=[93*mm, 96*mm], style=[("VALIGN",(0,0),(-1,-1),"TOP"), ("LEFTPADDING",(0,0),(-1,-1),0)]))
story.append(Spacer(1, 4))
story.append(Paragraph(f"Average earning per salesperson: <b>{inr(v(PS,'B20'))}</b>. Highest: {inr(v(PS,'B21'))} (21 plots). "
                       f"Lowest: {inr(v(PS,'B22'))} (1 plot).", B))

story.append(Paragraph("5. Who earns what (sample plan)", H2))
prof = [["Item", "Amount"],
    ["Landowner: received from developer", cr(land_paid)],
    ["Developer: plot sales value (DP + EMI principal)", cr(sales)],
    ["Developer: EMI interest income", cr(emi_int)],
    ["Developer: less paid to landowner", "-" + cr(land_paid)],
    ["Developer: less sales commission", "-" + cr(comm)],
    ["Developer: net profit (before development cost and tax)", cr(dev_profit)],
    ["Developer: return on land cost", pct(v(PS,"B16"))],
    ["Developer: net per sq yd sold after commission", inr(v(PS,"B17"))],
    ["Sales team: total commission (30 people)", cr(comm)]]
story.append(tbl(prof, [110, 65], highlight=(1, 6, 9)))
story.append(Spacer(1, 6))
story.append(Paragraph("Key assumptions", H2))
for n in [
    "166 whole plots fit in 25,000 sq yd at 150 sq yd each; 100 sq yd is left over. Common area is 0% unless changed.",
    "The landowner is paid on the full 25,000 sq yd. The developer bears commission, selling risk and all development costs.",
    "Commission is slab based on the plots each person sells (not marginal), paid on sale value, not on EMI interest.",
    "EMI buyers pay no booking amount; flat interest of 15% p.a. for 4 months. DP means full payment within 14 days.",
    "Profit excludes development, legal, registration and marketing costs, GST, stamp duty and income tax. Add costs in the Excel Inputs sheet.",
]:
    story.append(Paragraph("&bull; " + n, SM))

def footer(c, d):
    c.saveState(); c.setFont("Helvetica", 7.5); c.setFillColor(colors.grey)
    c.drawString(15*mm, 9*mm, "Changodar Plotting Project: sales plan summary. Figures from Plotting_Sales_Plan.xlsx")
    c.drawRightString(195*mm, 9*mm, f"Page {d.page}"); c.restoreState()
doc = SimpleDocTemplate("output/Plotting_Sales_Plan_Summary.pdf", pagesize=A4,
                        leftMargin=15*mm, rightMargin=15*mm, topMargin=14*mm, bottomMargin=15*mm,
                        title="Changodar Plotting Project: Sales Plan")
doc.build(story, onFirstPage=footer, onLaterPages=footer)
print("ok")
