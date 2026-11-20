"""Landowner deal plan: land sold to developer at a flat rate, paid over 8 months.
Output: output/Land_Deal_Payment_Plan.xlsx (recalculate with LibreOffice afterwards)."""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule

OUT = "output/Land_Deal_Payment_Plan.xlsx"
INR = '[>=10000000]##\\,##\\,##\\,##0;[>=100000]##\\,##\\,##0;##,##0'
NEG = '[<=-10000000]-##\\,##\\,##\\,##0;[<=-100000]-##\\,##\\,##0;#,##0'
CR = '0.00 "Cr"'; PCT = '0.0%'; DATE = 'dd-mmm-yyyy'
NAVY = "0F3D2E"
HDR = PatternFill("solid", fgColor=NAVY); INP = PatternFill("solid", fgColor="FFF2CC")
KEY = PatternFill("solid", fgColor="E2EFDA"); SUB = PatternFill("solid", fgColor="D9E1F2"); BAD = PatternFill("solid", fgColor="F8CBAD")
thin = Side(style="thin", color="BFBFBF"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
wb = Workbook()

def title(ws, t, sub=None):
    ws["A1"] = t; ws["A1"].font = Font(bold=True, size=14, color=NAVY)
    if sub: ws["A2"] = sub; ws["A2"].font = Font(italic=True, color="595959")
def header(ws, row, labels, col=1):
    for i, l in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=l); c.fill = HDR; c.font = Font(bold=True, color="FFFFFF"); c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 32
def put(ws, ref, val, fmt=None, fill=None, bold=False, inp=False):
    c = ws[ref]; c.value = val; c.border = BOX
    if fmt: c.number_format = fmt
    if inp: c.fill = INP; c.font = Font(color="0000FF", bold=bold)
    else:
        if fill: c.fill = fill
        if bold: c.font = Font(bold=True)
    return c
def widths(ws, w):
    for k, v in w.items(): ws.column_dimensions[k].width = v

# ---------------- Inputs ----------------
ws = wb.active; ws.title = "Inputs"
title(ws, "Land Deal Plan: Inputs", "Yellow cells are inputs. Every other sheet updates from them.")
header(ws, 4, ["Deal input", "Value", "Unit", "Note"])
deal = [("Land area sold to developer", 25000, "sq yd", "", '#,##0'),
        ("Price to landowners", 15000, "INR / sq yd", "Flat rate", INR),
        ("Agreement signing date", dt.date(2026, 12, 1), "date", "Month 0 = signing", DATE),
        ("Partner 1 name", "Ashish", "", "", None),
        ("Partner 1 share", 0.5, "%", "Partner 2 gets the rest", PCT),
        ("Partner 2 name", "Partner", "", "Type your partner's name", None),
        ("TDS deducted by developer", 0.01, "%", "Sec 194-IA: 1% on property over Rs 50 lakh; claim it back in your tax return", PCT),
        ("Late payment interest (for agreement)", 0.18, "% per year", "Suggested penalty if an instalment is late", PCT)]
for i, (a, b, c, d, f) in enumerate(deal):
    r = 5 + i; put(ws, f"A{r}", a); put(ws, f"B{r}", b, f, inp=True); put(ws, f"C{r}", c); put(ws, f"D{r}", d)
LAND, PRICE, SIGN, P1N, P1S, P2N, TDS, LATE = [f"Inputs!$B${5+i}" for i in range(8)]

ws["A14"] = "Payment schedule (% of total, must add up to 100%)"; ws["A14"].font = Font(bold=True, color=NAVY)
header(ws, 15, ["Month", "% of total", "Milestone", "Note"])
sched = [(0, .10, "Token / advance at signing", "Registered agreement to sell (banakhat) signed"),
         (1, .05, "Instalment 1", ""), (2, .10, "Instalment 2", ""), (3, .10, "Instalment 3", ""),
         (4, .15, "Instalment 4", ""), (5, .15, "Instalment 5", ""), (6, .15, "Instalment 6", ""),
         (7, .10, "Instalment 7", ""), (8, .10, "Final instalment", "Sale deed for remaining land on full payment")]
for i, (m, p, ms, n) in enumerate(sched):
    r = 16 + i; put(ws, f"A{r}", m, '0'); put(ws, f"B{r}", p, PCT, inp=True); put(ws, f"C{r}", ms, inp=True); put(ws, f"D{r}", n)
put(ws, "A25", "Total", bold=True, fill=SUB); put(ws, "B25", "=SUM(B16:B24)", PCT, bold=True, fill=SUB)
put(ws, "C25", '=IF(ROUND(B25,4)=1,"OK: adds to 100%","Fix: must add to 100%")', fill=SUB)
ws.conditional_formatting.add("C25", CellIsRule(operator="equal", formula=['"Fix: must add to 100%"'], fill=BAD))

ws["A27"] = "Developer sales assumptions (to check the developer can pay)"; ws["A27"].font = Font(bold=True, color=NAVY)
header(ws, 28, ["Input", "Value", "Unit", "Note"])
dev = [("Plot size", 150, "sq yd", "", '#,##0'), ("Regular selling price", 20000, "INR / sq yd", "", INR),
       ("DP discount", 500, "INR / sq yd", "DP = full payment within 2 weeks", INR),
       ("Share of buyers choosing DP", 0.5, "%", "Rest choose 4-month EMI", PCT),
       ("EMI interest rate (flat)", 0.15, "% per year", "EMI is 4 months; first EMI in the month of sale", PCT),
       ("Sales commission", 0.075, "% of sale value", "R1 ladder averages about 7.5%; R2 is about 9.8%", PCT)]
for i, (a, b, c, d, f) in enumerate(dev):
    r = 29 + i; put(ws, f"A{r}", a); put(ws, f"B{r}", b, f, inp=True); put(ws, f"C{r}", c); put(ws, f"D{r}", d)
PLOT, SELL, DPD, DPSH, RATE, COMM = [f"Inputs!$B${29+i}" for i in range(6)]
header(ws, 36, ["Month", "Plots sold", "", "Note"])
sales = [0, 10, 15, 20, 25, 25, 25, 25, 21]
for m, n in enumerate(sales):
    r = 37 + m; put(ws, f"A{r}", m, '0'); put(ws, f"B{r}", n, '0', inp=True)
put(ws, "A46", "Total", bold=True, fill=SUB); put(ws, "B46", "=SUM(B37:B45)", '0', bold=True, fill=SUB)
put(ws, "D46", '="Plots available: "&INT(B5/B29)')
widths(ws, {"A": 40, "B": 16, "C": 24, "D": 62})

# ---------------- Payment Schedule ----------------
ps = wb.create_sheet("Payment Schedule")
title(ps, "How We Get Paid: 8-Month Schedule", "Total land value is paid in full by month 8.")
put(ps, "A3", "Total land value"); put(ps, "B3", f"={LAND}*{PRICE}", INR, fill=KEY, bold=True); put(ps, "C3", "=B3/10000000", CR, fill=KEY, bold=True)
header(ps, 5, ["Month", "Due date", "Milestone", "% of total", "Amount due (INR)", "Cumulative (INR)", "Cumulative %",
               "TDS (INR)", "Net received (INR)", "=Inputs!B8&\" share (net)\"", "=Inputs!B10&\" share (net)\""])
for m in range(9):
    r = 6 + m; ir = 16 + m
    put(ps, f"A{r}", m, '0')
    put(ps, f"B{r}", f"=EDATE({SIGN},A{r})", DATE)
    put(ps, f"C{r}", f"=Inputs!C{ir}")
    put(ps, f"D{r}", f"=Inputs!B{ir}", PCT)
    put(ps, f"E{r}", f"=$B$3*D{r}", INR, fill=KEY, bold=True)
    put(ps, f"F{r}", f"=SUM($E$6:E{r})", INR)
    put(ps, f"G{r}", f"=F{r}/$B$3", PCT)
    put(ps, f"H{r}", f"=E{r}*{TDS}", INR)
    put(ps, f"I{r}", f"=E{r}-H{r}", INR)
    put(ps, f"J{r}", f"=I{r}*{P1S}", INR)
    put(ps, f"K{r}", f"=I{r}*(1-{P1S})", INR)
put(ps, "C15", "TOTAL", bold=True, fill=SUB)
for c, f in (("D", PCT), ("E", INR), ("H", INR), ("I", INR), ("J", INR), ("K", INR)):
    put(ps, f"{c}15", f"=SUM({c}6:{c}14)", f, bold=True, fill=SUB)
ps["A17"] = "TDS is not lost: the developer deposits it against your PAN and you claim it in your income tax return."
ps["A18"] = "Capital gains tax on the land sale is separate; plan it with your CA."
for r in (17, 18): ps[f"A{r}"].font = Font(italic=True, color="595959")
ps.freeze_panes = "A6"
widths(ps, {"A": 8, "B": 13, "C": 28, "D": 11, "E": 18, "F": 18, "G": 12, "H": 14, "I": 18, "J": 18, "K": 18})

# ---------------- Developer Cash Check ----------------
dc = wb.create_sheet("Developer Cash Check")
title(dc, "Can the Developer Pay on Time?", "Developer's cash from plot sales vs what is owed to us. Months 9 to 11 show EMI money still coming in.")
MONTHS = list(range(12)); C0 = 2  # column B = month 0
header(dc, 4, ["Item"] + [f"Month {m}" for m in MONTHS] + ["Total"])
col = lambda m: L(C0 + m)
labels = ["Plots sold", "DP plots", "EMI plots", "Cash from DP plots", "Cash from EMI plots", "Sales commission paid",
          "Developer net cash in", "Cumulative net cash in", "Land payment due to us", "Cumulative land payment",
          "Developer's own money needed (cumulative)"]
for i, lab in enumerate(labels):
    put(dc, f"A{5+i}", lab, bold=i in (6, 8, 10))
EMIPP = f"{PLOT}*{SELL}*(1+{RATE}*4/12)/4"   # monthly EMI per plot
for m in MONTHS:
    c = col(m)
    put(dc, f"{c}5", f"=Inputs!B{37+m}" if m <= 8 else 0, '0')
    put(dc, f"{c}6", f"={c}5*{DPSH}", '0.0')
    put(dc, f"{c}7", f"={c}5-{c}6", '0.0')
    put(dc, f"{c}8", f"={c}6*{PLOT}*({SELL}-{DPD})", INR)
    window = ",".join(f"{col(k)}7" for k in range(max(0, m - 3), m + 1))
    put(dc, f"{c}9", f"=SUM({window})*{EMIPP}", INR)
    put(dc, f"{c}10", f"=-({c}6*{PLOT}*({SELL}-{DPD})+{c}7*{PLOT}*{SELL})*{COMM}", NEG)
    put(dc, f"{c}11", f"={c}8+{c}9+{c}10", INR, fill=KEY, bold=True)
    put(dc, f"{c}12", f"=SUM(${col(0)}$11:{c}11)", INR)
    put(dc, f"{c}13", f"='Payment Schedule'!E{6+m}" if m <= 8 else 0, INR, bold=True)
    put(dc, f"{c}14", f"=SUM(${col(0)}$13:{c}13)", INR)
    put(dc, f"{c}15", f"=MAX(0,{c}14-{c}12)", INR, bold=True)
T = col(12)
for r, f in ((5, '0'), (6, '0.0'), (7, '0.0'), (8, INR), (9, INR), (10, NEG), (11, INR), (13, INR)):
    put(dc, f"{T}{r}", f"=SUM({col(0)}{r}:{col(11)}{r})", f, fill=SUB, bold=True)
dc.conditional_formatting.add(f"{col(0)}15:{col(11)}15", CellIsRule(operator="greaterThan", formula=["0"], fill=BAD))
put(dc, "A17", "Peak own money the developer must arrange", bold=True); put(dc, "B17", f"=MAX({col(0)}15:{col(11)}15)", INR, fill=KEY, bold=True)
put(dc, "C17", "=B17/10000000", CR, fill=KEY, bold=True)
put(dc, "A18", "Month of peak need"); put(dc, "B18", f"=MATCH(B17,{col(0)}15:{col(11)}15,0)-1", '0')
put(dc, "A19", "First month sales alone cover all payments due"); put(dc, "B19", f'=IFERROR(INDEX({col(0)}4:{col(11)}4,MATCH(TRUE,INDEX(({col(0)}15:{col(11)}15=0)*({col(0)}4:{col(11)}4<>"Month 0")>0,0),0)),"Not within 11 months")')
put(dc, "A20", "Cash from sales collected by month 8"); put(dc, "B20", f"={col(8)}12", INR)
put(dc, "A21", "Our total due by month 8"); put(dc, "B21", f"={col(8)}14", INR)
dc["A23"] = "Red cells: the developer must use their own funds or a loan to pay us on time that month. This is normal early on; agree it upfront."
dc["A23"].font = Font(italic=True, color="595959")
dc.freeze_panes = "B5"
widths(dc, {"A": 42, **{col(m): 13 for m in MONTHS}, T: 15})

# ---------------- Partner Summary ----------------
pt = wb.create_sheet("Partner Summary")
title(pt, "Partner Summary")
header(pt, 4, ["Item", "=Inputs!B8", "=Inputs!B10", "Total"])
rows = [("Share", f"={P1S}", f"=1-{P1S}", "=B5+C5", PCT),
        ("Gross amount", f"='Payment Schedule'!$E$15*B5", f"='Payment Schedule'!$E$15*C5", "=B6+C6", INR),
        ("TDS (claim back in tax return)", f"='Payment Schedule'!$H$15*B5", f"='Payment Schedule'!$H$15*C5", "=B7+C7", INR),
        ("Net cash received by month 8", "=B6-B7", "=C6-C7", "=B8+C8", INR),
        ("Gross in crore", "=B6/10000000", "=C6/10000000", "=D6/10000000", CR),
        ("Received at signing (month 0, net)", "='Payment Schedule'!J6", "='Payment Schedule'!K6", "=B10+C10", INR),
        ("Received by month 4 (net)", "=SUM('Payment Schedule'!J6:J10)", "=SUM('Payment Schedule'!K6:K10)", "=B11+C11", INR)]
for i, (a, b, c, d, f) in enumerate(rows):
    r = 5 + i; put(pt, f"A{r}", a); put(pt, f"B{r}", b, f); put(pt, f"C{r}", c, f); put(pt, f"D{r}", d, f)
for col_ in "BCD": pt[f"{col_}8"].fill = KEY; pt[f"{col_}8"].font = Font(bold=True)
widths(pt, {"A": 38, "B": 18, "C": 18, "D": 18})

# ---------------- Safeguards ----------------
sg = wb.create_sheet("Safeguards")
title(sg, "Safeguards for the Landowners", "Checklist to discuss with your lawyer before signing")
header(sg, 4, ["#", "Safeguard", "Why it matters", "Done?"])
items = [("Registered agreement to sell (banakhat) with the full schedule, dates and amounts", "Legally binds the developer to the price and timeline"),
         ("Token / advance of at least 10% at signing", "Shows commitment; forfeited if the developer walks away"),
         ("Post-dated cheques (or bank guarantee) for every instalment, given at signing", "Payment is ready on each due date; a bounced cheque is a legal offence"),
         ("All payments by bank transfer (RTGS/NEFT), never cash", "Clear record for tax and for any dispute"),
         ("Title stays with us until full payment; final sale deed only after month 8 payment", "We keep ownership if the developer defaults"),
         ("Release land in phases: plot registries allowed only for land already paid for", "Developer cannot sell land to buyers that we have not been paid for"),
         ("Late payment interest (e.g. 18% p.a.) and a cure period of 15 to 30 days", "Discourages delays"),
         ("Termination clause: after 2 missed instalments we can cancel and keep the token", "Protects us if sales are slow"),
         ("Developer pays stamp duty, registration and all development / approval costs", "Our Rs 37.5 Cr stays clean"),
         ("Partnership / co-owner agreement between the two of us on shares and decisions", "Avoids disputes between partners"),
         ("TDS certificates (Form 16B) collected for every payment", "Needed to claim TDS back"),
         ("Plan capital gains tax with a CA before the first payment", "Avoids a surprise tax bill")]
for i, (a, b) in enumerate(items):
    r = 5 + i
    put(sg, f"A{r}", i + 1, '0'); put(sg, f"B{r}", a); put(sg, f"C{r}", b); put(sg, f"D{r}", "", inp=True)
    for cc in "BC": sg[f"{cc}{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    sg.row_dimensions[r].height = 32
widths(sg, {"A": 5, "B": 62, "C": 52, "D": 10})

for s in wb.worksheets: s.sheet_view.showGridLines = False
wb.save(OUT); print("saved")
