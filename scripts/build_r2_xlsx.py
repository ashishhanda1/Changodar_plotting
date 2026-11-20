"""R2 plan: salespeople earn a fixed INR per sq yd based on plots sold."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "output/R2_Sales_Plan.xlsx"
INR = '[>=10000000]##\\,##\\,##\\,##0;[>=100000]##\\,##\\,##0;##,##0'
NEG_INR = '[<=-10000000]-##\\,##\\,##\\,##0;[<=-100000]-##\\,##\\,##0;#,##0'
PCT = '0.0%'
NAVY = "0F3D2E"
HDR = PatternFill("solid", fgColor=NAVY); INP = PatternFill("solid", fgColor="FFF2CC")
KEY = PatternFill("solid", fgColor="E2EFDA"); SUB = PatternFill("solid", fgColor="D9E1F2")
thin = Side(style="thin", color="BFBFBF"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WHITE = Font(bold=True, color="FFFFFF"); BLUE = Font(color="0000FF")
wb = Workbook()

def title(ws, t, sub=None):
    ws["A1"] = t; ws["A1"].font = Font(bold=True, size=14, color=NAVY)
    if sub: ws["A2"] = sub; ws["A2"].font = Font(italic=True, color="595959")
def header(ws, row, labels):
    for i, l in enumerate(labels):
        c = ws.cell(row=row, column=1 + i, value=l); c.fill = HDR; c.font = WHITE; c.border = BOX
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

# Inputs
ws = wb.active; ws.title = "Inputs"
title(ws, "R2 Sales Plan: Inputs", "Yellow cells are inputs. Every other sheet updates from them.")
header(ws, 4, ["Input", "Value", "Unit", "Note"])
ins = [("Total land area", 25000, "sq yd", "", '#,##0'),
       ("Plot size", 150, "sq yd", "", '#,##0'),
       ("Common area (roads, parks)", 0, "%", "Usually 30% to 40% in a real layout; 0% as planned", PCT),
       ("Landowner price to developer", 15000, "INR / sq yd", "Flat, on full land", INR),
       ("Regular selling price", 20000, "INR / sq yd", "", INR),
       ("DP discount", 500, "INR / sq yd", "DP = full payment within 2 weeks", INR),
       ("EMI interest rate", 0.15, "% per year", "Flat interest", PCT),
       ("EMI tenure", 4, "months", "", '0'),
       ("Number of salespeople", 30, "people", "", '0'),
       ("Development & other costs (optional)", 0, "INR lump sum", "Not given, so 0", INR)]
for i, (a, b, c, d, f) in enumerate(ins):
    r = 5 + i; put(ws, f"A{r}", a); put(ws, f"B{r}", b, f, inp=True); put(ws, f"C{r}", c); put(ws, f"D{r}", d)
LAND, PLOT, COMMON, LOPRICE, SELL, DPDISC, RATE, TENURE, NSALES, DEVCOST = [f"Inputs!$B${5+i}" for i in range(10)]
ws["A17"] = "R2 commission slabs (fixed INR per sq yd sold)"; ws["A17"].font = Font(bold=True, color=NAVY)
header(ws, 18, ["Plots sold (from)", "INR per sq yd", "Slab", "Note"])
r2 = [(0, 0, "No sale", ""), (1, 1700, "1 plot", ""), (2, 1800, "2 to 4 plots", ""), (5, 2000, "5 plots and above", "Maximum")]
for i, (a, b, c, d) in enumerate(r2):
    r = 19 + i; put(ws, f"A{r}", a, '0', inp=True); put(ws, f"B{r}", b, INR, inp=True); put(ws, f"C{r}", c); put(ws, f"D{r}", d)
R2T = "Inputs!$A$19:$B$22"
ws["A24"] = "R1 percentage ladder (for comparison only)"; ws["A24"].font = Font(bold=True, color=NAVY)
header(ws, 25, ["Plots sold (from)", "Rate", "Slab", "Note"])
r1 = [(0, 0, "No sale", ""), (1, .05, "1 plot", ""), (2, .06, "2 plots", ""), (3, .065, "3 to 4 plots", ""),
      (5, .075, "5 plots", ""), (6, .08, "6+ plots", "Maximum")]
for i, (a, b, c, d) in enumerate(r1):
    r = 26 + i; put(ws, f"A{r}", a, '0', inp=True); put(ws, f"B{r}", b, PCT, inp=True); put(ws, f"C{r}", c); put(ws, f"D{r}", d)
R1T = "Inputs!$A$26:$B$31"
notes = ["Assumptions",
         "1. R2 rate is slab based: the INR per sq yd for a person's total plots applies to ALL their plots.",
         "2. R2 pays the same per sq yd whether the buyer chooses DP or EMI (it does not depend on the price).",
         "3. R1 (percentage) is paid on the sale value: DP plots at 19,500 and EMI plots at 20,000 per sq yd.",
         "4. Profit is before development cost (unless entered), stamp duty, GST and income tax."]
for i, n in enumerate(notes):
    c = ws.cell(row=33 + i, column=1, value=n)
    if i == 0: c.font = Font(bold=True, color=NAVY)
widths(ws, {"A": 38, "B": 16, "C": 18, "D": 50})

# R2 Earnings ladder
el = wb.create_sheet("R2 Earnings")
title(el, "What a Salesperson Earns under R2", "Fixed INR per sq yd. Compared with R1 at the regular price.")
header(el, 4, ["Plots sold", "R2 INR / sq yd", "R2 earning (INR)", "R2 per plot (INR)", "R2 as % of sale", "R1 rate", "R1 earning (INR)", "R2 extra vs R1 (INR)"])
for k in range(1, 13):
    r = 4 + k
    put(el, f"A{r}", k, '0')
    put(el, f"B{r}", f"=VLOOKUP(A{r},{R2T},2,TRUE)", INR)
    put(el, f"C{r}", f"=B{r}*A{r}*{PLOT}", INR, fill=KEY, bold=True)
    put(el, f"D{r}", f"=C{r}/A{r}", INR)
    put(el, f"E{r}", f"=B{r}/{SELL}", PCT)
    put(el, f"F{r}", f"=VLOOKUP(A{r},{R1T},2,TRUE)", PCT)
    put(el, f"G{r}", f"=A{r}*{PLOT}*{SELL}*F{r}", INR)
    put(el, f"H{r}", f"=C{r}-G{r}", INR)
widths(el, {"A": 11, "B": 14, "C": 18, "D": 17, "E": 14, "F": 10, "G": 17, "H": 20})

# Tracker
st = wb.create_sheet("Salesperson Tracker")
title(st, "R2 Salesperson Tracker (30 people)", "Enter plots sold and DP plots. Sample figures sell all 166 plots; replace with actuals.")
header(st, 4, ["#", "Salesperson", "Plots sold", "of which DP", "of which EMI", "Sq yd sold", "Sale value (INR)", "R2 INR / sq yd", "R2 commission (INR)", "R1 commission (INR)", "R2 minus R1 (INR)"])
sample = [1,1,1,1, 2,2,2,2, 3,3,3,3,3, 4,4,4,4, 5,5,5,5, 7,8,9,9,10,11,13,15,21]
assert len(sample) == 30 and sum(sample) == 166
for i in range(30):
    r = 5 + i
    put(st, f"A{r}", i + 1, '0'); put(st, f"B{r}", f"Salesperson {i+1}", inp=True)
    put(st, f"C{r}", sample[i], '0', inp=True); put(st, f"D{r}", sample[i] // 2, '0', inp=True)
    put(st, f"E{r}", f"=C{r}-D{r}", '0')
    put(st, f"F{r}", f"=C{r}*{PLOT}", '#,##0')
    put(st, f"G{r}", f"=D{r}*{PLOT}*({SELL}-{DPDISC})+E{r}*{PLOT}*{SELL}", INR)
    put(st, f"H{r}", f"=IF(C{r}<=0,0,VLOOKUP(C{r},{R2T},2,TRUE))", INR)
    put(st, f"I{r}", f"=F{r}*H{r}", INR, fill=KEY, bold=True)
    put(st, f"J{r}", f"=IF(C{r}<=0,0,G{r}*VLOOKUP(C{r},{R1T},2,TRUE))", INR)
    put(st, f"K{r}", f"=I{r}-J{r}", INR)
T = 35
put(st, f"B{T}", "TOTAL", fill=SUB, bold=True)
for col, f in (("C", '0'), ("D", '0'), ("E", '0'), ("F", '#,##0'), ("G", INR), ("I", INR), ("J", INR), ("K", INR)):
    put(st, f"{col}{T}", f"=SUM({col}5:{col}34)", f, fill=SUB, bold=True)
put(st, f"H{T}", f"=IF(F{T}=0,0,I{T}/F{T})", INR, fill=SUB, bold=True)
put(st, f"B{T+2}", "Plots available"); put(st, f"C{T+2}", f"=INT({LAND}*(1-{COMMON})/{PLOT})", '0')
put(st, f"B{T+3}", "Plots still unsold"); put(st, f"C{T+3}", f"=C{T+2}-C{T}", '0')
put(st, f"B{T+4}", "Check"); put(st, f"C{T+4}", f'=IF(C{T}>C{T+2},"Over-allotted!",IF(C{T}=C{T+2},"All plots sold","OK"))')
put(st, f"H{T+2}", "Average R2 per person"); put(st, f"I{T+2}", f"=I{T}/{NSALES}", INR)
put(st, f"H{T+3}", "Highest R2 earner"); put(st, f"I{T+3}", f"=MAX(I5:I34)", INR)
put(st, f"H{T+4}", "Lowest R2 earner"); put(st, f"I{T+4}", f"=MIN(I5:I34)", INR)
dv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0", showErrorMessage=True, error="Enter a whole number, 0 or more")
st.add_data_validation(dv); dv.add("C5:D34")
st.freeze_panes = "C5"
widths(st, {"A": 5, "B": 17, "C": 10, "D": 10, "E": 10, "F": 11, "G": 17, "H": 22, "I": 19, "J": 18, "K": 17})

# Developer margin: R2 vs R1
dm = wb.create_sheet("Developer Margin")
title(dm, "Developer Margin: R2 vs R1", "Uses the plots on the Salesperson Tracker.")
header(dm, 4, ["Item", "R2 (INR per sq yd)", "R1 (percentage)", "Difference (R2 - R1)"])
TR = "'Salesperson Tracker'!"
rows = [("Plot sales value (DP + EMI principal)", f"={TR}G35", f"={TR}G35", INR),
        ("EMI interest income", f"={TR}E35*{PLOT}*{SELL}*{RATE}*{TENURE}/12", None, INR),
        ("Total inflow", "=B5+B6", "=C5+C6", INR),
        ("Paid to landowner", f"=-{LAND}*{LOPRICE}", "=B8", NEG_INR),
        ("Sales commission", f"=-{TR}I35", f"=-{TR}J35", NEG_INR),
        ("Development & other costs", f"=-{DEVCOST}", "=B10", NEG_INR),
        ("Developer net profit", "=SUM(B7:B10)", "=SUM(C7:C10)", INR),
        ("Commission as % of sales", f"=-B9/B5", f"=-C9/C5", PCT),
        ("Average commission per sq yd sold", f"=-B9/{TR}F35", f"=-C9/{TR}F35", INR),
        ("Developer net per sq yd sold (sales less commission)", f"=(B5+B9)/{TR}F35", f"=(C5+C9)/{TR}F35", INR),
        ("Return on land cost", "=B11/-B8", "=C11/-C8", PCT)]
for i, (lab, b, c, f) in enumerate(rows):
    r = 5 + i
    put(dm, f"A{r}", lab); put(dm, f"B{r}", b, f); put(dm, f"C{r}", c if c else f"=B{r}", f)
    put(dm, f"D{r}", f"=B{r}-C{r}", PCT if f == PCT else NEG_INR)
for col in "ABCD":
    dm[f"{col}11"].fill = KEY; dm[f"{col}11"].font = Font(bold=True)
dm["A18"] = "Break-even check: R2 at 2,000 per sq yd is 10% of the regular price, more than R1's 8% maximum. R2 pays salespeople more and costs the developer more."
dm["A18"].font = Font(italic=True, color="595959")
widths(dm, {"A": 50, "B": 20, "C": 20, "D": 22})

for s in wb.worksheets: s.sheet_view.showGridLines = False
wb.save(OUT); print("saved")
