from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

OUT = "output/Plotting_Sales_Plan.xlsx"
INR = '[>=10000000]##\\,##\\,##\\,##0;[>=100000]##\\,##\\,##0;##,##0'
PCT = '0.0%'
NAVY = "1F3864"
HDR = PatternFill("solid", fgColor=NAVY)
INP = PatternFill("solid", fgColor="FFF2CC")
KEY = PatternFill("solid", fgColor="E2EFDA")
SUB = PatternFill("solid", fgColor="D9E1F2")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WHITE = Font(bold=True, color="FFFFFF")
BLUE = Font(color="0000FF")

wb = Workbook()

def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = Font(bold=True, size=14, color=NAVY)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(italic=True, color="595959")

def header(ws, row, labels, col=1):
    for i, l in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=l)
        c.fill = HDR; c.font = WHITE; c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def put(ws, ref, val, fmt=None, fill=None, bold=False, inp=False):
    c = ws[ref]; c.value = val; c.border = BOX
    if fmt: c.number_format = fmt
    if inp: c.fill = INP; c.font = BLUE
    elif fill: c.fill = fill
    if bold: c.font = Font(bold=True, color="0000FF" if inp else "000000")
    return c

def widths(ws, w):
    for k, v in w.items(): ws.column_dimensions[k].width = v

# ---------------- Inputs ----------------
ws = wb.active; ws.title = "Inputs"
title(ws, "Changodar Plotting Project: Inputs", "Yellow cells are inputs. Change them and every other sheet updates.")
header(ws, 4, ["Input", "Value", "Unit", "Note"])
inputs = [
    ("Total land area", 25000, "sq yd", "Whole land the landowner owns", '#,##0'),
    ("Plot size", 150, "sq yd", "Size of each plot", '#,##0'),
    ("Common area (roads, parks, amenities)", 0, "%", "Usually 30% to 40% in a real layout. Set to 0% as asked, so all land becomes plots.", PCT),
    ("Landowner sale price to developer", 15000, "INR / sq yd", "Flat price paid to the landowner on the full land", INR),
    ("Developer selling price (regular)", 20000, "INR / sq yd", "Regular price to plot buyers", INR),
    ("DP discount", 500, "INR / sq yd", "DP = full payment within the DP window", INR),
    ("DP payment window", 14, "days", "Full payment within 2 weeks", '0'),
    ("EMI interest rate", 0.15, "% per year", "Charged on EMI buyers", PCT),
    ("EMI tenure", 4, "months", "", '0'),
    ("Number of salespeople", 30, "people", "", '0'),
    ("Development & other costs (optional)", 0, "INR lump sum", "Roads, electricity, legal, registration, marketing. Not given, so 0.", INR),
]
for i, (lab, val, unit, note, fmt) in enumerate(inputs):
    r = 5 + i
    put(ws, f"A{r}", lab); put(ws, f"B{r}", val, fmt, inp=True); put(ws, f"C{r}", unit); put(ws, f"D{r}", note)
    ws[f"D{r}"].alignment = Alignment(wrap_text=True, vertical="top")
# refs
LAND, PLOT, COMMON, LOPRICE, SELL, DPDISC, DPDAYS, RATE, TENURE, NSALES, DEVCOST = [f"Inputs!$B${5+i}" for i in range(11)]
r = 18
ws[f"A{r}"] = "Assumptions"; ws[f"A{r}"].font = Font(bold=True, color=NAVY)
notes = [
    "1. Plots = whole number of 150 sq yd plots that fit in the saleable land (25,000 / 150 = 166.67, so 166 plots; 100 sq yd is left over).",
    "2. The landowner sells the full land to the developer at a flat rate. The developer bears all sales commission and selling risk.",
    "3. \"25,000 x 19,000\" is read as the developer's net realisation of about 19,000 per sq yd after paying 5% to 8% commission on 20,000.",
    "4. Commission is slab based: the rate for the number of plots a person sells applies to ALL plots that person sold (not marginal).",
    "5. Ladder: 1 plot 5%, 2 plots 6%, 3 to 4 plots 6.5%, 5 plots 7.5%, 6 or more plots 8% (maximum). The 4 and 6+ slabs are assumed; edit them on the Commission Ladder sheet.",
    "6. Commission is paid on the plot sale value (DP price or regular price), not on EMI interest.",
    "7. EMI: flat interest = principal x 15% x 4/12 = 5% of principal, split into 4 equal monthly EMIs. A reducing-balance EMI is shown for comparison. No booking amount is assumed for EMI buyers.",
    "8. Profit figures are before development cost (unless entered above), stamp duty, GST and income tax.",
]
for i, n in enumerate(notes):
    ws.cell(row=r + 1 + i, column=1, value=n)
    ws.merge_cells(start_row=r + 1 + i, start_column=1, end_row=r + 1 + i, end_column=4)
    ws.cell(row=r + 1 + i, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r + 1 + i].height = 30
widths(ws, {"A": 40, "B": 16, "C": 14, "D": 70})
for rr in range(5, 16): ws.row_dimensions[rr].height = 30

# ---------------- Commission Ladder ----------------
cl = wb.create_sheet("Commission Ladder")
title(cl, "Salesperson Ladder Commission", "Slab based: the rate for your total plots applies to every plot you sold. Edit yellow cells.")
header(cl, 4, ["Plots sold (from)", "Commission rate", "Slab"])
ladder = [(0, 0, "No sale"), (1, 0.05, "1 plot"), (2, 0.06, "2 plots"), (3, 0.065, "3 to 4 plots"),
          (5, 0.075, "5 plots"), (6, 0.08, "6+ plots (maximum)")]
for i, (p, rt, s) in enumerate(ladder):
    r = 5 + i
    put(cl, f"A{r}", p, '0', inp=True); put(cl, f"B{r}", rt, PCT, inp=True); put(cl, f"C{r}", s)
LADDER = "'Commission Ladder'!$A$5:$B$10"
cl["A12"] = "What one salesperson earns (all plots at regular price)"; cl["A12"].font = Font(bold=True, color=NAVY)
header(cl, 13, ["Plots sold", "Rate", "Sale value (INR)", "Commission (INR)", "Commission per plot (INR)", "Developer net per sq yd (INR)"])
for i in range(1, 11):
    r = 13 + i
    put(cl, f"A{r}", i, '0')
    put(cl, f"B{r}", f"=VLOOKUP(A{r},{LADDER},2,TRUE)", PCT)
    put(cl, f"C{r}", f"=A{r}*{PLOT}*{SELL}", INR)
    put(cl, f"D{r}", f"=C{r}*B{r}", INR, fill=KEY)
    put(cl, f"E{r}", f"=D{r}/A{r}", INR)
    put(cl, f"F{r}", f"={SELL}*(1-B{r})", INR)
cl["A25"] = "Developer net per sq yd = 20,000 less commission. At 5% it is 19,000 (the figure in the brief); at 8% it is 18,400."
cl["A25"].font = Font(italic=True, color="595959")
widths(cl, {"A": 18, "B": 16, "C": 20, "D": 20, "E": 22, "F": 24})

# ---------------- Plot & Revenue ----------------
pr = wb.create_sheet("Plot & Revenue")
title(pr, "Plots and Revenue Summary")
header(pr, 4, ["Item", "Value", "Unit", "Formula / note"])
rows = [
    ("Total land", f"={LAND}", "sq yd", "Input", '#,##0'),
    ("Common area", f"=B5*{COMMON}", "sq yd", "Land x common area %", '#,##0'),
    ("Saleable land", "=B5-B6", "sq yd", "", '#,##0'),
    ("Plot size", f"={PLOT}", "sq yd", "", '#,##0'),
    ("Number of plots", "=INT(B7/B8)", "plots", "Whole plots only", '#,##0'),
    ("Area in plots", "=B9*B8", "sq yd", "", '#,##0'),
    ("Leftover land (not a full plot)", "=B7-B10", "sq yd", "Can be merged into a larger plot or common area", '#,##0'),
    ("Regular price per sq yd", f"={SELL}", "INR", "", INR),
    ("Regular price per plot", "=B12*B8", "INR", "", INR),
    ("DP price per sq yd", f"={SELL}-{DPDISC}", "INR", "Full payment within 2 weeks", INR),
    ("DP price per plot", "=B14*B8", "INR", "", INR),
    ("EMI total price per plot (flat interest)", f"=B13*(1+{RATE}*{TENURE}/12)", "INR", "See Payment Plans", INR),
    ("Gross revenue if all plots sold at regular price", "=B9*B13", "INR", "", INR),
    ("Gross revenue if all plots sold on DP", "=B9*B15", "INR", "", INR),
    ("Gross revenue on full land at regular price (25,000 x 20,000)", "=B5*B12", "INR", "Theoretical, if every sq yd were sold", INR),
    ("Net revenue on full land at 19,000 (the brief's figure)", "=B5*19000", "INR", "25,000 x 19,000", INR),
    ("Paid to landowner", f"=B5*{LOPRICE}", "INR", "Full land x 15,000", INR),
    ("Developer cost per saleable sq yd", "=B21/B10", "INR", "Land cost spread over plot area", INR),
]
for i, (lab, f, unit, note, fmt) in enumerate(rows):
    r = 5 + i
    put(pr, f"A{r}", lab); put(pr, f"B{r}", f, fmt); put(pr, f"C{r}", unit); put(pr, f"D{r}", note)
for r in (9, 13, 17, 21): pr[f"B{r}"].fill = KEY; pr[f"B{r}"].font = Font(bold=True)
widths(pr, {"A": 52, "B": 18, "C": 10, "D": 46})

# ---------------- Salesperson Tracker ----------------
st = wb.create_sheet("Salesperson Tracker")
title(st, "Salesperson Tracker (30 people)", "Enter plots sold and how many of those were DP. Sample figures distribute all 166 plots; replace with actuals.")
header(st, 4, ["#", "Salesperson", "Plots sold", "of which DP", "of which EMI", "Sale value (INR)", "Slab rate", "Commission (INR)", "EMI interest earned by developer (INR)"])
sample = [1,1,1,1, 2,2,2,2, 3,3,3,3,3, 4,4,4,4, 5,5,5,5, 7,8,9,9,10,11,13,15,21]
assert len(sample) == 30
diff = 166 - sum(sample); sample[-1] += diff
assert sum(sample) == 166
for i in range(30):
    r = 5 + i
    put(st, f"A{r}", i + 1, '0')
    put(st, f"B{r}", f"Salesperson {i+1}", inp=True)
    put(st, f"C{r}", sample[i], '0', inp=True)
    put(st, f"D{r}", sample[i] // 2, '0', inp=True)
    put(st, f"E{r}", f"=C{r}-D{r}", '0')
    put(st, f"F{r}", f"=D{r}*'Plot & Revenue'!$B$15+E{r}*'Plot & Revenue'!$B$13", INR)
    put(st, f"G{r}", f"=IF(C{r}<=0,0,VLOOKUP(C{r},{LADDER},2,TRUE))", PCT)
    put(st, f"H{r}", f"=F{r}*G{r}", INR, fill=KEY)
    put(st, f"I{r}", f"=E{r}*'Plot & Revenue'!$B$13*{RATE}*{TENURE}/12", INR)
T = 35
put(st, f"B{T}", "TOTAL", bold=True, fill=SUB)
for col, fmt in (("C", '0'), ("D", '0'), ("E", '0'), ("F", INR), ("H", INR), ("I", INR)):
    put(st, f"{col}{T}", f"=SUM({col}5:{col}34)", fmt, fill=SUB, bold=True)
put(st, f"G{T}", f"=IF(F{T}=0,0,H{T}/F{T})", PCT, fill=SUB, bold=True)
put(st, f"B{T+2}", "Plots available"); put(st, f"C{T+2}", "='Plot & Revenue'!B9", '0')
put(st, f"B{T+3}", "Plots still unsold"); put(st, f"C{T+3}", f"=C{T+2}-C{T}", '0')
put(st, f"B{T+4}", "Check"); put(st, f"C{T+4}", f'=IF(C{T}>C{T+2},"Over-allotted!",IF(C{T}=C{T+2},"All plots sold","OK"))')
st.merge_cells(f"C{T+4}:E{T+4}")
put(st, f"G{T+2}", "Average commission per person"); put(st, f"H{T+2}", f"=H{T}/{NSALES}", INR)
put(st, f"G{T+3}", "Highest commission"); put(st, f"H{T+3}", f"=MAX(H5:H34)", INR)
put(st, f"G{T+4}", "Lowest commission"); put(st, f"H{T+4}", f"=MIN(H5:H34)", INR)
dv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0", showErrorMessage=True, error="Enter a whole number, 0 or more")
st.add_data_validation(dv); dv.add("C5:D34")
dv2 = DataValidation(type="custom", formula1="D5<=C5", showErrorMessage=True, error="DP plots cannot exceed plots sold")
st.add_data_validation(dv2); dv2.add("D5:D34")
st.conditional_formatting.add(f"C{T+4}", CellIsRule(operator="equal", formula=['"Over-allotted!"'], fill=PatternFill("solid", fgColor="F8CBAD")))
st.freeze_panes = "C5"
widths(st, {"A": 5, "B": 18, "C": 11, "D": 11, "E": 11, "F": 18, "G": 30, "H": 18, "I": 22})

# ---------------- Payment Plans ----------------
pp = wb.create_sheet("Payment Plans")
title(pp, "Payment Plans per Plot: DP vs 4-month EMI")
header(pp, 4, ["Item", "DP (pay in 2 weeks)", "EMI, flat interest", "EMI, reducing balance", "Note"])
items = [
    ("Price per sq yd (base)", f"={SELL}-{DPDISC}", f"={SELL}", f"={SELL}", "DP gets 500 off per sq yd"),
    ("Plot size (sq yd)", f"={PLOT}", f"={PLOT}", f"={PLOT}", ""),
    ("Plot price / principal", "=B5*B6", "=C5*C6", "=D5*D6", ""),
    ("Interest rate per year", 0, f"={RATE}", f"={RATE}", ""),
    ("Tenure (months)", "=0", f"={TENURE}", f"={TENURE}", ""),
    ("Total interest", 0, "=C7*C8*C9/12", "=D11*D9-D7", "Flat = principal x 15% x 4/12"),
    ("Monthly EMI", "-", "=(C7+C10)/C9", "=PMT(D8/12,D9,-D7)", "4 equal payments"),
    ("Total paid by buyer", "=B7", "=C7+C10", "=D11*D9", ""),
    ("Effective price per sq yd", "=B12/B6", "=C12/C6", "=D12/D6", ""),
    ("Difference vs regular price", "=B12-$C$7", "=C12-$C$7", "=D12-$C$7", "Negative = buyer saves"),
    ("When paid", f'="Full amount within "&{DPDAYS}&" days"', f'={TENURE}&" monthly EMIs"', f'={TENURE}&" monthly EMIs"', ""),
]
for i, (lab, b, c, d, note) in enumerate(items):
    r = 5 + i
    put(pp, f"A{r}", lab)
    fmt = PCT if "rate" in lab else ('0' if "Tenure" in lab or "size" in lab else INR)
    put(pp, f"B{r}", b, fmt); put(pp, f"C{r}", c, fmt); put(pp, f"D{r}", d, fmt); put(pp, f"E{r}", note)
for col in "BCD": pp[f"{col}11"].fill = KEY; pp[f"{col}12"].fill = KEY; pp[f"{col}12"].font = Font(bold=True)
pp["A18"] = "EMI schedule per plot"; pp["A18"].font = Font(bold=True, color=NAVY)
header(pp, 19, ["Month", "Flat: EMI", "Flat: balance due", "Reducing: opening balance", "Reducing: interest", "Reducing: principal", "Reducing: EMI", "Reducing: closing balance"])
for m in range(1, 13):
    r = 19 + m
    put(pp, f"A{r}", f"=IF({m}<=$C$9,{m},\"\")", '0')
    put(pp, f"B{r}", f"=IF(A{r}=\"\",\"\",$C$11)", INR)
    put(pp, f"C{r}", f"=IF(A{r}=\"\",\"\",$C$12-$C$11*A{r})", INR)
    put(pp, f"D{r}", f"=IF(A{r}=\"\",\"\",IF(A{r}=1,$D$7,H{r-1}))" if m > 1 else f"=IF(A{r}=\"\",\"\",$D$7)", INR)
    put(pp, f"E{r}", f"=IF(A{r}=\"\",\"\",D{r}*$D$8/12)", INR)
    put(pp, f"F{r}", f"=IF(A{r}=\"\",\"\",G{r}-E{r})", INR)
    put(pp, f"G{r}", f"=IF(A{r}=\"\",\"\",$D$11)", INR)
    put(pp, f"H{r}", f"=IF(A{r}=\"\",\"\",D{r}-F{r})", INR)
pp["A33"] = "Schedule shows up to 12 months; rows beyond the EMI tenure stay blank."
pp["A33"].font = Font(italic=True, color="595959")
widths(pp, {"A": 30, "B": 20, "C": 20, "D": 24, "E": 34, "F": 20, "G": 16, "H": 22})

# ---------------- Profit Summary ----------------
ps = wb.create_sheet("Profit Summary")
title(ps, "Profit Summary: Landowner, Developer, Sales Team", "Uses the plots entered on the Salesperson Tracker.")
header(ps, 4, ["Item", "Amount (INR)", "In crore", "Note"])
TR = "'Salesperson Tracker'!"
prow = [
    ("LANDOWNER", None, None),
    ("Received from developer", f"='Plot & Revenue'!B21", "Full land x 15,000 per sq yd"),
    ("DEVELOPER", None, None),
    ("Plot sales value (DP + EMI principal)", f"={TR}F35", "From tracker"),
    ("EMI interest income", f"={TR}I35", "Flat interest on EMI plots"),
    ("Total inflow", "=B8+B9", ""),
    ("Less: paid to landowner", "=-B6", ""),
    ("Less: sales commission", f"=-{TR}H35", ""),
    ("Less: development & other costs", f"=-{DEVCOST}", "Optional input"),
    ("Developer net profit", "=SUM(B10:B13)", "Before taxes"),
    ("Developer margin on inflow", "=IF(B10=0,0,B14/B10)", ""),
    ("Return on land cost", "=IF(B6=0,0,B14/B6)", ""),
    ("Net realisation per sq yd sold (after commission)", f"=IF({TR}C35=0,0,(B8-{TR}H35)/({TR}C35*{PLOT}))", "Compare with 19,000 in the brief"),
    ("SALES TEAM (30 people)", None, None),
    ("Total commission", f"={TR}H35", ""),
    ("Average per salesperson", f"={TR}H37", ""),
    ("Highest earner", f"={TR}H38", ""),
    ("Lowest earner", f"={TR}H39", ""),
    ("Overall commission rate", f"={TR}G35", "Commission / sales value"),
    ("CHECK", None, None),
    ("Plots sold in tracker", f"={TR}C35", ""),
    ("Plots available", "='Plot & Revenue'!B9", ""),
]
for i, (lab, f, note) in enumerate(prow):
    r = 5 + i
    if f is None:
        c = put(ps, f"A{r}", lab, fill=SUB, bold=True)
        for col in "BCD": put(ps, f"{col}{r}", None, fill=SUB)
        continue
    is_pct = "margin" in lab or "Return" in lab or "rate" in lab
    is_cnt = "Plots" in lab
    put(ps, f"A{r}", lab)
    put(ps, f"B{r}", f, PCT if is_pct else ('0' if is_cnt else INR))
    if not (is_pct or is_cnt or "per sq yd" in lab):
        put(ps, f"C{r}", f"=B{r}/10000000", '0.00')
    else:
        put(ps, f"C{r}", None)
    put(ps, f"D{r}", note)
for r in (6, 14, 19): ps[f"B{r}"].fill = KEY; ps[f"B{r}"].font = Font(bold=True); ps[f"C{r}"].fill = KEY
widths(ps, {"A": 48, "B": 18, "C": 12, "D": 36})

for s in wb.worksheets:
    s.sheet_view.showGridLines = False
wb.save(OUT)
print("saved")
