# Changodar Plotting Project: Sales Plan

Plot sales plan for 25,000 sq yd of land divided into 150 sq yd plots (166 plots).

## Files

- `output/Plotting_Sales_Plan.xlsx`: workbook with live formulas. Yellow cells are inputs.
  Sheets: Inputs, Commission Ladder, Plot & Revenue, Salesperson Tracker (30 people), Payment Plans (DP vs EMI), Profit Summary.
- `output/Plotting_Sales_Plan_Summary.pdf`: 3-page summary of the plan.
- `scripts/build_xlsx.py`, `scripts/build_pdf.py`: regenerate the files.

## Key numbers (default inputs, sample sales split)

| Item | Figure |
|---|---|
| Plots | 166 x 150 sq yd (100 sq yd left over) |
| Landowner receives (25,000 x Rs 15,000) | Rs 37.50 Cr |
| Regular plot price (Rs 20,000/sq yd) | Rs 30.00 lakh |
| DP plot price (Rs 19,500/sq yd, pay in 14 days) | Rs 29.25 lakh |
| 4-month EMI, 15% p.a. flat | Rs 7,87,500 x 4 = Rs 31.50 lakh |
| Sales commission (30 people, 5% to 8% ladder) | Rs 3.69 Cr |
| Developer net profit (before costs and tax) | Rs 9.46 Cr |

## Regenerate

```bash
pip install openpyxl reportlab
python scripts/build_xlsx.py
# Recalculate so cached values are stored (needed by the PDF script):
soffice --headless --convert-to xlsx --outdir output/recalc output/Plotting_Sales_Plan.xlsx
mv output/recalc/Plotting_Sales_Plan.xlsx output/ && rmdir output/recalc
python scripts/build_pdf.py
```
