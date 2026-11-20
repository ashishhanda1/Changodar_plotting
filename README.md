# Changodar Plotting Project: Sales Plan

Plot sales plan for 25,000 sq yd of land divided into 150 sq yd plots (166 plots).

## Files

- `output/Plotting_Sales_Plan.xlsx`: workbook with live formulas. Yellow cells are inputs.
  Sheets: Inputs, Commission Ladder, Plot & Revenue, Salesperson Tracker (30 people), Payment Plans (DP vs EMI), Profit Summary.
- `output/Plotting_Sales_Plan_Summary.pdf`: 3-page summary of the plan.
- `output/Changodar_Plots_Marketing.pptx`: 10-slide marketing deck for buyers and sales partners (layout, pricing, DP vs EMI, ladder income, how to book). Contact on the last slide: Harman Ram, +91 7777999896.
- `output/R2_Sales_Plan.xlsx`: R2 plan, salespeople earn a fixed INR per sq yd (1 plot Rs 1,700; 2 to 4 plots Rs 1,800; 5+ plots Rs 2,000, applied to all plots). 30-person tracker, R1 vs R2 comparison, developer margin.
- `output/R2_Sales_Partner_Plan.pptx`: 9-slide deck for sales partners only, explaining R2.
- Hindi and Gujarati versions of both decks: `*_Hindi.pptx`, `*_Gujarati.pptx`. They use the Nirmala UI font (included with Windows and Office).
- `assets/<lang>/`: illustrations per language (en, hi, gu).
- `scripts/`: `build_xlsx.py`, `build_pdf.py`, `build_images.py`, `build_deck.py` regenerate everything.

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
pip install openpyxl reportlab matplotlib python-pptx
python scripts/build_xlsx.py
# Recalculate so cached values are stored (needed by the PDF script):
soffice --headless --convert-to xlsx --outdir output/recalc output/Plotting_Sales_Plan.xlsx
mv output/recalc/Plotting_Sales_Plan.xlsx output/ && rmdir output/recalc
python scripts/build_pdf.py
for l in en hi gu; do python scripts/build_images.py $l; python scripts/build_deck.py $l; python scripts/build_r2_deck.py $l; done
python scripts/build_r2_xlsx.py   # then recalculate with soffice as above
```

Translations live in `scripts/strings_hi.py` and `scripts/strings_gu.py` (English text is the key). Images need Noto Sans Devanagari / Gujarati fonts (`apt install fonts-noto-core`) and Pillow with raqm.
