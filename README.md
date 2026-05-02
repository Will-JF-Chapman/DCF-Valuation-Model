# DCF Valuation Model

Python tool that builds a fully-formulated Discounted Cash Flow model for any public company and exports it as an Excel workbook. Pulls live financials from Yahoo Finance, projects 5 years of free cash flow, computes WACC from the company's actual capital structure, and runs a sensitivity analysis across WACC and terminal growth. Output is a 5-sheet workbook with linked formulas. Edit any input and the whole model recalculates.

## Stack

Python 3.10+, yfinance, pandas, NumPy, openpyxl

## Setup

bash
pip install yfinance pandas numpy openpyxl

## Usage

Note that the default ticker is AAPL

"""
bash
python dcf_model.py
"""

To pass a different ticker as an argument:

"""
bash
python dcf_model.py MSFT
python dcf_model.py NVDA
"""

Output file `DCF_Valuation.xlsx` is created in the working directory. Open in Excel and the formulas recalculate automatically.

## What the model does

- Pulls 4-5 years of historical financials from Yahoo Finance
- Projects free cash flow forward 5 years using historical defaults for growth, margins, and reinvestment ratios
- Computes WACC: cost of equity from CAPM, after-tax cost of debt from interest expense ÷ total debt, weighted by the actual capital structure
- Discounts projected FCFs and adds a Gordon Growth terminal value
- Subtracts net debt, divides by share count → implied intrinsic share price
- Compares to current market price and outputs a BUY / ACCUMULATE / HOLD / REDUCE / SELL recommendation

## Limitations

- 5-year explicit forecast means terminal value typically dominates enterprise value. The valuation is highly sensitive to terminal growth — flex it on the Sensitivity Table sheet to see how much.
- Single-segment view. Conglomerates with structurally different business lines would benefit from a sum-of-the-parts approach.
- Static capital structure across the projection period — no modeling of deleveraging or buybacks.
- yfinance data quality is best for liquid US large-caps; smaller or international names may have missing fields.

## Files

DCF_Valuation_Model/
├── dcf_model.py
├── DCF_Valuation.xlsx
└── README.md
