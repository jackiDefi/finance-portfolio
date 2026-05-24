"""
Build telecom-model.xlsx — a 16-sheet, formula-driven three-statement
financial model for a stylised integrated telco.

Default values approximate Saudi Telecom Company (Tadawul: 7010);
balance-sheet starting positions are illustrative — they are NOT
actual STC FY2024 figures.

Run from this directory:
    python3 build_excel.py

Convention:
    - Inputs:   blue font   (0563C1)  on the Assumptions sheet
    - Formulas: black font  (000000)
    - Outputs:  bold font   (000000)
    - Headers:  white bold on navy fill
    - Bands:    light-cream fill on totals / key rollup rows
    - Checks:   green if pass, red if fail

Workbook layout (16 sheets):
    1.  Cover          —  title, model map, colour legend, instructions
    2.  Assumptions    —  every driver / starting balance in one place
    3.  Drivers        —  revenue and subscriber build-up by segment
    4.  IS             —  full income statement, Y1-Y5
    5.  BS             —  balance sheet, Y0 opening + Y1-Y5
    6.  CFS            —  cash flow statement, Y1-Y5
    7.  Equity         —  equity rollforward (NI, dividends, OCI)
    8.  Debt           —  debt schedule (gross debt, interest, movements)
    9.  PPE            —  PP&E rollforward (capex, D&A)
    10. WC             —  working capital schedule (AR, inventory, AP)
    11. Tax            —  tax expense schedule
    12. Budget         —  Y1 monthly budget with quarterly subtotals
    13. Scenarios      —  base / upside / downside switcher
    14. Sensitivity    —  two-way data tables on WACC × g and margin × growth
    15. Valuation      —  DCF + multiple cross-check + per-share
    16. Checks         —  BS balance, CFS reconciliation, cash positivity
"""

from datetime import date
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName


# =============================================================================
# 1. STYLE PRIMITIVES
# =============================================================================

BLUE     = "0563C1"
NAVY     = "0A1828"
COPPER   = "B87333"
LIGHT_BG = "F1EFE6"
PASS_BG  = "DDF0DC"
FAIL_BG  = "F8D7D5"
PASS_FG  = "2C7A3E"
FAIL_FG  = "B33231"

input_font    = Font(name="Calibri", size=11, color=BLUE,    bold=False)
formula_font  = Font(name="Calibri", size=11, color="000000", bold=False)
output_font   = Font(name="Calibri", size=11, color="000000", bold=True)
header_font   = Font(name="Calibri", size=11, color="FFFFFF", bold=True)
title_font    = Font(name="Calibri", size=18, color=NAVY,    bold=True)
sub_font      = Font(name="Calibri", size=10, color="5A6573", italic=True)
section_font  = Font(name="Calibri", size=10, color=COPPER,  bold=True)
note_font     = Font(name="Calibri", size= 9, color="5A6573")
pass_font     = Font(name="Calibri", size=11, color=PASS_FG, bold=True)
fail_font     = Font(name="Calibri", size=11, color=FAIL_FG, bold=True)

header_fill   = PatternFill("solid", fgColor=NAVY)
band_fill     = PatternFill("solid", fgColor=LIGHT_BG)
pass_fill     = PatternFill("solid", fgColor=PASS_BG)
fail_fill     = PatternFill("solid", fgColor=FAIL_BG)

right  = Alignment(horizontal="right")
left   = Alignment(horizontal="left")
center = Alignment(horizontal="center")

thin   = Side(border_style="thin", color="D4D2C8")
medium = Side(border_style="medium", color="0A1828")
bottom_border    = Border(bottom=thin)
top_border       = Border(top=thin)
top_bottom       = Border(top=thin, bottom=thin)
double_underline = Border(top=thin, bottom=medium)

NUM_BN  = '#,##0.0;[Red]-#,##0.0'
NUM_PCT = '0.0%;[Red]-0.0%'
NUM_X   = '0.00"x"'
NUM_INT = '#,##0'
NUM_DAYS = '#,##0" d"'


# =============================================================================
# 2. DEFAULT VALUES  (mirror the JS DEFAULTS where overlapping)
# =============================================================================

D = {
    # ---- Operating drivers ----
    "rev0":              75.0,     # SAR bn, service revenue Y0
    "growth":            0.06,     # annual %
    "ebitda_target":     0.36,     # display only; actual EBITDA is sum-of-costs
    "da_pct":            0.16,     # D&A / revenue
    "capex_pct":         0.18,     # capex / revenue

    # ---- Cost structure (must sum to 1 - ebitda_target = 0.64) ----
    "cogs_pct":          0.22,     # COGS / interconnect / content
    "network_opex_pct":  0.14,     # network operating costs
    "employee_pct":      0.12,     # employee costs
    "cust_acq_pct":      0.05,     # commercial / customer acquisition
    "other_opex_pct":    0.11,     # G&A and other

    # ---- Other P&L items ----
    "finance_income_pct": 0.005,   # finance income / revenue (treasury yield etc.)

    # ---- Financial assumptions ----
    "tax":               0.15,     # tax / zakat
    "wacc":              0.09,
    "g":                 0.03,     # terminal growth
    "kd":                0.05,     # cost of debt on net debt (interest)
    "div_payout":        0.50,     # payout of net income

    # ---- Working-capital ratios ----
    "dso":               75,       # days sales outstanding
    "dio":               30,       # days inventory outstanding
    "dpo":               90,       # days payables outstanding

    # ---- Balance sheet Y0 (illustrative — NOT actual STC figures) ----
    # Assets
    "cash_y0":           10.0,
    "ar_y0":             15.4,     # = 75 * 75/365
    "inv_y0":             2.5,
    "other_ca_y0":        5.0,
    "ppe_y0":            80.0,
    "goodwill_y0":       25.0,
    "intangibles_y0":    15.0,
    "investments_y0":     8.0,
    "dta_y0":             2.0,
    # Liabilities
    "ap_y0":              7.5,
    "other_cl_y0":       10.0,
    "st_debt_y0":         5.0,
    "lt_debt_y0":        20.0,
    "lease_liab_y0":     18.0,
    "provisions_y0":      5.0,
    "dtl_y0":             3.0,
    "other_ncl_y0":       5.0,
    # Equity (Retained earnings is a plug so total balances)
    "share_capital":      2.0,
    "oci_reserve_y0":     0.0,
    "minority_y0":        7.9,
    # Retained earnings plug: TotalAssets - TotalLiab - SC - OCI - Min
    # = 162.9 - 73.5 - 2 - 0 - 7.9 = 79.5

    # ---- Capital structure ----
    "shares":          5000,       # millions

    # ---- Scenario settings (1 = downside, 2 = base, 3 = upside) ----
    "scen_default":       2,
    "scen_growth":       [0.02, 0.06, 0.09],
    "scen_margin":       [0.33, 0.36, 0.38],
    "scen_capex":        [0.20, 0.18, 0.16],

    # ---- Budget seasonality (Q1, Q2, Q3, Q4 monthly weights sum to 1) ----
    # Telecom revenue is fairly flat with slight Q4 uplift from device sales.
    "month_weights": [
        0.0810, 0.0780, 0.0820,    # Q1 = 24.10%
        0.0830, 0.0820, 0.0850,    # Q2 = 25.00%
        0.0820, 0.0830, 0.0840,    # Q3 = 24.90%
        0.0860, 0.0870, 0.0870,    # Q4 = 26.00%
    ],
}

# Derived: starting retained earnings to balance
TOTAL_ASSETS_Y0 = (D["cash_y0"] + D["ar_y0"] + D["inv_y0"] + D["other_ca_y0"]
                   + D["ppe_y0"] + D["goodwill_y0"] + D["intangibles_y0"]
                   + D["investments_y0"] + D["dta_y0"])
TOTAL_LIAB_Y0   = (D["ap_y0"] + D["other_cl_y0"] + D["st_debt_y0"] + D["lt_debt_y0"]
                   + D["lease_liab_y0"] + D["provisions_y0"] + D["dtl_y0"]
                   + D["other_ncl_y0"])
RE_Y0_PLUG = TOTAL_ASSETS_Y0 - TOTAL_LIAB_Y0 - D["share_capital"] - D["oci_reserve_y0"] - D["minority_y0"]
D["re_y0"] = round(RE_Y0_PLUG, 2)


# =============================================================================
# 3. SHEET NAMES
# =============================================================================

SH_COVER  = "Cover"
SH_ASSUM  = "Assumptions"
SH_DRIV   = "Drivers"
SH_IS     = "IS"
SH_BS     = "BS"
SH_CFS    = "CFS"
SH_EQ     = "Equity"
SH_DEBT   = "Debt"
SH_PPE    = "PPE"
SH_WC     = "WC"
SH_TAX    = "Tax"
SH_BUD    = "Budget"
SH_SCEN   = "Scenarios"
SH_SENS   = "Sensitivity"
SH_VAL    = "Valuation"
SH_CHK    = "Checks"


# =============================================================================
# 4. CELL REGISTRY  (key -> "Sheet!$Col$Row")
# =============================================================================

COORD = {}

def reg(key, sheet, col, row):
    """Register an absolute cross-sheet reference."""
    COORD[key] = f"{sheet}!${col}${row}"
    return COORD[key]

def C(key):
    return COORD[key]

def yc(t):
    """Year column letter. t=0 => B (Y0); t=1..5 => C..G."""
    return get_column_letter(2 + t)


# =============================================================================
# 5. HELPERS
# =============================================================================

def section_header(ws, row, text):
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = section_font

def year_header(ws, row, start_t=0, end_t=5, label_y0="Y0 (base)"):
    """Write Y0..Y5 (or subset) header band in row `row`."""
    ws.cell(row=row, column=1).fill = header_fill
    for t in range(start_t, end_t + 1):
        col = 2 + t
        label = label_y0 if t == 0 else f"Y{t}"
        c = ws.cell(row=row, column=col, value=label)
        c.font = header_font
        c.fill = header_fill
        c.alignment = right

def write_label(ws, row, label, *, bold=False, banded=False, indent=0):
    cell = ws.cell(row=row, column=1, value=("  " * indent) + label)
    cell.font = output_font if bold else formula_font
    if banded:
        cell.fill = band_fill
    return cell

def write_input(ws, row, col, value, fmt=NUM_BN, note=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = input_font
    c.number_format = fmt
    c.alignment = right
    if note:
        n = ws.cell(row=row, column=col + 1, value=note)
        n.font = sub_font
    return c

def write_formula(ws, row, col, formula, fmt=NUM_BN, *, bold=False, banded=False):
    c = ws.cell(row=row, column=col, value=formula)
    c.font = output_font if bold else formula_font
    c.number_format = fmt
    c.alignment = right
    if banded:
        c.fill = band_fill
    return c

def fill_row(ws, row, fmt=NUM_BN, start_t=1, end_t=5, formula_for_t=None,
             bold=False, banded=False):
    """Write Y1..Y5 (or subset) formulas using formula_for_t(t) -> str."""
    for t in range(start_t, end_t + 1):
        formula = formula_for_t(t)
        write_formula(ws, row, 2 + t, formula, fmt, bold=bold, banded=banded)


# =============================================================================
# 6. COVER SHEET
# =============================================================================

def build_cover(ws):
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 70

    ws["B2"] = "Telecom & IT — Operating Model"
    ws["B2"].font = title_font
    ws["B3"] = f"Author: Togrul Mirzayev   •   Generated: {date.today().isoformat()}"
    ws["B3"].font = sub_font

    ws["B5"] = "PORTFOLIO"
    ws["B5"].font = section_font
    ws["B6"] = "https://jackidefi.github.io/finance-portfolio/"
    ws["B6"].font = Font(name="Calibri", size=11, color="000000", underline="single")

    ws["B8"] = "MODEL MAP"
    ws["B8"].font = section_font
    sheets = [
        ("Assumptions",  "Every driver and Y0 starting balance — blue cells, change freely."),
        ("Drivers",      "Subscriber and ARPU buildup; service revenue ties back to Assumptions."),
        ("IS",           "Income Statement, fully formula-driven, Y1-Y5."),
        ("BS",           "Balance Sheet, Y0 opening + Y1-Y5 forecast. Must balance."),
        ("CFS",          "Cash Flow Statement, indirect method. Ties to change in cash on BS."),
        ("Equity",       "Equity rollforward — share capital, retained earnings, OCI, minority."),
        ("Debt",         "Debt schedule — gross debt, interest, period movements."),
        ("PPE",          "PP&E rollforward — capex in, D&A out, closing balance to BS."),
        ("WC",           "Working capital schedule — AR, inventory, AP from days ratios."),
        ("Tax",          "Tax expense schedule — current tax, effective rate."),
        ("Budget",       "Y1 broken into 12 months with quarterly subtotals (planning view)."),
        ("Scenarios",    "Downside / Base / Upside switcher; flexes growth, margin, capex."),
        ("Sensitivity",  "Two-way tables: WACC × terminal growth, margin × revenue growth."),
        ("Valuation",    "DCF (Gordon-growth terminal), EV/EBITDA cross-check, per-share."),
        ("Checks",       "BS balance, CFS reconciliation, cash positivity, sources = uses."),
    ]
    for i, (name, desc) in enumerate(sheets):
        ws.cell(row=10 + i, column=2, value=name).font = output_font
        ws.cell(row=10 + i, column=3, value=desc).font = formula_font

    r = 10 + len(sheets) + 2
    ws.cell(row=r, column=2, value="COLOUR LEGEND").font = section_font
    ws.cell(row=r+1, column=2, value="Blue").font = input_font
    ws.cell(row=r+1, column=3, value="Hard-coded input — change freely.").font = formula_font
    ws.cell(row=r+2, column=2, value="Black").font = formula_font
    ws.cell(row=r+2, column=3, value="Formula referencing other cells.").font = formula_font
    ws.cell(row=r+3, column=2, value="Bold").font = output_font
    ws.cell(row=r+3, column=3, value="Subtotal or key model output.").font = formula_font
    ws.cell(row=r+4, column=2, value="Green / red fills").font = formula_font
    ws.cell(row=r+4, column=3, value="Checks sheet — pass / fail status.").font = formula_font

    r += 6
    ws.cell(row=r, column=2, value="DATA DISCLAIMER").font = section_font
    disclaimer = (
        "Default values are calibrated to roughly approximate STC (Tadawul: 7010) at the "
        "consolidated level, but starting balance-sheet line items are illustrative — "
        "they are NOT extracted from actual STC filings. The model demonstrates structure "
        "and methodology. For a publication-grade analysis, replace all blue input cells "
        "with figures verified against STC's audited financial statements."
    )
    cell = ws.cell(row=r+1, column=2, value=disclaimer)
    cell.font = note_font
    cell.alignment = Alignment(horizontal="left", wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r+1, start_column=2, end_row=r+5, end_column=3)
    ws.row_dimensions[r+1].height = 15
    for rr in range(r+1, r+6):
        ws.row_dimensions[rr].height = 15

    r += 7
    ws.cell(row=r, column=2, value="MODEL CONVENTIONS").font = section_font
    conventions = [
        "All figures in SAR billion unless noted (subscribers in millions, ARPU in SAR / month).",
        "Forecast horizon: 5 explicit years (Y1-Y5). Y0 is the base / opening period.",
        "Interest expense uses cost-of-debt × opening total debt to avoid circular references.",
        "Working capital is built from DSO / DIO / DPO ratios applied to revenue and COGS.",
        "Depreciation is driven by D&A as % of revenue; PP&E rolls forward as opening + capex − D&A.",
        "Tax is a single effective rate applied to pre-tax income; deferred tax held flat.",
        "Dividends = payout ratio × net income; equity rolls forward through the Equity sheet.",
        "Terminal value via Gordon growth: FCFF_Y5 × (1+g) ÷ (WACC − g).",
        "BS plugs to cash from the CFS — if every BS movement is correctly mirrored in CFS, the BS balances.",
    ]
    for i, conv in enumerate(conventions):
        c = ws.cell(row=r+1+i, column=2, value=f"•  {conv}")
        c.font = formula_font
        ws.merge_cells(start_row=r+1+i, start_column=2, end_row=r+1+i, end_column=3)


# =============================================================================
# 7. ASSUMPTIONS SHEET
# =============================================================================

def build_assumptions(ws):
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 60

    ws["A1"] = "ASSUMPTIONS"
    ws["A1"].font = section_font
    ws["A2"] = "All inputs to the model live on this sheet. Change blue cells freely."
    ws["A2"].font = sub_font

    rows = [
        # (row, key, label, value, fmt, note)
        # ---- Operating drivers ----
        ( 4, None,              "Operating drivers",                  None,                NUM_BN,  None),
        ( 5, "rev0",            "Service revenue Y0 (SAR bn)",        D["rev0"],           NUM_BN,  "Year 0 base; everything scales from here."),
        ( 6, "growth",          "Revenue growth (annual)",            D["growth"],         NUM_PCT, "Applied to every forecast year (Scenarios sheet can override)."),
        ( 7, "ebitda_target",   "EBITDA margin (target, info only)",  D["ebitda_target"],  NUM_PCT, "Reference value. Actual EBITDA is computed bottom-up from cost lines below."),
        ( 8, "da_pct",          "D&A / revenue",                      D["da_pct"],         NUM_PCT, "Drives PP&E rollforward and depreciation expense."),
        ( 9, "capex_pct",       "Capex intensity",                    D["capex_pct"],      NUM_PCT, "Capex ÷ revenue, fed into PP&E and CFS."),

        # ---- Cost split ----
        (11, None,              "Cost structure (% of revenue)",      None,                NUM_BN,  None),
        (12, "cogs_pct",        "COGS / interconnect / content",      D["cogs_pct"],       NUM_PCT, "Variable; partly indexed to revenue."),
        (13, "network_opex_pct","Network opex",                       D["network_opex_pct"],NUM_PCT,"Site lease, energy, maintenance — mostly fixed."),
        (14, "employee_pct",    "Employee costs",                     D["employee_pct"],   NUM_PCT, "Quasi-fixed."),
        (15, "cust_acq_pct",    "Customer acquisition / commercial",  D["cust_acq_pct"],   NUM_PCT, "Subsidies, dealer commissions."),
        (16, "other_opex_pct",  "Other opex / G&A",                   D["other_opex_pct"], NUM_PCT, "Catch-all; sum of all five lines + EBITDA margin should = 100%."),
        (17, "finance_income_pct","Finance income / revenue",         D["finance_income_pct"],NUM_PCT,"Treasury yield on cash and short-term investments."),

        # ---- Financial ----
        (19, None,              "Financial",                          None,                NUM_BN,  None),
        (20, "tax",             "Tax / zakat rate",                   D["tax"],            NUM_PCT, "Effective rate applied to pre-tax income."),
        (21, "wacc",            "WACC",                               D["wacc"],           NUM_PCT, "Discount rate for DCF."),
        (22, "g",               "Terminal growth (g)",                D["g"],              NUM_PCT, "Long-run nominal growth; must be < WACC."),
        (23, "kd",              "Cost of debt (interest)",            D["kd"],             NUM_PCT, "Applied to opening total debt for period interest."),
        (24, "div_payout",      "Dividend payout (% of NI)",          D["div_payout"],     NUM_PCT, "Used to compute dividends in equity rollforward & CFS."),

        # ---- Working capital ratios ----
        (26, None,              "Working capital (days)",             None,                NUM_BN,  None),
        (27, "dso",             "Days sales outstanding (DSO)",       D["dso"],            NUM_DAYS,"Drives accounts receivable on BS."),
        (28, "dio",             "Days inventory outstanding (DIO)",   D["dio"],            NUM_DAYS,"Drives inventory on BS."),
        (29, "dpo",             "Days payables outstanding (DPO)",    D["dpo"],            NUM_DAYS,"Drives accounts payable on BS."),

        # ---- Capital structure ----
        (31, None,              "Capital structure",                  None,                NUM_BN,  None),
        (32, "shares",          "Shares outstanding (millions)",      D["shares"],         NUM_INT, "Used for per-share value."),

        # ---- Balance sheet Y0 ----
        (34, None,              "Balance sheet Y0 — Assets",          None,                NUM_BN,  None),
        (35, "cash_y0",         "Cash & equivalents",                 D["cash_y0"],        NUM_BN,  "Opening cash; plugs through CFS thereafter."),
        (36, "ar_y0",           "Accounts receivable",                D["ar_y0"],          NUM_BN,  "Calibrated to DSO × revenue."),
        (37, "inv_y0",          "Inventory",                          D["inv_y0"],         NUM_BN,  None),
        (38, "other_ca_y0",     "Other current assets",               D["other_ca_y0"],    NUM_BN,  "Held flat in forecast."),
        (39, "ppe_y0",          "Property, plant & equipment",        D["ppe_y0"],         NUM_BN,  "Net book value; rolls forward via PPE sheet."),
        (40, "goodwill_y0",     "Goodwill",                           D["goodwill_y0"],    NUM_BN,  "Held flat (no impairment in default case)."),
        (41, "intangibles_y0",  "Intangibles (incl. spectrum)",       D["intangibles_y0"], NUM_BN,  "Held flat."),
        (42, "investments_y0",  "Investments / associates",           D["investments_y0"], NUM_BN,  None),
        (43, "dta_y0",          "Deferred tax assets",                D["dta_y0"],         NUM_BN,  "Held flat."),

        (45, None,              "Balance sheet Y0 — Liabilities",     None,                NUM_BN,  None),
        (46, "ap_y0",           "Accounts payable",                   D["ap_y0"],          NUM_BN,  "Calibrated to DPO × COGS."),
        (47, "other_cl_y0",     "Other current liabilities",          D["other_cl_y0"],    NUM_BN,  None),
        (48, "st_debt_y0",      "Short-term debt",                    D["st_debt_y0"],     NUM_BN,  None),
        (49, "lt_debt_y0",      "Long-term debt",                     D["lt_debt_y0"],     NUM_BN,  None),
        (50, "lease_liab_y0",   "Lease liabilities (IFRS 16)",        D["lease_liab_y0"],  NUM_BN,  "Held flat in forecast (simplification)."),
        (51, "provisions_y0",   "Provisions",                         D["provisions_y0"],  NUM_BN,  None),
        (52, "dtl_y0",          "Deferred tax liabilities",           D["dtl_y0"],         NUM_BN,  "Held flat."),
        (53, "other_ncl_y0",    "Other non-current liabilities",      D["other_ncl_y0"],   NUM_BN,  None),

        (55, None,              "Balance sheet Y0 — Equity",          None,                NUM_BN,  None),
        (56, "share_capital",   "Share capital",                      D["share_capital"],  NUM_BN,  "Held flat (no issuance / buyback)."),
        (57, "re_y0",           "Retained earnings",                  D["re_y0"],          NUM_BN,  "Computed plug so total balance sheet balances at Y0."),
        (58, "oci_reserve_y0",  "OCI reserve",                        D["oci_reserve_y0"], NUM_BN,  "Default: nil movement."),
        (59, "minority_y0",     "Minority interest",                  D["minority_y0"],    NUM_BN,  "Held flat."),

        # ---- Scenarios ----
        (61, None,              "Scenario selector",                  None,                NUM_BN,  None),
        (62, "scen_default",    "Active scenario (1=Down, 2=Base, 3=Up)", D["scen_default"], NUM_INT, "Drives the Scenarios sheet."),
    ]

    for row, key, label, value, fmt, note in rows:
        if value is None:
            # Section header row
            section_header(ws, row, label)
        else:
            ws.cell(row=row, column=1, value=label).font = formula_font
            cell = write_input(ws, row, 2, value, fmt, note)
            if key:
                reg(f"a.{key}", SH_ASSUM, "B", row)


# =============================================================================
# 8. DRIVERS SHEET (revenue buildup by segment)
# =============================================================================

def build_drivers(ws):
    ws.column_dimensions["A"].width = 34
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = 13

    ws["A1"] = "REVENUE & SUBSCRIBER DRIVERS"
    ws["A1"].font = section_font
    ws["A2"] = "Illustrative segment buildup. Segment revenues scale to match the total service revenue on Assumptions."
    ws["A2"].font = sub_font

    year_header(ws, 4, 0, 5)

    # --- Subscribers (millions) ---
    r = 6
    section_header(ws, r, "Subscribers (millions)"); r += 1

    seg_subs_y0 = {
        "Mobile consumer": 35.0,
        "Fixed broadband":  1.5,
        "B2B connections":  0.5,
    }
    seg_arpu_y0 = {
        "Mobile consumer":  80.0,    # SAR / month
        "Fixed broadband": 400.0,
        "B2B connections":3000.0,
    }
    # Sub growth assumptions (annual %)
    seg_sub_growth = {"Mobile consumer": 0.02, "Fixed broadband": 0.04, "B2B connections": 0.06}
    seg_arpu_growth = {"Mobile consumer": 0.03, "Fixed broadband": 0.02, "B2B connections": 0.04}

    sub_rows = {}
    for name, sub_y0 in seg_subs_y0.items():
        write_label(ws, r, name)
        write_input(ws, r, 2, sub_y0, NUM_BN)  # blue, since input
        g_pct = seg_sub_growth[name]
        for t in range(1, 6):
            prev = ws.cell(row=r, column=2 + t - 1).coordinate
            write_formula(ws, r, 2 + t, f"={prev}*(1+{g_pct})", NUM_BN)
        sub_rows[name] = r
        r += 1

    # ---- ARPU (SAR/month) ----
    r += 1
    section_header(ws, r, "ARPU (SAR / month)"); r += 1
    arpu_rows = {}
    for name, a in seg_arpu_y0.items():
        write_label(ws, r, name)
        write_input(ws, r, 2, a, '#,##0')
        g_pct = seg_arpu_growth[name]
        for t in range(1, 6):
            prev = ws.cell(row=r, column=2 + t - 1).coordinate
            write_formula(ws, r, 2 + t, f"={prev}*(1+{g_pct})", '#,##0')
        arpu_rows[name] = r
        r += 1

    # ---- Implied revenue (SAR bn) ----
    r += 1
    section_header(ws, r, "Implied revenue (SAR bn)"); r += 1
    imp_rows = {}
    for name in seg_subs_y0.keys():
        write_label(ws, r, name)
        for t in range(0, 6):
            sub_cell = ws.cell(row=sub_rows[name], column=2 + t).coordinate
            arpu_cell = ws.cell(row=arpu_rows[name], column=2 + t).coordinate
            # subs (mm) * ARPU (SAR/mo) * 12 / 1000 = SAR bn
            write_formula(ws, r, 2 + t, f"={sub_cell}*{arpu_cell}*12/1000", NUM_BN)
        imp_rows[name] = r
        r += 1

    # Sum of subs-driven segments
    subs_total_row = r
    write_label(ws, r, "Subs-driven revenue", bold=True, banded=True)
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        cells = [f"{col}{imp_rows[n]}" for n in seg_subs_y0]
        write_formula(ws, r, 2 + t, "=" + "+".join(cells), NUM_BN, bold=True, banded=True)
    r += 2

    # ---- Other revenue lines (ICT/cloud + wholesale + equipment) ----
    section_header(ws, r, "Other revenue (residual)"); r += 1
    write_label(ws, r, "ICT / cloud / wholesale (plug)")
    # Plug so total = service revenue from Assumptions
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        if t == 0:
            target = C("a.rev0")
        else:
            target = f"{C('a.rev0')}*(1+{C('a.growth')})^{t}"
        plug = f"={target}-{col}{subs_total_row}"
        write_formula(ws, r, 2 + t, plug, NUM_BN)
    other_row = r
    r += 1

    # ---- TOTAL service revenue (ties to Assumptions × growth) ----
    r += 1
    write_label(ws, r, "TOTAL SERVICE REVENUE", bold=True, banded=True)
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{subs_total_row}+{col}{other_row}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
        reg(f"dr.rev_t{t}", SH_DRIV, col, r)
    total_row = r


# =============================================================================
# 9. INCOME STATEMENT
# =============================================================================

# We track IS row positions in a dict so other sheets can reference precisely.
IS_ROWS = {}

def build_is(ws):
    ws.column_dimensions["A"].width = 32
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = 13

    ws["A1"] = "INCOME STATEMENT"
    ws["A1"].font = section_font
    ws["A2"] = "All figures in SAR billion. Five-year explicit forecast."
    ws["A2"].font = sub_font

    # Header row: just Y1..Y5
    ws.cell(row=4, column=1).fill = header_fill
    for t in range(1, 6):
        c = ws.cell(row=4, column=2 + t, value=f"Y{t}")
        c.font = header_font; c.fill = header_fill; c.alignment = right
    # Leave column B blank (Y0 not shown on IS; would be historical)

    r = 5

    # Service revenue (from Drivers)
    write_label(ws, r, "Service revenue", bold=True)
    fill_row(ws, r, formula_for_t=lambda t: f"={C(f'dr.rev_t{t}')}", bold=True)
    IS_ROWS["revenue"] = r
    r += 2

    # Costs
    write_label(ws, r, "Cost lines (% of revenue)"); r += 1

    cost_defs = [
        ("COGS / interconnect / content",  "cogs_pct",         "cogs"),
        ("Network operating costs",        "network_opex_pct", "network"),
        ("Employee costs",                 "employee_pct",     "emp"),
        ("Customer acquisition / commercial","cust_acq_pct",   "cust_acq"),
        ("Other opex / G&A",               "other_opex_pct",   "other"),
    ]
    for label, akey, ikey in cost_defs:
        write_label(ws, r, label, indent=1)
        fill_row(ws, r, formula_for_t=lambda t, k=akey: f"=-{C('a.' + k)}*{ws.cell(row=IS_ROWS['revenue'], column=2 + t).coordinate}")
        IS_ROWS[ikey] = r
        r += 1

    # EBITDA
    r += 1
    write_label(ws, r, "EBITDA", bold=True, banded=True)
    cost_rows = [IS_ROWS[k] for _, _, k in cost_defs]
    def ebitda_formula(t):
        col = get_column_letter(2 + t)
        terms = [f"{col}{IS_ROWS['revenue']}"] + [f"{col}{cr}" for cr in cost_rows]
        return "=" + "+".join(terms)
    fill_row(ws, r, formula_for_t=ebitda_formula, bold=True, banded=True)
    IS_ROWS["ebitda"] = r
    r += 2

    # D&A
    write_label(ws, r, "Depreciation & amortisation")
    fill_row(ws, r, formula_for_t=lambda t: f"=-{C('a.da_pct')}*{ws.cell(row=IS_ROWS['revenue'], column=2 + t).coordinate}")
    IS_ROWS["da"] = r
    r += 1

    # EBIT
    r += 1
    write_label(ws, r, "EBIT (operating profit)", bold=True, banded=True)
    fill_row(ws, r, formula_for_t=lambda t: f"={get_column_letter(2+t)}{IS_ROWS['ebitda']}+{get_column_letter(2+t)}{IS_ROWS['da']}", bold=True, banded=True)
    IS_ROWS["ebit"] = r
    r += 2

    # Finance income (% of revenue)
    write_label(ws, r, "Finance income")
    fill_row(ws, r, formula_for_t=lambda t: f"={C('a.finance_income_pct')}*{get_column_letter(2+t)}{IS_ROWS['revenue']}")
    IS_ROWS["fin_inc"] = r
    r += 1

    # Interest expense (negative): from Debt sheet
    write_label(ws, r, "Interest expense")
    fill_row(ws, r, formula_for_t=lambda t: f"=-{SH_DEBT}!{get_column_letter(2+t)}{DEBT_ROWS_PLACEHOLDER}")
    IS_ROWS["interest"] = r
    r += 1

    # Pre-tax income
    r += 1
    write_label(ws, r, "Pre-tax income", bold=True)
    fill_row(ws, r, formula_for_t=lambda t: f"={get_column_letter(2+t)}{IS_ROWS['ebit']}+{get_column_letter(2+t)}{IS_ROWS['fin_inc']}+{get_column_letter(2+t)}{IS_ROWS['interest']}", bold=True)
    IS_ROWS["pretax"] = r
    r += 1

    # Tax (from Tax schedule)
    write_label(ws, r, "Income tax")
    fill_row(ws, r, formula_for_t=lambda t: f"=-{SH_TAX}!{get_column_letter(2+t)}{TAX_ROWS_PLACEHOLDER}")
    IS_ROWS["tax"] = r
    r += 1

    # Net income
    r += 1
    write_label(ws, r, "NET INCOME", bold=True, banded=True)
    fill_row(ws, r, formula_for_t=lambda t: f"={get_column_letter(2+t)}{IS_ROWS['pretax']}+{get_column_letter(2+t)}{IS_ROWS['tax']}", bold=True, banded=True)
    IS_ROWS["ni"] = r
    r += 2

    # Memo: EBITDA margin
    section_header(ws, r, "MEMO"); r += 1
    write_label(ws, r, "EBITDA margin")
    fill_row(ws, r, fmt=NUM_PCT, formula_for_t=lambda t: f"={get_column_letter(2+t)}{IS_ROWS['ebitda']}/{get_column_letter(2+t)}{IS_ROWS['revenue']}")
    IS_ROWS["ebitda_margin"] = r
    r += 1
    write_label(ws, r, "EBIT margin")
    fill_row(ws, r, fmt=NUM_PCT, formula_for_t=lambda t: f"={get_column_letter(2+t)}{IS_ROWS['ebit']}/{get_column_letter(2+t)}{IS_ROWS['revenue']}")
    IS_ROWS["ebit_margin"] = r
    r += 1
    write_label(ws, r, "Net margin")
    fill_row(ws, r, fmt=NUM_PCT, formula_for_t=lambda t: f"={get_column_letter(2+t)}{IS_ROWS['ni']}/{get_column_letter(2+t)}{IS_ROWS['revenue']}")
    IS_ROWS["net_margin"] = r

    # Register cross-sheet references
    for key, row in IS_ROWS.items():
        for t in range(1, 6):
            col = get_column_letter(2 + t)
            reg(f"is.{key}_t{t}", SH_IS, col, row)


# Placeholders — these get patched after Debt/Tax sheets are built
DEBT_ROWS_PLACEHOLDER = 0
TAX_ROWS_PLACEHOLDER = 0
# (We solve this differently — by passing the row number to build_is via globals
# after Debt and Tax are constructed. Simpler approach: build IS LAST among the
# linked sheets. But IS feeds Equity which feeds BS, so order is: Debt -> PPE ->
# WC -> Tax -> IS -> Equity -> BS -> CFS. Let's restructure.)


# =============================================================================
# 10. DEBT SCHEDULE
# =============================================================================

DEBT_ROWS = {}

def build_debt(ws):
    ws.column_dimensions["A"].width = 32
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = 13

    ws["A1"] = "DEBT SCHEDULE"
    ws["A1"].font = section_font
    ws["A2"] = "Total debt = ST + LT + lease liabilities. Held flat unless overridden (simplification)."
    ws["A2"].font = sub_font

    year_header(ws, 4, 0, 5)

    r = 6

    # Opening total debt = Y0 ST + LT + lease
    write_label(ws, r, "Short-term debt")
    write_formula(ws, r, 2, f"={C('a.st_debt_y0')}", NUM_BN)  # Y0
    for t in range(1, 6):
        prev = ws.cell(row=r, column=2 + t - 1).coordinate
        write_formula(ws, r, 2 + t, f"={prev}", NUM_BN)  # held flat
    DEBT_ROWS["st"] = r; r += 1

    write_label(ws, r, "Long-term debt")
    write_formula(ws, r, 2, f"={C('a.lt_debt_y0')}", NUM_BN)
    for t in range(1, 6):
        prev = ws.cell(row=r, column=2 + t - 1).coordinate
        write_formula(ws, r, 2 + t, f"={prev}", NUM_BN)
    DEBT_ROWS["lt"] = r; r += 1

    write_label(ws, r, "Lease liabilities (IFRS 16)")
    write_formula(ws, r, 2, f"={C('a.lease_liab_y0')}", NUM_BN)
    for t in range(1, 6):
        prev = ws.cell(row=r, column=2 + t - 1).coordinate
        write_formula(ws, r, 2 + t, f"={prev}", NUM_BN)
    DEBT_ROWS["lease"] = r; r += 1

    # Total gross debt
    r += 1
    write_label(ws, r, "Gross debt (total)", bold=True, banded=True)
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{DEBT_ROWS['st']}+{col}{DEBT_ROWS['lt']}+{col}{DEBT_ROWS['lease']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
        if t == 0:
            reg(f"d.gross_y0", SH_DEBT, col, r)
        else:
            reg(f"d.gross_t{t}", SH_DEBT, col, r)
    DEBT_ROWS["gross"] = r; r += 2

    # Net debt = gross - cash (cash comes from BS; for now use Y0 directly and let
    # subsequent years pick up cash from BS sheet rows we'll write in build_bs)
    write_label(ws, r, "Less: cash (from BS)")
    # Y0 cash
    write_formula(ws, r, 2, f"=-{C('a.cash_y0')}", NUM_BN)
    # Y1..Y5: pull from BS
    for t in range(1, 6):
        # BS cash row to be defined; we'll reference by name later via key registry
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"=-BS!{col}{BS_PLACEHOLDER_CASH}", NUM_BN)
    DEBT_ROWS["cash_offset"] = r; r += 1

    r += 1
    write_label(ws, r, "Net debt", bold=True, banded=True)
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{DEBT_ROWS['gross']}+{col}{DEBT_ROWS['cash_offset']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
        if t == 0:
            reg("d.net_y0", SH_DEBT, col, r)
        else:
            reg(f"d.net_t{t}", SH_DEBT, col, r)
    DEBT_ROWS["net"] = r; r += 2

    # Interest expense (positive number)
    write_label(ws, r, "Interest expense (= kd × opening gross debt)")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        prev_col = get_column_letter(2 + t - 1)
        formula = f"={C('a.kd')}*{prev_col}{DEBT_ROWS['gross']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN)
        reg(f"d.interest_t{t}", SH_DEBT, col, r)
    DEBT_ROWS["interest"] = r


BS_PLACEHOLDER_CASH = 0   # patched below by build_bs writing the actual row


# =============================================================================
# 11. PPE SCHEDULE
# =============================================================================

PPE_ROWS = {}

def build_ppe(ws):
    ws.column_dimensions["A"].width = 32
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = 13

    ws["A1"] = "PP&E ROLLFORWARD"
    ws["A1"].font = section_font
    ws["A2"] = "Opening PP&E + capex − depreciation = closing PP&E. Net basis (no gross / accumulated split)."
    ws["A2"].font = sub_font

    year_header(ws, 4, 0, 5)

    r = 6
    # Opening PP&E
    write_label(ws, r, "Opening PP&E")
    write_formula(ws, r, 2, '""', '@')  # Y0 blank
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        prev_col = get_column_letter(2 + t - 1)
        if t == 1:
            write_formula(ws, r, 2 + t, f"={C('a.ppe_y0')}", NUM_BN)
        else:
            write_formula(ws, r, 2 + t, f"={prev_col}{r + 3}", NUM_BN)  # prev closing
    PPE_ROWS["opening"] = r; r += 1

    write_label(ws, r, "+ Capex")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={C('a.capex_pct')}*{C(f'is.revenue_t{t}')}"
        write_formula(ws, r, 2 + t, formula, NUM_BN)
        reg(f"p.capex_t{t}", SH_PPE, col, r)
    PPE_ROWS["capex"] = r; r += 1

    write_label(ws, r, "− Depreciation & amortisation")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"=-{C('a.da_pct')}*{C(f'is.revenue_t{t}')}"
        write_formula(ws, r, 2 + t, formula, NUM_BN)
        reg(f"p.da_t{t}", SH_PPE, col, r)
    PPE_ROWS["da"] = r; r += 1

    write_label(ws, r, "Closing PP&E", bold=True, banded=True)
    write_formula(ws, r, 2, f"={C('a.ppe_y0')}", NUM_BN, bold=True, banded=True)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{PPE_ROWS['opening']}+{col}{PPE_ROWS['capex']}+{col}{PPE_ROWS['da']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
        reg(f"p.closing_t{t}", SH_PPE, col, r)
    reg("p.closing_y0", SH_PPE, "B", r)
    PPE_ROWS["closing"] = r


# =============================================================================
# 12. WORKING CAPITAL SCHEDULE
# =============================================================================

WC_ROWS = {}

def build_wc(ws):
    ws.column_dimensions["A"].width = 32
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = 13

    ws["A1"] = "WORKING CAPITAL SCHEDULE"
    ws["A1"].font = section_font
    ws["A2"] = "AR = DSO × revenue / 365. Inventory = DIO × COGS / 365. AP = DPO × COGS / 365."
    ws["A2"].font = sub_font

    year_header(ws, 4, 0, 5)

    r = 6
    # Memo: revenue and COGS used for ratios
    section_header(ws, r, "Reference"); r += 1
    write_label(ws, r, "Revenue")
    write_formula(ws, r, 2, f"={C('a.rev0')}", NUM_BN)
    for t in range(1, 6):
        write_formula(ws, r, 2 + t, f"={C(f'is.revenue_t{t}')}", NUM_BN)
    rev_row = r; r += 1
    write_label(ws, r, "COGS (positive)")
    write_formula(ws, r, 2, f"={C('a.cogs_pct')}*{C('a.rev0')}", NUM_BN)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={C('a.cogs_pct')}*{C(f'is.revenue_t{t}')}", NUM_BN)
    cogs_row = r; r += 2

    # Receivables
    section_header(ws, r, "Working capital balances (closing)"); r += 1
    write_label(ws, r, "Accounts receivable")
    # Y0 = input
    write_formula(ws, r, 2, f"={C('a.ar_y0')}", NUM_BN)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={C('a.dso')}*{col}{rev_row}/365"
        write_formula(ws, r, 2 + t, formula, NUM_BN)
        reg(f"w.ar_t{t}", SH_WC, col, r)
    reg("w.ar_y0", SH_WC, "B", r)
    WC_ROWS["ar"] = r; r += 1

    write_label(ws, r, "Inventory")
    write_formula(ws, r, 2, f"={C('a.inv_y0')}", NUM_BN)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={C('a.dio')}*{col}{cogs_row}/365"
        write_formula(ws, r, 2 + t, formula, NUM_BN)
        reg(f"w.inv_t{t}", SH_WC, col, r)
    reg("w.inv_y0", SH_WC, "B", r)
    WC_ROWS["inv"] = r; r += 1

    write_label(ws, r, "Accounts payable")
    write_formula(ws, r, 2, f"={C('a.ap_y0')}", NUM_BN)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={C('a.dpo')}*{col}{cogs_row}/365"
        write_formula(ws, r, 2 + t, formula, NUM_BN)
        reg(f"w.ap_t{t}", SH_WC, col, r)
    reg("w.ap_y0", SH_WC, "B", r)
    WC_ROWS["ap"] = r; r += 2

    # ΔWC for CFS
    section_header(ws, r, "Working capital movements (to CFS)"); r += 1
    write_label(ws, r, "ΔReceivables (incr = use of cash)")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        prev = get_column_letter(2 + t - 1)
        formula = f"=-({col}{WC_ROWS['ar']}-{prev}{WC_ROWS['ar']})"
        write_formula(ws, r, 2 + t, formula, NUM_BN)
        reg(f"w.dar_t{t}", SH_WC, col, r)
    WC_ROWS["dar"] = r; r += 1

    write_label(ws, r, "ΔInventory (incr = use of cash)")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        prev = get_column_letter(2 + t - 1)
        formula = f"=-({col}{WC_ROWS['inv']}-{prev}{WC_ROWS['inv']})"
        write_formula(ws, r, 2 + t, formula, NUM_BN)
        reg(f"w.dinv_t{t}", SH_WC, col, r)
    WC_ROWS["dinv"] = r; r += 1

    write_label(ws, r, "ΔPayables (incr = source of cash)")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        prev = get_column_letter(2 + t - 1)
        formula = f"=({col}{WC_ROWS['ap']}-{prev}{WC_ROWS['ap']})"
        write_formula(ws, r, 2 + t, formula, NUM_BN)
        reg(f"w.dap_t{t}", SH_WC, col, r)
    WC_ROWS["dap"] = r; r += 2

    # Total
    write_label(ws, r, "Total Δ working capital", bold=True, banded=True)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{WC_ROWS['dar']}+{col}{WC_ROWS['dinv']}+{col}{WC_ROWS['dap']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
        reg(f"w.dwc_t{t}", SH_WC, col, r)
    WC_ROWS["dwc"] = r


# =============================================================================
# 13. TAX SCHEDULE
# =============================================================================

TAX_ROWS = {}

def build_tax(ws):
    ws.column_dimensions["A"].width = 32
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = 13

    ws["A1"] = "TAX SCHEDULE"
    ws["A1"].font = section_font
    ws["A2"] = "Simplified: current tax = max(0, pretax × effective rate). Deferred tax held flat."
    ws["A2"].font = sub_font

    # Header — only Y1-Y5
    ws.cell(row=4, column=1).fill = header_fill
    for t in range(1, 6):
        c = ws.cell(row=4, column=2 + t, value=f"Y{t}")
        c.font = header_font; c.fill = header_fill; c.alignment = right

    r = 6
    write_label(ws, r, "Pre-tax income (from IS)")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={C(f'is.pretax_t{t}')}", NUM_BN)
    TAX_ROWS["pretax"] = r; r += 1

    write_label(ws, r, "Effective tax rate")
    for t in range(1, 6):
        write_formula(ws, r, 2 + t, f"={C('a.tax')}", NUM_PCT)
    TAX_ROWS["rate"] = r; r += 1

    write_label(ws, r, "Current tax expense", bold=True, banded=True)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"=MAX(0,{col}{TAX_ROWS['pretax']}*{col}{TAX_ROWS['rate']})"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
        reg(f"t.expense_t{t}", SH_TAX, col, r)
    TAX_ROWS["expense"] = r


# =============================================================================
# 14. INCOME STATEMENT — rewritten to use real Debt/Tax row references
# =============================================================================

def build_is_v2(ws):
    """Built after Debt/Tax exist, so we can wire real cross-sheet refs."""
    ws.column_dimensions["A"].width = 34
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = 13

    ws["A1"] = "INCOME STATEMENT"
    ws["A1"].font = section_font
    ws["A2"] = "All figures in SAR billion. Five-year explicit forecast."
    ws["A2"].font = sub_font

    ws.cell(row=4, column=1).fill = header_fill
    for t in range(1, 6):
        c = ws.cell(row=4, column=2 + t, value=f"Y{t}")
        c.font = header_font; c.fill = header_fill; c.alignment = right

    r = 5
    # Service revenue
    write_label(ws, r, "Service revenue", bold=True)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={C(f'dr.rev_t{t}')}", NUM_BN, bold=True)
        reg(f"is.revenue_t{t}", SH_IS, col, r)
    IS_ROWS["revenue"] = r
    r += 2

    write_label(ws, r, "Cost lines (% of revenue)"); r += 1

    cost_defs = [
        ("COGS / interconnect / content",     "cogs_pct",         "cogs"),
        ("Network operating costs",           "network_opex_pct", "network"),
        ("Employee costs",                    "employee_pct",     "emp"),
        ("Customer acquisition / commercial", "cust_acq_pct",     "cust_acq"),
        ("Other opex / G&A",                  "other_opex_pct",   "other"),
    ]
    for label, akey, ikey in cost_defs:
        write_label(ws, r, label, indent=1)
        for t in range(1, 6):
            col = get_column_letter(2 + t)
            rev_cell = f"{col}{IS_ROWS['revenue']}"
            write_formula(ws, r, 2 + t, f"=-{C('a.' + akey)}*{rev_cell}", NUM_BN)
            reg(f"is.{ikey}_t{t}", SH_IS, col, r)
        IS_ROWS[ikey] = r
        r += 1

    r += 1
    write_label(ws, r, "EBITDA", bold=True, banded=True)
    cost_rows = [IS_ROWS[k] for _, _, k in cost_defs]
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        terms = [f"{col}{IS_ROWS['revenue']}"] + [f"{col}{cr}" for cr in cost_rows]
        write_formula(ws, r, 2 + t, "=" + "+".join(terms), NUM_BN, bold=True, banded=True)
        reg(f"is.ebitda_t{t}", SH_IS, col, r)
    IS_ROWS["ebitda"] = r
    r += 2

    write_label(ws, r, "Depreciation & amortisation")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={C(f'p.da_t{t}')}", NUM_BN)
        reg(f"is.da_t{t}", SH_IS, col, r)
    IS_ROWS["da"] = r
    r += 1

    r += 1
    write_label(ws, r, "EBIT (operating profit)", bold=True, banded=True)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{IS_ROWS['ebitda']}+{col}{IS_ROWS['da']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
        reg(f"is.ebit_t{t}", SH_IS, col, r)
    IS_ROWS["ebit"] = r
    r += 2

    write_label(ws, r, "Finance income")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={C('a.finance_income_pct')}*{col}{IS_ROWS['revenue']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN)
        reg(f"is.fin_inc_t{t}", SH_IS, col, r)
    IS_ROWS["fin_inc"] = r
    r += 1

    write_label(ws, r, "Interest expense")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"=-{C(f'd.interest_t{t}')}", NUM_BN)
        reg(f"is.interest_t{t}", SH_IS, col, r)
    IS_ROWS["interest"] = r
    r += 1

    r += 1
    write_label(ws, r, "Pre-tax income", bold=True)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{IS_ROWS['ebit']}+{col}{IS_ROWS['fin_inc']}+{col}{IS_ROWS['interest']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True)
        reg(f"is.pretax_t{t}", SH_IS, col, r)
    IS_ROWS["pretax"] = r
    r += 1

    write_label(ws, r, "Income tax")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"=-{C(f't.expense_t{t}')}", NUM_BN)
        reg(f"is.tax_t{t}", SH_IS, col, r)
    IS_ROWS["tax"] = r
    r += 1

    r += 1
    write_label(ws, r, "NET INCOME", bold=True, banded=True)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{IS_ROWS['pretax']}+{col}{IS_ROWS['tax']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
        reg(f"is.ni_t{t}", SH_IS, col, r)
    IS_ROWS["ni"] = r
    r += 2

    section_header(ws, r, "MEMO"); r += 1
    for label, key, top in [
        ("EBITDA margin", "ebitda_margin", IS_ROWS["ebitda"]),
        ("EBIT margin",   "ebit_margin",   IS_ROWS["ebit"]),
        ("Net margin",    "net_margin",    IS_ROWS["ni"]),
    ]:
        write_label(ws, r, label)
        for t in range(1, 6):
            col = get_column_letter(2 + t)
            write_formula(ws, r, 2 + t, f"={col}{top}/{col}{IS_ROWS['revenue']}", NUM_PCT)
        IS_ROWS[key] = r
        r += 1


# =============================================================================
# 15. EQUITY ROLLFORWARD
# =============================================================================

EQ_ROWS = {}

def build_equity(ws):
    ws.column_dimensions["A"].width = 32
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = 13

    ws["A1"] = "EQUITY ROLLFORWARD"
    ws["A1"].font = section_font
    ws["A2"] = "Closing equity = opening + net income − dividends + OCI movements. Drives BS equity section."
    ws["A2"].font = sub_font

    year_header(ws, 4, 0, 5)
    r = 6

    # Share capital (flat)
    write_label(ws, r, "Share capital")
    write_formula(ws, r, 2, f"={C('a.share_capital')}", NUM_BN)
    for t in range(1, 6):
        prev = get_column_letter(2 + t - 1)
        write_formula(ws, r, 2 + t, f"={prev}{r}", NUM_BN)
    EQ_ROWS["share_cap"] = r; r += 1
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        reg(f"eq.share_cap_t{t}", SH_EQ, col, EQ_ROWS["share_cap"])

    # Retained earnings rollforward
    write_label(ws, r, "Opening retained earnings")
    write_formula(ws, r, 2, '""', '@')
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        if t == 1:
            write_formula(ws, r, 2 + t, f"={C('a.re_y0')}", NUM_BN)
        else:
            prev_col = get_column_letter(2 + t - 1)
            write_formula(ws, r, 2 + t, f"={prev_col}{r + 3}", NUM_BN)  # prev closing RE
    EQ_ROWS["re_open"] = r; r += 1

    write_label(ws, r, "+ Net income for period")
    write_formula(ws, r, 2, '""', '@')
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={C(f'is.ni_t{t}')}", NUM_BN)
    EQ_ROWS["ni"] = r; r += 1

    write_label(ws, r, "− Dividends paid")
    write_formula(ws, r, 2, '""', '@')
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"=-{C('a.div_payout')}*{C(f'is.ni_t{t}')}"
        write_formula(ws, r, 2 + t, formula, NUM_BN)
        reg(f"eq.div_t{t}", SH_EQ, col, r)
    EQ_ROWS["div"] = r; r += 1

    write_label(ws, r, "Closing retained earnings", bold=True, banded=True)
    write_formula(ws, r, 2, f"={C('a.re_y0')}", NUM_BN, bold=True, banded=True)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{EQ_ROWS['re_open']}+{col}{EQ_ROWS['ni']}+{col}{EQ_ROWS['div']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
    EQ_ROWS["re_close"] = r
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        reg(f"eq.re_t{t}", SH_EQ, col, EQ_ROWS["re_close"])
    r += 2

    # OCI reserve (held flat in default)
    write_label(ws, r, "OCI reserve (held flat)")
    write_formula(ws, r, 2, f"={C('a.oci_reserve_y0')}", NUM_BN)
    for t in range(1, 6):
        prev = get_column_letter(2 + t - 1)
        write_formula(ws, r, 2 + t, f"={prev}{r}", NUM_BN)
    EQ_ROWS["oci"] = r
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        reg(f"eq.oci_t{t}", SH_EQ, col, EQ_ROWS["oci"])
    r += 1

    # Minority interest (held flat)
    write_label(ws, r, "Minority interest (held flat)")
    write_formula(ws, r, 2, f"={C('a.minority_y0')}", NUM_BN)
    for t in range(1, 6):
        prev = get_column_letter(2 + t - 1)
        write_formula(ws, r, 2 + t, f"={prev}{r}", NUM_BN)
    EQ_ROWS["minority"] = r
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        reg(f"eq.minority_t{t}", SH_EQ, col, EQ_ROWS["minority"])
    r += 2

    # Total equity
    write_label(ws, r, "TOTAL EQUITY", bold=True, banded=True)
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{EQ_ROWS['share_cap']}+{col}{EQ_ROWS['re_close']}+{col}{EQ_ROWS['oci']}+{col}{EQ_ROWS['minority']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
        reg(f"eq.total_t{t}", SH_EQ, col, r)
    EQ_ROWS["total"] = r


# =============================================================================
# 16. BALANCE SHEET
# =============================================================================

BS_ROWS = {}

def build_bs(ws):
    ws.column_dimensions["A"].width = 34
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = 13

    ws["A1"] = "BALANCE SHEET"
    ws["A1"].font = section_font
    ws["A2"] = "All figures in SAR billion. Y0 opening + Y1-Y5 forecast. Must balance (see Checks)."
    ws["A2"].font = sub_font

    year_header(ws, 4, 0, 5)

    r = 6
    # ---- ASSETS ----
    section_header(ws, r, "ASSETS"); r += 1

    # Cash (plug, from CFS)
    write_label(ws, r, "Cash & equivalents")
    write_formula(ws, r, 2, f"={C('a.cash_y0')}", NUM_BN)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        prev = get_column_letter(2 + t - 1)
        write_formula(ws, r, 2 + t, f"={prev}{r}+CFS!{col}{CFS_PLACEHOLDER_NETCASH}", NUM_BN)
    BS_ROWS["cash"] = r
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        reg(f"b.cash_t{t}", SH_BS, col, r)
    r += 1
    # Patch debt schedule placeholder for cash row
    global BS_PLACEHOLDER_CASH
    BS_PLACEHOLDER_CASH = r - 1  # because we already advanced r

    # AR
    write_label(ws, r, "Accounts receivable")
    write_formula(ws, r, 2, f"={C('a.ar_y0')}", NUM_BN)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={C(f'w.ar_t{t}')}", NUM_BN)
    BS_ROWS["ar"] = r; r += 1

    write_label(ws, r, "Inventory")
    write_formula(ws, r, 2, f"={C('a.inv_y0')}", NUM_BN)
    for t in range(1, 6):
        write_formula(ws, r, 2 + t, f"={C(f'w.inv_t{t}')}", NUM_BN)
    BS_ROWS["inv"] = r; r += 1

    write_label(ws, r, "Other current assets (flat)")
    write_formula(ws, r, 2, f"={C('a.other_ca_y0')}", NUM_BN)
    for t in range(1, 6):
        prev = get_column_letter(2 + t - 1)
        write_formula(ws, r, 2 + t, f"={prev}{r}", NUM_BN)
    BS_ROWS["other_ca"] = r; r += 1

    # Current assets total
    write_label(ws, r, "Total current assets", bold=True)
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{BS_ROWS['cash']}+{col}{BS_ROWS['ar']}+{col}{BS_ROWS['inv']}+{col}{BS_ROWS['other_ca']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True)
    BS_ROWS["ca_total"] = r; r += 2

    # Non-current assets
    write_label(ws, r, "PP&E")
    write_formula(ws, r, 2, f"={C('a.ppe_y0')}", NUM_BN)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={C(f'p.closing_t{t}')}", NUM_BN)
    BS_ROWS["ppe"] = r; r += 1

    write_label(ws, r, "Goodwill (flat)")
    write_formula(ws, r, 2, f"={C('a.goodwill_y0')}", NUM_BN)
    for t in range(1, 6):
        prev = get_column_letter(2 + t - 1)
        write_formula(ws, r, 2 + t, f"={prev}{r}", NUM_BN)
    BS_ROWS["goodwill"] = r; r += 1

    write_label(ws, r, "Intangibles (flat)")
    write_formula(ws, r, 2, f"={C('a.intangibles_y0')}", NUM_BN)
    for t in range(1, 6):
        prev = get_column_letter(2 + t - 1)
        write_formula(ws, r, 2 + t, f"={prev}{r}", NUM_BN)
    BS_ROWS["intangibles"] = r; r += 1

    write_label(ws, r, "Investments / associates (flat)")
    write_formula(ws, r, 2, f"={C('a.investments_y0')}", NUM_BN)
    for t in range(1, 6):
        prev = get_column_letter(2 + t - 1)
        write_formula(ws, r, 2 + t, f"={prev}{r}", NUM_BN)
    BS_ROWS["investments"] = r; r += 1

    write_label(ws, r, "Deferred tax assets (flat)")
    write_formula(ws, r, 2, f"={C('a.dta_y0')}", NUM_BN)
    for t in range(1, 6):
        prev = get_column_letter(2 + t - 1)
        write_formula(ws, r, 2 + t, f"={prev}{r}", NUM_BN)
    BS_ROWS["dta"] = r; r += 1

    write_label(ws, r, "Total non-current assets", bold=True)
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        cells = [BS_ROWS["ppe"], BS_ROWS["goodwill"], BS_ROWS["intangibles"], BS_ROWS["investments"], BS_ROWS["dta"]]
        formula = "=" + "+".join(f"{col}{x}" for x in cells)
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True)
    BS_ROWS["nca_total"] = r; r += 2

    # Total assets
    write_label(ws, r, "TOTAL ASSETS", bold=True, banded=True)
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{BS_ROWS['ca_total']}+{col}{BS_ROWS['nca_total']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
        reg(f"b.total_assets_t{t}", SH_BS, col, r)
    BS_ROWS["total_assets"] = r; r += 2

    # ---- LIABILITIES ----
    section_header(ws, r, "LIABILITIES"); r += 1

    write_label(ws, r, "Accounts payable")
    write_formula(ws, r, 2, f"={C('a.ap_y0')}", NUM_BN)
    for t in range(1, 6):
        write_formula(ws, r, 2 + t, f"={C(f'w.ap_t{t}')}", NUM_BN)
    BS_ROWS["ap"] = r; r += 1

    write_label(ws, r, "Other current liabilities (flat)")
    write_formula(ws, r, 2, f"={C('a.other_cl_y0')}", NUM_BN)
    for t in range(1, 6):
        prev = get_column_letter(2 + t - 1)
        write_formula(ws, r, 2 + t, f"={prev}{r}", NUM_BN)
    BS_ROWS["other_cl"] = r; r += 1

    write_label(ws, r, "Short-term debt")
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={SH_DEBT}!{col}{DEBT_ROWS['st']}", NUM_BN)
    BS_ROWS["st_debt"] = r; r += 1

    write_label(ws, r, "Total current liabilities", bold=True)
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{BS_ROWS['ap']}+{col}{BS_ROWS['other_cl']}+{col}{BS_ROWS['st_debt']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True)
    BS_ROWS["cl_total"] = r; r += 2

    write_label(ws, r, "Long-term debt")
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={SH_DEBT}!{col}{DEBT_ROWS['lt']}", NUM_BN)
    BS_ROWS["lt_debt"] = r; r += 1

    write_label(ws, r, "Lease liabilities")
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={SH_DEBT}!{col}{DEBT_ROWS['lease']}", NUM_BN)
    BS_ROWS["lease"] = r; r += 1

    write_label(ws, r, "Provisions (flat)")
    write_formula(ws, r, 2, f"={C('a.provisions_y0')}", NUM_BN)
    for t in range(1, 6):
        prev = get_column_letter(2 + t - 1)
        write_formula(ws, r, 2 + t, f"={prev}{r}", NUM_BN)
    BS_ROWS["provisions"] = r; r += 1

    write_label(ws, r, "Deferred tax liabilities (flat)")
    write_formula(ws, r, 2, f"={C('a.dtl_y0')}", NUM_BN)
    for t in range(1, 6):
        prev = get_column_letter(2 + t - 1)
        write_formula(ws, r, 2 + t, f"={prev}{r}", NUM_BN)
    BS_ROWS["dtl"] = r; r += 1

    write_label(ws, r, "Other non-current liabilities (flat)")
    write_formula(ws, r, 2, f"={C('a.other_ncl_y0')}", NUM_BN)
    for t in range(1, 6):
        prev = get_column_letter(2 + t - 1)
        write_formula(ws, r, 2 + t, f"={prev}{r}", NUM_BN)
    BS_ROWS["other_ncl"] = r; r += 1

    write_label(ws, r, "Total non-current liabilities", bold=True)
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        cells = [BS_ROWS["lt_debt"], BS_ROWS["lease"], BS_ROWS["provisions"], BS_ROWS["dtl"], BS_ROWS["other_ncl"]]
        formula = "=" + "+".join(f"{col}{x}" for x in cells)
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True)
    BS_ROWS["ncl_total"] = r; r += 2

    write_label(ws, r, "TOTAL LIABILITIES", bold=True, banded=True)
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{BS_ROWS['cl_total']}+{col}{BS_ROWS['ncl_total']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
        reg(f"b.total_liab_t{t}", SH_BS, col, r)
    BS_ROWS["total_liab"] = r; r += 2

    # ---- EQUITY ----
    section_header(ws, r, "EQUITY"); r += 1

    write_label(ws, r, "Share capital")
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={C(f'eq.share_cap_t{t}')}", NUM_BN)
    r += 1
    write_label(ws, r, "Retained earnings")
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={C(f'eq.re_t{t}')}", NUM_BN)
    r += 1
    write_label(ws, r, "OCI reserve")
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={C(f'eq.oci_t{t}')}", NUM_BN)
    r += 1
    write_label(ws, r, "Minority interest")
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={C(f'eq.minority_t{t}')}", NUM_BN)
    r += 1
    write_label(ws, r, "TOTAL EQUITY", bold=True, banded=True)
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={C(f'eq.total_t{t}')}", NUM_BN, bold=True, banded=True)
        reg(f"b.total_equity_t{t}", SH_BS, col, r)
    BS_ROWS["total_equity"] = r; r += 2

    # Check row: A - L - E
    write_label(ws, r, "Balance check (A − L − E)", bold=True)
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{BS_ROWS['total_assets']}-{col}{BS_ROWS['total_liab']}-{col}{BS_ROWS['total_equity']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True)
    BS_ROWS["check"] = r


CFS_PLACEHOLDER_NETCASH = 0


# =============================================================================
# 17. CASH FLOW STATEMENT
# =============================================================================

CFS_ROWS = {}

def build_cfs(ws):
    ws.column_dimensions["A"].width = 34
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = 13

    ws["A1"] = "CASH FLOW STATEMENT (indirect method)"
    ws["A1"].font = section_font
    ws["A2"] = "All figures in SAR billion. Net change in cash flows to BS."
    ws["A2"].font = sub_font

    ws.cell(row=4, column=1).fill = header_fill
    for t in range(1, 6):
        c = ws.cell(row=4, column=2 + t, value=f"Y{t}")
        c.font = header_font; c.fill = header_fill; c.alignment = right

    r = 6
    section_header(ws, r, "OPERATING"); r += 1

    write_label(ws, r, "Net income")
    for t in range(1, 6):
        write_formula(ws, r, 2 + t, f"={C(f'is.ni_t{t}')}", NUM_BN)
    CFS_ROWS["ni"] = r; r += 1

    write_label(ws, r, "+ D&A (non-cash addback)")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        # da on IS is negative; we add back the absolute value
        write_formula(ws, r, 2 + t, f"=-{C(f'is.da_t{t}')}", NUM_BN)
    CFS_ROWS["da"] = r; r += 1

    write_label(ws, r, "+/− Δ working capital")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={C(f'w.dwc_t{t}')}", NUM_BN)
    CFS_ROWS["dwc"] = r; r += 1

    write_label(ws, r, "Cash from operations (CFO)", bold=True, banded=True)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{CFS_ROWS['ni']}+{col}{CFS_ROWS['da']}+{col}{CFS_ROWS['dwc']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
        reg(f"c.cfo_t{t}", SH_CFS, col, r)
    CFS_ROWS["cfo"] = r; r += 2

    section_header(ws, r, "INVESTING"); r += 1

    write_label(ws, r, "− Capex")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"=-{C(f'p.capex_t{t}')}", NUM_BN)
    CFS_ROWS["capex"] = r; r += 1

    write_label(ws, r, "Other investing (flat / nil)")
    for t in range(1, 6):
        write_formula(ws, r, 2 + t, "=0", NUM_BN)
    CFS_ROWS["other_inv"] = r; r += 1

    write_label(ws, r, "Cash from investing (CFI)", bold=True, banded=True)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{CFS_ROWS['capex']}+{col}{CFS_ROWS['other_inv']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
        reg(f"c.cfi_t{t}", SH_CFS, col, r)
    CFS_ROWS["cfi"] = r; r += 2

    section_header(ws, r, "FINANCING"); r += 1

    write_label(ws, r, "+/− Net debt issuance (flat assumption)")
    for t in range(1, 6):
        write_formula(ws, r, 2 + t, "=0", NUM_BN)  # debt held flat
    CFS_ROWS["debt_chg"] = r; r += 1

    write_label(ws, r, "− Dividends paid")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={C(f'eq.div_t{t}')}", NUM_BN)
    CFS_ROWS["div"] = r; r += 1

    write_label(ws, r, "Cash from financing (CFF)", bold=True, banded=True)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{CFS_ROWS['debt_chg']}+{col}{CFS_ROWS['div']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
        reg(f"c.cff_t{t}", SH_CFS, col, r)
    CFS_ROWS["cff"] = r; r += 2

    write_label(ws, r, "Net change in cash", bold=True, banded=True)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{CFS_ROWS['cfo']}+{col}{CFS_ROWS['cfi']}+{col}{CFS_ROWS['cff']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
        reg(f"c.netcash_t{t}", SH_CFS, col, r)
    CFS_ROWS["netcash"] = r

    # Patch BS placeholder so cash row picks up the netcash
    global CFS_PLACEHOLDER_NETCASH
    CFS_PLACEHOLDER_NETCASH = r

    r += 1
    write_label(ws, r, "Opening cash")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        if t == 1:
            write_formula(ws, r, 2 + t, f"={C('a.cash_y0')}", NUM_BN)
        else:
            prev = get_column_letter(2 + t - 1)
            write_formula(ws, r, 2 + t, f"={prev}{r + 1}", NUM_BN)
    CFS_ROWS["open_cash"] = r; r += 1

    write_label(ws, r, "Closing cash", bold=True, banded=True)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{CFS_ROWS['open_cash']}+{col}{CFS_ROWS['netcash']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
    CFS_ROWS["close_cash"] = r


# =============================================================================
# 18. BUDGET (Y1 monthly)
# =============================================================================

def build_budget(ws):
    ws.column_dimensions["A"].width = 30
    for c in range(2, 16):
        ws.column_dimensions[get_column_letter(c)].width = 11

    ws["A1"] = "Y1 MONTHLY BUDGET"
    ws["A1"].font = section_font
    ws["A2"] = "Y1 P&L line items, distributed by month per the seasonality weights below. Variance columns are placeholders for actuals."
    ws["A2"].font = sub_font

    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

    # Header
    ws.cell(row=4, column=1).fill = header_fill
    for i, m in enumerate(months):
        c = ws.cell(row=4, column=2 + i, value=m)
        c.font = header_font; c.fill = header_fill; c.alignment = right
    c = ws.cell(row=4, column=14, value="Y1 Total")
    c.font = header_font; c.fill = header_fill; c.alignment = right

    # Weights row
    r = 5
    write_label(ws, r, "Seasonality weight")
    for i, w in enumerate(D["month_weights"]):
        write_input(ws, r, 2 + i, w, NUM_PCT)
    write_formula(ws, r, 14, f"=SUM(B{r}:M{r})", NUM_PCT, bold=True)
    weights_row = r; r += 2

    # P&L lines (using IS Y1 references)
    section_header(ws, r, "P&L lines (SAR bn)"); r += 1

    line_defs = [
        ("Service revenue",         f"={C('is.revenue_t1')}"),
        ("COGS / interconnect",     f"={C('is.cogs_t1')}"),
        ("Network opex",            f"={C('is.network_t1')}"),
        ("Employee costs",          f"={C('is.emp_t1')}"),
        ("Customer acquisition",    f"={C('is.cust_acq_t1')}"),
        ("Other opex / G&A",        f"={C('is.other_t1')}"),
        ("EBITDA",                  f"={C('is.ebitda_t1')}"),
        ("D&A",                     f"={C('is.da_t1')}"),
        ("EBIT",                    f"={C('is.ebit_t1')}"),
        ("Net income",              f"={C('is.ni_t1')}"),
    ]
    for label, y1_formula in line_defs:
        bold = label in ("EBITDA", "EBIT", "Net income")
        write_label(ws, r, label, bold=bold, banded=bold)
        # Per-month = weight × annual
        for i in range(12):
            col = get_column_letter(2 + i)
            formula = f"={col}{weights_row}*({y1_formula[1:]})"
            write_formula(ws, r, 2 + i, formula, NUM_BN, bold=bold, banded=bold)
        # Total = SUM
        write_formula(ws, r, 14, f"=SUM(B{r}:M{r})", NUM_BN, bold=True, banded=bold)
        r += 1

    r += 1
    section_header(ws, r, "Quarterly summary"); r += 1
    write_label(ws, r, "Service revenue (quarterly)")
    # Columns N..Q for quarterly. Actually simpler: print Q1..Q4 in rows.
    ws.cell(row=r, column=2, value="Q1"); ws.cell(row=r, column=2).font = sub_font
    ws.cell(row=r, column=3, value="Q2"); ws.cell(row=r, column=3).font = sub_font
    ws.cell(row=r, column=4, value="Q3"); ws.cell(row=r, column=4).font = sub_font
    ws.cell(row=r, column=5, value="Q4"); ws.cell(row=r, column=5).font = sub_font
    r += 1
    write_label(ws, r, "Quarterly service revenue")
    # First service revenue row is r-something; we tracked first line above
    sr_row = weights_row + 2  # first line under "P&L lines" section header
    for q, cols in enumerate([(2,3,4),(5,6,7),(8,9,10),(11,12,13)]):
        col = get_column_letter(2 + q)
        terms = "+".join(f"{get_column_letter(c)}{sr_row}" for c in cols)
        write_formula(ws, r, 2 + q, f"={terms}", NUM_BN, bold=True, banded=True)


# =============================================================================
# 19. SCENARIOS
# =============================================================================

def build_scenarios(ws):
    ws.column_dimensions["A"].width = 30
    for c in range(2, 6):
        ws.column_dimensions[get_column_letter(c)].width = 14

    ws["A1"] = "SCENARIO MANAGER"
    ws["A1"].font = section_font
    ws["A2"] = "Reference table for downside / base / upside parameter sets. Active scenario set on Assumptions sheet."
    ws["A2"].font = sub_font

    # Header
    ws.cell(row=4, column=1).fill = header_fill
    for i, name in enumerate(["1: Downside", "2: Base", "3: Upside"]):
        c = ws.cell(row=4, column=2 + i, value=name)
        c.font = header_font; c.fill = header_fill; c.alignment = right

    r = 6
    write_label(ws, r, "Revenue growth")
    for i, v in enumerate(D["scen_growth"]):
        write_input(ws, r, 2 + i, v, NUM_PCT)
    r += 1
    write_label(ws, r, "EBITDA margin (target)")
    for i, v in enumerate(D["scen_margin"]):
        write_input(ws, r, 2 + i, v, NUM_PCT)
    r += 1
    write_label(ws, r, "Capex intensity")
    for i, v in enumerate(D["scen_capex"]):
        write_input(ws, r, 2 + i, v, NUM_PCT)
    r += 2

    write_label(ws, r, "Active scenario (1/2/3)", bold=True)
    write_formula(ws, r, 2, f"={C('a.scen_default')}", NUM_INT, bold=True, banded=True)
    r += 2

    section_header(ws, r, "Active values (use INDEX in Assumptions to wire these)"); r += 1
    write_label(ws, r, "Note")
    note = ws.cell(row=r, column=2, value=("Manually overwrite Assumptions!B6/B7/B9 with these "
                                          "values, or build INDEX formulas. Kept manual for clarity."))
    note.font = sub_font; note.alignment = left
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)


# =============================================================================
# 20. SENSITIVITY
# =============================================================================

def build_sensitivity(ws):
    ws.column_dimensions["A"].width = 30
    for c in range(2, 10):
        ws.column_dimensions[get_column_letter(c)].width = 12

    ws["A1"] = "SENSITIVITY TABLES"
    ws["A1"].font = section_font
    ws["A2"] = "Two-way tables on the key valuation levers. Outputs computed via TABLE-style formulas — change axis values to flex."
    ws["A2"].font = sub_font

    # ---- Table 1: WACC × g => Per-share value ----
    r = 4
    section_header(ws, r, "Per-share value (SAR) — WACC × terminal growth"); r += 1

    waccs = [0.07, 0.08, 0.09, 0.10, 0.11]
    gs    = [0.01, 0.02, 0.03, 0.04]

    # Header row (WACCs)
    write_label(ws, r, "WACC →")
    for i, w in enumerate(waccs):
        write_input(ws, r, 2 + i, w, NUM_PCT)
    r += 1
    # Body
    for j, g in enumerate(gs):
        write_input(ws, r, 1, g, NUM_PCT)
        # We compute approximate per-share by replicating the DCF math in-formula.
        for i, w in enumerate(waccs):
            # Use IS revenue Y5, EBITDA margin from a.ebitda_target, derive simple FCFF
            # For simplicity reference Valuation EV computation pattern but with these w, g
            # Per-share = (PV(explicit FCFF using w) + PV(TV using w,g) - net_debt) * 1000 / shares
            # We construct the formula inline. Since FCFF depends on a.wacc, we cannot
            # use Valuation sheet directly. Replicate inline:
            terms = []
            for t in range(1, 6):
                ebit = f"{C(f'is.ebit_t{t}')}"
                da = f"(-{C(f'is.da_t{t}')})"
                capex = f"{C(f'p.capex_t{t}')}"
                dwc = f"{C(f'w.dwc_t{t}')}"
                fcff = f"({ebit}*(1-{C('a.tax')})+{da}-{capex}+{dwc})"
                terms.append(f"({fcff}/(1+{get_column_letter(2 + i)}{r - 1 - j})^{t})")
            sum_fcff = "+".join(terms)
            tv = (f"(({C('is.ebit_t5')}*(1-{C('a.tax')})+(-{C('is.da_t5')})-{C('p.capex_t5')}+{C('w.dwc_t5')})"
                  f"*(1+A{r})/({get_column_letter(2 + i)}{r - 1 - j}-A{r}))")
            pv_tv = f"({tv}/(1+{get_column_letter(2 + i)}{r - 1 - j})^5)"
            ev_expr = f"({sum_fcff}+{pv_tv})"
            equity_expr = f"({ev_expr}-{C('d.net_y0')})"
            ps_expr = f"={equity_expr}*1000/{C('a.shares')}"
            write_formula(ws, r, 2 + i, ps_expr, '#,##0.0')
        r += 1

    r += 2
    section_header(ws, r, "Enterprise value (SAR bn) — Y1 growth × EBITDA target margin"); r += 1
    growths = [0.02, 0.04, 0.06, 0.08, 0.10]
    margins = [0.30, 0.33, 0.36, 0.39, 0.42]
    write_label(ws, r, "Growth →")
    for i, gv in enumerate(growths):
        write_input(ws, r, 2 + i, gv, NUM_PCT)
    r += 1
    # Body: highly simplified — EV ≈ FCFF_Y5 / (wacc - g_terminal) where FCFF_Y5 derived
    # from simple formula: rev0*(1+growth)^5 * margin * (1-tax) - rev0*(1+growth)^5 * capex_pct
    for j, m in enumerate(margins):
        write_input(ws, r, 1, m, NUM_PCT)
        for i, gv in enumerate(growths):
            rev5 = f"({C('a.rev0')}*(1+{get_column_letter(2 + i)}{r - 1 - j})^5)"
            fcff5 = f"({rev5}*({m}-{C('a.capex_pct')})*(1-{C('a.tax')}))"
            tv = f"({fcff5}*(1+{C('a.g')})/({C('a.wacc')}-{C('a.g')}))"
            pv_tv = f"({tv}/(1+{C('a.wacc')})^5)"
            # Quick PV explicit: 5x FCFF5 * discount (rough)
            sum_pv = f"({fcff5}*5/(1+{C('a.wacc')})^3)"  # simplification
            write_formula(ws, r, 2 + i, f"={sum_pv}+{pv_tv}", NUM_BN)
        r += 1


# =============================================================================
# 21. VALUATION
# =============================================================================

VAL_ROWS = {}

def build_valuation(ws):
    ws.column_dimensions["A"].width = 32
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = 13

    ws["A1"] = "DCF VALUATION"
    ws["A1"].font = section_font
    ws["A2"] = "Free cash flow to firm (FCFF), discounted at WACC. Terminal value via Gordon growth."
    ws["A2"].font = sub_font

    ws.cell(row=4, column=1).fill = header_fill
    for t in range(1, 6):
        c = ws.cell(row=4, column=2 + t, value=f"Y{t}")
        c.font = header_font; c.fill = header_fill; c.alignment = right

    r = 6
    write_label(ws, r, "EBIT (from IS)")
    for t in range(1, 6):
        write_formula(ws, r, 2 + t, f"={C(f'is.ebit_t{t}')}", NUM_BN)
    VAL_ROWS["ebit"] = r; r += 1

    write_label(ws, r, "× (1 − tax)")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={col}{VAL_ROWS['ebit']}*(1-{C('a.tax')})", NUM_BN)
    VAL_ROWS["nopat"] = r; r += 1

    write_label(ws, r, "+ D&A (non-cash addback)")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"=-{C(f'is.da_t{t}')}", NUM_BN)
    VAL_ROWS["da"] = r; r += 1

    write_label(ws, r, "− Capex")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"=-{C(f'p.capex_t{t}')}", NUM_BN)
    VAL_ROWS["capex"] = r; r += 1

    write_label(ws, r, "+/− Δ working capital")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"={C(f'w.dwc_t{t}')}", NUM_BN)
    VAL_ROWS["dwc"] = r; r += 1

    write_label(ws, r, "FCFF", bold=True, banded=True)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{VAL_ROWS['nopat']}+{col}{VAL_ROWS['da']}+{col}{VAL_ROWS['capex']}+{col}{VAL_ROWS['dwc']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True, banded=True)
    VAL_ROWS["fcff"] = r; r += 2

    write_label(ws, r, "Discount factor (1/(1+WACC)^t)")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        write_formula(ws, r, 2 + t, f"=1/(1+{C('a.wacc')})^{t}", '0.0000')
    VAL_ROWS["df"] = r; r += 1

    write_label(ws, r, "PV of FCFF", bold=True)
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        formula = f"={col}{VAL_ROWS['fcff']}*{col}{VAL_ROWS['df']}"
        write_formula(ws, r, 2 + t, formula, NUM_BN, bold=True)
    VAL_ROWS["pv_fcff"] = r; r += 2

    # Summary block (col A label, col B value)
    write_label(ws, r, "Sum PV (explicit Y1-Y5)", bold=True)
    write_formula(ws, r, 2, f"=SUM(C{VAL_ROWS['pv_fcff']}:G{VAL_ROWS['pv_fcff']})", NUM_BN, bold=True)
    VAL_ROWS["sum_pv"] = r; r += 1

    write_label(ws, r, "Terminal value (Gordon)")
    write_formula(ws, r, 2, f"=G{VAL_ROWS['fcff']}*(1+{C('a.g')})/({C('a.wacc')}-{C('a.g')})", NUM_BN)
    VAL_ROWS["tv"] = r; r += 1

    write_label(ws, r, "PV of terminal value", bold=True)
    write_formula(ws, r, 2, f"=B{VAL_ROWS['tv']}*G{VAL_ROWS['df']}", NUM_BN, bold=True)
    VAL_ROWS["pv_tv"] = r; r += 2

    write_label(ws, r, "ENTERPRISE VALUE", bold=True, banded=True)
    write_formula(ws, r, 2, f"=B{VAL_ROWS['sum_pv']}+B{VAL_ROWS['pv_tv']}", NUM_BN, bold=True, banded=True)
    VAL_ROWS["ev"] = r
    reg("v.ev", SH_VAL, "B", r); r += 1

    write_label(ws, r, "− Net debt (Y0)")
    write_formula(ws, r, 2, f"=-{C('d.net_y0')}", NUM_BN)
    r += 1

    write_label(ws, r, "EQUITY VALUE", bold=True, banded=True)
    write_formula(ws, r, 2, f"=B{VAL_ROWS['ev']}+B{r-1}", NUM_BN, bold=True, banded=True)
    VAL_ROWS["equity"] = r; r += 1

    write_label(ws, r, "Per-share value (SAR)", bold=True, banded=True)
    write_formula(ws, r, 2, f"=B{VAL_ROWS['equity']}*1000/{C('a.shares')}", '#,##0.00', bold=True, banded=True)
    VAL_ROWS["per_share"] = r; r += 2

    section_header(ws, r, "Cross-checks"); r += 1
    write_label(ws, r, "Implied FY+1 EV/EBITDA")
    write_formula(ws, r, 2, f"=B{VAL_ROWS['ev']}/{C('is.ebitda_t1')}", NUM_X)
    r += 1
    write_label(ws, r, "Implied FY+1 EV/EBIT")
    write_formula(ws, r, 2, f"=B{VAL_ROWS['ev']}/{C('is.ebit_t1')}", NUM_X)
    r += 1
    write_label(ws, r, "Implied FY+1 P/E")
    write_formula(ws, r, 2, f"=B{VAL_ROWS['equity']}/{C('is.ni_t1')}", NUM_X)


# =============================================================================
# 22. CHECKS
# =============================================================================

def build_checks(ws):
    ws.column_dimensions["A"].width = 36
    for c in range(2, 10):
        ws.column_dimensions[get_column_letter(c)].width = 13

    ws["A1"] = "MODEL INTEGRITY CHECKS"
    ws["A1"].font = section_font
    ws["A2"] = "All checks should evaluate to PASS. Tolerance is 0.01 (SAR 10 million)."
    ws["A2"].font = sub_font

    year_header(ws, 4, 0, 5)

    r = 6
    # 1. BS balance check
    write_label(ws, r, "1. Balance sheet balances (A − L − E)")
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        check = f"=IF(ABS({SH_BS}!{col}{BS_ROWS['check']})<0.01,\"PASS\",\"FAIL: \"&TEXT({SH_BS}!{col}{BS_ROWS['check']},\"+0.00;-0.00\"))"
        c = ws.cell(row=r, column=2 + t, value=check)
        c.alignment = right
        c.font = formula_font
    r += 1

    # 2. Cash on BS = closing cash on CFS rollforward
    write_label(ws, r, "2. BS cash = CFS rollforward")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        check = (f"=IF(ABS({SH_BS}!{col}{BS_ROWS['cash']}-{SH_CFS}!{col}{CFS_ROWS['close_cash']})<0.01,"
                 f"\"PASS\",\"FAIL\")")
        c = ws.cell(row=r, column=2 + t, value=check)
        c.alignment = right
        c.font = formula_font
    r += 1

    # 3. Equity rollforward = BS equity
    write_label(ws, r, "3. Equity rollforward = BS equity")
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        check = (f"=IF(ABS({SH_EQ}!{col}{EQ_ROWS['total']}-{SH_BS}!{col}{BS_ROWS['total_equity']})<0.01,"
                 f"\"PASS\",\"FAIL\")")
        c = ws.cell(row=r, column=2 + t, value=check)
        c.alignment = right
        c.font = formula_font
    r += 1

    # 4. Cash positivity
    write_label(ws, r, "4. Cash balance ≥ 0")
    for t in range(0, 6):
        col = get_column_letter(2 + t)
        check = f"=IF({SH_BS}!{col}{BS_ROWS['cash']}>=0,\"PASS\",\"FAIL (negative)\")"
        c = ws.cell(row=r, column=2 + t, value=check)
        c.alignment = right
        c.font = formula_font
    r += 1

    # 5. Net income > 0 in all years (just a sanity)
    write_label(ws, r, "5. Net income positive")
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        check = f"=IF({C(f'is.ni_t{t}')}>0,\"PASS\",\"WARNING\")"
        c = ws.cell(row=r, column=2 + t, value=check)
        c.alignment = right
        c.font = formula_font
    r += 1

    # Apply conditional formatting via openpyxl's CellIsRule
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
    pass_dxf = pass_fill
    fail_dxf = fail_fill
    # Range covering all check cells: B6:G10
    rng = f"B6:G{r-1}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal",
                                                  formula=['"PASS"'],
                                                  fill=pass_fill,
                                                  font=pass_font))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(B6<>"PASS",B6<>"")'],
                                                   fill=fail_fill,
                                                   font=fail_font))


# =============================================================================
# 23. MAIN
# =============================================================================

def main():
    wb = Workbook()

    cover = wb.active
    cover.title = SH_COVER
    build_cover(cover)

    # Pre-register IS rows so any downstream sheet (PPE, WC, Tax) can reference them.
    # build_is_v2 verifies that actual rows match this pre-registration.
    PRE_IS_LAYOUT = {
        "revenue":  5,
        "cogs":     8,
        "network":  9,
        "emp":     10,
        "cust_acq":11,
        "other":   12,
        "ebitda":  14,
        "da":      16,
        "ebit":    18,
        "fin_inc": 20,
        "interest":21,
        "pretax":  23,
        "tax":     24,
        "ni":      26,
        "ebitda_margin": 29,
        "ebit_margin":   30,
        "net_margin":    31,
    }
    for key, row in PRE_IS_LAYOUT.items():
        for t in range(1, 6):
            col = get_column_letter(2 + t)
            reg(f"is.{key}_t{t}", SH_IS, col, row)

    # Build order matters: shared registries get populated as we go.
    build_assumptions(wb.create_sheet(SH_ASSUM))
    build_drivers(wb.create_sheet(SH_DRIV))

    # Build downstream schedules first so IS can reference them.
    build_debt(wb.create_sheet(SH_DEBT))
    build_ppe(wb.create_sheet(SH_PPE))
    build_wc(wb.create_sheet(SH_WC))
    build_tax(wb.create_sheet(SH_TAX))

    # IS — references Tax/Debt/PPE rows now that they exist.
    is_ws = wb.create_sheet(SH_IS)
    build_is_v2(is_ws)
    # Sanity: verify actual IS rows match the pre-registered layout.
    for key, predicted in PRE_IS_LAYOUT.items():
        actual = IS_ROWS.get(key)
        if actual is not None and actual != predicted:
            raise RuntimeError(f"IS layout drift: '{key}' predicted row {predicted}, "
                               f"actual row {actual}. Update PRE_IS_LAYOUT.")

    build_equity(wb.create_sheet(SH_EQ))
    build_bs(wb.create_sheet(SH_BS))
    build_cfs(wb.create_sheet(SH_CFS))

    # Patch the BS cash cell that the Debt sheet references.
    # We wrote it as BS!<col>{BS_PLACEHOLDER_CASH}, but at write-time the placeholder
    # was the row offset BEFORE BS was built. Now we need to fix it to BS_ROWS["cash"].
    # We need to rewrite Debt's cash_offset row formulas with the correct cash row.
    debt_ws = wb[SH_DEBT]
    target_cash_row = BS_ROWS["cash"]
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        debt_ws.cell(row=DEBT_ROWS["cash_offset"], column=2 + t,
                     value=f"=-BS!{col}{target_cash_row}").font = formula_font

    # Patch BS cash formula that references CFS!{col}{CFS_PLACEHOLDER_NETCASH}.
    # When BS was built, CFS rows weren't known yet. Now CFS_ROWS["netcash"] is known.
    bs_ws = wb[SH_BS]
    target_netcash_row = CFS_ROWS["netcash"]
    for t in range(1, 6):
        col = get_column_letter(2 + t)
        prev = get_column_letter(2 + t - 1)
        bs_ws.cell(row=BS_ROWS["cash"], column=2 + t,
                   value=f"={prev}{BS_ROWS['cash']}+CFS!{col}{target_netcash_row}").font = formula_font

    build_budget(wb.create_sheet(SH_BUD))
    build_scenarios(wb.create_sheet(SH_SCEN))
    build_sensitivity(wb.create_sheet(SH_SENS))
    build_valuation(wb.create_sheet(SH_VAL))
    build_checks(wb.create_sheet(SH_CHK))

    wb.active = 0
    out = "telecom-model.xlsx"
    wb.save(out)
    print(f"Wrote {out}")
    print(f"Sheets: {wb.sheetnames}")


if __name__ == "__main__":
    main()
