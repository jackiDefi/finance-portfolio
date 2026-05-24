"""
Build telecom-model.xlsx — a formula-driven mirror of the interactive
operating model on the Telecom & IT deep-dive page.

Run from this directory:
    python3 build_excel.py

Convention:
    - Inputs:   blue font (0563C1)
    - Formulas: black
    - Outputs:  bold
"""

from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ---------- Style primitives -------------------------------------------------

BLUE = "0563C1"
NAVY = "0A1828"
COPPER = "B87333"
LIGHT_BG = "F1EFE6"

input_font   = Font(name="Calibri", size=11, color=BLUE, bold=False)
formula_font = Font(name="Calibri", size=11, color="000000", bold=False)
output_font  = Font(name="Calibri", size=11, color="000000", bold=True)
header_font  = Font(name="Calibri", size=11, color="FFFFFF", bold=True)
title_font   = Font(name="Calibri", size=18, color=NAVY, bold=True)
sub_font     = Font(name="Calibri", size=10, color="5A6573", italic=True)
section_font = Font(name="Calibri", size=10, color=COPPER, bold=True)

header_fill = PatternFill("solid", fgColor=NAVY)
band_fill   = PatternFill("solid", fgColor=LIGHT_BG)

thin = Side(border_style="thin", color="D4D2C8")
bottom_border = Border(bottom=thin)
top_bottom    = Border(top=thin, bottom=thin)

right = Alignment(horizontal="right")
left  = Alignment(horizontal="left")
center = Alignment(horizontal="center")

NUM_BN = '#,##0.0;[Red]-#,##0.0'
NUM_PCT = '0.0%;[Red]-0.0%'
NUM_X   = '0.00"x"'
NUM_INT = '#,##0'


# ---------- Default driver values (mirror of the JS DEFAULTS) ---------------

DEFAULTS = {
    "rev0":    75.0,   # SAR bn
    "growth":  0.06,
    "ebitda":  0.36,
    "da":      0.16,
    "capex":   0.18,
    "tax":     0.15,
    "wacc":    0.09,
    "g":       0.03,
    "netdebt": 15.0,   # SAR bn
    "shares":  5000.0, # millions
    "kd":      0.05,   # cost of debt for interest
    "wc_pct":  0.02,   # WC as % of incremental revenue
}


def main() -> None:
    wb = Workbook()

    cover = wb.active
    cover.title = "Cover"
    build_cover(cover)

    inputs = wb.create_sheet("Inputs")
    build_inputs(inputs)

    forecast = wb.create_sheet("Forecast")
    build_forecast(forecast)

    valuation = wb.create_sheet("Valuation")
    build_valuation(valuation)

    wb.active = 0
    out = "telecom-model.xlsx"
    wb.save(out)
    print(f"Wrote {out}")


# ---------- Cover -----------------------------------------------------------

def build_cover(ws) -> None:
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 36

    ws["B2"] = "Telecom & IT — Operating Model"
    ws["B2"].font = title_font

    ws["B3"] = f"Author: Togrul Mirzayev   •   Generated: {date.today().isoformat()}"
    ws["B3"].font = sub_font

    ws["B5"] = "PORTFOLIO"
    ws["B5"].font = section_font
    ws["B6"] = "https://jackidefi.github.io/finance-portfolio/"
    ws["B6"].font = Font(name="Calibri", size=11, color="000000", underline="single")

    ws["B8"] = "STRUCTURE"
    ws["B8"].font = section_font
    structure = [
        ("Inputs",    "All drivers in one place. Blue cells are hard-coded; change anything here."),
        ("Forecast",  "5-year P&L and FCF build, fully formula-driven from Inputs."),
        ("Valuation", "DCF, terminal value, EV, equity, per-share."),
    ]
    for i, (name, desc) in enumerate(structure):
        ws.cell(row=9 + i, column=2, value=name).font = output_font
        ws.cell(row=9 + i, column=3, value=desc).font = formula_font

    ws["B13"] = "COLOUR LEGEND"
    ws["B13"].font = section_font
    ws.cell(row=14, column=2, value="Blue text").font = input_font
    ws.cell(row=14, column=3, value="Hard-coded input — change freely").font = formula_font
    ws.cell(row=15, column=2, value="Black text").font = formula_font
    ws.cell(row=15, column=3, value="Formula referencing other cells").font = formula_font
    ws.cell(row=16, column=2, value="Bold").font = output_font
    ws.cell(row=16, column=3, value="Subtotal or model output").font = formula_font

    ws["B18"] = "NOTES"
    ws["B18"].font = section_font
    notes = [
        "Default values mirror the interactive model on the website and roughly approximate STC (Tadawul: 7010).",
        "Interest expense is simplified: cost of debt × Year 0 net debt, held flat across all five years.",
        "Working capital absorbs 2% of incremental revenue — used as a simplifying placeholder.",
        "Terminal value uses Gordon growth: FCFF_Y5 × (1+g) ÷ (WACC − g).",
        "FCFF = NOPAT + D&A − Capex − ΔWC (pre-financing).",
    ]
    for i, n in enumerate(notes):
        ws.cell(row=19 + i, column=2, value=f"•  {n}").font = formula_font
        ws.merge_cells(start_row=19 + i, start_column=2, end_row=19 + i, end_column=4)


# ---------- Inputs ----------------------------------------------------------

# We define a stable layout so Forecast and Valuation can reference these cells.
INPUT_CELLS = {
    "rev0":    "B5",
    "growth":  "B6",
    "ebitda":  "B7",
    "da":      "B8",
    "capex":   "B9",
    "tax":     "B12",
    "wacc":    "B13",
    "g":       "B14",
    "netdebt": "B17",
    "shares":  "B18",
    "kd":      "B19",
    "wc_pct":  "B20",
}


def build_inputs(ws) -> None:
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 50

    ws["A1"] = "INPUTS"
    ws["A1"].font = section_font

    rows = [
        (3,  "Operating drivers",                None,                NUM_BN, None),
        (5,  "Service revenue Y0 (SAR bn)",      DEFAULTS["rev0"],    NUM_BN, "Year 0 baseline; bump for larger/smaller operators"),
        (6,  "Revenue growth (annual)",          DEFAULTS["growth"],  NUM_PCT, "Applied to every forecast year"),
        (7,  "EBITDA margin",                    DEFAULTS["ebitda"],  NUM_PCT, "Post-IFRS 16; held flat across the explicit period"),
        (8,  "D&A / revenue",                    DEFAULTS["da"],      NUM_PCT, "Reflects asset base; held flat"),
        (9,  "Capex intensity",                  DEFAULTS["capex"],   NUM_PCT, "Capex ÷ revenue"),

        (11, "Financial drivers",                None,                NUM_BN, None),
        (12, "Tax / zakat rate",                 DEFAULTS["tax"],     NUM_PCT, "KSA zakat + corporate-tax blend; simplified"),
        (13, "WACC",                             DEFAULTS["wacc"],    NUM_PCT, "Discount rate for DCF"),
        (14, "Terminal growth",                  DEFAULTS["g"],       NUM_PCT, "Long-run nominal growth in TV; must be < WACC"),

        (16, "Balance sheet",                    None,                NUM_BN, None),
        (17, "Net debt Y0 (SAR bn)",             DEFAULTS["netdebt"], NUM_BN, "Subtracted from EV to get equity value"),
        (18, "Shares outstanding (millions)",    DEFAULTS["shares"],  NUM_INT, "Used for per-share value"),
        (19, "Cost of debt (interest)",          DEFAULTS["kd"],      NUM_PCT, "Applied to net debt for interest expense"),
        (20, "WC as % of ΔRevenue",              DEFAULTS["wc_pct"],  NUM_PCT, "Simplifying assumption"),
    ]

    for row, label, value, fmt, note in rows:
        ws.cell(row=row, column=1, value=label)
        if value is None:
            ws.cell(row=row, column=1).font = section_font
        else:
            ws.cell(row=row, column=1).font = formula_font
            c = ws.cell(row=row, column=2, value=value)
            c.font = input_font
            c.alignment = right
            c.number_format = fmt
            if note:
                ws.cell(row=row, column=3, value=note).font = sub_font


# ---------- Forecast --------------------------------------------------------

# Forecast row layout
F_ROW = {
    "rev":    5,
    "ebitda": 6,
    "da":     7,
    "ebit":   8,
    "int":    9,
    "pretax": 10,
    "tax":    11,
    "ni":     12,
    "capex":  14,
    "wc":     15,
    "fcff":   16,
    "nopat":  18,  # used by Valuation sheet, kept here for clarity
}


def ic(key: str) -> str:
    """Absolute reference to an Inputs cell, prefixed with sheet name."""
    return f"Inputs!${INPUT_CELLS[key][0]}${INPUT_CELLS[key][1:]}"


def build_forecast(ws) -> None:
    ws.column_dimensions["A"].width = 28
    for col in range(2, 8):
        ws.column_dimensions[get_column_letter(col)].width = 13

    ws["A1"] = "5-YEAR FORECAST"
    ws["A1"].font = section_font
    ws["A2"] = "All figures in SAR billion unless noted."
    ws["A2"].font = sub_font

    # Header row (year labels)
    ws.cell(row=4, column=1, value="").fill = header_fill
    for i, y in enumerate(["Y1", "Y2", "Y3", "Y4", "Y5"]):
        c = ws.cell(row=4, column=2 + i, value=y)
        c.font = header_font
        c.fill = header_fill
        c.alignment = right

    # Revenue: Y0 * (1+g)^t
    for i in range(5):
        col = 2 + i
        cell = ws.cell(row=F_ROW["rev"], column=col,
                       value=f"={ic('rev0')}*(1+{ic('growth')})^{i + 1}")
        cell.number_format = NUM_BN
        cell.font = formula_font
        cell.alignment = right
    ws.cell(row=F_ROW["rev"], column=1, value="Service revenue").font = formula_font

    # EBITDA = revenue * margin
    label_row(ws, "EBITDA",     F_ROW["ebitda"], lambda c: f"={ws.cell(row=F_ROW['rev'], column=c).coordinate}*{ic('ebitda')}", NUM_BN, bold=True)

    # D&A = revenue * da%
    label_row(ws, "D&A",        F_ROW["da"], lambda c: f"={ws.cell(row=F_ROW['rev'], column=c).coordinate}*{ic('da')}", NUM_BN)

    # EBIT = EBITDA - D&A
    label_row(ws, "EBIT",       F_ROW["ebit"], lambda c: f"={ws.cell(row=F_ROW['ebitda'], column=c).coordinate}-{ws.cell(row=F_ROW['da'], column=c).coordinate}", NUM_BN, bold=True)

    # Interest = net debt * cost of debt (flat across all years)
    label_row(ws, "Interest",   F_ROW["int"], lambda c: f"={ic('netdebt')}*{ic('kd')}", NUM_BN)

    # Pre-tax = EBIT - interest
    label_row(ws, "Pre-tax income", F_ROW["pretax"], lambda c: f"={ws.cell(row=F_ROW['ebit'], column=c).coordinate}-{ws.cell(row=F_ROW['int'], column=c).coordinate}", NUM_BN)

    # Tax = max(0, pretax * tax rate)
    label_row(ws, "Tax / zakat", F_ROW["tax"], lambda c: f"=MAX(0,{ws.cell(row=F_ROW['pretax'], column=c).coordinate}*{ic('tax')})", NUM_BN)

    # Net income
    label_row(ws, "Net income", F_ROW["ni"], lambda c: f"={ws.cell(row=F_ROW['pretax'], column=c).coordinate}-{ws.cell(row=F_ROW['tax'], column=c).coordinate}", NUM_BN, bold=True)

    # spacer

    # Capex
    label_row(ws, "Capex",      F_ROW["capex"], lambda c: f"={ws.cell(row=F_ROW['rev'], column=c).coordinate}*{ic('capex')}", NUM_BN)

    # ΔWC = (rev_t - rev_t-1) * wc_pct; for Y1 use rev_Y0 from inputs
    def wc_formula(col):
        rev_t = ws.cell(row=F_ROW["rev"], column=col).coordinate
        if col == 2:
            rev_prev = ic("rev0")
        else:
            rev_prev = ws.cell(row=F_ROW["rev"], column=col - 1).coordinate
        return f"=({rev_t}-{rev_prev})*{ic('wc_pct')}"
    label_row(ws, "ΔWorking capital", F_ROW["wc"], wc_formula, NUM_BN)

    # FCFF = NOPAT + D&A - capex - dWC. NOPAT = EBIT * (1 - tax)
    def fcff_formula(col):
        ebit = ws.cell(row=F_ROW["ebit"],  column=col).coordinate
        da   = ws.cell(row=F_ROW["da"],    column=col).coordinate
        cx   = ws.cell(row=F_ROW["capex"], column=col).coordinate
        wc   = ws.cell(row=F_ROW["wc"],    column=col).coordinate
        return f"={ebit}*(1-{ic('tax')})+{da}-{cx}-{wc}"
    label_row(ws, "FCFF", F_ROW["fcff"], fcff_formula, NUM_BN, bold=True, banded=True)

    # NOPAT (used by valuation sheet) — separate line for clarity, but not strictly required
    label_row(ws, "Memo: NOPAT", F_ROW["nopat"], lambda c: f"={ws.cell(row=F_ROW['ebit'], column=c).coordinate}*(1-{ic('tax')})", NUM_BN)

    # Ratios block
    ws.cell(row=20, column=1, value="KEY RATIOS").font = section_font

    label_row(ws, "EBITDA margin",     21, lambda c: f"={ws.cell(row=F_ROW['ebitda'], column=c).coordinate}/{ws.cell(row=F_ROW['rev'], column=c).coordinate}", NUM_PCT)
    label_row(ws, "FCF / revenue",     22, lambda c: f"={ws.cell(row=F_ROW['fcff'],   column=c).coordinate}/{ws.cell(row=F_ROW['rev'], column=c).coordinate}", NUM_PCT)
    label_row(ws, "Capex / revenue",   23, lambda c: f"={ws.cell(row=F_ROW['capex'],  column=c).coordinate}/{ws.cell(row=F_ROW['rev'], column=c).coordinate}", NUM_PCT)
    label_row(ws, "Net debt / EBITDA", 24, lambda c: f"={ic('netdebt')}/{ws.cell(row=F_ROW['ebitda'], column=c).coordinate}", NUM_X)


def label_row(ws, label: str, row: int, formula_for_col, num_fmt: str,
              bold: bool = False, banded: bool = False) -> None:
    """Write a label cell + 5 formula cells (Y1..Y5)."""
    lbl = ws.cell(row=row, column=1, value=label)
    lbl.font = output_font if bold else formula_font
    if banded:
        lbl.fill = band_fill
    for i in range(5):
        col = 2 + i
        c = ws.cell(row=row, column=col, value=formula_for_col(col))
        c.font = output_font if bold else formula_font
        c.number_format = num_fmt
        c.alignment = right
        if banded:
            c.fill = band_fill


# ---------- Valuation -------------------------------------------------------

def build_valuation(ws) -> None:
    ws.column_dimensions["A"].width = 32
    for col in range(2, 8):
        ws.column_dimensions[get_column_letter(col)].width = 13

    ws["A1"] = "DCF VALUATION"
    ws["A1"].font = section_font
    ws["A2"] = "WACC, terminal growth, net debt, and shares all live on the Inputs sheet."
    ws["A2"].font = sub_font

    # Header
    ws.cell(row=4, column=1, value="").fill = header_fill
    for i, y in enumerate(["Y1", "Y2", "Y3", "Y4", "Y5"]):
        c = ws.cell(row=4, column=2 + i, value=y)
        c.font = header_font
        c.fill = header_fill
        c.alignment = right

    # FCFF pulled from Forecast sheet
    ws.cell(row=5, column=1, value="FCFF (from Forecast)").font = formula_font
    for i in range(5):
        col = 2 + i
        ref = f"Forecast!{get_column_letter(col)}{F_ROW['fcff']}"
        c = ws.cell(row=5, column=col, value=f"={ref}")
        c.font = formula_font
        c.number_format = NUM_BN
        c.alignment = right

    # Discount factor 1/(1+WACC)^t
    ws.cell(row=6, column=1, value="Discount factor").font = formula_font
    for i in range(5):
        col = 2 + i
        c = ws.cell(row=6, column=col, value=f"=1/(1+{ic('wacc')})^{i + 1}")
        c.font = formula_font
        c.number_format = '0.0000'
        c.alignment = right

    # PV of FCFF
    ws.cell(row=7, column=1, value="PV of FCFF").font = output_font
    for i in range(5):
        col = 2 + i
        fc = ws.cell(row=5, column=col).coordinate
        df = ws.cell(row=6, column=col).coordinate
        c = ws.cell(row=7, column=col, value=f"={fc}*{df}")
        c.font = output_font
        c.number_format = NUM_BN
        c.alignment = right

    # Terminal value and outputs (single column on the right)
    ws.cell(row=9,  column=1, value="Sum PV (explicit Y1-Y5)").font = output_font
    ws.cell(row=9,  column=2, value="=SUM(B7:F7)").font = output_font
    ws.cell(row=9,  column=2).number_format = NUM_BN
    ws.cell(row=9,  column=2).alignment = right

    ws.cell(row=10, column=1, value="Terminal value (Gordon)").font = formula_font
    ws.cell(row=10, column=2, value=f"=F5*(1+{ic('g')})/({ic('wacc')}-{ic('g')})").font = formula_font
    ws.cell(row=10, column=2).number_format = NUM_BN
    ws.cell(row=10, column=2).alignment = right

    ws.cell(row=11, column=1, value="PV of terminal value").font = output_font
    ws.cell(row=11, column=2, value="=B10*F6").font = output_font
    ws.cell(row=11, column=2).number_format = NUM_BN
    ws.cell(row=11, column=2).alignment = right

    ws.cell(row=13, column=1, value="Enterprise value").font = output_font
    ws.cell(row=13, column=2, value="=B9+B11").font = output_font
    ws.cell(row=13, column=2).number_format = NUM_BN
    ws.cell(row=13, column=2).fill = band_fill
    ws.cell(row=13, column=1).fill = band_fill
    ws.cell(row=13, column=2).alignment = right

    ws.cell(row=14, column=1, value="Less: net debt").font = formula_font
    ws.cell(row=14, column=2, value=f"={ic('netdebt')}").font = formula_font
    ws.cell(row=14, column=2).number_format = NUM_BN
    ws.cell(row=14, column=2).alignment = right

    ws.cell(row=15, column=1, value="Equity value").font = output_font
    ws.cell(row=15, column=2, value="=B13-B14").font = output_font
    ws.cell(row=15, column=2).number_format = NUM_BN
    ws.cell(row=15, column=2).fill = band_fill
    ws.cell(row=15, column=1).fill = band_fill
    ws.cell(row=15, column=2).alignment = right

    ws.cell(row=16, column=1, value="Per-share value (SAR)").font = output_font
    ws.cell(row=16, column=2, value=f"=B15*1000/{ic('shares')}").font = output_font
    ws.cell(row=16, column=2).number_format = '0.00'
    ws.cell(row=16, column=2).alignment = right

    ws.cell(row=18, column=1, value="Implied FY+1 EV/EBITDA").font = formula_font
    ws.cell(row=18, column=2, value=f"=B13/Forecast!B{F_ROW['ebitda']}").font = formula_font
    ws.cell(row=18, column=2).number_format = NUM_X
    ws.cell(row=18, column=2).alignment = right


if __name__ == "__main__":
    main()
