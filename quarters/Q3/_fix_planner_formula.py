"""Fix the Q4_Planner operating-capacity formula and reflect the full web/Rio upside.

The original pad used 1 + (1 - productivity), which understates the requirement.
Correct relationship: effective units = OC * productivity * days, so
    required OC/day = forecast units / (productivity * days).
At 74% productivity the wrong pad gives 1.26x when the right factor is 1/0.74 = 1.35x.
"""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

DST = "quarters/Q3/Q3Data.xlsx"
wb = load_workbook(DST)
ws = wb["Q4_Planner"]
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
yellow = PatternFill("solid", fgColor="FFEB9C")
good = PatternFill("solid", fgColor="C6EFCE")
bad = PatternFill("solid", fgColor="FFC7CE")

ws["A2"] = ("Rule: required OC/day = forecast units / (productivity x 65 days). "
            "Never schedule below the forecast again. Q3 proof: OC 8/day vs 24 owned -> 204 stock-outs "
            "AND labour cost/unit up $130 -> $205.")
ws["A2"].alignment = Alignment(wrap_text=True)

# B9 was a pad factor of 1+(1-p); replace with the correct 1/p and fix B10/B11.
ws["A9"] = "Capacity factor (1 / productivity)"
ws["B9"] = "=1/B8"
ws["C9"] = "CORRECTED: was 1+(1-p)=1.26, which understated the requirement. 1/0.74 = 1.35."
ws["C9"].alignment = Alignment(wrap_text=True)
ws["B9"].fill = good

ws["B10"] = "=B7/(B8*65)"
ws["C10"] = "units / (productivity x days) - schedule AT LEAST this"
ws["B11"] = "=B10*65"

# Scenario block: what the full Q4 plan does to the capacity requirement.
r = 46
rows = [
    ["DEMAND SCENARIOS -> REQUIRED OPERATING CAPACITY", "Demand", "Required OC/day", "Fits in 24/day?"],
    ["A. Do nothing (Q3 demand less 16.3% ill will)", 525, "=B48/(0.74*65)", "=IF(C48<=24,\"yes\",\"NO\")"],
    ["B. Base forecast (planner B7)", "=B7", "=B49/(0.74*65)", "=IF(C49<=24,\"yes\",\"NO\")"],
    ["C. + fill Rio to the 7-person cap", "=B49+200", "=B50/(0.74*65)", "=IF(C50<=24,\"yes\",\"NO\")"],
    ["D. + full web rebuild (7 staff, 4 tactics)", "=B50+315", "=B51/(0.74*65)", "=IF(C51<=24,\"yes\",\"NO\")"],
    ["E. + New York store", "=B51+400", "=B52/(0.74*65)", "=IF(C52<=24,\"yes\",\"NO\")"],
    ["Max units at 24/day and 74% productivity", "=24*0.74*65", None, None],
    ["Read", None, None,
     "Scenarios C and D fit inside our existing 3 printers. Scenario E (adding New York on top) does NOT - "
     "it would need roughly 28/day. If we open New York we must either buy a 4th printer or accept that "
     "web/Rio growth is the cheaper path this quarter."],
]
for i, row in enumerate(rows):
    for c, v in enumerate(row, start=1):
        cell = ws.cell(r + i, c, v)
        if i == 0:
            cell.fill = hdr
            cell.font = bold
for i in range(1, 6):
    ws.cell(r + i, 3).number_format = "0.0"
ws.cell(r + 6, 2).fill = good
ws.cell(r + 7, 4).alignment = Alignment(wrap_text=True)
ws.cell(r + 7, 4).fill = yellow
ws.cell(r + 5, 4).fill = bad

# Web + Rio benchmarks appended to the benchmark block
r2 = 56
rows = [
    ["WEB & STAFFING TARGETS (from Q3_Web_Ops / Q3_Staffing)", "WeBike Q3", "Industry", "Q4 target"],
    ["Web staff (cap is 7)", 3, "7 (four firms)", 7],
    ["Web productivity tactics funded (of 4)", 1, 4, 4],
    ["Web productivity budget", 6000, "24,000 - 33,000", 28000],
    ["Web demand", 120, "435 - 569", 435],
    ["Rio store staff (cap is 7)", 4, "-", 7],
    ["Amsterdam store staff (cap is 7)", 7, "7", "7 (AT CAP - no room)"],
    ["Untrained staff", 2, 0, 0],
]
for i, row in enumerate(rows):
    for c, v in enumerate(row, start=1):
        cell = ws.cell(r2 + i, c, v)
        if i == 0:
            cell.fill = hdr
            cell.font = bold
        elif c == 4:
            cell.fill = yellow
ws.cell(r2 + 7, 4).fill = bad

wb.save(DST)
print("Q4_Planner: OC formula corrected to units/(productivity*days); scenarios A-E added")
