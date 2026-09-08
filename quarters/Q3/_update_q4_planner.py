"""Update Q4_Planner with graph-informed assumptions (productivity ceiling, unit-cost upside)."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

DST = "quarters/Q3/Q3Data.xlsx"
wb = load_workbook(DST)
ws = wb["Q4_Planner"]
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
yellow = PatternFill("solid", fgColor="FFEB9C")
good = PatternFill("solid", fgColor="C6EFCE")

# Corrected productivity assumption: industry ceiling is ~73%, not 85%.
ws["B8"] = 0.74
ws["C8"] = "Graph-checked: industry runs 68-73%. Assume 74% WITH a pay raise. Never plan on 85%."
ws["C6"] = "Q3 demand growth was +235% (highest in industry) on 2 stores; 1.25x is conservative"

ws["A2"] = ("Rule: OC/day = forecast demand / 65 * (1 + (1 - productivity)). Never schedule below forecast again. "
            "Q3 proof: OC 8/day vs 24 owned -> 204 stock-outs AND labor cost/unit up $130->$205.")
ws["A2"].alignment = Alignment(wrap_text=True)

r = 28
rows = [
    ["INDUSTRY BENCHMARKS FOR Q4 (from Strategic Graphs)", "WeBike Q3", "Industry", "Q4 target"],
    ["Scheduled operating capacity [units/day]", 8, "14 - 30", 15],
    ["Fixed capacity [units/day]", 24, "16 - 32", 24],
    ["Overtime capacity [units/day]", 1, "0 - 7", "0 - 2"],
    ["Average COGS per unit [$]", 620, "BB LLC lowest", "< 550"],
    ["Average labor cost per unit [$]", 205, None, "< 150"],
    ["Sales compensation [$]", 24425, "22,000 - 27,000", 26500],
    ["Production compensation [$]", 20757, "20,500 - 23,500", 22500],
    ["Demand per sales person", 45, "60 typical / 72 best", 60],
    ["Store + web center spend [$]", 121000, "up to 360,000", "raise — open city #3"],
    ["Organic SEM clicks", 150, "440 best", 300],
    ["Number of stores", 2, "2 (all but Bike Bros)", 3],
]
for i, row in enumerate(rows):
    for c, v in enumerate(row, start=1):
        cell = ws.cell(r + i, c, v)
        if i == 0:
            cell.fill = hdr
            cell.font = bold
        elif c == 4:
            cell.fill = yellow

r2 = r + len(rows) + 2
notes = [
    ["UNIT-COST UPSIDE — why right-sizing OC pays twice", None],
    ["Q3 COGS per unit", 620],
    ["Target COGS per unit at ~15/day scale", 550],
    ["Saving per unit", "=B{0}-B{1}".format(r2 + 1, r2 + 2)],
    ["Q4 forecast units", "=B7"],
    ["COGS saving on forecast volume", "=B{0}*B{1}".format(r2 + 3, r2 + 4)],
    ["+ Gross margin recovered from ending stock-outs", "=Q3_Unit_Economics!I6"],
    ["= Combined Q4 swing vs Q3 behaviour", "=B{0}+B{1}".format(r2 + 5, r2 + 6)],
]
for i, row in enumerate(notes):
    for c, v in enumerate(row, start=1):
        cell = ws.cell(r2 + i, c, v)
        if i == 0:
            cell.font = bold
ws.cell(r2 + 7, 2).fill = good

ws.column_dimensions["A"].width = 46
ws.column_dimensions["B"].width = 16
ws.column_dimensions["C"].width = 26
ws.column_dimensions["D"].width = 24

wb.save(DST)
print("Q4_Planner updated with industry benchmarks and unit-cost upside")
