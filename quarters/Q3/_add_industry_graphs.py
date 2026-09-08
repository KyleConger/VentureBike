"""Add Q3 industry-graph intelligence (Strategic Graphs workspace) to Q3Data.xlsx.

Values marked approx are read off the Strategic Graphs charts and are estimates
(+/- one gridline). Exact figures come from the numeric reports and live in the
other Q3_* sheets.
"""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

DST = "quarters/Q3/Q3Data.xlsx"
wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
bad = PatternFill("solid", fgColor="FFC7CE")
good = PatternFill("solid", fgColor="C6EFCE")
yellow = PatternFill("solid", fgColor="FFEB9C")

for name in ("Q3_Industry_Graphs", "Q3_Graph_Reads"):
    if name in wb.sheetnames:
        del wb[name]


def put(ws, start, rows, header_row=True):
    r = start
    for i, row in enumerate(rows):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(r, c, v)
            if i == 0 and header_row:
                cell.fill = hdr
                cell.font = bold
        r += 1
    return r


ws = wb.create_sheet("Q3_Industry_Graphs")
ws["A1"] = "Q3 Strategic Graphs — industry comparison (WeBike vs 6 rivals)"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("Source: Strategic Graphs > Workspace, Quarter 4 view (shows Q0-Q3 actuals). "
            "Chart-read values are approximate; numeric report values live in Q3_BSC / Q3_Market / Q3_Financials.")
ws["A2"].alignment = Alignment(wrap_text=True)

# ---- Manufacturing: the smoking gun ----
r = put(ws, 4, [
    ["MANUFACTURING — capacity (units/day)", "Q2", "Q3", "Note"],
    ["WeBike fixed capacity", 24, 24, "3 printers; tied 2nd highest in industry"],
    ["WeBike SCHEDULED operating capacity", 20, 8, "CONFIRMS the back-solve: we cut OC by 60%"],
    ["WeBike overtime capacity", 0, 1, "ran OT on top of a tiny base"],
    ["SpaceBikes operating capacity", 20, 30, "leader — expanded fixed capacity to 32"],
    ["MILC Bikes operating capacity", 16, 24, None],
    ["Bike Bros operating capacity", 10, 15, None],
    ["LiteCycle operating capacity", 8, 15, None],
    ["BB LLC operating capacity", 6, 15, "+ ~7 overtime — highest OT in industry"],
    ["Spoke'd Up operating capacity", 5, 14, None],
    ["WeBike RANK on operating capacity", None, "LAST (7 of 7)", "Only firm that REDUCED operating capacity in Q3"],
])
ws.cell(6, 3).fill = bad
ws.cell(r - 1, 3).fill = bad

r = put(ws, r + 1, [
    ["MANUFACTURING — WeBike unit costs", "Q2", "Q3", "Read"],
    ["Capacity utilization [%]", 16, 73, "Q2 = idle plant; Q3 = tight plant, but tiny plant"],
    ["Average labor cost per unit [$]", 130, 205, "ROSE 58% — small OC + overtime is expensive per unit"],
    ["Average cost of goods sold per unit [$]", 520, 620, "ROSE ~19% despite selling 4x the units"],
    ["Production cost [$]", 100928, 272340, None],
    ["Stock-outs [units]", 0, 204, "the whipsaw"],
    ["Excess capacity cost [$]", 148652, 0, "the whipsaw, other direction"],
    ["Lesson", None, None, "Under-scheduling cost us BOTH revenue AND margin. Volume normally lowers unit cost; we broke that by running a sub-scale OC with overtime."],
])
ws.cell(r - 6, 3).fill = bad
ws.cell(r - 5, 3).fill = bad

# ---- Market ----
r = put(ws, r + 1, [
    ["MARKET — segment size (units)", "Q2", "Q3", "Growth", "Note"],
    ["Speed", 1500, 2878, "=C{0}/B{0}-1".format(r + 1), "LARGEST and fastest growing; all 7 firms compete"],
    ["Recreation", 1100, 2060, "=C{0}/B{0}-1".format(r + 2), "we do not compete"],
    ["Mountain", 1050, 1621, "=C{0}/B{0}-1".format(r + 3), "SMALLEST and slowest; our primary; only 4 firms"],
    ["Total market", 3700, 6559, "=C{0}/B{0}-1".format(r + 4), None],
])
ws.cell(r - 1, 5).value = "Strategic tension: our least-crowded segment is also the smallest and slowest-growing."

r = put(ws, r + 1, [
    ["MARKET — searches (web interest, Q3)", "Approx", "Note"],
    ["Recreation", 1350, "highest search volume — segment we ignore"],
    ["Mountain", 1030, None],
    ["Speed", 900, "large demand but lowest search volume"],
])

r = put(ws, r + 1, [
    ["MARKET — Q3 share by segment [%]", "Recreation", "Mountain", "Speed", "Overall", "Note"],
    ["BB LLC", 4.8, 46.3, 22.8, 23.0, "LEADER — all 3 segments, dominates Mountain"],
    ["SpaceBikes", 34.5, 0.0, 20.4, 19.8, "#2 — Rec + Speed only, no Mountain"],
    ["Spoke'd Up", 23.9, 0.0, 12.6, 13.0, None],
    ["LiteCycle", 4.7, 21.0, 19.0, 15.0, "our closest Mountain rival"],
    ["MILC Bikes", 22.9, 0.0, 11.4, 12.2, None],
    ["WeBike", 2.6, 22.1, 7.5, 9.6, "#2 Mountain, 6th Speed"],
    ["Bike Bros", 6.7, 10.6, 6.3, 7.5, "FALLING in every segment (Mtn 19->11)"],
])

r = put(ws, r + 1, [
    ["MARKET — demand productivity per head (Q3, approx)", "WeBike", "Industry best", "Industry typical", "Note"],
    ["Demand per sales person", 45, 72, 60, "LAST — every rival gets more demand per head"],
    ["Demand per store sales person", 46, 88, 70, "BB LLC ~88; our stores under-produce"],
    ["Demand per web sales center person", 42, 95, 75, "we CUT web staff in Q3 — wrong lever"],
    ["Total sales people", 14, 21, 14, "headcount is at parity; output per head is not"],
])
ws.cell(r - 4, 2).fill = bad

r = put(ws, r + 1, [
    ["MARKET — demand vs units sold (the leakage)", "Value", "Note"],
    ["WeBike share of total DEMAND [%]", 9.6, "what we created"],
    ["WeBike share of UNITS SOLD [%]", 7.3, "what we captured"],
    ["Leakage [pct points]", "=B{0}-B{1}".format(r - 2, r - 1), "pure stock-out loss, visible in the graphs"],
    ["Q3 demand growth rate [%]", 235, "highest in industry — demand generation is NOT our problem"],
])

# ---- Marketing inputs ----
r = put(ws, r + 1, [
    ["MARKETING INPUTS (Q3, approx)", "WeBike", "Industry high", "Rank", "Note"],
    ["Advertising expenses [$]", 119165, 145000, "3rd", "spend is competitive"],
    ["Total regional ads [count]", 19, 24, "2nd", "ad VOLUME is a strength"],
    ["Organic SEM clicks", 150, 440, "6th", "weak; BB LLC 440, SpaceBikes 400"],
    ["Organic SEM expenses [$]", 2000, 5000, "3rd", "Bike Bros spends 5k for 200 clicks — poor ROI benchmark"],
    ["Average price [$]", 1400, 1480, "3rd", "BB LLC prices HIGHER and still leads share"],
    ["Number of stores", 2, 2, "tied", "everyone but Bike Bros has 2 -> city #3 is a differentiator"],
    ["Number of web sales centers", 1, 1, "tied", None],
    ["Store + web center expenses [$]", 121000, 360000, "lowest", "we UNDER-invest in channel"],
])
ws.cell(r - 8, 2).fill = good
ws.cell(r - 7, 2).fill = good
ws.cell(r - 1, 2).fill = bad

# ---- HR: correction ----
r = put(ws, r + 1, [
    ["HUMAN RESOURCES (Q3, approx) — CORRECTS earlier read", "WeBike", "Industry range", "Note"],
    ["Production worker compensation [$]", 20757, "20,500 - 23,500", "low end, not an outlier"],
    ["Sales force compensation [$]", 24425, "22,000 - 27,000", "middle; MILC ~27k, BB LLC ~26.5k highest"],
    ["Production worker productivity [%]", 72, "68 - 73", "MID-PACK, not behind"],
    ["Sales force productivity [%]", 70, "70 - 75", "MID-PACK; BB LLC ~75 best"],
    ["Comp satisfaction", 68, "65 - 72", None],
    ["CORRECTION", None, None, "The whole industry runs ~70%. Our 85% projection was never realistic. "
     "The error was PLANNING CAPACITY on 85%, not being badly underpaid. Pay raises buy a few points, not 15."],
])
ws.cell(r - 1, 1).fill = yellow

# ---- BSC comparison ----
r = put(ws, r + 1, [
    ["BALANCED SCORECARD — Q3 by firm (approx from graph)", "Total Perf", "Financial", "Market", "Asset Mgmt", "Wealth"],
    ["BB LLC", 22.0, 37.7, 0.345, 0.920, 0.87],
    ["SpaceBikes", 11.0, 34.0, 0.270, 0.670, 0.95],
    ["MILC Bikes", 2.8, 15.0, 0.175, 0.480, 0.79],
    ["Bike Bros", 2.5, 14.0, 0.085, 0.460, 0.82],
    ["LiteCycle", 2.0, 14.5, 0.155, 0.530, 0.80],
    ["Spoke'd Up", 1.0, 6.0, 0.150, 0.420, 0.79],
    ["WeBike", 0.411, 5.319, 0.098, 0.353, 0.668],
])
for c in range(1, 7):
    ws.cell(r - 1, c).fill = bad

r = put(ws, r + 1, [
    ["CUMULATIVE BSC (approx)", "Value", "Note"],
    ["SpaceBikes", 6.3, "cumulative leader"],
    ["BB LLC", 6.2, "closing fast — Q3 single-quarter leader"],
    ["Bike Bros", 1.3, None],
    ["MILC Bikes", 1.3, None],
    ["Spoke'd Up", 0.7, None],
    ["LiteCycle", 0.6, None],
    ["WeBike", 0.033, "Q2 Manufacturing Productivity of 0.20 still poisons the cumulative product"],
])

# ---- Rival profiles ----
r = put(ws, r + 1, [
    ["RIVAL PROFILES — what to copy and what to avoid", "Read"],
    ["BB LLC (leader)", "All 3 segments. Highest price (~$1,480) AND highest share. Most regional ads (24). "
     "Asset Mgmt 0.920 — best capital efficiency. Runs heavy overtime (~7) on modest fixed capacity: "
     "buys flexibility instead of printers. 21 sales people, ~88 demand per store rep."],
    ["SpaceBikes (cumulative leader)", "Skips Mountain entirely; owns Recreation (34.5%) + Speed (20.4%). "
     "Expanded fixed capacity to 32/day — the only firm that out-built us. Highest ad spend (~$145k). Wealth 0.95."],
    ["LiteCycle", "Our Mountain rival at 21.0%, plus 19.0% of Speed. Organic SEM clicks COLLAPSED 400->180."],
    ["Bike Bros (cautionary tale)", "1 store, 0 web sales centers, 7 sales people. Share falling in all three "
     "segments (Mountain 19->11). Under-channeling is fatal — this is the risk of our 'cut web spend' instinct."],
    ["MILC / Spoke'd Up", "Mid-pack. Spoke'd Up has the highest market growth but lowest Marketing Effectiveness (0.66)."],
])
for cell_row in range(r - 5, r):
    ws.cell(cell_row, 2).alignment = Alignment(wrap_text=True)

# ---- Q4 implications ----
r = put(ws, r + 1, [
    ["Q4 IMPLICATIONS FROM THE GRAPHS", "Action"],
    ["1. Operating capacity", "We were LAST at 8/day while owning 24. Schedule to forecast (~15/day). "
     "This is confirmed by the Operating Capacity chart, not just inferred."],
    ["2. Unit cost", "Labor cost/unit rose $130->$205 and COGS/unit $520->$620 because OC was sub-scale. "
     "Right-sizing OC cuts unit cost AND ends stock-outs — one decision, two wins."],
    ["3. Do NOT cut channel", "Store+web spend is the LOWEST in industry and demand per head is LAST. "
     "Bike Bros shows where that path ends. Open city #3 and re-staff web."],
    ["4. Pay expectations", "Industry productivity is ~70% across the board. Raise pay to upper-quartile "
     "(sales ~$26-27k, production ~$22-23k) but PLAN on ~73-75%, not 85%."],
    ["5. Segment strategy", "Mountain is the smallest, slowest segment (1,621) and BB LLC holds 46%. "
     "Speed is 2,878 and growing. Recreation has the most searches (1,350) and we hold 2.6%."],
    ["6. Price", "BB LLC charges MORE than us and leads. Our ~$1,400 average is not the constraint — supply is."],
    ["7. Advertising", "Ad count 2nd, spend 3rd, Marketing Effectiveness at par. Marketing is NOT the problem. "
     "Weakest marketing input is organic SEM clicks (150 vs 440)."],
])
for cell_row in range(r - 7, r):
    ws.cell(cell_row, 2).alignment = Alignment(wrap_text=True)

ws.column_dimensions["A"].width = 46
for col in "BCDEF":
    ws.column_dimensions[col].width = 16
ws.column_dimensions["D"].width = 22
ws.column_dimensions["E"].width = 22
ws.column_dimensions["F"].width = 60

wb.save(DST)
print(f"Added Q3_Industry_Graphs to {DST}")
