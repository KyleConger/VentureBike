"""Add Q3 web sales/service staffing + web productivity budgets, and fix the Q4_Planner OC formula."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

DST = "quarters/Q3/Q3Data.xlsx"

# company: (web sales, web service, total, web demand)
WEB = {
    "BB LLC":     (5, 2, 7, 569),
    "SpaceBikes": (5, 2, 7, 491),
    "LiteCycle":  (5, 2, 7, 435),
    "MILC Bikes": (6, 1, 7, 435),
    "Spoke'd Up": (2, 1, 3, 138),
    "WeBike":     (2, 1, 3, 120),
    "Bike Bros":  (0, 0, 0,   0),
}

TACTICS = ["Toll-free phone advising / service calls",
           "Advanced shopping cart & checkout",
           "Continuous page upgrades (content, appeal, navigation)",
           "Order tracking software"]
# company -> budget per tactic
BUDGET = {
    "LiteCycle":  [7000, 7000,  7000, 8000],
    "WeBike":     [6000,    0,     0,    0],
    "BB LLC":     [9000, 7000,  9000, 8000],
    "Spoke'd Up": [3000,    0,  6000,    0],
    "SpaceBikes": [6000, 7000, 12000, 8000],
    "MILC Bikes": [3000, 7000,  6000, 8000],
}

print("WEB STAFF vs WEB DEMAND")
for co, (s, sv, t, d) in sorted(WEB.items(), key=lambda x: -x[1][3]):
    per = d / t if t else 0
    print(f"  {co:<12} staff {t}  ({s} sales / {sv} service)  demand {d:>4}  per head {per:5.1f}")
print(f"\n  MAX web staff observed: {max(v[2] for v in WEB.values())} -> 7 is the cap (4 firms sit exactly there)")
print(f"  WeBike at 3 -> 4 SLOTS OPEN\n")

print("WEB PRODUCTIVITY BUDGET vs WEB DEMAND")
for co in sorted(BUDGET, key=lambda x: -sum(BUDGET[x])):
    b = BUDGET[co]
    live = sum(1 for x in b if x > 0)
    print(f"  {co:<12} total ${sum(b):>6,}  tactics live {live}/4  web demand {WEB[co][3]:>4}")
print()
print("  COHORT SPLIT")
strong = [c for c in BUDGET if sum(1 for x in BUDGET[c] if x > 0) == 4]
weak = [c for c in BUDGET if sum(1 for x in BUDGET[c] if x > 0) < 4]
print(f"    all 4 tactics: {', '.join(strong)}")
print(f"      budgets ${min(sum(BUDGET[c]) for c in strong):,}-${max(sum(BUDGET[c]) for c in strong):,}"
      f"  web demand {min(WEB[c][3] for c in strong)}-{max(WEB[c][3] for c in strong)}")
print(f"    partial:       {', '.join(weak)}")
print(f"      budgets ${min(sum(BUDGET[c]) for c in weak):,}-${max(sum(BUDGET[c]) for c in weak):,}"
      f"  web demand {min(WEB[c][3] for c in weak)}-{max(WEB[c][3] for c in weak)}")

wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
bad = PatternFill("solid", fgColor="FFC7CE")
good = PatternFill("solid", fgColor="C6EFCE")
yellow = PatternFill("solid", fgColor="FFEB9C")
ours = PatternFill("solid", fgColor="FFF2CC")

if "Q3_Web_Ops" in wb.sheetnames:
    del wb["Q3_Web_Ops"]
ws = wb.create_sheet("Q3_Web_Ops")

ws["A1"] = "Web Sales & Service + Web Productivity Budgets — Q3 actual"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("Two findings: web staff is ALSO capped at 7 and we sit at 3, and there are FOUR web productivity "
            "tactics of which we fund ONE. Every firm running all four gets 435-569 web units; the two "
            "running fewer get 120-138.")
ws["A2"].alignment = Alignment(wrap_text=True)


def block(start, rows, wrap_col=None):
    for i, row in enumerate(rows):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(start + i, c, v)
            if i == 0:
                cell.fill = hdr
                cell.font = bold
            elif wrap_col and c == wrap_col:
                cell.alignment = Alignment(wrap_text=True)
    return start + len(rows) + 2


r = 4
rows = [["WEB STAFF", "Sales", "Service", "Total", "Web demand", "Demand / web head", "Slots open to 7"]]
for co, (s, sv, t, d) in sorted(WEB.items(), key=lambda x: -x[1][3]):
    rows.append([co, s, sv, t, d, (d / t) if t else 0, (7 - t) if t else "no web centre"])
r2 = block(r, rows)
for i in range(1, len(rows)):
    ws.cell(r + i, 6).number_format = "0.0"
    if rows[i][0] == "WeBike":
        for c in range(1, 8):
            ws.cell(r + i, c).fill = ours
            ws.cell(r + i, c).font = bold
        ws.cell(r + i, 6).fill = bad
        ws.cell(r + i, 7).fill = bad
r = r2

rows = [
    ["THE WEB STAFF CAP", "Value", "Note"],
    ["Max web staff observed", 7, "LiteCycle, BB LLC, SpaceBikes and MILC all sit exactly at 7"],
    ["WeBike web staff", 3, "2 sales + 1 service - we CUT to this in Q3"],
    ["Slots open", 4, "same 7-person ceiling as the city stores"],
    ["Our demand per web head", 40.0, "LOWEST of all six firms with a web centre"],
    ["Leader demand per web head", 81.3, "BB LLC"],
    ["Read", None, "The web centre has the same 7-person ceiling as a store, but unlike Amsterdam we are "
     "nowhere near it. Four open slots and the worst output per head is the largest single pool of "
     "unclaimed demand we have found."],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 2, 2).fill = bad
ws.cell(r + 4, 2).fill = bad
ws.cell(r + 6, 3).fill = yellow
r = r2

# budgets
rows = [["WEB PRODUCTIVITY BUDGETS"] + list(BUDGET.keys()) + ["Average"]]
AVG = [5667, 7000, 6667, 8000]
for i, t in enumerate(TACTICS):
    rows.append([t] + [BUDGET[c][i] for c in BUDGET] + [AVG[i]])
rows.append(["TOTAL"] + [sum(BUDGET[c]) for c in BUDGET] + [sum(AVG)])
rows.append(["Tactics funded (of 4)"] + [sum(1 for x in BUDGET[c] if x > 0) for c in BUDGET] + [None])
rows.append(["Web demand"] + [WEB[c][3] for c in BUDGET] + [None])
r2 = block(r, rows)
we_col = list(BUDGET.keys()).index("WeBike") + 2
for i in range(1, len(rows)):
    ws.cell(r + i, we_col).fill = ours
    ws.cell(r + i, we_col).font = bold
for i in range(1, 5):
    if BUDGET["WeBike"][i - 1] == 0:
        ws.cell(r + i, we_col).fill = bad
r = r2

rows = [
    ["THE PATTERN IS ALMOST PERFECT", "Tactics funded", "Total budget", "Web demand"],
    ["BB LLC", 4, 33000, 569],
    ["SpaceBikes", 4, 33000, 491],
    ["LiteCycle", 4, 29000, 435],
    ["MILC Bikes", 4, 24000, 435],
    ["Spoke'd Up", 2, 9000, 138],
    ["WeBike", 1, 6000, 120],
    ["Read", None, None, "Every firm funding all four tactics ($24k-33k) lands at 435-569 web units. "
     "Both firms funding fewer land at 120-138. Headcount moves with budget (7 vs 3) so the two effects "
     "cannot be fully separated - but we are at the bottom of BOTH."],
]
r2 = block(r, rows)
for i in range(1, 5):
    ws.cell(r + i, 4).fill = good
for i in (5, 6):
    ws.cell(r + i, 4).fill = bad
ws.cell(r + 7, 4).alignment = Alignment(wrap_text=True)
ws.cell(r + 7, 4).fill = yellow
r = r2

rows = [
    ["WHAT WE NEVER STARTED", "Our budget", "Industry average", "Note"],
    ["Toll-free phone", 6000, 5667, "ABOVE average - our one good web decision (raised from $3k)"],
    ["Advanced shopping cart & checkout", 0, 7000, "NEVER STARTED - 4 of 6 firms fund it at $7,000"],
    ["Continuous page upgrades", 0, 6667, "WE STOPPED THIS IN Q3 - SpaceBikes funds it at $12,000"],
    ["Order tracking software", 0, 8000, "NEVER STARTED - 4 of 6 firms fund it at $8,000"],
    ["Our total", 6000, 27500, "average of the four full-tactic firms"],
]
r2 = block(r, rows, wrap_col=4)
ws.cell(r + 1, 2).fill = good
for i in (2, 3, 4):
    ws.cell(r + i, 2).fill = bad
r = r2

rows = [
    ["THE FULL WEB REBUILD CASE", "Value", "Note"],
    ["Current web demand", 120, "3 staff, 1 tactic, $6,000"],
    ["Full-tactic cohort web demand range", "435 - 569", "7 staff, 4 tactics, $24k-33k"],
    ["Conservative target (cohort floor)", 435, "LiteCycle / MILC level"],
    ["Additional demand", "=C{0}-C{1}".format(r + 3, r + 1), None],
    ["Contribution per unit", 750, None],
    ["Gross margin opportunity", "=C{0}*C{1}".format(r + 4, r + 5), None],
    ["Cost: +4 web staff per quarter", 24425, "4 x $24,425/yr = ~$6,106/qtr each"],
    ["Cost: raise web budget $6,000 -> $28,000", 22000, "add 3 tactics, top up toll-free"],
    ["Total incremental cost", "=C{0}+C{1}".format(r + 7, r + 8), None],
    ["NET QUARTERLY GAIN", "=C{0}-C{1}".format(r + 6, r + 9), None],
    ["Note", None, "This supersedes the earlier +$89,600 estimate, which only modelled matching the "
     "average web MIX. Matching the leaders' web OPERATION is worth far more."],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 6, 2).fill = good
ws.cell(r + 10, 2).fill = good
ws.cell(r + 10, 2).font = Font(bold=True, size=12)
ws.cell(r + 11, 3).alignment = Alignment(wrap_text=True)
r = r2

rows = [
    ["CAPACITY WARNING", "Units", "Note"],
    ["Store demand today", 507, None],
    ["Rio +3 staff at ~67/head", 200, "see Q3_Staffing"],
    ["Web at cohort floor", 435, None],
    ["Indicative total demand", "=C{0}+C{1}+C{2}".format(r + 1, r + 2, r + 3), "before ill will"],
    ["Max units at 24/day, 65 days, 74% productivity", "=24*65*0.74", None],
    ["Required scheduled OC per day", "=C{0}/(0.74*65)".format(r + 4), "units / (productivity x days)"],
    ["Fixed capacity per day", 24, None],
    ["Headroom", "=C{0}-C{1}".format(r + 7, r + 6), None],
    ["Read", None, "Doing all of this at once pushes demand toward ~1,140 and required OC to roughly "
     "21/day - inside our 24/day fixed capacity, but with little slack. Schedule OC to the FORECAST and "
     "re-run the production simulation before committing."],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 6, 2).fill = yellow
ws.cell(r + 9, 3).fill = yellow
ws.cell(r + 9, 3).alignment = Alignment(wrap_text=True)

ws.column_dimensions["A"].width = 50
ws.column_dimensions["B"].width = 16
ws.column_dimensions["C"].width = 20
ws.column_dimensions["D"].width = 58
for col in "EFGH":
    ws.column_dimensions[col].width = 14

wb.save(DST)
print("\nAdded Q3_Web_Ops")
