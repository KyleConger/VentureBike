"""Add Q3 sales & service staffing by city + the 7-person cap finding."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

DST = "quarters/Q3/Q3Data.xlsx"

# city, company, total, service, recreation, mountain, speed
STAFF = [
    ("New York City",  "BB LLC",     7, 1, 0, 3, 3),
    ("New York City",  "Spoke'd Up", 5, 1, 2, 0, 2),
    ("New York City",  "SpaceBikes", 7, 2, 2, 0, 3),
    ("Rio de Janeiro", "WeBike",     4, 1, 0, 1, 1),
    ("Amsterdam",      "Bike Bros",  7, 2, 0, 2, 2),
    ("Amsterdam",      "LiteCycle",  5, 1, 0, 1, 3),
    ("Amsterdam",      "WeBike",     7, 1, 0, 3, 2),
    ("Amsterdam",      "BB LLC",     7, 1, 0, 3, 3),
    ("Amsterdam",      "Spoke'd Up", 6, 1, 2, 0, 3),
    ("Amsterdam",      "SpaceBikes", 7, 2, 2, 0, 3),
    ("Amsterdam",      "MILC Bikes", 7, 1, 3, 0, 3),
    ("Bangalore",      "LiteCycle",  5, 1, 0, 1, 3),
]

# firm -> (store demand, web demand, mountain demand, speed demand, rec demand)
DEM = {
    "BB LLC":     (938, 569, 751, 657,  99),
    "SpaceBikes": (807, 491,   0, 588, 710),
    "LiteCycle":  (549, 435, 340, 548,  96),
    "MILC Bikes": (364, 435,   0, 328, 471),
    "Spoke'd Up": (716, 138,   0, 362, 492),
    "WeBike":     (507, 120, 358, 215,  54),
    "Bike Bros":  (490,   0, 172, 180, 138),
}
CITY_DEM = {  # WeBike only
    "Rio de Janeiro": 270,
    "Amsterdam": 295,
}

print("UNTRAINED HEADCOUNT CHECK (total vs sum of specialisations)")
untrained = {}
for city, co, tot, svc, rec, mtn, spd in STAFF:
    gap = tot - (svc + rec + mtn + spd)
    untrained[co] = untrained.get(co, 0) + gap
    flag = "  <-- UNTRAINED" if gap else ""
    print(f"  {city:<16}{co:<12} total {tot}  assigned {svc+rec+mtn+spd}  gap {gap}{flag}")
print()
print("MAX STAFF IN ANY CITY:", max(s[2] for s in STAFF), "-> 7 appears to be the hard cap")
print()

wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
bad = PatternFill("solid", fgColor="FFC7CE")
good = PatternFill("solid", fgColor="C6EFCE")
yellow = PatternFill("solid", fgColor="FFEB9C")
ours = PatternFill("solid", fgColor="FFF2CC")

if "Q3_Staffing" in wb.sheetnames:
    del wb["Q3_Staffing"]
ws = wb.create_sheet("Q3_Staffing")

ws["A1"] = "Sales and Service People — World Market, Q3 actual"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("Two hard findings: 7 people per city is the cap (nobody exceeds it, and Amsterdam has us AT it), "
            "and we are the only firm besides Bike Bros carrying UNTRAINED staff.")
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


# raw table
r = 4
rows = [["City", "Company", "Total", "Service", "Recreation", "Mountain", "Speed",
         "Assigned", "UNTRAINED", "At 7-cap?"]]
for city, co, tot, svc, rec, mtn, spd in STAFF:
    a = svc + rec + mtn + spd
    rows.append([city, co, tot, svc, rec, mtn, spd, a, tot - a, "YES" if tot == 7 else ""])
r2 = block(r, rows)
for i in range(1, len(rows)):
    if rows[i][1] == "WeBike":
        for c in range(1, 11):
            ws.cell(r + i, c).fill = ours
            ws.cell(r + i, c).font = bold
    if rows[i][8]:
        ws.cell(r + i, 9).fill = bad
r = r2

# firm rollup
rows = [["FIRM ROLLUP", "Store people", "Service", "Recreation", "Mountain", "Speed", "Untrained",
         "Store demand", "Demand / store person"]]
agg = {}
for city, co, tot, svc, rec, mtn, spd in STAFF:
    a = agg.setdefault(co, [0, 0, 0, 0, 0, 0])
    a[0] += tot; a[1] += svc; a[2] += rec; a[3] += mtn; a[4] += spd
    a[5] += tot - (svc + rec + mtn + spd)
for co in sorted(agg, key=lambda x: -(DEM[x][0] / agg[x][0])):
    a = agg[co]
    rows.append([co, a[0], a[1], a[2], a[3], a[4], a[5], DEM[co][0], DEM[co][0] / a[0]])
r2 = block(r, rows)
for i in range(1, len(rows)):
    ws.cell(r + i, 9).number_format = "0.0"
    if rows[i][0] == "WeBike":
        for c in range(1, 10):
            ws.cell(r + i, c).fill = ours
            ws.cell(r + i, c).font = bold
        ws.cell(r + i, 9).fill = bad
        ws.cell(r + i, 7).fill = bad
r = r2

others = [c for c in agg if c != "WeBike"]
avg_other = sum(DEM[c][0] / agg[c][0] for c in others) / len(others)
rows = [
    ["PRODUCTIVITY GAP", "Value", "Note"],
    ["WeBike demand per store person", DEM["WeBike"][0] / agg["WeBike"][0], "LOWEST in the industry"],
    ["Industry average (excl. us)", avg_other, "Bike Bros 70.0 is best; BB LLC 67.0"],
    ["Shortfall per person", "=C{0}-C{1}".format(r + 2, r + 1), None],
    ["Our store people", 11, None],
    ["Store demand if we matched the average", "=C{0}*C{1}".format(r + 2, r + 4), None],
    ["Additional units available at current headcount", "=C{0}-507".format(r + 5), "no new hires needed"],
]
r2 = block(r, rows, wrap_col=3)
for i in (1, 2, 3):
    ws.cell(r + i, 2).number_format = "0.0"
ws.cell(r + 1, 2).fill = bad
ws.cell(r + 2, 2).fill = good
ws.cell(r + 6, 2).fill = good
r = r2

# THE CAP
rows = [
    ["THE 7-PERSON CITY CAP — this settles depth vs breadth", "Value", "Note"],
    ["Maximum staff observed in any city (all 12 rows)", 7, "BB LLC, SpaceBikes, Bike Bros, MILC and our own Amsterdam all sit at 7"],
    ["WeBike Amsterdam headcount", 7, "AT THE CAP - we cannot add a single person there"],
    ["WeBike Rio headcount", 4, "3 SLOTS OPEN - the only store depth available to us"],
    ["Read", None, "'Deepen Amsterdam' is not a legal move. Amsterdam is maxed at 7. The only ways to add "
     "store selling capacity are (a) fill Rio's 3 open slots, or (b) open a new city, which unlocks 7 more."],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 2, 2).fill = bad
ws.cell(r + 3, 2).fill = good
ws.cell(r + 4, 3).fill = yellow
r = r2

rows = [
    ["RIO vs AMSTERDAM — where a body is worth more", "Rio", "Amsterdam"],
    ["Store people", 4, 7],
    ["City demand (all channels)", 270, 295],
    ["Demand per person", "=C{0}/C{1}".format(r + 2, r + 1), "=D{0}/D{1}".format(r + 2, r + 1)],
    ["Rival stores in city", 0, 6],
    ["Open slots to the 7-cap", 3, 0],
    ["Read", "A Rio head is worth ~60% more than an Amsterdam head because Rio has no rival store. "
     "Filling the 3 open Rio slots is the highest-return staffing move on the board.", None],
    ["CAVEAT", "City demand includes web orders, so per-person figures are directional rather than "
     "pure store productivity. The ranking is robust; the exact multiple is not.", None],
]
r2 = block(r, rows)
ws.cell(r + 3, 2).number_format = "0.0"
ws.cell(r + 3, 3).number_format = "0.0"
ws.cell(r + 3, 2).fill = good
ws.cell(r + 3, 3).fill = bad
ws.cell(r + 6, 2).alignment = Alignment(wrap_text=True)
ws.cell(r + 6, 2).fill = yellow
ws.cell(r + 7, 2).alignment = Alignment(wrap_text=True)
r = r2

rows = [
    ["UNTRAINED STAFF — a free win we skipped", "Value", "Note"],
    ["WeBike untrained people", 2, "1 in Rio, 1 in Amsterdam"],
    ["Firms with any untrained staff", "WeBike (2), Bike Bros (1)", "the other five are 100% trained"],
    ["Mountain training cost", 400, "per quarter"],
    ["Speed training cost", 400, "per quarter"],
    ["Cost to train both", 800, "trivial against ~$750 contribution PER UNIT"],
    ["Read", None, "We are paying two full salaries ($24,425/yr each) for untrained sellers while every "
     "serious competitor trains everyone. Fix this first - it costs $800."],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 1, 2).fill = bad
ws.cell(r + 5, 2).fill = good
ws.cell(r + 6, 3).fill = yellow
r = r2

rows = [
    ["SPECIALISATION vs SEGMENT RESULT", "Specialists", "Segment demand", "Demand per specialist"],
    ["MOUNTAIN", None, None, None],
    ["LiteCycle", 2, 340, "=C{0}/B{0}".format(r + 2)],
    ["BB LLC", 6, 751, "=C{0}/B{0}".format(r + 3)],
    ["WeBike", 4, 358, "=C{0}/B{0}".format(r + 4)],
    ["Bike Bros", 2, 172, "=C{0}/B{0}".format(r + 5)],
    ["SPEED", None, None, None],
    ["BB LLC", 6, 657, "=C{0}/B{0}".format(r + 7)],
    ["MILC Bikes", 3, 328, "=C{0}/B{0}".format(r + 8)],
    ["SpaceBikes", 6, 588, "=C{0}/B{0}".format(r + 9)],
    ["LiteCycle", 6, 548, "=C{0}/B{0}".format(r + 10)],
    ["Spoke'd Up", 5, 362, "=C{0}/B{0}".format(r + 11)],
    ["WeBike", 3, 215, "=C{0}/B{0}".format(r + 12)],
    ["RECREATION", None, None, None],
    ["SpaceBikes", 4, 710, "=C{0}/B{0}".format(r + 14)],
    ["MILC Bikes", 3, 471, "=C{0}/B{0}".format(r + 15)],
    ["Spoke'd Up", 4, 492, "=C{0}/B{0}".format(r + 16)],
    ["WeBike", 0, 54, "spillover only - no Rec specialists"],
]
r2 = block(r, rows)
for rr in (r + 1, r + 6, r + 13):
    ws.cell(rr, 1).font = bold
    ws.cell(rr, 1).fill = hdr
for rr in (r + 4, r + 12, r + 17):
    ws.cell(rr, 1).fill = ours
for i in range(1, len(rows)):
    ws.cell(r + i, 4).number_format = "0.0"
ws.cell(r + 12, 2).fill = bad
r = r2

rows = [
    ["Q4 STAFFING PLAN", "Action", "Cost", "Expected units"],
    ["1. Train the 2 untrained", "1 Mountain (Rio) + 1 Speed (Amsterdam)", 800, "free productivity"],
    ["2. Fill Rio to the 7-cap", "+3 people: 1 Mountain, 1 Speed, 1 Service", "~$18,300/qtr salary + hire + train",
     "~+200 at Rio's ~67/head"],
    ["3. Rebuild web staff", "3 -> 6 web sales/support", "~$18,300/qtr", "~+152 (see Q3_Channel)"],
    ["4. Add Speed specialists", "we have 3 vs BB LLC/LiteCycle/SpaceBikes 6", "included above",
     "Speed is our weakest segment AND the largest"],
    ["5. New York store (if funded)", "unlocks 7 fresh slots; staff Speed-heavy", "setup + lease + salaries",
     "~400-570 at incumbent rates"],
    ["NOT possible", "Adding anyone in Amsterdam - we are at the 7-person cap", None, None],
]
block(r, rows, wrap_col=2)
for i in range(1, len(rows)):
    ws.cell(r + i, 2).fill = yellow
ws.cell(r + len(rows) - 1, 2).fill = bad

ws.column_dimensions["A"].width = 44
ws.column_dimensions["B"].width = 40
ws.column_dimensions["C"].width = 34
ws.column_dimensions["D"].width = 30
for col in "EFGHIJ":
    ws.column_dimensions[col].width = 13

wb.save(DST)
print("Added Q3_Staffing")
