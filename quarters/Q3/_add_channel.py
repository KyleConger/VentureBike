"""Add Q3 demand distribution by channel (stores vs web) and the web rebuild case."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

DST = "quarters/Q3/Q3Data.xlsx"

# company, store demand, web demand, total, inferred stores, inferred web centers
CH = [
    ("BB LLC",      938, 569, 1507, 2, 1),
    ("SpaceBikes",  807, 491, 1298, 2, 1),
    ("LiteCycle",   549, 435,  984, 2, 1),
    ("MILC Bikes",  364, 435,  799, 1, 1),
    ("Spoke'd Up",  716, 138,  854, 2, 1),
    ("WeBike",      507, 120,  627, 2, 1),
    ("Bike Bros",   490,   0,  490, 1, 0),
]

print("RECONCILIATION")
bad_rows = [c for c in CH if c[1] + c[2] != c[3]]
print("  store + web = total for every firm" if not bad_rows else f"  MISMATCH: {bad_rows}")
assert not bad_rows
tot = sum(c[3] for c in CH)
print(f"  industry total {tot} (report says 6559) {'OK' if tot == 6559 else 'MISMATCH'}")
assert tot == 6559
print("  reconciles\n")

wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
bad = PatternFill("solid", fgColor="FFC7CE")
good = PatternFill("solid", fgColor="C6EFCE")
yellow = PatternFill("solid", fgColor="FFEB9C")
ours = PatternFill("solid", fgColor="FFF2CC")

if "Q3_Channel" in wb.sheetnames:
    del wb["Q3_Channel"]
ws = wb.create_sheet("Q3_Channel")

ws["A1"] = "Demand Distribution by Channel — Q3 actual"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("We cut web staff and spend in Q3. This is the bill: 120 web units, the lowest of any firm that "
            "operates a web centre, and 19.1% web mix against an industry average of 34.9%.")
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
rows = [["Company", "Store demand", "Store %", "Web demand", "Web %", "Total demand",
         "Stores", "Demand / store", "Web centres", "Demand / web centre"]]
for co, sd, wd, t, ns, nw in sorted(CH, key=lambda x: -x[3]):
    rows.append([co, sd, sd / t, wd, wd / t, t, ns, sd / ns, nw, (wd / nw) if nw else 0])
r2 = block(r, rows)
for i in range(1, len(CH) + 1):
    ws.cell(r + i, 3).number_format = "0.0%"
    ws.cell(r + i, 5).number_format = "0.0%"
    ws.cell(r + i, 8).number_format = "0"
    ws.cell(r + i, 10).number_format = "0"
    if rows[i][0] == "WeBike":
        for c in range(1, 11):
            ws.cell(r + i, c).fill = ours
            ws.cell(r + i, c).font = bold
r = r2

web_avg = sum(c[2] / c[3] for c in CH if c[5]) / len([c for c in CH if c[5]])
rows = [
    ["OUR CHANNEL MIX vs INDUSTRY", "Value", "Note"],
    ["WeBike web mix [%]", 0.191, "we CUT web staff and spend in Q3"],
    ["Industry average web mix [%] (firms with a web centre)", web_avg, "MILC 54.4% - LiteCycle 44.2% - BB LLC 37.8% - SpaceBikes 37.8%"],
    ["Gap [pct points]", "=C{0}-C{1}".format(r + 2, r + 1), None],
    ["WeBike web demand [units]", 120, "LOWEST of any firm operating a web centre"],
    ["Industry web demand range", "138 - 569", "Spoke'd Up 138 is the only firm near us"],
    ["Rank on web demand", "6 of 7", "only Bike Bros is lower, and it has NO web centre"],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 1, 2).number_format = "0.0%"
ws.cell(r + 2, 2).number_format = "0.0%"
ws.cell(r + 3, 2).number_format = "0.0%"
ws.cell(r + 1, 2).fill = bad
ws.cell(r + 4, 2).fill = bad
r = r2

rows = [
    ["WHY WEB MATTERS MORE THAN WE THOUGHT", "Units", "Note"],
    ["Our demand from New York City (no store)", 38, "web only"],
    ["Our demand from Bangalore (no store)", 24, "web only"],
    ["Subtotal from cities where we have NO store", "=B{0}+B{1}".format(r + 1, r + 2), None],
    ["Total web demand", 120, None],
    ["Share of web demand coming from non-store cities", "=B{0}/B{1}".format(r + 3, r + 4), None],
    ["Store demand (Amsterdam + Rio)", 507, None],
    ["Amsterdam + Rio total demand", 565, "so 58 units of web demand also came from store cities"],
    ["CHECK", "=B{0}-B{1}".format(r + 7, r + 6), "58 = web demand inside store cities; 58 + 62 = 120 web total"],
    ["Read", None, "The web centre is HALF our reach into cities where we have no store. "
     "It is the only way to touch New York and Bangalore today, and we shrank it."],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 5, 2).number_format = "0.0%"
ws.cell(r + 9, 3).alignment = Alignment(wrap_text=True)
ws.cell(r + 9, 3).fill = yellow
r = r2

rows = [
    ["THE WEB REBUILD CASE", "Value", "Note"],
    ["Current store demand [units]", 507, "hold constant"],
    ["Current web demand [units]", 120, None],
    ["Current total", "=B{0}+B{1}".format(r + 1, r + 2), None],
    ["If web mix matched industry average (34.9%)", None, None],
    ["  implied total demand", "=B{0}/(1-{1})".format(r + 1, round(web_avg, 4)), "store demand becomes 65.1% of total"],
    ["  implied web demand", "=B{0}-B{1}".format(r + 6, r + 1), None],
    ["  ADDITIONAL demand vs today", "=B{0}-B{1}".format(r + 7, r + 2), None],
    ["Contribution per unit (Q3 actual, avg)", 750, "Hike Bike $747 / Swift Bike $756"],
    ["Gross margin opportunity", "=B{0}*B{1}".format(r + 8, r + 9), None],
    ["Cost of ~4 additional web staff per quarter", 24425, "$24,425/yr each = ~$6,106/qtr x 4"],
    ["Net quarterly gain", "=B{0}-B{1}".format(r + 10, r + 11), None],
    ["Read", None, "Web headcount is the cheapest demand in the game: no lease, no store setup, "
     "and it reaches all four cities. This is a far better first dollar than a third store."],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 10, 2).fill = good
ws.cell(r + 12, 2).fill = good
ws.cell(r + 12, 2).font = bold
ws.cell(r + 13, 3).alignment = Alignment(wrap_text=True)
r = r2

rows = [
    ["STORE PRODUCTIVITY — we under-staff what we own", "Demand / store", "Note"],
    ["Bike Bros", 490, "ONE store, highest per-store demand in the industry"],
    ["BB LLC", 469, "leader"],
    ["SpaceBikes", 404, None],
    ["Spoke'd Up", 358, None],
    ["LiteCycle", 275, None],
    ["WeBike", 254, "LOWEST of the two-store firms"],
    ["MILC Bikes", 364, "web-first model; Amsterdam store base"],
    ["Read", None, "BB LLC pulls 1.8x our demand from the same number of stores. Depth inside Amsterdam "
     "and Rio is unfinished business before we add a third location."],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 6, 2).fill = bad
ws.cell(r + 8, 3).alignment = Alignment(wrap_text=True)
r = r2

rows = [
    ["CHANNEL MODELS IN THIS INDUSTRY", "Profile", "Result"],
    ["Web-first", "MILC Bikes - 54.4% web, lowest store demand (364)", "12.2% share on 1 store base"],
    ["Balanced", "LiteCycle 44.2% - BB LLC 37.8% - SpaceBikes 37.8%", "the three strongest firms sit here"],
    ["Store-heavy", "Spoke'd Up 83.8% - WeBike 80.9%", "13.0% and 9.6% share"],
    ["Store-only", "Bike Bros 100% store, no web centre", "7.5% share, LAST, falling everywhere"],
    ["Read", None, "The top three firms all run 38-44% web. Store-only is last. We are 80.9% store-heavy "
     "and drifting toward the Bike Bros failure mode."],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 2, 3).fill = good
ws.cell(r + 4, 3).fill = bad
ws.cell(r + 5, 3).alignment = Alignment(wrap_text=True)
r = r2

rows = [
    ["Q4 CHANNEL ACTIONS", "Decision"],
    ["1. Reverse the Q3 web cut", "Rehire web sales + support toward 6 people. Cheapest demand available: "
     "no lease, no setup, reaches all 4 cities."],
    ["2. Restart page upgrades", "We stopped this tactic in Q3. Toll-free stayed on. Both were cheap "
     "(~$3k and ~$6k/qtr) relative to ~$750 contribution per unit."],
    ["3. Deepen Amsterdam + Rio", "254 demand/store vs BB LLC's 469. Add store sales staff before adding a "
     "third city - same money, faster payback."],
    ["4. Re-evaluate city #3 order", "Web already reaches New York (38 units). A store there is still the "
     "biggest prize, but web staff + store depth may be the better FIRST dollar this quarter."],
    ["5. Never cut channel again", "Bike Bros is the control group: 100% store, no web, last place, share "
     "falling in all three segments."],
]
block(r, rows, wrap_col=2)
for i in range(1, len(rows)):
    ws.cell(r + i, 2).fill = yellow

ws.column_dimensions["A"].width = 50
ws.column_dimensions["B"].width = 20
ws.column_dimensions["C"].width = 56
for col in "DEFGHIJ":
    ws.column_dimensions[col].width = 15

wb.save(DST)
print("Added Q3_Channel")
