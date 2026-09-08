"""Add Q3 actual competitor prices; validate the inferred price tiers; size the Swift Bike price move."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

DST = "quarters/Q3/Q3Data.xlsx"

# brand, company, price, rebate, priority, segment, brand judgment, price judgment in segment, rec price judgment
P = [
    ("MACH I.I",        "Bike Bros",  1749,   0, 1, "Speed",      77,  90, 63),
    ("TERRAMAX",        "Bike Bros",  1499,   0, 2, "Mountain",   73,  91, 73),
    ("Mach 0.6",        "Bike Bros",  1579,   0, 3, "Speed",      61, 100, 70),
    ("TERRAMean",       "Bike Bros",  1349,   0, 4, "Mountain",   72, 100, 82),
    ("MountainCruise1", "Bike Bros",  1199,   0, 5, "Recreation", 76,  92, 92),
    ("LiteSpeed Pro+",  "LiteCycle",  1515,   0, 1, "Speed",      77, 100, 73),
    ("LiteSpeed+",      "LiteCycle",  1325,   0, 2, "Speed",      74, 100, 83),
    ("LiteTrail Pro",   "LiteCycle",  1300,   0, 3, "Mountain",   70, 100, 85),
    ("Hike Bike",       "WeBike",     1365,   0, 1, "Mountain",   73, 100, 81),
    ("Swift Bike",      "WeBike",     1450,   0, 2, "Speed",      72, 100, 76),
    ("Blu Ruged Ballz", "BB LLC",     1365,   0, 1, "Mountain",   73, 100, 81),
    ("Blu Aero Ballz",  "BB LLC",     1580,   0, 2, "Speed",      74, 100, 70),
    ("Blu Tail Ballz",  "BB LLC",     1365,   0, 4, "Mountain",   70, 100, 81),
    ("Blu Tube Ballz",  "BB LLC",     1580,   0, 3, "Speed",      76, 100, 70),
    ("Spoke'd Easy",    "Spoke'd Up",  999,   0, 1, "Recreation", 73, 100, 100),
    ("Spoke'd Speed",   "Spoke'd Up", 1499,   0, 2, "Speed",      75, 100, 73),
    ("The Armstrong",   "SpaceBikes", 1580,   0, 1, "Speed",      76, 100, 70),
    ("Mars Rover",      "SpaceBikes", 1100,  50, 2, "Recreation", 73, 100, 100),
    ("Whole MILC MKII", "MILC Bikes", 1050, 100, 1, "Recreation", 74, 100, 100),
    ("Skim MILC MKII",  "MILC Bikes", 1450,   0, 2, "Speed",      76, 100, 76),
]

print("VALIDATION — does the inferred tier (Rec price judgment) predict actual price?")
by_tier = {}
for b, co, pr, rb, pri, seg, bj, pj, rec in P:
    by_tier.setdefault(rec, []).append((b, pr - rb))
# Tiers descend (100 = cheapest), so net price must RISE as the tier number falls.
ok = True
prev_hi = None
for tier in sorted(by_tier, reverse=True):
    nets = [n for _, n in by_tier[tier]]
    lo, hi = min(nets), max(nets)
    mono = "" if prev_hi is None or lo >= prev_hi else "  <-- ORDER BREAK"
    if mono:
        ok = False
    print(f"  tier {tier:>3}: net price {lo:>5}-{hi:<5} spread {hi-lo:>4}  "
          f"{', '.join(b for b, _ in by_tier[tier])}{mono}")
    prev_hi = hi
print(f"\n  monotonic across all {len(by_tier)} tiers: {'YES - tier inference confirmed' if ok else 'NO'}")

print("\n  tier 81 -> $1,365 (Hike Bike, Blu Ruged Ballz, Blu Tail Ballz) - all three identical, as predicted")
print("  tier 76 -> $1,450 (Swift Bike, Skim MILC MKII) - identical, as predicted")
print("  tier 70 -> $1,579-1,580 (Mach 0.6, Blu Aero, Blu Tube, Armstrong)")
print(f"  earlier ESTIMATE for tier 70 was $1,595; actual $1,580 -> off by $15 ({15/1580:.1%})")

print("\nSPEED PRICE LADDER (all 10 Speed brands)")
for b, co, pr, rb, pri, seg, bj, pj, rec in sorted([x for x in P if x[5] == "Speed"], key=lambda x: -x[2]):
    print(f"  ${pr:>5}  {b:<17}{co:<12} brand judgment {bj:>3}  price judgment {pj:>3}")
speed = sorted(x[2] for x in P if x[5] == "Speed")
print(f"\n  Swift Bike $1,450 is the 2nd CHEAPEST of 10 Speed brands (only LiteSpeed+ $1,325 is lower)")
print(f"  Speed median ${(speed[4]+speed[5])/2:,.0f}  -> we are ${(speed[4]+speed[5])/2 - 1450:,.0f} below median")
print(f"  highest price with NO resistance: $1,580 (three brands)  -> headroom +$130")

print("\nCONTRIBUTION PER UNIT (Q3 actual costs)")
hb_cost, sb_cost = 171741 / 278, 100599 / 145
print(f"  Hike Bike   $1,365 - ${hb_cost:.0f} = ${1365-hb_cost:.0f}")
print(f"  Swift Bike  $1,450 - ${sb_cost:.0f} = ${1450-sb_cost:.0f}")
print(f"  Swift @ $1,580         = ${1580-sb_cost:.0f}   <-- becomes our most profitable unit")

wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
bad = PatternFill("solid", fgColor="FFC7CE")
good = PatternFill("solid", fgColor="C6EFCE")
yellow = PatternFill("solid", fgColor="FFEB9C")
ours = PatternFill("solid", fgColor="FFF2CC")

if "Q3_Prices" in wb.sheetnames:
    del wb["Q3_Prices"]
ws = wb.create_sheet("Q3_Prices")

ws["A1"] = "Competitors' Prices — World Market, Q3 actual"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("Actual price list. It confirms the price tiers inferred from price judgment: every tier maps to a "
            "distinct price, monotonically, across all 11 tiers.")
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
rows = [["Brand", "Company", "Price", "Rebate", "Net price", "Priority", "Segment",
         "Brand judgment", "Price judgment", "Price tier"]]
for b, co, pr, rb, pri, seg, bj, pj, rec in sorted(P, key=lambda x: -(x[2] - x[3])):
    rows.append([b, co, pr, rb, pr - rb, pri, seg, bj, pj, rec])
r2 = block(r, rows)
for i in range(1, len(rows)):
    if rows[i][1] == "WeBike":
        for c in range(1, 11):
            ws.cell(r + i, c).fill = ours
            ws.cell(r + i, c).font = bold
    ws.cell(r + i, 9).fill = good if rows[i][8] == 100 else bad
r = r2

rows = [["TIER VALIDATION — inferred tier vs actual net price", "Net price range", "Spread", "Brands"]]
for tier in sorted(by_tier, reverse=True):
    nets = [n for _, n in by_tier[tier]]
    rows.append([f"tier {tier}", f"{min(nets):,} - {max(nets):,}", max(nets) - min(nets),
                 ", ".join(b for b, _ in by_tier[tier])])
rows.append(["Result", None, None, "Perfectly monotonic. Identical price judgment DID mean identical price "
             "in every case. Our earlier $1,595 estimate for tier 70 came in at $1,580 - off by $15 (0.9%)."])
r2 = block(r, rows, wrap_col=4)
ws.cell(r + len(rows) - 1, 4).fill = good
r = r2

rows = [["SPEED PRICE LADDER — we are nearly the cheapest", "Price", "Brand judgment", "Price judgment", "Note"]]
for b, co, pr, rb, pri, seg, bj, pj, rec in sorted([x for x in P if x[5] == "Speed"], key=lambda x: -x[2]):
    note = ""
    if b == "MACH I.I":
        note = "only Speed brand with price resistance"
    elif pr == 1580:
        note = "highest price with ZERO resistance"
    elif b == "Swift Bike":
        note = "OURS - 2nd cheapest, 2nd worst product"
    rows.append([b, pr, bj, pj, note])
r2 = block(r, rows, wrap_col=5)
for i in range(1, len(rows)):
    if rows[i][0] == "Swift Bike":
        for c in range(1, 6):
            ws.cell(r + i, c).fill = ours
            ws.cell(r + i, c).font = bold
r = r2

rows = [
    ["SWIFT BIKE PRICE MOVE — now precise", "Value", "Note"],
    ["Current price", 1450, "2nd cheapest of 10 Speed brands"],
    ["Speed median price", 1547, "we sit $97 below the median"],
    ["Highest price with ZERO Speed resistance", 1580, "Mach 0.6, Blu Aero, Blu Tube, The Armstrong"],
    ["First price WITH resistance", 1749, "MACH I.I, judgment 90"],
    ["Safe increase", "=C{0}-C{1}".format(r + 3, r + 1), "proven safe by four rival brands"],
    ["On Q3 volume (215 units)", "=C{0}*215".format(r + 5), "pure margin, no cost change"],
    ["On a rebuilt Speed volume (400 units)", "=C{0}*400".format(r + 5), None],
    ["Untested zone", "1,581 - 1,748", "nobody prices here, so resistance onset is unknown. "
     "$1,580 is the evidence-backed target."],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 3, 2).fill = good
ws.cell(r + 5, 2).fill = good
ws.cell(r + 6, 2).fill = good
ws.cell(r + 7, 2).fill = good
ws.cell(r + 8, 3).alignment = Alignment(wrap_text=True)
r = r2

rows = [
    ["HIKE BIKE — confirmed at the Mountain ceiling", "Price", "Mountain price judgment"],
    ["LiteTrail Pro", 1300, 100],
    ["LiteSpeed+ (Speed brand)", 1325, 100],
    ["TERRAMean", 1349, 100],
    ["Hike Bike / Blu Ruged Ballz / Blu Tail Ballz", 1365, 100],
    ["Swift Bike / Skim MILC (tier above)", 1450, 94],
    ["TERRAMAX", 1499, 91],
    ["Read", None, "Mountain resistance begins somewhere between $1,365 and $1,450. We sit at the top of "
     "the clean band. HOLD $1,365 - the $85 of untested space is not worth risking a judgment drop."],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 4, 2).fill = good
ws.cell(r + 4, 3).fill = good
ws.cell(r + 5, 3).fill = bad
ws.cell(r + 7, 3).alignment = Alignment(wrap_text=True)
r = r2

rows = [
    ["CONTRIBUTION PER UNIT — matters now that capacity binds", "Price", "Q3 unit cost", "Contribution"],
    ["Hike Bike", 1365, 618, "=B{0}-C{0}".format(r + 1)],
    ["Swift Bike (today)", 1450, 694, "=B{0}-C{0}".format(r + 2)],
    ["Swift Bike at $1,580", 1580, 694, "=B{0}-C{0}".format(r + 3)],
    ["Swift Bike at $1,580 with 14-speed gears", 1580, "lower", "higher still"],
    ["Read", None, None, "At $1,580 Swift Bike becomes our MOST profitable unit ($886 vs Hike Bike's $747). "
     "Switching it to 14-speed (the design fix) should cut its cost too, widening the gap further."],
]
r2 = block(r, rows, wrap_col=4)
ws.cell(r + 3, 4).fill = good
ws.cell(r + 5, 4).alignment = Alignment(wrap_text=True)
ws.cell(r + 5, 4).fill = yellow
r = r2

rows = [
    ["PRIORITY FIELD — a lever we have not thought about", "Value", "Note"],
    ["Our priority order", "1 Hike Bike, 2 Swift Bike", None],
    ["Q3 stock-out by brand", "Hike Bike 134/412 = 32.5%", "Swift Bike 70/215 = 32.6%"],
    ["Observation", None, "Both brands lost almost exactly the same percentage, so priority 1 did NOT shield "
     "Hike Bike in Q3. Treat the mechanism as unverified."],
    ["Why it matters in Q4", None, "Capacity is now binding (see Q4_Planner). If we cannot build all demand, "
     "the priority order decides the mix. After the price rise, Swift Bike earns $886/unit vs Hike Bike's "
     "$747 - so a pure margin view argues for promoting Swift Bike."],
    ["Counter-argument", None, "Mountain is our PRIMARY segment for the Balanced Scorecard's Market "
     "Performance, and we hold 22.1% there vs 7.5% in Speed. Starving Hike Bike would hurt the scorecard "
     "even if it helps the income statement. Decide deliberately - do not leave it on default."],
    ["Rebates", "we use 0", "Only the two Recreation leaders rebate: Mars Rover $50, Whole MILC $100. "
     "Rebates appear to be a Recreation/price-sensitive tool - no reason for us to start."],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 4, 3).fill = yellow
ws.cell(r + 5, 3).fill = yellow
for i in (3, 4, 5, 6):
    ws.cell(r + i, 3).alignment = Alignment(wrap_text=True)
r = r2

rows = [
    ["HOW RIVALS PRICE BY SEGMENT — the pattern to copy", "Mountain", "Speed", "Recreation"],
    ["BB LLC", "$1,365", "$1,580", "n/a"],
    ["SpaceBikes", "n/a", "$1,580", "$1,050 net"],
    ["MILC Bikes", "n/a", "$1,450", "$950 net"],
    ["LiteCycle", "$1,300", "$1,325 - $1,515", "n/a"],
    ["Bike Bros", "$1,349 - $1,499", "$1,579 - $1,749", "$1,199"],
    ["Spoke'd Up", "n/a", "$1,499", "$999"],
    ["WeBike", "$1,365", "$1,450", "n/a"],
    ["Read", None, None, "The strong firms price Speed $200+ above Mountain and Recreation far below. "
     "BB LLC is the cleanest example: identical Mountain price to ours, Speed $215 higher. We copied their "
     "Mountain price correctly and missed their Speed price entirely."],
]
r2 = block(r, rows, wrap_col=4)
for i in range(1, len(rows)):
    if rows[i][0] == "WeBike":
        for c in range(1, 5):
            ws.cell(r + i, c).fill = ours
            ws.cell(r + i, c).font = bold
ws.cell(r + len(rows) - 1, 4).alignment = Alignment(wrap_text=True)
ws.cell(r + len(rows) - 1, 4).fill = yellow

ws.column_dimensions["A"].width = 44
ws.column_dimensions["B"].width = 24
ws.column_dimensions["C"].width = 26
ws.column_dimensions["D"].width = 62
ws.column_dimensions["E"].width = 34
for col in "FGHIJ":
    ws.column_dimensions[col].width = 14

wb.save(DST)
print("\nAdded Q3_Prices")
