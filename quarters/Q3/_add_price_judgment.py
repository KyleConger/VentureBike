"""Add Q3 price judgment + inferred price tiers and the pricing headroom analysis."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

DST = "quarters/Q3/Q3Data.xlsx"

# brand, company, rec, mtn, speed, target segment, brand judgment in target
PJ = [
    ("MACH I.I",        "Bike Bros",  63,  78,  90, "Speed", 77),
    ("TERRAMAX",        "Bike Bros",  73,  91, 100, "Mountain", 73),
    ("Mach 0.6",        "Bike Bros",  70,  86, 100, "Speed", 61),
    ("TERRAMean",       "Bike Bros",  82, 100, 100, "Mountain", 72),
    ("MountainCruise1", "Bike Bros",  92, 100, 100, "Recreation", 76),
    ("LiteSpeed Pro+",  "LiteCycle",  73,  90, 100, "Speed", 77),
    ("LiteSpeed+",      "LiteCycle",  83, 100, 100, "Speed", 74),
    ("LiteTrail Pro",   "LiteCycle",  85, 100, 100, "Mountain", 70),
    ("Hike Bike",       "WeBike",     81, 100, 100, "Mountain", 73),
    ("Swift Bike",      "WeBike",     76,  94, 100, "Speed", 72),
    ("Blu Ruged Ballz", "BB LLC",     81, 100, 100, "Mountain", 73),
    ("Blu Aero Ballz",  "BB LLC",     70,  86, 100, "Speed", 74),
    ("Blu Tail Ballz",  "BB LLC",     81, 100, 100, "Mountain", 70),
    ("Blu Tube Ballz",  "BB LLC",     70,  86, 100, "Speed", 76),
    ("Spoke'd Easy",    "Spoke'd Up", 100, 100, 100, "Recreation", 73),
    ("Spoke'd Speed",   "Spoke'd Up",  73,  91, 100, "Speed", 75),
    ("The Armstrong",   "SpaceBikes",  70,  86, 100, "Speed", 76),
    ("Mars Rover",      "SpaceBikes", 100, 100, 100, "Recreation", 73),
    ("Whole MILC MKII", "MILC Bikes", 100, 100, 100, "Recreation", 74),
    ("Skim MILC MKII",  "MILC Bikes",  76,  94, 100, "Speed", 76),
]

# Recreation judgment is the most price-sensitive column, so it ranks price cleanly.
tiers = {}
for b, co, rec, mtn, spd, seg, bj in PJ:
    tiers.setdefault((rec, mtn, spd), []).append(f"{b} ({co})")

print("INFERRED PRICE TIERS (cheapest -> most expensive, by Recreation price judgment)")
for k in sorted(tiers, key=lambda x: -x[0]):
    print(f"  Rec {k[0]:>3} / Mtn {k[1]:>3} / Speed {k[2]:>3}   {', '.join(tiers[k])}")

print("\nSEGMENT PRICE RESISTANCE BY TIER")
print("  tier(Rec)  Mountain  Speed")
seen = set()
for b, co, rec, mtn, spd, seg, bj in sorted(PJ, key=lambda x: -x[2]):
    if rec in seen:
        continue
    seen.add(rec)
    print(f"  {rec:>9}  {mtn:>8}  {spd:>5}")

print("\nKEY READS")
print("  Speed stays at 100 all the way down to tier 70; only tier 63 (MACH I.I) drops to 90.")
print("  Mountain stays at 100 only to tier 81 - Hike Bike sits exactly at that edge.")
print("  Recreation is the most price-sensitive column: only the 3 Rec brands reach 100.")
print("\n  Hike Bike  tier 81, Mountain 100 -> AT THE TOP of zero-resistance. Hold price.")
print("  Swift Bike tier 76, Speed 100    -> tier 70 brands ALSO score 100 -> HEADROOM to raise.")

wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
bad = PatternFill("solid", fgColor="FFC7CE")
good = PatternFill("solid", fgColor="C6EFCE")
yellow = PatternFill("solid", fgColor="FFEB9C")
ours = PatternFill("solid", fgColor="FFF2CC")

if "Q3_Price_Judgment" in wb.sheetnames:
    del wb["Q3_Price_Judgment"]
ws = wb.create_sheet("Q3_Price_Judgment")

ws["A1"] = "Price Judgment — World Market, Q3 actual"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("100 = no price resistance. Lower = buyers find the price high. Because the score depends only on "
            "price vs segment willingness-to-pay, identical triplets imply identical prices - which lets us "
            "rank every rival's price without seeing the price list.")
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
rows = [["Brand", "Company", "Recreation", "Mountain", "Speed", "Targets",
         "Price judgment in target", "Resistance?"]]
for b, co, rec, mtn, spd, seg, bj in PJ:
    tgt = {"Recreation": rec, "Mountain": mtn, "Speed": spd}[seg]
    rows.append([b, co, rec, mtn, spd, seg, tgt, "none" if tgt == 100 else f"-{100-tgt}"])
r2 = block(r, rows)
for i in range(1, len(rows)):
    if rows[i][1] == "WeBike":
        for c in range(1, 9):
            ws.cell(r + i, c).fill = ours
            ws.cell(r + i, c).font = bold
    ws.cell(r + i, 7).fill = good if rows[i][6] == 100 else bad
r = r2

rows = [["INFERRED PRICE TIERS (cheapest first)", "Rec", "Mountain", "Speed", "Brands at this price"]]
for k in sorted(tiers, key=lambda x: -x[0]):
    rows.append([f"tier {k[0]}", k[0], k[1], k[2], ", ".join(tiers[k])])
r2 = block(r, rows, wrap_col=5)
for i in range(1, len(rows)):
    if "WeBike" in str(rows[i][4]):
        for c in range(1, 6):
            ws.cell(r + i, c).fill = ours
r = r2

rows = [
    ["SEGMENT PRICE SENSITIVITY — resistance sets in at different tiers", "Mountain", "Speed", "Recreation"],
    ["tier 100 (cheapest)", 100, 100, 100],
    ["tier 92", 100, 100, 92],
    ["tier 85", 100, 100, 85],
    ["tier 83", 100, 100, 83],
    ["tier 82", 100, 100, 82],
    ["tier 81  <- Hike Bike", 100, 100, 81],
    ["tier 76  <- Swift Bike", 94, 100, 76],
    ["tier 73", 91, 100, 73],
    ["tier 70", 86, 100, 70],
    ["tier 63 (most expensive)", 78, 90, 63],
    ["Read", None, None, "SPEED is the least price-sensitive segment in the game - it tolerates every price "
     "up to tier 70 with zero resistance. MOUNTAIN starts resisting immediately below tier 81. "
     "RECREATION resists almost everything."],
]
r2 = block(r, rows)
ws.cell(r + 6, 2).fill = good
ws.cell(r + 7, 3).fill = good
ws.cell(r + 9, 3).fill = good
ws.cell(r + 10, 3).fill = bad
ws.cell(r + 11, 4).alignment = Alignment(wrap_text=True)
ws.cell(r + 11, 4).fill = yellow
r = r2

rows = [
    ["HIKE BIKE — price is already optimal, hold it", "Value", "Note"],
    ["Price", 1365, None],
    ["Mountain price judgment", 100, "ZERO price resistance"],
    ["Price tier", 81, "the HIGHEST-priced tier that still scores 100 in Mountain"],
    ["Next tier up (73)", 91, "TERRAMAX sits there and takes a 9-point hit"],
    ["Verdict", None, "We are extracting the maximum price Mountain will bear without penalty. "
     "Raising it drops us into the resistance band. HOLD $1,365."],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 2, 2).fill = good
ws.cell(r + 5, 3).fill = good
r = r2

rows = [
    ["SWIFT BIKE — we are UNDER-priced in Speed", "Value", "Note"],
    ["Price", 1450, None],
    ["Speed price judgment", 100, "zero resistance"],
    ["Our price tier", 76, None],
    ["Tier 70 brands (Armstrong, Blu Tube, Blu Aero, Mach 0.6)", 100, "priced ABOVE us and STILL score 100 in Speed"],
    ["First tier with Speed resistance", 63, "only MACH I.I, at 90"],
    ["Headroom", None, "roughly 1-2 price tiers before Speed resistance appears"],
    ["Estimated tier-70 price", 1595, "INFERRED - see the estimate note below"],
    ["Potential price increase", "=C{0}-C{1}".format(r + 7, r + 1), None],
    ["On Q3 volume (215 units)", "=C{0}*215".format(r + 8), "pure margin - no cost change"],
    ["On a rebuilt Speed volume (say 400 units)", "=C{0}*400".format(r + 8), None],
    ["CAVEAT", None, "The $1,595 figure is inferred from BB LLC's ~$1,480 average price across 4 brands, "
     "two of which appear to sit at our $1,365 tier. VERIFY against the actual price list before locking. "
     "The DIRECTION (we have room to raise) is solid; the exact number is not."],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 4, 2).fill = good
ws.cell(r + 9, 2).fill = good
ws.cell(r + 10, 2).fill = good
ws.cell(r + 11, 3).fill = yellow
ws.cell(r + 11, 3).alignment = Alignment(wrap_text=True)
r = r2

rows = [
    ["THE BB LLC COMPARISON IS NOW AIRTIGHT", "Hike Bike (WeBike)", "Blu Ruged Ballz (BB LLC)"],
    ["Brand judgment - Mountain", 73, 73],
    ["Component recipe", "56/73/1", "56/73/1 (identical)"],
    ["Price judgment", "81 / 100 / 100", "81 / 100 / 100 (identical)"],
    ["Implied price", "$1,365", "$1,365 (same tier)"],
    ["Mountain share", 0.221, 0.463],
    ["Conclusion", None, "Same product, same price, same price perception - and they take 2.1x our share. "
     "There is now no product or pricing explanation left. The entire gap is SUPPLY (we stocked out 33%), "
     "COVERAGE (14 store staff vs 14 but 7 web vs our 3) and ADS (24 regional vs our 19)."],
]
r2 = block(r, rows)
ws.cell(r + 5, 2).number_format = "0.0%"
ws.cell(r + 5, 3).number_format = "0.0%"
ws.cell(r + 5, 2).fill = bad
ws.cell(r + 5, 3).fill = good
ws.cell(r + 6, 3).alignment = Alignment(wrap_text=True)
ws.cell(r + 6, 3).fill = yellow
r = r2

rows = [
    ["ANOTHER REASON TO SKIP RECREATION", "Value", "Note"],
    ["Hike Bike Recreation price judgment", 81, "Rec buyers resist our $1,365"],
    ["All three Recreation brands", 100, "Spoke'd Easy, Mars Rover, Whole MILC - all at tier 100"],
    ["Read", None, "Recreation is the most price-sensitive segment AND the winning brands are the cheapest "
     "in the game. Entering would mean a low-price, low-margin product - the opposite of our "
     "'profit margin leader' strategy. Combined with the 56-vs-76 brand judgment gap, Recreation stays parked."],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 1, 2).fill = bad
ws.cell(r + 3, 3).alignment = Alignment(wrap_text=True)
r = r2

rows = [
    ["Q4 PRICING ACTIONS", "Decision"],
    ["Hike Bike", "HOLD $1,365. Mountain judgment 100 and we sit at the top tier that achieves it."],
    ["Swift Bike", "RAISE. Speed tolerates 1-2 tiers above us with zero resistance. Verify the exact price "
     "list, then move toward the Armstrong / Blu Tube tier."],
    ["Sequencing note", "Raise the Swift Bike price in the SAME quarter we fix its design to 77 - better "
     "product plus a price the segment does not resist is the cleanest margin gain available."],
    ["Capacity note", "A price rise also RELIEVES the capacity squeeze: more margin per unit on the same "
     "24/day. If we cannot build 1,171 units, we should at least earn more on the ones we do build."],
]
block(r, rows, wrap_col=2)
for i in range(1, len(rows)):
    ws.cell(r + i, 2).fill = yellow

ws.column_dimensions["A"].width = 52
ws.column_dimensions["B"].width = 22
ws.column_dimensions["C"].width = 30
ws.column_dimensions["D"].width = 60
ws.column_dimensions["E"].width = 56
for col in "FGH":
    ws.column_dimensions[col].width = 16

wb.save(DST)
print("\nAdded Q3_Price_Judgment")
