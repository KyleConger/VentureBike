"""Add Q3 ad judgment, verify the Marketing Effectiveness formula, and size the Swift Bike ad fix."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

DST = "quarters/Q3/Q3Data.xlsx"

# ad, company, brand, rec, mtn, speed
ADS = [
    ("Excluspeedity",    "Bike Bros",  "MACH I.I",         0, 12, 72),
    ("TerraTech",        "Bike Bros",  "TERRAMAX",        23, 78, 30),
    ("Easy Rider",       "Bike Bros",  "MountainCruise1", 75, 25, 17),
    ("Mountainability",  "Bike Bros",  "TERRAMAX",        35, 77, 32),
    ("AndStill",         "Bike Bros",  "MACH I.I",        22, 24, 78),
    ("Speed of Lite 1",  "LiteCycle",  "LiteSpeed Pro+",  10, 22, 77),
    ("Speed of Lite 2",  "LiteCycle",  "LiteSpeed+",      16, 13, 73),
    ("Trail Blazing 1",  "LiteCycle",  "LiteTrail Pro",   18, 79, 34),
    ("HikeBike 1",       "WeBike",     "Hike Bike",       33, 80, 26),
    ("Swift Bike",       "WeBike",     "Swift Bike",      52, 36, 70),
    ("BB Big Momma",     "BB LLC",     "Blu Ruged Ballz", 28, 81, 27),
    ("Unleash lil pap",  "BB LLC",     "Blu Tube Ballz",  15, 16, 76),
    ("Spoke'd Easy",     "Spoke'd Up", "Spoke'd Easy",    59, 27, 21),
    ("Spoke'd Speed",    "Spoke'd Up", "Spoke'd Speed",    4, 18, 57),
    ("The Armstrong 1",  "SpaceBikes", "The Armstrong",    2, 21, 77),
    ("Mars Rover 1",     "SpaceBikes", "Mars Rover",      79, 25, 24),
    ("Skim MILC MKII",   "MILC Bikes", "Skim MILC MKII",  10, 19, 75),
    ("Whole MILC MKII",  "MILC Bikes", "Whole MILC MKII", 77, 47, 41),
]
BRAND_SEG = {
    "MACH I.I": "Speed", "TERRAMAX": "Mountain", "MountainCruise1": "Recreation",
    "LiteSpeed Pro+": "Speed", "LiteSpeed+": "Speed", "LiteTrail Pro": "Mountain",
    "Hike Bike": "Mountain", "Swift Bike": "Speed",
    "Blu Ruged Ballz": "Mountain", "Blu Tube Ballz": "Speed",
    "Spoke'd Easy": "Recreation", "Spoke'd Speed": "Speed",
    "The Armstrong": "Speed", "Mars Rover": "Recreation",
    "Skim MILC MKII": "Speed", "Whole MILC MKII": "Recreation",
}
IDX = {"Recreation": 3, "Mountain": 4, "Speed": 5}

# Marketing Effectiveness = (avg brand judgment/100 + avg ad judgment/100) / 2, over targeted segments.
# firm -> (primary seg, brand, brand judgment, ad, ad judgment), (secondary ...)
ME_IN = {
    "BB LLC":     [("Mountain", "Blu Ruged Ballz", 73, "BB Big Momma", 81),
                   ("Speed", "Blu Tube Ballz", 76, "Unleash lil pap", 76)],
    "SpaceBikes": [("Recreation", "Mars Rover", 73, "Mars Rover 1", 79),
                   ("Speed", "The Armstrong", 76, "The Armstrong 1", 77)],
    "MILC Bikes": [("Recreation", "Whole MILC MKII", 74, "Whole MILC MKII", 77),
                   ("Speed", "Skim MILC MKII", 76, "Skim MILC MKII", 75)],
    "WeBike":     [("Mountain", "Hike Bike", 73, "HikeBike 1", 80),
                   ("Speed", "Swift Bike", 72, "Swift Bike", 70)],
    "Spoke'd Up": [("Recreation", "Spoke'd Easy", 73, "Spoke'd Easy", 59),
                   ("Speed", "Spoke'd Speed", 75, "Spoke'd Speed", 57)],
}
REPORTED = {"WeBike": 0.738, "Spoke'd Up": 0.660, "BB LLC": 0.765}


def me(entries):
    b = sum(e[2] for e in entries) / len(entries)
    a = sum(e[4] for e in entries) / len(entries)
    return (b / 100 + a / 100) / 2, b, a


print("MARKETING EFFECTIVENESS FORMULA CHECK")
print("  ME = (avg brand judgment/100 + avg ad judgment/100) / 2")
for co, e in sorted(ME_IN.items(), key=lambda x: -me(x[1])[0]):
    v, b, a = me(e)
    tag = ""
    if co in REPORTED:
        tag = f"   reported {REPORTED[co]:.3f}  {'MATCH' if abs(v-REPORTED[co]) < 0.0015 else 'MISMATCH'}"
    print(f"  {co:<12} brands {b:5.1f}  ads {a:5.1f}  ME {v:.4f}{tag}")
print("\n  Industry reported min 0.660 (Spoke'd Up) and max 0.765 (BB LLC) both reproduce exactly.\n")

print("SPEED AD LADDER")
for ad, co, br, rec, mtn, spd in sorted([a for a in ADS if BRAND_SEG[a[2]] == "Speed"], key=lambda x: -x[5]):
    print(f"  Speed {spd:>3}  {ad:<17}{co:<12} off-target leakage (Rec+Mtn) {rec+mtn:>3}")

print("\nMOUNTAIN AD LADDER")
for ad, co, br, rec, mtn, spd in sorted([a for a in ADS if BRAND_SEG[a[2]] == "Mountain"], key=lambda x: -x[4]):
    print(f"  Mtn   {mtn:>3}  {ad:<17}{co:<12} off-target leakage (Rec+Speed) {rec+spd:>3}")

print("\nWHAT FIXING SWIFT BIKE DOES TO MARKETING EFFECTIVENESS")
base = me(ME_IN["WeBike"])[0]
scen = [
    ("Q3 actual (brand 72, ad 70)", 72, 70),
    ("Fix ad only (design to Speed focus)", 72, 77),
    ("Fix brand only (14-speed + decals)", 77, 70),
    ("Fix BOTH", 77, 77),
]
for name, bj, aj in scen:
    e = [ME_IN["WeBike"][0], ("Speed", "Swift Bike", bj, "Swift Bike", aj)]
    v = me(e)[0]
    note = "  <-- BEATS the industry max of 0.765" if v > 0.765 else ""
    print(f"  {name:<38} ME {v:.4f}{note}")

wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
bad = PatternFill("solid", fgColor="FFC7CE")
good = PatternFill("solid", fgColor="C6EFCE")
yellow = PatternFill("solid", fgColor="FFEB9C")
ours = PatternFill("solid", fgColor="FFF2CC")

if "Q3_Ad_Judgment" in wb.sheetnames:
    del wb["Q3_Ad_Judgment"]
ws = wb.create_sheet("Q3_Ad_Judgment")

ws["A1"] = "Ad Judgment — World Market, Q3 actual"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("HikeBike 1 is essentially best-in-class (80 vs BB LLC's 81). The Swift Bike ad is 8th of 9 "
            "Speed ads and shows by far the highest off-target appeal in the industry.")
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
rows = [["Ad", "Company", "Brand advertised", "Recreation", "Mountain", "Speed",
         "Target segment", "Score in target", "Off-target leakage"]]
for ad, co, br, rec, mtn, spd in ADS:
    seg = BRAND_SEG[br]
    vals = {"Recreation": rec, "Mountain": mtn, "Speed": spd}
    leak = sum(v for k, v in vals.items() if k != seg)
    rows.append([ad, co, br, rec, mtn, spd, seg, vals[seg], leak])
r2 = block(r, rows)
for i in range(1, len(rows)):
    if rows[i][1] == "WeBike":
        for c in range(1, 10):
            ws.cell(r + i, c).fill = ours
            ws.cell(r + i, c).font = bold
r = r2

rows = [["SPEED AD LADDER", "Speed score", "Off-target leakage", "Note"]]
for ad, co, br, rec, mtn, spd in sorted([a for a in ADS if BRAND_SEG[a[2]] == "Speed"], key=lambda x: -x[5]):
    note = ""
    if ad == "AndStill":
        note = "BEST Speed ad in the industry"
    elif ad == "Swift Bike":
        note = "OURS - 8th of 9"
    elif ad == "Spoke'd Speed":
        note = "worst; drags Spoke'd Up to last on Marketing Effectiveness"
    rows.append([f"{ad} ({co})", spd, rec + mtn, note])
r2 = block(r, rows, wrap_col=4)
for i in range(1, len(rows)):
    if "WeBike" in str(rows[i][0]):
        for c in range(1, 5):
            ws.cell(r + i, c).fill = ours
            ws.cell(r + i, c).font = bold
        ws.cell(r + i, 2).fill = bad
        ws.cell(r + i, 3).fill = bad
r = r2

rows = [["MOUNTAIN AD LADDER", "Mountain score", "Off-target leakage", "Note"]]
for ad, co, br, rec, mtn, spd in sorted([a for a in ADS if BRAND_SEG[a[2]] == "Mountain"], key=lambda x: -x[4]):
    note = "BEST" if ad == "BB Big Momma" else ("OURS - 1 point off best" if ad == "HikeBike 1" else "")
    rows.append([f"{ad} ({co})", mtn, rec + spd, note])
r2 = block(r, rows, wrap_col=4)
for i in range(1, len(rows)):
    if "WeBike" in str(rows[i][0]):
        for c in range(1, 5):
            ws.cell(r + i, c).fill = ours
            ws.cell(r + i, c).font = bold
        ws.cell(r + i, 2).fill = good
r = r2

rows = [
    ["MARKETING EFFECTIVENESS — formula verified against 3 reported values", "Avg brand", "Avg ad", "ME", "Reported"],
]
for co, e in sorted(ME_IN.items(), key=lambda x: -me(x[1])[0]):
    v, b, a = me(e)
    rows.append([co, b, a, v, REPORTED.get(co, "")])
rows.append(["Result", None, None, None, "Industry min 0.660 (Spoke'd Up) and max 0.765 (BB LLC) both "
             "reproduce exactly, as does our 0.738. Formula confirmed."])
r2 = block(r, rows, wrap_col=5)
for i in range(1, len(rows) - 1):
    ws.cell(r + i, 4).number_format = "0.0000"
    if rows[i][0] == "WeBike":
        for c in range(1, 6):
            ws.cell(r + i, c).fill = ours
            ws.cell(r + i, c).font = bold
ws.cell(r + len(rows) - 1, 5).fill = good
r = r2

rows = [
    ["FIXING SWIFT BIKE — the Marketing Effectiveness payoff", "Brand judgment", "Ad judgment", "ME", "vs industry max 0.765"],
    ["Q3 actual", 72, 70, 0.7375, "below average (0.743)"],
    ["Fix ad only", 72, 77, 0.7550, "3rd best"],
    ["Fix brand only (14-speed + decals)", 77, 70, 0.7500, "4th best"],
    ["FIX BOTH", 77, 77, 0.7675, "NEW INDUSTRY BEST"],
]
r2 = block(r, rows, wrap_col=5)
for i in range(1, 5):
    ws.cell(r + i, 4).number_format = "0.0000"
ws.cell(r + 4, 4).fill = good
ws.cell(r + 4, 5).fill = good
ws.cell(r + 4, 4).font = Font(bold=True)
ws.cell(r + 1, 4).fill = bad
r = r2

rows = [
    ["DIAGNOSIS — the Swift Bike ad is unfocused", "Value", "Note"],
    ["Our Swift Bike ad: Recreation score", 52, "highest Rec score of ANY Speed ad"],
    ["Our Swift Bike ad: Mountain score", 36, "highest Mtn score of ANY Speed ad"],
    ["Total off-target leakage", 88, "next highest is AndStill at 46; The Armstrong 1 is just 23"],
    ["Leading Speed ads for comparison", "Rec 2-22, Mtn 13-24", "The Armstrong 1 scores Rec 2 - almost zero"],
    ["Hypothesis", None, "The ad appears to carry Mountain- and Recreation-flavoured benefit claims rather "
     "than Speed-specific ones. This MIRRORS the component error: Swift Bike inherited Hike Bike's 24-speed "
     "drivetrain AND, it seems, Hike Bike's messaging. Same root cause - we cloned our Mountain assets into "
     "a Speed launch."],
    ["CAUTION", None, "Leakage does NOT cleanly predict score across the field (AndStill leaks 46 and scores "
     "best), so this is a strong hypothesis, not a solved model like the component analysis. Rebuild the ad "
     "around Speed cues and re-test in the ad designer."],
]
r2 = block(r, rows, wrap_col=3)
ws.cell(r + 3, 2).fill = bad
ws.cell(r + 5, 3).fill = yellow
ws.cell(r + 6, 3).fill = yellow
for i in (5, 6):
    ws.cell(r + i, 3).alignment = Alignment(wrap_text=True)
r = r2

rows = [
    ["AD PORTFOLIO PATTERNS", "Ads", "Brands", "Note"],
    ["Bike Bros", 5, 5, "TWO ads each for MACH I.I and TERRAMAX - A/B testing. Still finished LAST."],
    ["LiteCycle", 3, 3, "one ad per brand"],
    ["BB LLC", 2, 4, "advertises only Blu Ruged Ballz + Blu Tube Ballz; Blu Aero and Blu Tail run with NO ad"],
    ["WeBike", 2, 2, "one per brand"],
    ["SpaceBikes / MILC / Spoke'd Up", 2, 2, "one per brand"],
    ["Read", None, None, "The LEADER runs the FEWEST ads relative to its brand count - 2 ads for 4 brands - "
     "and both are best or near-best in class. Bike Bros runs the most ads and is last. Ad QUALITY beats ad "
     "COUNT. A second Swift Bike ad is not the answer; a better one is."],
    ["Portfolio idea worth noting", None, None, "BB LLC pairs an ADVERTISED brand with an UNADVERTISED "
     "second brand in each segment (Blu Ruged + Blu Tail in Mountain, Blu Tube + Blu Aero in Speed). The ad "
     "builds the segment and the second brand catches overflow at a higher price. Not a Q4 move for us - "
     "capacity binds - but note it for later."],
]
r2 = block(r, rows, wrap_col=4)
ws.cell(r + 6, 4).fill = yellow
for i in (6, 7):
    ws.cell(r + i, 4).alignment = Alignment(wrap_text=True)
r = r2

rows = [
    ["Q4 ADVERTISING ACTIONS", "Decision"],
    ["HikeBike 1", "KEEP. 80 vs BB LLC's 81 - one point off the industry best. Redesign cost is not "
     "justified; put the money into media inserts and capacity instead."],
    ["Swift Bike ad", "REDESIGN around Speed cues. Strip Mountain/Recreation claims (aerodynamic, racing "
     "tires, precision brakes, drop bars, lightweight). Target 77 to match Speed of Lite 1 / The Armstrong 1."],
    ["Do NOT add more ads", "The leader runs 2 ads for 4 brands; the last-place firm runs 5 for 5. "
     "Quality over count."],
    ["Combined payoff", "Swift Bike brand 72->77 AND ad 70->77 lifts Marketing Effectiveness from 0.738 to "
     "0.7675 - past BB LLC's 0.765 to become the industry best. Both fixes are cheap relative to capacity."],
]
block(r, rows, wrap_col=2)
for i in range(1, len(rows)):
    ws.cell(r + i, 2).fill = yellow

ws.column_dimensions["A"].width = 40
ws.column_dimensions["B"].width = 22
ws.column_dimensions["C"].width = 22
ws.column_dimensions["D"].width = 62
ws.column_dimensions["E"].width = 40
for col in "FGHI":
    ws.column_dimensions[col].width = 15

wb.save(DST)
print("\nAdded Q3_Ad_Judgment")
