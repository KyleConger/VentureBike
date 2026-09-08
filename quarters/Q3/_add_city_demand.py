"""Add Q3 city x brand x segment demand detail and the city #3 analysis."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

DST = "quarters/Q3/Q3Data.xlsx"

# (brand, company, city, rec, mtn, speed)
D = [
    ("LiteSpeed Pro+", "LiteCycle", "New York City", 0, 0, 39),
    ("LiteSpeed+", "LiteCycle", "New York City", 0, 0, 29),
    ("LiteTrail Pro", "LiteCycle", "New York City", 13, 62, 0),
    ("Hike Bike", "WeBike", "New York City", 3, 23, 0),
    ("Swift Bike", "WeBike", "New York City", 0, 0, 12),
    ("Blu Ruged Ballz", "BB LLC", "New York City", 19, 265, 0),
    ("Blu Aero Ballz", "BB LLC", "New York City", 0, 0, 135),
    ("Blu Tail Ballz", "BB LLC", "New York City", 22, 145, 0),
    ("Blu Tube Ballz", "BB LLC", "New York City", 0, 0, 166),
    ("Spoke'd Easy", "Spoke'd Up", "New York City", 237, 0, 0),
    ("Spoke'd Speed", "Spoke'd Up", "New York City", 0, 0, 175),
    ("The Armstrong", "SpaceBikes", "New York City", 0, 0, 272),
    ("Mars Rover", "SpaceBikes", "New York City", 302, 0, 0),
    ("Whole MILC MKII", "MILC Bikes", "New York City", 68, 0, 0),
    ("Skim MILC MKII", "MILC Bikes", "New York City", 0, 0, 52),

    ("LiteSpeed Pro+", "LiteCycle", "Rio de Janeiro", 0, 0, 27),
    ("LiteSpeed+", "LiteCycle", "Rio de Janeiro", 0, 0, 20),
    ("LiteTrail Pro", "LiteCycle", "Rio de Janeiro", 13, 46, 0),
    ("Hike Bike", "WeBike", "Rio de Janeiro", 25, 166, 0),
    ("Swift Bike", "WeBike", "Rio de Janeiro", 0, 0, 79),
    ("Blu Ruged Ballz", "BB LLC", "Rio de Janeiro", 6, 48, 0),
    ("Blu Aero Ballz", "BB LLC", "Rio de Janeiro", 0, 0, 23),
    ("Blu Tail Ballz", "BB LLC", "Rio de Janeiro", 7, 26, 0),
    ("Blu Tube Ballz", "BB LLC", "Rio de Janeiro", 0, 0, 29),
    ("Spoke'd Easy", "Spoke'd Up", "Rio de Janeiro", 21, 0, 0),
    ("Spoke'd Speed", "Spoke'd Up", "Rio de Janeiro", 0, 0, 11),
    ("The Armstrong", "SpaceBikes", "Rio de Janeiro", 0, 0, 43),
    ("Mars Rover", "SpaceBikes", "Rio de Janeiro", 77, 0, 0),
    ("Whole MILC MKII", "MILC Bikes", "Rio de Janeiro", 70, 0, 0),
    ("Skim MILC MKII", "MILC Bikes", "Rio de Janeiro", 0, 0, 37),

    ("MACH I.I", "Bike Bros", "Amsterdam", 0, 0, 155),
    ("TERRAMAX", "Bike Bros", "Amsterdam", 6, 92, 0),
    ("Mach 0.6", "Bike Bros", "Amsterdam", 0, 0, 25),
    ("TERRAMean", "Bike Bros", "Amsterdam", 10, 80, 0),
    ("MountainCruise1", "Bike Bros", "Amsterdam", 122, 0, 0),
    ("LiteSpeed Pro+", "LiteCycle", "Amsterdam", 0, 0, 126),
    ("LiteSpeed+", "LiteCycle", "Amsterdam", 0, 0, 93),
    ("LiteTrail Pro", "LiteCycle", "Amsterdam", 36, 110, 0),
    ("Hike Bike", "WeBike", "Amsterdam", 24, 156, 0),
    ("Swift Bike", "WeBike", "Amsterdam", 0, 0, 115),
    ("Blu Ruged Ballz", "BB LLC", "Amsterdam", 16, 137, 0),
    ("Blu Aero Ballz", "BB LLC", "Amsterdam", 0, 0, 110),
    ("Blu Tail Ballz", "BB LLC", "Amsterdam", 19, 75, 0),
    ("Blu Tube Ballz", "BB LLC", "Amsterdam", 0, 0, 136),
    ("Spoke'd Easy", "Spoke'd Up", "Amsterdam", 217, 0, 0),
    ("Spoke'd Speed", "Spoke'd Up", "Amsterdam", 0, 0, 163),
    ("The Armstrong", "SpaceBikes", "Amsterdam", 0, 0, 223),
    ("Mars Rover", "SpaceBikes", "Amsterdam", 267, 0, 0),
    ("Whole MILC MKII", "MILC Bikes", "Amsterdam", 275, 0, 0),
    ("Skim MILC MKII", "MILC Bikes", "Amsterdam", 0, 0, 197),

    ("LiteSpeed Pro+", "LiteCycle", "Bangalore", 0, 0, 123),
    ("LiteSpeed+", "LiteCycle", "Bangalore", 0, 0, 91),
    ("LiteTrail Pro", "LiteCycle", "Bangalore", 34, 122, 0),
    ("Hike Bike", "WeBike", "Bangalore", 2, 13, 0),
    ("Swift Bike", "WeBike", "Bangalore", 0, 0, 9),
    ("Blu Ruged Ballz", "BB LLC", "Bangalore", 5, 36, 0),
    ("Blu Aero Ballz", "BB LLC", "Bangalore", 0, 0, 26),
    ("Blu Tail Ballz", "BB LLC", "Bangalore", 5, 19, 0),
    ("Blu Tube Ballz", "BB LLC", "Bangalore", 0, 0, 32),
    ("Spoke'd Easy", "Spoke'd Up", "Bangalore", 17, 0, 0),
    ("Spoke'd Speed", "Spoke'd Up", "Bangalore", 0, 0, 13),
    ("The Armstrong", "SpaceBikes", "Bangalore", 0, 0, 50),
    ("Mars Rover", "SpaceBikes", "Bangalore", 64, 0, 0),
    ("Whole MILC MKII", "MILC Bikes", "Bangalore", 58, 0, 0),
    ("Skim MILC MKII", "MILC Bikes", "Bangalore", 0, 0, 42),
]

CITIES = ["Amsterdam", "New York City", "Rio de Janeiro", "Bangalore"]
FIRMS = ["BB LLC", "SpaceBikes", "LiteCycle", "Spoke'd Up", "MILC Bikes", "WeBike", "Bike Bros"]

# ---- reconciliation ----
tot = sum(r + m + s for _, _, _, r, m, s in D)
we = sum(r + m + s for _, co, _, r, m, s in D if co == "WeBike")
hb = sum(r + m + s for b, _, _, r, m, s in D if b == "Hike Bike")
sb = sum(r + m + s for b, _, _, r, m, s in D if b == "Swift Bike")
print("RECONCILIATION")
print(f"  total demand  {tot:>6}  (report says 6559)  {'OK' if tot == 6559 else 'MISMATCH'}")
print(f"  WeBike demand {we:>6}  (report says  627)  {'OK' if we == 627 else 'MISMATCH'}")
print(f"  Hike Bike     {hb:>6}  (report says  412)  {'OK' if hb == 412 else 'MISMATCH'}")
print(f"  Swift Bike    {sb:>6}  (report says  215)  {'OK' if sb == 215 else 'MISMATCH'}")
assert (tot, we, hb, sb) == (6559, 627, 412, 215)
print("  all four reconcile exactly\n")


def city_seg(city, idx):
    return sum(row[3 + idx] for row in D if row[2] == city)


def firm_city(firm, city):
    return sum(r + m + s for _, co, ci, r, m, s in D if co == firm and ci == city)


def firm_city_seg(firm, city, idx):
    return sum(row[3 + idx] for row in D if row[1] == firm and row[2] == city)


wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
bad = PatternFill("solid", fgColor="FFC7CE")
good = PatternFill("solid", fgColor="C6EFCE")
yellow = PatternFill("solid", fgColor="FFEB9C")
ours = PatternFill("solid", fgColor="FFF2CC")

for n in ("Q3_City_Demand", "Q3_City_Strategy"):
    if n in wb.sheetnames:
        del wb[n]

# ================= Q3_City_Demand (raw) =================
ws = wb.create_sheet("Q3_City_Demand")
ws["A1"] = "Detailed Brand Demand — World Market by city, Q3 actual"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = "Reconciles exactly to 6,559 total / 627 WeBike / 412 Hike Bike / 215 Swift Bike."
r = 4
for c, v in enumerate(["Brand", "Company", "City", "Recreation", "Mountain", "Speed", "Total"], start=1):
    cell = ws.cell(r, c, v)
    cell.fill = hdr
    cell.font = bold
r += 1
for brand, co, city, rec, mtn, spd in D:
    vals = [brand, co, city, rec, mtn, spd, rec + mtn + spd]
    for c, v in enumerate(vals, start=1):
        cell = ws.cell(r, c, v)
        if co == "WeBike":
            cell.fill = ours
            cell.font = bold
    r += 1
ws.column_dimensions["A"].width = 20
ws.column_dimensions["B"].width = 13
ws.column_dimensions["C"].width = 16
ws.freeze_panes = "A5"

# ================= Q3_City_Strategy =================
ws = wb.create_sheet("Q3_City_Strategy")
ws["A1"] = "City-level competitive map and the City #3 decision"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("Amsterdam + New York are 77% of the entire market. We hold 9.9% of Amsterdam and 1.9% of New York, "
            "but 34.9% of Rio - the one city where we are the only firm with a store.")
ws["A2"].alignment = Alignment(wrap_text=True)


def block(start, rows, wrap_col=None, width=None):
    for i, row in enumerate(rows):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(start + i, c, v)
            if i == 0:
                cell.fill = hdr
                cell.font = bold
            elif wrap_col and c == wrap_col:
                cell.alignment = Alignment(wrap_text=True)
    return start + len(rows) + 2


# city size table
r = 4
rows = [["CITY SIZE & OUR POSITION", "Recreation", "Mountain", "Speed", "Total demand",
         "% of market", "WeBike units", "WeBike share"]]
for city in CITIES:
    t = city_seg(city, 0) + city_seg(city, 1) + city_seg(city, 2)
    wu = firm_city("WeBike", city)
    rows.append([city, city_seg(city, 0), city_seg(city, 1), city_seg(city, 2), t,
                 t / 6559, wu, wu / t])
rows.append(["TOTAL", city_seg("Amsterdam", 0) + city_seg("New York City", 0) + city_seg("Rio de Janeiro", 0) + city_seg("Bangalore", 0),
             sum(city_seg(c, 1) for c in CITIES), sum(city_seg(c, 2) for c in CITIES), 6559, 1.0, 627, 627 / 6559])
r2 = block(r, rows)
for i in range(1, 5):
    ws.cell(r + i, 6).number_format = "0.0%"
    ws.cell(r + i, 8).number_format = "0.0%"
ws.cell(r + 5, 6).number_format = "0.0%"
ws.cell(r + 5, 8).number_format = "0.0%"
ws.cell(r + 1, 8).fill = bad   # Amsterdam share
ws.cell(r + 2, 8).fill = bad   # NYC share
ws.cell(r + 3, 8).fill = good  # Rio share
r = r2

# share matrix
rows = [["SHARE BY CITY [%] — who owns what"] + CITIES]
for firm in FIRMS:
    row = [firm]
    for city in CITIES:
        t = sum(city_seg(city, i) for i in range(3))
        row.append(firm_city(firm, city) / t)
    rows.append(row)
r2 = block(r, rows)
for i in range(1, len(FIRMS) + 1):
    for c in range(2, 6):
        ws.cell(r + i, c).number_format = "0.0%"
        if rows[i][0] == "WeBike":
            ws.cell(r + i, c).fill = ours
r = r2

# store inference
rows = [
    ["INFERRED STORE LOCATIONS (from demand concentration)", "Firms with a store", "City demand", "Read"],
    ["Amsterdam", "ALL SEVEN", 2985, "Most contested city AND the largest. Nobody exceeds 16.5%."],
    ["New York City", "BB LLC, SpaceBikes, Spoke'd Up", 2039, "2nd largest. We have NO store -> only 1.9% share."],
    ["Rio de Janeiro", "WeBike ONLY", 774, "Uncontested by any rival store -> we take 34.9%."],
    ["Bangalore", "LiteCycle ONLY", 761, "Uncontested by others -> LiteCycle takes 48.6%."],
    ["Bike Bros", "Amsterdam only (1 store, 0 web)", None, "Appears in NO other city. Confirms the 1-store read."],
    ["CAUTION", None, None, "Store locations are INFERRED from demand patterns, not reported directly. "
     "MILC Bikes is Amsterdam-concentrated but its 2nd store is not identifiable."],
]
r2 = block(r, rows, wrap_col=4)
ws.cell(r + 3, 2).fill = good
ws.cell(r + 2, 2).fill = bad
ws.cell(r + len(rows) - 1, 1).fill = yellow
r = r2

# the store effect
rows = [
    ["THE STORE EFFECT — what a local store is worth", "Share with store", "Share without store"],
    ["WeBike in Rio (only store in city)", 0.349, None],
    ["LiteCycle in Bangalore (only store in city)", 0.486, None],
    ["WeBike in Amsterdam (store, 6 rivals)", 0.099, None],
    ["WeBike in New York City (no store)", None, 0.019],
    ["WeBike in Bangalore (no store)", None, 0.032],
    ["Read", None, None],
]
r2 = block(r, rows)
for i in (1, 2, 3):
    ws.cell(r + i, 2).number_format = "0.0%"
for i in (4, 5):
    ws.cell(r + i, 3).number_format = "0.0%"
ws.cell(r + 1, 2).fill = good
ws.cell(r + 2, 2).fill = good
ws.cell(r + 4, 3).fill = bad
ws.cell(r + 5, 3).fill = bad
ws.cell(r + 6, 2, "A store is worth roughly 10x the share of web-only coverage (2-3% without, 10-35% with). "
                  "An UNCONTESTED store is worth 35-49%. This is the highest-leverage decision we have.")
ws.cell(r + 6, 2).alignment = Alignment(wrap_text=True)
r = r2

# Mountain by city
rows = [["OUR PRIMARY SEGMENT (Mountain) BY CITY", "Segment size", "WeBike units", "WeBike share",
         "Leader", "Leader units", "Leader share"]]
mtn_leader = {
    "Amsterdam": ("WeBike", 156),
    "New York City": ("BB LLC (both brands)", 410),
    "Rio de Janeiro": ("WeBike", 166),
    "Bangalore": ("LiteCycle", 122),
}
for city in CITIES:
    size = city_seg(city, 1)
    wu = firm_city_seg("WeBike", city, 1)
    lname, lunits = mtn_leader[city]
    rows.append([city, size, wu, wu / size, lname, lunits, lunits / size])
r2 = block(r, rows)
for i in range(1, 5):
    ws.cell(r + i, 4).number_format = "0.0%"
    ws.cell(r + i, 7).number_format = "0.0%"
ws.cell(r + 3, 4).fill = good   # Rio
ws.cell(r + 2, 4).fill = bad    # NYC
r = r2

# City #3 decision
rows = [
    ["CITY #3 DECISION — New York City vs Bangalore", "New York City", "Bangalore"],
    ["Total city demand (Q3)", 2039, 761],
    ["Share of total market", "=B{0}/6559".format(r + 1), "=C{0}/6559".format(r + 1)],
    ["Rival stores already there", 3, 1],
    ["Our current share (web only)", 0.019, 0.032],
    ["Mountain segment size", 495, 190],
    ["Mountain leader hold", "BB LLC 82.8%", "LiteCycle 64.2%"],
    ["Speed segment size", 880, 386],
    ["Recreation segment size (untapped)", 664, 185],
    ["Units at a conservative 10% share", "=B{0}*0.10".format(r + 1), "=C{0}*0.10".format(r + 1)],
    ["Units at a Rio-like 25% share", "=B{0}*0.25".format(r + 1), "=C{0}*0.25".format(r + 1)],
    ["VERDICT", "LARGER PRIZE — 2.7x the demand. Even at a weak 10% it beats Bangalore at 25%. "
     "Fits our locked 'largest geos even if expensive' strategy.",
     "SAFER but SMALLER — only LiteCycle to fight, yet the whole city is worth less than "
     "NYC's Speed segment alone."],
]
r2 = block(r, rows, wrap_col=None)
ws.cell(r + 4, 2).number_format = "0.0%"
ws.cell(r + 4, 3).number_format = "0.0%"
ws.cell(r + 2, 2).number_format = "0.0%"
ws.cell(r + 2, 3).number_format = "0.0%"
for c in (2, 3):
    ws.cell(r + 11, c).alignment = Alignment(wrap_text=True)
ws.cell(r + 11, 2).fill = good
r = r2

# capacity check
rows = [
    ["CAPACITY CHECK — can we even serve a 3rd city?", "Units"],
    ["Fixed capacity per day", 24],
    ["Days per quarter", 65],
    ["Theoretical max units", "=B{0}*B{1}".format(r + 1, r + 2)],
    ["At 74% worker productivity", "=B{0}*0.74".format(r + 3)],
    ["Q4 forecast demand (existing cities, after ill will)", 656],
    ["Headroom for a new city", "=B{0}-B{1}".format(r + 4, r + 5)],
    ["Read", "We can absorb roughly 500 more units without buying a printer. A New York store at 10% "
     "(~204 units) fits comfortably. DO NOT repeat Q3 - schedule the operating capacity to match."],
]
r2 = block(r, rows)
ws.cell(r + 6, 2).fill = good
ws.cell(r + 7, 2).alignment = Alignment(wrap_text=True)
r = r2

# risks
rows = [
    ["RISKS & PROTECTIONS", "Note"],
    ["Rio is our crown jewel", "58.0% of Rio Mountain and 34.9% of the city. Hike Bike's single best market "
     "(166 Mountain units, more than Amsterdam's 156). The Q3 stock-out damaged our strongest position - "
     "ill will bites hardest exactly where we lead."],
    ["Amsterdam is the biggest miss", "2,985 units = 45.5% of the market and we hold only 9.9%. Adding sales "
     "coverage here may be cheaper per unit than a new city, and no rival exceeds 16.5% - it is genuinely open."],
    ["New York is BB LLC's fortress", "They take 36.9% of the city and 82.8% of its Mountain. Attack via SPEED "
     "(880 units, no single firm dominant) and Recreation, not head-on in Mountain."],
    ["Bangalore is LiteCycle's fortress", "48.6% city share, 64.2% of its Mountain - the most concentrated "
     "position any firm holds in any city."],
    ["Do not spread too thin", "Bike Bros ran 5 brands from 1 store and finished last. Our constraint is supply "
     "and coverage depth, not map coverage."],
]
r2 = block(r, rows, wrap_col=2)
ws.cell(r + 1, 2).fill = yellow

ws.column_dimensions["A"].width = 46
ws.column_dimensions["B"].width = 30
ws.column_dimensions["C"].width = 30
for col in "DEFGH":
    ws.column_dimensions[col].width = 16

wb.save(DST)
print("Added Q3_City_Demand and Q3_City_Strategy")
