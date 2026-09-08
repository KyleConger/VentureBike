"""Add Q3 component teardown of all 20 industry brands + reverse-engineered scoring model."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

DST = "quarters/Q3/Q3Data.xlsx"
wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
bad = PatternFill("solid", fgColor="FFC7CE")
good = PatternFill("solid", fgColor="C6EFCE")
yellow = PatternFill("solid", fgColor="FFEB9C")
ours = PatternFill("solid", fgColor="FFF2CC")

for n in ("Q3_Components", "Q3_Design_Model"):
    if n in wb.sheetnames:
        del wb[n]

BRANDS = [
    ("MACH I.I", "Bike Bros", "Speed", 77),
    ("TERRAMAX", "Bike Bros", "Mountain", 73),
    ("Mach 0.6", "Bike Bros", "Speed", 61),
    ("TERRAMean", "Bike Bros", "Mountain", 72),
    ("MountainCruise1", "Bike Bros", "Recreation", 76),
    ("LiteSpeed Pro+", "LiteCycle", "Speed", 77),
    ("LiteSpeed+", "LiteCycle", "Speed", 74),
    ("LiteTrail Pro", "LiteCycle", "Mountain", 70),
    ("Hike Bike", "WeBike", "Mountain", 73),
    ("Swift Bike", "WeBike", "Speed", 72),
    ("Blu Ruged Ballz", "BB LLC", "Mountain", 73),
    ("Blu Aero Ballz", "BB LLC", "Speed", 74),
    ("Blu Tail Ballz", "BB LLC", "Mountain", 70),
    ("Blu Tube Ballz", "BB LLC", "Speed", 76),
    ("Spoke'd Easy", "Spoke'd Up", "Recreation", 73),
    ("Spoke'd Speed", "Spoke'd Up", "Speed", 75),
    ("The Armstrong", "SpaceBikes", "Speed", 76),
    ("Mars Rover", "SpaceBikes", "Recreation", 73),
    ("Whole MILC MKII", "MILC Bikes", "Recreation", 74),
    ("Skim MILC MKII", "MILC Bikes", "Speed", 76),
]

# 1 = included, 0 = not. Order matches BRANDS.
COMPONENTS = [
    ("Frame", "Comfort (relaxed)",        [0,0,0,0,1,0,0,0,0,0,0,0,0,0,1,0,0,1,1,0]),
    ("Frame", "Rugged (rough terrain)",   [0,1,0,1,0,0,0,1,1,0,1,0,1,0,0,0,0,0,0,0]),
    ("Frame", "Aerodynamic (speed)",      [1,0,1,0,0,1,1,0,0,1,0,1,0,1,0,1,1,0,0,1]),
    ("Tires", "Mountain - high grip",     [0,1,0,1,0,0,0,1,1,0,1,0,1,0,0,0,0,0,0,0]),
    ("Tires", "Hybrid - road/off-road",   [0,0,0,0,1,0,0,0,0,0,0,0,0,0,1,0,0,1,1,0]),
    ("Tires", "Racing - fast",            [1,0,1,0,0,1,1,0,0,1,0,1,0,1,0,1,1,0,0,1]),
    ("Brakes", "Standard",                [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0]),
    ("Brakes", "Precision",               [1,0,1,0,0,1,1,0,0,1,0,1,0,1,1,1,1,0,1,1]),
    ("Brakes", "Standard disc",           [0,1,0,1,1,0,0,1,1,0,1,0,1,0,0,0,0,0,0,0]),
    ("Handlebars", "Basic straight",      [0,1,0,1,0,0,0,1,1,0,1,0,1,0,0,0,0,0,0,0]),
    ("Handlebars", "Comfort straight",    [0,0,0,0,1,0,0,0,0,0,0,0,0,0,1,0,0,1,1,0]),
    ("Handlebars", "Basic drop down",     [1,0,1,0,0,1,1,0,0,1,0,1,0,1,0,1,1,0,0,1]),
    ("Gears", "7 speed (1x7)",            [0,0,1,0,1,0,0,0,0,0,0,0,0,0,1,0,0,1,1,0]),
    ("Gears", "14 speed (2x7)",           [1,0,0,0,0,1,1,0,0,0,0,1,0,1,0,1,1,0,0,1]),
    ("Gears", "24 speed (3x8)",           [0,1,0,1,0,0,0,1,1,1,1,0,1,0,0,0,0,0,0,0]),
    ("Seat", "Polymer gel all-purpose",   [0,1,0,1,0,0,0,0,1,0,1,0,0,0,0,0,0,0,0,0]),
    ("Seat", "Polymer gel comfort",       [0,0,0,0,1,0,0,1,0,0,0,0,1,0,1,0,0,1,1,0]),
    ("Seat", "Polymer gel racing",        [1,0,1,0,0,1,1,0,0,1,0,1,0,1,0,1,1,0,0,1]),
    ("Accessories", "Reflectors",         [1,0,1,1,1,1,0,0,0,1,0,1,0,0,1,1,0,1,1,0]),
    ("Decals", "Colorful brushstrokes",   [1,1,1,1,1,1,0,1,1,0,1,0,1,1,1,0,1,0,0,1]),
    ("Lights", "Standard",                [1,0,1,0,1,1,1,0,0,1,0,0,0,1,0,1,1,1,1,1]),
    ("Carriers", "Plastic basket",        [0,0,0,0,1,0,0,0,0,0,0,0,0,0,1,0,0,1,1,0]),
    ("Suspension", "Front shocks",        [0,1,0,1,1,0,0,1,1,0,1,0,1,0,1,0,0,1,1,0]),
]

# ---------------- Q3_Components ----------------
ws = wb.create_sheet("Q3_Components")
ws["A1"] = "Competitors' Brands — component teardown, Q3 actual (all 20 brands)"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = "1 = included, blank = not included. 'Base components' and 'Standard carbon fiber' are in every brand and omitted."
ws["A2"].alignment = Alignment(wrap_text=True)

# header block
for j, (name, co, seg, score) in enumerate(BRANDS):
    c = 3 + j
    ws.cell(4, c, co).font = bold
    ws.cell(5, c, name).font = bold
    ws.cell(6, c, seg)
    sc = ws.cell(7, c, score)
    sc.font = bold
    ws.cell(4, c).fill = hdr
    ws.cell(5, c).fill = ours if co == "WeBike" else hdr
    if co == "WeBike":
        ws.cell(6, c).fill = ours
        ws.cell(7, c).fill = ours
    ws.cell(5, c).alignment = Alignment(textRotation=90, vertical="bottom")
ws.cell(4, 1, "Company").font = bold
ws.cell(5, 1, "Brand").font = bold
ws.cell(6, 1, "Targets").font = bold
ws.cell(7, 1, "Judgment in target segment").font = bold

r = 8
last_group = None
for group, label, vals in COMPONENTS:
    if group != last_group:
        ws.cell(r, 1, group).font = bold
        ws.cell(r, 1).fill = hdr
        r += 1
        last_group = group
    ws.cell(r, 2, label)
    for j, v in enumerate(vals):
        if v:
            cell = ws.cell(r, 3 + j, 1)
            cell.alignment = Alignment(horizontal="center")
            cell.fill = ours if BRANDS[j][1] == "WeBike" else good
    r += 1

ws.column_dimensions["A"].width = 14
ws.column_dimensions["B"].width = 26
for j in range(len(BRANDS)):
    ws.column_dimensions[ws.cell(1, 3 + j).column_letter].width = 4.5
ws.row_dimensions[5].height = 110
ws.freeze_panes = "C8"

# ---------------- Q3_Design_Model ----------------
ws = wb.create_sheet("Q3_Design_Model")
ws["A1"] = "Reverse-engineered brand judgment model (fitted to all 20 brands)"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("Every brand in each segment shares one base recipe; scores differ only by a few components. "
            "The point values below reproduce all 20 observed scores exactly.")
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
r = block(r, [
    ["SPEED — base recipe (all 10 Speed brands share this)", "Component"],
    ["Frame", "Aerodynamic"],
    ["Tires", "Racing - fast"],
    ["Brakes", "Precision"],
    ["Handlebars", "Basic drop down"],
    ["Seat", "Polymer gel racing"],
    ["Carriers / Suspension", "none"],
])

r = block(r, [
    ["SPEED — what the variable components are worth", "Points", "Evidence"],
    ["Gears 14 speed (2x7)", "optimal", "all 8 top Speed brands use it"],
    ["Gears 24 speed (3x8) instead", -3, "Swift Bike is the ONLY Speed brand with it"],
    ["Gears 7 speed (1x7) instead", -16, "Mach 0.6 = 61 with an otherwise full kit"],
    ["Decals (colorful brushstrokes)", 2, "every 76-77 brand has them; every 72-75 brand lacks them"],
    ["Reflectors", 1, "MACH I.I 77 (with) vs Blu Tube 76 (without), else identical"],
    ["Lights", 1, "Blu Aero 74 lacks lights; otherwise matches 75-point kit"],
    ["MAXIMUM ACHIEVABLE (Speed)", 77, "MACH I.I and LiteSpeed Pro+ both hit it"],
], wrap_col=3)
ws.cell(r - 3, 2).fill = good

r = block(r, [
    ["SWIFT BIKE — exact fix to reach the industry ceiling", "Now", "Change to", "Points"],
    ["Frame", "Aerodynamic", "keep", 0],
    ["Tires", "Racing", "keep", 0],
    ["Brakes", "Precision", "keep", 0],
    ["Handlebars", "Basic drop down", "keep", 0],
    ["Seat", "Polymer gel racing", "keep", 0],
    ["Reflectors", "included", "keep", 0],
    ["Lights", "included", "keep", 0],
    ["Gears", "24 speed (3x8)", "14 speed (2x7)", 3],
    ["Decals", "not included", "ADD colorful brushstrokes", 2],
    ["Current judgment", 72, None, None],
    ["Projected judgment after both changes", "=B{0}+SUM(D{1}:D{2})".format(r + 10, r + 9, r + 10), "= industry ceiling (tied best)", None],
])
for rr in (r + 8, r + 9):
    ws.cell(rr, 3).fill = yellow
    ws.cell(rr, 4).fill = yellow
ws.cell(r + 11, 2).fill = good
ws.cell(r + 11, 2).font = bold

r = block(r, [
    ["Why the gears mistake happened", "Note"],
    ["Root cause", "Swift Bike inherited the 24-speed (3x8) mountain drivetrain from Hike Bike. "
     "24-speed is OPTIMAL in Mountain (all 6 Mountain brands use it) and WRONG in Speed "
     "(all 8 leading Speed brands use 14-speed). We copied our Mountain spec into a Speed bike."],
    ["Cost side effect", "14 speed (2x7) is a simpler drivetrain than 24 speed (3x8), so this change should "
     "also LOWER unit cost. Verify against the component price list before locking."],
], wrap_col=2)

r = block(r, [
    ["MOUNTAIN — base recipe (all 6 Mountain brands share this)", "Component"],
    ["Frame", "Rugged"],
    ["Tires", "Mountain - high grip"],
    ["Brakes", "Standard disc"],
    ["Handlebars", "Basic straight"],
    ["Gears", "24 speed (3x8)"],
    ["Suspension", "Front shocks"],
    ["Decals", "Colorful brushstrokes"],
    ["Lights / Carriers", "none"],
])

r = block(r, [
    ["MOUNTAIN — what the variable components are worth", "Points", "Evidence"],
    ["Seat: Polymer gel ALL-PURPOSE", "optimal", "Hike Bike, TERRAMAX, Blu Ruged Ballz all = 73"],
    ["Seat: Polymer gel COMFORT instead", -3, "LiteTrail Pro and Blu Tail Ballz both = 70"],
    ["Reflectors added", -1, "TERRAMean = 72; identical to Hike Bike except reflectors"],
    ["MAXIMUM ACHIEVABLE (Mountain)", 73, "Hike Bike already there"],
], wrap_col=3)
ws.cell(r - 2, 2).fill = bad
ws.cell(r - 3, 2).fill = good

r = block(r, [
    ["HIKE BIKE — verified optimal, do not touch", "Status"],
    ["Seat", "Polymer gel all-purpose = correct (comfort seat would cost 3 points)"],
    ["Reflectors", "correctly EXCLUDED (adding them costs 1 point - they are a Recreation cue)"],
    ["Lights", "correctly excluded"],
    ["Judgment", "73 of a possible 73 - nothing left to gain from redesign"],
    ["Implication", "Any Q4 budget aimed at improving Hike Bike is wasted. Spend it on capacity, "
     "sales coverage and ads instead."],
], wrap_col=2)

r = block(r, [
    ["RECREATION — blueprint if we ever enter (Q5 option)", "Component"],
    ["Frame", "Comfort (relaxed)"],
    ["Tires", "Hybrid - road/off-road"],
    ["Brakes", "Standard disc  <-- NOT precision (costs 2) and NOT standard (costs 1)"],
    ["Handlebars", "Comfort straight"],
    ["Gears", "7 speed (1x7)"],
    ["Seat", "Polymer gel comfort"],
    ["Accessories", "Reflectors + Decals + Lights + Plastic basket + Front shocks"],
    ["Result", "76 = MountainCruise1 (Bike Bros), the only brand at the Recreation ceiling"],
    ["CAUTION", "Only 4 Recreation brands exist and they differ in 3 places (brakes, decals, lights), "
     "so the individual penalties are NOT uniquely identifiable. Observed constraints: "
     "precision+no-decals = 2 - precision+no-lights = 3 - standard-brakes+no-decals = 3. "
     "What IS certain: standard disc brakes are optimal, and MountainCruise1's exact recipe scores 76."],
], wrap_col=2)

r = block(r, [
    ["CROSS-SEGMENT RULE OF THUMB", "Note"],
    ["Reflectors", "+1 in Speed, -1 in Mountain, required in Recreation. Segment-dependent, not universally good."],
    ["Decals", "+2 in Speed, required in Mountain and Recreation. Cheap and always worth having except where noted."],
    ["Lights", "+1 in Speed, excluded in Mountain, required in Recreation."],
    ["Gears", "The single biggest lever in Speed: 7sp = -16, 24sp = -3, 14sp = optimal."],
    ["Seat", "The single biggest lever in Mountain: comfort seat = -3 vs all-purpose."],
], wrap_col=2)

ws.column_dimensions["A"].width = 44
ws.column_dimensions["B"].width = 30
ws.column_dimensions["C"].width = 34
ws.column_dimensions["D"].width = 10

wb.save(DST)
print("Added Q3_Components and Q3_Design_Model")
