"""Add Q3 Competitors' Profiles by city — measured store vs web split + people."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

DST = "quarters/Q3/Q3Data.xlsx"

FIRMS = ["Bike Bros", "LiteCycle", "WeBike", "BB LLC", "Spoke'd Up", "SpaceBikes", "MILC Bikes"]
CITIES = ["Amsterdam", "New York City", "Rio de Janeiro", "Bangalore"]

# (company, city, total, store, web, brands, avg_price, inserts,
#  store_sf, dem_per_store, web_sf, dem_per_web)
# None = N/A (no presence / not applicable)
P = [
    # New York City
    ("LiteCycle",  "New York City", 143,   0, 143, 3, 1380, 12, None, None, 7, 20),
    ("WeBike",     "New York City",  38,   0,  38, 2, 1408, 19, None, None, 3, 13),
    ("BB LLC",     "New York City", 752, 567, 185, 4, 1473, 24,    7,   81, 7, 26),
    ("Spoke'd Up", "New York City", 412, 371,  41, 2, 1249, 12,    5,   74, 3, 14),
    ("SpaceBikes", "New York City", 574, 435, 139, 2, 1340, 17,    7,   62, 7, 20),
    ("MILC Bikes", "New York City", 120,   0, 120, 2, 1250, 12, None, None, 7, 17),
    # Rio de Janeiro
    ("LiteCycle",  "Rio de Janeiro", 106,   0, 106, 3, 1380, 12, None, None, 7, 15),
    ("WeBike",     "Rio de Janeiro", 270, 240,  30, 2, 1408, 19,    4,   60, 3, 10),
    ("BB LLC",     "Rio de Janeiro", 139,   0, 139, 4, 1473, 24, None, None, 7, 20),
    ("Spoke'd Up", "Rio de Janeiro",  32,   0,  32, 2, 1249, 12, None, None, 3, 11),
    ("SpaceBikes", "Rio de Janeiro", 120,   0, 120, 2, 1340, 17, None, None, 7, 17),
    ("MILC Bikes", "Rio de Janeiro", 107,   0, 107, 2, 1250, 12, None, None, 7, 15),
    # Amsterdam
    ("Bike Bros",  "Amsterdam", 490, 490,   0, 5, 1475, 10,    7,   70, None, None),
    ("LiteCycle",  "Amsterdam", 365, 272,  93, 3, 1380, 12,    5,   54,    7,   13),
    ("WeBike",     "Amsterdam", 295, 267,  28, 2, 1408, 19,    7,   38,    3,    9),
    ("BB LLC",     "Amsterdam", 493, 372, 121, 4, 1473, 24,    7,   53,    7,   17),
    ("Spoke'd Up", "Amsterdam", 380, 344,  36, 2, 1249, 12,    6,   57,    3,   12),
    ("SpaceBikes", "Amsterdam", 490, 371, 119, 2, 1340, 17,    7,   53,    7,   17),
    ("MILC Bikes", "Amsterdam", 472, 365, 107, 2, 1250, 12,    7,   52,    7,   15),
    # Bangalore
    ("LiteCycle",  "Bangalore", 370, 277,  93, 3, 1380, 12,    5,   55,    7,   13),
    ("WeBike",     "Bangalore",  24,   0,  24, 2, 1408, 19, None, None,    3,    8),
    ("BB LLC",     "Bangalore", 123,   0, 123, 4, 1473, 24, None, None,    7,   18),
    ("Spoke'd Up", "Bangalore",  30,   0,  30, 2, 1249, 12, None, None,    3,   10),
    ("SpaceBikes", "Bangalore", 114,   0, 114, 2, 1340, 17, None, None,    7,   16),
    ("MILC Bikes", "Bangalore", 100,   0, 100, 2, 1250, 12, None, None,    7,   14),
]

# company-level brand / ad scores (identical in every city report)
BRANDS = {
    "Bike Bros":  {"rec": (76, "MountainCruise1"), "mtn": (73, "TERRAMAX"),
                   "spd": (77, "MACH I.I"),
                   "ad_rec": (75, "Easy Rider"), "ad_mtn": (78, "TerraTech"),
                   "ad_spd": (78, "AndStill")},
    "LiteCycle":  {"rec": None, "mtn": (70, "LiteTrail Pro"),
                   "spd": (77, "LiteSpeed Pro+"),
                   "ad_rec": None, "ad_mtn": (79, "Trail Blazing 1"),
                   "ad_spd": (77, "Speed of Lite 1")},
    "WeBike":     {"rec": None, "mtn": (73, "Hike Bike"),
                   "spd": (72, "Swift Bike"),
                   "ad_rec": None, "ad_mtn": (80, "HikeBike 1"),
                   "ad_spd": (70, "Swift Bike")},
    "BB LLC":     {"rec": None, "mtn": (73, "Blu Ruged Ballz"),
                   "spd": (76, "Blu Tube Ballz"),
                   "ad_rec": None, "ad_mtn": (81, "BB Big Momma"),
                   "ad_spd": (76, "Unleash lil pap")},
    "Spoke'd Up": {"rec": (73, "Spoke'd Easy"), "mtn": None,
                   "spd": (75, "Spoke'd Speed"),
                   "ad_rec": None, "ad_mtn": None, "ad_spd": None},
    "SpaceBikes": {"rec": (73, "Mars Rover"), "mtn": None,
                   "spd": (76, "The Armstrong"),
                   "ad_rec": (79, "Mars Rover 1"), "ad_mtn": None,
                   "ad_spd": (77, "The Armstrong 1")},
    "MILC Bikes": {"rec": (74, "Whole MILC MKII"), "mtn": None,
                   "spd": (76, "Skim MILC MKII"),
                   "ad_rec": (77, "Whole MILC MKII"), "ad_mtn": None,
                   "ad_spd": (75, "Skim MILC MKII")},
}


def rows_for(city):
    return [r for r in P if r[1] == city]


def firm_row(firm, city):
    for r in P:
        if r[0] == firm and r[1] == city:
            return r
    return None


print("RECONCILIATION")
grand_t = grand_s = grand_w = 0
for city in CITIES:
    rs = rows_for(city)
    t = sum(r[2] for r in rs)
    s = sum(r[3] for r in rs)
    w = sum(r[4] for r in rs)
    grand_t += t
    grand_s += s
    grand_w += w
    print(f"  {city:<16} total {t:>5}  store {s:>5}  web {w:>5}  s+w {s+w:>5}  {'OK' if t == s + w else 'SPLIT MISMATCH'}")
print(f"  {'GRAND':<16} total {grand_t:>5}  store {grand_s:>5}  web {grand_w:>5}")
assert grand_t == 6559, grand_t
assert grand_t == grand_s + grand_w

# WeBike identity
we_t = sum(r[2] for r in P if r[0] == "WeBike")
we_s = sum(r[3] for r in P if r[0] == "WeBike")
we_w = sum(r[4] for r in P if r[0] == "WeBike")
print(f"  WeBike          total {we_t:>5}  store {we_s:>5}  web {we_w:>5}")
assert (we_t, we_s, we_w) == (627, 507, 120)

# per-person identity
print("\nPER-PERSON CHECK")
for r in P:
    co, city, tot, store, web, *rest, ssf, dsp, wsf, dpw = r
    if ssf is not None:
        calc = round(store / ssf)
        flag = "OK" if calc == dsp else f"MISMATCH calc={calc}"
        print(f"  store {city:<16}{co:<12} {store}/{ssf} -> {calc} reported {dsp}  {flag}")
        assert calc == dsp
    if wsf is not None:
        calc = round(web / wsf)
        flag = "OK" if calc == dpw else f"MISMATCH calc={calc}"
        print(f"  web   {city:<16}{co:<12} {web}/{wsf} -> {calc} reported {dpw}  {flag}")
        assert calc == dpw

# channel-report rounding (±1) note
print("\nFIRM TOTALS vs Q3_Channel (expect ±1 rounding)")
CHANNEL = {
    "BB LLC": (938, 569, 1507),
    "SpaceBikes": (807, 491, 1298),
    "LiteCycle": (549, 435, 984),
    "Spoke'd Up": (716, 138, 854),
    "MILC Bikes": (364, 435, 799),
    "WeBike": (507, 120, 627),
    "Bike Bros": (490, 0, 490),
}
for f in FIRMS:
    s = sum(r[3] for r in P if r[0] == f)
    w = sum(r[4] for r in P if r[0] == f)
    t = sum(r[2] for r in P if r[0] == f)
    cs, cw, ct = CHANNEL[f]
    print(f"  {f:<12} city {s:>4}/{w:>4}/{t:>5}  channel {cs:>4}/{cw:>4}/{ct:>5}  dS {s-cs:+d} dW {w-cw:+d}")
    assert t == ct
    assert abs(s - cs) <= 1 and abs(w - cw) <= 1

wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
bad = PatternFill("solid", fgColor="FFC7CE")
good = PatternFill("solid", fgColor="C6EFCE")
yellow = PatternFill("solid", fgColor="FFEB9C")
ours = PatternFill("solid", fgColor="FFF2CC")

if "Q3_City_Profiles" in wb.sheetnames:
    del wb["Q3_City_Profiles"]
ws = wb.create_sheet("Q3_City_Profiles")

ws["A1"] = "Competitors' Profiles by city — Q3 actual (measured store vs web)"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("Source: Detailed report for city, all four cities. Totals reconcile to 6,559. "
            "WeBike 627 / 507 store / 120 web. Firm store+web vs Q3_Channel is ±1 rounding; "
            "city figures are the ground truth. Bike Bros is N/A outside Amsterdam.")
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
rows = [["Company", "City", "Total", "Store", "Web", "Brands", "Avg price", "Inserts",
         "Store people", "Demand / store person", "Web people", "Demand / web person"]]
for row in P:
    rows.append(list(row))
r2 = block(r, rows)
for i, row in enumerate(P, start=1):
    if row[0] == "WeBike":
        for c in range(1, 13):
            ws.cell(r + i, c).fill = ours
            ws.cell(r + i, c).font = bold
r = r2

# city store vs web mix
rows = [["CITY CHANNEL MIX (measured)", "Total", "Store units", "Store %", "Web units", "Web %",
         "Store firms", "Store people in city"]]
city_mix = {}
for city in CITIES:
    rs = rows_for(city)
    t = sum(x[2] for x in rs)
    s = sum(x[3] for x in rs)
    w = sum(x[4] for x in rs)
    n_store_firms = sum(1 for x in rs if x[8] is not None)
    n_people = sum(x[8] or 0 for x in rs)
    city_mix[city] = (t, s, w, n_store_firms, n_people)
    rows.append([city, t, s, s / t, w, w / t, n_store_firms, n_people])
r2 = block(r, rows)
for i in range(1, 5):
    ws.cell(r + i, 4).number_format = "0.0%"
    ws.cell(r + i, 6).number_format = "0.0%"
ws.cell(r + 3, 4).fill = bad   # Rio store % low
ws.cell(r + 2, 4).fill = yellow  # NYC 67%
r = r2

# store productivity by city-firm
rows = [["STORE PRODUCTIVITY (pure store units / store person)", "City", "Store units",
         "People", "Per person", "Note"]]
store_rows = sorted(
    [x for x in P if x[8] is not None],
    key=lambda x: -(x[3] / x[8]),
)
for co, city, tot, store, web, *_, ssf, dsp, wsf, dpw in store_rows:
    note = ""
    if co == "WeBike" and city == "Amsterdam":
        note = "LAST in our home city"
    elif co == "WeBike" and city == "Rio de Janeiro":
        note = "Already 1.6x our Amsterdam rate; 3 slots open"
    elif city == "New York City" and co == "Spoke'd Up":
        note = "5 people, NO Mountain brand — NYC floor analog"
    elif city == "New York City" and co == "BB LLC":
        note = "NYC ceiling analog (Mountain + Speed, 7 people)"
    rows.append([co, city, store, ssf, dsp, note])
r2 = block(r, rows, wrap_col=6)
for i, row in enumerate(store_rows, start=1):
    if row[0] == "WeBike":
        for c in range(1, 7):
            ws.cell(r + i, c).fill = ours
        if row[1] == "Amsterdam":
            ws.cell(r + i, 5).fill = bad
        if row[1] == "Rio de Janeiro":
            ws.cell(r + i, 5).fill = good
    if row[1] == "New York City":
        ws.cell(r + i, 5).fill = good
r = r2

# NYC store prize
rows = [
    ["NYC STORE PRIZE — retire the 10% = 204 estimate", "People", "Store units",
     "Web units", "Total", "Per store person", "Read"],
    ["BB LLC (has store)", 7, 567, 185, 752, 81, "Mountain + Speed, 4 brands, 24 inserts"],
    ["SpaceBikes (has store)", 7, 435, 139, 574, 62, "Recreation + Speed, worst NYC store rate"],
    ["Spoke'd Up (has store)", 5, 371, 41, 412, 74, "Rec + Speed, NO Mountain, only 5 people"],
    ["LiteCycle (web only)", None, 0, 143, 143, None, "Full web op, no store — web-only ceiling"],
    ["MILC Bikes (web only)", None, 0, 120, 120, None, "Web-first firm"],
    ["WeBike (web only)", None, 0, 38, 38, None, "Last among web-only firms"],
    ["NYC physical store total", 19, 1373, 365, 1738, round(1373 / 19),
     "67.3% of city is store; 85.2% is store-firm (store+their web)"],
    ["If we open as 4th store, floor", 7, 371, None, None, 53,
     "Match Spoke'd Up's STORE units (they have 5 people and no Mountain)."],
    ["If we staff 7 at SpaceBikes rate", 7, 434, None, None, 62,
     "Worst incumbent NYC rate × our likely 7-cap staff."],
    ["If 4th store splits today's 1,373", 7, 343, None, None, 49,
     "Conservative: store pool unchanged, split 4 ways instead of 3."],
    ["VERDICT", None, None, None, None, None,
     "The 10% = 204 figure is retired. A NYC store is a 340–570 unit prize, "
     "not a 204 unit prize. Enter via Speed (Spoke'd Up's Speed ad is <70). "
     "Stacked on Rio-fill + web rebuild this needs printers — see Q4_Planner scenario E."],
]
r2 = block(r, rows, wrap_col=7)
ws.cell(r + 3, 3).fill = yellow  # Spoke'd Up 371 floor
ws.cell(r + 8, 3).fill = good
ws.cell(r + 11, 7).fill = yellow
ws.cell(r + 11, 7).alignment = Alignment(wrap_text=True)
r = r2

# Amsterdam last
rows = [
    ["AMSTERDAM — we are last in our own city", "Store units", "People", "Per person",
     "Web units", "Total"],
    ["Bike Bros", 490, 7, 70, 0, 490],
    ["Spoke'd Up", 344, 6, 57, 36, 380],
    ["LiteCycle", 272, 5, 54, 93, 365],
    ["BB LLC", 372, 7, 53, 121, 493],
    ["SpaceBikes", 371, 7, 53, 119, 490],
    ["MILC Bikes", 365, 7, 52, 107, 472],
    ["WeBike", 267, 7, 38, 28, 295],
    ["Industry excl. us", 2214, 39, round(2214 / 39), 476, 2690],
    ["Read", None, None, None, None,
     "Same 7 people, same city: Bike Bros 70 vs WeBike 38. We cannot add a head "
     "(at cap). Closing the gap is training, filling stock-outs, and the Swift Bike "
     "fix — not hiring. Matching industry 57/head would be ~399 store units (+132) "
     "with the people we already have."],
]
r2 = block(r, rows, wrap_col=6)
ws.cell(r + 7, 1).fill = ours
ws.cell(r + 7, 4).fill = bad
ws.cell(r + 1, 4).fill = good
r = r2

# Rio vs Bangalore sole-store, now with store/web split
rows = [
    ["SOLE-STORE CITIES — store vs web split", "WeBike in Rio", "LiteCycle in Bangalore"],
    ["City demand", 774, 761],
    ["Physical store units", 240, 277],
    ["Store people", 4, 5],
    ["Store units per person", 60, 55],
    ["Own web in the city", 30, 93],
    ["Total as sole store holder", 270, 370],
    ["Share of city", 0.349, 0.486],
    ["Rivals' web (no store)", 504, 391],
    ["Open store slots", 3, 2],
    ["Fill to 7 at current store rate", 420, 385],
    ["Read", "Filling 3 Rio slots at 60/head = +180 store units, no new lease. "
     "Our 30 web in our own city is 1/3 of LiteCycle's 93 in theirs — another "
     "web-rebuild tell. BB LLC takes 139 Rio units with no store.",
     "LiteCycle extracts more with 5 people than we do with 4, and still has "
     "2 slots open. Their city web (93) is 3x ours in Rio (30)."],
]
r2 = block(r, rows, wrap_col=None)
ws.cell(r + 7, 2).number_format = "0.0%"
ws.cell(r + 7, 3).number_format = "0.0%"
ws.cell(r + 4, 2).fill = good
ws.cell(r + 10, 2).fill = good
ws.cell(r + 11, 2).alignment = Alignment(wrap_text=True)
ws.cell(r + 11, 3).alignment = Alignment(wrap_text=True)
r = r2

# web raid by city
rows = [["WEB RAID BY CITY (web units, one shared web centre)",
         "Amsterdam", "New York City", "Rio de Janeiro", "Bangalore", "Total", "Web people", "Per head"]]
web_firms = ["BB LLC", "SpaceBikes", "LiteCycle", "MILC Bikes", "Spoke'd Up", "WeBike"]
for f in web_firms:
    vals = []
    for city in CITIES:
        row = firm_row(f, city)
        vals.append(row[4] if row else 0)
    total = sum(vals)
    wsf = next(x[10] for x in P if x[0] == f and x[10] is not None)
    rows.append([f] + vals + [total, wsf, round(total / wsf, 1)])
r2 = block(r, rows)
for i, f in enumerate(web_firms, start=1):
    if f == "WeBike":
        for c in range(1, 9):
            ws.cell(r + i, c).fill = ours
            ws.cell(r + i, c).font = bold
r = r2

rows = [
    ["WHAT THE WEB RAID SHOWS", "Note"],
    ["BB LLC is even across all 4 cities",
     "Amsterdam 121 · NYC 185 · Rio 139 · Bangalore 123. They raid cities they "
     "never entered at the same scale they sell online in cities they own."],
    ["We are even and tiny",
     "Amsterdam 28 · NYC 38 · Rio 30 · Bangalore 24. Same shape, ~1/5 the scale."],
    ["Even in Amsterdam, where we have a store",
     "BB LLC 121 web vs our 28. Store presence does not automatically produce "
     "web demand — the web operation does."],
    ["LiteCycle with no NYC store beats our NYC web 3.8x",
     "143 vs 38. That is the web-only NYC benchmark once we staff 7 and fund "
     "all 4 tactics."],
    ["Spoke'd Up is our twin",
     "Web staff 3, web total 139 vs our 120. Same failure mode. They still "
     "out-produce us in NYC store (371 with 5 people) because they showed up."],
]
r2 = block(r, rows, wrap_col=2)
ws.cell(r + 3, 2).fill = bad
ws.cell(r + 5, 2).fill = yellow
r = r2

# company brands (for completeness; already in other sheets)
rows = [["COMPANY BRAND / AD SCORES (repeated identically in every city report)",
         "Rec brand", "Mtn brand", "Speed brand", "Rec ad", "Mtn ad", "Speed ad", "Inserts"]]
inserts = {f: next(x[7] for x in P if x[0] == f) for f in FIRMS}
for f in FIRMS:
    b = BRANDS[f]

    def fmt(v):
        return f"{v[0]} {v[1]}" if v else "—"
    rows.append([f, fmt(b["rec"]), fmt(b["mtn"]), fmt(b["spd"]),
                 fmt(b["ad_rec"]), fmt(b["ad_mtn"]), fmt(b["ad_spd"]), inserts[f]])
r2 = block(r, rows)
for i, f in enumerate(FIRMS, start=1):
    if f == "WeBike":
        for c in range(1, 9):
            ws.cell(r + i, c).fill = ours
        ws.cell(r + i, 7).fill = bad  # Speed ad 70
r = r2

rows = [
    ["Q4 IMPLICATIONS (this report)", "Decision"],
    ["1. Train 2 untrained — $800",
     "Unchanged. Cheapest productivity in both cities we already occupy."],
    ["2. Fill Rio 4 → 7",
     "Revised: +180 store units at the measured 60/head (was ~+200 using "
     "total city demand / heads). Still the highest-return staffing move. "
     "Fits 24/day easily."],
    ["3. Full web rebuild 3 → 7 + 4 tactics",
     "Confirmed from a second direction: we are last in web units in every "
     "city, including the two we have stores in. Amsterdam 28 vs BB LLC 121."],
    ["4. NYC store sizing",
     "Floor ~340–370 store units (4-way split or Spoke'd Up match). "
     "Staff-7 at worst NYC rate ~434. BB LLC analog ~567. Retire 10% = 204. "
     "Enter via Speed; Spoke'd Up's Speed ad scores <70."],
    ["5. Printers if NYC is in Q4",
     "Rio-fill + web rebuild already 24.3/day vs 24 owned. NYC at 340–570 "
     "on top needs ~2 more printers. Idle cash $1,010,838 funds it."],
    ["NOT a priority",
     "Adding Amsterdam heads (illegal — at cap). Bangalore vs LiteCycle's "
     "48.6%. A 3rd brand."],
]
block(r, rows, wrap_col=2)
for i in range(1, 6):
    ws.cell(r + i, 2).fill = yellow

ws.column_dimensions["A"].width = 48
ws.column_dimensions["B"].width = 22
ws.column_dimensions["C"].width = 18
ws.column_dimensions["D"].width = 16
ws.column_dimensions["E"].width = 16
ws.column_dimensions["F"].width = 18
ws.column_dimensions["G"].width = 22
for col in "HIJKLM":
    ws.column_dimensions[col].width = 18
ws.row_dimensions[2].height = 48

wb.save(DST)
print("\nAdded Q3_City_Profiles")
