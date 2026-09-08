"""Add Q3 World Market major-media inserts (all 106 industry inserts) to Q3Data.xlsx."""
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

DST = "quarters/Q3/Q3Data.xlsx"

COST = {
    "Leisure & Entertainment": 10000,
    "Health & Fitness Magazines": 7000,
    "Biking Magazines": 4500,
    "Sport Magazines": 10000,
    "Business Magazines": 9500,
    "New Venture Magazines": 5500,
    "General News Magazines": 8000,
}
MEDIA = list(COST)

# media, company, ad, brand, inserts — exact from Q3 World Market report
RAW = [
    ("Leisure & Entertainment", "Bike Bros", "Easy Rider", "MountainCruise1", 1),
    ("Leisure & Entertainment", "Spoke'd Up", "Spoke'd Easy", "Spoke'd Easy", 3),
    ("Leisure & Entertainment", "SpaceBikes", "Mars Rover 1", "Mars Rover", 2),
    ("Health & Fitness Magazines", "Bike Bros", "TerraTech", "TERRAMAX", 1),
    ("Health & Fitness Magazines", "Bike Bros", "Easy Rider", "MountainCruise1", 1),
    ("Health & Fitness Magazines", "WeBike", "HikeBike 1", "Hike Bike", 1),
    ("Health & Fitness Magazines", "Spoke'd Up", "Spoke'd Easy", "Spoke'd Easy", 2),
    ("Health & Fitness Magazines", "SpaceBikes", "Mars Rover 1", "Mars Rover", 3),
    ("Health & Fitness Magazines", "MILC Bikes", "Whole MILC MKII", "Whole MILC MKII", 5),
    ("Biking Magazines", "Bike Bros", "Excluspeedity", "MACH I.I", 1),
    ("Biking Magazines", "Bike Bros", "TerraTech", "TERRAMAX", 1),
    ("Biking Magazines", "LiteCycle", "Speed of Lite 1", "LiteSpeed Pro+", 2),
    ("Biking Magazines", "LiteCycle", "Speed of Lite 2", "LiteSpeed+", 2),
    ("Biking Magazines", "LiteCycle", "Trail Blazing 1", "LiteTrail Pro", 2),
    ("Biking Magazines", "WeBike", "HikeBike 1", "Hike Bike", 6),
    ("Biking Magazines", "WeBike", "Swift Bike", "Swift Bike", 4),
    ("Biking Magazines", "BB LLC", "BB Big Momma", "Blu Ruged Ballz", 12),
    ("Biking Magazines", "BB LLC", "Unleash lil pap", "Blu Tube Ballz", 12),
    ("Biking Magazines", "Spoke'd Up", "Spoke'd Easy", "Spoke'd Easy", 1),
    ("Biking Magazines", "Spoke'd Up", "Spoke'd Speed", "Spoke'd Speed", 1),
    ("Biking Magazines", "SpaceBikes", "The Armstrong 1", "The Armstrong", 2),
    ("Biking Magazines", "MILC Bikes", "Skim MILC MKII", "Skim MILC MKII", 5),
    ("Sport Magazines", "Bike Bros", "Mountainability", "TERRAMAX", 1),
    ("Sport Magazines", "Bike Bros", "AndStill", "MACH I.I", 1),
    ("Sport Magazines", "LiteCycle", "Trail Blazing 1", "LiteTrail Pro", 1),
    ("Sport Magazines", "WeBike", "HikeBike 1", "Hike Bike", 2),
    ("Sport Magazines", "WeBike", "Swift Bike", "Swift Bike", 1),
    ("Sport Magazines", "Spoke'd Up", "Spoke'd Speed", "Spoke'd Speed", 2),
    ("Sport Magazines", "SpaceBikes", "The Armstrong 1", "The Armstrong", 2),
    ("Business Magazines", "Bike Bros", "Excluspeedity", "MACH I.I", 1),
    ("Business Magazines", "LiteCycle", "Speed of Lite 1", "LiteSpeed Pro+", 1),
    ("Business Magazines", "LiteCycle", "Speed of Lite 2", "LiteSpeed+", 1),
    ("Business Magazines", "WeBike", "Swift Bike", "Swift Bike", 1),
    ("Business Magazines", "Spoke'd Up", "Spoke'd Speed", "Spoke'd Speed", 1),
    ("Business Magazines", "SpaceBikes", "The Armstrong 1", "The Armstrong", 2),
    ("New Venture Magazines", "Bike Bros", "AndStill", "MACH I.I", 1),
    ("New Venture Magazines", "LiteCycle", "Speed of Lite 1", "LiteSpeed Pro+", 1),
    ("New Venture Magazines", "LiteCycle", "Speed of Lite 2", "LiteSpeed+", 1),
    ("New Venture Magazines", "WeBike", "HikeBike 1", "Hike Bike", 1),
    ("New Venture Magazines", "WeBike", "Swift Bike", "Swift Bike", 2),
    ("New Venture Magazines", "Spoke'd Up", "Spoke'd Speed", "Spoke'd Speed", 1),
    ("New Venture Magazines", "SpaceBikes", "The Armstrong 1", "The Armstrong", 2),
    ("New Venture Magazines", "MILC Bikes", "Skim MILC MKII", "Skim MILC MKII", 1),
    ("New Venture Magazines", "MILC Bikes", "Whole MILC MKII", "Whole MILC MKII", 1),
    ("General News Magazines", "Bike Bros", "Mountainability", "TERRAMAX", 1),
    ("General News Magazines", "LiteCycle", "Speed of Lite 1", "LiteSpeed Pro+", 1),
    ("General News Magazines", "WeBike", "HikeBike 1", "Hike Bike", 1),
    ("General News Magazines", "Spoke'd Up", "Spoke'd Easy", "Spoke'd Easy", 1),
    ("General News Magazines", "SpaceBikes", "The Armstrong 1", "The Armstrong", 2),
    ("General News Magazines", "SpaceBikes", "Mars Rover 1", "Mars Rover", 2),
]

# company -> reported insert total (Major Media Inserts, Q3)
REPORTED = {
    "Bike Bros": 10,
    "LiteCycle": 12,
    "WeBike": 19,
    "BB LLC": 24,
    "Spoke'd Up": 12,
    "SpaceBikes": 17,
    "MILC Bikes": 12,
}

# brand -> (company, target segment)
BRAND_SEG = {
    "MountainCruise1": ("Bike Bros", "Recreation"),
    "TERRAMAX": ("Bike Bros", "Mountain"),
    "MACH I.I": ("Bike Bros", "Speed"),
    "LiteSpeed Pro+": ("LiteCycle", "Speed"),
    "LiteSpeed+": ("LiteCycle", "Speed"),
    "LiteTrail Pro": ("LiteCycle", "Mountain"),
    "Hike Bike": ("WeBike", "Mountain"),
    "Swift Bike": ("WeBike", "Speed"),
    "Blu Ruged Ballz": ("BB LLC", "Mountain"),
    "Blu Tube Ballz": ("BB LLC", "Speed"),
    "Spoke'd Easy": ("Spoke'd Up", "Recreation"),
    "Spoke'd Speed": ("Spoke'd Up", "Speed"),
    "The Armstrong": ("SpaceBikes", "Speed"),
    "Mars Rover": ("SpaceBikes", "Recreation"),
    "Skim MILC MKII": ("MILC Bikes", "Speed"),
    "Whole MILC MKII": ("MILC Bikes", "Recreation"),
}

# Q2 HikeBike1 lock (for the Q2 vs Q3 comparison)
Q2_HIKE = {
    "Leisure & Entertainment": 0,
    "Health & Fitness Magazines": 1,
    "Biking Magazines": 12,
    "Sport Magazines": 2,
    "Business Magazines": 0,
    "New Venture Magazines": 0,
    "General News Magazines": 1,
}

by_co = defaultdict(int)
by_co_media = defaultdict(lambda: defaultdict(int))
by_co_ad = defaultdict(lambda: defaultdict(int))
by_co_spend = defaultdict(int)
by_media = defaultdict(int)
by_brand_biking = defaultdict(int)
rows_out = []
for media, co, ad, brand, n in RAW:
    spend = n * COST[media]
    by_co[co] += n
    by_co_media[co][media] += n
    by_co_ad[co][ad] += n
    by_co_spend[co] += spend
    by_media[media] += n
    if media == "Biking Magazines":
        by_brand_biking[brand] += n
    seg = BRAND_SEG[brand][1]
    rows_out.append((media, co, ad, brand, seg, n, COST[media], spend))

print("RECONCILIATION — company totals vs World Market rows")
mismatch = False
for co, reported in REPORTED.items():
    got = by_co[co]
    ok = got == reported
    print(f"  {co:<12} report {reported:>3}  rows {got:>3}  {'OK' if ok else 'MISMATCH'}")
    if not ok:
        mismatch = True
assert not mismatch
grand = sum(by_co.values())
print(f"  {'TOTAL':<12}          {grand:>3}")
assert grand == 106
assert sum(REPORTED.values()) == 106
print("  World Market IS the entire Major Media Inserts report\n")

print("INSERT SPEND (at published $/insert)")
for co, spend in sorted(by_co_spend.items(), key=lambda x: -x[1]):
    n = by_co[co]
    print(f"  {co:<12} {n:>3} inserts  ${spend:>9,}  ${spend // n:,}/insert")
print(f"  {'WeBike residual vs $119,165 ad expense:':} ${119165 - by_co_spend['WeBike']:,} "
      f"(web ads + design)\n")

wbike_hike_q3 = {m: 0 for m in MEDIA}
wbike_swift_q3 = {m: 0 for m in MEDIA}
for media, co, ad, brand, n in RAW:
    if co != "WeBike":
        continue
    if ad == "HikeBike 1":
        wbike_hike_q3[media] += n
    else:
        wbike_swift_q3[media] += n

hike_q3_n = sum(wbike_hike_q3.values())
hike_q3_spend = sum(wbike_hike_q3[m] * COST[m] for m in MEDIA)
hike_q2_n = sum(Q2_HIKE.values())
hike_q2_spend = sum(Q2_HIKE[m] * COST[m] for m in MEDIA)
swift_n = sum(wbike_swift_q3.values())
swift_spend = sum(wbike_swift_q3[m] * COST[m] for m in MEDIA)

print("WEBIKE Q2 vs Q3 — HikeBike 1")
print(f"  Q2 {hike_q2_n} inserts ${hike_q2_spend:,}  (Biking {Q2_HIKE['Biking Magazines']})")
print(f"  Q3 {hike_q3_n} inserts ${hike_q3_spend:,}  (Biking {wbike_hike_q3['Biking Magazines']})")
print(f"  Swift Bike Q3 {swift_n} inserts ${swift_spend:,}  (Biking {wbike_swift_q3['Biking Magazines']})")
print(f"  Combined Q3 {hike_q3_n + swift_n} inserts ${hike_q3_spend + swift_spend:,}")
assert hike_q3_n + swift_n == 19
assert hike_q3_spend + swift_spend == by_co_spend["WeBike"]
assert by_co_spend["BB LLC"] == 24 * 4500
print("  BB LLC spend $108,000 = 24 x $4,500 (all Biking). Checks.\n")

biking_total = by_media["Biking Magazines"]
print(f"BIKING MAGAZINES: {biking_total} of {grand} industry inserts "
      f"({biking_total / grand:.1%})")
print(f"  BB LLC {by_co_media['BB LLC']['Biking Magazines']} = "
      f"{by_co_media['BB LLC']['Biking Magazines'] / biking_total:.1%} of all Biking")
print(f"  WeBike {by_co_media['WeBike']['Biking Magazines']}")
print(f"  Mountain Biking: Blu Ruged 12 vs Hike Bike 6")
print(f"  Speed Biking:    Blu Tube 12 vs Swift Bike 4")

# ---------------- write sheet ----------------
wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
bad = PatternFill("solid", fgColor="FFC7CE")
good = PatternFill("solid", fgColor="C6EFCE")
yellow = PatternFill("solid", fgColor="FFEB9C")
ours = PatternFill("solid", fgColor="FFF2CC")

if "Q3_Media" in wb.sheetnames:
    del wb["Q3_Media"]
ws = wb.create_sheet("Q3_Media")

ws["A1"] = "Major Media Inserts — World Market, Q3 actual (all 106 industry inserts)"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = (
    "World Market rows sum exactly to the company insert totals, so this IS the whole "
    "Major Media Inserts report (not a subset). BB LLC put all 24 inserts in Biking Magazines "
    "($4,500, highest Mountain AND Speed preference) and spent LESS than we did. We cut "
    "HikeBike 1 Biking from 12 (Q2) to 6 (Q3) to 'diversify' — that was the mistake."
)
ws["A2"].alignment = Alignment(wrap_text=True)
ws.row_dimensions[2].height = 48


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
rows = [["Company", "Inserts", "Rank", "Insert spend [$]", "$ / insert",
         "Biking inserts", "Biking % of own buy", "Other inserts"]]
ranked = sorted(REPORTED, key=lambda c: -REPORTED[c])
for i, co in enumerate(ranked, start=1):
    n = REPORTED[co]
    b = by_co_media[co]["Biking Magazines"]
    rows.append([co, n, i, by_co_spend[co], round(by_co_spend[co] / n),
                 b, b / n, n - b])
r2 = block(r, rows)
for i in range(1, len(REPORTED) + 1):
    ws.cell(r + i, 7).number_format = "0.0%"
    ws.cell(r + i, 4).number_format = '"$"#,##0'
    ws.cell(r + i, 5).number_format = '"$"#,##0'
    if rows[i][0] == "WeBike":
        for c in range(1, 9):
            ws.cell(r + i, c).fill = ours
            ws.cell(r + i, c).font = bold
    if rows[i][0] == "BB LLC":
        ws.cell(r + i, 6).fill = good
        ws.cell(r + i, 7).fill = good
r = r2

rows = [
    ["THE EFFICIENCY FINDING", "Value", "Note"],
    ["BB LLC inserts", 24, "industry high — 100% Biking Magazines"],
    ["WeBike inserts", 19, "2nd — but spread across 5 media"],
    ["BB LLC insert spend [$]", 108000, "24 x $4,500"],
    ["WeBike insert spend [$]", by_co_spend["WeBike"], "Health+Biking+Sport+Business+NV+News"],
    ["WeBike reported ad expense [$]", 119165, "inserts $116k + ~$3,165 web/design"],
    ["WeBike $ / insert", round(by_co_spend["WeBike"] / 19), "we pay 36% more per insert than BB LLC"],
    ["BB LLC $ / insert", 4500, "the Biking rate — cheapest medium, best Mtn+Speed pref"],
    ["Inserts we could have bought at $4,500 with our $116k",
     by_co_spend["WeBike"] // 4500,
     "25 Biking inserts vs the 10 we actually bought"],
]
r = block(r, rows, wrap_col=3)
ws.cell(r - 5, 2).fill = good
ws.cell(r - 4, 2).fill = yellow
ws.cell(r - 1, 2).fill = yellow

# mix by medium
rows = [["Inserts by medium (industry)", "Bike Bros", "LiteCycle", "WeBike",
         "BB LLC", "Spoke'd Up", "SpaceBikes", "MILC Bikes", "Industry"]]
firms = ["Bike Bros", "LiteCycle", "WeBike", "BB LLC", "Spoke'd Up", "SpaceBikes", "MILC Bikes"]
for media in MEDIA:
    row = [media] + [by_co_media[co][media] for co in firms] + [by_media[media]]
    rows.append(row)
rows.append(["TOTAL"] + [REPORTED[co] for co in firms] + [grand])
r = block(r, rows)
# highlight WeBike column (D = 4) and BB LLC (E = 5)
for i in range(1, len(MEDIA) + 2):
    ws.cell(r - len(MEDIA) - 2 + i, 4).fill = ours
    if i <= len(MEDIA) and MEDIA[i - 1] == "Biking Magazines":
        ws.cell(r - len(MEDIA) - 2 + i, 5).fill = good

# WeBike vs Q2
rows = [
    ["WEBIKE MIX — Q2 HikeBike1 lock vs Q3 actual", "Q2 HikeBike1", "Q3 HikeBike 1",
     "Q3 Swift Bike", "Q3 total", "Q2→Q3 Hike change"],
]
for media in MEDIA:
    q2 = Q2_HIKE[media]
    h = wbike_hike_q3[media]
    s = wbike_swift_q3[media]
    rows.append([media, q2, h, s, h + s, h - q2])
rows.append([
    "Inserts", hike_q2_n, hike_q3_n, swift_n, hike_q3_n + swift_n, hike_q3_n - hike_q2_n,
])
rows.append([
    "Spend [$]", hike_q2_spend, hike_q3_spend, swift_spend,
    hike_q3_spend + swift_spend, hike_q3_spend - hike_q2_spend,
])
r = block(r, rows)
# color the Biking row change
for i, media in enumerate(MEDIA, start=1):
    if media == "Biking Magazines":
        ws.cell(r - len(MEDIA) - 3 + i, 6).fill = bad
        ws.cell(r - len(MEDIA) - 3 + i, 3).fill = yellow

# identical-product Biking gap
rows = [
    ["IDENTICAL MOUNTAIN PRODUCT — Biking Magazines only", "Inserts", "Spend [$]",
     "Mountain share", "Note"],
    ["BB LLC BB Big Momma → Blu Ruged Ballz (56/73/1, $1,365)", 12, 54000, 0.463,
     "same design, same price as Hike Bike"],
    ["WeBike HikeBike 1 → Hike Bike (56/73/1, $1,365)", 6, 27000, 0.221,
     "we ran this at 12 in Q2 and cut it in half"],
    ["Gap", 6, 27000, 0.242, "half the Biking weight, half-ish the share — plus stock-outs"],
]
r = block(r, rows, wrap_col=5)
ws.cell(r - 3, 4).number_format = "0.0%"
ws.cell(r - 2, 4).number_format = "0.0%"
ws.cell(r - 1, 4).number_format = "0.0%"
for c in range(1, 6):
    ws.cell(r - 2, c).fill = ours
    ws.cell(r - 2, c).font = bold
    ws.cell(r - 1, c).fill = yellow

rows = [
    ["SPEED — Biking Magazines", "Inserts", "Spend [$]", "Speed share", "Note"],
    ["BB LLC Unleash lil pap → Blu Tube Ballz (judgment 76, $1,580)", 12, 54000, None,
     "their advertised Speed brand"],
    ["WeBike Swift Bike (judgment 72, $1,450)", 4, 18000, 0.075,
     "2nd-worst product, 2nd-cheapest, 1/3 the Biking weight"],
    ["MILC Skim MILC MKII (judgment 76)", 5, 22500, None, "more Biking than us on Speed"],
]
r = block(r, rows, wrap_col=5)
ws.cell(r - 2, 4).number_format = "0.0%"
for c in range(1, 6):
    ws.cell(r - 2, c).fill = ours
    ws.cell(r - 2, c).font = bold

# who skipped which brands
rows = [
    ["UNADVERTISED BRANDS (zero major-media inserts)", "Company", "Target", "Note"],
    ["Blu Tail Ballz", "BB LLC", "Mountain", "judgment 70; Blu Ruged (73) takes all 12 Mtn inserts"],
    ["Blu Aero Ballz", "BB LLC", "Speed", "judgment 74; Blu Tube (76) takes all 12 Speed inserts"],
    ["TERRAMean", "Bike Bros", "Mountain", "judgment 72; they advertise TERRAMAX instead"],
    ["Mach 0.6", "Bike Bros", "Speed", "judgment 61; they advertise MACH I.I instead"],
]
r = block(r, rows, wrap_col=4)

# Q4 options — NOT locked
rows = [
    ["Q4 MEDIA OPTIONS — not locked; team vote", "HikeBike 1", "Swift Bike",
     "Inserts", "Insert spend [$]", "vs Q3 spend", "Read"],
    ["A. Copy BB LLC (all Biking)",
     "Biking 12", "Biking 12", 24, 108000, 108000 - by_co_spend["WeBike"],
     "Same insert count AND lower spend than Q3. Drops Sport/News/Business."],
    ["B. Restore Q2 Hike + match BB LLC Speed Biking (recommended)",
     "Health 1 · Biking 12 · Sport 2 · News 1", "Biking 12",
     29, 89000 + 54000, (89000 + 54000) - by_co_spend["WeBike"],
     "Puts Hike Bike back on its proven mix; matches BB LLC's 12 Speed Biking."],
    ["C. Restore Q2 Hike + Speed Biking and Business",
     "Health 1 · Biking 12 · Sport 2 · News 1", "Biking 12 · Business 1",
     30, 89000 + 54000 + 9500, (89000 + 54000 + 9500) - by_co_spend["WeBike"],
     "Adds Speed's #2 medium. Still well inside the $1.01M cash pile."],
    ["Do NOT", "keep Biking at 6", "keep Biking at 4 + NV 2", 19, by_co_spend["WeBike"], 0,
     "Diversifying off Biking is what Q3 did. BB LLC does the opposite."],
]
r = block(r, rows, wrap_col=7)
ws.cell(r - 3, 1).fill = good
for c in range(1, 8):
    ws.cell(r - 3, c).fill = good
ws.cell(r - 1, 1).fill = bad

rows = [
    ["READS", "Detail"],
    ["1. Placement, not ad count",
     "Ad CREATIVE count is not the lever (BB LLC 2 ads, Bike Bros 5). INSERT mix is: "
     "BB LLC 24/24 Biking vs our 10/19. Distinct from the Swift Bike ad-copy problem (judgment 70)."],
    ["2. Q3 'diversify Mountain media' cut the best medium in half",
     "HikeBike 1 Biking 12 → 6. That $27,000 left Biking and went to Swift + New Venture. "
     "Mountain still loves Biking (pref 134). BB LLC kept 12 on the identical product."],
    ["3. We already outspend BB LLC on inserts",
     "$116,000 vs $108,000. Matching them is a MIX change, not a budget increase. "
     "Option B costs +$27k vs Q3 — idle cash covers it 37 times over."],
    ["4. Do not add a 3rd ad or a Recreation buy",
     "Leisure is Rec-skewed (pref 60 for Mountain). Rec stays parked. Improve the Swift Bike "
     "ad (70 → ~77) rather than adding a second Speed creative."],
    ["5. HikeBike 1 copy stays",
     "Ad judgment 80, 2nd in the field. Restore its Q2 media weight; do not redesign the ad."],
]
r = block(r, rows, wrap_col=2)
for cell_row in range(r - 5, r):
    ws.cell(cell_row, 2).alignment = Alignment(wrap_text=True)

# raw dump at the bottom for audit
r += 1
ws.cell(r, 1, "SOURCE DUMP — World Market rows (audit)")
ws.cell(r, 1).font = bold
r += 1
raw_rows = [["Media", "Company", "Ad", "Brand advertised", "Segment", "Inserts",
             "Cost / insert [$]", "Spend [$]"]]
for row in rows_out:
    raw_rows.append(list(row))
block(r, raw_rows)

ws.column_dimensions["A"].width = 62
ws.column_dimensions["B"].width = 42
ws.column_dimensions["C"].width = 28
ws.column_dimensions["D"].width = 22
ws.column_dimensions["E"].width = 22
ws.column_dimensions["F"].width = 22
ws.column_dimensions["G"].width = 72
ws.column_dimensions["H"].width = 16
ws.column_dimensions["I"].width = 14

wb.save(DST)
print(f"\nAdded Q3_Media to {DST}")
