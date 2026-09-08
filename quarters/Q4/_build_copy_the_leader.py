"""Build the BB LLC 'copy the leader' gap analysis and Q4 change list."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DST = "quarters/Q3/Q3Data.xlsx"

# ---------------------------------------------------------------- segment demand
# Brand demand by city, Q3 actual. brand -> (company, segment, [rec, mtn, speed] summed)
BRAND_DEMAND = {
    # BB LLC
    "Blu Ruged Ballz": ("BB LLC", "Mountain", 19 + 6 + 16 + 5, 265 + 48 + 137 + 36, 0),
    "Blu Tail Ballz":  ("BB LLC", "Mountain", 22 + 7 + 19 + 5, 145 + 26 + 75 + 19, 0),
    "Blu Tube Ballz":  ("BB LLC", "Speed", 0, 0, 166 + 29 + 136 + 32),
    "Blu Aero Ballz":  ("BB LLC", "Speed", 0, 0, 135 + 23 + 110 + 26),
    # WeBike
    "Hike Bike":  ("WeBike", "Mountain", 3 + 25 + 24 + 2, 23 + 166 + 156 + 13, 0),
    "Swift Bike": ("WeBike", "Speed", 0, 0, 12 + 79 + 115 + 9),
}
SEG_TOTAL = {"Recreation": 2060, "Mountain": 1621, "Speed": 2878}

print("=" * 78)
print("SEGMENT DEMAND: BB LLC vs WeBike (built from brand-by-city report)")
print("=" * 78)
tot = {}
for co in ("BB LLC", "WeBike"):
    rec = sum(v[2] for v in BRAND_DEMAND.values() if v[0] == co)
    mtn = sum(v[3] for v in BRAND_DEMAND.values() if v[0] == co)
    spd = sum(v[4] for v in BRAND_DEMAND.values() if v[0] == co)
    tot[co] = (rec, mtn, spd)
    print(f"\n{co}:  Rec {rec}  Mountain {mtn}  Speed {spd}  TOTAL {rec+mtn+spd}")
    print(f"   Mountain share {mtn/SEG_TOTAL['Mountain']:6.1%}   Speed share {spd/SEG_TOTAL['Speed']:6.1%}")

print("\nReported totals: BB LLC 1,507 / WeBike 627  ->  reconciles")
bb, we = tot["BB LLC"], tot["WeBike"]
print(f"\nGAP TO CLOSE:  Mountain +{bb[1]-we[1]} units   Speed +{bb[2]-we[2]} units")

# ---------------------------------------------------------------- market perf
print("\n" + "=" * 78)
print("MARKET PERFORMANCE = (avg share in targeted segments/100) x (pct demand served/100)")
print("=" * 78)


def market_perf(mtn, spd, served):
    share = (mtn / SEG_TOTAL["Mountain"] + spd / SEG_TOTAL["Speed"]) / 2
    return share * served, share


v, s = market_perf(we[1], we[2], 423 / 627)
print(f"  WeBike Q3   avg share {s:6.2%}  served {423/627:6.1%}  -> {v:.3f}   (reported 0.098)")
v, s = market_perf(bb[1], bb[2], 1.0)
print(f"  BB LLC Q3   avg share {s:6.2%}  served {1.0:6.1%}  -> {v:.3f}")
v, s = market_perf(we[1], we[2], 1.0)
print(f"  WeBike, stock-outs fixed ONLY (same share)      -> {v:.3f}  = +{v/0.098-1:.0%}")

# ---------------------------------------------------------------- capacity
print("\n" + "=" * 78)
print("CAPACITY: what OC do we need, and do we need a printer?")
print("=" * 78)
PROD, DAYS, OWNED = 0.74, 65, 24
SCEN = [
    ("A  do nothing (post 16.3% ill will)", 525),
    ("B  base forecast", 656),
    ("C  + Rio filled to the 7-person cap", 856),
    ("D  + full web rebuild  <-- THE PLAN", 1171),
    ("E  + NYC store", 1571),
]
print(f"  Owned 24/day at {PROD:.0%} x {DAYS} days = {OWNED*PROD*DAYS:,.0f} units\n")
for name, d in SCEN:
    need = d / (PROD * DAYS)
    print(f"  {name:<38} {d:>5,} units  needs {need:5.2f}/day  "
          f"{'FITS 24 owned' if need <= 24 else 'NEEDS A PRINTER'}")
print(f"\n  Scenario D needs 24.35/day vs 24.00 owned -> short by "
      f"{1171 - OWNED*PROD*DAYS:.0f} units ({(1171-OWNED*PROD*DAYS)/1171:.1%}). No capex required.")
print(f"  Scenario E needs 32.68/day -> 1 printer (32/day = {32*PROD*DAYS:,.0f}) + ~0.7 OT, or 2 printers.")

print("\n  MANUFACTURING PRODUCTIVITY = (pct of SCHEDULED OC used) - (OT as pct of OC used / 2)")
for name, oc, ot in [("WeBike Q3 (8 sched, 1.00 OT)", 8, 1.00),
                     ("BB LLC   (24 sched, 6.90 OT)", 24, 6.90),
                     ("SpaceBikes (30 sched, 0 OT)", 30, 0.0),
                     ("WeBike Q4 plan (24 sched, 0 OT)", 24, 0.0)]:
    print(f"    {name:<32} = {1.0 - (ot/oc)/2:.3f}")
print("  -> Copying BB LLC's 6.90 OT would COST us this indicator. Copy SpaceBikes here: schedule to")
print("     forecast, zero overtime, 1.000. Our 0.938 in Q3 came from 1.00 OT on only 8 scheduled.")

# ---------------------------------------------------------------- media
print("\n" + "=" * 78)
print("MEDIA: copying BB LLC's insert plan COSTS LESS THAN WHAT WE DID")
print("=" * 78)
COST = {"Leisure": 10000, "Health": 7000, "Biking": 4500, "Sport": 10000,
        "Business": 9500, "NewVenture": 5500, "News": 8000}
Q3_US = {"HikeBike 1": {"Health": 1, "Biking": 6, "Sport": 2, "NewVenture": 1, "News": 1},
         "Swift Bike": {"Biking": 4, "Sport": 1, "Business": 1, "NewVenture": 2}}
us_total = 0
for ad, mix in Q3_US.items():
    c = sum(COST[m] * n for m, n in mix.items())
    us_total += c
    print(f"  WeBike Q3 {ad:<12} {sum(mix.values()):>2} inserts  ${c:>7,}   {mix}")
print(f"  WeBike Q3 TOTAL          19 inserts  ${us_total:>7,}   only 10 in Biking")
bb_cost = 24 * COST["Biking"]
print(f"\n  BB LLC Q3                24 inserts  ${bb_cost:>7,}   12 Biking Mountain + 12 Biking Speed")
print(f"\n  Copying BB LLC exactly: {24-19:+d} inserts, ${bb_cost-us_total:+,} cash, "
      f"Biking 10 -> 24.  STRICTLY DOMINANT.")
print(f"  Biking pages in Q3: 51 total, BB LLC owned 24 (47%). At 24 each we co-own the medium.")

# ---------------------------------------------------------------- SEM
print("\n" + "=" * 78)
print("ORGANIC SEM: rank is a step function on clicks, and Mountain rank = ad judgment order")
print("=" * 78)
SEARCHES = {"Mountain": 1027, "Speed": 894}
CTR = {1: 0.2280, 2: 0.1108, 3: 0.0670, 4: 0.0464, 5: 0.0324}
print("  Mountain SERP is EXACTLY ad-judgment order:  81 > 80 > 79 > 78 > 77")
print(f"    pos 1  BB Big Momma (ad 81)   237 clicks")
print(f"    pos 2  HikeBike 1   (ad 80)   115 clicks   <-- us")
print(f"  Beating 81 moves us to pos 1: {237-115:+d} clicks on a page we ALREADY pay $1,000 for.")
print(f"  BB Big Momma's snippet: gears / high tread / 'Highest rated Mountain bike'")
print(f"  HikeBike 1 snippet:     picture / gears / high tread  (we burn a slot on the picture,")
print(f"                          and we omit 'Highest rated' - which is LEGAL, we are tied at 73)")
print("\n  Speed SERP does NOT follow ad judgment (AndStill 78 sits 7th) - likely page tenure.")
for pos, clicks, aj, who in [(1, 200, 76, "Unleash lil pap (BB LLC)"), (2, 98, 77, "The Armstrong 1"),
                             (3, 61, 77, "Speed of Lite 1"), (4, 43, 75, "Skim MILC"),
                             (5, 29, 70, "Swift Bike  <-- us")]:
    print(f"    pos {pos}  {who:<26} ad {aj}  {clicks:>3} clicks")
print(f"  Top 4 all sit at ad judgment 75-77; we are the only sub-75 in the top 5.")
print(f"  Swift Bike at 77 should join that cluster: pos 2-4 = 43-98 clicks (+14 to +69). Landing")
print(f"  position is NOT predictable - treat as upside, not a forecast.")

# ---------------------------------------------------------------- web ops
print("\n" + "=" * 78)
print("WEB OPS: the clearest copy-paste in the whole game")
print("=" * 78)
TACT = [("Toll-free phone", 9000, 6000), ("Advanced shopping cart", 7000, 0),
        ("Page upgrades", 9000, 0), ("Order tracking", 8000, 0)]
bbt = sum(t[1] for t in TACT)
wet = sum(t[2] for t in TACT)
for n, b, w in TACT:
    print(f"  {n:<24} BB LLC ${b:>6,}   WeBike ${w:>6,}   {'MATCH' if b == w else 'FUND IT'}")
print(f"  {'TOTAL':<24} BB LLC ${bbt:>6,}   WeBike ${wet:>6,}   delta ${bbt-wet:+,}/qtr")
print(f"  Web staff:  BB LLC 7 (5 sales + 2 service)   WeBike 3 (2 + 1)   need +4")
print(f"  Web demand: BB LLC 569   WeBike 120")
print(f"  Every firm funding all 4 tactics lands 435-569 web units. Every firm funding fewer: 120-138.")

# ---------------------------------------------------------------- staffing
print("\n" + "=" * 78)
print("STORE STAFFING: we can clone BB LLC's template EXACTLY in both cities")
print("=" * 78)
print("  BB LLC runs the identical crew in both stores: 1 Service / 0 Rec / 3 Mountain / 3 Speed = 7")
print("  (7 per city is the hard cap - no firm anywhere exceeds it.)\n")
print("  Amsterdam  WeBike now: 1 Svc / 3 Mtn / 2 Speed / 1 UNTRAINED = 7  (AT CAP)")
print("             -> train the untrained one into SPEED  = 1/3/3 = EXACT BB LLC clone. Cost $400.")
print("  Rio        WeBike now: 1 Svc / 1 Mtn / 1 Speed / 1 UNTRAINED = 4  (3 slots open)")
print("             -> train untrained into MOUNTAIN, hire 1 Mountain + 2 Speed")
print("                = 1 Svc / 3 Mtn / 3 Speed = 7 = EXACT BB LLC clone.")
print(f"\n  Result: store force 11 -> 14 = BB LLC's exact 14, in the exact same mix.")
print(f"  Rio measured store productivity is 60 units/head (no rival store there) vs Amsterdam's 38.")
print(f"  Filling Rio's 3 slots at 60/head ~ +180 store units.")

# ---------------------------------------------------------------- comp
print("\n" + "=" * 78)
print("COMPENSATION: BB LLC leads BOTH productivity scores in the industry")
print("=" * 78)
for grp, sal, hlth, vac, pen, tot_, prod in [
        ("BB LLC sales", 22000, "Expanded 3,300", 1222, "4% 880", 27402, 78),
        ("WeBike sales", 19000, "FULL 4,180", 1055, "1% 190", 24425, 70),
        ("BB LLC production", 18500, "Expanded 2,775", 1027, "4% 740", 23042, 75),
        ("WeBike production", 16800, "Expanded 2,520", 933, "3% 504", 20757, 72)]:
    print(f"  {grp:<20} salary ${sal:>6,}  {hlth:<16} vac ${vac:>5,}  pen {pen:<8} "
          f"= ${tot_:>7,}  prod {prod}%")
print(f"\n  Both gaps are salary + the Full-coverage mistake. MILC pays MORE than BB LLC on sales")
print(f"  ($27,630 vs $27,402) and scores 8.4 points WORSE - because they use Full coverage + 1 week.")
print(f"  HR indicator = (sales prod + worker prod)/2:  WeBike {(70+72)/2/100:.3f}  BB LLC {(78+75)/2/100:.3f}")
print(f"  Cost to clone BB LLC on 21 sales people: {21*(27402-24425):+,}/yr = "
      f"{21*(27402-24425)/4:+,.0f}/qtr. Against $1,010,838 idle.")

# ---------------------------------------------------------------- ME
print("\n" + "=" * 78)
print("MARKETING EFFECTIVENESS: the copy lets us PASS the leader")
print("=" * 78)
print(f"  BB LLC : brands (73 Mtn + 76 Speed)/2 = 74.5   ads (81 + 76)/2 = 78.5  -> 0.7650  (max)")
print(f"  WeBike : brands (73 + 72)/2 = 72.5             ads (80 + 70)/2 = 75.0  -> 0.7375")
print(f"  WeBike after Swift 77/77: brands 75.0          ads 78.5              -> 0.7675  BEST")
print(f"  If HikeBike 1 also clears 82: brands 75.0      ads 79.5              -> 0.7725")

# ---------------------------------------------------------------- cash
print("\n" + "=" * 78)
print("CASH CHECK (incremental quarterly, vs $1,010,838 on hand and a $300,000 floor)")
print("=" * 78)
ITEMS = [
    ("Media: 24 Biking inserts instead of Q3's mix", bb_cost - us_total),
    ("Web productivity tactics (4 funded, BB LLC levels)", bbt - wet),
    ("Web staff +4 at BB LLC sales pay (quarterly)", round(4 * 27402 / 4)),
    ("Rio store staff +3 at BB LLC sales pay (quarterly)", round(3 * 27402 / 4)),
    ("Compensation clone on the 14 existing sales heads (qtr)", round(14 * (27402 - 24425) / 4)),
    ("Training: 2 untrained + 3 new hires", 5 * 400),
]
sub = sum(i[1] for i in ITEMS)
for n, c in ITEMS:
    print(f"  {n:<56} {c:>+9,}")
print(f"  {'RECURRING SUBTOTAL':<56} {sub:>+9,}")
ONE = [("Web tactic setup fees (assume = 1st quarterly budget)", 27000),
       ("Hiring + design/redesign fees (Swift brand, Swift ad, HikeBike ad)", 40000)]
for n, c in ONE:
    print(f"  {n:<56} {c:>+9,}  (one-time, ESTIMATE - verify in sim)")
one = sum(i[1] for i in ONE)
print(f"\n  Total first-quarter cash impact ~ ${sub+one:,}")
print(f"  Headroom above the $300,000 floor: ${1010838-300000:,}  -> "
      f"{(sub+one)/(1010838-300000):.0%} of it. Comfortable.")
print(f"  Production worker payroll rises too (count not in the report) - confirm in the pro forma.")

print("\n" + "=" * 78)
print("REVENUE / MARGIN SANITY CHECK")
print("=" * 78)
print(f"  Contribution per unit: Hike Bike $747 · Swift Bike at $1,580 $886 (14-speed cuts cost further)")
for mtn_u, spd_u in [(600, 450), (650, 500)]:
    rev = mtn_u * 1365 + spd_u * 1580
    con = mtn_u * 747 + spd_u * 886
    print(f"  {mtn_u} Mountain + {spd_u} Speed = {mtn_u+spd_u:,} units -> "
          f"revenue ${rev:,}  contribution ${con:,}")
print(f"  Q3 actual: revenue $589,720, net income -$111,431 on 423 units built.")
print("=" * 78)

# ================================================================ workbook
wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
bad = PatternFill("solid", fgColor="FFC7CE")
good = PatternFill("solid", fgColor="C6EFCE")
yellow = PatternFill("solid", fgColor="FFEB9C")
ours = PatternFill("solid", fgColor="FFF2CC")
thin = Border(*[Side(style="thin", color="D0D0D0")] * 4)

if "Q4_Copy_The_Leader" in wb.sheetnames:
    del wb["Q4_Copy_The_Leader"]
ws = wb.create_sheet("Q4_Copy_The_Leader")


def block(start, rows, widths=None, wrap=()):
    for i, row in enumerate(rows):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(start + i, c, v)
            cell.border = thin
            if i == 0:
                cell.fill = hdr
                cell.font = bold
            if c in wrap:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
    return start + len(rows) + 2


ws["A1"] = "Q4 — COPY BB LLC: complete change list"
ws["A1"].font = Font(bold=True, size=14)
ws["A2"] = ("BB LLC scored 22 on the Q3 quarterly scorecard (next best 11, WeBike 0) and targets the SAME two "
            "segments we do — Mountain and Speed. Their Mountain bike is byte-identical to ours: 56/73/1 "
            "components, brand judgment 73, price $1,365, price judgment 81/100/100. Same product, same price, "
            "46.3% share vs our 22.1%. There is no product or pricing explanation left. Everything below is "
            "execution we can copy.")
ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A2:F2")
ws.row_dimensions[2].height = 62

r = 4
rows = [["THE FULL GAP", "BB LLC", "WeBike", "Gap", "Copyable in Q4?"]]
GAP = [
    ("Q3 quarterly BSC", "22", "0", "—", "outcome"),
    ("Total demand", "1,507", "627", "2.40x", "outcome"),
    ("Mountain demand (share)", "751 (46.3%)", "358 (22.1%)", "+393 units", "partly"),
    ("Speed demand (share)", "657 (22.8%)", "216 (7.5%)", "+441 units", "partly"),
    ("Demand actually served", "99.9%", "67.5%", "204 units lost", "YES — fully"),
    ("Mountain brand judgment", "73 (Blu Ruged)", "73 (Hike Bike)", "TIED", "already equal"),
    ("Mountain price", "$1,365", "$1,365", "IDENTICAL", "already equal"),
    ("Speed brand judgment", "76 (Blu Tube)", "72 (Swift)", "−4", "YES — we can hit 77"),
    ("Speed price", "$1,580", "$1,450", "−$130", "YES — exactly"),
    ("Mountain ad judgment", "81", "80", "−1", "YES — worth 122 clicks"),
    ("Speed ad judgment", "76", "70", "−6", "YES — we can hit 77"),
    ("Media inserts / spend", "24 / $108,000", "19 / $116,000", "we paid MORE for LESS", "YES — and it SAVES $8,000"),
    ("Inserts in Biking Magazines", "24 (100%)", "10 (53%)", "−14", "YES — exactly"),
    ("Organic SEM clicks (same $2,000)", "437", "144", "−293", "partly"),
    ("Organic rank Mountain / Speed", "#1 / #1", "#2 / #5", "—", "Mountain yes, Speed likely"),
    ("Stores", "2 (AMS + NYC)", "2 (AMS + Rio)", "same count", "n/a"),
    ("Store sales force", "14 (7 + 7)", "11 (7 + 4)", "−3", "YES — exactly 14"),
    ("Store crew mix per city", "1 Svc / 3 Mtn / 3 Speed", "AMS 1/3/2+1 untrained", "mix + training", "YES — exactly"),
    ("Web sales force", "7 (5 + 2)", "3 (2 + 1)", "−4", "YES — exactly"),
    ("Web productivity tactics funded", "4 of 4 ($33,000)", "1 of 4 ($6,000)", "−3 tactics", "YES — exactly"),
    ("Web units", "569", "120", "−449", "YES — pattern says 435+"),
    ("Fixed capacity / day", "24", "24", "IDENTICAL", "already equal"),
    ("Scheduled operating capacity / day", "24 (100%)", "8 (33%)", "−16", "YES — fully"),
    ("Overtime / day", "6.90", "1.00", "—", "DON'T copy (see below)"),
    ("Sales compensation / productivity", "$27,402 / 78%", "$24,425 / 70%", "−$2,977 / −8 pts", "YES — exactly"),
    ("Worker compensation / productivity", "$23,042 / 75%", "$20,757 / 72%", "−$2,285 / −3 pts", "YES — exactly"),
    ("Marketing Effectiveness", "0.765 (max)", "0.738", "−0.027", "YES — we can reach 0.7675"),
    ("Asset Management", "0.920", "0.353", "idle cash", "YES — spend it on the above"),
]
for g in GAP:
    rows.append(list(g))
r2 = block(r, rows, wrap=(5,))
for i in range(1, len(rows)):
    lab = rows[i][0]
    if "IDENTICAL" in rows[i][3] or "TIED" in rows[i][3]:
        for c in (2, 3, 4):
            ws.cell(r + i, c).fill = good
    if lab in ("Demand actually served", "Scheduled operating capacity / day",
               "Web productivity tactics funded", "Web units", "Media inserts / spend"):
        ws.cell(r + i, 3).fill = bad
        ws.cell(r + i, 5).fill = good
    if lab == "Overtime / day":
        ws.cell(r + i, 5).fill = yellow
r = r2

rows = [["#", "CHANGE — enter these in the sim", "Exact value", "Cash / qtr", "Why (evidence)"]]
CHANGES = [
    (1, "Schedule operating capacity", "24 / day (100% of owned)", "$0",
     "BB LLC schedules all 24 owned; we scheduled 8 and stocked out on 204 of 627 units. "
     "24/day x 74% x 65 = 1,154 units, which covers the 1,171 forecast within 1.5%. "
     "Also cuts unit cost: under-scheduling pushed labour/unit from $130 to $205."),
    (2, "Overtime", "0.00 — do NOT copy BB LLC's 6.90", "$0",
     "Manufacturing Productivity = utilisation of scheduled OC minus half the OT ratio. BB LLC's "
     "6.90 OT scores 0.856; scheduling 24 with zero OT scores 1.000. Copy SpaceBikes here, not BB LLC."),
    (3, "Printers", "Buy NONE (unless NYC — then 1)", "$0",
     "We already own the same 24/day BB LLC does. The Q2 $148,652 excess-capacity charge came from "
     "SCHEDULING unused capacity, not owning it. Only scenario E (NYC) needs 32.7/day."),
    (4, "Swift Bike design: gears", "24-speed -> 14-speed (2 x 7)", "design fee",
     "+3 judgment. Swift Bike is the ONLY Speed brand in the industry running a 24-speed mountain "
     "drivetrain — it inherited Hike Bike's. 14-speed is also simpler and should cut unit cost."),
    (5, "Swift Bike design: decals", "ADD colourful thin brushstrokes", "included",
     "+2 judgment. Takes Swift Bike 72 -> 77, tying the industry Speed ceiling and BEATING BB LLC's "
     "Blu Tube at 76. Verified against a model that reproduces all 16 Mountain/Speed scores exactly."),
    (6, "Hike Bike design", "NO CHANGE — 56/73/1", "$0",
     "73 is the industry Mountain ceiling, tied with BB LLC's own Blu Ruged Ballz. Nothing to gain."),
    (7, "Swift Bike price", "$1,450 -> $1,580", "revenue +",
     "$1,580 is BB LLC's exact Speed price and draws ZERO Speed resistance (proven by Blu Aero, Blu "
     "Tube, Armstrong at the same price). Resistance starts at $1,749. Contribution $756 -> $886."),
    (8, "Hike Bike price", "HOLD $1,365", "$0",
     "Identical to BB LLC's Mountain price, and tier 81 is the highest tier still scoring 100 in "
     "Mountain. Already optimal."),
    (9, "Brand priority", "Hike Bike 1, Swift Bike 2 (unchanged)", "$0",
     "Same pattern BB LLC uses (Mountain brand at priority 1). Once OC is scheduled to the forecast, "
     "priority stops rationing anything."),
    (10, "Swift Bike ad — REBUILD", "7 benefits, ranked (see next block)", "design fee",
     "70 -> target 77. Our ad is the only Speed ad in the industry selling lights, reflectors and "
     "'great price'; it leaks 52 Rec / 36 Mtn against 12-46 for everyone else."),
    (11, "HikeBike 1 ad — try to clear 82", "add 'Highest rated Mountain bike' + 'Local sales & service'",
     "design fee",
     "Mountain SERP rank is EXACTLY ad-judgment order (81>80>79>78>77). Beating BB Big Momma's 81 "
     "moves us from position 2 to 1: 115 -> 237 clicks on a page we already pay $1,000 for. Both "
     "claims are legal (we are tied at 73) and BB LLC runs both. PREVIEW FIRST — keep Q3 copy if it drops."),
    (12, "Media: HikeBike 1 inserts", "Biking 12 — and nothing else", "$54,000",
     "Exactly BB LLC's BB Big Momma plan. We cut HikeBike 1 Biking from 12 to 6 in Q3 to 'diversify' "
     "and spent the money on Sport/News/Venture instead. Biking is the #1 Mountain medium."),
    (13, "Media: Swift Bike inserts", "Biking 12 — and nothing else", "$54,000",
     "Exactly BB LLC's Unleash lil pap plan. We ran only 4 Biking for Swift Bike. BB LLC owned 24 of "
     "the industry's 51 Biking pages; at 24 each we co-own the medium."),
    (14, "Media total", "24 inserts / $108,000 (was 19 / $116,000)", "-$8,000",
     "The winning media plan is CHEAPER than what we did. More inserts, best medium, less money. "
     "Strictly dominant — the single easiest change on this list."),
    (15, "Web tactic: toll-free phone", "$6,000 -> $9,000", "$3,000",
     "BB LLC's level. Our one already-funded tactic."),
    (16, "Web tactic: advanced shopping cart", "$0 -> $7,000", "$7,000",
     "Never funded. BB LLC, LiteCycle, SpaceBikes and MILC all fund it."),
    (17, "Web tactic: page upgrades", "$0 -> $9,000", "$9,000",
     "We funded this in Q2 and STOPPED it in Q3. BB LLC spends $9,000."),
    (18, "Web tactic: order tracking", "$0 -> $8,000", "$8,000",
     "Never funded. Every firm that funds all four lands 435-569 web units; every firm that funds "
     "fewer lands 120-138. We are at 120."),
    (19, "Web staff", "3 -> 7 (5 sales + 2 service)", "~$27,400",
     "BB LLC's exact web crew. 7 is the hard cap. Web demand per head: BB LLC 81.3, WeBike 40.0 (last)."),
    (20, "Amsterdam store: train the untrained", "1 untrained -> SPEED specialist", "$400",
     "Gives 1 Service / 3 Mountain / 3 Speed = BB LLC's exact crew. Amsterdam is AT the 7-person cap, "
     "so training is the only legal move there."),
    (21, "Rio store: train the untrained", "1 untrained -> MOUNTAIN specialist", "$400",
     "Only Bike Bros (last place) also carries untrained staff. The other five firms train everyone."),
    (22, "Rio store: hire 3", "+1 Mountain, +2 Speed -> 1 Svc / 3 Mtn / 3 Speed", "~$20,600",
     "Clones BB LLC's crew in our second store and takes total store force to 14 = BB LLC's 14. Rio "
     "measures 60 units/head (we are the ONLY store in the city) vs Amsterdam's 38, so this is our "
     "highest-return staffing move: roughly +180 store units."),
    (23, "Sales compensation", "$22,000 salary / Expanded / 2 wk / 4% = $27,402", "~$15,600",
     "BB LLC's exact package — the industry's best sales productivity at 78%. Drop Full coverage: MILC "
     "pays MORE ($27,630) on Full + 1 week and scores 8.4 points WORSE. Salary is the #1 lever (87)."),
    (24, "Production compensation", "$18,500 salary / Expanded / 2 wk / 4% = $23,042", "~payroll",
     "BB LLC's exact package — best worker productivity at 75%. Our mix is already right; we are "
     "simply under on salary."),
    (25, "Productivity assumption in the planner", "74% — NEVER 85%", "$0",
     "Q3 planned 85% and delivered 70/72. BB LLC's 78% sales is the industry ceiling. Pay buys a few "
     "HR points, not fifteen productivity points."),
    (26, "Stock issue", "NONE", "$0",
     "$1,010,838 already sits idle; that is what Asset Management 0.353 punishes. Spend cash, "
     "don't raise more."),
    (27, "Target segments", "Mountain (primary) / Speed (secondary)", "$0",
     "Unchanged — identical to BB LLC's targeting. No action, just don't touch it."),
    (28, "Organic SEM pages", "2 pages, $1,000 each (unchanged)", "$0",
     "BB LLC gets 437 clicks from the SAME 2 pages and SAME $2,000 we spend. Rank, not budget, is the "
     "gap — and rank follows ad judgment, so changes 10 and 11 are the SEM fix."),
    (29, "Rebates", "NONE", "$0", "Only the two Recreation price leaders use them."),
    (30, "Recreation brand", "DO NOT ADD — park for Q5", "$0",
     "BB LLC has no Recreation brand scoring 70+ either. Brand count does not drive share: Bike Bros "
     "has 5 brands and finished last; SpaceBikes has 2 and leads cumulatively. Capacity binds."),
]
for c in CHANGES:
    rows.append(list(c))
r2 = block(r, rows, wrap=(5,))
for i in range(1, len(rows)):
    ws.cell(r + i, 5).alignment = Alignment(wrap_text=True, vertical="top")
    if rows[i][0] in (1, 2, 14, 19, 22, 23):
        ws.cell(r + i, 2).fill = good
        ws.cell(r + i, 2).font = bold
    if rows[i][0] in (2, 3, 6, 8, 11, 26, 30):
        ws.cell(r + i, 3).fill = yellow
    if rows[i][0] == 14:
        ws.cell(r + i, 4).fill = good
r = r2

rows = [["SWIFT BIKE AD — rank order to enter (draft, preview before committing)", "Note"]]
for i, (b, n) in enumerate([
        ("Mention brand name", "fills the Ad/Brand columns; does NOT consume an organic snippet slot"),
        ("Ride a wind-cheater", "8 of 8 rival Speed ads run this; we were missing it. Aero frame makes it true."),
        ("Roll fast with racing tires", "7 of 8 rivals; the bike HAS racing tires"),
        ("Elite look - a ride of distinction", "6 of 8 rivals, including both 77-point ads"),
        ("A tailor-made bike just for you! 3D printing", "on both best ads; we actually 3D-print"),
        ("Local sales & service", "we are 80.9% store — and never claimed it"),
        ("Picture of road race", "the one thing Q3 got right; ranked lower so text fills the snippet")], 1):
    rows.append([f"{i}. {b}", n])
rows.append(["DROP: lights, reflectors, 'carbon fiber quality at a great price', carbon-light at #2",
             "Zero of 8 rival Speed ads run lights, reflectors or great price. Ours ran all three. "
             "Do not sell 'great price' in the quarter we raise the price to $1,580."])
rows.append(["Resulting snippet: wind-cheater / racing tires / elite look",
             "Identical to Unleash lil pap, the Speed SERP #1 with 200 clicks."])
rows.append(["Do NOT add a second Swift Bike ad",
             "BB LLC runs 2 ads for 4 brands and leads. Bike Bros runs 5 for 5 and is last."])
r2 = block(r, rows, wrap=(2,))
ws.cell(r + 8, 1).fill = bad
ws.cell(r + 9, 1).fill = good
r = r2

rows = [["EXPECTED RESULT", "Q3 actual", "Q4 after the copy", "Basis"]]
for row in [
    ("Units built", "423", "~1,154", "24/day x 74% x 65 days, zero overtime"),
    ("Demand served", "67.5%", "~99%", "capacity finally matches the forecast"),
    ("Web units", "120", "435+", "all 4 tactics + 7 staff — the pattern holds for 4 of 4 firms"),
    ("Rio store units", "240", "~420", "3 more heads at Rio's measured 60/head"),
    ("Mountain share", "22.1%", "~35%", "supply + 12 Biking inserts + Rio depth (BB LLC proves 46.3% is reachable on this product)"),
    ("Speed share", "7.5%", "~17%", "brand 77 + ad 77 + $1,580 + 12 Biking inserts + 3 Speed heads per store"),
    ("Marketing Effectiveness", "0.738", "0.7675", "arithmetic — formula verified on 3 firms"),
    ("Manufacturing Productivity", "0.938", "1.000", "schedule to forecast with zero OT"),
    ("Market Performance", "0.098", "~0.26", "avg share ~26% x ~99% served (BB LLC scored ~0.345)"),
    ("Revenue", "$589,720", "~$1.5-1.7M", "600 Mountain at $1,365 + 450 Speed at $1,580"),
    ("Asset Management", "0.353", "improves", "idle cash converted into staff, web and inserts"),
]:
    rows.append(list(row))
r2 = block(r, rows, wrap=(4,))
for i in range(1, len(rows)):
    ws.cell(r + i, 3).fill = good
    ws.cell(r + i, 3).font = bold
r = r2

rows = [["WHAT WE DELIBERATELY DO NOT COPY", "Reason"]]
for row in [
    ("BB LLC's 6.90 overtime", "It costs Manufacturing Productivity (0.856 vs 1.000). We own the same 24 "
     "printers — we can schedule to the forecast without it. SpaceBikes is the better model here: "
     "32 owned, 30 scheduled, zero OT."),
    ("BB LLC's 4 brands", "Their 3rd and 4th brands (Blu Tail, Blu Aero) run UNADVERTISED as overflow "
     "catchers at a higher price. That only works with spare capacity. Capacity binds for us in Q4. "
     "File it for Q5."),
    ("BB LLC's NYC store", "Tier 2 — they hold 82.8% of NYC Mountain and it needs a printer plus 7 new "
     "hires. If we go, go via SPEED (880 units, no dominant firm, Spoke'd Up's Speed ad is under 70), "
     "not Mountain. Depth in Rio and the web pays back faster."),
    ("Bike Bros' 5 ads and 5 brands", "Most ads and most brands in the industry — and last place. "
     "Their two Speed pages cannibalise each other into positions 7 and 8."),
    ("MILC's compensation", "Highest sales spend in the industry ($27,630) and 8.4 points WORSE "
     "satisfaction than BB LLC, because of Full coverage + 1 week vacation. Mix beats total."),
]:
    rows.append(list(row))
r2 = block(r, rows, wrap=(2,))
for i in range(1, len(rows)):
    ws.cell(r + i, 2).fill = yellow
r = r2

rows = [["UNVERIFIED / VOTE REQUIRED", "Status"]]
for row in [
    ("Every item above", "ANALYSIS ONLY — nothing is locked. Majority vote required before entry."),
    ("Swift Bike ad landing score", "Draft ranks are modelled on the two 77-point ads. PREVIEW in the "
     "designer and read the score before committing."),
    ("HikeBike 1 modification", "Only commit if the preview clears 82. If it drops below 80, keep the "
     "Q3 copy exactly as is."),
    ("Speed organic rank after the fix", "Speed SERP does NOT follow ad judgment (AndStill scores 78 and "
     "sits 7th), so the landing position is unpredictable. Treat as upside, not forecast."),
    ("Web tactic setup fees", "Q2 charged setup equal to the first quarterly budget. Assume ~$27,000 "
     "one-time on the three restarts — CONFIRM in the sim."),
    ("Hiring, design and redesign fees", "Estimated at ~$40,000 combined. Read the actual figures."),
    ("Production worker headcount", "Not in the Competitors' Profiles report, so the production payroll "
     "delta is not costed. Check the pro forma."),
    ("Printer price", "~$240,000 assumed for the NYC tier. Verify before any purchase."),
    ("Share targets", "Directional. They assume the unit build lands and the mix follows Q3 proportions."),
    ("Before advancing", "Run Production Simulation, then Cash Flow and pro forma. Confirm ending "
     "Cash + CD >= $300,000."),
]:
    rows.append(list(row))
block(r, rows, wrap=(2,))

ws.column_dimensions["A"].width = 46
ws.column_dimensions["B"].width = 40
ws.column_dimensions["C"].width = 30
ws.column_dimensions["D"].width = 16
ws.column_dimensions["E"].width = 78
ws.column_dimensions["F"].width = 20

wb.save(DST)
print("\nAdded Q4_Copy_The_Leader to Q3Data.xlsx")
