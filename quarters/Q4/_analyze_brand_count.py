"""Does a second brand per segment actually work? Test it against every firm, then spec the copy."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DST = "quarters/Q3/Q3Data.xlsx"

SEG_TOTAL = {"Recreation": 2060, "Mountain": 1621, "Speed": 2878}

# brand -> (company, target segment, judgment, price, advertised?, [(rec, mtn, speed) per city summed])
B = {
    "MACH I.I":        ("Bike Bros",  "Speed",      77, 1749, True,  (0, 0, 155)),
    "Mach 0.6":        ("Bike Bros",  "Speed",      61, 1579, False, (0, 0, 25)),
    "TERRAMAX":        ("Bike Bros",  "Mountain",   73, 1499, True,  (6, 92, 0)),
    "TERRAMean":       ("Bike Bros",  "Mountain",   72, 1349, False, (10, 80, 0)),
    "MountainCruise1": ("Bike Bros",  "Recreation", 76, 1199, True,  (122, 0, 0)),
    "LiteSpeed Pro+":  ("LiteCycle",  "Speed",      77, 1515, True,  (0, 0, 39 + 27 + 126 + 123)),
    "LiteSpeed+":      ("LiteCycle",  "Speed",      74, 1325, True,  (0, 0, 29 + 20 + 93 + 91)),
    "LiteTrail Pro":   ("LiteCycle",  "Mountain",   70, 1300, True,  (13 + 13 + 36 + 34, 62 + 46 + 110 + 122, 0)),
    "Hike Bike":       ("WeBike",     "Mountain",   73, 1365, True,  (3 + 25 + 24 + 2, 23 + 166 + 156 + 13, 0)),
    "Swift Bike":      ("WeBike",     "Speed",      72, 1450, True,  (0, 0, 12 + 79 + 115 + 9)),
    "Blu Ruged Ballz": ("BB LLC",     "Mountain",   73, 1365, True,  (19 + 6 + 16 + 5, 265 + 48 + 137 + 36, 0)),
    "Blu Tail Ballz":  ("BB LLC",     "Mountain",   70, 1365, False, (22 + 7 + 19 + 5, 145 + 26 + 75 + 19, 0)),
    "Blu Tube Ballz":  ("BB LLC",     "Speed",      76, 1580, True,  (0, 0, 166 + 29 + 136 + 32)),
    "Blu Aero Ballz":  ("BB LLC",     "Speed",      74, 1580, False, (0, 0, 135 + 23 + 110 + 26)),
    "Spoke'd Easy":    ("Spoke'd Up", "Recreation", 73, 999,  True,  (237 + 21 + 217 + 17, 0, 0)),
    "Spoke'd Speed":   ("Spoke'd Up", "Speed",      75, 1499, True,  (0, 0, 175 + 11 + 163 + 13)),
    "The Armstrong":   ("SpaceBikes", "Speed",      76, 1580, True,  (0, 0, 272 + 43 + 223 + 50)),
    "Mars Rover":      ("SpaceBikes", "Recreation", 73, 1050, True,  (302 + 77 + 267 + 64, 0, 0)),
    "Whole MILC MKII": ("MILC Bikes", "Recreation", 74, 950,  True,  (68 + 70 + 275 + 58, 0, 0)),
    "Skim MILC MKII":  ("MILC Bikes", "Speed",      76, 1450, True,  (0, 0, 52 + 37 + 197 + 42)),
}
SEG_I = {"Recreation": 0, "Mountain": 1, "Speed": 2}
# Bike Bros only sells in Amsterdam, so their city rows are the Amsterdam figures only.

CAPACITY = {  # firm -> (fixed/day, scheduled OC, OT, worker prod %)
    "BB LLC": (24, 24, 6.90, 75), "SpaceBikes": (32, 30, 0.0, 73),
    "LiteCycle": (16, 14, 1.75, 74), "Spoke'd Up": (16, 15, 1.88, 72),
    "MILC Bikes": (24, 17, 0.0, 73), "WeBike": (24, 8, 1.00, 72),
    "Bike Bros": (16, 15, 0.0, 70),
}
STORE_SF = {"BB LLC": 14, "SpaceBikes": 14, "LiteCycle": 10, "Spoke'd Up": 11,
            "MILC Bikes": 7, "WeBike": 11, "Bike Bros": 7}
WEB_SF = {"BB LLC": 7, "SpaceBikes": 7, "LiteCycle": 7, "Spoke'd Up": 3,
          "MILC Bikes": 7, "WeBike": 3, "Bike Bros": 0}

print("=" * 92)
print("DOES BRAND COUNT DRIVE SEGMENT SHARE?  (one row per firm per segment they target)")
print("=" * 92)

rows = []
for seg in ("Speed", "Mountain", "Recreation"):
    firms = {}
    for name, (co, tgt, j, price, adv, dem) in B.items():
        if tgt != seg:
            continue
        firms.setdefault(co, []).append((name, j, price, adv, dem[SEG_I[seg]]))
    print(f"\n  {seg} (segment total {SEG_TOTAL[seg]:,} units)")
    print(f"    {'Firm':<12}{'Brands':>7}{'Units':>7}{'Share':>8}{'OC/day':>8}{'Store SF':>9}{'Web SF':>7}   brands")
    out = []
    for co, bl in firms.items():
        units = sum(x[4] for x in bl)
        out.append((co, len(bl), units, units / SEG_TOTAL[seg]))
    for co, n, units, sh in sorted(out, key=lambda x: -x[3]):
        eff = CAPACITY[co][1] + CAPACITY[co][2]
        detail = ", ".join(f"{x[0]} {x[1]}{'' if x[3] else ' (no ad)'}" for x in firms[co])
        star = "  <-- us" if co == "WeBike" else ""
        print(f"    {co:<12}{n:>7}{units:>7}{sh:>7.1%}{eff:>8.2f}{STORE_SF[co]:>9}{WEB_SF[co]:>7}   {detail}{star}")
        rows.append((seg, co, n, units, sh, eff, STORE_SF[co] + WEB_SF[co]))

print("\n" + "=" * 92)
print("THE VERDICT ON BRAND COUNT")
print("=" * 92)
print("  Speed:    SpaceBikes gets 20.4% with ONE brand. LiteCycle gets 19.0% with TWO.")
print("            Bike Bros gets 6.3% with TWO - worse than our 7.5% with one.")
print("  Mountain: Bike Bros has TWO Mountain brands and 10.6%. LiteCycle has ONE and 21.0%.")
print("  -> Brand count explains nothing. Effective capacity and sales coverage explain everything.")
print("\n  Same table sorted by capacity instead:")
for seg in ("Speed", "Mountain"):
    print(f"    {seg}:")
    for seg2, co, n, units, sh, eff, sf in sorted([r for r in rows if r[0] == seg], key=lambda x: -x[5]):
        print(f"      OC {eff:>5.2f}/day  {sf:>2} sales people  {n} brand(s)  ->  {sh:>5.1%}   {co}")

print("\n" + "=" * 92)
print("BUT THE SECOND BRAND IS NOT NOTHING - it is 41% of the leader's business")
print("=" * 92)
for co in ("BB LLC", "LiteCycle", "Bike Bros"):
    bl = [(n, v) for n, v in B.items() if v[0] == co]
    tot = sum(sum(v[5]) for _, v in bl)
    second = 0
    for seg in ("Speed", "Mountain"):
        same = sorted([(n, v) for n, v in bl if v[1] == seg],
                      key=lambda x: -sum(x[1][5]))
        if len(same) > 1:
            second += sum(sum(v[5]) for _, v in same[1:])
    if second:
        print(f"  {co:<12} total demand {tot:>5}   from 2nd brands in a segment {second:>5}  "
              f"({second/tot:>5.1%})")
print("\n  BB LLC's two unadvertised brands (Blu Tail 318, Blu Aero 294) pulled 612 units.")
print("  Bike Bros' second brands pulled only 105 - they have 1 store, 0 web and 15 OC.")
print("  -> A second brand converts capacity and coverage into units. With neither, it does nothing.")

print("\n" + "=" * 92)
print("WHAT BB LLC'S SECOND BRANDS ACTUALLY LOOK LIKE (this is the pattern to copy)")
print("=" * 92)
print(f"  {'Brand':<18}{'Seg':<10}{'Judg':>5}{'Price':>7}{'Ad?':>6}{'Units':>7}")
for n in ("Blu Ruged Ballz", "Blu Tail Ballz", "Blu Tube Ballz", "Blu Aero Ballz"):
    co, seg, j, p, adv, dem = B[n]
    print(f"  {n:<18}{seg:<10}{j:>5}{p:>7}{'yes' if adv else 'NO':>6}{sum(dem):>7}")
print("\n  KEY: the second brand carries the SAME price as the lead brand, not a higher one,")
print("  and a slightly LOWER judgment (70 vs 73, 74 vs 76). Same price, cheaper spec.")
print("  Blu Aero = 14sp, decals NO, reflectors yes, lights NO   -> 74")
print("  Blu Tube = 14sp, decals yes, reflectors no,  lights yes -> 76")
print("  LiteSpeed+ = 14sp, decals NO, reflectors NO, lights yes -> 74  (same idea, different firm)")
print("\n  HYPOTHESIS (unverified - we have no component price list): the second brand deliberately")
print("  omits cheap-but-costly accessories to cut unit cost while charging the same price, since")
print("  Speed shows ZERO price resistance up to $1,749. Alternative reading: they simply did not")
print("  bother optimising brand #2. Read component costs on the Design Brand screen to decide.")

print("\n" + "=" * 92)
print("CAN WE AFFORD A THIRD BRAND IN Q4?  (this is the whole decision)")
print("=" * 92)
PROD, DAYS, OWNED = 0.74, 65, 24
cap = OWNED * PROD * DAYS
plan = 1171
print(f"  Q4 plan of record demand forecast          {plan:>6,} units")
print(f"  Capacity at 24/day x 74% x 65 days         {cap:>6,.0f} units")
print(f"  Slack before any new brand                 {cap-plan:>6,.0f} units  ({(cap-plan)/plan:+.1%})")
print("\n  A second Speed brand, scaled from the firms that run one:")
for co, lead, second in [("BB LLC", 363, 294), ("LiteCycle", 315, 233)]:
    print(f"    {co:<12} lead Speed brand {lead}, second {second}  -> second = {second/lead:.0%} of lead")
our_speed = 490
for ratio, label in ((0.74, "LiteCycle ratio"), (0.81, "BB LLC ratio")):
    add = round(our_speed * ratio)
    tot = plan + add
    need = tot / (PROD * DAYS)
    short = tot - cap
    print(f"\n  At the {label} ({ratio:.0%}) on ~{our_speed} Speed units: +{add} units of demand")
    print(f"    total demand {tot:,}  vs capacity {cap:,.0f}  ->  {short:,.0f} unfilled ({short/tot:.0%} stock-out)")
    print(f"    required OC {need:.1f}/day vs 24 owned  ->  needs {'1 more printer' if need <= 32 else '2 more printers'}")
    print(f"    Q5 ill will would be about {short/tot/2:.1%} (half the unmet percentage)")
print(f"\n  Q3 for reference: 204 of 627 unfilled = 33% stock-out -> 16.3% ill will -> LAST place.")
print("  Adding a brand in Q4 without a printer recreates the exact failure we are fixing.")

print("\n" + "=" * 92)
print("SPEC SHEET FOR THE BRANDS WE SHOULD EVENTUALLY ADD")
print("=" * 92)
print("""
  BRAND 3 - second SPEED brand  (highest priority once capacity exists)
    Why Speed: largest segment (2,878 vs Mountain 1,621), we hold only 7.5%, and it is the
    least price-sensitive segment in the game - judgment stays 100 all the way to $1,580 and
    resistance only starts at $1,749. Contribution at $1,580 is $886/unit, our best.
    Spec (copy Blu Tube / The Armstrong / Skim MILC - three firms converged on this exact 76):
      aerodynamic frame - racing tires - precision brakes - basic drop-down bars
      polymer gel racing seat - 14 speed (2x7) - decals - lights - NO reflectors
      no carrier, no suspension                                        -> judgment 76
    Optional: add reflectors for 77 if the screen shows they are cheap.
    Price $1,580 (same as Swift Bike, exactly as BB LLC prices both of theirs).
    Advertising: NONE at first. BB LLC leaves both second brands unadvertised and leads the
    industry. Media money stays concentrated on the two ads we already have.

  BRAND 4 - second MOUNTAIN brand  (lower priority)
    Mountain is the smallest and slowest segment and we are already #2 at 22.1%.
    Spec (copy Blu Tail / LiteTrail Pro = 70): the Hike Bike recipe but with the polymer gel
    COMFORT seat instead of all-purpose (-3). Price $1,365, same as Hike Bike. No ad.
    Only worth it after Speed is saturated.

  BRAND 5 - RECREATION  (Q5+ at the earliest, and it fights our margin strategy)
    Spec = MountainCruise1's exact 76, the only brand at the Recreation ceiling:
      comfort frame - hybrid tires - STANDARD DISC brakes - comfort straight bars
      7 speed (1x7) - polymer gel comfort seat - reflectors + decals + lights
      + plastic basket + front shocks                                  -> judgment 76
    WARNING: every Recreation winner sits at price judgment 100, which means cheapest.
    Mars Rover $1,050 net, Whole MILC $950 net, Spoke'd Easy $999. Our $1,365 scores 81 in
    Recreation. Entering means low margin, which contradicts our locked profit-margin-leader
    intent. Hike Bike already scores 56 with Recreation buyers and pulled 54 units of pure
    spillover, so there is no cheap way in.
""")
print("=" * 92)

# ------------------------------------------------------------------ workbook
wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
bad = PatternFill("solid", fgColor="FFC7CE")
good = PatternFill("solid", fgColor="C6EFCE")
yellow = PatternFill("solid", fgColor="FFEB9C")
ours = PatternFill("solid", fgColor="FFF2CC")
thin = Border(*[Side(style="thin", color="D0D0D0")] * 4)

if "Q4_Brand_Count" in wb.sheetnames:
    del wb["Q4_Brand_Count"]
ws = wb.create_sheet("Q4_Brand_Count")


def block(start, rws, wrap=()):
    for i, row in enumerate(rws):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(start + i, c, v)
            cell.border = thin
            if i == 0:
                cell.fill = hdr
                cell.font = bold
            if c in wrap:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
    return start + len(rws) + 2


ws["A1"] = "Do we add more bikes?  ANSWER: not in Q4 — capacity binds. Spec for Q5 below."
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("Brand count does not explain segment share anywhere in this industry. Effective capacity and "
            "sales coverage do. BB LLC's second brands work because they own 30.9 effective capacity and 21 "
            "sales people; Bike Bros' second brands do nothing because they own 15 and 7.")
ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A2:G2")
ws.row_dimensions[2].height = 46

r = 4
rws = [["Segment", "Firm", "Brands in segment", "Units", "Share", "Effective OC/day", "Sales people"]]
for seg, co, n, units, sh, eff, sf in rows:
    rws.append([seg, co, n, units, round(sh, 4), round(eff, 2), sf])
r2 = block(r, rws)
for i in range(1, len(rws)):
    ws.cell(r + i, 5).number_format = "0.0%"
    ws.cell(r + i, 6).number_format = "0.00"
    if rws[i][1] == "WeBike":
        for c in range(1, 8):
            ws.cell(r + i, c).fill = ours
            ws.cell(r + i, c).font = bold
r = r2

rws = [
    ["THE COUNTEREXAMPLES THAT KILL THE 'MORE BRANDS' ARGUMENT", "Brands", "Share", "Effective OC", "Sales people"],
    ["SpaceBikes in Speed", 1, "20.4%", 30.00, 21],
    ["LiteCycle in Speed", 2, "19.0%", 15.75, 17],
    ["Bike Bros in Speed", 2, "6.3%", 15.00, 7],
    ["WeBike in Speed", 1, "7.5%", 9.00, 14],
    ["LiteCycle in Mountain", 1, "21.0%", 15.75, 17],
    ["Bike Bros in Mountain", 2, "10.6%", 15.00, 7],
    ["Read", None, None, None, "One brand with capacity beats two without it, every time. Bike Bros runs "
     "the most brands in the industry (5) and finished last."],
]
r2 = block(r, rws, wrap=(5,))
ws.cell(r + 1, 2).fill = good
ws.cell(r + 3, 2).fill = bad
ws.cell(r + 6, 2).fill = good
ws.cell(r + 7, 4).fill = bad
r = r2

rws = [
    ["BUT SECOND BRANDS ARE 41% OF THE LEADER'S DEMAND", "Judgment", "Price", "Advertised?", "Units"],
    ["Blu Ruged Ballz (Mountain, lead)", 73, 1365, "yes — BB Big Momma 81", 532],
    ["Blu Tail Ballz (Mountain, second)", 70, 1365, "NO AD AT ALL", 318],
    ["Blu Tube Ballz (Speed, lead)", 76, 1580, "yes — Unleash lil pap 76", 363],
    ["Blu Aero Ballz (Speed, second)", 74, 1580, "NO AD AT ALL", 294],
    ["Second brands combined", None, None, None, 612],
    ["Pattern", None, None, None, "Same price as the lead brand — NOT higher. Slightly lower judgment. "
     "No advertising. 612 units, or 41% of their 1,507 total."],
    ["Hypothesis (UNVERIFIED)", None, None, None, "The second brand omits accessories to cut unit cost "
     "while charging the same price, because Speed shows zero resistance to $1,580. Blu Aero drops decals "
     "AND lights; LiteSpeed+ drops decals AND reflectors. Alternative reading: they just didn't optimise "
     "brand #2. We have no component price list — read the costs on the Design Brand screen."],
]
r2 = block(r, rws, wrap=(5,))
for i in (2, 4):
    ws.cell(r + i, 4).fill = yellow
ws.cell(r + 5, 5).fill = good
ws.cell(r + 5, 5).font = bold
ws.cell(r + 7, 5).fill = yellow
r = r2

rws = [
    ["WHY NOT IN Q4 — the capacity arithmetic", "Units", "Note"],
    ["Q4 plan of record demand forecast", 1171, "Rio filled to cap + full web rebuild"],
    ["Capacity: 24/day x 74% x 65 days", 1154, "100% of what we own, zero overtime"],
    ["Slack before adding anything", -17, "we are ALREADY 1.4% over — there is no room"],
    ["A second Speed brand would add", 363, "74–81% of the lead brand, per BB LLC and LiteCycle"],
    ["Resulting unfilled demand", 380, "25% stock-out rate"],
    ["Resulting Q5 ill will", "~12%", "half the unmet percentage"],
    ["Q3 for comparison", 204, "33% stock-out, 16.3% ill will, LAST place in the industry"],
    ["Required operating capacity", "31.9/day", "needs a 4th printer (~$240,000) to serve it"],
    ["VERDICT", None, "Adding a brand in Q4 recreates the exact failure we are fixing. It is not a brand "
     "decision, it is a printer decision — and the printer money is better spent on Rio staff and the web "
     "rebuild, which need no capex at all."],
]
r2 = block(r, rws, wrap=(3,))
ws.cell(r + 3, 2).fill = bad
ws.cell(r + 5, 2).fill = bad
ws.cell(r + 6, 2).fill = bad
ws.cell(r + 10, 3).fill = bad
ws.cell(r + 10, 1).font = Font(bold=True)
r = r2

rws = [
    ["BRAND 3 — second SPEED brand (first to add, once capacity exists)", "Value"],
    ["Why Speed", "Largest segment at 2,878 vs Mountain's 1,621, we hold only 7.5%, and it is the least "
     "price-sensitive segment in the game — judgment holds at 100 up to $1,580 and resistance only starts "
     "at $1,749. Contribution at $1,580 is $886/unit, our best."],
    ["Frame", "Aerodynamic"],
    ["Tires", "Racing - fast"],
    ["Brakes", "Precision"],
    ["Handlebars", "Basic drop down"],
    ["Seat", "Polymer gel racing"],
    ["Gears", "14 speed (2x7)"],
    ["Decals", "Colorful thin brushstrokes — INCLUDE"],
    ["Lights", "Standard — INCLUDE"],
    ["Reflectors", "OMIT (this is what makes it 76 not 77)"],
    ["Carrier / suspension", "none"],
    ["Judgment", "76 — copies Blu Tube Ballz, The Armstrong AND Skim MILC, three firms that independently "
     "converged on this exact spec. Add reflectors for 77 only if the screen shows they are cheap."],
    ["Price", "$1,580 — identical to Swift Bike, exactly as BB LLC prices both of theirs"],
    ["Advertising", "NONE. BB LLC leaves both second brands unadvertised and leads the industry. Keep media "
     "concentrated on the two ads we already have."],
    ["Trigger to build it", "Owned capacity comfortably exceeds forecast demand AND both existing brands are "
     "fully supplied. Realistically Q5, after Q4 proves the supply fix."],
]
r2 = block(r, rws, wrap=(2,))
for i in (8, 9, 10, 12, 13, 14):
    ws.cell(r + i, 2).fill = good
ws.cell(r + 15, 2).fill = yellow
r = r2

rws = [
    ["BRAND 4 — second MOUNTAIN brand (lower priority)", "Value"],
    ["Why later", "Mountain is the smallest and slowest-growing segment (1,621) and we are already #2 at "
     "22.1%. Speed has four times the headroom."],
    ["Spec", "The exact Hike Bike recipe but with the polymer gel COMFORT seat instead of all-purpose "
     "(−3 → judgment 70). This is Blu Tail Ballz and LiteTrail Pro."],
    ["Price", "$1,365 — identical to Hike Bike, as BB LLC does"],
    ["Advertising", "None"],
]
r2 = block(r, rws, wrap=(2,))
r = r2

rws = [
    ["BRAND 5 — RECREATION (Q5+ at the earliest; fights our margin strategy)", "Value"],
    ["Spec = MountainCruise1, the only brand at the Recreation ceiling", "judgment 76"],
    ["Frame", "Comfort (relaxed)"],
    ["Tires", "Hybrid - road and off-road"],
    ["Brakes", "STANDARD DISC — not precision, not standard"],
    ["Handlebars", "Comfort straight"],
    ["Gears", "7 speed (1x7)"],
    ["Seat", "Polymer gel comfort"],
    ["Accessories", "Reflectors + decals + lights + plastic basket + front shocks — all of them"],
    ["THE PROBLEM", "Every Recreation winner sits at price judgment 100, which means cheapest in the game: "
     "Whole MILC $950 net, Spoke'd Easy $999, Mars Rover $1,050 net. Our $1,365 scores only 81 in "
     "Recreation. Entering means accepting low margin, which contradicts our locked profit-margin-leader "
     "intent. Hike Bike already scores 56 with Recreation buyers and pulled 54 units of pure spillover, so "
     "there is no cheap way in."],
    ["Caveat on the spec", "Only 4 Recreation brands exist and they differ in 3 places, so individual "
     "penalties are NOT uniquely identifiable. MountainCruise1's exact recipe is known to score 76; the "
     "reason each component matters is not."],
]
r2 = block(r, rws, wrap=(2,))
ws.cell(r + 9, 2).fill = bad
ws.cell(r + 10, 2).fill = yellow
r = r2

rws = [
    ["WHAT WE DO INSTEAD IN Q4", "Why it beats a third brand"],
    ["Fill Rio to the 7-person cap", "~+180 store units at Rio's measured 60/head. No design fee, no lease, "
     "no printer."],
    ["Rebuild the web (4 tactics + 7 staff)", "120 → 435+ units. Every firm funding all four lands 435–569. "
     "No printer."],
    ["Swift Bike to 77 and $1,580", "Raises share AND margin in the segment where a third brand would go — "
     "without adding a unit of demand we cannot build."],
    ["Restore Biking inserts to 12 + 12", "Five more inserts than Q3 for $8,000 LESS."],
    ["Summary", "Every one of these lifts the same segments a third brand would target, costs less, and does "
     "not create demand we would fail to serve. A third brand is a Q5 move that follows a printer."],
]
block(r, rws, wrap=(2,))

ws.column_dimensions["A"].width = 52
ws.column_dimensions["B"].width = 22
ws.column_dimensions["C"].width = 20
ws.column_dimensions["D"].width = 24
ws.column_dimensions["E"].width = 76
ws.column_dimensions["F"].width = 18
ws.column_dimensions["G"].width = 14

wb.save(DST)
print("\nAdded Q4_Brand_Count to Q3Data.xlsx")
