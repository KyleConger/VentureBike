"""Q4 sales force and stores: exact hire-by-hire roster, and why no new store this quarter."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DST = "quarters/Q3/Q3Data.xlsx"

CAP_CITY = 7
PER_PERSON_Q = 6_851          # BB LLC package $27,402/yr / 4
TRAIN = 400
CAP_Q4, FCST_Q4 = 1_154, 1_171
CAP_Q5 = round(40 * 0.74 * 65)

# city -> (service, rec, mountain, speed, untrained)
NOW = {"Amsterdam": (1, 0, 3, 2, 1), "Rio de Janeiro": (1, 0, 1, 1, 1)}
BB = {"Amsterdam": (1, 0, 3, 3, 0), "New York City": (1, 0, 3, 3, 0)}

print("=" * 98)
print("PART 1 — STORES: NO NEW STORE IN Q4. THE PRINTERS DECIDE THIS.")
print("=" * 98)
print(f"""  New York is unambiguously the right third city, and it is still unclaimed:
    NYC demand      2,039 units (31.1% of the market) — only 3 of 7 firms have a store there
    Bangalore         761 units — LiteCycle alone
    Nobody in the industry has THREE stores. Five firms have two, two firms have one.
    Store effect is roughly 10x: no store = 1.9-3.2% share · store against 6 rivals = 9.9% ·
    SOLE store in a city = 34.9% (us in Rio) to 48.6% (LiteCycle in Bangalore).

  BUT IT CANNOT OPEN THIS QUARTER, and the reason is the printer lag:
    Q4 capacity          {CAP_Q4:,} units  (24/day x 0.74 x 65)
    Q4 forecast demand   {FCST_Q4:,} units  — ALREADY includes Rio to the cap and the full web rebuild
    Spare capacity          {CAP_Q4-FCST_Q4:>3} units
    An NYC store adds 340-570 units against that. It would take us straight back to the Q3
    failure: 33% stocked out, 16.3% ill will carried into the next quarter — and it would happen
    in the one city where BB LLC already holds 82.8% of Mountain.

  THE SEQUENCE IS ALREADY SET BY THE FINANCING DECISION:
    Q4  buy printers #4 and #5 ($480,000, funded by the loan). Capacity stays {CAP_Q4:,}.
    Q5  printers online -> {CAP_Q5:,} units. OPEN NEW YORK into capacity that exists, and enter
        via SPEED (NYC Speed is 880 units with no dominant firm — the best is Armstrong at 272)
        rather than Mountain, on the back of the Swift Bike fix to 77.""")

print("\n" + "=" * 98)
print("PART 2 — AMSTERDAM IS AT THE HARD CAP. 'DEEPEN OUR BEST CITY' IS NOT A LEGAL MOVE.")
print("=" * 98)
print(f"""  {CAP_CITY} people per city is the ceiling. Not one firm in any of the 12 city rows exceeds it,
  and BB LLC, SpaceBikes, Bike Bros, MILC and our own Amsterdam all sit exactly at {CAP_CITY}.

    WeBike Amsterdam   {CAP_CITY} people  -> AT THE CAP. We cannot add a single person.
    WeBike Rio         4 people  -> 3 SLOTS OPEN. The only store depth available to us.
    WeBike web centre  3 people  -> 4 SLOTS OPEN. Same 7-person ceiling as a store.

  So there are only three ways to add selling capacity at all: fill Rio, fill the web centre, or
  open a new city. Two of those are free of the store cap and both are on the table this quarter.

  AND A RIO HEAD IS WORTH MORE THAN AN AMSTERDAM HEAD:
    Rio        270 city demand / 4 people, with ZERO rival stores in the city
    Amsterdam  295 city demand / 7 people, against SIX rival stores
  Amsterdam is the most contested city in the game — all 7 firms are there, which is exactly why
  nobody exceeds 16.5% share in it. Rio we own outright.""")

print("\n" + "=" * 98)
print("PART 3 — THE ROSTER: WE CAN CLONE BB LLC'S TEMPLATE EXACTLY")
print("=" * 98)
print("""  BB LLC runs the IDENTICAL template in both of its cities:
    1 Service · 0 Recreation · 3 Mountain · 3 Speed = 7
  That is not a coincidence of two data points — it is the same allocation twice, in the two
  biggest cities, by the firm with a quarterly BSC of 22 against our 0.
""")
rows = []
for city, (sv, rc, mt, sp, un) in NOW.items():
    tgt = (1, 0, 3, 3, 0)
    rows.append((city, sv, rc, mt, sp, un, tgt))
print(f"  {'City':<17}{'Role':<12}{'Now':>5}{'Target':>8}{'Change':>8}   How")
for city, sv, rc, mt, sp, un, tgt in rows:
    cur = {"Service": sv, "Recreation": rc, "Mountain": mt, "Speed": sp}
    t = {"Service": tgt[0], "Recreation": tgt[1], "Mountain": tgt[2], "Speed": tgt[3]}
    for role in ("Service", "Recreation", "Mountain", "Speed"):
        d = t[role] - cur[role]
        how = ""
        if city == "Amsterdam" and role == "Speed" and d > 0:
            how = "TRAIN the 1 untrained -> Speed ($400). NO HIRE — city is at cap."
        elif city.startswith("Rio") and role == "Mountain" and d > 0:
            how = "TRAIN the 1 untrained -> Mountain ($400), then hire +1"
        elif city.startswith("Rio") and role == "Speed" and d > 0:
            how = f"HIRE +{d}"
        elif d == 0:
            how = "no change"
        print(f"  {city:<17}{role:<12}{cur[role]:>5}{t[role]:>8}{d:>+8}   {how}")
    print(f"  {city:<17}{'TOTAL':<12}{sv+rc+mt+sp+un:>5}{sum(tgt):>8}"
          f"{sum(tgt)-(sv+rc+mt+sp+un):>+8}   {'AT CAP already' if city=='Amsterdam' else 'fills the 3 open slots'}")
    print()

print("  WEB CENTRE (same 7-person cap, and we are nowhere near it):")
print(f"    {'Role':<12}{'Now':>5}{'Target':>8}{'Change':>8}   BB LLC runs 5 sales + 2 service = 7")
for role, now, tgt in (("Sales", 2, 5), ("Service", 1, 2)):
    print(f"    {role:<12}{now:>5}{tgt:>8}{tgt-now:>+8}")
print(f"    {'TOTAL':<12}{3:>5}{7:>8}{4:>+8}")

print("\n" + "=" * 98)
print("PART 4 — WHAT THAT ADDS UP TO")
print("=" * 98)
SPEC_NOW = {"Service": 2, "Recreation": 0, "Mountain": 4, "Speed": 3}
SPEC_TGT = {"Service": 2, "Recreation": 0, "Mountain": 6, "Speed": 6}
BB_SPEC = {"Service": 2, "Recreation": 0, "Mountain": 6, "Speed": 6}
print(f"  {'Specialists':<14}{'WeBike now':>12}{'WeBike Q4':>11}{'BB LLC':>9}   Note")
for k in SPEC_NOW:
    note = ""
    if k == "Speed":
        note = "DOUBLED — largest segment (2,878), and we had HALF the leaders' coverage"
    elif k == "Mountain":
        note = "matches BB LLC, who took 46.3% Mountain share with these 6"
    elif k == "Recreation":
        note = "stay at 0 — BB LLC also runs 0. Recreation is parked."
    print(f"  {k:<14}{SPEC_NOW[k]:>12}{SPEC_TGT[k]:>11}{BB_SPEC[k]:>9}   {note}")
print(f"\n  {'Store people':<14}{11:>12}{14:>11}{14:>9}   exact match")
print(f"  {'Web people':<14}{3:>12}{7:>11}{7:>9}   exact match")
print(f"  {'TOTAL':<14}{14:>12}{21:>11}{21:>9}   exact match")

print(f"""
  SPEED IS THE HEADLINE. We carry 3 Speed specialists against 6 for BB LLC, SpaceBikes AND
  LiteCycle — in the LARGEST segment in the game (2,878 units) while holding 7.5% of it. And we
  are fixing the Speed product (72 -> 77) and the Speed ad (70 -> 77) in this same quarter.
  Doubling the coverage is what lets those two fixes actually convert.""")

print("\n" + "=" * 98)
print("PART 5 — COST AND RETURN")
print("=" * 98)
ITEMS = [
    ("Train 2 untrained (1 Rio Mtn, 1 AMS Speed)", 2 * TRAIN, "free productivity — we are paying 2 full salaries for untrained sellers"),
    ("Rio: hire +3 (1 Mtn, 2 Speed)", 3 * PER_PERSON_Q, "at Rio's measured ~60-67 per head"),
    ("Web: hire +4 (3 sales, 1 service)", 4 * PER_PERSON_Q, "must pair with funding all 4 web tactics"),
]
tot = sum(i[1] for i in ITEMS)
for n, c, note in ITEMS:
    print(f"  {n:<44}{c:>10,}/qtr   {note}")
print(f"  {'-'*56}")
print(f"  {'TOTAL PEOPLE COST':<44}{tot:>10,}/qtr")
print(f"\n  Returns:")
RIO_U, WEB_U = 190, 315
print(f"    Rio +3 people   ~+{RIO_U} units x ~$800 contribution = ${RIO_U*800:>9,}  vs ${3*PER_PERSON_Q:,} cost"
      f"  => +${RIO_U*800-3*PER_PERSON_Q:,}/qtr")
print(f"    Web 3 -> 7      ~+{WEB_U} units x ~$750 = ${WEB_U*750:>9,}  less ${4*PER_PERSON_Q:,} staff"
      f" and $22,000 tactics => +${WEB_U*750-4*PER_PERSON_Q-22000:,}/qtr")
print(f"""
  Both pay back INSIDE the quarter, several times over. And the training is the cheapest thing on
  the entire Q4 board: $800 total. We are the only firm besides Bike Bros — dead last — carrying
  untrained staff at all; the other five train everyone.""")

print("\n" + "=" * 98)
print("PART 6 — WHY THIS DOES NOT CONTRADICT CUTTING MEDIA INSERTS")
print("=" * 98)
print(f"""  Fair challenge: I just argued for cutting media because Q4 demand cannot exceed {CAP_Q4:,} units.
  So why hire 7 people who exist to generate demand?

  BECAUSE THE STAFFING IS ALREADY INSIDE THE FORECAST. The {FCST_Q4:,}-unit number is not a
  do-nothing baseline — it is scenario D, which explicitly assumes Rio filled to the cap AND the
  full web rebuild. The hires ARE the plan that produces {FCST_Q4:,} against {CAP_Q4:,} of capacity.
  Take the hires away and demand falls to roughly 656, leaving us at 13.6/day against 24/day of
  owned capacity — under-scheduling, which is precisely the Q3 mistake that cost us $148,652 in
  Q2 and 204 stocked-out units in Q3.

  What pushed us OVER the line was the ad and media improvements stacked on top. That is why the
  media insert cut is the right offset: it is the discretionary, purely demand-generating item,
  while the people are the capability that produces the base plan.

  AND PEOPLE ARE DURABLE. In Q5 capacity roughly doubles to {CAP_Q5:,} units and New York unlocks 7
  MORE slots. These 7 hires need to be trained and productive BEFORE that, not during it. Same
  timing argument as the rapid R&D schedule and the social media start.""")

print("\n" + "=" * 98)
print("PART 7 — WHAT NOT TO DO")
print("=" * 98)
print("""  DO NOT cut anyone. We cut web staff from 5 to 3 in Q3 and web demand per head is now 40, the
    worst of any firm with a web centre. Bike Bros is the cautionary tale: 1 store, 0 web, 7
    people, and LAST place.
  DO NOT hire Recreation specialists. We have 0 and so does BB LLC. Rec spillover gave us 54
    units on a product that scores 56 with Rec buyers against dedicated brands at 73-76.
  DO NOT try to add anyone in Amsterdam. It is at the 7-person cap. Illegal move.
  DO NOT open Bangalore. It is 761 units against New York's 2,039, and LiteCycle already holds
    48.6% of it.
  DO NOT put new hires on the old compensation package. The plan of record moves sales to BB
    LLC's exact terms ($22,000 / Expanded / 2 weeks / 4% = $27,402), which delivered the
    industry-best 78% sales productivity. And budget 73-75% productivity, never the 85% we
    assumed in Q3.""")
print("=" * 98)

# ---------------------------------------------------------------- workbook
wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
good = PatternFill("solid", fgColor="C6EFCE")
badf = PatternFill("solid", fgColor="FFC7CE")
yel = PatternFill("solid", fgColor="FFEB9C")
thin = Border(*[Side(style="thin", color="D0D0D0")] * 4)

if "Q4_Staffing_Stores" in wb.sheetnames:
    del wb["Q4_Staffing_Stores"]
ws = wb.create_sheet("Q4_Staffing_Stores")


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


ws["A1"] = "Q4 sales force and stores — hire 7, train 2, open NO new store"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("ANSWER: YES change the people, NO change the stores. Train both untrained staff ($800), fill Rio's 3 open "
            "slots and take the web centre from 3 to 7. That clones BB LLC exactly — 14 store people, 7 web, 21 total, "
            "with 6 Mountain and 6 Speed specialists. But NO new store: printers arrive a quarter late, so Q4 capacity "
            "is 1,154 against a forecast that already assumes these hires. New York opens in Q5, into capacity that "
            "exists, entering via SPEED.")
ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A2:G2")
ws.row_dimensions[2].height = 58

r = 4
rws = [["ROSTER — city by city", "Role", "Now", "Target", "Change", "How"]]
for city, (sv, rc, mt, sp, un) in NOW.items():
    cur = {"Service": sv, "Recreation": rc, "Mountain": mt, "Speed": sp}
    t = {"Service": 1, "Recreation": 0, "Mountain": 3, "Speed": 3}
    for role in ("Service", "Recreation", "Mountain", "Speed"):
        d = t[role] - cur[role]
        if city == "Amsterdam" and role == "Speed":
            how = "TRAIN the 1 untrained -> Speed ($400). NO HIRE — city is at the 7-person cap."
        elif city.startswith("Rio") and role == "Mountain":
            how = "TRAIN the 1 untrained -> Mountain ($400), then HIRE +1"
        elif city.startswith("Rio") and role == "Speed":
            how = "HIRE +2"
        else:
            how = "no change"
        rws.append([city, role, cur[role], t[role], d, how])
    rws.append([city, "TOTAL", sv + rc + mt + sp + un, 7, 7 - (sv + rc + mt + sp + un),
                "AT CAP already — 0 hires" if city == "Amsterdam" else "fills all 3 open slots"])
for role, now, tgt in (("Sales", 2, 5), ("Service", 1, 2)):
    rws.append(["Web centre", role, now, tgt, tgt - now, "BB LLC runs 5 sales + 2 service = 7"])
rws.append(["Web centre", "TOTAL", 3, 7, 4, "4 open slots — same 7-person ceiling as a store"])
r2 = block(r, rws, wrap=(6,))
for i in range(1, len(rws)):
    if rws[i][1] == "TOTAL":
        for c in range(1, 7):
            ws.cell(r + i, c).font = bold
    if rws[i][4] and rws[i][4] > 0:
        ws.cell(r + i, 5).fill = good
r = r2

rws = [["SPECIALIST COUNT", "WeBike now", "WeBike Q4", "BB LLC", "Note"]]
for k in ("Service", "Recreation", "Mountain", "Speed"):
    note = {"Speed": "DOUBLED — the largest segment (2,878 units) where we hold only 7.5%, and where we are ALSO fixing "
                     "the product 72->77 and the ad 70->77 this quarter. Coverage is what lets those fixes convert.",
            "Mountain": "Matches BB LLC, who took 46.3% Mountain share with exactly these 6.",
            "Recreation": "Stay at 0 — BB LLC also runs 0. Recreation is parked.",
            "Service": "Already at BB LLC's level."}[k]
    rws.append([k, SPEC_NOW[k], SPEC_TGT[k], BB_SPEC[k], note])
rws += [["Store people", 11, 14, 14, "exact match"], ["Web people", 3, 7, 7, "exact match"],
        ["TOTAL PEOPLE", 14, 21, 21, "exact match"]]
r2 = block(r, rws, wrap=(5,))
for i in range(1, len(rws)):
    if rws[i][0] in ("Store people", "Web people", "TOTAL PEOPLE"):
        for c in range(1, 6):
            ws.cell(r + i, c).font = bold
ws.cell(r + 4, 5).fill = good
r = r2

rws = [["COST AND RETURN", "Cost/qtr", "Return", "Net/qtr"]]
rws.append(["Train 2 untrained (1 Rio Mountain, 1 AMS Speed)", 800,
            "free productivity — we pay 2 full salaries for untrained sellers", "cheapest item on the Q4 board"])
rws.append(["Rio: hire +3 (1 Mountain, 2 Speed)", 3 * PER_PERSON_Q,
            f"~+{RIO_U} units at Rio's measured 60-67/head", RIO_U * 800 - 3 * PER_PERSON_Q])
rws.append(["Web: hire +4 (3 sales, 1 service)", 4 * PER_PERSON_Q,
            f"~+{WEB_U} units (120 -> 435, the full-tactic cohort floor)", WEB_U * 750 - 4 * PER_PERSON_Q - 22000])
rws.append(["TOTAL", tot, "", ""])
r2 = block(r, rws, wrap=(3,))
for i in range(1, len(rws)):
    ws.cell(r + i, 2).number_format = "#,##0"
    if isinstance(rws[i][3], int):
        ws.cell(r + i, 4).number_format = "+#,##0"
        ws.cell(r + i, 4).fill = good
r = r2
ws.cell(r - 1, 1, "The web hires ONLY deliver alongside funding all four web productivity tactics ($22,000 more). "
                  "Every firm funding all four lands at 435-569 web units; both firms funding fewer land at 120-138. "
                  "Staff and tactics move together, so do not split them.")
ws.cell(r - 1, 1).alignment = Alignment(wrap_text=True)

rws = [["STORES — no change in Q4", "Detail"]]
for a, b in [
    ("New York IS the right third city", "2,039 units and 31.1% of the market against Bangalore's 761. Only 3 of 7 firms "
     "have a store there, and NOBODY in the industry has three stores — city #3 is unclaimed."),
    ("The store effect is roughly 10x", "No store = 1.9-3.2% share. A store against 6 rivals = 9.9%. SOLE store in a city "
     "= 34.9% (us in Rio) to 48.6% (LiteCycle in Bangalore). Highest-leverage decision we have."),
    ("BUT IT CANNOT OPEN IN Q4", f"Printers arrive a quarter late, so Q4 capacity is {CAP_Q4:,} units against a forecast of "
     f"{FCST_Q4:,} that ALREADY assumes these hires — {CAP_Q4-FCST_Q4} units of spare. An NYC store adds 340-570 on top. That is "
     "a direct repeat of Q3: 33% stocked out, 16.3% ill will carried forward."),
    ("Worst possible city to stock out in", "BB LLC already holds 82.8% of NYC Mountain. Arriving and immediately failing "
     "to supply hands them the city."),
    ("THE SEQUENCE", f"Q4 buy printers #4 and #5 ($480,000, loan-funded), capacity stays {CAP_Q4:,}. Q5 printers come online "
     f"at {CAP_Q5:,} units, THEN open New York — entering via SPEED (NYC Speed is 880 units with no dominant firm, best is "
     "Armstrong at 272) on the back of the Swift Bike fix to 77."),
    ("Amsterdam cannot be deepened", "It is at the 7-person hard cap. Not one firm in any of the 12 city rows exceeds 7. "
     "'Invest more in our best city' is not a legal move."),
]:
    rws.append([a, b])
r2 = block(r, rws, wrap=(2,))
ws.cell(r + 3, 2).fill = badf
ws.cell(r + 5, 2).fill = good
r = r2

rws = [["WHY THIS DOES NOT CONTRADICT CUTTING MEDIA INSERTS", "Detail"]]
for a, b in [
    ("The staffing is already INSIDE the forecast", f"The {FCST_Q4:,}-unit number is not a do-nothing baseline — it is "
     "scenario D, which explicitly assumes Rio filled to the cap AND the full web rebuild. These hires ARE the plan that "
     f"produces {FCST_Q4:,} against {CAP_Q4:,} of capacity."),
    ("Removing them causes the OPPOSITE failure", "Without the hires demand falls to roughly 656 units, or 13.6/day against "
     "24/day of owned capacity. That is under-scheduling — the Q3 mistake that followed the Q2 mistake of $148,652 in "
     "excess-capacity charges."),
    ("Ads and media were the OVERSHOOT", "The ad rebuilds and Biking concentration stack on top of the staffed plan. That is "
     "why the media insert cut is the right offset: inserts are discretionary and purely demand-generating, while people are "
     "the capability that produces the base plan."),
    ("And people are DURABLE", f"Q5 capacity roughly doubles to {CAP_Q5:,} units and New York unlocks 7 MORE slots. These 7 "
     "hires need to be trained and productive BEFORE that, not during it. Same timing argument as rapid R&D and the social "
     "media start."),
]:
    rws.append([a, b])
r2 = block(r, rws, wrap=(2,))
ws.cell(r + 4, 2).fill = good
r = r2

rws = [["WHAT NOT TO DO", "Why"]]
for a, b in [
    ("Do not cut anyone", "We cut web staff 5 -> 3 in Q3 and web demand per head is now 40, worst of any firm with a web "
     "centre. Bike Bros is the cautionary tale: 1 store, 0 web, 7 people, LAST place."),
    ("Do not hire Recreation specialists", "We run 0 and so does BB LLC. Hike Bike scores 56 with Rec buyers against dedicated "
     "Rec brands at 73-76, so spillover cannot win. Rec stays parked."),
    ("Do not try to add in Amsterdam", "At the 7-person cap. Illegal move."),
    ("Do not open Bangalore", "761 units against New York's 2,039, and LiteCycle already holds 48.6% of it."),
    ("Do not put new hires on the old package", "Move sales to BB LLC's exact terms ($22,000 / Expanded / 2 weeks / 4% = "
     "$27,402), which produced the industry-best 78% sales productivity. And budget 73-75% productivity, NEVER the 85% we "
     "assumed in Q3."),
]:
    rws.append([a, b])
block(r, rws, wrap=(2,))

ws.column_dimensions["A"].width = 46
ws.column_dimensions["B"].width = 14
ws.column_dimensions["C"].width = 100
for c in "DEF":
    ws.column_dimensions[c].width = 12
ws.column_dimensions["F"].width = 66
ws.column_dimensions["G"].width = 14

wb.save(DST)
print("\nAdded Q4_Staffing_Stores to Q3Data.xlsx")
