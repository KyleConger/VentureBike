"""Q4 World Market: retail price, rebate and sales priority."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DST = "quarters/Q3/Q3Data.xlsx"

# Price tiers from Q3_Price_Judgment. Identical component triplets => identical prices,
# so the tier ladder ranks every rival's price and tells us the judgment at each price point.
TIERS = [
    (1_365, 81, 100, 100, "Hike Bike, Blu Ruged Ballz, Blu Tail Ballz"),
    (1_450, 76, 94, 100, "Swift Bike, Skim MILC MKII"),
    (1_499, 73, 91, 100, "TERRAMAX, Spoke'd Speed"),
    (1_515, 73, 91, 100, "LiteSpeed Pro+"),
    (1_579, 70, 86, 100, "Mach 0.6"),
    (1_580, 70, 86, 100, "Blu Aero Ballz, Blu Tube Ballz, The Armstrong"),
    (1_749, 63, 78, 90, "MACH I.I — most expensive in the game"),
]
MTN_DEMAND, SPD_DEMAND = 1_621, 2_878
CAPACITY = 1_154
FORECAST = 1_171

print("=" * 98)
print("PART 1 — RETAIL PRICE: the tier ladder answers both brands outright")
print("=" * 98)
print(f"  {'Price':>8}{'Rec':>6}{'Mtn':>6}{'Speed':>7}   Brands at this price")
for p, rec, mtn, spd, who in TIERS:
    mark = ""
    if p == 1_365:
        mark = "  <== HIKE BIKE. Mountain 100, and the HIGHEST price that still scores 100."
    elif p == 1_450:
        mark = "  <== SWIFT BIKE today. Speed 100, but we are leaving money on the table."
    elif p == 1_580:
        mark = "  <== SWIFT BIKE TARGET. Speed STILL 100, proven by 3 rival brands."
    print(f"  {p:>8,}{rec:>6}{mtn:>6}{spd:>7}   {who}{mark}")

print(f"""
  HIKE BIKE -> HOLD $1,365. Mountain price judgment is 100 here and the very next step up
  ($1,450) drops it to 94. $1,365 is the highest price in the game that draws ZERO Mountain
  price resistance, so it is already optimal. BB LLC's Blu Ruged Ballz confirms it from the
  outside: identical components (56/73/1), identical brand judgment (73), identical price
  ($1,365), identical price judgment (81/100/100).

  SWIFT BIKE -> RAISE $1,450 to $1,580, a gain of $130 per unit.
    Speed is by far the least price-sensitive segment: judgment stays at 100 all the way down
    the ladder to $1,580 and only breaks at $1,749 (MACH I.I, 90). Three rival brands sit at
    exactly $1,580 with no resistance — Blu Aero Ballz, Blu Tube Ballz and The Armstrong — plus
    Mach 0.6 at $1,579 on a brand judgment of only 61.
    Today Swift Bike is the 2nd CHEAPEST of 10 Speed brands while carrying the 2nd WORST
    product, $97 below the Speed median of $1,547. That is the wrong end of both scales.
    Rivals price Speed $200+ above Mountain. BB LLC runs Mountain $1,365 (identical to ours)
    and Speed $1,580. We copied their Mountain price and missed their Speed price.""")

print("\n" + "=" * 98)
print("PART 2 — DO NOT GAMBLE INTO THE UNTESTED $1,581-$1,748 ZONE")
print("=" * 98)
print(f"""  Nobody in the industry prices between $1,580 and $1,749, so we cannot see where Speed
  judgment starts to break. Two reasons to stop at the proven $1,580 this quarter:

  1. Capacity binds. Forecast demand {FORECAST:,} against capacity {CAPACITY:,} leaves only
     {FORECAST-CAPACITY} units of slack. A price rise that dents Speed demand could leave the
     printers under-filled — and Q2 already cost us $148,652 for scheduled capacity we did not
     use. Not worth the experiment in a quarter this tight.
  2. It is a one-way door within the quarter. $1,580 is verified by four separate brands and
     captures $130 of the available $384 headroom with zero risk.

  The right time to probe higher is a quarter when demand comfortably EXCEEDS capacity. Note
  that Q5 is not it either: two new printers take capacity to ~40/day, so Q5 needs MORE demand,
  not less.""")

print("\n" + "=" * 98)
print("PART 3 — PRICE REBATE: keep both at $0")
print("=" * 98)
print("""  Only two brands in the entire industry use a rebate, and both are Recreation:
    Mars Rover      $1,100 - $50  = $1,050 net   (SpaceBikes)
    Whole MILC MKII $1,050 - $100 =   $950 net   (MILC Bikes)
  ZERO Mountain brands and ZERO Speed brands rebate.

  A rebate is just a price cut that prints as its own line. It would walk us back DOWN the tier
  ladder we are trying to climb, and it signals discount in the two segments where we are
  pursuing margin leadership. Recreation is the only segment where price is the battleground
  (its winners are simply the three cheapest bikes) and Recreation stays parked.""")

print("\n" + "=" * 98)
print("PART 4 — SALES PRIORITY: the field is a genuine trade-off, and one side is arithmetic")
print("=" * 98)
mtn_pts = 100 / MTN_DEMAND
spd_pts = 100 / SPD_DEMAND
print(f"""  Market Performance = (avg share across targeted segments) x (pct of demand served).
  Verified on Q3: Mountain 358/{MTN_DEMAND:,} = 22.1%, Speed 215/{SPD_DEMAND:,} = 7.5%,
  average 14.8% x 67.5% served = 0.100 against a reported 0.098.

  THE ARITHMETIC NOBODY WOULD GUESS: because share is units divided by SEGMENT demand, and
  Mountain is the SMALLER segment, one Mountain unit buys more share than one Speed unit.
    one Mountain unit = {mtn_pts:.4f} share points   (1 / {MTN_DEMAND:,})
    one Speed unit    = {spd_pts:.4f} share points   (1 / {SPD_DEMAND:,})
    -> a Mountain unit is worth {mtn_pts/spd_pts:.2f}x a Speed unit for Market Performance.

  AGAINST THAT, MARGIN POINTS THE OTHER WAY:
    Hike Bike contribution  $747/unit  ($1,365 - $618 COGS)
    Swift Bike at $1,580    $886/unit  ($1,580 - $694 COGS)
    -> Swift earns $139/unit MORE, and the 14-speed swap should widen it further since the
       feature text says fewer parts than the 24-speed it replaces.

  So the two objectives genuinely conflict: share says promote Hike Bike, margin says Swift.""")

print("\n  Rationing scenarios (priority only matters for units we CANNOT build):\n")
print(f"  {'Shortfall':>10}{'Hike first: lost $':>20}{'Swift first: lost $':>21}{'Avg share cost, Hike 1st':>26}{'Avg share cost, Swift 1st':>27}")
for short in (17, 100, 250):
    a_cash, b_cash = short * 886, short * 747
    a_share = (short / SPD_DEMAND) * 100 / 2
    b_share = (short / MTN_DEMAND) * 100 / 2
    print(f"  {short:>10}{a_cash:>20,}{b_cash:>21,}{a_share:>25.2f}pt{b_share:>26.2f}pt")
print(f"""
  READ: at the PLANNED shortfall of {FORECAST-CAPACITY} units the whole question is worth about
  ${(FORECAST-CAPACITY)*139:,} — immaterial. Priority only becomes consequential if demand
  overshoots the forecast, which it plausibly will once the ad rebuilds, the web rebuild, Rio
  staffing and the Biking media concentration all land in the same quarter.""")

print("\n" + "=" * 98)
print("PART 5 — BUT Q3 SAYS THE PRIORITY FIELD DID NOT ACTUALLY PROTECT ANYTHING")
print("=" * 98)
print("""  We set Hike Bike to priority 1 in Q3. The outcome:
    Hike Bike  stocked out 32.5%
    Swift Bike stocked out 32.6%
  Essentially identical. If priority rationed production toward priority 1, Hike Bike should
  have been filled FIRST and shown a far lower stock-out rate until Swift ran dry. It did not.

  So on our own evidence the shortfall was allocated roughly PRO-RATA and the priority field
  did not shield the favoured brand. The mechanism is UNVERIFIED.

  TWO CONSEQUENCES:
  1. Do not treat priority as insurance. The real protection against stock-outs is SCHEDULING
     PRODUCTION TO THE FORECAST — schedule the full 24/day with zero overtime, which is
     already the plan of record.
  2. Because the field's effect is unproven and the planned shortfall is trivial, set it on
     strategic grounds rather than trying to optimise a mechanism we cannot see.""")

print("\n" + "=" * 98)
print("PART 6 — RECOMMENDATION: LEAVE PRIORITY AS IS (Hike Bike 1, Swift Bike 2)")
print("=" * 98)
print(f"""  Four reasons, in order of how much I trust them:

  1. ARITHMETIC. A Mountain unit is worth {mtn_pts/spd_pts:.2f}x a Speed unit for Market Performance,
     because Mountain's segment demand ({MTN_DEMAND:,}) is smaller than Speed's ({SPD_DEMAND:,}).
     This is division, not estimation.
  2. DEFENSIBILITY. Mountain is the one position we can actually hold: 22.1% share, a product
     TIED for best in the industry at 73, and a Rio store monopoly delivering 58% of Rio
     Mountain. In Speed we are 6th at 7.5% with the 2nd-worst product. Protect the franchise.
  3. RIO. Q3's ill will hit hardest exactly where we lead. Stocking out Mountain a second
     consecutive quarter would damage the single most profitable position we own.
  4. TEAM GOAL. The stated segment priority is Mountain -> Speed -> Recreation.

  WHY I AM NOT PERSUADED BY THE MARGIN ARGUMENT: it looks strong through Financial Performance,
  but that metric's numerator is currently only ~$132,975, so ANY marginal dollar swings it by a
  huge percentage. That hypersensitivity is an artefact of us being barely profitable, and it
  breaks down entirely if Q4 posts a loss — which it likely will once R&D is expensed. The share
  arithmetic is stable; the FP arithmetic is not. Prefer the stable one.

  AND NOTE THE MARGIN GAIN IS NOT LOST EITHER WAY: raising Swift to $1,580 captures the $130
  per unit on every Swift bike we DO sell. Priority only decides who loses the last few units.""")

print("\n" + "=" * 98)
print("PART 7 — TWO THINGS TO CHECK ON THE SCREEN")
print("=" * 98)
print("""  1. 'AVAILABLE FOR SALE' IS BLANK for both brands. Make sure both are actually enabled for
     World Market. We are spending ~$33,000 on web productivity tactics and staffing the web
     centre from 3 people to 7 this quarter — if a brand is not flagged available in the web
     channel, that entire investment cannot convert. Worth a deliberate look rather than an
     assumption, since a blank field is exactly how a brand silently fails to list.

  2. IS THIS SCREEN GLOBAL OR PER-MARKET? It is headed 'World Market', which is our web centre.
     Check whether Amsterdam and Rio have their own price screens. Q3 reported ONE price per
     brand across the industry, which suggests pricing is global — but if it is per-market, the
     $1,580 Swift price and the $1,365 Hike price must be entered in all three places, and a
     mismatch would be an easy and expensive mistake.""")

print("\n" + "=" * 98)
print("PART 8 — WHAT TO ENTER")
print("=" * 98)
print(f"""  {'Priority':<10}{'Brand':<14}{'Retail price':>14}{'Rebate':>9}   Change
  {'1':<10}{'Hike Bike':<14}{'$1,365':>14}{'$0':>9}   NO CHANGE — already optimal
  {'2':<10}{'Swift Bike':<14}{'$1,580':>14}{'$0':>9}   RAISE $130 from $1,450

  Do this in the SAME quarter as the design fix (24-speed -> 14-speed, add decals => 77). A
  price rise on the 2nd-worst Speed product invites resistance; a price rise on a joint-best
  product does not. The two decisions belong together.""")
print("=" * 98)

# ---------------------------------------------------------------- workbook
wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
good = PatternFill("solid", fgColor="C6EFCE")
badf = PatternFill("solid", fgColor="FFC7CE")
yel = PatternFill("solid", fgColor="FFEB9C")
thin = Border(*[Side(style="thin", color="D0D0D0")] * 4)

if "Q4_Price_Priority" in wb.sheetnames:
    del wb["Q4_Price_Priority"]
ws = wb.create_sheet("Q4_Price_Priority")


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


ws["A1"] = "Q4 World Market — retail price, rebate, sales priority"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("BOTTOM LINE: hold Hike Bike at $1,365 (already the highest price that draws zero Mountain resistance), "
            "raise Swift Bike $1,450 -> $1,580 (+$130, verified riskless by three rival brands at that exact price), "
            "keep both rebates at $0, and LEAVE the priority order alone. The priority reasoning changed: a Mountain "
            "unit is worth 1.78x a Speed unit for Market Performance because Mountain is the SMALLER segment.")
ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A2:F2")
ws.row_dimensions[2].height = 48

r = 4
rws = [["WHAT TO ENTER", "Priority", "Retail price", "Rebate", "Change"]]
rws.append(["Hike Bike", 1, 1365, 0, "NO CHANGE — Mountain price judgment is already 100 and the next step up drops it to 94"])
rws.append(["Swift Bike", 2, 1580, 0, "RAISE $130 from $1,450 — Speed judgment stays at 100, proven by Blu Aero, Blu Tube and The Armstrong"])
r2 = block(r, rws, wrap=(5,))
for i in (1, 2):
    ws.cell(r + i, 3).number_format = "$#,##0"
    ws.cell(r + i, 4).number_format = "$#,##0"
for c in range(1, 6):
    ws.cell(r + 2, c).fill = good
r = r2

rws = [["PRICE TIER LADDER (from Q3_Price_Judgment)", "Price", "Rec", "Mtn", "Speed", "Brands at this price"]]
for p, rec, mtn, spd, who in TIERS:
    rws.append(["", p, rec, mtn, spd, who])
r2 = block(r, rws, wrap=(6,))
for i in range(1, len(rws)):
    ws.cell(r + i, 2).number_format = "$#,##0"
    if rws[i][1] == 1_365:
        for c in range(1, 7):
            ws.cell(r + i, c).fill = good
    if rws[i][1] == 1_580:
        for c in range(1, 7):
            ws.cell(r + i, c).fill = yel
r = r2
ws.cell(r - 1, 1, "Speed is the least price-sensitive segment in the game: judgment holds at 100 all the way to "
                  "$1,580 and only breaks at $1,749. Mountain holds 100 only to $1,365. Swift Bike is currently the "
                  "2nd CHEAPEST of 10 Speed brands while carrying the 2nd WORST product — $97 below the $1,547 Speed "
                  "median. Rivals price Speed $200+ above Mountain; BB LLC runs Mountain $1,365 (same as ours) and "
                  "Speed $1,580. We copied their Mountain price and missed their Speed price.")
ws.cell(r - 1, 1).alignment = Alignment(wrap_text=True)

rws = [["WHY NOT GO ABOVE $1,580?", "Reason"]]
for a, b in [
    ("$1,581-$1,748 is untested", "Nobody in the industry prices in that band, so we cannot see where Speed judgment breaks."),
    ("Capacity binds this quarter", f"Forecast {FORECAST:,} vs capacity {CAPACITY:,} leaves only {FORECAST-CAPACITY} units of slack. "
     "A price rise that dents Speed demand could leave printers under-filled, and Q2 already cost us $148,652 for scheduled capacity we did not use."),
    ("$1,580 is free money; beyond is a bet", "It is verified by four separate brands and captures $130 of the $384 available headroom at zero risk."),
    ("Q5 is not the time either", "Two new printers take capacity to ~40/day, so Q5 needs MORE demand, not less. Probe higher only when demand comfortably exceeds capacity."),
]:
    rws.append([a, b])
r2 = block(r, rws, wrap=(2,))
r = r2

rws = [["REBATE — keep both at $0", "Evidence"]]
for a, b in [
    ("Only 2 brands in the game rebate", "Mars Rover $1,100-$50 and Whole MILC $1,050-$100. Both Recreation."),
    ("Zero Mountain, zero Speed rebates", "Not one rival in either of our segments discounts."),
    ("A rebate is a tier walk-back", "It is a price cut that prints as its own line — it moves us back DOWN the ladder we are climbing and signals discount in the two segments where we chase margin leadership."),
    ("Recreation is the price battleground", "Its three winners are simply the cheapest bikes in the game. Recreation stays parked."),
]:
    rws.append([a, b])
r2 = block(r, rws, wrap=(2,))
r = r2

rws = [["SALES PRIORITY — the trade-off", "Detail"]]
for a, b in [
    ("SHARE says promote Hike Bike", f"Market Performance = avg share x pct served, and share is units / SEGMENT demand. "
     f"Mountain demand is {MTN_DEMAND:,} vs Speed {SPD_DEMAND:,}, so one Mountain unit = {mtn_pts:.4f} share points against "
     f"{spd_pts:.4f} for a Speed unit — a Mountain unit is worth {mtn_pts/spd_pts:.2f}x. This is division, not estimation."),
    ("MARGIN says promote Swift Bike", "Hike Bike contributes $747/unit; Swift at $1,580 contributes $886 — $139 more, and "
     "the 14-speed swap should widen it since the feature text says fewer parts than the 24-speed it replaces."),
    ("But Q3 says the field did nothing", "We set Hike Bike to priority 1 in Q3 and it stocked out 32.5% while Swift stocked out "
     "32.6%. If priority rationed production, Hike Bike should have filled FIRST. It did not — the shortfall went roughly "
     "PRO-RATA. The mechanism is UNVERIFIED; do not treat priority as insurance."),
    ("And the stake is tiny at plan", f"At the planned {FORECAST-CAPACITY}-unit shortfall the whole question is worth about "
     f"${(FORECAST-CAPACITY)*139:,}. It only matters if demand overshoots — which it plausibly will once the ad rebuilds, web "
     "rebuild, Rio staffing and Biking concentration all land together."),
    ("DECISION: leave it (Hike 1, Swift 2)", "Mountain is the one position we can hold — 22.1% share, a product TIED for best at "
     "73, and a Rio monopoly delivering 58% of Rio Mountain, versus 6th place and the 2nd-worst product in Speed. The margin "
     "case looks strong only through Financial Performance, whose numerator is ~$132,975, so any marginal dollar swings it "
     "wildly — an artefact of being barely profitable that breaks entirely if Q4 posts a loss. Prefer the stable arithmetic. "
     "And the $130 price rise is captured on every Swift bike we DO sell either way; priority only decides the last few units."),
]:
    rws.append([a, b])
r2 = block(r, rws, wrap=(2,))
ws.cell(r + 3, 2).fill = yel
for c in (1, 2):
    ws.cell(r + 5, c).fill = good
r = r2

rws = [["CHECK ON THE SCREEN", "Why"]]
for a, b in [
    ("'Available for Sale' is BLANK for both", "Confirm both brands are actually enabled for World Market. We are spending ~$33,000 "
     "on web tactics and taking web staff from 3 to 7 this quarter — if a brand is not flagged available in the web channel, none "
     "of that can convert. A blank field is exactly how a brand silently fails to list."),
    ("Is this screen global or per-market?", "It is headed 'World Market', our web centre. Check whether Amsterdam and Rio have "
     "their own price screens. Q3 reported ONE price per brand industry-wide, suggesting pricing is global — but if it is "
     "per-market, both prices must be entered in all three places."),
    ("Sequencing", "Do the price rise in the SAME quarter as the design fix (24sp -> 14sp, add decals => 77). A price rise on the "
     "2nd-worst Speed product invites resistance; the same rise on a joint-best product does not."),
]:
    rws.append([a, b])
block(r, rws, wrap=(2,))
ws.cell(r + 1, 2).fill = yel

ws.column_dimensions["A"].width = 42
ws.column_dimensions["B"].width = 14
ws.column_dimensions["C"].width = 100
for c in "DEF":
    ws.column_dimensions[c].width = 13
ws.column_dimensions["F"].width = 46

wb.save(DST)
print("\nAdded Q4_Price_Priority to Q3Data.xlsx")
