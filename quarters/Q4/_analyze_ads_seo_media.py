"""Q4 ad design, organic SEO and media placement — grounded in the Q3 dumps."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DST = "quarters/Q3/Q3Data.xlsx"

# ---------------------------------------------------------------- SERP data
# segment -> [(pos, company, ad, ad_judgment, clicks, total_inserts)]
SERP = {
    "Recreation": [(1, "SpaceBikes", "Mars Rover 1", 79, 303, 7), (2, "MILC", "Whole MILC MKII", 77, 144, 11),
                   (3, "Bike Bros", "Easy Rider", 75, 88, 2), (4, "Spoke'd Up", "Spoke'd Easy", 59, 61, 7)],
    "Mountain": [(1, "BB LLC", "BB Big Momma", 81, 237, 12), (2, "WeBike", "HikeBike 1", 80, 115, 11),
                 (3, "LiteCycle", "Trail Blazing 1", 79, 69, 3), (4, "Bike Bros", "TerraTech", 78, 46, 2),
                 (5, "Bike Bros", "Mountainability", 77, 33, 2)],
    "Speed": [(1, "BB LLC", "Unleash lil pap", 76, 200, 12), (2, "SpaceBikes", "The Armstrong 1", 77, 98, 10),
              (3, "LiteCycle", "Speed of Lite 1", 77, 61, 5), (4, "MILC", "Skim MILC MKII", 75, 43, 5),
              (5, "WeBike", "Swift Bike", 70, 29, 8), (6, "LiteCycle", "Speed of Lite 2", 73, 23, 4),
              (7, "Bike Bros", "AndStill", 78, 23, 2), (8, "Bike Bros", "Excluspeedity", 72, 23, 2),
              (9, "Spoke'd Up", "Spoke'd Speed", 57, 21, 5)],
}

print("=" * 100)
print("PART 1 — WHAT SETS SERP POSITION? (it is ad judgment, and Recreation PROVES it)")
print("=" * 100)
for seg, rows in SERP.items():
    js = [r[3] for r in rows]
    ins = [r[5] for r in rows]
    mono_j = all(js[i] >= js[i + 1] for i in range(len(js) - 1))
    mono_i = all(ins[i] >= ins[i + 1] for i in range(len(ins) - 1))
    print(f"\n  {seg}:")
    print(f"    {'Pos':>4}{'Ad':<22}{'Ad judgment':>13}{'Inserts':>9}{'Clicks':>8}")
    for p, co, ad, j, cl, i in rows:
        print(f"    {p:>4}{ad:<22}{j:>13}{i:>9}{cl:>8}")
    print(f"    ad judgment strictly descending? {'YES' if mono_j else 'NO'}"
          f"     insert count strictly descending? {'YES' if mono_i else 'NO'}")

print("""
  READ:
  - Mountain and Recreation rank EXACTLY by ad judgment.
  - RECREATION IS THE PROOF that judgment drives rank and media spend does not: Whole MILC ran
    11 inserts against Mars Rover's 7 and still ranked BELOW it, on a judgment of 77 vs 79. In
    Mountain the two explanations happen to agree, so only Recreation separates them.
  - SPEED DOES NOT FOLLOW AD JUDGMENT and we have not explained why. AndStill has the best Speed
    ad in the game at 78 and sits 7th. A partial clue: Bike Bros runs TWO ads for the SAME brand
    (both AndStill and Excluspeedity push MACH I.I) and both sink to 7th and 8th, consistent with
    the cannibalisation we already saw in their Mountain pages. But even excluding those two, the
    order is still not monotonic (Unleash at 76 outranks two 77s).
  - PRACTICAL CONSEQUENCE, and it differs by segment:
      Mountain -> raising ad judgment above 81 buys position 1 DETERMINISTICALLY.
      Speed    -> raising ad judgment should help but the payoff is NOT predictable. Treat any
                  Speed click gain as upside, never as a forecast.
""")

print("=" * 100)
print("PART 2 — CTR IS A STEP FUNCTION OF POSITION ONLY")
print("=" * 100)
CTR = {1: 0.228, 2: 0.111, 3: 0.067, 4: 0.046, 5: 0.032}
MTN_SEARCHES = 1027
print(f"  {'Position':>9}{'CTR':>8}{'Mountain clicks':>18}")
for p, c in CTR.items():
    print(f"  {p:>9}{c:>8.1%}{c*MTN_SEARCHES:>18.0f}")
print(f"""
  Position 1 is worth roughly TWICE position 2, and position 2 twice position 3. Snippet wording
  does not move CTR — position does. So the entire SEO game is: maximise ad judgment to win a
  better position, then let the step function pay out.

  HikeBike 1 moving from position 2 to position 1 = {CTR[1]*MTN_SEARCHES:.0f} - {CTR[2]*MTN_SEARCHES:.0f}
  = +{(CTR[1]-CTR[2])*MTN_SEARCHES:.0f} clicks, for the price of an ad redesign.""")

print("\n" + "=" * 100)
print("PART 3 — OUR SEO IS NOT THE PROBLEM. OUR WEB OPERATIONS ARE.")
print("=" * 100)
CONV = [("LiteCycle", 153, 435), ("MILC Bikes", 187, 435), ("BB LLC", 437, 569),
        ("SpaceBikes", 401, 491), ("Spoke'd Up", 82, 138), ("WeBike", 144, 120),
        ("Bike Bros", 213, 0)]
print(f"  {'Firm':<14}{'Organic clicks':>16}{'Web units':>11}{'Units per click':>17}")
for f, c, u in sorted(CONV, key=lambda x: -(x[2] / x[1] if x[1] else 0)):
    print(f"  {f:<14}{c:>16}{u:>11}{(u/c if c else 0):>17.2f}")
print("""
  LiteCycle turns 153 clicks into 435 web units. We turn 144 clicks into 120. Almost identical
  traffic, 3.6x the sales. Bike Bros generates 213 clicks and sells ZERO online because it has
  no web centre at all.

  So do NOT spend the quarter chasing clicks. The web rebuild (staff 3 -> 7, fund all four
  productivity tactics) is what converts the traffic we ALREADY have. SEO copy fixes are cheap
  and worth doing, but they are not the constraint.""")

print("\n" + "=" * 100)
print("PART 4 — HIKEBIKE 1: COPY BB LLC'S STACK ALMOST EXACTLY")
print("=" * 100)
print("""  Ours (80, position 2)          |  BB Big Momma (81, position 1)
  1 Mention brand name           |  1 Tackle steep climbs with more gears
  2 Picture of rider, steep trail|  2 Grab the path with high tread tires
  3 Tackle steep climbs w/ gears |  3 Highest rated Mountain bike
  4 Grab the path high tread     |  4 Local sales & service
  5 Have an adventure carbon bike|  5 Mention brand name
  6 Go anywhere tough carbon bike|  6 Picture of a rider on a steep trail
  7 Mountains no longer difficult|  (6 benefits total)

  THE DIFFERENCE IS EXACTLY FIVE ITEMS. BB LLC's ad = ours, minus three filler claims
  (adventure / tough carbon / mountains-no-longer-difficult), plus two real claims:

    + Highest rated Mountain bike   TRUE — Hike Bike is TIED for best Mountain at 73, with
                                    TERRAMAX and Blu Ruged Ballz. The claim is legal.
                                    Hits 'Status symbol/exclusivity' (Mtn 121).
    + Local sales & service         TRUE — Amsterdam + Rio. 3 of 4 rival Mountain ads run it.
                                    Hits 'Repair, parts, support nearby' (Mtn 121).

  Note BB LLC wins with SIX benefits while we run SEVEN. More claims is not better. Compare
  Spoke'd Speed: it has all the right Speed ingredients, uses the full 9-benefit cap, and scores
  57 — the worst ad in the game.

  RECOMMENDED HIKEBIKE 1 — build BB LLC's stack verbatim, then test ONE addition:
    1 Tackle steep climbs with more gears      (Mountain's #1 need: incline 131; 4 of 4 rivals)
    2 Grab the path with high tread tires      (rough terrain 128; 3 of 4 rivals)
    3 Highest rated Mountain bike              (NEW)
    4 Local sales & service                    (NEW)
    5 Mention brand name
    6 Picture of a rider on a steep trail
    then PREVIEW. If it reads 81, add '7 Mountains are no longer difficult' and preview again.
    Keep whichever scores higher. Revert to the Q3 copy if both come in under 80.

  WHY THIS IS THE HIGHEST-CONFIDENCE AD MOVE WE HAVE: Mountain SERP rank is exactly ad-judgment
  order, so 82 wins position 1 outright and 81 at minimum ties the leader. Everything about the
  mechanism is already verified.""")
gain = (CTR[1] - CTR[2]) * MTN_SEARCHES
print(f"""
  VALUE: +{gain:.0f} Mountain clicks. At BB LLC's conversion of 1.30 web units per click that is
  about {gain*1.30:.0f} units; at our current broken 0.83 it is {gain*0.83:.0f}. Against a ~$6,000 ad design
  fee, worth it on either number — but note it lands as WEB demand, which is exactly what the
  capacity warning in Part 6 is about.""")

print("\n" + "=" * 100)
print("PART 5 — SWIFT BIKE: COPY ANDSTILL, THE 78-POINT AD")
print("=" * 100)
print("""  Ours (70, 8th of 9)              |  AndStill (78, best Speed ad)
  1 Picture of road race           |  1 Picture of road race
  2 Enjoy your ride - carbon light |  2 Mention brand name
  3 Carbon fiber quality GREAT PRICE|  3 Ride a wind-cheater
  4 Mention brand name             |  4 Elite look - a ride of distinction
  5 Ride safely after dark, LIGHTS |  5 A tailor-made bike just for you! 3D printing
  6 Added safety - REFLECTORS      |  6 Roll fast with racing tires
                                   |  7 Local sales & service

  We share exactly ONE line with the best Speed ad: the road-race picture at rank 1.

  What we are missing, with rival adoption:
    Ride a wind-cheater          8 of 8 rivals run it. We do not. Aero frame supports it.
    Roll fast with racing tires  7 of 8. The bike HAS racing tires. Speed's #1 need is
                                 'Speed' at 134, and its #5 is 'Aerodynamic' at 122.
    Elite look - a ride of distinction  6 of 8, including both 77-point ads.
    3D printing                  3 of 8, and on BOTH of the two best ads. We 3D-print.
    Local sales & service        3 of 8. We are 80.9% store and never claim it.

  What we run that NOBODY runs — 0 of 8 rivals:
    Lights, reflectors, and 'great price'. These are RECREATION cues. They are why the ad leaks
    52 Recreation and 36 Mountain units and scores 70. Note the bike KEEPS its lights and
    reflectors (both are worth +1 each in the Speed design recipe) — we simply stop advertising
    them. And 'great price' has to go regardless, because we are raising Swift to $1,580.

  RECOMMENDED SWIFT BIKE — AndStill's stack verbatim:
    1 Picture of road race
    2 Mention brand name
    3 Ride a wind-cheater
    4 Elite look - a ride of distinction
    5 A tailor-made bike just for you! 3D printing
    6 Roll fast with racing tires
    7 Local sales & service
  Every one of those seven is truthful for us today.

  THEN TEST ONE SWAP: substitute 'Highest rated Speed bike' for 'Local sales & service'.
  The Armstrong runs that claim and sits at position 2. BUT: it only becomes legal once Swift
  Bike's design actually reaches 77 (tying MACH I.I and LiteSpeed Pro+), and the sim may validate
  the claim against Q3 data where Swift is still 72. So DESIGN THE BRAND FIRST, THEN THE AD, and
  if the designer rejects the claim, keep local sales & service.

  SEQUENCING THAT MATTERS: the Swift fix is TWO changes in the same quarter —
    design  24-speed -> 14-speed (+3) and ADD decals (+2)  => brand judgment 72 -> 77
    ad      the stack above                                => ad judgment 70 -> ~77
  Together those take Marketing Effectiveness to 0.7675, ahead of BB LLC's 0.765, the best in
  the industry. Ad alone gets 0.755; design alone gets 0.750. Do both or the gain is halved.

  AND DO NOT add a second Speed ad. Bike Bros ran two ads for MACH I.I and both sank to
  positions 7 and 8.""")

print("\n" + "=" * 100)
print("PART 6 — MEDIA PLACEMENT: I NEED TO REVISE MY EARLIER RECOMMENDATION")
print("=" * 100)
RATE = {"Biking Magazines": 4500, "New Venture": 5500, "Health & Fitness": 7000,
        "General News": 8000, "Business": 9500, "Sport": 10000, "Leisure & Entertainment": 10000}
print("  Cost per insert by medium:")
for m, c in sorted(RATE.items(), key=lambda x: x[1]):
    tag = "  <- CHEAPEST, and the top-preference medium for BOTH our segments (Mountain pref 134)" \
        if c == 4500 else ""
    print(f"    {m:<26}${c:>7,}{tag}")
print(f"""
  Q3 efficiency, the finding that still stands:
    BB LLC   24 inserts, 100% Biking, $108,000 = $4,500 each
    WeBike   19 inserts,  53% Biking, $116,000 = $6,105 each  -> we paid 36% MORE per insert
    Our $116,000 would have bought 25 Biking inserts. We bought 10.
    Biking is 51 of the 106 industry inserts; BB LLC alone owns 24 of those 51.
    Identical Mountain product: Blu Ruged 12 Biking -> 46.3% share. Hike Bike 6 -> 22.1%.
    We ran 12 in Q2 and cut to 6 in Q3 to 'diversify'. That $27,000 left the best medium.

  *** BUT THE CAPACITY TIMELINE CHANGES THE SIZING, AND I HAD THIS WRONG BEFORE. ***

  I previously recommended Option B: 29 inserts for $143,000, +$27,000 versus Q3. That was
  written before we established that printers arrive a quarter late. Q4 capacity is hard-capped
  at 24/day x 0.74 x 65 = 1,154 units, and the plan of record already forecasts 1,171 units of
  demand. There is NO room for more demand in Q4.

  Market Performance = (average share in targeted segments) x (percent of demand served).
  With output capped at 1,154, extra demand CANNOT raise units sold — so it only lowers the
  percent-served term. In a capacity-bound quarter, demand generation beyond capacity is
  STRICTLY NEGATIVE for the scorecard. It is also what earns next quarter's ill will.

  THE KEY SEPARATION, and it resolves the tension cleanly:
    AD JUDGMENT is free of the capacity constraint. Marketing Effectiveness =
    (avg brand judgment + avg ad judgment) / 2 — there is NO volume term. Better copy raises a
    BSC indicator at zero capacity cost.
    MEDIA INSERT VOLUME is pure demand generation and must be sized to capacity.
  => Improve the COPY aggressively. Restrain the PLACEMENT.""")

OPTS = [
    ("Q3 actual (what we did)", 11, 8, 116_000, 10, "5 media, 53% Biking. The mix error."),
    ("A. Copy BB LLC exactly", 12, 12, 108_000, 24, "Right mix, but +5 inserts of demand we cannot build."),
    ("B. My earlier rec (SUPERSEDED)", 16, 12, 143_000, 24, "+$27,000 and the most demand. Wrong for a capped quarter."),
    ("C. MIX FIX ONLY (recommended)", 12, 4, 72_000, 16, "Matches BB LLC's 12 Biking on Mountain, holds Speed at Q3's 4, drops all off-Biking. -$44,000."),
    ("D. Minimum viable", 8, 4, 54_000, 12, "If the team wants to protect the loan headroom harder."),
]
print(f"\n  {'Option':<34}{'Hike':>6}{'Swift':>7}{'Inserts':>9}{'Spend':>11}{'vs Q3':>11}{'Biking':>8}")
for name, h, s, sp, bk, note in OPTS:
    print(f"  {name:<34}{h:>6}{s:>7}{h+s:>9}{sp:>11,}{sp-116_000:>+11,}{bk:>8}")
print("""
  RECOMMENDED — OPTION C: HikeBike 1 Biking 12, Swift Bike Biking 4. 16 inserts, $72,000.
    - Puts our PRIMARY segment on BB LLC's exact Biking weight (12), on a product tied for best
      in the industry, in the one segment where SERP rank is deterministic.
    - Holds Speed at its Q3 Biking count. Swift's brand and ad fixes will lift Speed demand on
      their own this quarter; buying more Speed reach can wait for the printers.
    - Drops every off-Biking insert: Health 1, Sport 3, Business 1, New Venture 3, News 1.
    - Biking exposure goes 10 -> 16 while spend FALLS $44,000, which goes straight to loan
      headroom against the unknown rapid-R&D premium.
    - Then in Q5, with 40/day of capacity and New York opening, expand Swift to 12 Biking and
      take the full 24-insert BB LLC buy.

  DO NOT: add a third ad, buy Leisure & Entertainment (Rec-skewed, Mountain preference 60), or
  add a Recreation page. Rec stays parked.""")

print("\n" + "=" * 100)
print("PART 7 — SUMMARY OF Q4 MARKETING DECISIONS")
print("=" * 100)
print("""
  HIKEBIKE 1 AD   redesign to BB LLC's 6-benefit stack; add 'Highest rated Mountain bike' and
                  'Local sales & service'; drop adventure / tough carbon / mountains-easier.
                  Preview; target 82 for position 1. Deterministic payoff in Mountain.
  SWIFT BIKE AD   redesign to AndStill's 7-benefit stack. Drop lights, reflectors, great price.
                  Add wind-cheater, racing tires, elite look, 3D printing, local sales.
                  Design the brand FIRST so 'Highest rated Speed bike' can be tested.
  ORGANIC SEO     keep both pages at $1,000 (nobody spends more; not a lever). No third page,
                  no Recreation page. Rank matters, snippet wording does not. The real web fix
                  is operations, not traffic.
  MEDIA           Option C: Biking only, HikeBike 12 + Swift 4 = 16 inserts, $72,000, -$44,000.
                  Restrain volume because Q4 output is capped at 1,154 units.
  UNCHANGED       Hike Bike's DESIGN (73 = tied best, verified optimal) and price ($1,365).
""")
print("=" * 100)

# ---------------------------------------------------------------- workbook
wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
good = PatternFill("solid", fgColor="C6EFCE")
badf = PatternFill("solid", fgColor="FFC7CE")
yel = PatternFill("solid", fgColor="FFEB9C")
thin = Border(*[Side(style="thin", color="D0D0D0")] * 4)

if "Q4_Ads_SEO_Media" in wb.sheetnames:
    del wb["Q4_Ads_SEO_Media"]
ws = wb.create_sheet("Q4_Ads_SEO_Media")


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


ws["A1"] = "Q4 ad design, organic SEO and media placement"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("THE ORGANISING PRINCIPLE: ad JUDGMENT is free of the capacity constraint (Marketing Effectiveness has "
            "no volume term), but media INSERT VOLUME is pure demand generation and Q4 output is hard-capped at "
            "1,154 units against a forecast of 1,171. So improve the COPY aggressively and RESTRAIN the PLACEMENT. "
            "This supersedes the earlier Option B media recommendation, which was written before we established "
            "that printers arrive a quarter late.")
ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A2:G2")
ws.row_dimensions[2].height = 58

r = 4
rws = [["Segment", "Pos", "Ad", "Ad judgment", "Inserts", "Clicks", "Rank = judgment?"]]
for seg, rows in SERP.items():
    js = [x[3] for x in rows]
    mono = all(js[i] >= js[i + 1] for i in range(len(js) - 1))
    for p, co, ad, j, cl, i in rows:
        rws.append([seg, p, f"{ad} ({co})", j, i, cl, "YES" if mono else "NO"])
r2 = block(r, rws)
for i in range(1, len(rws)):
    ws.cell(r + i, 7).fill = good if rws[i][6] == "YES" else badf
    if rws[i][2].startswith(("HikeBike", "Swift")):
        for c in range(1, 8):
            ws.cell(r + i, c).font = bold
r = r2
ws.cell(r - 1, 1, "RECREATION IS THE PROOF that ad judgment sets rank and media spend does not: Whole MILC ran 11 "
                  "inserts vs Mars Rover's 7 and still ranked below it (judgment 77 vs 79). In Mountain both "
                  "explanations agree, so only Rec separates them. SPEED IS UNEXPLAINED — AndStill has the best "
                  "Speed ad at 78 and sits 7th. Partial clue: Bike Bros runs two ads for the SAME brand (MACH I.I) "
                  "and both sank to 7th/8th. But even excluding those, Speed is not monotonic. So treat Speed click "
                  "gains as UPSIDE, never as a forecast.")
ws.cell(r - 1, 1).alignment = Alignment(wrap_text=True)

rws = [["CTR by position (verified, near-identical across segments)", "CTR", "Mountain clicks at 1,027 searches"]]
for p, c in CTR.items():
    rws.append([f"Position {p}", c, round(c * MTN_SEARCHES)])
rws.append(["HikeBike 1 moving position 2 -> 1", (CTR[1] - CTR[2]), round(gain)])
r2 = block(r, rws)
for i in range(1, len(rws)):
    ws.cell(r + i, 2).number_format = "0.0%"
for c in (1, 2, 3):
    ws.cell(r + len(rws) - 1, c).fill = good
    ws.cell(r + len(rws) - 1, c).font = bold
r = r2

rws = [["HIKEBIKE 1 — recommended stack (BB LLC's verbatim)", "Status", "Why"]]
for a, b, c in [
    ("1 Tackle steep climbs with more gears", "keep, promote", "Mountain's #1 need: easily handles change in incline, 131. 4 of 4 rivals run it."),
    ("2 Grab the path with high tread tires", "keep, promote", "Can handle rough terrain 128. 3 of 4 rivals."),
    ("3 Highest rated Mountain bike", "ADD", "TRUE — Hike Bike is TIED for best Mountain at 73 with TERRAMAX and Blu Ruged Ballz, so the claim is legal. Hits Status symbol/exclusivity 121. This is BB LLC's differentiator."),
    ("4 Local sales & service", "ADD", "TRUE — Amsterdam + Rio. Hits Repair/parts/support nearby 121. 3 of 4 rival Mountain ads run it and we never have, despite being 80.9% store."),
    ("5 Mention brand name", "keep, demote", "Fills the Ad/Brand columns and does NOT consume a SERP snippet slot, so rank 1 was wasted on it."),
    ("6 Picture of a rider on a steep trail", "keep, demote", "DOES consume a snippet slot. Rec #1 and Mountain #1 both lead with a CLAIM, not a picture."),
    ("-- Have an adventure on a carbon bike!", "DROP", "Only 2 of 4 rivals. BB LLC omits it and scores higher with fewer benefits."),
    ("-- Go anywhere on a tough carbon bike", "DROP", "Weakest adoption at 1 of 4."),
    ("-- Mountains are no longer difficult", "DROP, then TEST", "Drop first. If the preview reads 81, add it back at rank 7 and preview again; keep the higher score."),
]:
    rws.append([a, b, c])
r2 = block(r, rws, wrap=(3,))
for i in range(1, len(rws)):
    v = rws[i][1]
    ws.cell(r + i, 2).fill = good if "ADD" in v else (badf if "DROP" in v else yel)
r = r2
ws.cell(r - 1, 1, "BB LLC wins with SIX benefits while we run SEVEN — more claims is not better. Spoke'd Speed has "
                  "all the right Speed ingredients, uses the full 9-benefit cap, and scores 57, the worst in the "
                  "game. PREVIEW BEFORE COMMITTING; revert to Q3 copy if it comes in under 80.")
ws.cell(r - 1, 1).alignment = Alignment(wrap_text=True)

rws = [["SWIFT BIKE — recommended stack (AndStill's verbatim, the 78)", "Status", "Why"]]
for a, b, c in [
    ("1 Picture of road race", "keep", "8 of 8 rivals. The one line we already share with the best ad."),
    ("2 Mention brand name", "promote", "8 of 8 rivals rank it 2nd. Ours sat at 4. Does not consume a snippet slot."),
    ("3 Ride a wind-cheater", "ADD", "8 of 8 rivals run it and we do not. Aerodynamic is Speed need #5 at 122, and the aero frame supports the claim."),
    ("4 Elite look - a ride of distinction", "ADD", "6 of 8, including BOTH 77-point ads. Status symbol 116."),
    ("5 A tailor-made bike just for you! 3D printing", "ADD", "3 of 8, but on BOTH of the two best ads. We actually 3D-print — this is our licensed core capability."),
    ("6 Roll fast with racing tires", "ADD", "7 of 8. The bike HAS racing tires. Speed's #1 need is Speed at 134."),
    ("7 Local sales & service", "ADD", "3 of 8. We are 80.9% store and have never claimed it."),
    ("-- Carbon fiber quality at a GREAT PRICE", "DROP", "0 of 8 rivals. And we are raising Swift to $1,580 this quarter — the claim would contradict the price."),
    ("-- Ride safely after dark with LIGHTS", "DROP", "0 of 8. A Recreation cue. The bike KEEPS its lights (+1 in the Speed recipe); we just stop advertising them."),
    ("-- Added safety - REFLECTORS everywhere", "DROP", "0 of 8. Same: keep the part, drop the claim."),
    ("-- Enjoy your ride - carbon fiber light", "DEMOTE or drop", "4 of 8 run it, but it was our rank 2 and reads as Recreation comfort."),
    ("TEST: Highest rated Speed bike", "TEST ONLY", "The Armstrong runs it and sits position 2. Legal ONLY once Swift's design actually hits 77. The sim may validate against Q3 where Swift is 72 — so DESIGN THE BRAND FIRST, then the ad. If rejected, keep local sales & service."),
]:
    rws.append([a, b, c])
r2 = block(r, rws, wrap=(3,))
for i in range(1, len(rws)):
    v = rws[i][1]
    ws.cell(r + i, 2).fill = good if "ADD" in v else (badf if "DROP" in v else yel)
r = r2
ws.cell(r - 1, 1, "Our current Speed ad shares exactly ONE line with the best Speed ad. The three claims nobody else "
                  "runs (lights / reflectors / great price) are Recreation cues and are why the ad leaks 52 Rec and "
                  "36 Mtn units at a judgment of 70. DO BOTH the design fix (24sp->14sp +3, add decals +2 => 77) "
                  "AND the ad fix in the SAME quarter: together Marketing Effectiveness hits 0.7675, past BB LLC's "
                  "0.765 for industry best. Ad alone = 0.755, design alone = 0.750.")
ws.cell(r - 1, 1).alignment = Alignment(wrap_text=True)

rws = [["MEDIA — cost per insert", "Cost", "Note"]]
for m, c in sorted(RATE.items(), key=lambda x: x[1]):
    rws.append([m, c, "CHEAPEST and the top-preference medium for BOTH our segments (Mountain pref 134)"
                if c == 4500 else ("Rec-skewed — Mountain preference 60. Do not buy." if m.startswith("Leisure") else "")])
r2 = block(r, rws, wrap=(3,))
for i in range(1, len(rws)):
    ws.cell(r + i, 2).number_format = "#,##0"
for c in (1, 2, 3):
    ws.cell(r + 1, c).fill = good
r = r2

rws = [["MEDIA OPTIONS", "Hike Biking", "Swift Biking", "Inserts", "Spend", "vs Q3", "Biking total", "Read"]]
for name, h, s, sp, bk, note in OPTS:
    rws.append([name, h, s, h + s, sp, sp - 116_000, bk, note])
r2 = block(r, rws, wrap=(8,))
for i in range(1, len(rws)):
    for c in (5, 6):
        ws.cell(r + i, c).number_format = "#,##0"
    if "recommended" in rws[i][0]:
        for c in range(1, 9):
            ws.cell(r + i, c).fill = good
            ws.cell(r + i, c).font = bold
    if "SUPERSEDED" in rws[i][0] or "Q3 actual" in rws[i][0]:
        for c in range(1, 9):
            ws.cell(r + i, c).fill = badf
r = r2
ws.cell(r - 1, 1, "WHY VOLUME MUST BE RESTRAINED: Market Performance = (avg share in targeted segments) x (percent "
                  "of demand served). Q4 output is capped at 1,154 units and the plan already forecasts 1,171. With "
                  "output capped, extra demand CANNOT raise units sold, so it only lowers the percent-served term — "
                  "strictly negative — and it earns Q5 ill will at half the unmet percentage. Fix the MIX, hold the "
                  "VOLUME, and expand to the full 24-insert BB LLC buy in Q5 when the printers land.")
ws.cell(r - 1, 1).alignment = Alignment(wrap_text=True)

rws = [["ORGANIC SEO — settled points", "Detail"]]
for a, b in [
    ("Page investment is not a lever", "Every listing in the industry cost exactly $1,000 and nobody spent more. Leave both pages at $1,000."),
    ("Rank sets clicks; snippet wording does not", "CTR is a step function of position only: 22.8% / 11.1% / 6.7% / 4.6% / 3.2%. So optimise for AD JUDGMENT, which sets position, not for snippet aesthetics."),
    ("Brand name is free on the SERP", "'Mention brand name' fills the Ad/Brand columns and does NOT consume a snippet slot. A PICTURE does. Rec #1 and Mountain #1 both lead with a claim."),
    ("Never a second page in one segment", "Bike Bros' two Mountain ads drew 79 clicks combined against our single page's 115, and their two MACH I.I Speed pages sank to positions 7 and 8."),
    ("No Recreation page", "We have no Rec brand, and every Rec winner leads with comfort-seat copy we cannot truthfully run on Hike Bike."),
    ("SEO is NOT our web problem", "LiteCycle converts 153 clicks into 435 web units; we convert 144 into 120 — 3.6x worse on nearly identical traffic. The web rebuild (staff 3->7, all four tactics) is the fix, not more clicks."),
]:
    rws.append([a, b])
r2 = block(r, rws, wrap=(2,))
ws.cell(r + 6, 2).fill = yel
r = r2

rws = [["Q4 DECISION SUMMARY", "Action"]]
for a, b in [
    ("HikeBike 1 ad", "Redesign to BB LLC's 6-benefit stack. Add 'Highest rated Mountain bike' and 'Local sales & service'; drop adventure, tough carbon, mountains-easier. Preview and target 82 for position 1 — the payoff is deterministic in Mountain."),
    ("Swift Bike ad", "Redesign to AndStill's 7-benefit stack. Drop lights, reflectors and great price. Design the brand FIRST so 'Highest rated Speed bike' can be tested."),
    ("Organic SEO", "Both pages stay at $1,000. No third page, no Rec page. Fix web OPERATIONS, not traffic."),
    ("Media", "Option C — Biking only, HikeBike 12 + Swift 4 = 16 inserts, $72,000, down $44,000 from Q3."),
    ("Leave alone", "Hike Bike's design (73, tied best, verified optimal) and its price ($1,365, already the optimal tier)."),
    ("Status", "ANALYSIS ONLY — majority vote required. Preview every ad score in the designer before committing."),
]:
    rws.append([a, b])
block(r, rws, wrap=(2,))

ws.column_dimensions["A"].width = 48
ws.column_dimensions["B"].width = 17
ws.column_dimensions["C"].width = 100
for c in "DEFG":
    ws.column_dimensions[c].width = 13
ws.column_dimensions["H"].width = 70

wb.save(DST)
print("\nAdded Q4_Ads_SEO_Media to Q3Data.xlsx")
