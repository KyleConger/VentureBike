"""Score the 15 R&D features against Q1 segment needs importance, then pick 3."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DST = "quarters/Q3/Q3Data.xlsx"

# Need -> (Recreation, Mountain, Speed) importance. Product needs only (web needs excluded).
NEEDS = {
    "Colorful": (116, 120, 109),
    "Fits rider's personality": (116, 113, 116),
    "Simple to use": (121, 102, 69),
    "Safety": (121, 112, 113),
    "Low price point": (117, 98, 70),
    "Sleek": (104, 83, 119),
    "Stylish": (111, 118, 116),
    "Easy to ride": (124, 100, 102),
    "Comfortable": (124, 95, 101),
    "Soft ride": (115, 89, 87),
    "Can handle rough terrain": (111, 128, 82),
    "Can stop quickly": (114, 123, 110),
    "Precise speed control-turns, hills": (103, 106, 117),
    "On-road and off-road capabilities": (114, 95, 69),
    "Can carry things": (117, 104, 69),
    "Durable": (108, 127, 117),
    "Feel young at heart": (120, 115, 117),
    "Can turn sharply": (113, 124, 115),
    "Light weight": (115, 121, 133),
    "High visibility": (117, 103, 119),
    "Easily handles change in incline": (107, 131, 113),
    "Speed": (110, 109, 134),
    "Shifts smoothly": (118, 118, 118),
    "Provides competitive advantage": (98, 117, 124),
    "Aerodynamic": (80, 88, 122),
    "Rider expertise required": (103, 123, 126),
    "Status symbol/exclusivity": (111, 121, 116),
    "Navigation assistance": (115, 117, 118),
    "Soften the impact of rough terrain": (116, 124, 90),
}

# feature -> (dev cost, unit cost, frames compatible, [needs the description claims])
# Need mapping is OUR JUDGMENT from the Workspace feature text, not sim output.
RD = {
    "Enriched carbon fiber": (
        1023325, 40, {"Comfort", "Rugged", "Aerodynamic"},
        ["Light weight", "Durable", "Provides competitive advantage"]),
    "Tires: Mountain super traction": (
        254694, 30, {"Comfort", "Rugged", "Aerodynamic"},
        ["Can handle rough terrain", "Can turn sharply", "Durable", "Can stop quickly"]),
    "Tires: Hybrid comfort": (
        254694, 14, {"Comfort", "Rugged", "Aerodynamic"},
        ["Comfortable", "Soft ride", "Can turn sharply", "On-road and off-road capabilities"]),
    "Tires: Racing sleek very fast": (
        254694, 35, {"Comfort", "Rugged", "Aerodynamic"},
        ["Speed", "Can turn sharply", "Durable", "Sleek", "Precise speed control-turns, hills"]),
    "Bars: Power straight": (
        181924, 17, {"Comfort", "Rugged", "Aerodynamic"},
        ["Can handle rough terrain", "Rider expertise required", "Durable", "Easily handles change in incline"]),
    "Bars: Carbon fiber riser": (
        227406, 24, {"Comfort", "Rugged", "Aerodynamic"},
        ["Comfortable", "Soften the impact of rough terrain", "Can turn sharply", "Light weight"]),
    "Bars: Carbon aero drop down": (
        336560, 40, {"Comfort", "Rugged", "Aerodynamic"},
        ["Aerodynamic", "Light weight", "Sleek", "Comfortable"]),
    "Gears: 11 speed (1x11)": (
        568514, 50, {"Comfort", "Rugged", "Aerodynamic"},
        ["Shifts smoothly", "Light weight", "Easily handles change in incline", "Simple to use"]),
    "Performance bike computer": (
        682217, 60, {"Comfort", "Rugged", "Aerodynamic"},
        ["Navigation assistance", "Provides competitive advantage", "Precise speed control-turns, hills"]),
    "Puncture resistant slime": (
        159184, 3, {"Comfort", "Rugged", "Aerodynamic"},
        ["Durable", "Safety", "Simple to use"]),
    "Pedal-powered charger": (
        291079, 28, {"Comfort", "Rugged", "Aerodynamic"},
        ["Navigation assistance", "Feel young at heart"]),
    "Decals: bright, per segment": (
        109155, 8, {"Comfort", "Rugged", "Aerodynamic"},
        ["Colorful", "Stylish", "Fits rider's personality", "Status symbol/exclusivity"]),
    "Lights: high intensity LED": (
        140991, 13, {"Comfort", "Rugged", "Aerodynamic"},
        ["High visibility", "Safety"]),
    "Carrier: mesh folding tote": (
        90962, 12, {"Comfort"},
        ["Can carry things", "Simple to use"]),
    "Suspension: full (front+back)": (
        618543, 85, {"Comfort", "Rugged"},
        ["Soften the impact of rough terrain", "Can handle rough terrain", "Comfortable",
         "Soft ride", "Precise speed control-turns, hills"]),
}

OUR_FRAMES = {"Hike Bike": "Rugged", "Swift Bike": "Aerodynamic"}

# Frame compatibility is necessary but NOT sufficient. Each segment has a fixed recipe for the
# tire/handlebar/brake slots, so a part can fit a frame physically and still be wrong for the
# segment: racing tires bolt onto a Rugged frame, but the Mountain recipe requires mountain
# high-grip tires and fitting racing tires would break it. This map is which of OUR bikes would
# ACTUALLY be built with the feature, per the verified design model.
WOULD_FIT = {
    "Enriched carbon fiber":          ["Hike Bike", "Swift Bike"],   # material slot, no segment cue
    "Gears: 11 speed (1x11)":         ["Hike Bike", "Swift Bike"],   # gears slot exists in both recipes
    "Decals: bright, per segment":    ["Hike Bike", "Swift Bike"],   # both bikes carry decals
    "Puncture resistant slime":       ["Hike Bike", "Swift Bike"],   # accessory, segment-neutral
    "Performance bike computer":      ["Hike Bike", "Swift Bike"],   # accessory, segment-neutral
    "Tires: Mountain super traction": ["Hike Bike"],                 # Mountain recipe tire
    "Bars: Power straight":           ["Hike Bike"],                 # Mountain uses straight bars
    "Bars: Carbon fiber riser":       ["Hike Bike"],                 # riser is Mountain/Rec
    "Suspension: full (front+back)":  ["Hike Bike"],                 # not Aerodynamic-compatible
    "Tires: Racing sleek very fast":  ["Swift Bike"],                # Speed recipe tire
    "Bars: Carbon aero drop down":    ["Swift Bike"],                # Speed uses drop bars
    "Lights: high intensity LED":     ["Swift Bike"],                # lights excluded from Mountain
    "Tires: Hybrid comfort":          [],                            # Recreation tire
    "Pedal-powered charger":          [],                            # no top need in our segments
    "Carrier: mesh folding tote":     [],                            # Comfort frames only
}

print("=" * 100)
print("STEP 1 — WHAT THE BRAND REVIEWS ACTUALLY SAY (this reframes everything)")
print("=" * 100)
BANDS = [("Poor", "<70", 0), ("Acceptable", "70-77", 19), ("Good", "78-84", 0),
         ("Very Good", "85-94", 0), ("Excellent", "95-100", 0)]
print(f"  {'Band':<12}{'Range':<10}{'Brands in the industry':>24}")
for name, rng, n in BANDS:
    flag = "  <-- ALL 19 rated brands sit here" if n else ""
    print(f"  {name:<12}{rng:<10}{n:>24}{flag}")
print("\n  Every rated brand in the game is 'Acceptable'. NOBODY has reached 'Good' (78+).")
print("  Component ceilings: Mountain 73 · Speed 77 · Recreation 76 — all inside 70-77.")
print("  The scale runs to 100. Three whole bands above the entire industry are UNCLAIMED,")
print("  and R&D is the only way to reach them. 'Customers are only partially satisfied' is literal.")

print("\n" + "=" * 100)
print("STEP 2 — COMPATIBILITY SCREEN (which of OUR bikes can even use each feature)")
print("=" * 100)
for f, (dev, unit, frames, _) in RD.items():
    frame_ok = [b for b, fr in OUR_FRAMES.items() if fr in frames]
    real = WOULD_FIT[f]
    fo = "both" if len(frame_ok) == 2 else (", ".join(frame_ok) if frame_ok else "NEITHER")
    rl = "BOTH BIKES" if len(real) == 2 else (", ".join(real) if real else "NEITHER - skip")
    note = "   (fits the frame but would break the segment recipe)" if len(frame_ok) > len(real) else ""
    print(f"  {f:<32}frame-compatible: {fo:<10} would be BUILT on: {rl}{note}")
print("\n  Full suspension is NOT compatible with Aerodynamic frames -> Hike Bike only.")
print("  The mesh tote basket is Comfort-only -> fits NEITHER of our bikes. Automatic exclusion.")

print("\n" + "=" * 100)
print("STEP 3 — NEEDS-WEIGHTED SCORE (sum of segment importance for the needs each feature claims)")
print("=" * 100)
rows = []
for f, (dev, unit, frames, needs) in RD.items():
    mtn = sum(NEEDS[n][1] for n in needs)
    spd = sum(NEEDS[n][2] for n in needs)
    fits = WOULD_FIT[f]
    eff_mtn = mtn if "Hike Bike" in fits else 0
    eff_spd = spd if "Swift Bike" in fits else 0
    rows.append((f, dev, unit, len(fits), mtn, spd, eff_mtn + eff_spd, needs))

print(f"  {'Feature':<32}{'Dev cost':>11}{'Unit':>6}{'Bikes':>6}{'Mtn':>6}{'Spd':>6}{'USABLE':>8}{'per $100k':>10}")
for f, dev, unit, nb, mtn, spd, eff, needs in sorted(rows, key=lambda x: -x[6]):
    print(f"  {f:<32}{dev:>11,}{unit:>6}{nb:>6}{mtn:>6}{spd:>6}{eff:>8}{eff/(dev/100000):>10.1f}")

print("\n  'USABLE' counts a segment ONLY when the bike serving it would actually be BUILT with the part")
print("  under the verified design recipe — not merely when it bolts onto the frame.")
print("  Need mapping is judgment from the Workspace feature text, NOT sim output. See caveats below.")

print("\n" + "=" * 100)
print("STEP 4 — TOP NEEDS IN OUR TWO SEGMENTS, and which features touch them")
print("=" * 100)
for seg, i in (("Mountain", 1), ("Speed", 2)):
    print(f"\n  {seg} — top 8 product needs:")
    top = sorted(NEEDS.items(), key=lambda x: -x[1][i])[:8]
    for n, v in top:
        hitters = [f for f, (_, _, fr, nd) in RD.items()
                   if n in nd and ("Hike Bike" if seg == "Mountain" else "Swift Bike") in WOULD_FIT[f]]
        print(f"    {v[i]:>4}  {n:<36}{', '.join(hitters) if hitters else '-- nothing on the R&D list --'}")

print("\n" + "=" * 100)
print("STEP 5 — THE DECISION RULE: 3 slots, 2 bikes -> both-bikes features are worth double")
print("=" * 100)
both = [r for r in rows if r[3] == 2]
print("  Features usable on BOTH bikes, ranked by usable needs score:")
for f, dev, unit, nb, mtn, spd, eff, needs in sorted(both, key=lambda x: -x[6])[:6]:
    print(f"    {eff:>4}  {f:<32}${dev:>10,}  ${unit}/unit")
print("\n  Single-segment features score high WITHIN a segment but zero in the other, and Market")
print("  Performance AVERAGES our share across both targeted segments. With only 3 slots, a")
print("  both-bikes feature does two segments' work per slot.")

print("\n" + "=" * 100)
print("STEP 6 — RECOMMENDED THREE")
print("=" * 100)
PICK = ["Decals: bright, per segment", "Enriched carbon fiber", "Gears: 11 speed (1x11)"]
tot_dev = sum(RD[p][0] for p in PICK)
tot_unit = sum(RD[p][1] for p in PICK)
for p in PICK:
    dev, unit, frames, needs = RD[p]
    mtn = sum(NEEDS[n][1] for n in needs)
    spd = sum(NEEDS[n][2] for n in needs)
    print(f"\n  {p}   ${dev:,}  |  ${unit}/unit  |  BOTH BIKES")
    print(f"    Mountain {mtn}  Speed {spd}  combined {mtn+spd}")
    for n in needs:
        print(f"      {n:<38} Mtn {NEEDS[n][1]:>4}   Speed {NEEDS[n][2]:>4}")
print(f"\n  TOTAL development ${tot_dev:,}   added component cost ${tot_unit}/unit (GROSS - see step 7)")

print("\n" + "=" * 100)
print("STEP 7 — MARGIN CHECK (from the brand P&L just supplied)")
print("=" * 100)
BRAND = {"Hike Bike": (379470, 171741, 207729, 68732, 138997, 1365),
         "Swift Bike": (210250, 100599, 109651, 82433, 27218, 1450)}
for b, (rev, cogs, gp, exp, prof, price) in BRAND.items():
    u = round(rev / price)
    print(f"\n  {b}: {u} units")
    print(f"    revenue/unit      ${rev/u:>8,.0f}")
    print(f"    COGS/unit         ${cogs/u:>8,.0f}")
    print(f"    gross margin/unit ${gp/u:>8,.0f}")
    print(f"    brand expense/unit${exp/u:>8,.0f}   <- ad + design, almost entirely FIXED")
    print(f"    profit/unit       ${prof/u:>8,.0f}")
print("\n  Swift Bike's $188 profit/unit is NOT a structural problem: it carries a one-time $30,000")
print("  design fee and spreads $45,433 of advertising over just 145 units ($313/unit). At BB LLC's")
print("  Speed volume of 363 the same ad dollars cost $125/unit. Advertising is fixed - VOLUME is")
print("  what crushes cost per unit, which is the whole case for the capacity fix.")
print(f"\n  Adding ${tot_unit}/unit of new components against gross margin:")
print(f"    Hike Bike  $747 -> ${747-tot_unit}/unit   (still {(747-tot_unit)/1365:.0%} of a $1,365 price)")
print(f"    Swift Bike $886 -> ${886-tot_unit}/unit   (at the new $1,580 price)")
print("\n  IMPORTANT: these are REPLACEMENT components, not additions. Enriched carbon replaces")
print("  standard carbon, the new decals replace the old decals, 11-speed replaces 24-speed (and the")
print("  text says 'fewer parts'). The true incremental cost is the DIFFERENCE from what we ship")
print("  today, which is smaller than $98 and possibly much smaller. Read it on the Design screen.")

print("\n" + "=" * 100)
print("STEP 8 — FUNDING AND THE 50-75% GUIDANCE")
print("=" * 100)
CASH, VC = 1010838, 2500000
avail = CASH + VC
print(f"  Cash on hand                     ${CASH:>10,}")
print(f"  VC raise (take the full amount)  ${VC:>10,}")
print(f"  Available                        ${avail:>10,}")
print(f"  50% guidance floor               ${avail*0.50:>10,.0f}")
print(f"  75% guidance ceiling             ${avail*0.75:>10,.0f}")
print(f"\n  Recommended three at listed cost ${tot_dev:>10,}  = {tot_dev/avail:.1%} of available")
if tot_dev < avail * 0.5:
    print(f"  -> ${avail*0.5-tot_dev:,.0f} BELOW the 50% floor at the listed prices.")
print("\n  The listed figures are 'Cost to finish' with no quarter selected yet. The decision tip says")
print("  a 1-quarter (rapid) schedule costs MORE in total than spreading over 2 quarters, so the")
print("  1-quarter prices will be higher than what is displayed and should close most of that gap.")
print("\n  We are SLOT-constrained (3 projects) and NOT cash-constrained ($3.5M available). When slots")
print("  are scarce and money is not, buy SPEED - schedule 1-quarter so features land in Q5 designs")
print("  instead of Q6, and Q5's expanded engineering capacity is free for new projects. If the total")
print("  still lands under 50%, close the gap by paying for rapid development, NOT by adding a")
print("  weaker fourth-choice feature.")

print("\n" + "=" * 100)
print("STEP 9 — WHAT WE SKIP AND WHY")
print("=" * 100)
SKIP = [
    ("Carrier: mesh folding tote", "Comfort frames ONLY - fits neither of our bikes. Hard exclusion."),
    ("Suspension: full (front+back)", "Not compatible with Aerodynamic, so Hike Bike only. Highest unit cost on "
     "the list at $85 against a $747 margin, and $618,543 to develop. Strong Mountain needs "
     "(rough terrain 128, soften impact 124) but half the reach for more money."),
    ("Tires: Hybrid comfort", "Recreation feature. Comfortable scores 95 in Mountain and 101 in Speed."),
    ("Pedal-powered charger", "Touches no top-10 need in either of our segments."),
    ("Performance bike computer", "$60/unit and $682,217 to develop. Navigation assistance is only 117/118 - "
     "mid-tier - and no rival ad in the game claims anything like it."),
    ("Tires: Mountain super traction", "Genuinely strong for Mountain (rough terrain 128, turn sharply 124, durable "
     "127) but scores only 82 on rough terrain in Speed. Best single-segment option - hold for Q5."),
    ("Tires: Racing sleek very fast", "Strong for Speed (Speed 134, sleek 119) and directly supports the new Swift "
     "ad copy. Best Speed-only option - hold for Q5."),
    ("Bars: Carbon aero drop down", "Supports 'ride a wind-cheater', the one claim on 8 of 8 rival Speed ads. "
     "Aerodynamic scores 122 in Speed but 88 in Mountain. Hold for Q5."),
    ("Bars: Power straight / CF riser", "Mountain and Recreation leaning, and the Power straight text admits a "
     "downside (reduced tree clearance)."),
    ("Puncture resistant slime", "Only $3/unit and $159,184 - the best value per unit on the list, hitting Durable "
     "(Mtn 127 / Spd 117) and Safety. It is the strongest CHEAP alternative if the team wants to "
     "conserve cash, but it touches only 2-3 needs, so it wastes a scarce slot."),
    ("Lights: high intensity LED", "High visibility is 119 in Speed but 103 in Mountain, and no Mountain brand in "
     "the industry ships lights at all. Our own ad analysis says drop lights from Speed copy."),
]
for f, why in SKIP:
    print(f"\n  {f}\n    {why}")

print("\n" + "=" * 100)
print("STEP 10 — CAVEATS THAT MATTER")
print("=" * 100)
print("""
  1. The needs-to-feature mapping above is OUR reading of the Workspace feature descriptions.
     The sim has not told us which needs each component satisfies. Treat the scores as a ranking
     aid, not a prediction of brand judgment points.
  2. Our design model measured point values for gears, seat, decals, reflectors and lights ONLY.
     Frame, tires, brakes and handlebars never varied within a segment, so we have NO measured
     evidence for what upgrading them is worth. Decals (+2 in Speed) and gears (14sp=0, 24sp=-3,
     7sp=-16) are the two component classes we have actually measured - and both recommended
     picks land on them.
  3. 11-speed is the highest-variance pick. Gears show a 16-point swing in Speed, the largest of
     any component we can measure, so the upside is real. But 24-speed is currently OPTIMAL in
     Mountain and our own ad claims 'more gears', and 11 is fewer than 24. The feature text
     answers that objection directly ('11 gears can do the job of 24... lighter, easier to shift,
     wide range') and 'easily handles change in incline' is Mountain's #1 need at 131. Still an
     inference. If the team wants lower variance, swap in puncture slime ($159,184, $3/unit).
  4. A bonus if 11-speed works: it could make ONE drivetrain optimal on both bikes. Today Mountain
     wants 24-speed and Speed wants 14-speed, and mixing them up is exactly the error that cost
     Swift Bike 3 points.
  5. Nothing here is locked. Three slots, majority vote.
""")
print("=" * 100)

# ---------------------------------------------------------------- workbook
wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
badf = PatternFill("solid", fgColor="FFC7CE")
good = PatternFill("solid", fgColor="C6EFCE")
yellow = PatternFill("solid", fgColor="FFEB9C")
thin = Border(*[Side(style="thin", color="D0D0D0")] * 4)

if "Q4_RD" in wb.sheetnames:
    del wb["Q4_RD"]
ws = wb.create_sheet("Q4_RD")


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


ws["A1"] = "Q4 R&D — 3 project slots, scored against segment needs importance"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("EVERY rated brand in the industry is 'Acceptable' (70-77). Nobody has reached 'Good' (78-84), "
            "'Very Good' (85-94) or 'Excellent' (95-100). The component ceilings are Mountain 73, Speed 77, "
            "Recreation 76 — all inside one band. R&D is the only route into three unclaimed bands. "
            "Needs importance from Q1 market research (`Needs` sheet); feature-to-need mapping is our reading "
            "of the Workspace descriptions, NOT sim output.")
ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A2:H2")
ws.row_dimensions[2].height = 62

r = 4
rws = [["Feature", "Dev cost", "Unit cost", "Fits our bikes", "Mtn score", "Speed score",
        "Usable score", "Score per $100k", "Needs claimed"]]
for f, dev, unit, nb, mtn, spd, eff, needs in sorted(rows, key=lambda x: -x[6]):
    fits = WOULD_FIT[f]
    rws.append([f, dev, unit, "BOTH" if nb == 2 else (", ".join(fits) if fits else "NEITHER"),
                mtn, spd, eff, round(eff / (dev / 100000), 1), ", ".join(needs)])
r2 = block(r, rws, wrap=(9,))
for i in range(1, len(rws)):
    ws.cell(r + i, 2).number_format = "#,##0"
    if rws[i][0] in PICK:
        for c in range(1, 10):
            ws.cell(r + i, c).fill = good
            ws.cell(r + i, c).font = bold
    if rws[i][3] == "NEITHER":
        ws.cell(r + i, 4).fill = badf
r = r2

rws = [["RECOMMENDED THREE — all usable on BOTH bikes", "Dev cost", "Unit cost", "Why this one"]]
for p, why in [
    ("Decals: bright, styled per segment & frame", "Highest combined needs score on the entire list (929). "
     "Hits Colourful (Mtn 120), Status symbol (Mtn 121), Stylish (Mtn 118), Fits personality (Spd 116) — four "
     "needs all scoring 109-121 in BOTH segments. Cheapest project ($109,155) and cheapest component ($8). "
     "And decals are one of only two component classes our design model has MEASURED: worth +2 in Speed. "
     "Highest confidence per dollar on the board."),
    ("Enriched carbon fiber", "Hits Light weight (Speed's #2 need at 133, Mtn 121), Durable (Mountain's #3 at "
     "127, Spd 117) and Provides competitive advantage (Spd 124, Mtn 117). Upgrades the material on BOTH "
     "bikes and every future brand, since every bike in the game carries carbon fiber. Also the single most "
     "on-strategy item we could fund — our entire company premise is licensed carbon-fiber 3D printing."),
    ("Gears: 11 speed (1x11)", "Hits Easily handles change in incline (Mountain's #1 need at 131), Light "
     "weight (121/133) and Shifts smoothly (118/118). Gears are the STRONGEST lever our model has measured — "
     "a 16-point swing in Speed between 14sp and 7sp. Highest variance of the three: 24-speed is currently "
     "optimal in Mountain and 11 is fewer gears, though the feature text answers that directly. Bonus if it "
     "lands: ONE drivetrain optimal on both bikes, eliminating the copy-paste error that cost Swift Bike 3 points."),
]:
    key = [k for k in RD if k.split(":")[0] in p or p.startswith(k.split(":")[0])]
    dev, unit = (RD[PICK[["Decals", "Enriched", "Gears"].index(p.split(":")[0].split()[0])]][0],
                 RD[PICK[["Decals", "Enriched", "Gears"].index(p.split(":")[0].split()[0])]][1]) \
        if p.split(":")[0].split()[0] in ("Decals", "Enriched", "Gears") else (0, 0)
    rws.append([p, dev, unit, why])
rws.append(["TOTAL", tot_dev, tot_unit, f"{tot_dev/avail:.1%} of the ${avail:,} available. Gross added "
            f"component cost ${tot_unit}/unit, but all three are REPLACEMENT parts — the true increment is "
            f"the difference from what we ship today and is smaller."])
r2 = block(r, rws, wrap=(4,))
for i in range(1, len(rws)):
    ws.cell(r + i, 2).number_format = "#,##0"
ws.cell(r + len(rws) - 1, 2).fill = yellow
r = r2

rws = [["SCHEDULING — 1 quarter or 2", "Reasoning"]]
for a, b in [
    ("Recommend 1-quarter (rapid) on all three", "We are SLOT-constrained (3 projects), not cash-constrained "
     "($3.5M available). When slots are scarce and money is not, buy speed: features land in Q5 designs "
     "instead of Q6, and Q5's expanded engineering capacity stays free for new projects."),
    ("The listed prices are not final", "The table shows 'Cost to finish' with no quarter selected. The tip "
     "says a 1-quarter schedule costs MORE in total than spreading over two, so the rapid prices will be "
     "higher than displayed. Read the real numbers after selecting a quarter."),
    ("If the total lands under the 50% floor", "Close the gap by paying for rapid development, NOT by adding a "
     "weaker fourth-choice feature. Three good features delivered a quarter early beats four mediocre ones."),
    ("Ask the team", "How many quarters remain in the simulation? If it ends at Q6, 1-quarter scheduling is "
     "mandatory — a Q6 feature would never ship. This is the one input we do not have."),
]:
    rws.append([a, b])
r2 = block(r, rws, wrap=(2,))
ws.cell(r + 1, 2).fill = good
ws.cell(r + 4, 2).fill = yellow
r = r2

rws = [["FUNDING", "Amount", "Note"]]
for a, b, c in [
    ("Cash on hand", CASH, "idle — this is what Asset Management 0.353 punishes"),
    ("VC raise", VC, "take the full amount. This RECONCILES with our earlier 'stop issuing stock' rule, which "
     "was conditional on the cash sitting idle. R&D is the productive use that justifies a raise."),
    ("Available", avail, ""),
    ("50% guidance floor", round(avail * 0.50), "instructor guidance: plan to spend 50-75% on R&D"),
    ("75% guidance ceiling", round(avail * 0.75), ""),
    ("Recommended three (listed cost)", tot_dev, f"{tot_dev/avail:.1%} — just under the floor at listed prices; "
     f"rapid scheduling should close it"),
    ("Also committed this quarter", 146000, "the copy-the-leader operating plan"),
    ("Hard floor", 300000, "ending Cash + CD must stay at or above this"),
]:
    rws.append([a, b, c])
r2 = block(r, rws, wrap=(3,))
for i in range(1, len(rws)):
    ws.cell(r + i, 2).number_format = "#,##0"
ws.cell(r + 2, 3).fill = yellow
r = r2

rws = [["BRAND P&L — what it tells us about cost per unit", "Hike Bike", "Swift Bike"]]
for label, h, s in [
    ("Units", 278, 145),
    ("Revenue", 379470, 210250),
    ("COGS", 171741, 100599),
    ("COGS per unit", 618, 694),
    ("Gross margin per unit", 747, 756),
    ("Brand expense (ad + design)", 68732, 82433),
    ("Brand expense PER UNIT", 247, 568),
    ("Profit per unit", 500, 188),
]:
    rws.append([label, h, s])
rws.append(["Read", None, "Swift Bike's $188 is not structural. It carries a one-time $30,000 design fee and "
            "spreads $45,433 of advertising over only 145 units ($313/unit). At BB LLC's Speed volume of 363 "
            "the same ad dollars cost $125/unit. Advertising is FIXED — volume is what crushes cost per unit. "
            "Swift also costs $76 more per unit to build than Hike Bike, and the 14-speed swap should narrow that."])
r2 = block(r, rws, wrap=(3,))
for i in range(1, len(rws) - 1):
    for c in (2, 3):
        ws.cell(r + i, c).number_format = "#,##0"
ws.cell(r + len(rws) - 1, 3).fill = yellow
r = r2

rws = [["HOLD FOR Q5 (engineering capacity expands next quarter)", "Why not now"]]
for f, why in SKIP:
    rws.append([f, why])
r2 = block(r, rws, wrap=(2,))
ws.cell(r + 1, 2).fill = badf
r = r2

rws = [["CAVEATS", "Detail"]]
for a, b in [
    ("Need mapping is our judgment", "The sim has not told us which needs each component satisfies. These are "
     "readings of the Workspace feature text. Use the scores as a ranking aid, not a points prediction."),
    ("Only 5 components are MEASURED", "Our design model derived point values for gears, seat, decals, "
     "reflectors and lights only. Frame, tires, brakes and handlebars never varied within a segment, so we "
     "have no measured evidence for upgrading them. Two of the three picks land on measured components."),
    ("11-speed is the risk", "24-speed is currently optimal in Mountain and our ad claims 'more gears'. The "
     "feature text answers this ('11 gears can do the job of 24... lighter, easier to shift, wide range') and "
     "'easily handles change in incline' is Mountain's #1 need at 131 — but it is still an inference. Lower-"
     "variance swap: puncture resistant slime, $159,184 and $3/unit."),
    ("Replacement not addition", "All three picks replace parts we already fit, so the incremental unit cost is "
     "less than the $98 sticker. Read the delta on the Design Brand screen."),
    ("Status", "ANALYSIS ONLY. Three slots, majority vote required."),
]:
    rws.append([a, b])
block(r, rws, wrap=(2,))

ws.column_dimensions["A"].width = 42
ws.column_dimensions["B"].width = 14
ws.column_dimensions["C"].width = 14
ws.column_dimensions["D"].width = 88
for col in "EFGH":
    ws.column_dimensions[col].width = 13
ws.column_dimensions["I"].width = 60

wb.save(DST)
print("\nAdded Q4_RD to Q3Data.xlsx")
