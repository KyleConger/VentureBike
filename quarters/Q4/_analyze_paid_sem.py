"""Q4 paid SEM (PPC): max bid and quarterly budget per web page.

Industry averages from the Workspace paid-SEM screen (prior quarter, all
companies selling lightweight / carbon-fiber bikes on the Internet):
  Mountain $2.40 · Speed $1.95
"""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DST = "quarters/Q3/Q3Data.xlsx"

MTN_AVG = 2.40
SPD_AVG = 1.95
MTN_SEARCHES = 1_027
SPD_SEARCHES = 894
CAPACITY = 1_154
FORECAST = 1_171
SHORTFALL = FORECAST - CAPACITY  # demand already 17 over capacity
HIKE_GM = 747
SWIFT_GM = 886  # at the $1,580 price
HIKE_ADJ = 80
SWIFT_ADJ = 70
HIKE_ORG = 115
SWIFT_ORG = 29

# Recommendation — match the posted average on the bid, throttle with the budget.
HIKE_BID = 2.40
HIKE_BUDGET = 150
SWIFT_BID = 1.95
SWIFT_BUDGET = 80


def clicks(budget, bid, cpc_frac=1.0):
    return budget / (bid * cpc_frac)


print("=" * 100)
print("PART 1 — THIS IS A VOLUME DECISION, NOT A PROFITABILITY DECISION")
print("=" * 100)
print(f"""
  Q4 printers produce {CAPACITY:,} units (24/day x 0.74 x 65). The plan of record already
  forecasts {FORECAST:,} units of demand. Demand exceeds capacity by {SHORTFALL} units.

  Market Performance = (avg share in targeted segments) x (percent of demand served).
  With output capped, extra paid clicks cannot raise units sold — they only lower
  percent-served. Same logic that restrained media volume (Option C). Paid SEM is
  demand generation that sits ABOVE the organic listings.

  Unit economics would say bid almost anything: at a 15% click-to-demand conversion,
    Hike Bike value/click  = 0.15 x ${HIKE_GM}  = ${0.15 * HIKE_GM:.0f}
    Swift Bike value/click = 0.15 x ${SWIFT_GM} = ${0.15 * SWIFT_GM:.0f}
  The industry averages (${MTN_AVG:.2f} / ${SPD_AVG:.2f}) are 2% of that. Cash ROI is
  not the constraint. Filling the order is.
""")

print("=" * 100)
print("PART 2 — DO NOT SIT OUT. PAID SITS ABOVE ORGANIC.")
print("=" * 100)
print(f"""
  HikeBike 1 is Mountain organic #2 ({HIKE_ORG} clicks, ad {HIKE_ADJ}). Swift Bike is
  Speed #5 ({SWIFT_ORG} clicks, ad {SWIFT_ADJ}). Sponsored links render above those
  listings. If BB LLC (Mountain organic #1, 237 clicks, and the firm that USED
  overtime to fill demand) buys paid and we bid $0, they intercept Mountain
  searchers before our organic result. We still keep organic, but we hand them
  the slot above it.

  So: appear in paid. Cap the clicks. That is the whole strategy.
""")

print("=" * 100)
print("PART 3 — BID = POSITION. BUDGET = VOLUME. SET THEM SEPARATELY.")
print("=" * 100)
print(f"""
  Posted industry averages (Workspace): Mountain ${MTN_AVG:.2f} · Speed ${SPD_AVG:.2f}.
  The tip says bid at, below, or above that average. Higher bid -> better position
  -> more clicks, and it raises NEXT quarter's suggested bid if everyone piles on.

  Actual CPC <= max bid. It depends on rival bids AND how the engine rates the
  page (ad judgment is the quality score). HikeBike 1 at {HIKE_ADJ} (heading to 81-82)
  should pay LESS than Swift at {SWIFT_ADJ} (heading to 77) at the same bid.

  Lecture trap: a high bid + small budget wins week 1 and disappears. A moderate
  bid + a budget that lasts can have the better AVERAGE position because the
  over-bidders hit their caps. Matching the posted average IS the moderate bid
  this quarter — classmates unlocking paid SEM for the first time will be the
  ones who bid $3–$4 to 'win' the auction.
""")

print("  CLICK CAPS AT THE RECOMMENDED BIDS")
print(f"  {'Page':<16}{'Bid':>8}{'Budget':>10}{'Clicks @ bid':>14}{'Clicks @ 75% CPC':>18}")
for name, bid, bud in (("HikeBike 1", HIKE_BID, HIKE_BUDGET),
                       ("Swift Bike", SWIFT_BID, SWIFT_BUDGET)):
    print(f"  {name:<16}{bid:>8.2f}{bud:>10,.0f}{clicks(bud, bid):>14.0f}"
          f"{clicks(bud, bid, 0.75):>18.0f}")
print(f"  {'TOTAL':<16}{'':>8}{HIKE_BUDGET + SWIFT_BUDGET:>10,.0f}"
      f"{clicks(HIKE_BUDGET, HIKE_BID) + clicks(SWIFT_BUDGET, SWIFT_BID):>14.0f}"
      f"{clicks(HIKE_BUDGET, HIKE_BID, 0.75) + clicks(SWIFT_BUDGET, SWIFT_BID, 0.75):>18.0f}")

hike_max = clicks(HIKE_BUDGET, HIKE_BID)
swift_max = clicks(SWIFT_BUDGET, SWIFT_BID)
hike_75 = clicks(HIKE_BUDGET, HIKE_BID, 0.75)
swift_75 = clicks(SWIFT_BUDGET, SWIFT_BID, 0.75)
print(f"""
  READ: ${HIKE_BUDGET + SWIFT_BUDGET} buys at most {hike_max + swift_max:.0f} paid clicks
  if we pay the full bid, ~{hike_75 + swift_75:.0f} if quality knocks CPC down 25%.
  Compare to organic (144) and to Mountain searches alone ({MTN_SEARCHES:,}). This is a
  probe, not a campaign.

  If even 20% of those clicks become demand, that is ~{0.20 * (hike_max + swift_max):.0f}
  extra units against a {SHORTFALL}-unit hole — a small overshoot, not a 204-unit stock-out.
  A $1,000 Hike budget at $2.40 would cap at {clicks(1000, HIKE_BID):.0f} clicks.
""")

print("=" * 100)
print("PART 4 — THE NUMBERS TO ENTER")
print("=" * 100)
print(f"""
  HIKEBIKE 1  (Mountain page)
    Keywords          Mountain set (agency — already chosen when the page was created)
    Maximum bid       ${HIKE_BID:.2f}     match the ${MTN_AVG:.2f} industry average
    Maximum budget    ${HIKE_BUDGET:,}       cap ~{clicks(HIKE_BUDGET, HIKE_BID):.0f} clicks
    Why this page gets more: primary segment, organic #2 to defend, 1 Mountain unit
    is worth 1.78x a Speed unit for Market Performance, ad quality {HIKE_ADJ} will
    actually pay below ${HIKE_BID:.2f}.

  SWIFT BIKE  (Speed page)
    Keywords          Speed set (agency)
    Maximum bid       ${SWIFT_BID:.2f}     match the ${SPD_AVG:.2f} industry average
    Maximum budget    ${SWIFT_BUDGET:,}        cap ~{clicks(SWIFT_BUDGET, SWIFT_BID):.0f} clicks
    Why less: secondary segment, smallest search pool ({SPD_SEARCHES}), Speed SERP
    rank is not predicted by ad judgment, and the brand/ad/price fixes are already
    adding Speed demand we may not be able to build.

  TOTAL PAID SEM      ${HIKE_BUDGET + SWIFT_BUDGET}  (noise against the ~$146k operating plan)

  DO NOT
    - Bid above the posted averages. That starts the Q5 bid inflation the lecture
      warns about, and it spends the small budget in days.
    - Set $500–$1,000 budgets. That is 200–500 extra clicks against a 17-unit hole.
    - Put Recreation keywords on either page. Rec stays parked.
    - Add a third page. Bike Bros' two Mountain pages drew 79 clicks combined
      against our one page's 115.
    - Skip paid entirely. The slot above organic would go to BB LLC.

  Q5 (printers land, ~40/day): raise budgets toward $400–$600/page and keep bidding
  at whatever the new posted average is. That is when extra clicks can be filled.
""")

# ---------------------------------------------------------------- workbook
wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
good = PatternFill("solid", fgColor="C6EFCE")
badf = PatternFill("solid", fgColor="FFC7CE")
yel = PatternFill("solid", fgColor="FFEB9C")
thin = Border(*[Side(style="thin", color="D0D0D0")] * 4)

if "Q4_Paid_SEM" in wb.sheetnames:
    del wb["Q4_Paid_SEM"]
ws = wb.create_sheet("Q4_Paid_SEM")


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


ws["A1"] = "Q4 paid SEM — match the posted average on the bid, throttle with the budget"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = (
    "Paid search is demand generation that sits ABOVE organic. Q4 output is capped at "
    f"{CAPACITY:,} against a {FORECAST:,} forecast (demand already {SHORTFALL} over), so extra clicks "
    "are strictly negative for Market Performance — the same reason media volume was "
    "restrained. Do not sit out: BB LLC will buy the slot above HikeBike 1's organic #2. "
    f"Enter ${HIKE_BID:.2f} / ${HIKE_BUDGET} on HikeBike 1 and ${SWIFT_BID:.2f} / ${SWIFT_BUDGET} "
    "on Swift Bike. Keywords = the agency set already attached to each page (Mountain / Speed). "
    "OPEN — majority vote required."
)
ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A2:G2")
ws.row_dimensions[2].height = 72

r = 4
rws = [
    ["Page", "Segment / keywords", "Industry avg bid", "OUR max bid", "OUR max budget",
     "Clicks at max bid", "Clicks if CPC = 75% of bid"],
    ["HikeBike 1", "Mountain (agency set)", MTN_AVG, HIKE_BID, HIKE_BUDGET,
     round(clicks(HIKE_BUDGET, HIKE_BID), 1), round(clicks(HIKE_BUDGET, HIKE_BID, 0.75), 1)],
    ["Swift Bike", "Speed (agency set)", SPD_AVG, SWIFT_BID, SWIFT_BUDGET,
     round(clicks(SWIFT_BUDGET, SWIFT_BID), 1), round(clicks(SWIFT_BUDGET, SWIFT_BID, 0.75), 1)],
    ["TOTAL", "", "", "", HIKE_BUDGET + SWIFT_BUDGET,
     round(clicks(HIKE_BUDGET, HIKE_BID) + clicks(SWIFT_BUDGET, SWIFT_BID), 1),
     round(clicks(HIKE_BUDGET, HIKE_BID, 0.75) + clicks(SWIFT_BUDGET, SWIFT_BID, 0.75), 1)],
]
r2 = block(r, rws)
for c in range(1, 8):
    ws.cell(r + 1, c).fill = good
    ws.cell(r + 2, c).fill = good

rws = [
    ["Why this split", "Detail"],
    ["Bid matches the posted average",
     "Classmates unlocking paid SEM will be the ones who bid $3–$4. Matching $2.40 / $1.95 "
     "IS the lecture's moderate bid: we stay in the auction all quarter as they exhaust."],
    ["Budget is the volume throttle",
     f"$150 / $80 caps us at ~{clicks(HIKE_BUDGET, HIKE_BID) + clicks(SWIFT_BUDGET, SWIFT_BID):.0f} "
     "clicks if we pay the full bid. A $1,000 Hike budget would cap at "
     f"{clicks(1000, HIKE_BID):.0f} clicks we cannot build."],
    ["Hike Bike gets the larger budget",
     f"Primary segment. Organic #2 ({HIKE_ORG} clicks) must not lose the slot above it to "
     "BB LLC. One Mountain unit = 1.78x a Speed unit for Market Performance. Ad 80 "
     "(heading to 81–82) is a quality-score advantage, so actual CPC should land below $2.40."],
    ["Swift Bike gets the probe budget",
     f"Secondary segment, smallest search pool ({SPD_SEARCHES}), Speed rank is not predicted "
     "by ad judgment, and the 72→77 brand + 70→77 ad + $1,450→$1,580 price are already "
     "adding Speed demand."],
    ["Do not bid $0",
     "Paid listings render above organic. Sitting out hands BB LLC the intercept on "
     "Mountain searches. We still have organic after the budget caps; we just do not "
     "surrender the sponsored slot for the whole quarter."],
    ["Do not bid above the average",
     "Starts the Q5 bid-inflation the lecture warns about, spends the small budget in "
     "days (high bid + small budget = the trap), and buys demand the 24 printers cannot fill."],
    ["Keywords are already set",
     "Choosing the segment when the page was created attached the agency keyword set. "
     "HikeBike 1 = Mountain. Swift Bike = Speed. Do not retarget either page at Recreation."],
    ["Cash is not the issue",
     f"${HIKE_BUDGET + SWIFT_BUDGET} is noise against the ~$146k operating plan and the $950k "
     "loan. The cost of over-bidding is unmet demand, not the click invoice."],
    ["Q5 is when this scales",
     "Two printers land next quarter (~40/day). Raise budgets toward $400–$600/page and "
     "re-read the new posted averages. Extra clicks can be filled then."],
]
r3 = block(r2, rws, wrap=(1, 2))
ws.column_dimensions["A"].width = 34
ws.column_dimensions["B"].width = 88
for col, w in [("C", 20), ("D", 16), ("E", 18), ("F", 20), ("G", 26)]:
    ws.column_dimensions[col].width = w

rws = [
    ["Rejected alternatives", "Bid Hike / Swift", "Budget Hike / Swift", "Why not"],
    ["Sit out paid SEM", "$0 / $0", "$0 / $0",
     "Hands BB LLC the sponsored slot above our Mountain organic #2."],
    ["Shade 15c below average", "$2.25 / $1.80", "$150 / $80",
     "Fine if the auction is sleepy; risky if 6 classmates open at $3+. Match is safer."],
    ["Copy a 'win the auction' bid", "$3.50 / $3.00", "$1,000 / $600",
     "Burns cash in days, inflates Q5 averages, and dumps 400+ clicks on a 17-unit hole."],
    ["Equal budgets", f"${HIKE_BID:.2f} / ${SWIFT_BID:.2f}", "$115 / $115",
     "Ignores Mountain-primary and the organic-#2 defense. Speed is the smaller search pool."],
]
block(r3, rws, wrap=(4,))

wb.save(DST)
print("Added Q4_Paid_SEM")
