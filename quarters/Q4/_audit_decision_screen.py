"""Audit the Q4 decision screen dump against the Q4 plan of record.

Writes Q4_Screen_Audit into quarters/Q3/Q3Data.xlsx.
"""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

XLSX = "quarters/Q3/Q3Data.xlsx"
SHEET = "Q4_Screen_Audit"

# ---------------------------------------------------------------- constants
DAYS = 65
PROD_73 = 0.73          # projected worker productivity on screen
PROD_75 = 0.75          # projected if production comp -> BB LLC package
OWNED_OC = 24


def cap(oc, prod):
    return round(oc * DAYS * prod)


# ------------------------------------------------------- line-by-line audit
# (screen area, what the screen shows, what the plan says, verdict, consequence)
AUDIT = [
    # ---- BLOCKERS -------------------------------------------------------
    ("Manufacturing / Operating capacity",
     "Operating capacity = 8/day (520/qtr); effective 5.84/day = 380 units",
     "Schedule the full 24/day owned, overtime 0.00",
     "BLOCKER",
     f"380 units vs ~1,142 projected demand. This is the Q3 error verbatim: "
     f"67% stock-out, Q5 ill will ~33%. 24/day at 73% = {cap(OWNED_OC, PROD_73)} units."),

    ("Improve Web Productivity",
     "Total expenses = 6,000. Only toll-free shows status 'operational'. "
     "Secure site / cart / page upgrades / order tracking all status '-' "
     "with budgets typed but Start not checked",
     "Fund ALL tactics: setup 48,000 + quarterly 25,000 on the four new ones",
     "BLOCKER",
     "Q3 evidence is a step function: firms funding all tactics land 435-569 web "
     "units; firms funding 1-2 land 120-138. Typing a budget without checking Start "
     "leaves us in the losing band while paying 45,501 for 7 web staff."),

    ("Finance / Stock",
     "Common Stock, Venture Capitalists, 25,000 shares @ 100 = 2,500,000, Quarter 4",
     "Do NOT take the 2.5M equity round; remove or zero this row",
     "BLOCKER",
     "2.5M against 677,669 of R&D recreates the idle-cash problem that made Asset "
     "Management 0.353 = LAST, and 25,000 new shares halve Financial Performance "
     "(divided by total shares issued)."),

    ("Finance / Short Term Loan",
     "'No outstanding conventional loans this quarter'",
     "Draw the ~950,000 conventional loan at 7% annual",
     "BLOCKER",
     "Cash 1,010,838 cannot cover R&D 677,669 + ~380,000 visible operating + "
     "2 printers and still clear the 300,000 floor. Nothing is funded without this."),

    ("Demand Projection",
     "Store demand/SP = 0, Web demand/SP = 0, all totals 0, "
     "operating capacity required = 0",
     "Store 50.5/head x 14 = 707; Web 62.1/head x 7 = 435; total 1,142",
     "BLOCKER",
     "This is the input the capacity decision is read against. Left at 0 it argues "
     "for OC 0 and hides the 8/day error."),

    ("Price and Priority / Swifter Bike",
     "Retail Price 1,450",
     "Raise to 1,580",
     "MISS",
     "1,580 draws ZERO Speed price resistance (proven by Blu Aero, Blu Tube, "
     "Armstrong at that exact price). +130/unit of pure margin on ~457 units "
     "= +59,410, and it relieves the capacity squeeze."),

    # ---- COMP -----------------------------------------------------------
    ("HR / Sales Force Compensation",
     "19,000 salary + Full coverage + 2 weeks + 1% = 24,425/yr; projected prod 70",
     "22,000 + Expanded + 2 weeks + 4% = 27,402/yr (BB LLC exact); project 75",
     "MISS",
     "Still the Q2 lock. Mix is the error, not the total: BB LLC hit 78.4% on "
     "Expanded/2wk/4%; MILC paid MORE (27,630) on Full/1wk and scored 8.4 pts worse. "
     "Cost +2,977/head/yr = +15,634/qtr across 21 heads."),

    ("HR / Production Worker Compensation",
     "16,800 + Expanded + 2 weeks + 3% = 20,757/yr; projected prod 73",
     "18,500 + Expanded + 2 weeks + 4% = 23,042/yr (BB LLC exact); project 75",
     "MISS",
     f"Buys capacity directly: 73% -> 75% takes 24/day from {cap(OWNED_OC, PROD_73)} "
     f"to {cap(OWNED_OC, PROD_75)} units (+{cap(OWNED_OC, PROD_75)-cap(OWNED_OC, PROD_73)}) "
     "and cuts labour cost/unit."),

    # ---- MEDIA ----------------------------------------------------------
    ("Regional Media / off-Biking inserts",
     "Swifter Bike 1: Sport 1 (10,000), Business 1 (9,500), New Venture 2 (11,000)",
     "Concentrate in Biking; New Venture and Sport are the Q3 diversification error",
     "REVIEW",
     "Biking is 51 industry pages and BB LLC owns 24 of them at 4,500 each - the "
     "cheapest insert on the board. Business is Speed's #2 medium so Business 1 is "
     "defensible; Sport 1 + New Venture 2 = 21,000 for 3 poorly-targeted pages."),

    ("Regional Media / totals do not reconcile",
     "Per-brand totals 42,119 + 45,433 = 87,552",
     "20 inserts at list = 12x4,500 + 4x4,500 + 10,000 + 9,500 + 2x5,500 = 102,500",
     "REVIEW",
     "45,433 is exactly Swift Bike's Q3 brand advertising expense, so the Total row "
     "looks like prior-period or amortised figures, not this quarter's bill. "
     "Confirm what will actually be charged before locking."),

    ("Modify Ad / HikeBike 2 is a new ad number",
     "Ad is named 'HikeBike 2'; Q3's Mountain ad was 'HikeBike 1'",
     "Modify HikeBike 1 in place if the sim allows it",
     "REVIEW",
     "HikeBike 1 was organic Mountain #2 (115 clicks). If a new ad number spawns a "
     "new web page, we may reset page tenure and pay a second design fee. Speed SERP "
     "rank is already suspected to reflect tenure (AndStill, ad 78, sits 7th)."),

    ("Social Media",
     "Social Networking: setup 8,000 + quarterly 3,500 = 11,500",
     "~10,000 quarterly was the earlier recommendation",
     "OPTIONAL",
     "3,500 is defensible during the ramp - the sim warns demand response is slow and "
     "spend mostly buys team learning. The 8,000 setup was not in the earlier estimate. "
     "Consider 6,000-8,000 quarterly now that R&D came in cheaper than modelled."),

    ("Price and Priority / Available for Sale",
     "'Available for Sale' column appears blank for both brands",
     "Both brands must be flagged available in World Market",
     "VERIFY",
     "Brand Production shows both as 'produced', so this is probably a rendering "
     "artefact - but a brand built and not offered for sale sells zero."),

    # ---- CORRECT --------------------------------------------------------
    ("Modify Brand / Swifter Bike",
     "Aero frame, standard carbon, racing tires, precision brakes, basic drop-down, "
     "14 speed (2x7), polymer gel racing seat, reflectors, colourful thin brushstroke "
     "decals, standard lights",
     "The 77 recipe exactly",
     "CORRECT",
     "24sp -> 14sp is +3 and decals are +2. This is 72 -> 77, tying the industry "
     "ceiling (MACH I.I, LiteSpeed Pro+). Do not touch it again."),

    ("Modify Brand / Hike Bike",
     "Not listed as modified",
     "Leave Hike Bike alone - 73 of 73, verified optimal",
     "CORRECT",
     "Tied for best Mountain brand in the industry. Any change is downside."),

    ("Modify Ad / benefit stacks",
     "HikeBike 2: gears, high tread, highest rated Mountain, local sales, brand, "
     "steep-trail picture. Swifter: road-race picture, brand, wind-cheater, elite "
     "look, 3D printing, racing tires, local sales",
     "BB LLC's Mountain stack and AndStill's Speed stack",
     "CORRECT",
     "Lights, reflectors and 'great price' are gone from the Speed ad - those were "
     "the Recreation cues that held it to 70, 8th of 9."),

    ("Hire Sales People",
     "Rio 1/0/3/3 = 7, Amsterdam 1/0/3/3 = 7, 11 -> 14 heads, 92,897",
     "Clone BB LLC's identical-in-both-cities store template",
     "CORRECT",
     "Reconciles exactly: 14 x 24,425/4 = 85,487 salaries + 5,400 training = 90,888. "
     "Speed specialists go 3 -> 6, matching BB LLC and LiteCycle. No untrained staff left."),

    ("Hire Web Sales People",
     "5 web sales + 2 web support = 7; 3 -> 7 heads, 45,501",
     "BB LLC's exact 5+2 web split at the 7-head cap",
     "CORRECT",
     "Only pays off if the web productivity tactics are actually started."),

    ("Manufacturing / Fixed Capacity",
     "Planned increase 16/day; 24 -> 40 available next quarter",
     "Two printers for Q5 NYC entry",
     "CORRECT",
     "Ordered in Q4, live in Q5 - the right sequencing for a Q5 NYC store. Does not "
     "raise Q4 capacity, which is why OC must be 24 now."),

    ("Feature R&D",
     "11 speed 568,514 + decals 109,155 = 677,669; both available Q5",
     "Two rapid projects, third slot deliberately empty",
     "CORRECT",
     "Rapid premium is now KNOWN: 2-quarter prices are 478,062 and 91,788, so rapid "
     "costs 107,819 (18.9%) more and lands Q5 instead of Q6. Earlier planning assumed "
     "a ~203,301 premium on top of 677,669 - that was wrong, so cash is ~203k better "
     "than modelled."),

    ("Search Engine Marketing (paid)",
     "Mountain bid 2.40 / budget 150; Speed bid 1.95 / budget 80; total 230",
     "Match the posted industry average bid, throttle volume with the budget",
     "CORRECT",
     "Matches Q4_Paid_SEM exactly. Bid stays in the auction all quarter; the small "
     "budget caps us near 104-138 clicks, which is all the printers can absorb."),

    ("Manage Web Pages",
     "HikeBike 2 -> Mountain, Swifter Bike 1 -> Speed, 1,000 each = 2,000",
     "Two pages, one per targeted segment, no Recreation page",
     "CORRECT",
     "Two pages in one segment cannibalise - Bike Bros ran 2 Mountain ads for 79 "
     "clicks against our single page's 115."),

    ("Open Store / Open Webcenter",
     "Rio + Amsterdam operational (61,000); World Market webcenter (60,000)",
     "No third store in Q4 - depth before breadth",
     "CORRECT",
     "NYC needs the printers that only arrive in Q5, plus 7 new hires. Adding it now "
     "would generate demand against a 1,139-unit ceiling."),

    ("Price and Priority / the rest",
     "Hike Bike 1,365, both rebates 0, priority Hike 1 / Swift 2",
     "Hold Hike at 1,365, no rebates, priority unchanged",
     "CORRECT",
     "1,365 is the highest tier that still scores 100 on Mountain price judgment. "
     "Only the two cheapest Recreation brands use rebates."),

    ("Overtime",
     "Overtime table present; no hours selected",
     "Overtime 0.00",
     "CORRECT",
     "Manufacturing Productivity = utilisation - (OT/OC / 2). 24 scheduled with 0 OT "
     "= 1.000; BB LLC's 6.90 OT drags them to 0.856. This is the one place we beat them."),

    ("Buy Market Research",
     "20,000",
     "Not previously sized",
     "CORRECT",
     "On-strategy in a closed system where competitor intelligence is the only route "
     "to the numbers this whole plan is built from."),

    ("Corporate / Goals and Strategy",
     "Mountain > Speed > Recreation; six strategic directions",
     "Unchanged from the locked directions",
     "CORRECT",
     "Matches the ratified team goals."),
]

# --------------------------------------------------- demand projection block
DEMAND = [
    ("Stores", 14, 50.5, "Q3 was 46.09/head on 11 untrained heads. Anchors: Rio "
     "store-only measured 60/head, industry average excluding us 61.1, BB LLC 67. "
     "50.5 is deliberately below all three to absorb the 16.3% Q4 ill will."),
    ("Web Sales Center", 7, 62.1, "Q3 was 40.00/head with one tactic funded. 62.1 is "
     "the FLOOR of the all-tactics cohort (LiteCycle and MILC both 435 units), not "
     "BB LLC's 81.3. Conditional on starting the tactics."),
]

BRAND_SPLIT = [
    ("Hike Bike", 0.60, "Q3 was 65.7%. Ad 80 -> ~82 and Biking inserts 6 -> 12 help, "
     "but Mountain is the smallest segment (1,621) and we already hold 22.1%."),
    ("Swifter Bike", 0.40, "Q3 was 34.3%. Brand 72 -> 77, ad 70 -> 77, Speed "
     "specialists 3 -> 6, price +130. Speed is the largest segment (2,878) and we hold "
     "only 7.5%, so the headroom is here. Judgment call, not a measured figure."),
]

# ----------------------------------------------------------- capacity ladder
LADDER = [
    ("Screen as entered", 8, PROD_73),
    ("Plan of record", OWNED_OC, PROD_73),
    ("Plan + production comp fix", OWNED_OC, PROD_75),
]

# --------------------------------------------------------------- cash uses
CASH = [
    ("Feature R&D (2 rapid projects)", 677669, "entered"),
    ("Sales and service people", 92897, "entered"),
    ("Media inserts (list price, 20 inserts)", 102500, "entered, total row unclear"),
    ("Web personnel", 45501, "entered"),
    ("Store operations (Rio + Amsterdam)", 61000, "entered"),
    ("Web sales center", 60000, "entered"),
    ("Market research", 20000, "entered"),
    ("Social media (8,000 setup + 3,500)", 11500, "entered"),
    ("Web productivity: toll-free", 6000, "entered"),
    ("Web pages", 2000, "entered"),
    ("Paid SEM", 230, "entered"),
    ("Web tactics setup (4 new)", 48000, "MISSING - Start not checked"),
    ("Web tactics quarterly (4 new)", 25000, "MISSING - Start not checked"),
    ("Sales comp uplift to BB LLC package", 15634, "MISSING"),
    ("Production comp uplift to BB LLC package", 12000, "MISSING (estimate)"),
    ("Operating capacity change expense 8 -> 24", 0, "UNKNOWN - screen shows 0 at OC 8"),
    ("Two printers (fixed capacity +16)", 480000, "committed, not priced on screen"),
]


# ------------------------------------------------------------------ writer
def main():
    wb = load_workbook(XLSX)
    if SHEET in wb.sheetnames:
        del wb[SHEET]
    ws = wb.create_sheet(SHEET)

    H = Font(bold=True, size=14)
    SUB = Font(bold=True, size=11)
    TH = Font(bold=True, color="FFFFFF")
    HDR = PatternFill("solid", fgColor="1F3864")
    FILLS = {
        "BLOCKER": PatternFill("solid", fgColor="FF7B7B"),
        "MISS": PatternFill("solid", fgColor="FFC000"),
        "REVIEW": PatternFill("solid", fgColor="FFE699"),
        "OPTIONAL": PatternFill("solid", fgColor="DDEBF7"),
        "VERIFY": PatternFill("solid", fgColor="DDEBF7"),
        "CORRECT": PatternFill("solid", fgColor="C6EFCE"),
    }
    wrap = Alignment(wrap_text=True, vertical="top")

    r = 1
    ws.cell(r, 1, "Q4 decision screen audit - dump vs plan of record").font = H
    r += 1
    n_block = sum(1 for a in AUDIT if a[3] == "BLOCKER")
    n_miss = sum(1 for a in AUDIT if a[3] == "MISS")
    n_ok = sum(1 for a in AUDIT if a[3] == "CORRECT")
    ws.cell(r, 1, f"{n_block} blockers, {n_miss} misses, {n_ok} items correct. "
                  "The brand, ads, staffing, R&D and printer decisions are all in "
                  "correctly; the gaps are capacity, financing and the web tactic "
                  "Start checkboxes.")
    r += 2

    # ---- audit table
    ws.cell(r, 1, "Line-by-line").font = SUB
    r += 1
    for c, h in enumerate(["Screen area", "What the screen shows", "What the plan says",
                           "Verdict", "Consequence"], start=1):
        cell = ws.cell(r, c, h)
        cell.font, cell.fill, cell.alignment = TH, HDR, wrap
    r += 1
    order = {"BLOCKER": 0, "MISS": 1, "REVIEW": 2, "VERIFY": 3, "OPTIONAL": 4, "CORRECT": 5}
    for area, shows, plan, verdict, why in sorted(AUDIT, key=lambda a: order[a[3]]):
        for c, v in enumerate([area, shows, plan, verdict, why], start=1):
            cell = ws.cell(r, c, v)
            cell.alignment = wrap
            if c == 4:
                cell.fill = FILLS[verdict]
                cell.font = Font(bold=True)
        r += 1
    r += 1

    # ---- capacity ladder
    ws.cell(r, 1, "Capacity ladder (units = OC x 65 days x worker productivity)").font = SUB
    r += 1
    for c, h in enumerate(["Scenario", "OC/day", "Productivity", "Units/qtr",
                           "vs 1,142 demand"], start=1):
        cell = ws.cell(r, c, h)
        cell.font, cell.fill, cell.alignment = TH, HDR, wrap
    r += 1
    for name, oc, prod in LADDER:
        u = cap(oc, prod)
        for c, v in enumerate([name, oc, f"{prod:.0%}", u, u - 1142], start=1):
            ws.cell(r, c, v).alignment = wrap
        r += 1
    r += 1

    # ---- demand projection
    ws.cell(r, 1, "Demand Projection - what to type into the zeros").font = SUB
    r += 1
    for c, h in enumerate(["Channel", "People", "Demand/person", "Total demand",
                           "Basis"], start=1):
        cell = ws.cell(r, c, h)
        cell.font, cell.fill, cell.alignment = TH, HDR, wrap
    r += 1
    total = 0
    for chan, people, per, basis in DEMAND:
        t = round(people * per)
        total += t
        for c, v in enumerate([chan, people, per, t, basis], start=1):
            ws.cell(r, c, v).alignment = wrap
        r += 1
    ws.cell(r, 1, "TOTAL").font = Font(bold=True)
    ws.cell(r, 2, 21).font = Font(bold=True)
    ws.cell(r, 4, total).font = Font(bold=True)
    ws.cell(r, 5, f"Required OC = {total} / (0.73 x 65) = "
                  f"{total/(PROD_73*DAYS):.2f}/day -> schedule 24").alignment = wrap
    r += 2

    ws.cell(r, 1, "Brand split of projected demand").font = SUB
    r += 1
    for c, h in enumerate(["Brand", "Share", "Units", "Basis"], start=1):
        cell = ws.cell(r, c, h)
        cell.font, cell.fill, cell.alignment = TH, HDR, wrap
    r += 1
    for brand, share, basis in BRAND_SPLIT:
        for c, v in enumerate([brand, f"{share:.0%}", round(total * share), basis], start=1):
            ws.cell(r, c, v).alignment = wrap
        r += 1
    r += 1

    # ---- cash
    ws.cell(r, 1, "Cash uses visible on the screen").font = SUB
    r += 1
    for c, h in enumerate(["Use", "Amount", "Status"], start=1):
        cell = ws.cell(r, c, h)
        cell.font, cell.fill, cell.alignment = TH, HDR, wrap
    r += 1
    entered = missing = 0
    for use, amt, status in CASH:
        for c, v in enumerate([use, amt, status], start=1):
            cell = ws.cell(r, c, v)
            cell.alignment = wrap
            if c == 2:
                cell.number_format = "#,##0"
            if c == 3 and status.startswith("MISSING"):
                cell.fill = FILLS["MISS"]
        if status == "entered":
            entered += amt
        elif status.startswith("MISSING"):
            missing += amt
        r += 1
    ws.cell(r, 1, "Entered subtotal").font = Font(bold=True)
    ws.cell(r, 2, entered).number_format = "#,##0"
    r += 1
    ws.cell(r, 1, "Still to add").font = Font(bold=True)
    ws.cell(r, 2, missing).number_format = "#,##0"
    r += 1
    ws.cell(r, 1, "Cash on hand entering Q4").font = Font(bold=True)
    ws.cell(r, 2, 1010838).number_format = "#,##0"
    r += 1
    ws.cell(r, 1, "Shortfall before the loan (incl. printers, excl. 300k floor)")
    ws.cell(r, 2, 1010838 - entered - missing - 480000 - 102500 + 102500).number_format = "#,##0"
    r += 2

    ws.cell(r, 1, "Note: rapid R&D premium is 107,819 (18.9%), not the ~203,301 used "
                  "in earlier cash planning. 677,669 IS the rapid price. Cash is "
                  "roughly 203k better than modelled, which comfortably funds the "
                  "73,000 of web tactics and the two comp uplifts.").alignment = wrap

    for col, w in zip("ABCDE", [38, 52, 42, 11, 72]):
        ws.column_dimensions[col].width = w

    wb.save(XLSX)
    print(f"Wrote {SHEET}")
    print(f"blockers={n_block} misses={n_miss} correct={n_ok}")
    for name, oc, prod in LADDER:
        print(f"  {name}: OC {oc} @ {prod:.0%} = {cap(oc, prod)} units")
    print(f"  projected demand = {total}")
    print(f"  entered={entered:,}  missing={missing:,}")


if __name__ == "__main__":
    main()
