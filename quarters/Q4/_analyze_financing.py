"""Q4 financing: loan vs equity, sized against actual sources and uses."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DST = "quarters/Q3/Q3Data.xlsx"

CASH = 1_010_838
SHARES = 25_000
PAR = 100
PAID_IN = SHARES * PAR          # 2,500,000
RE = -829_162
FLOOR = 300_000
WEALTH_REPORTED = 0.668
FP_REPORTED = 5.319

print("=" * 96)
print("STEP 1 — THE WEALTH FORMULA RECONCILES EXACTLY. My earlier claim was backwards.")
print("=" * 96)
w = (PAID_IN + RE) / PAID_IN
print(f"  Wealth = Net Equity / Total Stockholders Equity")
print(f"         = (paid-in capital + retained earnings) / paid-in capital")
print(f"         = ({PAID_IN:,} + {RE:,}) / {PAID_IN:,}")
print(f"         = {PAID_IN + RE:,} / {PAID_IN:,}")
print(f"         = {w:.4f}   vs REPORTED {WEALTH_REPORTED}   -> EXACT")
print(f"\n  Two independent facts produce it: {SHARES:,} shares and the ${PAR}/share Q1 price give")
print(f"  paid-in capital of exactly ${PAID_IN:,}. So Wealth = 1 + (RE / paid-in).")
print("\n  CONSEQUENCE: while retained earnings are NEGATIVE, issuing equity RAISES Wealth, because")
print("  it dilutes the accumulated deficit against a larger capital base. I previously told you")
print("  issuing shares would drag Wealth. That is wrong, and it reverses part of the argument.")

print("\n  Wealth at each raise size (holding RE constant):")
print(f"    {'Raise':>12}{'New paid-in':>14}{'Wealth':>10}{'vs 0.668':>11}")
for r in (0, 500_000, 750_000, 1_000_000, 2_500_000):
    p = PAID_IN + r
    nw = (p + RE) / p
    print(f"    {r:>12,}{p:>14,}{nw:>10.3f}{(nw/WEALTH_REPORTED-1):>10.0%}")
print("\n  A loan raises no paid-in capital, so Wealth stays 0.668 — dead last, unchanged.")

print("\n" + "=" * 96)
print("STEP 2 — BUT FINANCIAL PERFORMANCE PUNISHES SHARES DIRECTLY")
print("=" * 96)
num = FP_REPORTED * SHARES
print(f"  Financial Performance = ((Net Profit from Current Ops + Gross Profit) / 2) / Total Shares")
print(f"  Implied Q3 numerator = {FP_REPORTED} x {SHARES:,} = ${num:,.0f}")
print(f"\n  Caveat: that numerator does NOT tie exactly to brand-level gross profit")
print(f"  (Hike $207,729 + Swift $109,651 = $317,380, which would imply net profit of -$51,430")
print(f"  against a reported net income of -$111,431 — a $60,001 gap we have not explained).")
print(f"  The STRUCTURE is what matters here: shares are the denominator.")
print(f"\n  Holding the numerator flat, equity alone does this:")
print(f"    {'Raise':>12}{'Shares':>10}{'Fin Perf':>11}{'vs 5.319':>11}")
for r in (0, 500_000, 750_000, 1_000_000, 2_500_000):
    sh = SHARES + r // PAR
    fp = num / sh
    print(f"    {r:>12,}{sh:>10,}{fp:>11.3f}{(fp/FP_REPORTED-1):>10.0%}")

print("\n" + "=" * 96)
print("STEP 3 — HEAD TO HEAD ON THE SCORECARD (Total Performance is the PRODUCT of 9 indicators,")
print("          so what matters is the PERCENTAGE change in each)")
print("=" * 96)
RATE = 0.10  # assumed annual interest — WE DO NOT KNOW THE ACTUAL RATE
for label, r in (("$750,000", 750_000), ("$1,000,000", 1_000_000)):
    print(f"\n  {label} raise:")
    # equity
    p = PAID_IN + r
    we, fpe = (p + RE) / p, num / (SHARES + r // PAR)
    # debt: interest hits net profit -> half of it hits the FP numerator; RE falls by interest
    inter = r * RATE / 4
    wd = (PAID_IN + RE - inter) / PAID_IN
    fpd = (num - inter / 2) / SHARES
    print(f"    {'':<20}{'EQUITY':>12}{'LOAN':>12}")
    print(f"    {'Wealth':<20}{we:>12.3f}{wd:>12.3f}")
    print(f"    {'  change':<20}{(we/WEALTH_REPORTED-1):>11.0%}{(wd/WEALTH_REPORTED-1):>11.0%}")
    print(f"    {'Fin Performance':<20}{fpe:>12.3f}{fpd:>12.3f}")
    print(f"    {'  change':<20}{(fpe/FP_REPORTED-1):>11.0%}{(fpd/FP_REPORTED-1):>11.0%}")
    ce = (we / WEALTH_REPORTED) * (fpe / FP_REPORTED)
    cd = (wd / WEALTH_REPORTED) * (fpd / FP_REPORTED)
    print(f"    {'COMBINED on product':<20}{ce-1:>11.0%}{cd-1:>11.0%}   <- less bad wins")
print(f"\n  Interest assumed at {RATE:.0%}/yr. WE DO NOT KNOW THE REAL RATE — this is the one input")
print("  that could flip the comparison. At a punitive rate, debt loses.")
print("\n  Neither form helps Asset Management (0.353, last). That metric punishes cash SITTING")
print("  IDLE, so the fix is to raise only what we DEPLOY — identical either way.")

print("\n" + "=" * 96)
print("STEP 4 — WHERE THE MONEY ACTUALLY GOES (and the printer timing that resizes everything)")
print("=" * 96)
print("""
  CRITICAL SEQUENCING: printers come online the FOLLOWING quarter (our Q1 lock proves it —
  3 printers bought in Q1 gave '+24/day online next quarter, Q1 available = 0').

  So a printer bought in Q4 produces NOTHING in Q4. Which means the NYC store CANNOT open in
  Q4 — we would create 340-570 units of demand against a capacity that is already committed at
  24.35/day for Rio + the web rebuild, and earn another quarter of stock-out ill will in the one
  city where BB LLC already holds 82.8% of Mountain.

  The correct order is: Q4 BUYS the printers, Q5 OPENS New York into capacity that exists.
""")
USES = [
    ("R&D — decals + 11-speed, listed price", 677_669, "rapid scheduling costs MORE; read the real number on screen"),
    ("R&D rapid premium, if ~30%", 203_301, "ESTIMATE ONLY — the single biggest unknown in this table"),
    ("Operating plan, recurring", 146_000, "media +$27k, web tactics +$27k, web staff +$27.4k, Rio staff +$20.6k, comp +$10.4k, training +$2k"),
    ("Operating plan, one-time", 67_000, "hiring and setup estimates"),
    ("Swift Bike brand redesign fee", 30_000, "14-speed + decals; Q3 new-design fee was $30,000, a modify may be less"),
    ("Swift Bike ad redesign fee", 10_000, "estimate"),
    ("Printer #4", 240_000, "8 units/day. Online Q5. Pays back in ONE quarter (see below)"),
    ("Printer #5", 240_000, "OPTIONAL — only if we commit to New York in Q5"),
]
tot1 = sum(u[1] for u in USES[:-1])
tot2 = sum(u[1] for u in USES)
print(f"  {'Use':<44}{'Amount':>12}")
for name, amt, note in USES:
    print(f"  {name:<44}{amt:>12,}")
print(f"  {'-'*56}")
print(f"  {'TOTAL with 1 printer':<44}{tot1:>12,}")
print(f"  {'TOTAL with 2 printers':<44}{tot2:>12,}")
print(f"\n  {'Cash on hand':<44}{CASH:>12,}")
print(f"  {'Less hard floor':<44}{-FLOOR:>12,}")
print(f"  {'Deployable cash':<44}{CASH-FLOOR:>12,}")
print(f"\n  {'RAISE NEEDED — 1 printer':<44}{max(0,tot1-(CASH-FLOOR)):>12,}")
print(f"  {'RAISE NEEDED — 2 printers':<44}{max(0,tot2-(CASH-FLOOR)):>12,}")
print("\n  That is where $750,000-$1,000,000 came from: it is the TWO-printer case. One printer")
print("  needs only about $500,000. The raise size is really a decision about how hard we")
print("  commit to New York in Q5.")

print("\n" + "=" * 96)
print("STEP 5 — WHY A PRINTER IS THE ONE THING WORTH BORROWING FOR")
print("=" * 96)
PRINTER_DAY, PROD, DAYS = 8, 0.74, 65
units = PRINTER_DAY * PROD * DAYS
print(f"  One printer = {PRINTER_DAY} units/day x {PROD} productivity x {DAYS} days = {units:.0f} units/quarter")
print(f"  At our blended contribution (Hike $747 / Swift $886 at the new price), call it $800:")
print(f"    {units:.0f} units x $800 = ${units*800:,.0f} of contribution per quarter")
print(f"    against a ${240_000:,} purchase price")
print(f"  -> pays for itself in {240_000/(units*800):.2f} of a quarter, then repeats every quarter.")
print("\n  THIS is the textbook case for debt: a hard asset with a self-liquidating, measurable")
print("  payback inside one period. R&D is the opposite — speculative, no guaranteed payback, and")
print("  it should not carry a fixed repayment obligation.")

print("\n" + "=" * 96)
print("STEP 6 — RECOMMENDATION")
print("=" * 96)
print("""
  MATCH THE FUNDING TO THE ASSET.

  1. EQUITY for R&D and the operating plan (~$700k-$900k of the need). It is speculative
     spending with no contractual payback, so it belongs on equity. And the Wealth formula
     actively REWARDS it right now: our accumulated deficit is what drags Wealth to 0.668, and
     fresh paid-in capital dilutes that deficit. A $1M raise takes Wealth 0.668 -> 0.763.

  2. DEBT for the printers (~$240k-$480k), IF a loan is on the menu at a non-punitive rate. A
     printer returns its purchase price in under one quarter, so it services its own interest
     several times over, and debt keeps those dollars out of the Financial Performance
     denominator where we are already last.

  3. DO NOT take the full $2.5M in either form. Undeployed cash is precisely what put Asset
     Management at 0.353, last in the industry. Size the raise to the uses.

  WHAT I CANNOT ANSWER AND YOU NEED TO CHECK ON THE FINANCE SCREEN:
    a. Is a loan even offered this quarter? Everything we have been shown is a VC/equity round.
    b. What is the interest rate, and is there a borrowing cap tied to equity? With retained
       earnings at -$829,162 we may simply not qualify.
    c. Is the R&D spend EXPENSED in Q4 or capitalised? If expensed, Q4 posts a large loss even
       with the volume fix, retained earnings fall further, and Wealth drops with it. That
       weakens the debt case and is the number I would most want before voting.
    d. Confirm printers are still one quarter delayed. The entire New York sequence depends on it.

  If a loan is NOT available, take EQUITY of $750,000 and buy one printer. If it IS available on
  reasonable terms, split: ~$600,000 equity for R&D and operations, ~$480,000 debt for two
  printers, and open New York in Q5 with capacity already in place.
""")
print("=" * 96)

# ------------------------------------------------------------------ workbook
wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
good = PatternFill("solid", fgColor="C6EFCE")
bad = PatternFill("solid", fgColor="FFC7CE")
yel = PatternFill("solid", fgColor="FFEB9C")
thin = Border(*[Side(style="thin", color="D0D0D0")] * 4)

if "Q4_Financing" in wb.sheetnames:
    del wb["Q4_Financing"]
ws = wb.create_sheet("Q4_Financing")


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


ws["A1"] = "Q4 financing — loan vs equity, sized to actual uses"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("KEY CORRECTION: Wealth = (paid-in capital + retained earnings) / paid-in capital, which reconciles "
            "EXACTLY to the reported 0.668 from 25,000 shares at $100 and retained earnings of -$829,162. While "
            "retained earnings are negative, issuing equity RAISES Wealth by diluting the deficit. Earlier advice "
            "in this project said issuing shares would drag Wealth — that was wrong.")
ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A2:F2")
ws.row_dimensions[2].height = 48

r = 4
rws = [["Raise", "New paid-in", "Wealth (equity)", "vs 0.668", "Shares", "Fin Perf (equity)", "vs 5.319"]]
for rr in (0, 500_000, 750_000, 1_000_000, 2_500_000):
    p = PAID_IN + rr
    rws.append([rr, p, round((p + RE) / p, 3), (p + RE) / p / WEALTH_REPORTED - 1,
                SHARES + rr // PAR, round(num / (SHARES + rr // PAR), 3),
                num / (SHARES + rr // PAR) / FP_REPORTED - 1])
r2 = block(r, rws)
for i in range(1, len(rws)):
    for c in (1, 2, 5):
        ws.cell(r + i, c).number_format = "#,##0"
    for c in (4, 7):
        ws.cell(r + i, c).number_format = "0%"
    ws.cell(r + i, 3).fill = good
    ws.cell(r + i, 6).fill = bad
r = r2
ws.cell(r - 1, 1, "Equity helps Wealth and hurts Financial Performance. A loan does the opposite: Wealth stays "
                  "0.668 (last, unchanged) and shares stay at 25,000, with only interest touching the numerator.")
ws.cell(r - 1, 1).alignment = Alignment(wrap_text=True)

rws = [["SOURCES AND USES — Q4", "Amount", "Note"]]
for name, amt, note in USES:
    rws.append([name, amt, note])
rws += [["TOTAL with 1 printer", tot1, ""],
        ["TOTAL with 2 printers", tot2, ""],
        ["Cash on hand", CASH, ""],
        ["Hard floor (ending cash + CD)", FLOOR, "must not be broken"],
        ["Deployable cash", CASH - FLOOR, ""],
        ["RAISE NEEDED — 1 printer", max(0, tot1 - (CASH - FLOOR)), "~$500k"],
        ["RAISE NEEDED — 2 printers", max(0, tot2 - (CASH - FLOOR)),
         "this is where $750k-$1M came from — it is the two-printer case"]]
r2 = block(r, rws, wrap=(3,))
for i in range(1, len(rws)):
    ws.cell(r + i, 2).number_format = "#,##0"
for i in (len(rws) - 2, len(rws) - 1):
    for c in (1, 2, 3):
        ws.cell(r + i, c).fill = yel
        ws.cell(r + i, c).font = bold
r = r2

rws = [["PRINTER TIMING — resizes the whole plan", "Detail"]]
for a, b in [
    ("Printers arrive one quarter LATE", "Our Q1 lock proves it: 3 printers bought in Q1 gave '+24/day online next "
     "quarter, Q1 available = 0'. A printer bought in Q4 produces nothing in Q4."),
    ("So New York CANNOT open in Q4", "Q4 capacity is already committed at 24.35/day for Rio plus the web rebuild. "
     "Adding 340-570 NYC units against zero spare capacity earns another quarter of stock-out ill will, in the one "
     "city where BB LLC already holds 82.8% of Mountain."),
    ("Correct order", "Q4 BUYS printers and funds R&D. Q5 OPENS New York into capacity that already exists."),
    ("Printer payback", f"8 units/day x 0.74 x 65 days = {units:.0f} units/quarter. At ~$800 contribution that is "
     f"${units*800:,.0f} per quarter against a $240,000 price — it repays itself in "
     f"{240_000/(units*800):.2f} of a quarter, then repeats."),
]:
    rws.append([a, b])
r2 = block(r, rws, wrap=(2,))
ws.cell(r + 1, 2).fill = yel
r = r2

rws = [["RECOMMENDATION — match the funding to the asset", "Detail"]]
for a, b in [
    ("EQUITY for R&D + operating plan (~$700-900k)", "Speculative spending with no contractual payback belongs on "
     "equity. And Wealth actively rewards it now: the accumulated deficit is what drags us to 0.668, and fresh "
     "paid-in capital dilutes it. A $1M raise takes Wealth 0.668 -> 0.763."),
    ("DEBT for printers (~$240-480k), if offered", "A hard asset with a self-liquidating payback inside one "
     "quarter services its own interest many times over, and debt keeps those dollars out of the Financial "
     "Performance denominator where we are already last."),
    ("NOT the full $2.5M in either form", "Undeployed cash is exactly what put Asset Management at 0.353, last in "
     "the industry. Size the raise to the uses."),
    ("If no loan is available", "Take $750,000 of equity and buy one printer."),
    ("If a loan IS available on fair terms", "Split: ~$600,000 equity for R&D and operations, ~$480,000 debt for "
     "two printers, then open New York in Q5."),
]:
    rws.append([a, b])
r2 = block(r, rws, wrap=(2,))
for i in (1, 2):
    ws.cell(r + i, 2).fill = good
r = r2

rws = [["MUST CHECK BEFORE VOTING", "Why it matters"]]
for a, b in [
    ("Is a loan even offered this quarter?", "Everything the sim has shown us is a VC/equity round. The whole debt "
     "half of this analysis is moot if there is no debt window."),
    ("Interest rate and any borrowing cap", "The one input that could flip the comparison. This analysis assumes "
     "10%/yr. At a punitive rate debt loses. With retained earnings at -$829,162 we may not qualify at all."),
    ("Is R&D EXPENSED in Q4 or capitalised?", "THE NUMBER I WOULD MOST WANT. If the $677,669 is expensed, Q4 posts "
     "a large loss even with the volume fix, retained earnings fall further, and Wealth falls with them — which "
     "weakens the debt case and changes the raise size."),
    ("Confirm the one-quarter printer delay", "The entire New York sequence depends on it."),
    ("Status", "ANALYSIS ONLY — majority vote required."),
]:
    rws.append([a, b])
block(r, rws, wrap=(2,))
ws.cell(r + 3, 2).fill = yel

ws.column_dimensions["A"].width = 44
ws.column_dimensions["B"].width = 15
ws.column_dimensions["C"].width = 92
for c in "DEFG":
    ws.column_dimensions[c].width = 15

wb.save(DST)
print("\nAdded Q4_Financing to Q3Data.xlsx")
