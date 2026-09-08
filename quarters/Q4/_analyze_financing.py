"""Q4 financing DECIDED: real loan terms (7% annual) vs the $2.5M VC issue already on the screen."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DST = "quarters/Q3/Q3Data.xlsx"

CASH = 1_010_838
SHARES = 25_000
PAR = 100
PAID_IN = SHARES * PAR
RE = -829_162
FLOOR = 300_000
WEALTH_0 = 0.668
FP_0 = 5.319
FP_NUM = FP_0 * SHARES          # 132,975
QRATE = 0.0175                  # ACTUAL: 7.00% annual / 1.75% quarterly
CAPACITY = 2_506_258

print("=" * 98)
print("STEP 1 — THE DEBT CAPACITY FORMULA IS THE WEALTH NUMERATOR x 1.5")
print("=" * 98)
net_eq = PAID_IN + RE
print(f"  Reported debt capacity          ${CAPACITY:,}")
print(f"  1.5 x (Equity + Retained Earnings)")
print(f"    = 1.5 x (${PAID_IN:,} + ${RE:,})")
print(f"    = 1.5 x ${net_eq:,}")
print(f"    = ${net_eq*1.5:,.0f}   -> matches to the dollar")
print(f"\n  So debt capacity = 1.5 x NET EQUITY, and net equity is exactly the numerator of the")
print(f"  Wealth ratio we just solved. The two decisions are linked:")
print(f"    - every $1 of equity issued adds $1.50 of borrowing room")
print(f"    - every $1 of loss destroys $1.50 of borrowing room")
print(f"  Wealth and debt capacity are the same quantity wearing two hats.")

print("\n" + "=" * 98)
print("STEP 2 — THE $2.5M VC ISSUE IS ALREADY ENTERED AS A LIVE Q4 DECISION")
print("=" * 98)
print("""  Stock history rows:
    Q1  Executive Team        15,000 @ $100 = $1,500,000
    Q2  Executive Team         5,000 @ $100 =   $500,000
    Q3  Executive Team         5,000 @ $100 =   $500,000
    Q4  VENTURE CAPITALISTS   25,000 @ $100 = $2,500,000   <- CURRENTLY ENTERED

  The first three rows sum to 25,000 shares and $2,500,000 of paid-in capital, which is exactly
  what the Wealth reconciliation required. Good — the model is confirmed.

  But the Q4 row is a DECISION SITTING IN THE FORM, not history. As entered it would:
    - DOUBLE the share count, 25,000 -> 50,000
    - hand Venture Capitalists 25,000 of 50,000 shares = 50% of the company
    - raise $2,500,000 against a Q4 need of ~$1.6M, leaving the balance idle
  This must be reduced before submitting.""")

print("\n" + "=" * 98)
print("STEP 3 — USES (team voted TWO printers)")
print("=" * 98)
USES = [
    ("R&D — decals + 11-speed, listed", 677_669),
    ("R&D rapid premium (~30% ESTIMATE)", 203_301),
    ("Operating plan, recurring", 146_000),
    ("Operating plan, one-time", 67_000),
    ("Swift Bike brand redesign fee", 30_000),
    ("Swift Bike ad redesign fee", 10_000),
    ("Printer #4", 240_000),
    ("Printer #5", 240_000),
]
total = sum(u[1] for u in USES)
for n, a in USES:
    print(f"  {n:<44}{a:>12,}")
print(f"  {'-'*56}")
print(f"  {'TOTAL USES':<44}{total:>12,}")
print(f"  {'Cash on hand':<44}{CASH:>12,}")
print(f"  {'Hard floor':<44}{-FLOOR:>12,}")
print(f"  {'Deployable':<44}{CASH-FLOOR:>12,}")
gap = total - (CASH - FLOOR)
print(f"  {'FUNDING GAP':<44}{gap:>12,}")

print("\n" + "=" * 98)
print("STEP 4 — HEAD TO HEAD AT THE REAL 7% RATE (was assuming 10%)")
print("=" * 98)


def debt(amount):
    i = amount * QRATE
    w = (PAID_IN + RE - i) / PAID_IN
    fp = (FP_NUM - i / 2) / SHARES
    return w, fp, i


def equity(amount):
    sh = SHARES + round(amount / PAR)
    p = sh * PAR
    return (p + RE) / p, FP_NUM / sh, 0


print(f"  Funding the ${gap:,} gap:\n")
print(f"  {'':<24}{'ALL DEBT':>14}{'ALL EQUITY':>14}{'$2.5M VC (as entered)':>24}")
wd, fpd, i_d = debt(gap)
we, fpe, _ = equity(gap)
wv, fpv, _ = equity(2_500_000)
print(f"  {'Shares after':<24}{SHARES:>14,}{SHARES+round(gap/PAR):>14,}{50_000:>24,}")
print(f"  {'Quarterly interest':<24}{i_d:>14,.0f}{0:>14,}{0:>24,}")
print(f"  {'Wealth':<24}{wd:>14.3f}{we:>14.3f}{wv:>24.3f}")
print(f"  {'  vs 0.668':<24}{wd/WEALTH_0-1:>13.1%}{we/WEALTH_0-1:>13.1%}{wv/WEALTH_0-1:>23.1%}")
print(f"  {'Fin Performance':<24}{fpd:>14.3f}{fpe:>14.3f}{fpv:>24.3f}")
print(f"  {'  vs 5.319':<24}{fpd/FP_0-1:>13.1%}{fpe/FP_0-1:>13.1%}{fpv/FP_0-1:>23.1%}")
cd = (wd / WEALTH_0) * (fpd / FP_0)
ce = (we / WEALTH_0) * (fpe / FP_0)
cv = (wv / WEALTH_0) * (fpv / FP_0)
print(f"  {'COMBINED (BSC product)':<24}{cd-1:>13.1%}{ce-1:>13.1%}{cv-1:>23.1%}")
print(f"\n  DEBT WINS DECISIVELY at 7%. It is {abs(ce-1)/abs(cd-1):.1f}x less damaging than equity")
print(f"  and {abs(cv-1)/abs(cd-1):.1f}x less damaging than the $2.5M issue currently in the form.")
print(f"\n  Why: ${i_d:,.0f} of quarterly interest is a small, one-off drag. Doubling the share count")
print(f"  is a PERMANENT halving of Financial Performance — a metric where we are already last.")
print(f"  Note the FP numerator is held flat here; the point is structural, since shares are the")
print(f"  denominator and dilution scales the metric no matter how large the numerator grows.")

print("\n" + "=" * 98)
print("STEP 5 — IS THE INTEREST AFFORDABLE? AND IS THE PRINTER WORTH BORROWING FOR?")
print("=" * 98)
gm = 1_154 * 800
print(f"  Q4 gross margin at plan (~1,154 units x ~$800)      ${gm:>12,}")
print(f"  Quarterly interest on ${gap:,}                       ${gap*QRATE:>12,.0f}")
print(f"  Interest as a share of gross margin                 {gap*QRATE/gm:>12.1%}")
u = 8 * 0.74 * 65
print(f"\n  Two printers = {2*u:.0f} units/quarter x ~$800 = ${2*u*800:,.0f} contribution")
print(f"  Cost $480,000, financed at {QRATE:.2%}/qtr = ${480_000*QRATE:,.0f} interest per quarter")
print(f"  -> the printers out-earn their own interest {2*u*800/(480_000*QRATE):.0f}x over.")
print(f"  Borrowing at 1.75% a quarter to buy an asset returning {2*u*800/480_000:.0%} a quarter is")
print(f"  as close to free money as this simulation will ever offer.")

print("\n" + "=" * 98)
print("STEP 6 — THE ONE RISK: LOSSES SHRINK DEBT CAPACITY")
print("=" * 98)
print("  Capacity = 1.5 x (paid-in + RE), so a Q4 loss cuts next quarter's borrowing room by 1.5x.")
print("  If R&D is EXPENSED, Q4 posts a large loss. Where does that leave us?\n")
print(f"  {'Q4 net loss':>14}{'New RE':>14}{'New capacity':>15}{'Borrowed':>12}{'Headroom':>12}")
BORROW = 950_000
for loss in (0, 300_000, 500_000, 760_000, 1_000_000, 1_200_000):
    nre = RE - loss
    cap = 1.5 * (PAID_IN + nre)
    hr = cap - BORROW
    flag = "   BREACH" if hr < 0 else ""
    print(f"  {loss:>14,}{nre:>14,}{cap:>15,.0f}{BORROW:>12,}{hr:>12,.0f}{flag}")
print(f"\n  Borrowing ${BORROW:,} stays inside capacity until the Q4 loss exceeds about")
print(f"  ${(BORROW/1.5)-PAID_IN-RE:+,.0f} ... i.e. a loss beyond ${abs((BORROW/1.5)-PAID_IN-RE):,.0f}.")
print("\n  Plausibility: at plan (~1,154 units) gross margin is ~$920k against roughly $800k of")
print("  operating expense, so operations run near break-even BEFORE R&D. Expensing ~$880k of R&D")
print("  would put the Q4 loss near $760k — still inside capacity. It only breaches if the volume")
print("  plan ALSO underdelivers badly. Real but not likely.")
print("\n  MITIGATION IF THE TEAM WANTS INSURANCE: a small equity slice lifts the capacity base at")
print("  1.5:1. $250,000 of equity buys $375,000 of extra borrowing room and costs only 2,500")
print("  shares. That is the cheapest hedge available — but it is a hedge, not the base case.")

print("\n" + "=" * 98)
print("STEP 7 — RECOMMENDATION")
print("=" * 98)
print(f"""
  1. SET THE STOCK ISSUE TO $0. Delete the Q4 Venture Capitalists row. As entered it doubles
     the share count, gives away half the company, halves Financial Performance permanently,
     and raises $900k more than we can deploy — recreating the idle-cash problem that put
     Asset Management at 0.353, last in the industry.

  2. BORROW $950,000 as a conventional bank loan. Covers the ${gap:,} gap with a small
     buffer for the unknown rapid-R&D premium. Interest ${950_000*QRATE:,.0f}/quarter, which is
     {950_000*QRATE/gm:.1%} of planned gross margin and {1.5*(PAID_IN+RE)/950_000:.1f}x inside our
     ${CAPACITY:,} capacity.

  3. BUY BOTH PRINTERS ($480,000). They arrive Q5, which is exactly when New York opens.

  4. Wealth stays ~0.66 and remains our weakest link. The ONLY real fix is retained earnings —
     i.e. actually making money — not financial engineering. Equity would paper over it by
     diluting the deficit while wrecking Financial Performance to do so.

  STILL WORTH CONFIRMING: whether R&D is expensed or capitalised in Q4. It does not change the
  debt-vs-equity answer (debt wins at 7% either way), but it determines how much Q4 headroom
  we have left for Q5 and whether the $250k equity hedge is worth taking.
""")
print("=" * 98)

# ------------------------------------------------------------------ workbook
wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
good = PatternFill("solid", fgColor="C6EFCE")
bad = PatternFill("solid", fgColor="FFC7CE")
yel = PatternFill("solid", fgColor="FFEB9C")
thin = Border(*[Side(style="thin", color="D0D0D0")] * 4)

for s in ("Q4_Financing",):
    if s in wb.sheetnames:
        del wb[s]
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


ws["A1"] = "Q4 financing — DECIDED: borrow $950,000, issue no stock"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = ("Two formulas now reconcile EXACTLY. Wealth = (paid-in + RE) / paid-in = 0.6683 vs reported 0.668. "
            "Debt capacity = 1.5 x (paid-in + RE) = $2,506,257 vs reported $2,506,258. Net equity is the Wealth "
            "numerator AND the debt-capacity base, so every $1 of equity buys $1.50 of borrowing room and every "
            "$1 of loss destroys $1.50 of it. At the actual 7% annual rate, DEBT IS 2.5x LESS DAMAGING TO THE "
            "SCORECARD THAN EQUITY, and 4x less damaging than the $2.5M VC issue currently sitting in the form.")
ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A2:F2")
ws.row_dimensions[2].height = 62

r = 4
rws = [["LOAN TERMS (as offered)", "Value"]]
for a, b in [("Debt capacity = 1.5 x (Equity + RE)", CAPACITY), ("Total debt to date", 0),
             ("Available to borrow", CAPACITY), ("Annual interest rate", "7.00%"),
             ("Quarterly interest rate", "1.75%")]:
    rws.append([a, b])
r2 = block(r, rws)
ws.cell(r + 1, 2).number_format = "#,##0"
ws.cell(r + 3, 2).number_format = "#,##0"
r = r2

rws = [["STOCK HISTORY", "Owner", "Shares", "Price", "Total", "Quarter"]]
for o, s, p, t, q in [("Common Stock", "Executive Team", 15_000, 100, 1), ("Common Stock", "Executive Team", 5_000, 100, 2),
                      ("Common Stock", "Executive Team", 5_000, 100, 3),
                      ("Common Stock", "VENTURE CAPITALISTS", 25_000, 100, 4)]:
    rws.append([o, s, p, 100, p * 100, t if False else q])
rws = [rws[0]] + [["Common Stock", "Executive Team", 15_000, 100, 1_500_000, 1],
                  ["Common Stock", "Executive Team", 5_000, 100, 500_000, 2],
                  ["Common Stock", "Executive Team", 5_000, 100, 500_000, 3],
                  ["Common Stock", "VENTURE CAPITALISTS — CURRENTLY ENTERED, REDUCE TO $0", 25_000, 100, 2_500_000, 4]]
r2 = block(r, rws)
for i in range(1, len(rws)):
    for c in (3, 5):
        ws.cell(r + i, c).number_format = "#,##0"
for c in range(1, 7):
    ws.cell(r + 4, c).fill = bad
    ws.cell(r + 4, c).font = bold
r = r2
ws.cell(r - 1, 1, "Rows 1-3 sum to 25,000 shares and $2,500,000 paid-in — exactly what the Wealth reconciliation "
                  "required, so the model is confirmed. Row 4 is a LIVE DECISION, not history: as entered it doubles "
                  "the share count and hands VCs 50% of the company.")
ws.cell(r - 1, 1).alignment = Alignment(wrap_text=True)

rws = [["USES — Q4 (team voted 2 printers)", "Amount"]]
for n, a in USES:
    rws.append([n, a])
rws += [["TOTAL USES", total], ["Cash on hand", CASH], ["Hard floor", FLOOR],
        ["Deployable cash", CASH - FLOOR], ["FUNDING GAP", gap], ["RECOMMENDED BORROWING", 950_000]]
r2 = block(r, rws)
for i in range(1, len(rws)):
    ws.cell(r + i, 2).number_format = "#,##0"
for c in (1, 2):
    ws.cell(r + len(rws) - 1, c).fill = good
    ws.cell(r + len(rws) - 1, c).font = bold
r = r2

rws = [["OPTION", "Shares after", "Qtrly interest", "Wealth", "vs 0.668", "Fin Perf", "vs 5.319", "BSC combined"]]
for label, (w, fp, ii), sh in [("ALL DEBT (recommended)", debt(gap), SHARES),
                               ("ALL EQUITY", equity(gap), SHARES + round(gap / PAR)),
                               ("$2.5M VC as entered", equity(2_500_000), 50_000)]:
    rws.append([label, sh, round(ii), round(w, 3), w / WEALTH_0 - 1, round(fp, 3), fp / FP_0 - 1,
                (w / WEALTH_0) * (fp / FP_0) - 1])
r2 = block(r, rws)
for i in range(1, len(rws)):
    for c in (2, 3):
        ws.cell(r + i, c).number_format = "#,##0"
    for c in (5, 7, 8):
        ws.cell(r + i, c).number_format = "0.0%"
for c in range(1, 9):
    ws.cell(r + 1, c).fill = good
    ws.cell(r + 3, c).fill = bad
r = r2

rws = [["RISK — losses shrink debt capacity 1.5x", "New RE", "New capacity", "Borrowed $950k", "Headroom"]]
for loss in (0, 300_000, 500_000, 760_000, 1_000_000, 1_200_000):
    nre = RE - loss
    cap = 1.5 * (PAID_IN + nre)
    rws.append([f"Q4 net loss of ${loss:,}", nre, round(cap), 950_000, round(cap - 950_000)])
r2 = block(r, rws)
for i in range(1, len(rws)):
    for c in (2, 3, 4, 5):
        ws.cell(r + i, c).number_format = "#,##0"
    if rws[i][4] < 0:
        for c in range(1, 6):
            ws.cell(r + i, c).fill = bad
r = r2
ws.cell(r - 1, 1, "At plan (~1,154 units) gross margin ~$920k against ~$800k operating expense, so operations run "
                  "near break-even BEFORE R&D. Expensing ~$880k of R&D puts the Q4 loss near $760k — still inside "
                  "capacity. It only breaches if the volume plan ALSO underdelivers badly. HEDGE IF WANTED: "
                  "$250,000 of equity buys $375,000 of extra capacity for only 2,500 shares.")
ws.cell(r - 1, 1).alignment = Alignment(wrap_text=True)

rws = [["DECISION", "Detail"]]
for a, b in [
    ("1. Stock issue -> $0", "Delete the Q4 Venture Capitalists row. As entered it doubles shares 25,000 -> 50,000, "
     "gives VCs half the company, permanently halves Financial Performance where we are already last, and raises "
     "$900k more than we can deploy — recreating the idle-cash problem behind Asset Management 0.353."),
    ("2. Borrow $950,000", f"Covers the ${gap:,} gap plus a buffer for the unknown rapid-R&D premium. Interest "
     f"${950_000*QRATE:,.0f}/quarter = {950_000*QRATE/gm:.1%} of planned gross margin, and "
     f"{CAPACITY/950_000:.1f}x inside the ${CAPACITY:,} capacity."),
    ("3. Buy both printers ($480,000)", "They arrive Q5, which is exactly when New York opens. Two printers earn "
     f"${2*u*800:,.0f}/quarter against ${480_000*QRATE:,.0f} of interest — they out-earn their own financing "
     f"{2*u*800/(480_000*QRATE):.0f}x."),
    ("4. Accept that Wealth stays ~0.66", "It remains our weakest indicator. The only real fix is retained "
     "earnings — actually making money. Equity would paper over it by diluting the deficit while wrecking "
     "Financial Performance to do so."),
    ("Still worth confirming", "Whether R&D is expensed or capitalised in Q4. It does not change the debt-vs-equity "
     "answer, but it sets how much headroom Q5 has and whether the $250k equity hedge is worth taking."),
]:
    rws.append([a, b])
r2 = block(r, rws, wrap=(2,))
for i in (1, 2, 3):
    ws.cell(r + i, 2).fill = good
ws.cell(r + 5, 2).fill = yel

ws.column_dimensions["A"].width = 46
ws.column_dimensions["B"].width = 16
ws.column_dimensions["C"].width = 96
for c in "DEFGH":
    ws.column_dimensions[c].width = 14

wb.save(DST)
print("\nRewrote Q4_Financing in Q3Data.xlsx")
