"""Verify the corrected capacity arithmetic and the full web rebuild case."""

DAYS = 65
PROD = 0.74
FIXED = 24

print("CAPACITY FORMULA CHECK")
print(f"  max units at {FIXED}/day, {DAYS} days, {PROD:.0%} productivity = {FIXED*PROD*DAYS:.0f}")
print(f"  WRONG pad 1+(1-p) = {1+(1-PROD):.3f}   CORRECT factor 1/p = {1/PROD:.3f}")
print(f"  at 656 forecast: wrong gives {656/DAYS*(1+(1-PROD)):.1f}/day, correct is {656/(PROD*DAYS):.1f}/day\n")

print("DEMAND SCENARIOS -> REQUIRED OC/DAY")
scen = [
    ("A. Do nothing (627 less 16.3% ill will)", 525),
    ("B. Base forecast", 656),
    ("C. + Rio to 7-person cap (+200)", 856),
    ("D. + full web rebuild (+315)", 1171),
    ("E. + New York store (+400)", 1571),
]
for name, d in scen:
    oc = d / (PROD * DAYS)
    fits = "fits" if oc <= FIXED else "EXCEEDS 24/day"
    print(f"  {name:<42}{d:>5} units  OC {oc:5.1f}/day  {fits}")

print("\nFULL WEB REBUILD CASE")
extra = 435 - 120
margin = extra * 750
staff = 4 * 24425 / 4          # 4 people, one quarter
budget = 28000 - 6000
print(f"  web demand 120 -> 435 = +{extra} units")
print(f"  gross margin at $750    ${margin:,}")
print(f"  cost: +4 web staff/qtr  ${staff:,.0f}")
print(f"  cost: budget +$22,000   ${budget:,}")
print(f"  net quarterly gain      ${margin - staff - budget:,.0f}")
print(f"\n  (supersedes the earlier +$89,600, which only matched the average web MIX)")
