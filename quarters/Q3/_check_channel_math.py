"""Quick check of the web-rebuild arithmetic quoted in results.md."""
CH = [
    ("BB LLC",      938, 569, 1507, 2),
    ("SpaceBikes",  807, 491, 1298, 2),
    ("LiteCycle",   549, 435,  984, 2),
    ("MILC Bikes",  364, 435,  799, 1),
    ("Spoke'd Up",  716, 138,  854, 2),
    ("WeBike",      507, 120,  627, 2),
    ("Bike Bros",   490,   0,  490, 1),
]

with_web = [c for c in CH if c[2] > 0]
web_avg = sum(c[2] / c[3] for c in with_web) / len(with_web)
print(f"industry avg web mix (6 firms with a web centre): {web_avg:.1%}")
print(f"WeBike web mix: {120/627:.1%}   gap: {web_avg - 120/627:.1%}\n")

store = 507
total = store / (1 - web_avg)
extra = total - store - 120
print(f"hold store demand at {store}, match avg web mix:")
print(f"  implied total demand {total:.0f}")
print(f"  implied web demand   {total - store:.0f}")
print(f"  extra vs today       {extra:.0f} units")
print(f"  gross margin @ $750  ${extra * 750:,.0f}")
print(f"  less ~4 web staff    ${24425:,}")
print(f"  net quarterly gain   ${extra * 750 - 24425:,.0f}\n")

print("demand per store:")
for n, s, w, t, ns in sorted(CH, key=lambda x: -x[1] / x[4]):
    print(f"  {n:<12}{s/ns:6.0f}")
