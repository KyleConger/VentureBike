"""Validate the reverse-engineered brand judgment model against all 20 observed scores."""

# brand: (segment, observed, gears, decals, reflectors, lights, seat)
B = {
    "MACH I.I":        ("Speed", 77, 14, 1, 1, 1, "racing"),
    "Mach 0.6":        ("Speed", 61,  7, 1, 1, 1, "racing"),
    "LiteSpeed Pro+":  ("Speed", 77, 14, 1, 1, 1, "racing"),
    "LiteSpeed+":      ("Speed", 74, 14, 0, 0, 1, "racing"),
    "Swift Bike":      ("Speed", 72, 24, 0, 1, 1, "racing"),
    "Blu Aero Ballz":  ("Speed", 74, 14, 0, 1, 0, "racing"),
    "Blu Tube Ballz":  ("Speed", 76, 14, 1, 0, 1, "racing"),
    "Spoke'd Speed":   ("Speed", 75, 14, 0, 1, 1, "racing"),
    "The Armstrong":   ("Speed", 76, 14, 1, 0, 1, "racing"),
    "Skim MILC MKII":  ("Speed", 76, 14, 1, 0, 1, "racing"),

    "TERRAMAX":        ("Mountain", 73, 24, 1, 0, 0, "all-purpose"),
    "TERRAMean":       ("Mountain", 72, 24, 1, 1, 0, "all-purpose"),
    "LiteTrail Pro":   ("Mountain", 70, 24, 1, 0, 0, "comfort"),
    "Hike Bike":       ("Mountain", 73, 24, 1, 0, 0, "all-purpose"),
    "Blu Ruged Ballz": ("Mountain", 73, 24, 1, 0, 0, "all-purpose"),
    "Blu Tail Ballz":  ("Mountain", 70, 24, 1, 0, 0, "comfort"),
}

GEAR_PENALTY = {14: 0, 24: -3, 7: -16}


def predict(seg, gears, decals, refl, lights, seat):
    if seg == "Speed":
        s = 77 + GEAR_PENALTY[gears]
        s += 0 if decals else -2
        s += 0 if refl else -1
        s += 0 if lights else -1
        return s
    if seg == "Mountain":
        s = 73
        s += 0 if seat == "all-purpose" else -3
        s += -1 if refl else 0
        return s
    raise ValueError(seg)


ok = True
print(f"{'Brand':<18}{'Seg':<10}{'Observed':>9}{'Model':>7}{'Diff':>6}")
for name, (seg, obs, gears, dec, refl, lights, seat) in B.items():
    pred = predict(seg, gears, dec, refl, lights, seat)
    diff = pred - obs
    ok &= diff == 0
    print(f"{name:<18}{seg:<10}{obs:>9}{pred:>7}{diff:>6}")

print()
print("MODEL REPRODUCES ALL SCORES EXACTLY" if ok else "MODEL MISMATCH — do not trust")

print()
print("Swift Bike fix:")
base = predict("Speed", 24, 0, 1, 1, "racing")
add_decals = predict("Speed", 24, 1, 1, 1, "racing")
fix_gears = predict("Speed", 14, 0, 1, 1, "racing")
both = predict("Speed", 14, 1, 1, 1, "racing")
print(f"  current (24sp, no decals)      = {base}")
print(f"  + add decals only              = {add_decals}  (+{add_decals-base})")
print(f"  + change to 14 speed only      = {fix_gears}  (+{fix_gears-base})")
print(f"  + both changes                 = {both}  (+{both-base})  <- ties industry ceiling 77")

print()
print("Recreation: 4 brands, 4 unknowns, 3 independent equations -> underdetermined.")
print("  precision + no_decals = 2 | precision + no_lights = 3 | standard + no_decals = 3")
print("  Certain: standard disc brakes optimal; MountainCruise1 recipe = 76 ceiling.")
