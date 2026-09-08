"""Verify the design model on all 16 Mountain/Speed brands and print the exact entry sheet."""

BRANDS = [
    ("MACH I.I", "Bike Bros", "Speed", 77),
    ("TERRAMAX", "Bike Bros", "Mountain", 73),
    ("Mach 0.6", "Bike Bros", "Speed", 61),
    ("TERRAMean", "Bike Bros", "Mountain", 72),
    ("MountainCruise1", "Bike Bros", "Recreation", 76),
    ("LiteSpeed Pro+", "LiteCycle", "Speed", 77),
    ("LiteSpeed+", "LiteCycle", "Speed", 74),
    ("LiteTrail Pro", "LiteCycle", "Mountain", 70),
    ("Hike Bike", "WeBike", "Mountain", 73),
    ("Swift Bike", "WeBike", "Speed", 72),
    ("Blu Ruged Ballz", "BB LLC", "Mountain", 73),
    ("Blu Aero Ballz", "BB LLC", "Speed", 74),
    ("Blu Tail Ballz", "BB LLC", "Mountain", 70),
    ("Blu Tube Ballz", "BB LLC", "Speed", 76),
    ("Spoke'd Easy", "Spoke'd Up", "Recreation", 73),
    ("Spoke'd Speed", "Spoke'd Up", "Speed", 75),
    ("The Armstrong", "SpaceBikes", "Speed", 76),
    ("Mars Rover", "SpaceBikes", "Recreation", 73),
    ("Whole MILC MKII", "MILC Bikes", "Recreation", 74),
    ("Skim MILC MKII", "MILC Bikes", "Speed", 76),
]

COMPONENTS = [
    ("Frame", "Comfort (relaxed)",        [0,0,0,0,1,0,0,0,0,0,0,0,0,0,1,0,0,1,1,0]),
    ("Frame", "Rugged (rough terrain)",   [0,1,0,1,0,0,0,1,1,0,1,0,1,0,0,0,0,0,0,0]),
    ("Frame", "Aerodynamic (speed)",      [1,0,1,0,0,1,1,0,0,1,0,1,0,1,0,1,1,0,0,1]),
    ("Tires", "Mountain - high grip",     [0,1,0,1,0,0,0,1,1,0,1,0,1,0,0,0,0,0,0,0]),
    ("Tires", "Hybrid - road/off-road",   [0,0,0,0,1,0,0,0,0,0,0,0,0,0,1,0,0,1,1,0]),
    ("Tires", "Racing - fast",            [1,0,1,0,0,1,1,0,0,1,0,1,0,1,0,1,1,0,0,1]),
    ("Brakes", "Standard",                [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0]),
    ("Brakes", "Precision",               [1,0,1,0,0,1,1,0,0,1,0,1,0,1,1,1,1,0,1,1]),
    ("Brakes", "Standard disc",           [0,1,0,1,1,0,0,1,1,0,1,0,1,0,0,0,0,0,0,0]),
    ("Handlebars", "Basic straight",      [0,1,0,1,0,0,0,1,1,0,1,0,1,0,0,0,0,0,0,0]),
    ("Handlebars", "Comfort straight",    [0,0,0,0,1,0,0,0,0,0,0,0,0,0,1,0,0,1,1,0]),
    ("Handlebars", "Basic drop down",     [1,0,1,0,0,1,1,0,0,1,0,1,0,1,0,1,1,0,0,1]),
    ("Gears", "7 speed (1x7)",            [0,0,1,0,1,0,0,0,0,0,0,0,0,0,1,0,0,1,1,0]),
    ("Gears", "14 speed (2x7)",           [1,0,0,0,0,1,1,0,0,0,0,1,0,1,0,1,1,0,0,1]),
    ("Gears", "24 speed (3x8)",           [0,1,0,1,0,0,0,1,1,1,1,0,1,0,0,0,0,0,0,0]),
    ("Seat", "Polymer gel all-purpose",   [0,1,0,1,0,0,0,0,1,0,1,0,0,0,0,0,0,0,0,0]),
    ("Seat", "Polymer gel comfort",       [0,0,0,0,1,0,0,1,0,0,0,0,1,0,1,0,0,1,1,0]),
    ("Seat", "Polymer gel racing",        [1,0,1,0,0,1,1,0,0,1,0,1,0,1,0,1,1,0,0,1]),
    ("Accessories", "Reflectors",         [1,0,1,1,1,1,0,0,0,1,0,1,0,0,1,1,0,1,1,0]),
    ("Decals", "Colorful brushstrokes",   [1,1,1,1,1,1,0,1,1,0,1,0,1,1,1,0,1,0,0,1]),
    ("Lights", "Standard",                [1,0,1,0,1,1,1,0,0,1,0,0,0,1,0,1,1,1,1,1]),
    ("Carriers", "Plastic basket",        [0,0,0,0,1,0,0,0,0,0,0,0,0,0,1,0,0,1,1,0]),
    ("Suspension", "Front shocks",        [0,1,0,1,1,0,0,1,1,0,1,0,1,0,1,0,0,1,1,0]),
]

IDX = {n: i for i, (n, *_ ) in enumerate(BRANDS)}


def spec(brand):
    i = IDX[brand]
    out = {}
    for group, label, vals in COMPONENTS:
        if vals[i]:
            out.setdefault(group, []).append(label)
    return out


def gears(i):
    for g, label, vals in COMPONENTS:
        if g == "Gears" and vals[i]:
            return int(label.split()[0])
    return None


def has(i, label):
    for _, lab, vals in COMPONENTS:
        if lab == label:
            return bool(vals[i])
    return False


SPEED_BASE, MTN_BASE = 73, 73
GEAR_PEN = {14: 0, 24: -3, 7: -16}


def predict(i, seg):
    if seg == "Speed":
        s = SPEED_BASE + GEAR_PEN[gears(i)]
        s += 2 if has(i, "Colorful brushstrokes") else 0
        s += 1 if has(i, "Reflectors") else 0
        s += 1 if has(i, "Standard") and False else 0  # 'Standard' is ambiguous; use Lights row
        s += 1 if COMPONENTS[20][2][i] else 0          # Lights / Standard
        return s
    s = MTN_BASE
    if has(i, "Polymer gel comfort"):
        s -= 3
    if has(i, "Reflectors"):
        s -= 1
    return s


print("=" * 74)
print("MODEL VALIDATION — all 16 Mountain and Speed brands")
print("=" * 74)
print(f"{'Brand':<18}{'Seg':<10}{'Actual':>7}{'Model':>7}   {'':<4}")
bad = 0
for name, co, seg, actual in BRANDS:
    if seg == "Recreation":
        continue
    p = predict(IDX[name], seg)
    ok = "OK" if p == actual else "MISMATCH"
    if p != actual:
        bad += 1
    print(f"{name:<18}{seg:<10}{actual:>7}{p:>7}   {ok}")
print(f"\n{16-bad} of 16 exact. " + ("Model is sound." if bad == 0 else "FIX THE MODEL."))

print("\n" + "=" * 74)
print("SPEED: where the 77 ceiling comes from")
print("=" * 74)
print(f"  base (aero frame + racing tires + precision brakes + drop bars + racing seat) = {SPEED_BASE}")
print("  + 14 speed (2x7)   0   <- the correct drivetrain; 24sp costs 3, 7sp costs 16")
print("  + decals           +2")
print("  + reflectors       +1")
print("  + lights           +1")
print(f"  ------------------------\n  MAXIMUM            {SPEED_BASE+4}   (MACH I.I and LiteSpeed Pro+ both sit here)")

print("\n" + "=" * 74)
print("HIKE BIKE — current spec vs the Mountain optimum")
print("=" * 74)
i = IDX["Hike Bike"]
MTN_OPT = [
    ("Frame", "Rugged (rough terrain)", True),
    ("Tires", "Mountain - high grip", True),
    ("Brakes", "Standard disc", True),
    ("Handlebars", "Basic straight", True),
    ("Gears", "24 speed (3x8)", True),
    ("Seat", "Polymer gel all-purpose", True),
    ("Suspension", "Front shocks", True),
    ("Decals", "Colorful brushstrokes", True),
    ("Accessories", "Reflectors", False),
    ("Lights", "Standard", False),
    ("Carriers", "Plastic basket", False),
]
allok = True
for group, label, want in MTN_OPT:
    cur = COMPONENTS[20][2][i] if label == "Standard" else has(i, label)
    mark = "OK" if cur == want else "CHANGE"
    if cur != want:
        allok = False
    print(f"  {group:<12}{label:<26} have {'yes' if cur else 'no ':<4} want {'yes' if want else 'no ':<4} {mark}")
print(f"\n  Hike Bike = 73 of a possible 73. {'ZERO CHANGES NEEDED.' if allok else 'Changes above.'}")

print("\n" + "=" * 74)
print("SWIFT BIKE — current spec vs the Speed optimum")
print("=" * 74)
i = IDX["Swift Bike"]
SPD_OPT = [
    ("Frame", "Aerodynamic (speed)", True, 0),
    ("Tires", "Racing - fast", True, 0),
    ("Brakes", "Precision", True, 0),
    ("Handlebars", "Basic drop down", True, 0),
    ("Seat", "Polymer gel racing", True, 0),
    ("Accessories", "Reflectors", True, 1),
    ("Lights", "Standard", True, 1),
    ("Decals", "Colorful brushstrokes", True, 2),
    ("Carriers", "Plastic basket", False, 0),
    ("Suspension", "Front shocks", False, 0),
]
delta = 0
for group, label, want, worth in SPD_OPT:
    cur = COMPONENTS[20][2][i] if label == "Standard" else has(i, label)
    if cur == want:
        mark = "OK"
    else:
        mark = f"CHANGE  ({worth:+d})"
        delta += worth
    print(f"  {group:<12}{label:<26} have {'yes' if cur else 'no ':<4} want {'yes' if want else 'no ':<4} {mark}")
g = gears(i)
print(f"  {'Gears':<12}{'14 speed (2x7)':<26} have {g}sp  want 14sp  CHANGE  ({-GEAR_PEN[g]:+d})")
delta += -GEAR_PEN[g]
print(f"\n  Swift Bike 72 {delta:+d} = {72+delta} = the industry Speed ceiling (ties MACH I.I, LiteSpeed Pro+)")
print(f"  Beats BB LLC's best Speed brand, Blu Tube Ballz, at 76.")

print("\n" + "=" * 74)
print("WHO ELSE SITS AT EACH SCORE (so we know 77 really is the cap)")
print("=" * 74)
for seg in ("Speed", "Mountain"):
    print(f"\n  {seg}:")
    rows = sorted([b for b in BRANDS if b[2] == seg], key=lambda x: -x[3])
    for name, co, _, sc in rows:
        j = IDX[name]
        bits = []
        if seg == "Speed":
            bits.append(f"{gears(j)}sp")
            bits.append("decals" if has(j, "Colorful brushstrokes") else "NO decals")
            bits.append("refl" if has(j, "Reflectors") else "no refl")
            bits.append("lights" if COMPONENTS[20][2][j] else "no lights")
        else:
            bits.append("all-purpose seat" if has(j, "Polymer gel all-purpose") else "COMFORT seat")
            bits.append("REFLECTORS" if has(j, "Reflectors") else "no reflectors")
        star = "  <-- us" if co == "WeBike" else ""
        print(f"    {sc:>3}  {name:<18}{co:<12}{' · '.join(bits)}{star}")

print("\n" + "=" * 74)
print("MOUNTAIN: what would happen if we 'improved' Hike Bike")
print("=" * 74)
for change, effect in [
    ("Add reflectors", "-1  (TERRAMean proves it: identical bike + reflectors = 72)"),
    ("Swap to comfort seat", "-3  (LiteTrail Pro and Blu Tail Ballz both = 70)"),
    ("Add lights", "UNKNOWN - no Mountain brand in the industry has them. Reflectors are a"),
    ("", "     Recreation cue and cost a point, so lights are a live downside risk."),
    ("Add a basket", "UNKNOWN - same reasoning, and it is a Recreation component"),
    ("Remove decals", "UNKNOWN - all 6 Mountain brands have them, so the value is untestable"),
]:
    print(f"  {change:<22}{effect}")
print("\n  Every identifiable change to Hike Bike is negative or unknown. Leave it alone.")
print("=" * 74)
