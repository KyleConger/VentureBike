"""Append ad judgment analysis to results.md, before the open-decisions block."""
import io

PATH = "quarters/Q3/results.md"
s = io.open(PATH, encoding="utf-8").read()

ANCHOR = "## Open decisions for Q4"
assert ANCHOR in s, "anchor missing"

NEW = """## Ad judgment: HikeBike 1 is nearly perfect, the Swift Bike ad is not

| Ad | Company | Segment | Score in target | Off-target leakage |
|----|---------|---------|----------------:|-------------------:|
| BB Big Momma | BB LLC | Mountain | **81** | 55 |
| **HikeBike 1** | **WeBike** | **Mountain** | **80** | 59 |
| Trail Blazing 1 | LiteCycle | Mountain | 79 | 52 |
| AndStill | Bike Bros | Speed | **78** | 46 |
| Speed of Lite 1 | LiteCycle | Speed | 77 | 32 |
| The Armstrong 1 | SpaceBikes | Speed | 77 | 23 |
| Unleash lil pap | BB LLC | Speed | 76 | 31 |
| Skim MILC MKII | MILC Bikes | Speed | 75 | 29 |
| **Swift Bike** | **WeBike** | **Speed** | **70** | **88** |
| Spoke'd Speed | Spoke'd Up | Speed | 57 | 22 |

**HikeBike 1 scores 80 in Mountain, one point behind the industry best.** Of all the Q3 work we did, the Mountain ad is the piece that landed. It does not need a redesign, and redesigning costs money we need elsewhere.

**The Swift Bike ad scores 70 in Speed — eighth of nine.** Only Spoke'd Up's ad is worse, and their ads are so bad they finished last on Marketing Effectiveness because of it.

## The Marketing Effectiveness formula, confirmed

Ad judgment was the missing input. The class formula is:

> Marketing Effectiveness = (avg brand judgment / 100 + avg ad judgment / 100) / 2, across the two targeted segments

Running it on the field reproduces **all three** reported figures exactly:

| Company | Avg brand | Avg ad | Computed ME | Reported |
|---------|----------:|-------:|------------:|---------:|
| BB LLC | 74.5 | 78.5 | 0.7650 | **0.765 (industry max)** |
| SpaceBikes | 74.5 | 78.0 | 0.7625 | — |
| MILC Bikes | 75.0 | 76.0 | 0.7550 | — |
| **WeBike** | **72.5** | **75.0** | **0.7375** | **0.738** |
| Spoke'd Up | 74.0 | 58.0 | 0.6600 | **0.660 (industry min)** |

Three for three. This is now a solved formula, which means the Swift Bike fix can be priced precisely rather than argued about.

## Fixing Swift Bike makes us the industry's best marketer

Every input is a brand or ad decision we control, and both fixes are cheap:

| Scenario | Brand judgment | Ad judgment | Marketing Effectiveness |
|----------|---------------:|------------:|------------------------:|
| Q3 actual | 72 | 70 | 0.7375 (below the 0.743 average) |
| Fix brand only (14-speed + decals) | 77 | 70 | 0.7500 |
| Fix ad only | 72 | 77 | 0.7550 |
| **Fix both** | **77** | **77** | **0.7675 — new industry best** |

Doing both lifts us past BB LLC's 0.765 to first place on an indicator where we currently sit below average. The brand fix was already identified from the component teardown (drop to 14-speed, add decals); the ad fix is a redesign fee.

## Why the Swift Bike ad is weak: we cloned our Mountain playbook

Our Swift Bike ad scores **52 in Recreation and 36 in Mountain** — both the highest off-target scores of any Speed ad in the industry. Total leakage is 88, against 12–46 for everyone else. The Armstrong 1 scores just 2 in Recreation.

That pattern points at an ad carrying Mountain- and Recreation-flavored claims instead of Speed-specific ones. It is the same mistake as the drivetrain: **Swift Bike inherited Hike Bike's 24-speed gearing and, it appears, Hike Bike's messaging.** We cloned our Mountain assets into a Speed launch and paid for it in both brand and ad judgment.

One caveat worth stating plainly: leakage does not cleanly predict score across the whole field — AndStill leaks 46 and still posts the best Speed score of 78. So this is a well-supported hypothesis, not a solved model like the component analysis. The fix is to rebuild the ad around Speed cues (aerodynamic, racing tires, precision brakes, drop bars, light weight) and read the score in the designer before committing.

## Ad quantity is not the lever

| Company | Distinct ads | Brands | Note |
|---------|-------------:|-------:|------|
| Bike Bros | 5 | 5 | two ads each for MACH I.I and TERRAMAX — and finished **last** |
| LiteCycle | 3 | 3 | one per brand |
| **BB LLC** | **2** | **4** | Blu Aero and Blu Tail run with **no ad at all** |
| WeBike | 2 | 2 | one per brand |

The leader runs the fewest ads relative to its brand count and both are best or near-best in class. The last-place firm runs the most. A second Swift Bike ad is not the answer — a better one is.

Worth filing for later: BB LLC pairs an advertised brand with an unadvertised second brand in each segment (Blu Ruged + Blu Tail in Mountain, Blu Tube + Blu Aero in Speed). The ad builds segment awareness and the cheaper-to-run second brand catches overflow at a higher price. Not a Q4 move for us while capacity binds, but a real pattern.

"""

s = s.replace(ANCHOR, NEW + ANCHOR)
s = s.replace(
    "- [ ] Speed brand: invest vs refocus",
    "- [ ] Speed brand: invest vs refocus (brand + ad fix together = industry-best Marketing Effectiveness)",
)
io.open(PATH, "w", encoding="utf-8", newline="\n").write(s)
print("appended", len(NEW), "chars")
