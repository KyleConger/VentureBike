# Q3 Results — WeBike (filed 2026-09-08)

**Workbook:** `Q3Data.xlsx` (sheets `Q3_BSC`, `Q3_Market`, `Q3_Financials`, `Q3_Unit_Economics`, `Q4_Planner`)
**Headline:** WeBike finished **last** in Total Performance (**0.411** — exactly the industry minimum; average 5.897).

## What we did in Q3

- Designed a new Speed bike (**Swift Bike**) + new Speed ad
- Diversified media channels for the Mountain ad
- Added sales & service people for Speed
- **Reduced** web sales and support headcount; cut web spend
- **Stopped** the page-upgrade web productivity option

## Balanced Scorecard

| Indicator | WeBike | Average | Max | Read |
|-----------|-------:|--------:|----:|------|
| **Total Performance** | **0.411** | 5.897 | 21.531 | Last |
| Financial Performance | 5.319 | 19.006 | 37.687 | Last |
| **Market Performance** | **0.098** | 0.185 | 0.345 | 33% stock-outs |
| Marketing Effectiveness | 0.738 | 0.743 | 0.765 | At par |
| Investment in Future | **6.795** | 5.386 | 7.401 | Strength |
| Wealth | 0.668 | 0.831 | 1.017 | Last |
| HR Management | 0.713 | 0.724 | 0.767 | Below avg |
| **Asset Management** | **0.353** | 0.542 | 0.920 | Last |
| Manufacturing Productivity | **0.938** | 0.899 | 0.995 | Strength (misleading) |
| Financial Risk | 1.000 | 1.000 | 1.000 | No debt |

Because Total Performance **multiplies** every indicator, the two "last place" scores (Market 0.098, Asset 0.353) do most of the damage.

## Root cause — we under-scheduled the factory

| Step | Value |
|------|------:|
| Units produced | 423 |
| Overtime as % of OC used | 13% |
| Implied normal-hours output | ~374 units → ~5.8/day |
| Actual worker productivity | 72% |
| **Implied scheduled operating capacity** | **~8 / day** |
| **Fixed capacity we owned** | **24 / day (1,560/qtr)** |

We owned 24/day and scheduled roughly **8/day**, then ran overtime on top of it. This was an over-correction from Q2's **$148,652 excess capacity charge**.

Result: **204 of 627 units of demand (33%) were never built.**

### What the stock-out cost

| Brand | Price | Cost | Contribution | Lost units | Lost margin |
|-------|------:|-----:|-------------:|-----------:|------------:|
| Hike Bike | $1,365 | $618 | $747 | 134 | $100,098 |
| Swift Bike | $1,450 | $694 | $756 | 70 | $52,920 |
| **Total** | | | | **204** | **$153,018** |

Q3 operating profit was **−$111,431**. Serving that demand would have produced roughly **+$41,600** — a profitable quarter — before counting the unit-cost improvement from higher volume.

Manufacturing Productivity scored 0.938 precisely *because* we matched output to a too-small capacity plan. That indicator does not punish stock-outs; **Market Performance does**.

## Market position

| Segment | Our share | Rank | Note |
|---------|----------:|-----:|------|
| Mountain | **22.1%** | 2 of 7 | Only 4 firms compete; BB LLC leads at 46.3% |
| Speed | 7.5% | 6 of 7 | All 7 firms compete; largest segment (2,878) |
| Recreation | 2.6% | — | Not targeted |

Mountain remains our real franchise. Speed entry bought share of a crowded segment at a weak ad judgment (**70** vs Mountain's **80**).

## Carry-forward penalties into Q4

- **Ill will:** 32.5% of demand was lost → Q4 generated demand is cut by **16.3%**
- **Productivity came in at 70% sales / 72% production** against our 85% projection — see the correction below
- **Cash is $1,010,838 and idle** — the direct cause of Asset Management 0.353 (last)
- We issued another **$500,000** of stock (25,000 shares), which lowers per-share Financial Performance and further dilutes asset turnover

## What moves the needle in Q4 (sensitivity)

| Change | Total Performance |
|--------|------------------:|
| Q3 actual | 0.411 |
| Serve 100% of demand (same share) | **0.631** |
| + Asset turnover 0.353 → 0.55 | **0.983** |
| + HR 0.713 → 0.80 | **1.103** |

## Q4 priorities

1. **Schedule operating capacity to the forecast, not below it.** Planner in `Q3Data.xlsx` → `Q4_Planner`: at ~656 forecast demand and 72% productivity, required OC ≈ **13/day**; suggested **15/day** with cushion. Fixed 24/day covers it — no printers needed yet.
2. **Raise compensation, but fix the planning assumption.** Move toward upper-quartile pay (sales ~$26–27k, production ~$22–23k) and budget capacity on **73–75%** productivity — the industry ceiling is ~73%, so 85% was never achievable.
3. **Stop issuing stock** while $1M sits idle, and deploy the cash into revenue (city #3, sales people, ads).
4. **Speed: fix it, don't abandon it.** Redesign Swift Bike toward the 12/1/77 profile (we're at 10/1/72, and 8 of 9 rivals beat us) and strengthen the Speed ad (judgment 70). Leave Hike Bike's design alone — it's tied for best in the industry at 73.
5. Gross margin (53.8%, lowest in industry) rises automatically with volume — another reason to build to demand.

---

# Industry comparison — Strategic Graphs (Q4 view, Q0–Q3)

Chart-read values are approximate (±1 gridline). Full table: `Q3Data.xlsx` → **Q3_Industry_Graphs**.

## The capacity chart confirms the diagnosis outright

The Operating Capacity graph plots our scheduled capacity dropping **20 → 8 units/day** while every rival raised theirs. This is no longer an inference from the overtime math — it's on the chart.

| Firm | Fixed cap | Scheduled OC (Q3) |
|------|----------:|------------------:|
| SpaceBikes | 32 | 30 |
| MILC Bikes | 24 | 24 |
| Bike Bros | 16 | 15 |
| LiteCycle | 16 | 15 |
| BB LLC | 16 | 15 (+~7 overtime) |
| Spoke'd Up | 16 | 14 |
| **WeBike** | **24** | **8** |

We had the second-highest fixed capacity in the industry and scheduled the lowest operating capacity. We were the only firm that *reduced* operating capacity in Q3.

## New finding: under-scheduling also raised our unit cost

| WeBike | Q2 | Q3 |
|--------|---:|---:|
| Capacity utilization | 16% | 73% |
| Average labor cost per unit | ~$130 | **~$205** |
| Average COGS per unit | ~$520 | **~$620** |

This was the opposite of what I expected. We sold four times the units and unit cost still went **up**, because a sub-scale operating capacity run with overtime is expensive per unit. So right-sizing capacity in Q4 is not just a revenue fix — it should pull COGS down too. That also explains why our gross margin (53.8%) is the industry's lowest.

## Correction: we were not badly underpaid

I previously wrote that peers out-paid us into a productivity gap. The HR graphs don't support that.

| Metric | WeBike | Industry range |
|--------|-------:|----------------|
| Production compensation | $20,757 | $20,500 – $23,500 |
| Sales compensation | $24,425 | $22,000 – $27,000 |
| Production productivity | 72% | 68 – 73% |
| Sales productivity | 70% | 70 – 75% |

Our pay sits low-middle, but productivity is **mid-pack** — the entire industry runs about 70%. The real mistake was **planning capacity on an 85% assumption** that nobody in this market achieves. Raising pay is still worth doing (it buys a few points and helps the HR indicator), but budget on **73–75%**, not 85%.

## Marketing is not our problem

| Input | WeBike | Industry high | Rank |
|-------|-------:|--------------:|-----:|
| Regional ads | 19 | 24 | 2nd |
| Advertising spend | $119,165 | ~$145,000 | 3rd |
| Average price | ~$1,400 | ~$1,480 | 3rd |
| Organic SEM clicks | ~150 | ~440 | 6th |
| Store + web center spend | $121,000 | ~$360,000 | **lowest** |

Ad volume and spend are competitive, which is why Marketing Effectiveness sits at par. Two things stand out instead: organic SEM clicks are weak, and **we spend the least in the industry on channel**. Our Q3 instinct to cut web staff and spend was pointed at the wrong target.

Related: our **share of demand was 9.6% but share of units sold only 7.3%** — the stock-out leakage is visible directly in the graphs. And our Q3 demand growth rate (~235%) was the highest in the industry. **We are good at creating demand and bad at filling it.**

## Demand per head is our weakest operating metric

| Per-head demand (Q3) | WeBike | Best | Typical |
|----------------------|-------:|-----:|--------:|
| Per sales person | ~45 | ~72 | ~60 |
| Per store sales person | ~46 | ~88 | ~70 |
| Per web center person | ~42 | ~95 | ~75 |

Headcount is at parity (14 people vs a typical 14), so this is an output problem, not a staffing-level problem.

## Who's winning, and what to take from them

**BB LLC** (Q3 leader, Total Performance ~22) competes in all three segments, charges the *highest* price (~$1,480) and still leads share, runs the most regional ads (24), and posts Asset Management of 0.920. Notably it keeps fixed capacity at 16 and buys flexibility with ~7 units of overtime instead of printers.

**SpaceBikes** (cumulative leader, ~6.3) skips Mountain entirely, owns Recreation at 34.5% plus Speed at 20.4%, and was the only firm to out-build us — fixed capacity 32/day.

**Bike Bros** is the cautionary tale: one store, zero web sales centers, seven sales people, and share falling in all three segments (Mountain 19% → 11%). That's where under-investing in channel leads.

## The segment tension

| Segment | Q3 size | Growth | Searches | Our share | Rivals |
|---------|--------:|-------:|---------:|----------:|-------:|
| Speed | 2,878 | +92% | ~900 | 7.5% | all 7 |
| Recreation | 2,060 | +87% | ~1,350 | 2.6% | 3 |
| Mountain | 1,621 | +54% | ~1,030 | 22.1% | 4 |

Our least-crowded segment is also the smallest and slowest-growing, and BB LLC holds 46% of it. Mountain alone caps our ceiling — worth weighing before we decide whether to abandon or reinforce Speed.

---

# Brand Judgment — World Market, Q3 (all 20 industry brands)

Exact scores from the sim report. Full table: `Q3Data.xlsx` → **Q3_Brand_Judgment**.

## Hike Bike is tied for the best Mountain brand in the industry

| Brand | Company | Mountain judgment |
|-------|---------|------------------:|
| **Hike Bike** | **WeBike** | **73** |
| TERRAMAX | Bike Bros | 73 |
| Blu Ruged Ballz | BB LLC | 73 |
| TERRAMean | Bike Bros | 72 |
| LiteTrail Pro | LiteCycle | 70 |
| Blu Tail Ballz | BB LLC | 70 |

Nobody in the industry exceeds 73, so that looks like the ceiling — and we're at it.

This is the single most useful thing in the report, because **BB LLC takes 46.3% of Mountain to our 22.1% with a brand scored *identically* to ours (56/73/1)**. Same product appeal, 2.1× the share. The gap cannot be product. It is:

- **Supply** — we stocked out on 33% of demand; they didn't
- **Coverage** — 21 sales people to our 14, ~88 demand per store rep to our 46
- **Ad volume** — 24 regional ads to our 19

**Do not redesign Hike Bike. Feed it.** Redesign spend on Mountain is wasted money.

## Swift Bike is the weakest real Speed brand

| Brand | Company | Speed judgment |
|-------|---------|---------------:|
| MACH I.I | Bike Bros | 77 |
| LiteSpeed Pro+ | LiteCycle | 77 |
| Blu Tube Ballz | BB LLC | 76 |
| The Armstrong | SpaceBikes | 76 |
| Skim MILC MKII | MILC Bikes | 76 |
| Spoke'd Speed | Spoke'd Up | 75 |
| LiteSpeed+ | LiteCycle | 74 |
| Blu Aero Ballz | BB LLC | 74 |
| **Swift Bike** | **WeBike** | **72** |
| Mach 0.6 | Bike Bros | 61 |

Eight of nine rival Speed brands beat us; only Bike Bros' abandoned Mach 0.6 is worse. But the winning profile (12/1/77) is very close to ours (10/1/72), so a **modest redesign should close most of the 5-point gap**.

This answers the Speed question from the graph analysis. Weak product (72) *plus* weak ad (judgment 70) fully explains 7.5% share in a 2,878-unit segment. Speed isn't a lost cause — it's an under-built entry. Fixing both is cheaper than abandoning the segment.

## Brand design is commoditising

Multiple firms have landed on byte-identical profiles:

| Profile (Rec/Mtn/Speed) | Brands sharing it |
|-------------------------|-------------------|
| 56 / 73 / 1 | TERRAMAX, **Hike Bike**, Blu Ruged Ballz |
| 58 / 70 / 1 | LiteTrail Pro, Blu Tail Ballz |
| 12 / 1 / 77 | MACH I.I, LiteSpeed Pro+ |
| 9 / 1 / 76 | Blu Tube Ballz, The Armstrong, Skim MILC MKII |
| 7 / 1 / 74 | LiteSpeed+, Blu Aero Ballz |
| 73 / 1 / 1 | Spoke'd Easy, Mars Rover |

Everyone is converging on the same optima, and the ceilings (73 Mountain, 77 Speed) appear hard. **No one can win on product alone from here** — differentiation has to come from supply, channel, ads, and price. That is consistent with everything the Strategic Graphs showed.

## Brand count doesn't buy share

| Company | Brands | Overall share |
|---------|-------:|--------------:|
| Bike Bros | 5 | 7.5% |
| BB LLC | 4 | 23.0% |
| LiteCycle | 3 | 15.0% |
| **WeBike** | **2** | **9.6%** |
| Spoke'd Up | 2 | 13.0% |
| SpaceBikes | 2 | **19.8%** |
| MILC Bikes | 2 | 12.2% |

Bike Bros has the most brands and the worst share; SpaceBikes leads cumulatively on two. **A third brand is not our Q4 priority** — filling demand for the two we have is.

## One free option we're sitting on

Hike Bike scores **56 with Recreation buyers** and pulled **54 units of Recreation demand we never targeted** — in the segment with the highest search volume (~1,350) and 2,060 units of size. But 56 against a dedicated 76 means we can't win Recreation on spillover. Worth parking as a Q5 option with a real design, not a Q4 scramble.

---

# Component teardown & reverse-engineered design model

Full matrix: `Q3Data.xlsx` → **Q3_Components** (20 brands × 23 components). Fitted model: **Q3_Design_Model**. Validation: `_validate_design_model.py`.

Every brand in a segment shares one base recipe; scores differ only on a handful of components. Fitting point values to those differences **reproduces all 16 Mountain and Speed scores exactly, with zero error** — so this is a solved system, not a guess.

## Speed scoring model

Base recipe shared by all 10 Speed brands: aerodynamic frame, racing tires, precision brakes, basic drop-down handlebars, polymer gel racing seat, no carrier, no suspension.

| Variable component | Effect |
|--------------------|-------:|
| 14 speed (2×7) | optimal |
| 24 speed (3×8) instead | **−3** |
| 7 speed (1×7) instead | **−16** |
| Decals (colorful brushstrokes) | **+2** |
| Reflectors | +1 |
| Lights | +1 |
| **Maximum achievable** | **77** |

## Swift Bike: two changes reach the industry ceiling

| Component | Now | Change to | Points |
|-----------|-----|-----------|-------:|
| Gears | 24 speed (3×8) | **14 speed (2×7)** | +3 |
| Decals | not included | **add colorful brushstrokes** | +2 |
| | | **72 → 77** | **+5** |

Everything else on Swift Bike is already correct — frame, tires, brakes, handlebars, seat, reflectors, and lights all match the 77-point brands.

**The root cause is a copy-paste error.** Swift Bike inherited Hike Bike's **24-speed (3×8) mountain drivetrain**. That gearing is optimal in Mountain — all six Mountain brands use it — and wrong in Speed, where all eight leading Speed brands run 14-speed. Swift Bike is the *only* Speed brand in the industry with 24 speeds. We built a Speed bike with mountain gears.

Worth noting: 14-speed (2×7) is a simpler drivetrain than 24-speed (3×8), so this change should also **lower** unit cost. Verify against the component price list before locking, but this looks like a rare fix that raises appeal and cuts cost at once.

## Mountain scoring model — Hike Bike is verified optimal

Base recipe shared by all 6 Mountain brands: rugged frame, mountain high-grip tires, standard disc brakes, basic straight handlebars, 24 speed, front shocks, decals, no lights.

| Variable component | Effect |
|--------------------|-------:|
| Polymer gel **all-purpose** seat | optimal |
| Polymer gel **comfort** seat instead | **−3** |
| Reflectors added | **−1** |
| **Maximum achievable** | **73** |

That explains the whole Mountain field. LiteTrail Pro and Blu Tail Ballz score 70 because they chose a comfort seat. TERRAMean scores 72 because it added reflectors — otherwise identical to Hike Bike.

**Hike Bike scores 73 of a possible 73.** It has the all-purpose seat, correctly omits reflectors, and correctly omits lights. There is literally nothing to gain from redesigning it, which confirms the earlier conclusion from a second direction: Q4 money belongs in capacity, coverage, and ads.

## Components are segment-dependent, not universally good

| Component | Speed | Mountain | Recreation |
|-----------|------:|---------:|-----------:|
| Reflectors | **+1** | **−1** | required |
| Decals | **+2** | required | varies |
| Lights | **+1** | excluded | required |
| Biggest lever | gears | seat | brakes |

Reflectors help in Speed and *hurt* in Mountain — they read as a Recreation cue. This is the trap that cost TERRAMean a point and the reason a shared component list across brands is dangerous.

## Recreation blueprint (Q5 option, with a caveat)

MountainCruise1 holds the 76 ceiling with: comfort frame, hybrid tires, **standard disc brakes**, comfort straight handlebars, 7 speed, comfort seat, plus reflectors, decals, lights, plastic basket, and front shocks.

Caveat on precision: only four Recreation brands exist and they differ in three places (brakes, decals, lights), so the individual penalties **cannot be uniquely separated** from the data. The observed constraints are `precision + no_decals = 2`, `precision + no_lights = 3`, `standard + no_decals = 3`. What's certain is that standard disc brakes are optimal and MountainCruise1's exact recipe scores 76.

---

# City-level demand — the map changes the city #3 decision

Raw data: `Q3Data.xlsx` → **Q3_City_Demand** (65 brand × city rows). Analysis: **Q3_City_Strategy**. Totals reconcile exactly to 6,559 market / 627 WeBike / 412 Hike Bike / 215 Swift Bike.

## Where the market actually is

| City | Recreation | Mountain | Speed | Total | % of market | Our units | Our share |
|------|-----------:|---------:|------:|------:|------------:|----------:|----------:|
| **Amsterdam** | 992 | 650 | 1,343 | **2,985** | 45.5% | 295 | **9.9%** |
| **New York City** | 664 | 495 | 880 | **2,039** | 31.1% | 38 | **1.9%** |
| Rio de Janeiro | 219 | 286 | 269 | 774 | 11.8% | 270 | **34.9%** |
| Bangalore | 185 | 190 | 386 | 761 | 11.6% | 24 | 3.2% |

**Amsterdam and New York are 77% of the entire market.** We hold 9.9% of one and 1.9% of the other, while our best share by far sits in Rio — the smallest city.

## A store is worth roughly 10× web-only coverage

Store locations — **confirmed by the "Competitors in City" report** (see the Presence section below):

| City | Firms with a store | Our share |
|------|--------------------|----------:|
| Amsterdam | **all seven** | 9.9% |
| New York City | BB LLC, SpaceBikes, Spoke'd Up | **1.9%** (no store) |
| Rio de Janeiro | **WeBike only** | **34.9%** |
| Bangalore | LiteCycle only | 3.2% (no store) |

The pattern is stark. With no store we get 1.9–3.2%. With a store against six rivals we get 9.9%. With the *only* store in a city we get 34.9% — and LiteCycle, alone in Bangalore, gets 48.6%. Bike Bros appears in no city but Amsterdam, which independently confirms its one-store, zero-web-center profile.

This makes store placement one of the highest-leverage decisions available to us.

## Rio is our crown jewel, and we damaged it

| Mountain by city | Size | Our units | Our share | Leader |
|------------------|-----:|----------:|----------:|--------|
| Amsterdam | 650 | 156 | 24.0% | us |
| New York City | 495 | 23 | 4.6% | BB LLC 82.8% |
| **Rio de Janeiro** | 286 | **166** | **58.0%** | **us** |
| Bangalore | 190 | 13 | 6.8% | LiteCycle 64.2% |

Rio delivers Hike Bike's single best result — **166 Mountain units, more than Amsterdam's 156 out of a segment less than half the size**. We hold 58% of Rio Mountain.

That makes the Q3 stock-out worse than it first looked: the 16.3% ill-will penalty bites hardest exactly where we lead, in the one market where we have no competition for shelf space.

## City #3: New York is the larger prize

| | New York City | Bangalore |
|---|---:|---:|
| City demand | **2,039** | 761 |
| Rival stores present | 3 | 1 |
| Our share today | 1.9% | 3.2% |
| Mountain size / leader hold | 495 / BB LLC 82.8% | 190 / LiteCycle 64.2% |
| Speed size | **880** | 386 |
| Recreation size (untapped) | 664 | 185 |
| Units at a conservative 10% | **204** | 76 |
| Units at a Rio-like 25% | 510 | 190 |

New York at a weak 10% still beats Bangalore at 25%, and its Speed segment alone (880) is larger than all of Bangalore (761). It also fits our locked "largest geos even if expensive" direction.

The caution: New York is BB LLC's fortress — 36.9% of the city and **82.8% of its Mountain**. We should not attack there head-on in Mountain. The opening is **Speed** (880 units with no dominant firm; The Armstrong leads on only 272) — which is exactly the segment where the two-component Swift Bike fix takes us to the industry ceiling. Those two decisions reinforce each other.

## Capacity can absorb it

| | Units |
|---|---:|
| Fixed capacity 24/day × 65 days | 1,560 |
| At 74% productivity | 1,154 |
| Q4 forecast, existing cities | 656 |
| **Headroom** | **~498** |

We can absorb roughly 500 additional units without buying a printer, so a New York store at ~204 units fits comfortably. The condition is scheduling operating capacity to match — the Q3 failure repeated at three cities would be far more expensive.

## The competing use of the money

Worth weighing before locking city #3: **Amsterdam is 45.5% of the market and we hold 9.9% of it.** Our demand per store salesperson (~46) is roughly half BB LLC's (~88), so there is real headroom in the stores we already own — and deepening them costs no new lease.

> **Corrected below.** I initially read Amsterdam as "genuinely open" because no rival exceeds 16.5%. The Competitors in City report shows all seven firms hold an Amsterdam store — nobody exceeds 16.5% *because* all seven are there. See the Presence section.

---

# Channel mix — the Q3 web cut has a price tag

Data and analysis: `Q3Data.xlsx` → **Q3_Channel**. Store + web reconciles to total for every firm, and to 6,559 industry-wide.

| Company | Store | Store % | Web | Web % | Total |
|---------|------:|--------:|----:|------:|------:|
| BB LLC | 938 | 62.2% | **569** | 37.8% | 1,507 |
| SpaceBikes | 807 | 62.2% | **491** | 37.8% | 1,298 |
| LiteCycle | 549 | 55.8% | 435 | 44.2% | 984 |
| Spoke'd Up | 716 | 83.8% | 138 | 16.2% | 854 |
| MILC Bikes | 364 | 45.6% | 435 | **54.4%** | 799 |
| **WeBike** | **507** | **80.9%** | **120** | **19.1%** | **627** |
| Bike Bros | 490 | 100.0% | 0 | 0.0% | 490 |

Our 120 web units are the **lowest of any firm that operates a web centre**, and our 19.1% web mix sits far below the 34.9% average across the six firms with one. We cut web staff and spend in Q3; this is the bill.

## The top three firms all run 38–44% web

| Model | Who | Result |
|-------|-----|--------|
| Web-first | MILC Bikes, 54.4% | 12.2% share off a single store base |
| **Balanced (38–44%)** | LiteCycle, BB LLC, SpaceBikes | **the three strongest firms** |
| Store-heavy | Spoke'd Up 83.8%, **WeBike 80.9%** | 13.0% and 9.6% |
| Store-only | Bike Bros, 100% | **7.5%, last, falling everywhere** |

Bike Bros is the control group for this experiment, and we are drifting toward its failure mode.

## Web is half our reach into cities where we have no store

Cross-referencing the city data: our 38 New York units and 24 Bangalore units are **web-only** — we have no store in either. That's 62 of our 120 web units, or **52% of web demand coming from cities we can't otherwise touch**. (The remaining 58 units are web orders from Amsterdam and Rio, which reconciles: 565 total in store cities − 507 store demand = 58.)

So the web centre isn't a side channel. It's our only presence in 43% of the market by demand, and we shrank it.

## The rebuild is the cheapest demand available

Holding store demand constant at 507 and simply matching the industry-average web mix:

| | Units |
|---|---:|
| Implied total demand | 779 |
| Implied web demand | 272 |
| **Additional demand vs today** | **+152** |
| Gross margin at ~$750/unit | **$114,050** |
| Less ~4 web staff for a quarter | −$24,425 |
| **Net quarterly gain** | **≈ $89,600** |

No lease, no store setup, and it reaches all four cities. For comparison, Q3's entire operating loss was $111,431 — rebuilding the web channel alone recovers most of it.

We also stopped the page-upgrade tactic in Q3 (~$6,000/quarter) while keeping toll-free. Against ~$750 of contribution per unit, both were far too cheap to have cut.

## We also under-staff the stores we already own

| Company | Demand per store |
|---------|-----------------:|
| Bike Bros | 490 (one store) |
| BB LLC | 469 |
| SpaceBikes | 404 |
| MILC Bikes | 364 |
| Spoke'd Up | 358 |
| LiteCycle | 274 |
| **WeBike** | **254** |

BB LLC pulls **1.8× our demand from the same number of stores**. Bike Bros — with its single store — pulls nearly double ours per location.

## This shifts the Q4 spending order

The city #3 analysis said New York is the biggest prize, and that still holds. But this data says the **first** dollars should go to channel depth rather than channel breadth: rehire web staff, restart page upgrades, and add store sales people in Amsterdam and Rio. Those moves are cheaper, they pay back inside the quarter, and they don't add the fixed lease and setup cost of a third city while we're still carrying a $111,431 operating loss.

---

# Confirmed competitor footprints (Competitors in City)

Data and analysis: `Q3Data.xlsx` → **Q3_Presence**.

| Location | Bike Bros | LiteCycle | WeBike | BB LLC | Spoke'd Up | SpaceBikes | MILC | Store firms |
|----------|:---:|:---:|:---:|:---:|:---:|:---:|:---:|---:|
| Amsterdam | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | **7** |
| New York City | – | – | – | ✔ | ✔ | ✔ | – | **3** |
| Rio de Janeiro | – | – | **✔** | – | – | – | – | **1** |
| Bangalore | – | ✔ | – | – | – | – | – | **1** |
| Web Sales Centre | **–** | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | 6 |

**Every store location I inferred from the demand patterns is confirmed exactly.** One correction: **MILC Bikes has only one store** (Amsterdam), not two — so MILC runs a genuine web-first model, taking 54.4% of its demand online from a single physical location.

## Correction: Amsterdam is the *most* contested city

I earlier read Amsterdam as "genuinely open" because no rival exceeds 16.5% share. That was backwards. **All seven firms hold an Amsterdam store** — nobody exceeds 16.5% precisely because all seven are there.

| | Amsterdam | New York City |
|---|---:|---:|
| City demand | 2,985 | 2,039 |
| Store firms present | **7** | **3** |
| Avg demand per store firm | **426** | **579** |

A store in New York is worth *more* than deeper investment in Amsterdam. Amsterdam is 2,985 units split seven ways; New York is 2,039 units split three ways. Gaining Amsterdam share means prying it from six store-equipped rivals.

## A store never captures the whole city

| City | Store firms | Store-firm capture | Web-only rivals take |
|------|------------:|-------------------:|---------------------:|
| Amsterdam | 7 | 100% | 0 |
| New York City | 3 | 85.2% | 301 (14.8%) |
| Rio de Janeiro | **1** | **34.9%** | **504 (65.1%)** |
| Bangalore | 1 | 48.6% | 391 (51.4%) |

## Rio: we under-exploit our own monopoly

| | WeBike in Rio | LiteCycle in Bangalore |
|---|---:|---:|
| Sole store holder | yes | yes |
| City demand | 774 | 761 |
| Units | 270 | **370** |
| Share | **34.9%** | **48.6%** |

Two nearly identical structural positions — sole store in a ~770-unit city — and LiteCycle extracts **13.7 more points of share**. Some of that is timing, since Rio only opened in Q3 and we stocked out. But **65% of our exclusive city still went to rivals with no store there at all**, worth roughly 106 units if we merely matched LiteCycle's extraction rate.

## This makes the web case stronger, not weaker

BB LLC pulled **139 units out of Rio with no store there** — 51% of what we managed *with* a store. SpaceBikes took 120, MILC 107, LiteCycle 106.

Meanwhile we pull only 38 from New York and 24 from Bangalore by web. Rivals are three to five times better at the same manoeuvre.

Web reach isn't a consolation prize for cities you skip — it's how the strong firms raid cities they never entered. Our 19.1% web mix is costing us in **all four cities, including the two where we have stores**.

## Q4 sequencing

1. **Web rebuild** — cheapest demand, reaches all four cities, pays back in-quarter (≈ +$89,600). Rivals have proved web raids cities you never enter.
2. **Rio depth** — we're the only store there and extract 34.9% against LiteCycle's 48.6% in the same position. Closing that gap is ~106 units with no new lease.
3. **New York store** — biggest single prize (2,039 units, only three store rivals). Enter through **Speed** once Swift Bike is fixed to 77; BB LLC holds 82.8% of NYC Mountain.

Not priorities: deepening Amsterdam against six store-equipped rivals, or opening Bangalore against LiteCycle's entrenched 48.6%.

## Open decisions for Q4

- [ ] Operating capacity + overtime
- [ ] Compensation (sales + production) vs newly visible industry averages
- [ ] Stock issue: yes/no
- [ ] City #3
- [ ] Speed brand: invest vs refocus
- [ ] Prices, ads, media, SEM, hiring
- [ ] Pro forma + cash flow before advancing
