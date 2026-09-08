# Q4 — Invest in the future

**Plan:** `copy-the-leader.md` — the complete 30-item change list built by cloning BB LLC (Q3 BSC 22, same two segments as us, byte-identical Mountain bike at 46.3% share vs our 22.1%). Workbook: `../Q3/Q3Data.xlsx` → `Q4_Copy_The_Leader`.

Nothing below is locked. Majority vote required before entry.

## Manufacturing
- [ ] Schedule operating capacity **24/day** (was 8) → ~1,154 units, covers the 1,171 forecast
- [ ] Overtime **0.00** — do NOT copy BB LLC's 6.90 (costs Manufacturing Productivity 1.000 → 0.856)
- [ ] Buy **no printers** — we already own BB LLC's 24/day (only NYC would need one)

## Products & pricing
Design model re-validated **16 of 16 exact** on every Mountain/Speed brand: `_verify_brand_designs.py`. Swift Bike needs **exactly two** changes; Hike Bike needs **zero**.
- [ ] Swift Bike gears **24-speed (3×8) → 14-speed (2×7)** (+3)
- [ ] Swift Bike **add colourful thin brushstroke decals** (+2) → **77** = industry ceiling, beats BB LLC's Blu Tube at 76
- [ ] Swift Bike: **keep** aero frame · racing tires · precision brakes · basic drop-down bars · polymer gel racing seat · reflectors (+1) · lights (+1) · no carrier · no suspension
- [ ] **Read the component cost delta on the Design Brand screen** — 14-speed (2×7) should be cheaper than 24-speed (3×8), decals add a little. Net effect unconfirmed (no component price list in our data).
- [ ] Hike Bike design **unchanged** — 73 of 73. Every identifiable change is negative (reflectors −1, comfort seat −3) or untested (lights, basket, removing decals)
- [ ] Swift Bike price **$1,450 → $1,580** (BB LLC's exact Speed price, zero resistance)
- [ ] Hike Bike price **hold $1,365**
- [ ] Priority unchanged: Hike Bike 1, Swift Bike 2
- [ ] No rebates
- [ ] **No 3rd brand in Q4** — analysis in `../Q3/Q3Data.xlsx` → `Q4_Brand_Count`. Plan is already 1,171 demand vs 1,154 capacity; a 2nd Speed brand adds ~363–397 units → 25% stock-out and ~12% Q5 ill will, and needs a 4th printer (~$240,000). Brand count explains no firm's share — Bike Bros runs 2 Speed brands for 6.3% while SpaceBikes runs 1 for 20.4%.

## Advertising & media
- [ ] Swift Bike ad **rebuild** → target 77 (7 ranks in `copy-the-leader.md`; drop lights/reflectors/great price) — **preview the score**
- [ ] HikeBike 1: try **"Highest rated Mountain bike" + "Local sales & service"** to clear 82 → Mountain SERP #1 (+122 clicks). **Only commit if the preview clears 82**
- [ ] No second ad for either brand
- [ ] Media: **HikeBike 1 Biking ×12** and nothing else
- [ ] Media: **Swift Bike Biking ×12** and nothing else
- [ ] Total 24 inserts / **$108,000** — five more inserts than Q3 for **$8,000 less**
- [ ] Organic SEM: 2 pages at $1,000, unchanged (rank follows ad judgment, so the ad fixes are the SEM fix)

## Channel
- [ ] Toll-free phone **$6,000 → $9,000**
- [ ] Advanced shopping cart **$0 → $7,000**
- [ ] Page upgrades **$0 → $9,000** (we stopped this in Q3)
- [ ] Order tracking **$0 → $8,000**
- [ ] Web staff **3 → 7** (5 sales + 2 service)
- [ ] Amsterdam: train the untrained person into **Speed** → 1 Svc / 3 Mtn / 3 Speed (at cap, cannot hire)
- [ ] Rio: train the untrained person into **Mountain**
- [ ] Rio: hire **+1 Mountain, +2 Speed** → 1 Svc / 3 Mtn / 3 Speed; store force 11 → 14 = BB LLC's 14

## People & finance
- [ ] Sales comp **$22,000 / Expanded / 2 wk / 4% = $27,402** (BB LLC's exact package, 78% productivity)
- [ ] Production comp **$18,500 / Expanded / 2 wk / 4% = $23,042**
- [ ] Planner productivity stays **74%** — never 85%
- [ ] Issue **no stock** ($1,010,838 already idle)
- [ ] Target segments unchanged: Mountain primary / Speed secondary

## R&D — 3 project slots (analysis `../Q3/Q3Data.xlsx` → `Q4_RD`, script `_analyze_rd.py`)
**Framing:** every rated brand in the industry is **"Acceptable" (70–77)**. Nobody has reached Good (78–84), Very Good (85–94) or Excellent (95–100). Component ceilings are Mtn 73 / Speed 77 / Rec 76 — all one band. **R&D is the only route into three unclaimed bands.**
**Decision rule:** 3 slots, 2 bikes → prefer features that upgrade **both** bikes. Only the material, gears, decals and accessory slots are segment-neutral; tires/bars/suspension are locked by each segment's recipe.
- [ ] **Slot 1 — Decals: bright, styled per segment & frame** · $109,155 · $8/unit · needs score **929 (highest on the list)** · best value at 851 per $100k · decals are a **measured** +2 in Speed
- [ ] **Slot 2 — Gears: 11 speed (1×11)** · $568,514 · $50/unit · score **905** · hits Mountain's **#1** need (incline 131) + Light weight (121/133) + Shifts smoothly (118/118) · gears are our **strongest measured** component (16-pt swing in Speed) · **highest-variance pick**
- [ ] **Slot 3 — Enriched carbon fiber** · $1,023,325 · $40/unit · score 739 · Light weight (Speed **#2**, 133) + Durable (Mtn **#3**, 127) + Competitive advantage (Spd 124) · most on-strategy item on the board · **worst value per dollar (72/$100k)** — justified by the 50–75% spend mandate
- [ ] **Schedule all three 1-quarter (rapid)** — we are slot-constrained, not cash-constrained. Features land in Q5 designs, and Q5's expanded engineering capacity stays free.
- [ ] **Raise the full $2,500,000 VC.** Reconciles with the "stop issuing stock" rule, which was conditional on cash sitting *idle*; R&D is the productive use.
- [ ] Read the **1-quarter vs 2-quarter** prices after selecting a quarter — listed figures are "Cost to finish" with no quarter chosen, and rapid costs more in total
- [ ] Confirm the **incremental** unit cost on the Design Brand screen — all three are **replacements**, so the true add is less than the $98 gross
- [ ] **ASK THE TEAM: how many quarters remain?** If the sim ends at Q6, 2-quarter scheduling would never ship. Only missing input.
- Alternatives if the team wants less risk or less cash out: **puncture slime** ($159,184, $3/unit, score 640, both bikes) · **racing tires** ($254,694, score 602, Speed only) · **power straight bars** ($181,924, score 509, Mountain only)
- Excluded: mesh tote basket (**Comfort frames only — fits neither bike**) · full suspension (**not Aerodynamic-compatible**, Hike Bike only, $85/unit) · hybrid tires + pedal charger (Recreation) · bike computer ($60/unit, navigation only 117/118)

## Deferred (tier 2)
- [ ] **NYC store** — needs a printer + 7 hires; enter via **Speed** (880 units, no dominant firm), not Mountain (BB LLC holds 82.8%)
- [ ] Business plan + seek up to **$2,500,000** VC
- [ ] **Brand 3 — second Speed brand** (follows a printer): aero frame · racing tires · precision brakes · basic drop-down bars · polymer gel racing seat · **14-speed · decals · lights · NO reflectors** = **76** (Blu Tube, The Armstrong and Skim MILC all converged on this exact spec). Price **$1,580**, same as Swift Bike. **No ad** — BB LLC leaves both second brands unadvertised and leads the industry.
- [ ] Brand 4 — second Mountain brand: Hike Bike's recipe with the **comfort seat** instead of all-purpose = 70. Price $1,365. No ad. Lower priority; Mountain is the smallest segment and we are already #2.
- [ ] Brand 5 — Recreation (MountainCruise1's exact 76): comfort frame · hybrid tires · **standard disc** brakes · comfort straight bars · 7-speed · comfort seat · reflectors + decals + lights + basket + front shocks. **Fights our margin strategy** — every Rec winner is priced at judgment 100, i.e. cheapest in the game ($950–$1,050 net).
- [ ] City #4

## Before advancing
- [ ] Run **Production Simulation**
- [ ] Run **Cash Flow** + pro forma income statement and balance sheet
- [ ] Confirm ending **Cash + CD ≥ $300,000**
- [ ] Verify the estimated fees: web tactic setup (~$27,000), hiring + design (~$40,000), printer (~$240,000 if NYC)
