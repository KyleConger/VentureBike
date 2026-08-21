# Q1 — Manufacturing & Capacity Reference

**Workbook:** `Q1Data.xlsx` → **Manufacturing**, **Stores_Costs**  
**Facility location:** **Bangalore** (sim-wide; favorable incentives; **cannot relocate or add plants later**)

---

## Margin clarification (important)

Earlier strategy language about “high-margin segments” used **Price Willing to Pay** from the survey:

| Segment | Ideal-brand WTP ceiling |
|---------|-------------------------|
| Recreation | $1,100 |
| Mountain | $1,365 |
| Speed | $1,580 |

That is a **price ladder / willingness-to-pay proxy**, **not** proven contribution margin.  
We do **not** yet have official **COGS / unit cost / contribution margin** from the simulation. Do not treat WTP gaps as dollar profit until cost-of-production data exists.

---

## Production process

- Single plant receives all orders daily (custom dimensions / preferences).  
- Build-to-order; **no inventory** at plant; transit inventory ≤ ~2 days; **same-quarter fulfillment**.  
- Ship by **airfreight**; **firm pays shipping** (not customer).  
- Destination: store pickup if a store exists in that city; else break-bulk to buyer.

## Fixed vs operating capacity

| Concept | Meaning |
|---------|---------|
| **Fixed capacity** | # of 3D printers installed (capital asset + depreciation) |
| **Operating capacity** | How much of fixed capacity you run (printers on/off → workers) |
| Constraint | Operating capacity **≤** fixed capacity |

**Printer throughput (normal hours):** 1 frame/hour × 8 × 5 × 13 = **520 units/printer/quarter**.

### Fixed capacity menu

| Printers | Units/day | Units/quarter | Capital $ |
|---------:|----------:|--------------:|----------:|
| 0 | 0 | 0 | 0 |
| 1 | 8 | 520 | 240,000 |
| 2 | 16 | 1,040 | 480,000 |
| 3 | 24 | 1,560 | 720,000 |
| 4 | 32 | 2,080 | 960,000 |
| 5 | 40 | 2,600 | 1,200,000 |
| 7 | 56 | 3,640 | 1,680,000 |
| 10 | 80 | 5,200 | 2,400,000 |

---

## Store & web costs

| Outlet | Setup | Quarterly lease |
|--------|------:|----------------:|
| New York City | 200,000 | 62,000 |
| Rio de Janeiro | 90,000 | 17,000 |
| **Amsterdam** | **136,000** | **44,000** |
| Bangalore (store) | 84,000 | 13,000 |
| **Web center (Bangalore)** | **150,000** | **60,000** |

---

## Fixed capacity review (with these costs)

**Q1 cash start:** $1,500,000 · **Floor:** Cash + CD ≥ $300,000  
**Committed channel (locked):** Amsterdam setup **$136,000** (lease $44k/qtr once open)

### Major-outlay sketches (ignores other Q1 costs)

| Scenario | Printers | Printer $ | AMS | Web | Total | Cash left | Floor OK? |
|----------|---------:|----------:|----:|----:|------:|----------:|-----------|
| AMS + 2 printers | 2 | 480k | 136k | — | 616k | **884k** | Yes |
| AMS + 3 printers | 3 | 720k | 136k | — | 856k | **644k** | Yes |
| AMS + web + 2 | 2 | 480k | 136k | 150k | 766k | **734k** | Yes |
| AMS + web + 3 | 3 | 720k | 136k | 150k | 1,006k | **494k** | Yes |
| AMS + 5 printers | 5 | 1,200k | 136k | — | 1,336k | **164k** | **No** |

### Demand vs capacity (order of magnitude)

- Amsterdam Rec **12-mo potential** 5,468 → naïve ÷4 ≈ **1,367**/quarter — **early actual demand will be much lower**.  
- 1 printer = 520/qtr · 2 = 1,040 · 3 = 1,560.  
- For a **single-city Q2 test**, **2 printers** usually covers a cautious Rec-led launch; **3** adds headroom if web opens or demand surprises.  
- **5+** in Q1 is rarely justified before test results and **fails** the $300k floor with AMS alone at 5 printers under this sketch.  
- Remember: fixed capacity leads demand; you can add printers later (lead time applies in-sim).

### Distribution-cost note (Bangalore fab → Amsterdam store)

Plant is fixed in **Bangalore**; first store in **Amsterdam** means **airfreight on every bike**. That is a real cost drag we still need line-item rates for—another reason not to equate WTP with margin. Choosing Bangalore store later can shorten some lanes; it does not move the plant.

---

## Decision posture (not locked)

- **Printers:** favor **2 or 3** in Q1 pending team vote and final web yes/no.  
- **Do not** size fixed capacity to full annual potential ÷ 4.  
- Revisit after Q2 demand/COGS appear—then “margin leader” talk can become quantitative.
