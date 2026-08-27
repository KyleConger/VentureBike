# Q2 Manufacturing Background — WeBike

**Source:** Simulation lecture / Workspace (Lean Manufacturing, Operating Capacity, Overtime, Fixed Capacity)  
**Filed:** 2026-08-27 · Entering test market (Q2)  
**Status:** OC **LOCKED = 20**/day · OT **LOCKED = 0** · worker productivity **85%** (from production compensation).

---

## Lean philosophy

- Build **to demand**, not to forecast-driven inventory (“pulled manufacturing”).
- WeBike’s plant is **make-to-order / JIT**: orders → fab → ship; **no warehouse stock**.
- No inventory at plant, in transit (&gt;2 days), or at store → less cash tied in supply chain; all fulfilled **same quarter**.

### Daily flow

1. Sales orders (with custom frame dimensions) hit the **Bangalore** facility.  
2. 3D printers produce carbon frames to order.  
3. End of day: box, containerize, ship airfreight → store pickup or break-bulk delivery.  
4. Firm pays shipping.

---

## Brands → production

- If Marketing **prices a brand for sale**, it **automatically** enters the production plan.
- Manufacturing does not separately “choose” brands; it produces what Marketing sells.
- **Implication:** adding a Speed line ($30k) also adds a second demand stream Manufacturing must capacity-plan for.
- **Demand input:** Sales `Demand_Forecast` (units/SP × people → brand split). See `demand-projection-background.md`.

---

## Fixed vs operating capacity

| Term | Meaning |
|------|---------|
| **Fixed capacity** | Max frames/day from printers on hand. WeBike Q2: **24 units/day** (3 printers; 1,560/qtr at 8×5×13 without OT) |
| **Operating capacity** | How many units/day you **schedule** this quarter (≤ fixed). Sets workers and printer run-time |
| **Lead time** | Printers ordered **this** quarter available **next** quarter |

### Worker productivity (early quarters)

- Default early estimate ≈ **70%** productivity under standard compensation (HR can revise).
- Must **inflate** operating capacity for productivity loss.
- Effective OC after productivity = scheduled OC × productivity %.
- Needed OC for demand ≈ (Demand ÷ 65) × (1 + loss).

### Workspace Operating Capacity table (Q2)

| Metric | Units/Day | Units/Quarter |
|--------|----------:|--------------:|
| Fixed capacity (previous Q) | 0 | 0 |
| Operating capacity (previous Q) | 0 | 0 |
| Fixed capacity | **24** | **1,560** |
| Operating capacity | **20 LOCKED** | **1,300** |
| Effective OC after productivity | 20 × 85% = **~17**/day | **~1,105**/qtr |
| Effective OC needed for demand | from Demand_Forecast | |
| Projected worker productivity | **85% LOCKED** | |
| New direct labor $/unit OC | ~**$110** at OC 20 (curve) | |
| New overhead $/unit OC | ~**$48** at OC 20 (curve) | |
| Expense to change OC | Workspace | |
| Max overtime | **0 h LOCKED** | |

**Changeable (were):** daily OC / OT — now locked for Q2. See `Q2Data.xlsx` → Manufacturing · chart `assets/operating-capacity-costs-curve.png`.  
Demand ceiling without OT at this lock ≈ **1,130** units/qtr (`20/1.15 × 65`).

### Unit cost vs OC (Workspace curve)

- Both labor and overhead start very high at low OC, then fall (economies of scale).
- **Overhead** bottoms ~**$45–50**/unit by ~16–24/day (flat).
- **Labor** is U-shaped: trough ~**$110–120** around **18–22**/day; rises again near fixed max (**~$190** at 24/day).
- Planning: prefer scheduling near the labor trough unless demand/OT strategy requires otherwise; running at 24 maximizes capacity but **raises labor cost on every unit**.

**Lecture formula:**

\[
\text{Daily operating capacity} =
\frac{\text{Total forecasted demand (all brands)}}{65}
\times \bigl(1 + \text{proportional productivity loss}\bigr)
\]

Example: 70% productivity → ~**30% loss** → multiply demand/65 by **1.30**.  
Operating capacity **cannot exceed fixed capacity (24/day)** without using **overtime**.

---

## Cost behavior

- Labor + overhead **per unit** generally **fall** as operating capacity rises (efficiency).
- Near **100% of fixed capacity**, crowding can make unit costs **rise** again (U-shaped).
- Early PLC: **production cost high vs price** because volume is low; unit cost improves as volume grows.
- Tradeoff: max OC (more revenue, possibly higher unit cost) vs lower OC (efficiency, more stockouts).

---

## Overtime

| Rule | Detail |
|------|--------|
| Max | Up to **4 hours/day** |
| Pay | **2×** regular (100% premium) |
| Fatigue | Lowers productivity → raises effective unit cost |
| Use | Only as needed if demand &gt; scheduled OC |
| Ceiling | Can support up to ~**50% more** than fixed capacity when pushed |
| If demand &lt; OC | OT unused; risk is **excess capacity cost** (workers paid, idle) |

### Workspace OT productivity (Q2 chart)

| Max OT / day | Productivity |
|-------------:|-------------:|
| 0 | **70.0%** |
| 1 | **69.3%** |
| 2 | **68.0%** |
| 3 | **66.4%** |
| 4 | **64.4%** |

Chart: `assets/overtime-productivity-curve.png` · Tables: `Q2Data.xlsx` → Manufacturing (OT) + **Production_Sim**.

**Decision tip:** Under uncertainty, sometimes schedule **lower OC + OT safety net** vs high OC (excess capacity risk). Re-run Production_Sim at 50% / 100% / 150% demand (`Production_Sim!B3`).

### Production simulation outputs to watch

- Units produced vs net demand → **lost sales / stock-outs** → next-Q ill-will = ½ unmet %  
- Labor + overhead + materials vs selling price  
- Utilization vs **excess capacity** charges  
- Fixed capacity add (printers) so next-Q fixed ≥ padded next-Q demand  

**Changeable:** OC (Manufacturing!B39), max OT (B53), produce Y/N, demand %, printer add (Production_Sim!B62), price/materials.

---

## Each quarter — manufacturing checklist

1. Confirm brands Marketing will sell (auto on production plan).  
2. Take **total** demand forecast from Sales (all brands).  
3. Take **productivity** estimate from HR.  
4. Set **daily operating capacity** via formula (≤ 24 without OT).  
5. Set **max overtime** allowance.  
6. Run **production simulation** (and ±50% / half-demand scenarios).  
7. Decide **fixed capacity adds** for **next** quarter from next-quarter demand.  
8. Feed results into **pro forma**.

---

## Implications for Mountain-only vs Mountain + Speed

| Factor | Mountain only | Mountain + Speed |
|--------|---------------|------------------|
| Forecast objects | One demand stream | Two streams; total OC from **sum** |
| Make-to-order | Simpler priority | Capacity shared; Mtn priority must be explicit in forecast/sales |
| Productivity pad | Apply 1.30× to Mtn demand/65 | Apply 1.30× to **(Mtn+Speed)** demand/65 — hits **24/day ceiling sooner** |
| Excess capacity risk | Lower if forecast tight | Higher if Speed is red-ocean and undersells |
| Unit cost | All volume on one brand | Higher total volume **if** Speed sells → better efficiency; if not, pay for idle OC |
| Stockout risk | Concentrated | Split attention; OT more likely if both hot |

See `lean-scenario-mtn-vs-mtn-speed.md` for quantitative rerun.
