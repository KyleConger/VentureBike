# Demand Projection — Background (Q2 Lecture)

**Source:** Marketplace Workspace (Sales demand projection — current + next quarter).  
**Status:** Background + planner — **units per salesperson and brand split not locked**.  
**Related:** `Q2Data.xlsx` → **Demand_Forecast** · Sales_Ops · Manufacturing · Cost_of_Production · Pricing.

## Why it matters

Manufacturing needs demand to set **operating capacity** (this Q) and **fixed capacity / printers** (next Q).  
Marketing needs demand to set **prices** vs Cost of Production at that volume.  
Firm success depends heavily on forecast accuracy.

## Core formula

\[
\text{Demand Projection} = (\text{units per salesperson}) \times (\text{\# of sales people employed for the quarter})
\]

Estimate units/salesperson, multiply by headcount → **total demand**, then **allocate across brands** (sum of brand demands = total). Larger/target-priority segments get larger shares.

Workspace then shows **operating capacity required** to satisfy projected demand (units/quarter and units/day).

## This quarter (Q2 test market)

### Task
1. Enter units each salesperson can sell (headcount already from hiring).  
2. Allocate total across brands on sale.  
3. Note OC required.

### Decision tip (conservative)
- Consumers don’t know you yet; tactics may miss → **be conservative**.  
- Guide: **30–70 units per salesperson**.  
- Example: 6 people → ~180–420 units.  
- **Safe starting point: 40 units/person** (e.g. 6 people → **240** units).  
- Actual depends on # brands, marketing quality, segment/geo potential.  
- Demand should rise as experience, designs, ads, and sales management improve.

### Stock-out / ill-will penalty
If you underestimate and OC is too low → stock-outs → ill-will penalty **next quarter**:

1. Compute % unmet demand this quarter.  
2. Take **half** of that % as the penalty.  
3. Next quarter’s demand is reduced by that percent.

Overestimate → excess capacity costs (workers paid idle). Underestimate → lost sales + next-Q demand haircut.

## Next quarter (Q3) demand

Used to plan **printer / fixed capacity** purchases (one-quarter lead time).

### Formula
Same: units/salesperson × salespeople (for next quarter).

### Task
1. Decide expansion (new stores / web) **before** forecasting.  
2. Project headcount per open + planned outlet (4–8+ per outlet — pick reasonable).  
3. Sum store people + web people.  
4. Project demand per salesperson.

### Decision tip (second test market)
- More awareness + better decisions from Q2 feedback → higher sales/person likely.  
- Guide: **40–80 units/person** (perhaps a bit more).  
- **Reasonable start: 60 per salesperson**.  
- Underestimate → not enough printers → stock-outs. Overestimate → expensive excess printers.

## WeBike Workspace snapshot (entering plan)

| Block | Last Q (actual) | This Q (projected) |
|-------|-----------------|-------------------|
| Stores people / demand/SP / total | 0 / 0 / 0 | Yellow (link AMS headcount) |
| Web people / demand/SP / total | 0 / 0 / 0 | Yellow (link World Market) |
| Brand: Hike Bike | n/a | Allocate share of total (100% if only brand) |
| Outlets this → next | Stores 1 · Web 1 | Next: Stores 1 or 2 · Web 1 (if city #2) |

**Changeable (this Q):** demand per salesperson (stores, web); brand projected demand.  
**Changeable (next Q):** number of sales people; demand per salesperson (stores, web).

## Links

- Planner: `Q2Data.xlsx` → **Demand_Forecast**  
- Headcount: Sales_Ops (AMS + Web Personnel)  
- OC / printers: Manufacturing  
- Unit cost at volume: Cost_of_Production · Pricing
