# Compensation & Productivity — Background (Q2 Lecture)

**Source:** Marketplace Workspace (sales force compensation; same productivity logic for production).  
**Status:** Sales force **and** production packages **LOCKED** (Q2).  
**Related:** `Q2Data.xlsx` → **Comp_Industry** · Q1 baselines in `../Q1/compensation-industry-avg-sales.md` and `compensation-industry-avg-production.md` · team goal **pay > average**.

## Locked — Q2 Sales Force (World Market / sales people)

| Component | Our setting | Cost | Industry |
|-----------|-------------|-----:|----------|
| Annual salary | **$19,000** | $19,000 | $17,000 |
| Health benefits | **Full coverage** | $4,180 | Minimum $1,700 |
| Vacation | **2 weeks** | $1,055 | 1 week $440 |
| Pension | **1%** | $190 | 1% $170 |
| **Total / year** | | **$24,425** | **$19,310** |
| Projected productivity | | **85%** | ~70% norm |

Premium vs industry: **+$5,115 (~+26.5%)**. Applies to sales personnel including World Market web staff (Sales_Ops annual comp = $24,425).

## Locked — Q2 Production Workers (World Market)

| Component | Our setting | Cost | Industry |
|-----------|-------------|-----:|----------|
| Annual salary | **$16,800** | $16,800 | $14,000 |
| Health benefits | **Expanded coverage** | $2,520 | Minimum $1,400 |
| Vacation | **2 weeks** | $933 | 1 week $363 |
| Pension | **3%** | $504 | 2% $280 |
| **Total / year** | | **$20,757** | **$16,043** |
| Projected productivity | | **85%** | ~70% norm |

Premium vs industry: **+$4,714 (~+29.4%)**. Syncs to Manufacturing worker productivity (Manufacturing!B30 = 85%).

## Philosophy

Employees (yours and competitors’) always seek better compensation. Each quarter they are surveyed on which package area they most want improved.

- **Early quarters:** strong push for **base salary** gains.  
- As salaries rise, importance shifts toward **health**, **vacation**, and **pension**.  
- Importance scores change over time — re-read Workspace each quarter.

Competition for the best people is intense. Recruiting/retention and **motivation → demand creation** improve the further you sit **above industry norm**. Below or at norm risks losing talent and productivity to peers who raise pay.

## What you decide each quarter

For **sales people** (and separately for **production workers**):

1. Set package: annual salary · health benefits · vacation · pension % of salary.  
2. Freely change next quarter.  
3. **Estimate productivity** for the quarter (educated guess) — shared with the sales / production executive.

### Productivity norms

| Item | Value |
|------|------:|
| Current industry productivity norm (early) | **~70%** |
| If you improve vs industry | May estimate **above** 70% |
| If you make no adjustment | May fall **below** 70% if competitors raise pay |
| Last quarter’s industry average package | Available starting **Q3** (not yet for Q2 planning baseline — use Q1 industry snapshot) |

Same productivity starting point and Q3 industry-average visibility apply to **sales and production**.

## Sales force — industry snapshot (entering Q2)

| Component | Industry setting | Cost | Importance (improve) |
|-----------|------------------|-----:|---------------------:|
| Salary | $17,000 | $17,000 | **87** |
| Health | Minimum | $1,700 | **84** |
| Vacation | 1 week | $440 | **72** |
| Pension | 1% | $170 | **70** |
| **Total** | | **$19,310** | |

Priority for dollars (early): **Salary → Health → Vacation → Pension**.

## Sales — changeable metrics & costs (Workspace)

### Annual salary
- Dollar amount = dollar cost (direct).

### Health benefits package

| Tier | Cost |
|------|-----:|
| No health benefits | $0 |
| Minimum | $1,700 |
| Expanded | $2,550 |
| Full coverage | $3,740 |
| Comprehensive | $5,610 |

### Weeks of vacation

| Weeks | Cost |
|------:|-----:|
| 0 | $0 |
| 1 | $440 |
| 2 | $944 |
| 3 | $1,475 |
| 4 | $2,024 |
| 5 | $2,587 |
| 6 | $3,161 |
| 7 | $3,745 |

### Pension
- Percent of salary → cost = salary × (pension % / 100).

**Total yearly cost** = Salary + Health + Vacation + Pension.

## Production workers

- Separate industry package (Q1 snapshot): Salary $14,000 · Health Minimum $1,400 · Vacation 1 wk $363 · Pension 2% $280 · **Total $16,043**. Importance order same (87 / 84 / 72 / 70).  
- Q2 Workspace costs used in lock: Expanded health **$2,520** · 2 weeks vacation **$933** (differ from sales menus).  
- Productivity estimate starts at **~70%**; locked Q2 estimate **85%** feeds Manufacturing operating-capacity pad.

## Interactive planning approach (team method)

1. Lock a **productivity target** (e.g. 75%, 80%) as a performance goal hypothesis.  
2. Build packages that clear industry total **and** weight spend to high-importance levers (salary first early).  
3. Compare **$/productivity point** across scenarios in `Comp_Industry`.  
4. Use **Importance-Weighted Value Index** — default weights Salary **87** · Health **84** · Vacation **72** · Pension **70**:  
   - **OUR Planning Weights** (yellow `F68:F71`) are editable; Score mode `B73` = `Our` or `Industry`  
   - Value Points = max(0, uplift % vs industry) × Active Importance  
   - Package Value Index = Σ points / Active total × 100  
   - Allocate “extra $ above industry” by active weight share  
   - **History tables + charts** log Our/Industry weights and Value Index / productivity by quarter (Q1–Q6)  
5. Stress cash: headcount × package vs Cash+CD ≥ $300k and other Q2 spends.  
6. Revisit after classmates’ packages appear (Q3 industry averages); paste prior-quarter snapshots into History before changing packages.

Do **not** invent final numbers — yellow cells in workbook until majority vote.  
Planner location: `Q2Data.xlsx` → **Comp_Industry** (rows 65+ value scorecards + budget allocator).

## Open decisions

- [x] Sales package (salary / health / vacation / pension) — **LOCKED** $24,425 · 85% productivity  
- [x] Sales productivity estimate (%) — **85%**  
- [x] Production package — **LOCKED** $20,757 · 85% productivity  
- [x] Production productivity estimate (%) — **85%**  
- [ ] Service package when role unlocks  

## Links

- Workbook planner: `Q2Data.xlsx` → **Comp_Industry**  
- BSC: HR = (sales productivity + production worker productivity) / 2  
- Manufacturing OC uses production productivity (see `manufacturing-background.md`)
