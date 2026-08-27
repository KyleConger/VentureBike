# Q2 Scenario Rerun — Mountain only vs Mountain + Speed (lean manufacturing)

**After:** `manufacturing-background.md`  
**Fixed capacity Q2:** 24 units/day · 65 days · max OT ~+50%  
**Productivity assumption:** 70% → capacity pad **×1.30** (lecture formula)  
**Extra Speed line cost:** $30,000  

Same demand/share/price assumptions as prior blue/red analysis, now with **operating capacity** math.

---

## Shared market assumptions

| | Mountain | Speed |
|--|--------:|------:|
| Naïve quarterly potential | 3,455 | 2,969 |
| Early realization (15% of naïve) | **518** industry units | **445** industry units |
| Price | $1,300 | $1,420 (red-ocean pressure) |
| Base CM (assumption) | 40% | 32% |

| Scenario | Mtn share | Speed share |
|----------|----------:|------------:|
| **A — Mtn only (blue)** | 45% | — |
| **B — Mtn blue + Speed red** | 38% | 12% |

---

## Step 1 — Expected sales units

| Scenario | Units |
|----------|------:|
| A Mountain | \(0.45×518 = \mathbf{233}\) |
| B Mountain | \(0.38×518 = \mathbf{197}\) |
| B Speed | \(0.12×445 = \mathbf{53}\) |
| **B total** | **250** |

---

## Step 2 — Required daily operating capacity (lean formula)

\[
OC_{daily} = \frac{\text{forecast units}}{65} × 1.30
\]

| Scenario | Forecast units | Raw /65 | ×1.30 OC needed | vs Fixed 24 | Feasible w/o OT? |
|----------|---------------:|--------:|----------------:|------------:|------------------|
| A Mountain only | 233 | 3.58 | **4.7**/day | 24 | **Yes** (lots of headroom) |
| B Combined | 250 | 3.85 | **5.0**/day | 24 | **Yes** |

**Read:** At 15% early realization, **neither strategy is capacity-constrained**. Fixed 24/day is far above need. The fight is **forecast error, excess capacity cost, and unit cost at low volume** — not printer shortage.

*(If realization jumped to ~50% of naïve quarterly: Mtn industry ~1,728 → A at 45% = 778 units → OC ≈ 778/65×1.30 ≈ **15.6**/day — still OK. B at 250→ scaled ~833 → OC ≈ **16.7**/day — still OK. Constraint binds nearer full naïve shares.)*

---

## Step 3 — Revenue & contribution (unchanged core economics)

| | A Mtn only | B Mtn+Speed |
|--|----------:|------------:|
| Revenue | 233×1300 = **$302,900** | 197×1300 + 53×1420 = **$331,360** |
| Contribution @ 40%/32% CM | **$121,160** | **$126,539** |
| − Speed line cost | — | −$30,000 |
| **Net contribution proxy** | **$121,160** | **$96,539** |
| **B − A** | | **−$24,621** |

---

## Step 4 — Lean manufacturing overlay (where B gets worse)

### Excess capacity risk (underestimate Speed / overestimate total)

Suppose you schedule OC for B’s **250** unit forecast (5.0/day) but Speed sells **0** and Mountain only hits **197**:

- Paid operating capacity sized for 250; produce ~197  
- Idle ≈ **53 units of capacity** → **excess capacity cost** (workers paid, printers underused)  
- You still paid **$30k** for the Speed line  

Mountain-only schedules ~4.7/day for 233; miss is one-stream error only.

### Stockout / OT path (Speed unexpectedly hot)

If Speed is less “red” and B demand = 400 units:

- OC needed = 400/65×1.30 ≈ **8.0**/day — still &lt; 24  
- Still not fixed-capacity bound; OT optional  

Fixed capacity only bites if total served demand approaches **~24×65/1.30 ≈ 1,200** units/qtr before OT (or ~1,800 with heavy OT).

### Unit cost at low volume (lecture point)

- Early PLC: **production cost high vs price**.  
- **A** puts all volume on Mountain → slightly better chance to sit on the **efficient** part of the OC cost curve for one SKU.  
- **B** only lowers unit cost if **total** volume rises **and** you don’t leave paid OC idle when Speed fails.

---

## Step 5 — Verdict (with manufacturing recorded)

| Lens | Winner |
|------|--------|
| Revenue | B (+$28k) |
| Contribution after $30k line | **A (−$25k for B)** |
| Capacity feasibility @ 15% realization | Tie (both fine) |
| Lean / forecast risk | **A** (one pull stream; less excess-capacity exposure) |
| Margin-leader strategy | **A** |
| Learning / Q3+ option on Speed | B (strategic, not this-quarter math) |

### Bottom line
Under make-to-order + 70% productivity planning, **math still favors Mountain-only this quarter**. Diversifying into Speed adds a second pull stream and **$30k**, while expected Speed volume in a red ocean does **not** cover dilution + line cost — and raises **excess capacity** risk if you staff for Speed that does not appear.

**Revisit Speed when:** forecast confidence is higher, Mountain share won’t rise with focus, or you explicitly buy learning with OT/low OC (schedule Mtn-centric OC, small Speed upside via OT) rather than full dual forecast staffing.
