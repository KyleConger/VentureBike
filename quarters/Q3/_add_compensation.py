"""Add Q3 competitor compensation (exact) and Q4 package options to Q3Data.xlsx."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DST = "quarters/Q3/Q3Data.xlsx"

# company, sat, salary, health_tier, health, vac_weeks, vac, pension_pct, pension, total
SALES = [
    ("MILC Bikes", 0.700, 22000, "Full coverage", 4840, 1, 570, 0.01, 220, 27630),
    ("BB LLC",     0.784, 22000, "Expanded coverage", 3300, 2, 1222, 0.04, 880, 27402),
    ("LiteCycle",  0.728, 21000, "Expanded coverage", 3150, 2, 1166, 0.02, 420, 25736),
    ("SpaceBikes", 0.736, 20000, "Expanded coverage", 3000, 2, 1111, 0.03, 600, 24711),
    ("Spoke'd Up", 0.714, 20000, "Expanded coverage", 3000, 2, 1111, 0.02, 400, 24511),
    ("WeBike",     0.703, 19000, "Full coverage", 4180, 2, 1055, 0.01, 190, 24425),
    ("Bike Bros",  0.685, 18000, "Expanded coverage", 2700, 2, 1000, 0.02, 360, 22060),
]

PROD = [
    ("MILC Bikes", 0.727, 18400, "Full coverage", 4048, 1, 477, 0.02, 368, 23293),
    ("BB LLC",     0.777, 18500, "Expanded coverage", 2775, 2, 1027, 0.04, 740, 23042),
    ("LiteCycle",  0.743, 17500, "Expanded coverage", 2625, 2, 972, 0.03, 525, 21622),
    ("SpaceBikes", 0.734, 17000, "Expanded coverage", 2550, 2, 944, 0.03, 510, 21004),
    ("WeBike",     0.731, 16800, "Expanded coverage", 2520, 2, 933, 0.03, 504, 20757),
    ("Spoke'd Up", 0.731, 16800, "Expanded coverage", 2520, 2, 933, 0.03, 504, 20757),
    ("Bike Bros",  0.696, 16000, "Expanded coverage", 2400, 2, 889, 0.02, 320, 19609),
]

ALL_SALES_AVG = (18643, 2576, 737, 474, 22430)
ALL_PROD_AVG = (15643, 2088, 622, 458, 18811)

wb = load_workbook(DST)
bold = Font(bold=True)
hdr = PatternFill("solid", fgColor="DDEBF7")
ours = PatternFill("solid", fgColor="FFF2CC")
good = PatternFill("solid", fgColor="C6EFCE")
bad = PatternFill("solid", fgColor="FFC7CE")
yellow = PatternFill("solid", fgColor="FFEB9C")
thin = Border(
    left=Side(style="thin", color="D0D0D0"),
    right=Side(style="thin", color="D0D0D0"),
    top=Side(style="thin", color="D0D0D0"),
    bottom=Side(style="thin", color="D0D0D0"),
)

if "Q3_Compensation" in wb.sheetnames:
    del wb["Q3_Compensation"]
ws = wb.create_sheet("Q3_Compensation")

ws["A1"] = "Competitor Compensation — Q3 actual (exact from Workspace)"
ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = (
    "Among the 7 carbon firms WeBike is 6th of 7 on sales pay (only Bike Bros cheaper) and tied 5th/6th on "
    "production. The published metal+carbon average is easier to beat because it includes cheaper metal firms. "
    "Sales mix is the error: salary $19k is 6th of 7 while Full coverage makes health 2nd-highest. "
    "BB LLC (Expanded, 2 wk, 4%) scores 78.4% at $27,402; MILC (Full, 1 wk, 1%) scores 70.0% at $27,630. "
    "Importance still Salary 87 / Health 84 / Vacation 72 / Pension 70. Packages OPEN — not locked."
)
ws["A2"].alignment = Alignment(wrap_text=True)
ws.merge_cells("A2:J2")
ws.row_dimensions[2].height = 64


def block(start, rows, wrap_col=None):
    for i, row in enumerate(rows):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(start + i, c, v)
            cell.border = thin
            if i == 0:
                cell.fill = hdr
                cell.font = bold
            elif wrap_col and c == wrap_col:
                cell.alignment = Alignment(wrap_text=True)
    return start + len(rows) + 2


def roster(title, data, all_avg, start):
    ws.cell(start, 1, title).font = bold
    rows = [[
        "Company", "Satisfaction", "Salary", "Health tier", "Health $",
        "Vacation weeks", "Vacation $", "Pension %", "Pension $", "Total yearly cost",
    ]]
    for co, sat, sal, tier, h, vw, v, pp, p, tot in data:
        rows.append([co, sat, sal, tier, h, vw, v, pp, p, tot])
    n = len(data)
    mean_sal = sum(x[2] for x in data) / n
    mean_h = sum(x[4] for x in data) / n
    mean_tot = sum(x[9] for x in data) / n
    rows.append(["7-firm mean (carbon classmates)", None, round(mean_sal), None, round(mean_h),
                 None, None, None, None, round(mean_tot)])
    rows.append(["All metal + carbon average", None, all_avg[0], None, all_avg[1],
                 None, all_avg[2], None, all_avg[3], all_avg[4]])
    r2 = block(start + 1, rows)
    header_row = start + 1
    for i in range(1, n + 1):
        ws.cell(header_row + i, 2).number_format = "0.0%"
        ws.cell(header_row + i, 8).number_format = "0%"
        for col in (3, 5, 7, 9, 10):
            ws.cell(header_row + i, col).number_format = "#,##0"
        if rows[i][0] == "WeBike":
            for c in range(1, 11):
                ws.cell(header_row + i, c).fill = ours
                ws.cell(header_row + i, c).font = bold
        if rows[i][0] == "BB LLC":
            for c in range(1, 11):
                ws.cell(header_row + i, c).fill = good
    # mean / all-avg rows
    for offset in (n + 1, n + 2):
        for col in (3, 5, 7, 9, 10):
            ws.cell(header_row + offset, col).number_format = "#,##0"
        ws.cell(header_row + offset, 1).font = bold
    ws.cell(header_row + n + 1, 10).fill = yellow
    return r2, mean_sal, mean_h, mean_tot


r = 4
r, sales_mean_sal, sales_mean_h, sales_mean_tot = roster(
    "SALES FORCE — sorted by total cost (high to low)", SALES, ALL_SALES_AVG, r
)

rows = [
    ["SALES READ", "Value", "Note"],
    ["WeBike total rank among 7", "6 of 7", "only Bike Bros cheaper; we trail the 7-firm mean by $786"],
    ["WeBike salary rank among 7", "6 of 7", "importance 87 — this is the starved lever"],
    ["WeBike salary vs 7-firm mean", 19000 - sales_mean_sal, "Full coverage is why health is $727 above the 7-firm mean"],
    ["WeBike health vs 7-firm mean", 4180 - sales_mean_h, "2nd-highest health line; only MILC (also Full) is higher"],
    ["WeBike satisfaction rank", "5 of 7", "70.3% — MILC pays most and sits at 70.0%"],
    ["Satisfaction leader", "BB LLC 78.4%", "Expanded + 2 weeks + 4%, not Full coverage"],
    ["Expanded health formula", "15% of salary", "reproduces every Expanded row exactly"],
    ["Full coverage formula", "22% of salary", "reproduces WeBike $4,180 and MILC $4,840 exactly"],
    ["2-week vacation formula", "salary / 18", "reproduces every 2-week row (rounded)"],
]
r = block(r, rows, wrap_col=3)
ws.cell(r - 9, 2).fill = bad
ws.cell(r - 8, 2).fill = bad
ws.cell(r - 7, 2).number_format = "+#,##0;-#,##0"
ws.cell(r - 6, 2).number_format = "+#,##0;-#,##0"

r, prod_mean_sal, prod_mean_h, prod_mean_tot = roster(
    "PRODUCTION WORKERS — sorted by total cost (high to low)", PROD, ALL_PROD_AVG, r
)

rows = [
    ["PRODUCTION READ", "Value", "Note"],
    ["WeBike total rank among 7", "tied 5th/6th", "byte-identical package to Spoke'd Up; both 73.1% sat"],
    ["Mix vs industry pattern", "already correct", "Expanded + 2 weeks + 3% is the carbon-firm mode"],
    ["Gap vs 7-firm mean", 20757 - prod_mean_tot, "close it by raising salary $16,800 -> $18,000"],
    ["Satisfaction leader", "BB LLC 77.7%", "$18,500 Expanded 2 wk 4% = $23,042"],
    ["MILC again overpays the wrong mix", "72.7% at $23,293", "Full coverage + 1 week, highest production cost, 5th of 7 sat"],
]
r = block(r, rows, wrap_col=3)
ws.cell(r - 4, 2).number_format = "+#,##0;-#,##0"

rows = [
    ["BB LLC vs MILC — mix beats total (sales)", "BB LLC", "MILC Bikes", "Gap"],
    ["Salary", 22000, 22000, 0],
    ["Health", 3300, 4840, -1540],
    ["Vacation", 1222, 570, 652],
    ["Pension", 880, 220, 660],
    ["Total yearly cost", 27402, 27630, -228],
    ["Satisfaction", 0.784, 0.700, 0.084],
]
r = block(r, rows)
for i in range(1, 6):
    for c in range(2, 5):
        ws.cell(r - 7 + i, c).number_format = "#,##0;+#,##0;-#,##0"
ws.cell(r - 1, 2).number_format = "0.0%"
ws.cell(r - 1, 3).number_format = "0.0%"
ws.cell(r - 1, 4).number_format = "+0.0%;-0.0%"
ws.cell(r - 1, 2).fill = good
ws.cell(r - 1, 3).fill = bad

rows = [
    ["Q4 SALES PACKAGES (OPEN — majority vote)", "Current Q3", "A. Reallocate + raise (recommended)",
     "B. Clone BB LLC", "C. Keep Full"],
    ["Salary", 19000, 21500, 22000, 21000],
    ["Health tier", "Full coverage", "Expanded coverage", "Expanded coverage", "Full coverage"],
    ["Health $", 4180, 3225, 3300, 4620],
    ["Vacation", "2 weeks", "2 weeks", "2 weeks", "2 weeks"],
    ["Vacation $", 1055, 1194, 1222, 1167],
    ["Pension %", 0.01, 0.03, 0.04, 0.03],
    ["Pension $", 190, 645, 880, 630],
    ["Total yearly cost", 24425, 26564, 27402, 27417],
    ["vs 7-firm mean $25,211", 24425 - round(sales_mean_tot), 26564 - round(sales_mean_tot),
     27402 - round(sales_mean_tot), 27417 - round(sales_mean_tot)],
    ["vs all-firm avg $22,430", 24425 - 22430, 26564 - 22430, 27402 - 22430, 27417 - 22430],
]
r_sales_opt = r
r = block(r, rows)
ws.cell(r_sales_opt + 1, 3).fill = yellow
ws.cell(r_sales_opt + 8, 3).fill = yellow
for i in (1, 3, 5, 7, 8, 9, 10):
    for c in range(2, 6):
        cell = ws.cell(r_sales_opt + i, c)
        if i != 6:
            cell.number_format = "#,##0;+#,##0;-#,##0"
ws.cell(r_sales_opt + 6, 2).number_format = "0%"
ws.cell(r_sales_opt + 6, 3).number_format = "0%"
ws.cell(r_sales_opt + 6, 4).number_format = "0%"
ws.cell(r_sales_opt + 6, 5).number_format = "0%"

rows = [
    ["Q4 PRODUCTION PACKAGES (OPEN — majority vote)", "Current Q3",
     "A. Raise salary (recommended)", "B. Clone BB LLC"],
    ["Salary", 16800, 18000, 18500],
    ["Health tier", "Expanded coverage", "Expanded coverage", "Expanded coverage"],
    ["Health $", 2520, 2700, 2775],
    ["Vacation", "2 weeks", "2 weeks", "2 weeks"],
    ["Vacation $", 933, 1000, 1027],
    ["Pension %", 0.03, 0.03, 0.04],
    ["Pension $", 504, 540, 740],
    ["Total yearly cost", 20757, 22240, 23042],
    ["vs 7-firm mean $21,441", 20757 - round(prod_mean_tot), 22240 - round(prod_mean_tot),
     23042 - round(prod_mean_tot)],
]
r_prod_opt = r
r = block(r, rows)
ws.cell(r_prod_opt + 1, 3).fill = yellow
ws.cell(r_prod_opt + 8, 3).fill = yellow
for i in (1, 3, 5, 7, 8, 9):
    for c in range(2, 5):
        ws.cell(r_prod_opt + i, c).number_format = "#,##0;+#,##0;-#,##0"
for c in range(2, 5):
    ws.cell(r_prod_opt + 6, c).number_format = "0%"

rows = [
    ["CASH + PRODUCTIVITY GUARDRAILS", "Value", "Note"],
    ["Idle cash entering Q4", 1010838, "Asset Management 0.353 is idle cash, not wages"],
    ["Sales Option A vs current, $/person/year", 26564 - 24425, "about $535 per person per quarter"],
    ["Production Option A vs current, $/person/year", 22240 - 20757, "about $371 per person per quarter"],
    ["Q4 productivity to budget", 0.74, "industry ceiling ~73-75%. Never plan 85% again."],
    ["Q3 actual productivity", "70% sales / 72% production", "BB LLC sat 78% still sits inside that ceiling"],
]
r = block(r, rows, wrap_col=3)
ws.cell(r - 5, 2).number_format = "$#,##0"
ws.cell(r - 4, 2).number_format = "+$#,##0"
ws.cell(r - 3, 2).number_format = "+$#,##0"
ws.cell(r - 2, 2).number_format = "0%"
ws.cell(r - 2, 2).fill = yellow

ws.column_dimensions["A"].width = 52
ws.column_dimensions["B"].width = 18
ws.column_dimensions["C"].width = 38
ws.column_dimensions["D"].width = 22
ws.column_dimensions["E"].width = 16
for col in "FGHIJ":
    ws.column_dimensions[col].width = 14

# ---- patch Q4_Planner targets with exact 7-firm stats ----
if "Q4_Planner" in wb.sheetnames:
    p = wb["Q4_Planner"]
    for row in p.iter_rows(min_row=1, max_row=80, max_col=4):
        label = row[0].value
        if label == "Sales compensation [$]":
            row[1].value = 24425
            row[2].value = "7-firm mean 25,211 (range 22,060-27,630)"
            row[3].value = 26564
            row[3].fill = yellow
        elif label == "Production compensation [$]":
            row[1].value = 20757
            row[2].value = "7-firm mean 21,441 (range 19,609-23,293)"
            row[3].value = 22240
            row[3].fill = yellow

# ---- patch graph-read HR block with exact totals ----
if "Q3_Industry_Graphs" in wb.sheetnames:
    g = wb["Q3_Industry_Graphs"]
    for row in g.iter_rows(min_row=1, max_row=200, max_col=4):
        label = row[0].value
        if label == "Production worker compensation [$]":
            row[1].value = 20757
            row[2].value = "19,609 - 23,293 (7-firm)"
            row[3].value = "tied 5th/6th of 7; mix already right, salary low"
        elif label == "Sales force compensation [$]":
            row[1].value = 24425
            row[2].value = "22,060 - 27,630 (7-firm)"
            row[3].value = "6th of 7; salary starved, Full coverage bloated"

wb.save(DST)
print("Added Q3_Compensation; patched Q4_Planner + Q3_Industry_Graphs")
print(f"  sales 7-firm mean salary {sales_mean_sal:.0f} health {sales_mean_h:.0f} total {sales_mean_tot:.0f}")
print(f"  prod  7-firm mean salary {prod_mean_sal:.0f} health {prod_mean_h:.0f} total {prod_mean_tot:.0f}")
