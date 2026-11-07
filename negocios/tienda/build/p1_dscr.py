from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.formatting.rule import FormulaRule
from common import *
import re

OUT = "../productos/01-dscr-rental-analyzer/DSCR-Rental-Property-Analyzer.xlsx"
wb = Workbook()
start_sheet(wb, "DSCR Loan Qualifier & Rental Property Analyzer",
    "Know in 60 seconds if a rental qualifies for a DSCR loan, how much you can borrow, and whether it cash flows.",
    ["Go to the 'Analyzer' tab and type your numbers in the yellow cells. Up to 3 properties side by side (an example is pre-filled; overwrite it).",
     "Read the verdict: lender DSCR, QUALIFIES / BELOW TARGET, max loan and max price you can pay at your target DSCR.",
     "Open 'Sensitivity' to stress-test rate and rent, and 'Lender Checklist' to prepare your loan application."],
    [("Analyzer", "Lender view (PITIA, DSCR, max loan, rent needed) + investor view (NOI, cash flow, cap rate, cash-on-cash) for 3 deals side by side."),
     ("Sensitivity", "Three stress-test grids (DSCR, monthly cash flow, max price) across interest rates and rents. Pick which property to test."),
     ("Lender Checklist", "Typical DSCR loan requirements vs. your deal, with automatic checks for DSCR, LTV, credit score and reserves.")],
    DISCLAIMER_EN)

ws = wb.create_sheet("Analyzer")
base(ws, {"A": 2, "B": 38, "C": 9, "D": 17, "E": 17, "F": 17, "G": 3, "H": 52})
title(ws, "B2", "DSCR Loan Qualifier & Rental Analyzer", "Yellow = your inputs. Compare up to 3 properties.")
cols = ["D", "E", "F"]
for c, n in zip(cols, [1, 2, 3]):
    ws[f"{c}4"] = f"PROPERTY {n}"
    ws[f"{c}4"].font = Font(name=FONT, bold=True, color="FFFFFF")
    ws[f"{c}4"].fill = PatternFill("solid", fgColor=ACCENT)
    ws[f"{c}4"].alignment = Alignment(horizontal="center")

ex = {
 5: ["Maple St Duplex", "Oak Ave House", "Pine Ct Condo"],
 6: [350000, 280000, 220000], 7: [0.03, 0.03, 0.03], 8: [0, 15000, 0], 9: [3100, 2500, 1850], 10: [0, 0, 0],
 12: [0.75, 0.80, 0.75], 13: [0.0725, 0.075, 0.07875], 14: [30, 30, 30], 15: ["No", "No", "Yes"], 16: [0.015, 0.02, 0.01],
 18: [4200, 3600, 2400], 19: [1800, 1500, 900], 20: [0, 0, 250], 21: [0, 0, 0], 22: [0, 0, 50],
 23: [0.05, 0.05, 0.05], 24: [0.05, 0.05, 0.04], 25: [0.05, 0.05, 0.03], 26: [0.08, 0.08, 0.08], 28: [1.25, 1.25, 1.25],
}
rows = [
 ("band", 3, "PROPERTY"),
 (5, "Property nickname", "", None, ""),
 (6, "Purchase price (or as-is value for a refi)", "$", USD, ""),
 (7, "Closing costs", "% price", PCT, "Title, escrow, appraisal, lender fees. 2–4% is common."),
 (8, "Upfront repairs / rehab", "$", USD, "Cash you spend before renting."),
 (9, "Monthly market rent", "$/mo", USD, "Use the appraiser's rent schedule (Form 1007) or the lease, whichever your lender uses."),
 (10, "Other monthly income", "$/mo", USD, "Parking, laundry, storage. Lenders usually ignore it for DSCR."),
 ("band", 11, "LOAN"),
 (12, "Loan-to-value (LTV)", "%", PCT, "DSCR purchase loans are typically 70–80% LTV."),
 (13, "Interest rate", "%", PCT2, ""),
 (14, "Amortization", "years", NUM, "30 years is standard."),
 (15, "Interest-only?", "Yes/No", None, "Interest-only payments raise DSCR. Many lenders offer 5–10 yr IO."),
 (16, "Points / origination", "% loan", PCT2, ""),
 ("band", 17, "EXPENSES"),
 (18, "Property taxes", "$/yr", USD, ""),
 (19, "Insurance (hazard)", "$/yr", USD, ""),
 (20, "HOA dues", "$/mo", USD, "Included in PITIA by lenders."),
 (21, "Flood insurance", "$/yr", USD, ""),
 (22, "Utilities paid by owner", "$/mo", USD, "Investor view only."),
 (23, "Vacancy", "% rent", PCT, "Investor view only. Lenders use gross rent."),
 (24, "Repairs & maintenance", "% rent", PCT, "Investor view only."),
 (25, "CapEx reserve", "% rent", PCT, "Roof, HVAC, water heater. Investor view only."),
 (26, "Property management", "% rent", PCT, "Investor view only. 8–10% is typical."),
 ("band", 27, "LENDER TEST"),
 (28, "Target DSCR", "x", MULT, "Most lenders: 1.00–1.25x minimum. Better DSCR = better rate/LTV."),
]
for r in rows:
    if r[0] == "band":
        band(ws, f"B{r[1]+1 if r[1]==3 else r[1]}:F{r[1]+1 if r[1]==3 else r[1]}", r[2]) if False else None
for r in rows:
    if r[0] == "band":
        continue
    rr, lab, u, fmt, nt = r
    label(ws, f"B{rr}", lab); unit(ws, f"C{rr}", u)
    for c, v in zip(cols, ex[rr]):
        inp(ws, f"{c}{rr}", v, fmt)
    if nt:
        note(ws, f"H{rr}", nt)
for rr, t in [(11, "LOAN"), (17, "EXPENSES"), (27, "LENDER TEST")]:
    band(ws, f"B{rr}:F{rr}", t)
dropdown(ws, "D15:F15", ["Yes", "No"])

band(ws, "B30:F30", "LENDER VIEW (HOW THE BANK SEES IT)")
L = [
 (31, "Loan amount", "=D6*D12", USD, False, ""),
 (32, "Principal & interest", '=IF(D15="Yes",D31*D13/12,-PMT(D13/12,D14*12,D31))', USD, False, "Monthly."),
 (33, "Taxes + insurance + HOA + flood", "=D18/12+D19/12+D20+D21/12", USD, False, "Monthly."),
 (34, "PITIA (total monthly housing payment)", "=D32+D33", USD, True, "Principal, Interest, Taxes, Insurance, Association dues."),
 (35, "LENDER DSCR", "=IFERROR(D9/D34,0)", MULT, "key", "Gross rent ÷ PITIA: the most common DSCR-lender method."),
 (36, "Result vs target", '=IF(D35>=D28,"QUALIFIES","BELOW TARGET")', None, True, ""),
 (37, "Max loan at target DSCR", '=MAX(0,IF(D15="Yes",IFERROR((D9/D28-D33)/(D13/12),0),PV(D13/12,D14*12,-(D9/D28-D33))))', USD, "key", "The biggest loan this rent supports."),
 (38, "Max purchase price at your LTV", "=IFERROR(D37/D12,0)", USD, "key", "Your walk-away price for this rent and rate."),
 (39, "Break-even rent (DSCR 1.00x)", "=D34", USD, False, ""),
 (40, "Rent needed for target DSCR", "=D34*D28", USD, False, ""),
 (41, "Typical reserves required (6 months PITIA)", "=D34*6", USD, False, "Lenders ask for 3–12 months. See Lender Checklist."),
]
band(ws, "B42:F42", "CASH TO CLOSE")
C = [
 (43, "Down payment", "=D6-D31", USD, False, ""),
 (44, "Closing costs", "=D6*D7", USD, False, ""),
 (45, "Points", "=D31*D16", USD, False, ""),
 (46, "Upfront repairs", "=D8", USD, False, ""),
 (47, "TOTAL CASH NEEDED", "=SUM(D43:D46)", USD, "key", ""),
]
band(ws, "B48:F48", "INVESTOR VIEW (WHAT YOU ACTUALLY KEEP)")
I = [
 (49, "Effective gross income", "=(D9+D10)*(1-D23)", USD, False, "Monthly, after vacancy."),
 (50, "Fixed expenses (taxes, ins., HOA, flood, utilities)", "=D33+D22", USD, False, "Monthly."),
 (51, "Variable expenses (repairs, CapEx, management)", "=D49*(D24+D25+D26)", USD, False, "Monthly."),
 (52, "Net operating income (NOI)", "=D49-D50-D51", USD, True, "Monthly."),
 (53, "MONTHLY CASH FLOW", "=D52-D32", USD, "key", "After the mortgage payment."),
 (54, "Annual cash flow", "=D53*12", USD, False, ""),
 (55, "Cap rate", "=IFERROR(D52*12/D6,0)", PCT2, False, "Annual NOI ÷ price."),
 (56, "Cash-on-cash return", "=IFERROR(D54/D47,0)", PCT2, "key", "Annual cash flow ÷ total cash needed."),
 (57, "Investor DSCR (NOI ÷ debt service)", "=IFERROR(D52/D32,0)", MULT, False, "Commercial-style DSCR. Stricter than the lender method."),
 (58, "1% rule (rent ÷ price)", "=IFERROR(D9/D6,0)", PCT2, False, "Quick screen: 1%+ usually cash flows."),
]
band(ws, "B59:F59", "VERDICT")
V = [(60, "Deal verdict", '=IF(AND(D35>=D28,D53>0),"STRONG: qualifies + cash flows",IF(D35>=D28,"MIXED: qualifies, negative cash flow",IF(D53>0,"MIXED: cash flows, below DSCR target","WEAK: fails both tests")))', None, True, "")]
for rr, lab, fx, fmt, bold, nt in L + C + I + V:
    label(ws, f"B{rr}", lab, bold=bool(bold))
    for c in cols:
        fx_c = re.sub(r"(?<![A-Z$])D(?=\d)", c, fx)
        # replace column refs only (D -> c) safely: formulas only use D as column letter
        if bold == "key":
            key(ws, f"{c}{rr}", fx_c, fmt)
        else:
            out(ws, f"{c}{rr}", fx_c, fmt, bold=bool(bold))
    if nt:
        note(ws, f"H{rr}", nt)
for c in cols:
    ws[f"{c}60"].alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
ws.row_dimensions[60].height = 42
pass_fail(ws, "D36:F36", "QUALIFIES")
ws.conditional_formatting.add("D60:F60", FormulaRule(formula=['LEFT(D60,6)="STRONG"'], font=Font(name=FONT, bold=True, color="FFFFFF"), fill=PatternFill("solid", fgColor=GREEN)))
ws.conditional_formatting.add("D60:F60", FormulaRule(formula=['LEFT(D60,5)="MIXED"'], font=Font(name=FONT, bold=True, color="FFFFFF"), fill=PatternFill("solid", fgColor="A1580B")))
ws.conditional_formatting.add("D60:F60", FormulaRule(formula=['LEFT(D60,4)="WEAK"'], font=Font(name=FONT, bold=True, color="FFFFFF"), fill=PatternFill("solid", fgColor=RED)))
neg_red(ws, "D53:F54")
ws.freeze_panes = "D5"

# ---------- Sensitivity ----------
s = wb.create_sheet("Sensitivity")
base(s, {"A": 2, "B": 30, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 3, "I": 40})
title(s, "B2", "Stress test: interest rate × rent", "Change the property number to test another deal.")
label(s, "B4", "Test property #", bold=True); inp(s, "C4", 1); dropdown(s, "C4", ["1", "2", "3"])
label(s, "B5", "Property"); out(s, "C5", "=INDEX(Analyzer!$D$5:$F$5,$C$4)")
helpers = [(6, "Base rent ($/mo)", "=INDEX(Analyzer!$D$9:$F$9,$C$4)", USD), (7, "Base rate", "=INDEX(Analyzer!$D$13:$F$13,$C$4)", PCT2),
           (8, "Loan amount", "=INDEX(Analyzer!$D$31:$F$31,$C$4)", USD), (9, "Taxes+ins.+HOA+flood ($/mo)", "=INDEX(Analyzer!$D$33:$F$33,$C$4)", USD),
           (10, "Amortization (yrs)", "=INDEX(Analyzer!$D$14:$F$14,$C$4)", NUM), (11, "Interest-only?", "=INDEX(Analyzer!$D$15:$F$15,$C$4)", None),
           (12, "Target DSCR", "=INDEX(Analyzer!$D$28:$F$28,$C$4)", MULT), (13, "LTV", "=INDEX(Analyzer!$D$12:$F$12,$C$4)", PCT),
           (14, "Other income + vacancy/opex factor", "=(1-INDEX(Analyzer!$D$23:$F$23,$C$4))*(1-INDEX(Analyzer!$D$24:$F$24,$C$4)-INDEX(Analyzer!$D$25:$F$25,$C$4)-INDEX(Analyzer!$D$26:$F$26,$C$4))", '0.000'),
           (15, "Owner utilities ($/mo)", "=INDEX(Analyzer!$D$22:$F$22,$C$4)", USD), (16, "Other income ($/mo)", "=INDEX(Analyzer!$D$10:$F$10,$C$4)", USD)]
for rr, lab, fx, fmt in helpers:
    label(s, f"B{rr}", lab); out(s, f"C{rr}", fx, fmt)
for c in "CDEFG":
    pass

def grid(top, ttl, cell_fx, fmt, kind):
    band(s, f"B{top}:G{top}", ttl)
    s[f"B{top+1}"] = "Rate ↓  /  Rent →"; s[f"B{top+1}"].font = f(True, 9, GREY)
    rents = [0.8, 0.9, 1.0, 1.1, 1.2]
    for j, c in enumerate("CDEFG"):
        out(s, f"{c}{top+1}", f"=$C$6*{rents[j]}", USD, bold=True)
    for i, d in enumerate([-0.015, -0.01, -0.005, 0, 0.005, 0.01, 0.015]):
        rr = top + 2 + i
        out(s, f"B{rr}", f"=MAX(0.0001,$C$7+({d}))", PCT2, bold=True)
        for c in "CDEFG":
            pi = f'IF($C$11="Yes",$C$8*$B{rr}/12,-PMT($B{rr}/12,$C$10*12,$C$8))'
            out(s, f"{c}{rr}", cell_fx(c, top + 1, rr, pi), fmt)
    rng = f"C{top+2}:G{top+8}"
    if kind == "dscr":
        s.conditional_formatting.add(rng, FormulaRule(formula=[f"C{top+2}>=$C$12"], fill=PatternFill("solid", fgColor="D7F0E2")))
        s.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(C{top+2}>=1,C{top+2}<$C$12)"], fill=PatternFill("solid", fgColor="FBEFDF")))
        s.conditional_formatting.add(rng, FormulaRule(formula=[f"C{top+2}<1"], fill=PatternFill("solid", fgColor="F9DCD9")))
    else:
        s.conditional_formatting.add(rng, FormulaRule(formula=[f"C{top+2}<0"], fill=PatternFill("solid", fgColor="F9DCD9")))
        s.conditional_formatting.add(rng, FormulaRule(formula=[f"C{top+2}>=0"], fill=PatternFill("solid", fgColor="D7F0E2")))

grid(18, "LENDER DSCR  (green = meets target, amber = 1.00x to target, red = below 1.00x)",
     lambda c, hr, rr, pi: f"=IFERROR({c}${hr}/({pi}+$C$9),0)", MULT, "dscr")
grid(28, "MONTHLY CASH FLOW AFTER DEBT",
     lambda c, hr, rr, pi: f"=({c}${hr}+$C$16)*$C$14-$C$9-$C$15-{pi}", USD, "cf")
grid(38, "MAX PURCHASE PRICE THAT HITS YOUR TARGET DSCR",
     lambda c, hr, rr, pi: f'=MAX(0,IF($C$11="Yes",({c}${hr}/$C$12-$C$9)/($B{rr}/12),PV($B{rr}/12,$C$10*12,-({c}${hr}/$C$12-$C$9))))/$C$13', USD, "cf")
note(s, "I19", "Each cell recalculates the payment at that rate and divides the rent at that level by PITIA.")
note(s, "I39", "Use this as your offer ceiling: above it, the loan shrinks or the deal won't qualify.")

# ---------- Lender checklist ----------
k = wb.create_sheet("Lender Checklist")
base(k, {"A": 2, "B": 34, "C": 30, "D": 16, "E": 16, "F": 3, "G": 50})
title(k, "B2", "DSCR loan readiness checklist", "Typical 2026 market ranges. Every lender differs: confirm with yours.")
label(k, "B4", "Check property #", bold=True); inp(k, "D4", 1); dropdown(k, "D4", ["1", "2", "3"])
label(k, "B5", "Your credit score", bold=True); inp(k, "D5", 720, NUM)
label(k, "B6", "Liquid reserves you have after closing ($)", bold=True); inp(k, "D6", 20000, USD)
for c, t in zip("BCDE", ["Requirement", "Typical range", "Your deal", "Check"]):
    k[f"{c}8"] = t; k[f"{c}8"].font = Font(name=FONT, bold=True, color="FFFFFF"); k[f"{c}8"].fill = PatternFill("solid", fgColor=NAVY)
chk = [
 ("Minimum DSCR", "1.00x–1.25x", "=INDEX(Analyzer!$D$35:$F$35,$D$4)", MULT, '=IF(D9>=1.25,"PASS",IF(D9>=1,"PASS (some lenders)","FAIL"))'),
 ("Max LTV (purchase)", "70%–80%", "=INDEX(Analyzer!$D$12:$F$12,$D$4)", PCT, '=IF(D10<=0.75,"PASS",IF(D10<=0.8,"PASS (some lenders)","FAIL"))'),
 ("Minimum credit score", "660–700+", "=D5", NUM, '=IF(D11>=700,"PASS",IF(D11>=660,"PASS (some lenders)","FAIL"))'),
 ("Reserves (months of PITIA)", "3–12 months", "=IFERROR(D6/INDEX(Analyzer!$D$34:$F$34,$D$4),0)", '0.0" mo"', '=IF(D12>=6,"PASS",IF(D12>=3,"PASS (some lenders)","FAIL"))'),
 ("Minimum loan amount", "$75k–$100k", "=INDEX(Analyzer!$D$31:$F$31,$D$4)", USD, '=IF(D13>=100000,"PASS",IF(D13>=75000,"PASS (some lenders)","FAIL"))'),
]
for i, (req, rng_, fx, fmt, ck) in enumerate(chk):
    rr = 9 + i
    label(k, f"B{rr}", req); k[f"C{rr}"] = rng_; k[f"C{rr}"].font = f()
    out(k, f"D{rr}", fx, fmt); out(k, f"E{rr}", ck, None, bold=True)
pass_fail(k, "E9:E13", "PASS")
band(k, "B15:E15", "DOCUMENTS & TERMS TO EXPECT")
docs = ["Appraisal with rent schedule (Form 1007 for 1–4 units) or lease", "Entity docs if buying in an LLC (articles, operating agreement, EIN)",
        "2–3 months of bank statements for reserves and down payment", "Hazard insurance quote (and flood if in a flood zone)",
        "Prepayment penalty is common: e.g., 5-4-3-2-1 or 3-2-1 step-down", "Cash-out refi: seasoning of 3–6 months is common",
        "No personal income/DTI verification: the property's rent qualifies the loan"]
for i, d in enumerate(docs):
    k[f"B{16+i}"] = "☐  " + d; k[f"B{16+i}"].font = f()
note(k, "G9", "PASS (some lenders) = acceptable to part of the market, usually at a higher rate or lower LTV.")

wb.active = 1
import os; os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb.save(OUT); print("saved", OUT)
