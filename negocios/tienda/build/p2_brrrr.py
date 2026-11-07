from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.formatting.rule import FormulaRule
from common import *

OUT = "../productos/02-brrrr-calculator/BRRRR-Deal-Analyzer.xlsx"
wb = Workbook()
start_sheet(wb, "BRRRR Deal Analyzer",
    "Buy, Rehab, Rent, Refinance, Repeat: see exactly how much cash you get back at the refi, and the max price that recycles 100% of your money.",
    ["Open 'BRRRR Analyzer' and fill the yellow cells phase by phase (buy, rehab, rent, refinance). An example deal is pre-filled.",
     "Read the scorecard: cash left in the deal, % of capital recovered, cash flow after the refi, and your max offer price.",
     "Use 'Stress Test' to see what happens if the ARV comes in low or the rehab runs over budget."],
    [("BRRRR Analyzer", "The full deal across the 5 phases, with lender DSCR at refi, infinite-return detection and a pass/fail scorecard."),
     ("Stress Test", "Cash left in the deal across ARV × rehab overrun, and post-refi cash flow across refi rate × rent.")],
    DISCLAIMER_EN)
ws = wb.create_sheet("BRRRR Analyzer")
base(ws, {"A": 2, "B": 44, "C": 10, "D": 18, "E": 3, "F": 56})
title(ws, "B2", "BRRRR Deal Analyzer", "Yellow = your inputs. Results update instantly.")
R = {}
def row(r, lab, u, v, fmt, nt=""):
    label(ws, f"B{r}", lab); unit(ws, f"C{r}", u); inp(ws, f"D{r}", v, fmt)
    if nt: note(ws, f"F{r}", nt)
def calc(r, lab, fx, fmt, k=False, nt="", bold=False):
    label(ws, f"B{r}", lab, bold=bold or k)
    (key if k else out)(ws, f"D{r}", fx, fmt) if k else out(ws, f"D{r}", fx, fmt, bold=bold)
    if nt: note(ws, f"F{r}", nt)
band(ws, "B4:D4", "1 · BUY")
row(5, "Purchase price", "$", 150000, USD)
row(6, "Closing costs at purchase", "% price", 0.025, PCT)
row(7, "After-repair value (ARV)", "$", 300000, USD, "Use 3+ renovated sold comps within 0.5 mi / 6 months.")
row(8, "Max-offer rule", "% ARV", 0.70, PCT, "The classic 70% rule. Many BRRRR investors use 75%.")
band(ws, "B9:D9", "2 · REHAB")
row(10, "Rehab budget", "$", 45000, USD)
row(11, "Contingency", "% rehab", 0.10, PCT, "10–15% is prudent. Older houses: 20%.")
row(12, "Months until refinance", "months", 6, NUM, "Rehab + lease-up + lender seasoning (often 3–6 months).")
row(13, "Holding costs while vacant", "$/mo", 650, USD, "Taxes, insurance, utilities, lawn, security.")
band(ws, "B14:D14", "SHORT-TERM FINANCING (hard money / private)")
row(15, "Financed share of purchase", "%", 0.85, PCT, "Set to 0% if you buy cash.")
row(16, "Financed share of rehab", "%", 1.00, PCT)
row(17, "Interest rate", "%", 0.11, PCT2)
row(18, "Points", "% loan", 0.02, PCT2)
note(ws, "F19", "Simplification: interest is charged on the full loan for the whole period (conservative).")
band(ws, "B20:D20", "3 · RENT")
row(21, "Monthly rent", "$/mo", 2700, USD)
row(22, "Property taxes", "$/yr", 3600, USD)
row(23, "Insurance", "$/yr", 1500, USD)
row(24, "HOA", "$/mo", 0, USD)
row(25, "Vacancy", "% rent", 0.06, PCT)
row(26, "Repairs & maintenance", "% rent", 0.06, PCT)
row(27, "CapEx reserve", "% rent", 0.05, PCT)
row(28, "Property management", "% rent", 0.09, PCT)
band(ws, "B29:D29", "4 · REFINANCE")
row(30, "Cash-out refi LTV", "% ARV", 0.75, PCT, "DSCR / conventional investor loans: 70–75% cash-out.")
row(31, "Refi interest rate", "%", 0.0725, PCT2)
row(32, "Amortization", "years", 30, NUM)
row(33, "Refi closing costs", "% new loan", 0.025, PCT)

band(ws, "B35:D35", "RESULTS · CASH IN")
calc(36, "Rehab incl. contingency", "=D10*(1+D11)", USD)
calc(37, "Short-term loan", "=D5*D15+D36*D16", USD)
calc(38, "Points paid", "=D37*D18", USD)
calc(39, "Interest during hold", "=D37*D17/12*D12", USD)
calc(40, "Holding costs", "=D13*D12", USD)
calc(41, "Total project cost (all-in)", "=D5*(1+D6)+D36+D38+D39+D40", USD, bold=True)
calc(42, "YOUR CASH INVESTED", "=D41-D37", USD, k=True, nt="Down payment + closing + unfinanced rehab + points + interest + holding.")
band(ws, "B43:D43", "RESULTS · REFINANCE")
calc(44, "New refi loan", "=D7*D30", USD)
calc(45, "Pay off short-term loan", "=-D37", USD)
calc(46, "Refi closing costs", "=-D44*D33", USD)
calc(47, "CASH BACK TO YOU AT REFI", "=D44+D45+D46", USD, k=True, nt="Negative = you must bring cash to the refi.")
calc(48, "CASH LEFT IN THE DEAL", "=D42-D47", USD, k=True, nt="Zero or negative = all your money came back (infinite return).")
calc(49, "% of your capital recovered", "=IFERROR(D47/D42,0)", PCT)
calc(50, "Equity created (ARV − new loan)", "=D7-D44", USD)
band(ws, "B51:D51", "RESULTS · AFTER THE REFI (MONTHLY)")
calc(52, "Principal & interest", "=-PMT(D31/12,D32*12,D44)", USD)
calc(53, "PITIA", "=D52+D22/12+D23/12+D24", USD)
calc(54, "Lender DSCR (rent ÷ PITIA)", "=IFERROR(D21/D53,0)", MULT, nt="Most DSCR lenders need 1.00–1.25x to approve the refi.")
calc(55, "Net operating income", "=D21*(1-D25)*(1-D26-D27-D28)-D22/12-D23/12-D24", USD)
calc(56, "MONTHLY CASH FLOW", "=D55-D52", USD, k=True)
calc(57, "Cash-on-cash on cash left in deal", '=IF(D48<=0,"INFINITE",IFERROR(D56*12/D48,0))', PCT, k=True)
band(ws, "B58:D58", "5 · REPEAT: YOUR OFFER LIMITS")
calc(59, "Max offer by rule (ARV × rule − rehab)", "=D7*D8-D36", USD, k=True)
calc(60, "Max price to recover 100% of your cash", "=MAX(0,(D7*D30*(1-D33)-D36*(1+D16*(D18+D17*D12/12))-D13*D12)/(1+D6+D15*(D18+D17*D12/12)))", USD, k=True,
     nt="Solved algebraically: at this price, cash left in the deal = $0.")
band(ws, "B61:D61", "SCORECARD")
checks = [(62, "Cash left ≤ 10% of all-in cost", '=IF(D48<=0.1*D41,"PASS","FAIL")'),
          (63, "Refi DSCR ≥ 1.20x", '=IF(D54>=1.2,"PASS","FAIL")'),
          (64, "Positive cash flow after refi", '=IF(D56>0,"PASS","FAIL")'),
          (65, "Price within your max-offer rule", '=IF(D5<=D59,"PASS","FAIL")')]
for r, lab, fx in checks:
    calc(r, lab, fx, None, bold=True)
calc(66, "DEAL VERDICT", '=IF(COUNTIF(D62:D65,"PASS")=4,"GREAT BRRRR",IF(COUNTIF(D62:D65,"PASS")>=3,"WORKABLE: check the FAIL","PASS ON IT"))', None, k=True)
pass_fail(ws, "D62:D65", "PASS")
ws.conditional_formatting.add("D66", FormulaRule(formula=['D66="GREAT BRRRR"'], font=Font(name=FONT, bold=True, color="FFFFFF"), fill=PatternFill("solid", fgColor=GREEN)))
ws.conditional_formatting.add("D66", FormulaRule(formula=['LEFT(D66,8)="WORKABLE"'], font=Font(name=FONT, bold=True, color="FFFFFF"), fill=PatternFill("solid", fgColor="A1580B")))
ws.conditional_formatting.add("D66", FormulaRule(formula=['D66="PASS ON IT"'], font=Font(name=FONT, bold=True, color="FFFFFF"), fill=PatternFill("solid", fgColor=RED)))
neg_red(ws, "D47:D48"); neg_red(ws, "D56")
ws.freeze_panes = "A4"

s = wb.create_sheet("Stress Test")
base(s, {"A": 2, "B": 26, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 14, "I": 3, "J": 40})
title(s, "B2", "Stress test", "What if the appraisal comes in low or the rehab runs over?")
A = "'BRRRR Analyzer'!"
band(s, "B4:H4", "CASH LEFT IN THE DEAL  ·  ARV (columns) × rehab overrun (rows)")
s["B5"] = "Overrun ↓ / ARV →"; s["B5"].font = f(True, 9, GREY)
arv = [-0.15, -0.10, -0.05, 0, 0.05, 0.10]
for j, c in enumerate("CDEFGH"):
    out(s, f"{c}5", f"={A}$D$7*(1+({arv[j]}))", USD, bold=True)
for i, ov in enumerate([0, 0.1, 0.2, 0.3, 0.5]):
    r = 6 + i
    out(s, f"B{r}", ov, PCT, bold=True)
    for c in "CDEFGH":
        Rx = f"({A}$D$36*(1+$B{r}))"
        j = f"({A}$D$18+{A}$D$17*{A}$D$12/12)"
        fx = (f"={A}$D$5*(1+{A}$D$6+{A}$D$15*{j})+{Rx}*(1+{A}$D$16*{j})+{A}$D$13*{A}$D$12-{c}$5*{A}$D$30*(1-{A}$D$33)")
        out(s, f"{c}{r}", fx, USD)
s.conditional_formatting.add("C6:H10", FormulaRule(formula=["C6<=0"], fill=PatternFill("solid", fgColor="D7F0E2")))
s.conditional_formatting.add("C6:H10", FormulaRule(formula=[f"AND(C6>0,C6<=0.1*{A}$D$41)"], fill=PatternFill("solid", fgColor="FBEFDF")))
s.conditional_formatting.add("C6:H10", FormulaRule(formula=[f"C6>0.1*{A}$D$41"], fill=PatternFill("solid", fgColor="F9DCD9")))
note(s, "J6", "Green = all cash back. Amber = up to 10% of all-in cost left in. Red = more than 10% stuck in the deal.")
band(s, "B13:H13", "MONTHLY CASH FLOW AFTER REFI  ·  rent (columns) × refi rate (rows)")
s["B14"] = "Rate ↓ / Rent →"; s["B14"].font = f(True, 9, GREY)
rents = [-0.15, -0.10, -0.05, 0, 0.05, 0.10]
for j, c in enumerate("CDEFGH"):
    out(s, f"{c}14", f"={A}$D$21*(1+({rents[j]}))", USD, bold=True)
for i, d in enumerate([-0.01, -0.005, 0, 0.005, 0.01, 0.015]):
    r = 15 + i
    out(s, f"B{r}", f"={A}$D$31+({d})", PCT2, bold=True)
    for c in "CDEFGH":
        fx = (f"={c}$14*(1-{A}$D$25)*(1-{A}$D$26-{A}$D$27-{A}$D$28)-{A}$D$22/12-{A}$D$23/12-{A}$D$24+PMT($B{r}/12,{A}$D$32*12,{A}$D$44)")
        out(s, f"{c}{r}", fx, USD)
s.conditional_formatting.add("C15:H20", FormulaRule(formula=["C15<0"], fill=PatternFill("solid", fgColor="F9DCD9")))
s.conditional_formatting.add("C15:H20", FormulaRule(formula=["C15>=0"], fill=PatternFill("solid", fgColor="D7F0E2")))
wb.active = 1
import os; os.makedirs(os.path.dirname(OUT), exist_ok=True); wb.save(OUT); print("saved")
