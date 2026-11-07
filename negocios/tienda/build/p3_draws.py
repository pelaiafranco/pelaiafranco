from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.formatting.rule import FormulaRule
from common import *

OUT = "../productos/03-construction-draw-tracker/Construction-Draw-Tracker-Flip-Budget.xlsx"
wb = Workbook()
start_sheet(wb, "Construction Draw Tracker & Fix-and-Flip Budget",
    "Run your rehab like a lender does: line-item budget, 8 draws, a draw request ready to send, interest on drawn balance and your real flip profit.",
    ["In 'Setup', enter the property, loan terms and timeline. In 'Budget & Draws', enter your line-item budget (an example project is pre-filled).",
     "Each time you request funds, type the amounts in the next draw column and the month it is funded. Over-budget lines turn red.",
     "Pick the draw number in 'Draw Request' and send it to your lender. Check 'Interest' and 'Flip P&L' for your real cost and profit."],
    [("Setup", "Project, purchase, construction loan (Dutch or non-Dutch interest), fees, timeline and exit."),
     ("Budget & Draws", "20 rehab line items + contingency × 8 draws, with drawn to date, remaining, % complete and over-budget alerts."),
     ("Draw Request", "Lender-style draw request for any draw: previous, this request, to date, remaining, retainage, fees and net funding."),
     ("Interest", "Month-by-month loan balance and interest for 18 months. Compares Dutch (full commitment) vs non-Dutch (drawn balance)."),
     ("Flip P&L", "All-in cost, profit, ROI, annualized ROI and the lender ratios (LTC, LTV on ARV) with pass/fail checks.")],
    DISCLAIMER_EN)

# ---------------- Setup ----------------
st = wb.create_sheet("Setup")
base(st, {"A": 2, "B": 42, "C": 10, "D": 20, "E": 3, "F": 58})
title(st, "B2", "Project setup", "Yellow = your inputs.")
def srow(r, lab, u, v, fmt, nt=""):
    label(st, f"B{r}", lab); unit(st, f"C{r}", u); inp(st, f"D{r}", v, fmt)
    if nt: note(st, f"F{r}", nt)
band(st, "B4:D4", "PROPERTY")
srow(5, "Project name", "", "123 Elm St · Full gut rehab", None)
srow(6, "Purchase price", "$", 320000, USD)
srow(7, "Closing costs at purchase", "% price", 0.02, PCT)
srow(8, "After-repair value (ARV)", "$", 720000, USD)
band(st, "B9:D9", "CONSTRUCTION LOAN")
srow(10, "Share of purchase financed", "%", 0.80, PCT, "Fix-and-flip lenders: typically 80–90% of purchase.")
srow(11, "Share of rehab financed (holdback)", "%", 1.00, PCT, "Usually 100% of rehab, released in draws after inspection.")
srow(12, "Interest rate", "% / yr", 0.115, PCT2)
srow(13, "Points / origination", "% loan", 0.02, PCT2)
srow(14, "Interest method", "", "Non-Dutch", None, "Non-Dutch = interest only on money drawn. Dutch = interest on the full loan from day one.")
srow(15, "Retainage held per draw", "% draw", 0.10, PCT, "Held back from your contractor until completion. Set 0% if not used.")
srow(16, "Draw fee", "$ / draw", 250, USD)
srow(17, "Inspection fee", "$ / draw", 150, USD)
band(st, "B18:D18", "TIMELINE & EXIT")
srow(19, "Months from purchase to sale closing", "months", 8, NUM, "Rehab + listing + escrow.")
srow(20, "Holding costs (taxes, insurance, utilities)", "$/mo", 700, USD)
srow(21, "Selling costs", "% ARV", 0.065, PCT, "Agent commissions + seller closing costs.")
band(st, "B22:D22", "LENDER LIMITS (for checks)")
srow(23, "Max loan-to-cost (LTC)", "%", 0.90, PCT, "Typical 85–90% of purchase + rehab.")
srow(24, "Max loan-to-ARV", "%", 0.75, PCT, "Typical 65–75% of ARV.")
dropdown(st, "D14", ["Non-Dutch", "Dutch"])

# ---------------- Budget & Draws ----------------
b = wb.create_sheet("Budget & Draws")
base(b, {"A": 2, "B": 30, "C": 13, **{c: 11.5 for c in "DEFGHIJK"}, "L": 13, "M": 13, "N": 10, "O": 14})
title(b, "B2", "Budget & draws", "Enter the amount approved in each draw and the month it was funded.")
hdr = ["Line item", "Budget"] + [f"Draw {i}" for i in range(1, 9)] + ["Drawn to date", "Remaining", "% complete", "Status"]
for i, h in enumerate(hdr):
    c = b.cell(row=5, column=2 + i, value=h)
    c.font = Font(name=FONT, bold=True, color="FFFFFF", size=9); c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
b.row_dimensions[5].height = 30
label(b, "B4", "Month funded →", bold=True)
for c, m in zip("DEFGHIJK", [1, 2, 3, 4, 5, 6, None, None]):
    inp(b, f"{c}4", m, NUM)
items = [("Permits & plans", 6500, [6500]), ("Demolition & dumpsters", 9000, [9000]), ("Foundation & structural", 12000, [0, 12000]),
         ("Framing", 15000, [0, 9000, 6000]), ("Roofing", 14000, [0, 14000]), ("Windows & exterior doors", 11000, [0, 0, 11000]),
         ("Plumbing", 16000, [0, 8000, 0, 8000]), ("Electrical", 15000, [0, 7500, 0, 7500]), ("HVAC", 13000, [0, 0, 6500, 6500]),
         ("Insulation", 4500, [0, 0, 0, 4500]), ("Drywall", 12000, [0, 0, 0, 6000, 6000]), ("Cabinets", 14000, [0, 0, 0, 0, 14000]),
         ("Countertops", 7000, [0, 0, 0, 0, 0, 7000]), ("Flooring", 13000, [0, 0, 0, 0, 6500, 6500]), ("Interior paint & trim", 9500, [0, 0, 0, 0, 0, 9500]),
         ("Appliances", 6500, [0, 0, 0, 0, 0, 6500]), ("Plumbing & lighting fixtures", 5500, [0, 0, 0, 0, 0, 5500]),
         ("Siding & exterior paint", 10000, [0, 0, 0, 0, 10000]), ("Landscaping & driveway", 6000, [0, 0, 0, 0, 0, 6500]),
         ("Final clean & staging", 3500, [0, 0, 0, 0, 0, 0]), ("Contingency", 15000, [0, 0, 2500, 0, 0, 0])]
first = 6
for i, (name, bud, draws) in enumerate(items):
    r = first + i
    inp(b, f"B{r}", name); b[f"B{r}"].alignment = Alignment(horizontal="left")
    inp(b, f"C{r}", bud, USD)
    for j, c in enumerate("DEFGHIJK"):
        v = draws[j] if j < len(draws) else None
        inp(b, f"{c}{r}", v if v else None, USD)
    out(b, f"L{r}", f"=SUM(D{r}:K{r})", USD)
    out(b, f"M{r}", f"=C{r}-L{r}", USD)
    out(b, f"N{r}", f"=IFERROR(L{r}/C{r},0)", PCT)
    out(b, f"O{r}", f'=IF(L{r}>C{r},"OVER BUDGET",IF(L{r}=C{r},"Complete",IF(L{r}=0,"Not started","In progress")))', None, bold=True)
last = first + len(items) - 1
tot = last + 1
label(b, f"B{tot}", "TOTAL", bold=True)
for c in "CDEFGHIJKLM":
    key(b, f"{c}{tot}", f"=SUM({c}{first}:{c}{last})", USD)
key(b, f"N{tot}", f"=IFERROR(L{tot}/C{tot},0)", PCT)
b.conditional_formatting.add(f"O{first}:O{last}", FormulaRule(formula=[f'O{first}="OVER BUDGET"'], font=Font(name=FONT, bold=True, color="FFFFFF"), fill=PatternFill("solid", fgColor=RED)))
b.conditional_formatting.add(f"O{first}:O{last}", FormulaRule(formula=[f'O{first}="Complete"'], font=Font(name=FONT, bold=True, color=GREEN)))
b.conditional_formatting.add(f"M{first}:M{last}", FormulaRule(formula=[f"M{first}<0"], font=Font(name=FONT, bold=True, color=RED)))
rows2 = [(tot + 1, "Lender-funded share", f"=$D${{t}}*Setup!$D$11", USD), (tot + 2, "Retainage held", f"=-$D${{t}}*Setup!$D$15", USD),
         (tot + 3, "Draw + inspection fees", f'=IF($D${{t}}>0,-(Setup!$D$16+Setup!$D$17),0)', USD), (tot + 4, "NET FUNDED TO YOU", None, USD)]
label(b, f"B{tot+1}", "Lender-funded share"); label(b, f"B{tot+2}", "Less retainage held"); label(b, f"B{tot+3}", "Less draw + inspection fees"); label(b, f"B{tot+4}", "NET FUNDED", bold=True)
for c in "DEFGHIJK":
    out(b, f"{c}{tot+1}", f"={c}{tot}*Setup!$D$11", USD)
    out(b, f"{c}{tot+2}", f"=-{c}{tot+1}*Setup!$D$15", USD)
    out(b, f"{c}{tot+3}", f"=IF({c}{tot}>0,-(Setup!$D$16+Setup!$D$17),0)", USD)
    out(b, f"{c}{tot+4}", f"=SUM({c}{tot+1}:{c}{tot+3})", USD, bold=True)
note(b, f"B{tot+6}", "Example project pre-filled: draw 6 shows landscaping over budget, so you can see the alert. Overwrite all yellow cells with your project.")
b.freeze_panes = "C6"
TOT = tot

# ---------------- Draw Request ----------------
d = wb.create_sheet("Draw Request")
base(d, {"A": 2, "B": 30, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 11})
title(d, "B2", "Draw request", "Choose the draw number. Send this sheet as PDF to your lender (File > Print > Save as PDF).")
label(d, "B4", "Draw #", bold=True); inp(d, "C4", 3, NUM); dropdown(d, "C4", [str(i) for i in range(1, 9)])
label(d, "B5", "Project"); out(d, "C5", "=Setup!D5"); d["C5"].alignment = Alignment(horizontal="left")
label(d, "B6", "Funding month"); out(d, "C6", "=INDEX('Budget & Draws'!$D$4:$K$4,$C$4)", NUM)
for i, h in enumerate(["Line item", "Budget", "Previous draws", "This request", "Total to date", "Remaining", "% complete"]):
    c = d.cell(row=8, column=2 + i, value=h); c.font = Font(name=FONT, bold=True, color="FFFFFF", size=9)
    c.fill = PatternFill("solid", fgColor=NAVY); c.alignment = Alignment(horizontal="center", wrap_text=True)
BD = "'Budget & Draws'!"
for i in range(len(items)):
    r = 9 + i; src = first + i
    out(d, f"B{r}", f"={BD}B{src}"); d[f"B{r}"].alignment = Alignment(horizontal="left")
    out(d, f"C{r}", f"={BD}C{src}", USD)
    out(d, f"D{r}", f"=SUMPRODUCT((COLUMN({BD}$D${src}:$K${src})-COLUMN({BD}$D${src})+1<$C$4)*{BD}$D${src}:$K${src})", USD)
    out(d, f"E{r}", f"=INDEX({BD}$D${src}:$K${src},$C$4)", USD, bold=True)
    out(d, f"F{r}", f"=D{r}+E{r}", USD)
    out(d, f"G{r}", f"=C{r}-F{r}", USD)
    out(d, f"H{r}", f"=IFERROR(F{r}/C{r},0)", PCT)
dl = 9 + len(items) - 1; dt = dl + 1
label(d, f"B{dt}", "TOTAL", bold=True)
for c in "CDEFG":
    key(d, f"{c}{dt}", f"=SUM({c}9:{c}{dl})", USD)
key(d, f"H{dt}", f"=IFERROR(F{dt}/C{dt},0)", PCT)
label(d, f"B{dt+2}", "This request (gross)"); out(d, f"E{dt+2}", f"=E{dt}", USD)
label(d, f"B{dt+3}", "Lender-funded share"); out(d, f"E{dt+3}", f"=E{dt}*Setup!$D$11", USD)
label(d, f"B{dt+4}", "Less retainage"); out(d, f"E{dt+4}", f"=-E{dt+3}*Setup!$D$15", USD)
label(d, f"B{dt+5}", "Less draw + inspection fees"); out(d, f"E{dt+5}", f"=IF(E{dt}>0,-(Setup!$D$16+Setup!$D$17),0)", USD)
label(d, f"B{dt+6}", "NET AMOUNT TO BE FUNDED", bold=True); key(d, f"E{dt+6}", f"=SUM(E{dt+3}:E{dt+5})", USD)
label(d, f"B{dt+7}", "Retainage held to date"); out(d, f"E{dt+7}", f"=F{dt}*Setup!$D$11*Setup!$D$15", USD)
d[f"B{dt+9}"] = "Borrower signature: ____________________     Date: ________"; d[f"B{dt+9}"].font = f()
d[f"B{dt+11}"] = "Inspector / lender approval: ____________________     Date: ________"; d[f"B{dt+11}"].font = f()
d.print_area = f"B2:H{dt+11}"; d.page_setup.orientation = "portrait"; d.page_setup.fitToWidth = 1
d.sheet_properties.pageSetUpPr.fitToPage = True

# ---------------- Interest ----------------
it = wb.create_sheet("Interest")
base(it, {"A": 2, "B": 10, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16, "H": 3, "I": 50})
title(it, "B2", "Loan balance & interest by month", "Draws are assumed funded at the start of their month.")
label(it, "B4", "Purchase loan funded at closing", bold=True); out(it, "F4", "=Setup!D6*Setup!D10", USD)
label(it, "B5", "Total loan commitment", bold=True); out(it, "F5", f"=F4+{BD}C{TOT}*Setup!D11", USD)
for i, h in enumerate(["Month", "Draws funded", "Balance (end)", "Interest non-Dutch", "Interest Dutch", "Interest (your method)"]):
    c = it.cell(row=7, column=2 + i, value=h); c.font = Font(name=FONT, bold=True, color="FFFFFF", size=9)
    c.fill = PatternFill("solid", fgColor=NAVY); c.alignment = Alignment(horizontal="center", wrap_text=True)
for m in range(1, 19):
    r = 7 + m
    out(it, f"B{r}", m, NUM)
    out(it, f"C{r}", f"=SUMPRODUCT(({BD}$D$4:$K$4=B{r})*{BD}$D${TOT}:$K${TOT})*Setup!$D$11", USD)
    prev = "$F$4" if m == 1 else f"D{r-1}"
    out(it, f"D{r}", f"=IF(B{r}<=Setup!$D$19,{prev}+C{r},0)", USD)
    out(it, f"E{r}", f"=IF(B{r}<=Setup!$D$19,D{r}*Setup!$D$12/12,0)", USD)
    out(it, f"F{r}", f"=IF(B{r}<=Setup!$D$19,$F$5*Setup!$D$12/12,0)", USD)
    out(it, f"G{r}", f'=IF(Setup!$D$14="Dutch",F{r},E{r})', USD, bold=True)
label(it, "B26", "TOTAL", bold=True)
for c in "CEFG":
    key(it, f"{c}26", f"=SUM({c}8:{c}25)", USD)
label(it, "B27", "Dutch costs you extra:"); out(it, "F27", "=F26-E26", USD, bold=True, color=RED)
note(it, "I8", "Non-Dutch interest grows as draws are funded. Dutch charges the full commitment every month, which is why it costs more.")

# ---------------- Flip P&L ----------------
p = wb.create_sheet("Flip P&L")
base(p, {"A": 2, "B": 44, "C": 18, "D": 3, "E": 56})
title(p, "B2", "Flip P&L and lender ratios", "Everything calculated from the other tabs.")
band(p, "B4:C4", "ALL-IN COST")
lines = [(5, "Purchase price", "=Setup!D6"), (6, "Closing costs at purchase", "=Setup!D6*Setup!D7"), (7, "Rehab budget (incl. contingency)", f"={BD}C{TOT}"),
         (8, "Points", "=Interest!F5*Setup!D13"), (9, "Interest (your method)", "=Interest!G26"),
         (10, "Draw + inspection fees", f"=-SUM({BD}D{TOT+3}:K{TOT+3})"), (11, "Holding costs", "=Setup!D19*Setup!D20"),
         (12, "Selling costs", "=Setup!D8*Setup!D21")]
for r, lab, fx in lines:
    label(p, f"B{r}", lab); out(p, f"C{r}", fx, USD)
label(p, "B13", "TOTAL ALL-IN COST", bold=True); key(p, "C13", "=SUM(C5:C12)", USD)
band(p, "B15:C15", "PROFIT")
label(p, "B16", "Sale price (ARV)"); out(p, "C16", "=Setup!D8", USD)
label(p, "B17", "NET PROFIT", bold=True); key(p, "C17", "=C16-C13", USD)
label(p, "B18", "Loan at payoff"); out(p, "C18", "=Interest!F5", USD)
label(p, "B19", "Your cash invested"); out(p, "C19", "=C13-C12-C18", USD)
note(p, "E19", "All-in cost before selling costs, minus the loan. Selling costs come out of the sale proceeds.")
label(p, "B20", "ROI on your cash", bold=True); key(p, "C20", "=IFERROR(C17/C19,0)", PCT)
label(p, "B21", "Annualized ROI"); out(p, "C21", "=IFERROR((1+C20)^(12/Setup!D19)-1,0)", PCT)
label(p, "B22", "Profit margin on ARV"); out(p, "C22", "=IFERROR(C17/C16,0)", PCT)
note(p, "E22", "Many flippers target 10–15%+ of ARV.")
band(p, "B24:C24", "LENDER RATIOS")
label(p, "B25", "Loan-to-cost (loan ÷ purchase + rehab)"); out(p, "C25", "=IFERROR(Interest!F5/(C5+C7),0)", PCT)
label(p, "B26", "Loan-to-ARV"); out(p, "C26", "=IFERROR(Interest!F5/C16,0)", PCT)
label(p, "B27", "LTC within lender limit?", bold=True); out(p, "C27", '=IF(C25<=Setup!D23,"PASS","FAIL")', None, bold=True)
label(p, "B28", "Loan-to-ARV within lender limit?", bold=True); out(p, "C28", '=IF(C26<=Setup!D24,"PASS","FAIL")', None, bold=True)
label(p, "B29", "Margin ≥ 10% of ARV?", bold=True); out(p, "C29", '=IF(C22>=0.1,"PASS","FAIL")', None, bold=True)
pass_fail(p, "C27:C29", "PASS"); neg_red(p, "C17")
wb.active = 1
import os; os.makedirs(os.path.dirname(OUT), exist_ok=True); wb.save(OUT); print("saved")
