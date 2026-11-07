from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.chart import BarChart, Reference
from common import *

OUT = "../productos/04-jv-waterfall/Real-Estate-JV-Waterfall-Model.xlsx"
wb = Workbook()
start_sheet(wb, "Real Estate JV Equity Waterfall",
    "Split profits between investors (LP) and sponsor (GP) like institutional funds do: preferred return, IRR hurdles and promote, year by year.",
    ["In 'Deal Inputs', enter total equity, the sponsor's co-invest, the pref, the hurdle and the promote splits. An example deal is pre-filled.",
     "Enter the yearly cash flow available for distribution and the net sale proceeds in the exit year (up to 10 years).",
     "Read 'Summary' for LP and GP IRR, equity multiple and promote earned. 'Waterfall' shows every dollar, tier by tier."],
    [("Deal Inputs", "Equity, co-invest, 8% pref, IRR hurdle, two promote tiers, yearly cash flows, sale proceeds and an exit-price scenario lever."),
     ("Waterfall", "Year-by-year IRR hurdle accounting (compounding balances) across 3 tiers, for LP and GP."),
     ("Summary", "LP vs GP: distributions, profit, IRR, equity multiple, promote earned, distributions by tier, plus a chart.")],
    DISCLAIMER_EN + " Waterfall structures vary: confirm against your operating agreement.")

di = wb.create_sheet("Deal Inputs")
base(di, {"A": 2, "B": 40, "C": 10, "D": 18, "E": 3, "F": 60})
title(di, "B2", "Deal inputs", "Yellow = your inputs.")
def r_(r, lab, u, v, fmt, nt=""):
    label(di, f"B{r}", lab); unit(di, f"C{r}", u); inp(di, f"D{r}", v, fmt)
    if nt: note(di, f"F{r}", nt)
band(di, "B4:D4", "EQUITY")
r_(5, "Deal name", "", "Riverside 48-unit value-add", None)
r_(6, "Total equity raised", "$", 4000000, USD)
r_(7, "Sponsor (GP) co-invest", "% equity", 0.10, PCT, "GP's own money in the deal. 5–10% is common.")
band(di, "B8:D8", "WATERFALL TERMS")
r_(9, "Tier 1: preferred return (compounding)", "% / yr", 0.08, PCT, "Paid pro rata on all equity until investors earn this IRR.")
r_(10, "Tier 2 hurdle: LP IRR up to", "% / yr", 0.13, PCT, "Between the pref and this IRR, the GP earns the Tier 2 promote.")
r_(11, "Tier 2: GP promote", "% of tier", 0.20, PCT, "e.g., 80/20 split.")
r_(12, "Tier 3: GP promote above hurdle", "% of tier", 0.30, PCT, "e.g., 70/30 split.")
note(di, "F13", "Promote = sponsor's extra share. The rest of each tier is split pro rata by equity (so the GP's co-invest also earns its share).")
band(di, "B14:D14", "SCENARIO")
r_(15, "Exit price scenario vs. base", "%", 0.0, PCT, "Try -10% or +10% to see how the split changes.")
band(di, "B17:D17", "CASH FLOW TO EQUITY (after debt service)")
label(di, "B18", "Year", bold=True); di["C18"] = "Operating"; di["D18"] = "Net sale proceeds"
for c in "CD": di[f"{c}18"].font = f(True, 9, GREY)
ops = [180000, 260000, 320000, 360000, 380000, 0, 0, 0, 0, 0]
sale = [0, 0, 0, 0, 6400000, 0, 0, 0, 0, 0]
for y in range(1, 11):
    r = 18 + y
    label(di, f"B{r}", f"Year {y}")
    inp(di, f"C{r}", ops[y-1], USD); inp(di, f"D{r}", sale[y-1], USD)
di.column_dimensions["C"].width = 16
note(di, "F19", "Net sale proceeds = sale price − loan payoff − selling costs. Put it in the exit year; leave later years at 0.")

w = wb.create_sheet("Waterfall")
base(w, {"A": 2, "B": 38, **{c: 13 for c in "CDEFGHIJKLM"}})
title(w, "B2", "Waterfall by year", "Tier 1 pref → Tier 2 up to hurdle IRR → Tier 3 above.")
yc = "CDEFGHIJKLM"  # year 0..10
D = "'Deal Inputs'!"
for i, c in enumerate(yc):
    w[f"{c}4"] = f"Year {i}"; w[f"{c}4"].font = Font(name=FONT, bold=True, color="FFFFFF"); w[f"{c}4"].fill = PatternFill("solid", fgColor=NAVY)
    w[f"{c}4"].alignment = Alignment(horizontal="center")
labels = {5: "Cash available to distribute", 6: "LP share of equity", 8: "TIER 1 · Pref balance (start, compounded)", 9: "Tier 1 distribution (total)",
          10: "  to LP", 11: "  to GP", 12: "Remaining after Tier 1", 14: "TIER 2 · Hurdle balance (start, compounded)", 15: "Hurdle balance after Tier 1",
          16: "Tier 2 distribution (total)", 17: "  to LP", 18: "  to GP", 19: "Remaining after Tier 2", 21: "TIER 3 distribution (total)", 22: "  to LP", 23: "  to GP",
          25: "LP total distributions", 26: "GP total distributions", 27: "Pref balance (end)", 28: "Hurdle balance (end)",
          30: "LP net cash flow", 31: "GP net cash flow", 32: "Project net cash flow"}
for r, t in labels.items():
    label(w, f"B{r}", t, bold=t.isupper() or t.startswith("TIER") or "net" in t)
out(w, "C6", f"=1-{D}D7", PCT)
# year 0
out(w, "C27", f"={D}D6*C6", USD); out(w, "C28", f"={D}D6*C6", USD)
out(w, "C30", f"=-{D}D6*C6", USD); out(w, "C31", f"=-{D}D6*{D}D7", USD); out(w, "C32", f"=-{D}D6", USD)
for i in range(1, 11):
    c = yc[i]; p = yc[i-1]; ir = 18 + i
    out(w, f"{c}5", f"={D}C{ir}+{D}D{ir}*(1+{D}$D$15)", USD, bold=True)
    out(w, f"{c}6", "=$C$6", PCT)
    out(w, f"{c}8", f"={p}27*(1+{D}$D$9)", USD)
    out(w, f"{c}9", f"=MIN({c}5,IFERROR({c}8/{c}6,0))", USD)
    out(w, f"{c}10", f"={c}9*{c}6", USD); out(w, f"{c}11", f"={c}9-{c}10", USD)
    out(w, f"{c}12", f"={c}5-{c}9", USD)
    out(w, f"{c}14", f"={p}28*(1+{D}$D$10)", USD)
    out(w, f"{c}15", f"=MAX(0,{c}14-{c}10)", USD)
    out(w, f"{c}16", f"=MIN({c}12,IFERROR({c}15/((1-{D}$D$11)*{c}6),0))", USD)
    out(w, f"{c}17", f"={c}16*(1-{D}$D$11)*{c}6", USD); out(w, f"{c}18", f"={c}16-{c}17", USD)
    out(w, f"{c}19", f"={c}12-{c}16", USD)
    out(w, f"{c}21", f"={c}19", USD)
    out(w, f"{c}22", f"={c}21*(1-{D}$D$12)*{c}6", USD); out(w, f"{c}23", f"={c}21-{c}22", USD)
    out(w, f"{c}25", f"={c}10+{c}17+{c}22", USD, bold=True); out(w, f"{c}26", f"={c}11+{c}18+{c}23", USD, bold=True)
    out(w, f"{c}27", f"=MAX(0,{c}8-{c}25)", USD); out(w, f"{c}28", f"=MAX(0,{c}14-{c}25)", USD)
    out(w, f"{c}30", f"={c}25", USD); out(w, f"{c}31", f"={c}26", USD); out(w, f"{c}32", f"={c}5", USD)
w.freeze_panes = "C5"

s = wb.create_sheet("Summary")
base(s, {"A": 2, "B": 36, "C": 18, "D": 18, "E": 18, "F": 3})
title(s, "B2", "Returns summary", None)
out(s, "B3", f"={D}D5"); s["B3"].alignment = Alignment(horizontal="left"); s["B3"].font = f(False, 11, GREY, italic=True)
for c, t in zip("CDE", ["Investors (LP)", "Sponsor (GP)", "Total project"]):
    s[f"{c}5"] = t; s[f"{c}5"].font = Font(name=FONT, bold=True, color="FFFFFF"); s[f"{c}5"].fill = PatternFill("solid", fgColor=NAVY)
    s[f"{c}5"].alignment = Alignment(horizontal="center")
W = "Waterfall!"
rows = [(6, "Equity contributed", f"=-{W}C30", f"=-{W}C31", f"=-{W}C32", USD, False),
        (7, "Total distributions", f"=SUM({W}D25:M25)", f"=SUM({W}D26:M26)", f"=SUM({W}D5:M5)", USD, False),
        (8, "Profit", "=C7-C6", "=D7-D6", "=E7-E6", USD, False),
        (9, "IRR", f"=IFERROR(IRR({W}C30:M30),0)", f"=IFERROR(IRR({W}C31:M31),0)", f"=IFERROR(IRR({W}C32:M32),0)", PCT, True),
        (10, "Equity multiple", "=IFERROR(C7/C6,0)", "=IFERROR(D7/D6,0)", "=IFERROR(E7/E6,0)", MULT, True),
        (11, "Share of total profit", "=IFERROR(C8/E8,0)", "=IFERROR(D8/E8,0)", "=IFERROR(E8/E8,0)", PCT, False)]
for r, lab, a, b_, c_, fmt, k in rows:
    label(s, f"B{r}", lab, bold=k)
    for col, fx in zip("CDE", [a, b_, c_]):
        (key if k else out)(s, f"{col}{r}", fx, fmt)
band(s, "B13:E13", "SPONSOR PROMOTE")
label(s, "B14", "GP distributions if split purely pro rata"); out(s, "D14", f"=E7*{D}D7", USD)
label(s, "B15", "PROMOTE EARNED BY GP", bold=True); key(s, "D15", "=D7-D14", USD)
band(s, "B17:E17", "DISTRIBUTIONS BY TIER")
for r, lab, lp, gp in [(18, "Tier 1 · Return of capital + pref", f"=SUM({W}D10:M10)", f"=SUM({W}D11:M11)"),
                       (19, "Tier 2 · Up to hurdle IRR", f"=SUM({W}D17:M17)", f"=SUM({W}D18:M18)"),
                       (20, "Tier 3 · Above hurdle", f"=SUM({W}D22:M22)", f"=SUM({W}D23:M23)")]:
    label(s, f"B{r}", lab); out(s, f"C{r}", lp, USD); out(s, f"D{r}", gp, USD); out(s, f"E{r}", f"=C{r}+D{r}", USD)
label(s, "B21", "Check (must equal total distributions)"); out(s, "E21", "=SUM(E18:E20)-E7", USD)
note(s, "C21", "")
note(s, "B22", "Zero = every dollar was allocated.")
band(s, "B23:E23", "CHART DATA: DISTRIBUTIONS BY YEAR")
for i in range(1, 11):
    r = 24 + i
    s[f"B{r}"] = f"Year {i}"; s[f"B{r}"].font = f()
    out(s, f"C{r}", f"={W}{yc[i]}25", USD); out(s, f"D{r}", f"={W}{yc[i]}26", USD)
s["B24"] = "Year"; s["C24"] = "LP"; s["D24"] = "GP"
for c in "BCD": s[f"{c}24"].font = f(True, 9, GREY)
ch = BarChart(); ch.type = "col"; ch.grouping = "stacked"; ch.overlap = 100
ch.title = "Distributions by year (LP vs GP)"; ch.y_axis.title = "$"; ch.height = 8; ch.width = 18
data = Reference(s, min_col=3, max_col=4, min_row=24, max_row=34); cats = Reference(s, min_col=2, min_row=25, max_row=34)
ch.add_data(data, titles_from_data=True); ch.set_categories(cats)
ch.series[0].graphicalProperties.solidFill = ACCENT; ch.series[1].graphicalProperties.solidFill = "F0A85A"
s.add_chart(ch, "G5")
wb.active = 3
import os; os.makedirs(os.path.dirname(OUT), exist_ok=True); wb.save(OUT); print("saved")
