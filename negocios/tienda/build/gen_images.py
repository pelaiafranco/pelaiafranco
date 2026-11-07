"""Generate 5 listing images per product (HTML -> PNG via Playwright), using real values from the recalculated workbooks."""
import os, json, html
from openpyxl import load_workbook

P = "../productos"
def wbv(path): return load_workbook(path, data_only=True)
def usd(v, d=0): 
    s = f"{abs(v):,.{d}f}"; return ("-$" if v < 0 else "$") + s
def pct(v, d=1): return f"{v*100:.{d}f}%"
def x(v): return f"{v:.2f}x"
def ars(v): return "$" + f"{v:,.0f}".replace(",", ".")
def usds(v): return ("-US$" if v < 0 else "US$") + f"{abs(v):,.0f}".replace(",", ".")
def pct_es(v, d=1): return f"{v*100:.{d}f}%".replace(".", ",")

CSS = """
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,600;12..96,800&family=Public+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{width:1500px;height:1125px;overflow:hidden;font-family:'Public Sans',sans-serif;color:#14203A;background:var(--bg)}
:root{--navy:#14203A;--muted:#5A6478;--line:#DFE3EA;--card:#fff;--pos:#1F7A4D;--neg:#B42318;--amb:#A1580B}
.wrap{position:absolute;inset:0;padding:72px 80px;display:flex;flex-direction:column;gap:36px;justify-content:center}
td,th{padding-top:16px!important;padding-bottom:16px!important}
.eyebrow{font:600 20px/1 'IBM Plex Mono',monospace;letter-spacing:.08em;text-transform:uppercase;color:var(--acc)}
h1{font:800 76px/1.02 'Bricolage Grotesque',sans-serif;letter-spacing:-.02em;text-wrap:balance}
h2{font:800 56px/1.05 'Bricolage Grotesque',sans-serif;letter-spacing:-.015em;text-wrap:balance}
.sub{font-size:28px;line-height:1.4;color:var(--muted);max-width:30ch}
.card{background:var(--card);border-radius:22px;box-shadow:0 30px 60px -20px rgba(20,32,58,.28),0 0 0 1px rgba(20,32,58,.06);overflow:hidden}
.bar{height:46px;background:#F3F5F9;border-bottom:1px solid var(--line);display:flex;align-items:center;gap:9px;padding:0 18px}
.bar i{width:13px;height:13px;border-radius:50%;background:#D5DAE3;display:block}
.bar span{margin-left:14px;font:600 16px 'IBM Plex Mono',monospace;color:var(--muted)}
.tabs{display:flex;gap:4px;padding:0 18px;background:#F3F5F9;border-top:1px solid var(--line)}
.tabs b{font:600 15px 'Public Sans';padding:10px 16px;color:var(--muted)} .tabs b.on{background:#fff;color:var(--navy);border:1px solid var(--line);border-top:none}
.num{font-family:'IBM Plex Mono',monospace;font-weight:600;font-variant-numeric:tabular-nums}
.chip{display:inline-block;padding:9px 16px;border-radius:999px;font:700 18px 'Public Sans';color:#fff;letter-spacing:.02em}
.pos{background:var(--pos)} .neg{background:var(--neg)} .amb{background:var(--amb)}
.check{display:flex;gap:16px;align-items:flex-start;font-size:27px;line-height:1.35}
.check:before{content:"";flex:0 0 34px;height:34px;border-radius:50%;background:var(--acc) url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'><path d='M6 12.5l4 4 8-9' fill='none' stroke='white' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'/></svg>") center/22px no-repeat;margin-top:2px}
table{border-collapse:collapse;width:100%}
td,th{padding:13px 20px;border-bottom:1px solid var(--line);font-size:21px;text-align:right}
td:first-child,th:first-child{text-align:left}
th{font:700 15px 'Public Sans';letter-spacing:.06em;text-transform:uppercase;color:#fff;background:var(--navy)}
tr.sec td{background:var(--soft);font:700 15px 'Public Sans';letter-spacing:.07em;text-transform:uppercase;color:var(--acc);padding:9px 20px}
tr.key td{background:var(--soft);font-weight:700}
.foot{position:absolute;left:80px;right:80px;bottom:48px;display:flex;gap:14px;align-items:center}
.pill{border:2px solid rgba(20,32,58,.14);border-radius:999px;padding:10px 20px;font:600 19px 'Public Sans';background:rgba(255,255,255,.6)}
.big{font:800 120px/1 'Bricolage Grotesque';letter-spacing:-.03em}
</style>"""

def page(theme, body):
    acc, soft, bg = theme
    return f"<!doctype html><html><head><meta charset='utf-8'>{CSS}<style>:root{{--acc:{acc};--soft:{soft};--bg:{bg}}}</style></head><body>{body}</body></html>"

def pills(items): return "<div class='foot'>" + "".join(f"<span class='pill'>{html.escape(i)}</span>" for i in items) + "</div>"

def hero(theme, eyebrow, title, sub, checks, window_title, card_html, foot):
    return page(theme, f"""
<div class='wrap' style='flex-direction:row;gap:56px;align-items:center;padding-bottom:140px'>
  <div style='flex:1 1 0;display:flex;flex-direction:column;gap:30px'>
    <span class='eyebrow'>{eyebrow}</span><h1>{title}</h1><p class='sub'>{sub}</p>
    <div style='display:flex;flex-direction:column;gap:16px;margin-top:6px'>{''.join(f"<div class='check'>{c}</div>" for c in checks)}</div>
  </div>
  <div class='card' style='flex:0 0 640px'><div class='bar'><i></i><i></i><i></i><span>{window_title}</span></div>{card_html}</div>
</div>{pills(foot)}""")

def table_img(theme, eyebrow, title, window_title, tabs, thead, rows, foot, note=""):
    tr = ""
    for r in rows:
        if r[0] == "sec": tr += f"<tr class='sec'><td colspan='{len(thead)}'>{r[1]}</td></tr>"; continue
        cls = "key" if r[0] == "key" else ""
        tr += f"<tr class='{cls}'>" + "".join(f"<td class='{'num' if i else ''}'>{v}</td>" for i, v in enumerate(r[1:])) + "</tr>"
    th = "".join(f"<th>{h}</th>" for h in thead)
    tb = "".join(f"<b class='{'on' if i==0 else ''}'>{t}</b>" for i, t in enumerate(tabs))
    return page(theme, f"""
<div class='wrap' style='padding-bottom:140px;gap:30px'>
  <div style='display:flex;flex-direction:column;gap:14px'><span class='eyebrow'>{eyebrow}</span><h2>{title}</h2>{f"<p class='sub' style='max-width:60ch;font-size:24px'>{note}</p>" if note else ''}</div>
  <div class='card'><div class='bar'><i></i><i></i><i></i><span>{window_title}</span></div><table><tr>{th}</tr>{tr}</table><div class='tabs'>{tb}</div></div>
</div>{pills(foot)}""")

def heat_img(theme, eyebrow, title, sub, col_head, row_head, vals, fmt, color_fn, corner, foot):
    head = "".join(f"<th>{c}</th>" for c in col_head)
    body = ""
    for rh, row in zip(row_head, vals):
        body += f"<tr><td class='num' style='font-weight:700'>{rh}</td>" + "".join(
            f"<td class='num' style='background:{color_fn(v)};text-align:center;font-size:24px'>{fmt(v)}</td>" for v in row) + "</tr>"
    return page(theme, f"""
<div class='wrap' style='padding-bottom:140px;gap:30px'>
  <div style='display:flex;flex-direction:column;gap:14px'><span class='eyebrow'>{eyebrow}</span><h2>{title}</h2><p class='sub' style='max-width:62ch;font-size:24px'>{sub}</p></div>
  <div class='card'><table><tr><th>{corner}</th>{head}</tr>{body}</table></div>
</div>{pills(foot)}""")

def inside_img(theme, eyebrow, title, tabs, foot):
    cards = "".join(f"""<div class='card' style='padding:30px 32px;display:flex;flex-direction:column;gap:12px'>
      <span class='eyebrow' style='font-size:16px'>Tab {i+1}</span><div style='font:800 34px/1.1 Bricolage Grotesque'>{t}</div>
      <p style='font-size:21px;line-height:1.45;color:var(--muted)'>{d}</p></div>""" for i, (t, d) in enumerate(tabs))
    n = len(tabs); cols = 3 if n in (3, 5, 6) else 2
    return page(theme, f"""
<div class='wrap' style='padding-bottom:140px'>
  <div style='display:flex;flex-direction:column;gap:14px'><span class='eyebrow'>{eyebrow}</span><h2>{title}</h2></div>
  <div style='display:grid;grid-template-columns:repeat({cols},1fr);gap:24px'>{cards}</div>
</div>{pills(foot)}""")

def steps_img(theme, eyebrow, title, steps, who_title, who, foot):
    st = "".join(f"""<div style='display:flex;gap:24px;align-items:flex-start'>
      <div style='flex:0 0 72px;height:72px;border-radius:50%;background:var(--acc);color:#fff;display:grid;place-items:center;font:800 34px Bricolage Grotesque'>{i+1}</div>
      <div style='display:flex;flex-direction:column;gap:6px'><div style='font:800 34px/1.15 Bricolage Grotesque'>{a}</div><p style='font-size:23px;line-height:1.45;color:var(--muted)'>{b}</p></div></div>""" for i, (a, b) in enumerate(steps))
    wl = "".join(f"<div class='check' style='font-size:24px'>{w}</div>" for w in who)
    return page(theme, f"""
<div class='wrap' style='padding-bottom:140px'>
  <div style='display:flex;flex-direction:column;gap:14px'><span class='eyebrow'>{eyebrow}</span><h2>{title}</h2></div>
  <div style='display:grid;grid-template-columns:1.25fr 1fr;gap:48px;align-items:start'>
    <div style='display:flex;flex-direction:column;gap:34px'>{st}</div>
    <div class='card' style='padding:36px;display:flex;flex-direction:column;gap:18px'><div style='font:800 32px Bricolage Grotesque'>{who_title}</div>{wl}</div>
  </div>
</div>{pills(foot)}""")

def kpis(items):
    """items: (label, value, chip or None)"""
    out = "<div style='padding:30px 34px;display:flex;flex-direction:column;gap:22px'>"
    for lab, val, chip in items:
        c = f"<span class='chip {chip[1]}'>{chip[0]}</span>" if chip else ""
        out += f"<div style='display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid var(--line);padding-bottom:18px'><span style='font-size:22px;color:var(--muted)'>{lab}</span><span style='display:flex;gap:14px;align-items:center'>{c}<span class='num' style='font-size:34px'>{val}</span></span></div>"
    return out + "</div>"

EN_FOOT = ["Excel + Google Sheets", "Instant download", "No macros", "Example pre-filled"]
ES_FOOT = ["Excel + Google Sheets", "Descarga inmediata", "Sin macros", "Ejemplo precargado"]
jobs = {}

# ---------- 1 DSCR ----------
w = wbv(f"{P}/01-dscr-rental-analyzer/DSCR-Rental-Property-Analyzer.xlsx"); a = w["Analyzer"]; s = w["Sensitivity"]
th = ("#2446C7", "#E6EBFB", "#EEF1F8")
v = lambda r, c="D": a[f"{c}{r}"].value
jobs["01-dscr-rental-analyzer"] = [
 hero(th, "DSCR loan calculator · Excel & Google Sheets", "Will this rental qualify for a DSCR loan?",
      "Get the lender's answer in 60 seconds, before you make an offer.",
      ["Lender DSCR, PITIA and a clear QUALIFIES / BELOW verdict", "Max loan and max price you can pay", "Compare 3 properties side by side"],
      "DSCR-Rental-Property-Analyzer.xlsx",
      kpis([("Lender DSCR", x(v(35)), ("QUALIFIES", "pos")), ("PITIA / month", usd(v(34)), None), ("Max loan at 1.25x", usd(v(37)), None),
            ("Max price you can pay", usd(v(38)), None), ("Cash flow / month", usd(v(53)), None), ("Cash-on-cash", pct(v(56)), None)]), EN_FOOT),
 table_img(th, "Lender view + investor view", "3 deals side by side. One verdict each.", "Analyzer", ["Analyzer", "Sensitivity", "Lender Checklist"],
      ["", "Property 1", "Property 2", "Property 3"],
      [("", "Property", v(5), v(5, "E"), v(5, "F")), ("", "Purchase price", *[usd(v(6, c)) for c in "DEF"]), ("", "Monthly rent", *[usd(v(9, c)) for c in "DEF"]),
       ("sec", "Lender view"), ("", "PITIA", *[usd(v(34, c)) for c in "DEF"]), ("key", "Lender DSCR", *[x(v(35, c)) for c in "DEF"]),
       ("", "Max loan at target DSCR", *[usd(v(37, c)) for c in "DEF"]), ("sec", "Investor view"), ("", "NOI / month", *[usd(v(52, c)) for c in "DEF"]),
       ("key", "Cash flow / month", *[usd(v(53, c)) for c in "DEF"]), ("", "Cash-on-cash", *[pct(v(56, c)) for c in "DEF"]),
       ("key", "Verdict", *[f"<span class='chip {'pos' if v(60,c).startswith('STRONG') else 'amb' if v(60,c).startswith('MIXED') else 'neg'}' style='font-size:15px'>{v(60,c).split(':')[0]}</span>" for c in "DEF"])],
      EN_FOOT),
 heat_img(th, "Stress test", "What if rates go up or rent comes in low?", "Lender DSCR for Property 1 at every combination. Green meets your 1.25x target, amber is 1.00x–1.25x, red fails.",
      [usd(s[f"{c}19"].value) for c in "CDEFG"], [pct(s[f"B{r}"].value, 2) for r in range(20, 27)],
      [[s[f"{c}{r}"].value for c in "CDEFG"] for r in range(20, 27)], x,
      lambda v: "#D7F0E2" if v >= 1.25 else "#FBEFDF" if v >= 1 else "#F9DCD9", "Rate ↓ · Rent →", EN_FOOT),
 inside_img(th, "What's inside", "Everything a DSCR lender looks at, in 3 tabs.",
      [("Analyzer", "PITIA, lender DSCR, max loan, max price, rent needed, reserves, cash to close, NOI, cash flow, cap rate and cash-on-cash. For 3 deals at once."),
       ("Sensitivity", "3 stress-test grids: DSCR, monthly cash flow and max purchase price across 7 rates × 5 rent levels."),
       ("Lender Checklist", "Typical 2026 DSCR requirements with automatic PASS/FAIL checks for DSCR, LTV, credit score, reserves and loan size.")], EN_FOOT),
 steps_img(th, "How it works", "From listing to lender-ready in 3 steps.",
      [("Type the deal", "Price, rent, rate, taxes, insurance. Yellow cells only."), ("Read the verdict", "Qualifies or not, max loan, max price and cash flow."),
       ("Stress-test & apply", "Check rate and rent risk, then use the checklist with your lender.")],
      "Made for", ["Rental investors using DSCR / no-income loans", "Buyers comparing several listings", "Agents & loan officers who need a quick pre-check", "BRRRR investors planning the refi"], EN_FOOT)]

# ---------- 2 BRRRR ----------
w = wbv(f"{P}/02-brrrr-calculator/BRRRR-Deal-Analyzer.xlsx"); a = w["BRRRR Analyzer"]; s = w["Stress Test"]
th = ("#127A5A", "#E1F2EB", "#EDF5F1")
v = lambda r: a[f"D{r}"].value
coc = v(57)
jobs["02-brrrr-calculator"] = [
 hero(th, "BRRRR calculator · Excel & Google Sheets", "How much cash do you get back at the refi?",
      "Buy, Rehab, Rent, Refinance, Repeat. See the real numbers before you buy.",
      ["Cash left in the deal + % of your capital recovered", "Max price that gets 100% of your money back", "Refi DSCR, cash flow and a pass/fail scorecard"],
      "BRRRR-Deal-Analyzer.xlsx",
      kpis([("Your cash invested", usd(v(42)), None), ("Cash back at refi", usd(v(47)), None), ("Cash left in the deal", usd(v(48)), None),
            ("Refi DSCR", x(v(54)), None), ("Cash flow / month", usd(v(56)), None), ("Verdict", "", (v(66), "pos"))]), EN_FOOT),
 table_img(th, "The whole deal, phase by phase", "Buy → Rehab → Rent → Refinance → Repeat", "BRRRR Analyzer", ["BRRRR Analyzer", "Stress Test"], ["", "Amount"],
      [("sec", "Cash in"), ("key", "Your cash invested (all-in cost − loan)", usd(v(42))),
       ("sec", "Refinance"), ("", "New refi loan (75% of ARV)", usd(v(44))), ("key", "Cash back to you (after payoff + costs)", usd(v(47))),
       ("key", "Cash left in the deal", usd(v(48))), ("sec", "After the refi"), ("", "Lender DSCR", x(v(54))), ("", "Monthly cash flow", usd(v(56))),
       ("key", "Cash-on-cash on cash left", "INFINITE" if coc == "INFINITE" else pct(coc)), ("sec", "Offer limits"), ("key", "Max price for 100% cash back", usd(v(60)))],
      EN_FOOT),
 heat_img(th, "Stress test", "Low appraisal? Rehab overrun?", "Cash left in the deal across ARV and rehab overruns. Green = all your cash back, amber = up to 10% of all-in cost left in, red = more.",
      [usd(s[f"{c}5"].value) for c in "CDEFGH"], [pct(s[f"B{r}"].value, 0) + " over" for r in range(6, 11)],
      [[s[f"{c}{r}"].value for c in "CDEFGH"] for r in range(6, 11)], lambda v: usd(v),
      lambda val: "#D7F0E2" if val <= 0 else "#FBEFDF" if val <= 0.1 * a["D41"].value else "#F9DCD9", "Overrun ↓ · ARV →", EN_FOOT),
 inside_img(th, "What's inside", "Built for the refi, not just the purchase.",
      [("BRRRR Analyzer", "5 phases, hard-money costs, refi loan, cash back, cash left, equity created, refi DSCR, cash flow, infinite-return check, 70% rule and a solved max price."),
       ("Stress Test", "Cash left in the deal across ARV × rehab overrun, and cash flow across refi rate × rent."),
       ("Scorecard", "4 PASS/FAIL checks and a verdict: GREAT BRRRR, WORKABLE or PASS ON IT.")], EN_FOOT),
 steps_img(th, "How it works", "Know your refi before you buy.",
      [("Enter the deal", "Price, rehab, ARV, hard-money and refi terms, rent."), ("See the refi", "Cash back, cash left and whether it's an infinite return."),
       ("Set your offer", "Max price by your rule and the price that recycles 100% of your cash.")],
      "Made for", ["BRRRR & buy-and-hold investors", "Investors using hard money + DSCR refis", "Wholesalers showing buyers the numbers", "Coaches & mentors teaching BRRRR"], EN_FOOT)]

# ---------- 3 DRAWS ----------
w = wbv(f"{P}/03-construction-draw-tracker/Construction-Draw-Tracker-Flip-Budget.xlsx"); b = w["Budget & Draws"]; d = w["Draw Request"]; it = w["Interest"]; pl = w["Flip P&L"]
th = ("#C2410C", "#FDEBDD", "#F8F1EB")
jobs["03-construction-draw-tracker"] = [
 hero(th, "Construction draw tracker · Excel & Google Sheets", "Run your rehab like a lender does.",
      "Line-item budget, 8 draws, a draw request ready to send and your real flip profit.",
      ["Over-budget lines flagged automatically", "Lender-style draw request for any draw", "Dutch vs non-Dutch interest, LTC and LTV-on-ARV checks"],
      "Construction-Draw-Tracker.xlsx",
      kpis([("Rehab budget", usd(b["C27"].value), None), ("Drawn to date", usd(b["L27"].value), None), ("% complete", pct(b["N27"].value), None),
            ("Interest (non-Dutch)", usd(it["E26"].value), None), ("Net profit", usd(pl["C17"].value), None), ("Loan-to-cost", pct(pl["C25"].value), ("PASS", "pos"))]), EN_FOOT),
 table_img(th, "Budget & draws", "Every line item, every draw, always in sight.", "Budget & Draws", ["Budget & Draws", "Draw Request", "Interest", "Flip P&L"],
      ["Line item", "Budget", "Draw 1", "Draw 2", "Draw 3", "To date", "Status"],
      [("", b[f"B{r}"].value, usd(b[f"C{r}"].value), usd(b[f"D{r}"].value or 0), usd(b[f"E{r}"].value or 0), usd(b[f"F{r}"].value or 0), usd(b[f"L{r}"].value),
        f"<span class='chip {'neg' if b[f'O{r}'].value=='OVER BUDGET' else 'pos' if b[f'O{r}'].value=='Complete' else 'amb'}' style='font-size:14px'>{b[f'O{r}'].value}</span>") for r in [6, 9, 10, 12, 13, 17, 24]]
      + [("key", "TOTAL (21 lines)", usd(b["C27"].value), usd(b["D27"].value), usd(b["E27"].value), usd(b["F27"].value), usd(b["L27"].value), "")], EN_FOOT),
 table_img(th, "Draw request", "Pick a draw. Send it to your lender.", "Draw Request · Draw #3", ["Draw Request", "Budget & Draws"],
      ["Line item", "Budget", "Previous", "This request", "To date", "% done"],
      [("", d[f"B{r}"].value, usd(d[f"C{r}"].value), usd(d[f"D{r}"].value), usd(d[f"E{r}"].value), usd(d[f"F{r}"].value), pct(d[f"H{r}"].value, 0)) for r in [12, 13, 14, 15, 17, 29]]
      + [("key", "TOTAL", usd(d["C30"].value), usd(d["D30"].value), usd(d["E30"].value), usd(d["F30"].value), pct(d["H30"].value, 0)),
         ("", "Less retainage (10%)", "", "", usd(d["E34"].value), "", ""), ("", "Less draw + inspection fees", "", "", usd(d["E35"].value), "", ""),
         ("key", "NET AMOUNT TO BE FUNDED", "", "", usd(d["E36"].value), "", "")], EN_FOOT),
 inside_img(th, "What's inside", "5 tabs. The whole project, lender-ready.",
      [("Setup", "Purchase, construction loan, interest method, retainage, draw and inspection fees, timeline and exit."),
       ("Budget & Draws", "21 line items × 8 draws with remaining, % complete and over-budget alerts."),
       ("Draw Request", "Previous, this request, to date, retainage, fees and net funding for any draw."),
       ("Interest", "18-month balance and interest schedule. Shows what Dutch interest really costs."),
       ("Flip P&L", "All-in cost, profit, ROI, annualized ROI, LTC and LTV-on-ARV with PASS/FAIL checks.")], EN_FOOT),
 steps_img(th, "How it works", "Budget once. Track every draw.",
      [("Set up the project", "Loan terms, budget lines and timeline."), ("Log each draw", "Amount per line and the month it was funded."),
       ("Send & track", "Print the draw request, watch interest and profit update live.")],
      "Made for", ["Fix-and-flip investors using hard money", "Small builders & GCs managing draws", "Private lenders reviewing draw requests", "Owners doing a major renovation"], EN_FOOT)]

# ---------- 4 WATERFALL ----------
w = wbv(f"{P}/04-jv-waterfall/Real-Estate-JV-Waterfall-Model.xlsx"); s = w["Summary"]; di = w["Deal Inputs"]
th = ("#6D28D9", "#EEE7FB", "#F3F0FA")
jobs["04-jv-waterfall"] = [
 hero(th, "JV waterfall model · Excel & Google Sheets", "Split the profits like a fund does.",
      "Preferred return, IRR hurdle and promote, year by year, for your investors and for you.",
      ["8% pref + IRR hurdle + 2 promote tiers", "LP & GP IRR, equity multiple, promote earned", "Every dollar tracked tier by tier, with a check"],
      "Real-Estate-JV-Waterfall.xlsx",
      kpis([("Investors (LP) IRR", pct(s["C9"].value), None), ("LP equity multiple", x(s["C10"].value), None), ("Sponsor (GP) IRR", pct(s["D9"].value), None),
            ("GP equity multiple", x(s["D10"].value), None), ("Promote earned by GP", usd(s["D15"].value), None), ("Allocation check", "$0", ("BALANCED", "pos"))]), EN_FOOT),
 table_img(th, "Returns summary", "LP vs GP, side by side.", "Summary", ["Summary", "Waterfall", "Deal Inputs"], ["", "Investors (LP)", "Sponsor (GP)", "Project"],
      [("", "Equity contributed", usd(s["C6"].value), usd(s["D6"].value), usd(s["E6"].value)), ("", "Total distributions", usd(s["C7"].value), usd(s["D7"].value), usd(s["E7"].value)),
       ("", "Profit", usd(s["C8"].value), usd(s["D8"].value), usd(s["E8"].value)), ("key", "IRR", pct(s["C9"].value), pct(s["D9"].value), pct(s["E9"].value)),
       ("key", "Equity multiple", x(s["C10"].value), x(s["D10"].value), x(s["E10"].value)), ("sec", "Distributions by tier"),
       ("", "Tier 1 · Capital + 8% pref", usd(s["C18"].value), usd(s["D18"].value), usd(s["E18"].value)),
       ("", "Tier 2 · Up to 13% IRR (80/20)", usd(s["C19"].value), usd(s["D19"].value), usd(s["E19"].value)),
       ("", "Tier 3 · Above 13% IRR (70/30)", usd(s["C20"].value), usd(s["D20"].value), usd(s["E20"].value)),
       ("key", "Promote earned by GP", "", usd(s["D15"].value), "")], EN_FOOT),
 table_img(th, "Year-by-year waterfall", "See where every dollar goes.", "Waterfall", ["Waterfall", "Summary"], ["", "Yr 1", "Yr 2", "Yr 3", "Yr 4", "Yr 5"],
      [("", "Cash available", *[usd(w["Waterfall"][f"{c}5"].value) for c in "DEFGH"]), ("sec", "Tier 1 · Pref"),
       ("", "Pref balance (start)", *[usd(w["Waterfall"][f"{c}8"].value) for c in "DEFGH"]), ("", "Tier 1 paid", *[usd(w["Waterfall"][f"{c}9"].value) for c in "DEFGH"]),
       ("sec", "Tier 2 · Hurdle"), ("", "Tier 2 paid", *[usd(w["Waterfall"][f"{c}16"].value) for c in "DEFGH"]),
       ("sec", "Tier 3"), ("", "Tier 3 paid", *[usd(w["Waterfall"][f"{c}21"].value) for c in "DEFGH"]),
       ("key", "To investors (LP)", *[usd(w["Waterfall"][f"{c}25"].value) for c in "DEFGH"]), ("key", "To sponsor (GP)", *[usd(w["Waterfall"][f"{c}26"].value) for c in "DEFGH"])], EN_FOOT),
 inside_img(th, "What's inside", "Institutional logic, plain-English tabs.",
      [("Deal Inputs", "Equity, GP co-invest, pref, IRR hurdle, two promote tiers, 10 years of cash flow, sale proceeds and an exit-price scenario."),
       ("Waterfall", "Compounding pref and hurdle balances, 3 tiers, LP & GP distributions every year."),
       ("Summary", "IRR, equity multiple, profit share, promote earned, distributions by tier and a stacked chart.")], EN_FOOT),
 steps_img(th, "How it works", "Structure your raise in minutes.",
      [("Set the terms", "Equity, co-invest, pref, hurdle and promote."), ("Add cash flows", "Yearly distributions and net sale proceeds."),
       ("Show investors", "LP IRR and multiple. Try exit scenarios live.")],
      "Made for", ["Syndicators & first-time sponsors", "Investors reviewing a sponsor's offer", "JV partners negotiating a split", "Analysts learning waterfalls"], EN_FOOT)]

# ---------- 5 ARGENTINA ----------
w = wbv(f"{P}/05-tablero-inversiones-ars-usd/Tablero-Inversiones-Pesos-Dolares.xlsx"); t = w["Tablero"]; inv = w["Inversiones"]; pf = w["Plazo fijo vs dólar"]
th = ("#0B6FA4", "#E0F0F8", "#EEF5F9")
jobs["05-tablero-inversiones-ars-usd"] = [
 hero(th, "Planilla para Excel y Google Sheets", "¿Estás ganando en dólares o solo en pesos?",
      "Todas tus inversiones medidas en pesos, en dólar MEP y contra la inflación.",
      ["Resultado en USD de cada inversión", "Rendimiento real contra el IPC del INDEC", "Plazo fijo vs. dólar: tu dólar de equilibrio"],
      "Tablero-Inversiones.xlsx",
      kpis([("Valor total (USD MEP)", usds(t["C5"].value), None), ("Valor total (ARS)", ars(t["C6"].value), None), ("Ganancia en dólares", usds(t["C8"].value), None),
            ("Rendimiento en USD", pct_es(t["C9"].value), None), ("Inversiones", str(t["C11"].value), None), ("MEP usado", ars(t["C12"].value), None)]), ES_FOOT),
 table_img(th, "Mis inversiones", "Cada inversión, en pesos y en dólares.", "Inversiones", ["Inversiones", "Tablero", "Mercado", "Plazo fijo vs dólar"],
      ["Instrumento", "Tipo", "Invertido USD", "Hoy USD", "Resultado", "Rend. anual", "Real ARS"],
      [("", inv[f"C{r}"].value, inv[f"D{r}"].value, usds(inv[f"I{r}"].value), usds(inv[f"J{r}"].value), usds(inv[f"L{r}"].value),
        pct_es(inv[f"O{r}"].value) if isinstance(inv[f"O{r}"].value, (int, float)) else inv[f"O{r}"].value, pct_es(inv[f"Q{r}"].value)) for r in range(5, 12)], ES_FOOT,
      note="Valores de ejemplo: cargás los tuyos y el resto se calcula solo."),
 table_img(th, "Tablero", "Tu cartera, de un vistazo.", "Tablero", ["Tablero", "Inversiones", "Mercado"], ["Tipo", "Valor (USD)", "% cartera", "Resultado (USD)"],
      [("", t[f"B{r}"].value, usds(t[f"C{r}"].value), pct_es(t[f"D{r}"].value), usds(t[f"E{r}"].value)) for r in range(16, 26) if t[f"C{r}"].value]
      + [("key", "TOTAL", usds(t["C5"].value), "100,0%", usds(t["C8"].value))], ES_FOOT),
 table_img(th, "Plazo fijo vs dólar", "¿Pesos a tasa o dólar? Calculalo antes.", "Plazo fijo vs dólar", ["Plazo fijo vs dólar", "Tablero"],
      ["$1.000.000 a 30 días", "Plazo fijo", "Caución", "FCI MM"],
      [("", "TNA", *[pct_es(pf[f"{c}9"].value) for c in "CDE"]), ("", "TEA", *[pct_es(pf[f"{c}11"].value) for c in "CDE"]),
       ("", "Monto al final", *[ars(pf[f"{c}13"].value) for c in "CDE"]), ("key", "Dólar de equilibrio", *[ars(pf[f"{c}14"].value) for c in "CDE"]),
       ("", "Suba del MEP que empata", *[pct_es(pf[f"{c}15"].value, 2) for c in "CDE"]),
       ("key", "Si el MEP termina en $1.486", *[f"<span class='chip {'pos' if pf[f'{c}19'].value=='PESOS' else 'amb'}' style='font-size:15px'>{pf[f'{c}19'].value}</span>" for c in "CDE"])],
      ES_FOOT, note="Si el MEP queda por debajo del dólar de equilibrio, convenía la alternativa en pesos."),
 steps_img(th, "Cómo funciona", "En 3 pasos sabés si le ganás a la inflación.",
      [("Cargá el mercado", "Dólar MEP de hoy e inflación mensual del INDEC."), ("Cargá tus inversiones", "Plazo fijo, caución, FCI, bonos, ONs, CEDEARs, dólar o cripto."),
       ("Mirá el tablero", "Total en USD y ARS, ganancia en dólares y rendimiento real.")],
      "Ideal para", ["Ahorristas que invierten en pesos y en dólares", "Quienes usan un broker (IOL, Cocos, Balanz…)", "Saber si de verdad le ganan al dólar y a la inflación", "Ordenar todo en un solo lugar"], ES_FOOT)]

for slug, pages in jobs.items():
    d = f"{P}/{slug}/images/src"; os.makedirs(d, exist_ok=True)
    for i, h in enumerate(pages, 1):
        open(f"{d}/{i:02d}.html", "w").write(h)
print(json.dumps({k: len(v) for k, v in jobs.items()}))
