from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from datetime import date
from common import *

OUT = "../productos/05-tablero-inversiones-ars-usd/Tablero-Inversiones-Pesos-Dolares.xlsx"
wb = Workbook()
start_sheet(wb, "Tablero de Inversiones en Pesos y Dólares",
    "Sabé de verdad si estás ganando: cada inversión medida en pesos, en dólares MEP y contra la inflación, en un solo tablero.",
    ["En 'Mercado', cargá el dólar MEP de hoy y la inflación mensual del INDEC (se publica a mediados de cada mes).",
     "En 'Inversiones', cargá cada compra: fecha, instrumento, tipo, moneda, monto y el MEP de ese día. Actualizá el 'valor actual' cuando quieras.",
     "Mirá el 'Tablero': cuánto tenés en USD y en ARS, cuánto ganaste en dólares, el rendimiento real y la distribución. Usá 'Plazo fijo vs dólar' antes de decidir."],
    [("Mercado", "Dólar MEP, CCL y oficial de hoy + inflación mensual del INDEC (24 meses), con el índice acumulado calculado solo."),
     ("Inversiones", "Hasta 150 inversiones: plazo fijo, caución, FCI, bonos, ONs, CEDEARs, acciones, dólar MEP y cripto. Resultado en USD, rendimiento anualizado y rendimiento real."),
     ("Tablero", "Total en USD y ARS, ganancia en dólares, rendimiento real contra la inflación, distribución por tipo y por moneda, con gráfico."),
     ("Plazo fijo vs dólar", "Compará plazo fijo, caución y FCI: cuánto cobrás y el dólar de equilibrio, o sea, a qué precio del MEP te conviene más el dólar.")],
    DISCLAIMER_ES, lang="es")

m = wb.create_sheet("Mercado")
base(m, {"A": 2, "B": 34, "C": 16, "D": 16, "E": 3, "F": 56})
title(m, "B2", "Datos de mercado", "Actualizalos cuando quieras: todo el archivo se recalcula.")
band(m, "B4:C4", "DÓLAR HOY")
for r, lab, v in [(5, "Dólar MEP", 1450), (6, "Dólar CCL", 1480), (7, "Dólar oficial (BNA venta)", 1420)]:
    label(m, f"B{r}", lab); inp(m, f"C{r}", v, ARS)
note(m, "F5", "Fuente sugerida: dolarapi.com, la web de tu broker o los diarios. Valores precargados de EJEMPLO: reemplazalos.")
label(m, "B8", "Fecha de los datos"); inp(m, "C8", date(2026, 11, 6), "dd/mm/yyyy")
band(m, "B10:D10", "INFLACIÓN MENSUAL (IPC INDEC, NIVEL GENERAL)")
for c, t in zip("BCD", ["Mes", "Inflación del mes", "Índice acumulado"]):
    m[f"{c}11"] = t; m[f"{c}11"].font = f(True, 9, GREY)
ex_ipc = [0.022, 0.024, 0.037, 0.028, 0.015, 0.016, 0.019, 0.019, 0.021, 0.023, 0.025, 0.028,
          0.022, 0.021, 0.024, 0.026, 0.018, 0.017, 0.019, 0.020, 0.021, 0.022, None, None]
for i in range(24):
    r = 12 + i; y = 2025 + (i // 12); mo = i % 12 + 1
    inp(m, f"B{r}", date(y, mo, 1), "mmm-yyyy")
    inp(m, f"C{r}", ex_ipc[i], PCT)
    prev = "100" if i == 0 else f"D{r-1}"
    out(m, f"D{r}", f'=IF(C{r}="",{prev},{prev}*(1+C{r}))' if i else f'=IF(C{r}="",100,100*(1+C{r}))', '0.00')
label(m, "B37", "Último índice cargado", bold=True); key(m, "D37", "=INDEX(D12:D35,COUNT(C12:C35))", '0.00')
label(m, "B38", "Inflación acumulada del período", bold=True); out(m, "D38", "=D37/100-1", PCT)
note(m, "F12", "⚠ Los porcentajes precargados son de EJEMPLO para que veas cómo funciona. Reemplazalos con los datos oficiales del INDEC (indec.gob.ar). Podés cambiar los meses de la columna B para empezar desde la fecha que quieras.")
m.merge_cells("F12:F16")

inv = wb.create_sheet("Inversiones")
cols = {"A": 2, "B": 12, "C": 22, "D": 13, "E": 8, "F": 14, "G": 11, "H": 14, "I": 13, "J": 13, "K": 15, "L": 13, "M": 11, "N": 8, "O": 12, "P": 11, "Q": 12}
base(inv, cols)
title(inv, "B2", "Mis inversiones", "Amarillo = lo que cargás. El resto se calcula solo.")
hdr = ["Fecha compra", "Instrumento", "Tipo", "Moneda", "Monto invertido", "MEP al comprar", "Valor actual", "Invertido (USD)", "Valor actual (USD)",
       "Valor actual (ARS)", "Resultado (USD)", "Rend. USD", "Días", "Rend. anual USD", "Inflación desde compra", "Rend. real ARS"]
for i, h in enumerate(hdr):
    c = inv.cell(row=4, column=2 + i, value=h); c.font = Font(name=FONT, bold=True, color="FFFFFF", size=9)
    c.fill = PatternFill("solid", fgColor=NAVY); c.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
inv.row_dimensions[4].height = 32
tipos = ["Plazo fijo", "Caución", "FCI", "Bono", "ON", "CEDEAR", "Acción", "Dólar MEP", "Cripto", "Otro"]
ex = [(date(2025, 3, 10), "Dólar MEP ahorro", "Dólar MEP", "USD", 5000, 1230, 5000),
      (date(2025, 4, 2), "ON YPF 2031 (ejemplo)", "ON", "USD", 3000, 1240, 3240),
      (date(2025, 6, 15), "CEDEAR SPY (ejemplo)", "CEDEAR", "ARS", 2500000, 1190, 3350000),
      (date(2025, 9, 1), "AL30 (ejemplo)", "Bono", "ARS", 1800000, 1350, 2300000),
      (date(2026, 1, 20), "FCI money market", "FCI", "ARS", 1500000, 1420, 1690000),
      (date(2026, 8, 5), "Plazo fijo 30 días", "Plazo fijo", "ARS", 1000000, 1440, 1032000),
      (date(2026, 10, 1), "Caución 7 días", "Caución", "ARS", 800000, 1445, 806500)]
MK = "Mercado!"
N = 150
for i in range(N):
    r = 5 + i
    vals = ex[i] if i < len(ex) else (None,) * 7
    for c, v, fmt in zip("BCDEFGH", vals, ["dd/mm/yyyy", None, None, None, '#,##0', ARS, '#,##0']):
        inp(inv, f"{c}{r}", v, fmt)
    inv[f"C{r}"].alignment = Alignment(horizontal="left")
    ok = f'$B{r}<>""'
    out(inv, f"I{r}", f'=IF({ok},IF(E{r}="USD",F{r},IFERROR(F{r}/G{r},0)),"")', USDS)
    out(inv, f"J{r}", f'=IF({ok},IF(E{r}="USD",H{r},IFERROR(H{r}/{MK}$C$5,0)),"")', USDS)
    out(inv, f"K{r}", f'=IF({ok},IF(E{r}="ARS",H{r},H{r}*{MK}$C$5),"")', ARS)
    out(inv, f"L{r}", f'=IF({ok},J{r}-I{r},"")', USDS)
    out(inv, f"M{r}", f'=IF({ok},IFERROR(L{r}/I{r},0),"")', PCT)
    out(inv, f"N{r}", f'=IF({ok},{MK}$C$8-B{r},"")', NUM)
    out(inv, f"O{r}", f'=IF({ok},IF(N{r}>=30,IFERROR((1+M{r})^(365/N{r})-1,0),"< 30 días"),"")', PCT)
    ipc_buy = f"IFERROR(INDEX({MK}$D$12:$D$35,MATCH(DATE(YEAR(B{r}),MONTH(B{r}),1),{MK}$B$12:$B$35,0)-1),100)"
    out(inv, f"P{r}", f'=IF({ok},IFERROR({MK}$D$37/IF(MATCH(DATE(YEAR(B{r}),MONTH(B{r}),1),{MK}$B$12:$B$35,0)=1,100,{ipc_buy})-1,0),"")', PCT)
    inv_ars = f'IF(E{r}="ARS",F{r},F{r}*G{r})'
    out(inv, f"Q{r}", f'=IF({ok},IFERROR((K{r}/{inv_ars})/(1+P{r})-1,0),"")', PCT)
dropdown(inv, f"D5:D{4+N}", tipos); dropdown(inv, f"E5:E{4+N}", ["ARS", "USD"])
last = 4 + N
for rng in [f"L5:M{last}", f"Q5:Q{last}"]:
    neg_red(inv, rng)
inv.freeze_panes = "D5"
note(inv, "B157", "Valor actual = lo que vale hoy en la moneda de la inversión (mirá tu broker). 'Inflación desde compra' toma el IPC desde el mes anterior a la compra hasta el último mes cargado. Las filas precargadas son ejemplos: borralas y cargá las tuyas.")
inv.merge_cells("B157:Q158")

t = wb.create_sheet("Tablero")
base(t, {"A": 2, "B": 30, "C": 18, "D": 18, "E": 14, "F": 3})
title(t, "B2", "Tablero", None)
I = "Inversiones!"
R = f"$5:$" ; rng = lambda c: f"{I}${c}$5:${c}${last}"
band(t, "B4:D4", "RESUMEN")
items = [(5, "Valor actual total (USD MEP)", f"=SUM({rng('J')})", USDS, True), (6, "Valor actual total (ARS)", f"=SUM({rng('K')})", ARS, True),
         (7, "Total invertido (USD)", f"=SUM({rng('I')})", USDS, False), (8, "GANANCIA EN DÓLARES", "=C5-C7", USDS, True),
         (9, "Rendimiento en USD", "=IFERROR(C8/C7,0)", PCT, True),
         (10, "Rendimiento real promedio (ARS, ponderado)", f"=IFERROR(SUMPRODUCT(({rng('B')}<>\"\")*{rng('K')},{rng('Q')})/C6,0)", PCT, False),
         (11, "Cantidad de inversiones", f"=COUNT({rng('B')})", NUM, False),
         (12, "Dólar MEP usado", f"={MK}C5", ARS, False)]
for r, lab, fx, fmt, k in items:
    label(t, f"B{r}", lab, bold=k); (key if k else out)(t, f"C{r}", fx, fmt)
neg_red(t, "C8:C10")
band(t, "B14:E14", "DISTRIBUCIÓN POR TIPO")
for c, h in zip("BCDE", ["Tipo", "Valor (USD)", "% de la cartera", "Resultado (USD)"]):
    t[f"{c}15"] = h; t[f"{c}15"].font = f(True, 9, GREY)
for i, tp in enumerate(tipos):
    r = 16 + i
    t[f"B{r}"] = tp; t[f"B{r}"].font = f()
    out(t, f"C{r}", f'=SUMIFS({rng("J")},{rng("D")},B{r})', USDS)
    out(t, f"D{r}", f"=IFERROR(C{r}/$C$5,0)", PCT)
    out(t, f"E{r}", f'=SUMIFS({rng("L")},{rng("D")},B{r})', USDS)
neg_red(t, "E16:E25")
band(t, "B27:E27", "DISTRIBUCIÓN POR MONEDA")
for i, mo in enumerate(["ARS", "USD"]):
    r = 28 + i
    t[f"B{r}"] = "En pesos" if mo == "ARS" else "En dólares"; t[f"B{r}"].font = f()
    out(t, f"C{r}", f'=SUMIFS({rng("J")},{rng("E")},"{mo}")', USDS); out(t, f"D{r}", f"=IFERROR(C{r}/$C$5,0)", PCT)
pie = PieChart(); pie.title = "Cartera por tipo (USD)"; pie.height = 9; pie.width = 12
pie.add_data(Reference(t, min_col=3, min_row=15, max_row=25), titles_from_data=True)
pie.set_categories(Reference(t, min_col=2, min_row=16, max_row=25))
pie.dataLabels = DataLabelList(); pie.dataLabels.showPercent = True
t.add_chart(pie, "G4")

p = wb.create_sheet("Plazo fijo vs dólar")
base(p, {"A": 2, "B": 38, "C": 16, "D": 16, "E": 16, "F": 3, "G": 52})
title(p, "B2", "¿Pesos a tasa o dólar?", "Compará 3 alternativas en pesos contra comprar dólar MEP hoy.")
label(p, "B4", "Monto a invertir (ARS)", bold=True); inp(p, "C4", 1000000, ARS)
label(p, "B5", "Plazo (días)", bold=True); inp(p, "C5", 30, NUM)
label(p, "B6", "Dólar MEP hoy"); out(p, "C6", f"={MK}C5", ARS)
for c, h in zip("CDE", ["Plazo fijo", "Caución", "FCI money market"]):
    p[f"{c}8"] = h; p[f"{c}8"].font = Font(name=FONT, bold=True, color="FFFFFF"); p[f"{c}8"].fill = PatternFill("solid", fgColor=NAVY)
    p[f"{c}8"].alignment = Alignment(horizontal="center")
label(p, "B9", "TNA (tasa nominal anual)", bold=True)
for c, v in zip("CDE", [0.33, 0.30, 0.28]): inp(p, f"{c}9", v, PCT2)
label(p, "B10", "Costos / comisiones (% sobre el monto)")
for c, v in zip("CDE", [0, 0.001, 0]): inp(p, f"{c}10", v, PCT2)
rows_ = [(11, "TEA (tasa efectiva anual)", "=(1+{c}9*$C$5/365)^(365/$C$5)-1", PCT2, False),
         (12, "Interés ganado (ARS)", "=$C$4*{c}9*$C$5/365-$C$4*{c}10", ARS, False),
         (13, "Monto al final (ARS)", "=$C$4+{c}12", ARS, True),
         (14, "DÓLAR DE EQUILIBRIO", "=IFERROR($C$6*{c}13/$C$4,0)", ARS, "key"),
         (15, "Suba del MEP que empata", "=IFERROR({c}14/$C$6-1,0)", PCT2, False),
         (16, "Monto final en USD al MEP de hoy", "=IFERROR({c}13/$C$6,0)", USDS, False)]
for r, lab, fx, fmt, k in rows_:
    label(p, f"B{r}", lab, bold=bool(k))
    for c in "CDE":
        (key if k == "key" else out)(p, f"{c}{r}", "=" + fx[1:].replace("{c}", c), fmt) if k == "key" else out(p, f"{c}{r}", fx.replace("{c}", c), fmt, bold=bool(k))
note(p, "G14", "Si dentro del plazo el MEP queda POR DEBAJO de este valor, te convenía la alternativa en pesos. Si queda POR ENCIMA, te convenía el dólar.")
label(p, "B18", "Tu estimación del MEP al final del plazo"); inp(p, "C18", 1486, ARS)
label(p, "B19", "Con tu estimación conviene…", bold=True)
for c in "CDE":
    out(p, f"{c}19", f'=IF($C$18<{c}14,"PESOS","DÓLAR")', None, bold=True)
pass_fail(p, "C19:E19", "PESOS")
note(p, "B21", "Es un cálculo, no una recomendación. El rendimiento futuro del dólar y de las tasas es incierto.")
wb.active = 3
import os; os.makedirs(os.path.dirname(OUT), exist_ok=True); wb.save(OUT); print("saved")
