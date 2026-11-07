"""Shared styling helpers for all store workbooks (Excel + Google Sheets compatible)."""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule

FONT = "Arial"
NAVY = "14203A"
ACCENT = "2446C7"
SOFT = "E6EBFB"
INPUT_FILL = "FFF7D6"   # soft yellow = cells you edit
INPUT_BLUE = "0000FF"
GREEN = "1F7A4D"
RED = "B42318"
GREY = "5A6478"
LINE = "DFE3EA"

thin = Side(style="thin", color=LINE)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
BOTTOM = Border(bottom=thin)

USD = '$#,##0;($#,##0);"-"'
USD2 = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
MULT = '0.00"x"'
NUM = '#,##0;(#,##0);"-"'
ARS = '"$"#,##0;("$"#,##0);"-"'
USDS = '"US$"#,##0;("US$"#,##0);"-"'


def base(ws, widths):
    ws.sheet_view.showGridLines = False
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def f(bold=False, size=10, color=NAVY, italic=False):
    return Font(name=FONT, bold=bold, size=size, color=color, italic=italic)


def title(ws, cell, text, sub=None):
    ws[cell] = text
    ws[cell].font = f(True, 18, NAVY)
    if sub:
        r = ws[cell].row + 1
        c = ws[cell].column_letter
        ws[f"{c}{r}"] = sub
        ws[f"{c}{r}"].font = f(False, 10, GREY, italic=True)


def band(ws, rng, text):
    """Section header band across a range like 'B5:F5'."""
    first = rng.split(":")[0]
    ws.merge_cells(rng)
    ws[first] = text
    ws[first].font = Font(name=FONT, bold=True, size=10, color="FFFFFF")
    ws[first].fill = PatternFill("solid", fgColor=NAVY)
    ws[first].alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[ws[first].row].height = 20


def label(ws, cell, text, bold=False, note=None, indent=1):
    ws[cell] = text
    ws[cell].font = f(bold)
    ws[cell].alignment = Alignment(indent=indent, vertical="center")


def unit(ws, cell, text):
    ws[cell] = text
    ws[cell].font = f(False, 9, GREY)
    ws[cell].alignment = Alignment(horizontal="center")


def inp(ws, cell, value, fmt=None):
    ws[cell] = value
    ws[cell].font = Font(name=FONT, size=10, color=INPUT_BLUE)
    ws[cell].fill = PatternFill("solid", fgColor=INPUT_FILL)
    ws[cell].border = BOX
    ws[cell].alignment = Alignment(horizontal="right")
    if fmt:
        ws[cell].number_format = fmt


def out(ws, cell, formula, fmt=None, bold=False, color=NAVY, fill=None):
    ws[cell] = formula
    ws[cell].font = Font(name=FONT, size=10, bold=bold, color=color)
    ws[cell].alignment = Alignment(horizontal="right")
    ws[cell].border = BOTTOM
    if fill:
        ws[cell].fill = PatternFill("solid", fgColor=fill)
    if fmt:
        ws[cell].number_format = fmt


def key(ws, cell, formula, fmt=None):
    """Highlighted key result."""
    out(ws, cell, formula, fmt, bold=True, color=NAVY, fill=SOFT)
    ws[cell].font = Font(name=FONT, size=12, bold=True, color=NAVY)


def note(ws, cell, text, size=9):
    ws[cell] = text
    ws[cell].font = f(False, size, GREY, italic=True)
    ws[cell].alignment = Alignment(wrap_text=True, vertical="top")


def dropdown(ws, rng, options):
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(rng)


def pass_fail(ws, rng, good_text):
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'ISNUMBER(SEARCH("{good_text}",{rng.split(":")[0]}))'],
        font=Font(name=FONT, bold=True, color="FFFFFF"), fill=PatternFill("solid", fgColor=GREEN)))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'NOT(ISNUMBER(SEARCH("{good_text}",{rng.split(":")[0]})))'],
        font=Font(name=FONT, bold=True, color="FFFFFF"), fill=PatternFill("solid", fgColor=RED)))


def neg_red(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONT, color=RED, bold=True)))


def start_sheet(wb, product, tagline, steps, inside, disclaimer, lang="en"):
    ws = wb.active
    ws.title = "Start Here" if lang == "en" else "Inicio"
    base(ws, {"A": 3, "B": 30, "C": 80})
    title(ws, "B2", product, tagline)
    r = 5
    band(ws, f"B{r}:C{r}", "HOW TO USE (3 STEPS)" if lang == "en" else "CÓMO SE USA (3 PASOS)")
    for i, s in enumerate(steps, 1):
        r += 1
        ws[f"B{r}"] = f"Step {i}" if lang == "en" else f"Paso {i}"
        ws[f"B{r}"].font = f(True)
        ws[f"C{r}"] = s
        ws[f"C{r}"].font = f()
        ws[f"C{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30
    r += 2
    band(ws, f"B{r}:C{r}", "COLOR LEGEND" if lang == "en" else "REFERENCIA DE COLORES")
    r += 1
    inp(ws, f"B{r}", "1,000" if lang == "en" else "1.000")
    ws[f"C{r}"] = ("Yellow cell with blue text = an input you edit. Everything else calculates automatically."
                   if lang == "en" else "Celda amarilla con texto azul = dato que cargás vos. Todo lo demás se calcula solo.")
    ws[f"C{r}"].font = f()
    r += 1
    out(ws, f"B{r}", "1,000" if lang == "en" else "1.000")
    ws[f"C{r}"] = ("White cell = formula. Don't type over it (the example data shows the expected format)."
                   if lang == "en" else "Celda blanca = fórmula. No la sobrescribas (el ejemplo cargado muestra el formato esperado).")
    ws[f"C{r}"].font = f()
    r += 1
    key(ws, f"B{r}", "1,000" if lang == "en" else "1.000")
    ws[f"C{r}"] = "Blue cell = key result." if lang == "en" else "Celda azul = resultado clave."
    ws[f"C{r}"].font = f()
    r += 2
    band(ws, f"B{r}:C{r}", "WHAT'S INSIDE" if lang == "en" else "QUÉ INCLUYE")
    for tab, desc in inside:
        r += 1
        ws[f"B{r}"] = tab
        ws[f"B{r}"].font = f(True, color=ACCENT)
        ws[f"C{r}"] = desc
        ws[f"C{r}"].font = f()
        ws[f"C{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    r += 2
    band(ws, f"B{r}:C{r}", "WORKS IN" if lang == "en" else "FUNCIONA EN")
    r += 1
    ws[f"C{r}"] = ("Microsoft Excel (2010 or newer, Windows & Mac) and Google Sheets (File > Import). No macros."
                   if lang == "en" else "Microsoft Excel (2010 o más nuevo, Windows y Mac) y Google Sheets (Archivo > Importar). Sin macros.")
    ws[f"C{r}"].font = f()
    r += 2
    band(ws, f"B{r}:C{r}", "DISCLAIMER" if lang == "en" else "AVISO LEGAL")
    r += 1
    ws.merge_cells(f"B{r}:C{r+2}")
    note(ws, f"B{r}", disclaimer)
    ws.row_dimensions[r].height = 30
    return ws


DISCLAIMER_EN = ("Educational tool only. Not financial, lending, tax or legal advice. Results depend entirely on your inputs. "
                 "Loan programs and lender rules vary: always confirm terms with your lender and professionals before making decisions.")
DISCLAIMER_ES = ("Herramienta educativa. No constituye asesoramiento financiero, impositivo ni legal, ni recomendación de inversión "
                 "en los términos de la Ley 26.831. Los resultados dependen de los datos que cargues. Consultá con profesionales antes de decidir.")
