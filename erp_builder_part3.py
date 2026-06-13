"""
EL RINCÓN 3D LAB — ERP Excel Builder
Part 3: Sheets 13-20 (Embudo, P&L, Dashboards)
"""
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, DataBarRule
from openpyxl.styles.differential import DifferentialStyle

C_WHITE="FFFFFF"; C_LGRAY="F5F5F5"; C_GRAY="CCCCCC"; C_DGRAY="555555"
C_BORDER="DDDDDD"; C_ROW_ALT="F0F4FF"; C_ROW_EVEN="FFFFFF"
C_BLACK="0D0D0D"; C_DARK="1A1A2E"; C_MID="16213E"
C_ACCENT="E31837"; C_GOLD="FFD700"; C_GREEN="00C851"; C_TEAL="00BCD4"

def fill(c): return PatternFill("solid", fgColor=c)
def font(bold=False, color=C_BLACK, size=10, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic, name="Segoe UI")
def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def border_thin():
    s = Side(style="thin", color=C_BORDER)
    return Border(left=s, right=s, top=s, bottom=s)
def border_med():
    s = Side(style="medium", color=C_DGRAY)
    return Border(left=s, right=s, top=s, bottom=s)
def tab_color(ws, c): ws.sheet_properties.tabColor = c

def kpi_block(ws, row, col, label, formula_or_val, bg, fg="FFFFFF", size=18, fmt="#,##0"):
    # Label cell
    lc = ws.cell(row=row, column=col, value=label)
    lc.fill = fill(C_DARK); lc.font = font(bold=False, color=C_GRAY, size=8)
    lc.alignment = align("center","bottom"); lc.border = border_thin()
    # Value cell
    vc = ws.cell(row=row+1, column=col, value=formula_or_val)
    vc.fill = fill(bg); vc.font = font(bold=True, color=fg, size=size)
    vc.alignment = align("center","center"); vc.border = border_med()
    vc.number_format = fmt
    ws.row_dimensions[row].height = 16
    ws.row_dimensions[row+1].height = 36

def section_header(ws, row, col1, col2, text, bg=C_DARK, fg=C_GOLD):
    ws.merge_cells(f"{get_column_letter(col1)}{row}:{get_column_letter(col2)}{row}")
    c = ws.cell(row=row, column=col1, value=text)
    c.fill = fill(bg); c.font = font(bold=True, color=fg, size=10)
    c.alignment = align("left","center"); c.border = border_thin()
    ws.row_dimensions[row].height = 22

wb = load_workbook("/home/user/pitlane-fastf1-api/erp_temp_part2.xlsx")

# ═══════════════════════════════════════════════════════
# SHEET 13 — EMBUDO COMERCIAL
# ═══════════════════════════════════════════════════════
ws13 = wb.create_sheet("13_EMBUDO")
tab_color(ws13, "37474F")
ws13.sheet_view.showGridLines = False
ws13.merge_cells("A1:F1")
c = ws13["A1"]
c.value = "🔽  EMBUDO COMERCIAL — CONVERSIONES AUTOMÁTICAS"
c.fill = fill("37474F"); c.font = font(bold=True, color=C_WHITE, size=13)
c.alignment = align("center","center")
ws13.row_dimensions[1].height = 28

# Embudo data section
embudo_data = [
    ("Etapa","Cantidad","% del Total","Conversión siguiente","Observación"),
]
ws13.merge_cells("A2:E2")
ws13["A2"].value = "MÉTRICAS DEL EMBUDO — datos calculados desde hojas 04, 05, 06"
ws13["A2"].fill = fill(C_MID); ws13["A2"].font = font(italic=True, color=C_GRAY, size=9)
ws13["A2"].alignment = align("center","center")

headers_emb = ["Etapa","Cantidad","% del Total","Conv. → Siguiente","Observación"]
for c_idx, h in enumerate(headers_emb, 1):
    cell = ws13.cell(row=3, column=c_idx, value=h)
    cell.fill = fill("37474F"); cell.font = font(bold=True, color=C_WHITE, size=9)
    cell.alignment = align("center","center"); cell.border = border_thin()

etapas = [
    ("🔵 CONSULTAS (Leads totales)",    "=COUNTA(TBL_Leads[ID Lead])",         "=B4/B4",           "=B5/B4",  "Todos los leads registrados"),
    ("🟡 CONTACTADOS",                   "=COUNTIF(TBL_Leads[Estado],\"Contactado\")+COUNTIF(TBL_Leads[Estado],\"Presupuestado\")+COUNTIF(TBL_Leads[Estado],\"En negociación\")+COUNTIF(TBL_Leads[Estado],\"Ganado\")", "=B5/B4", "=B6/B5", "Respondidos"),
    ("🟠 PRESUPUESTADOS",                "=COUNTA(TBL_Cotizaciones[ID Cotización])", "=B6/B4",       "=B7/B6",  "Cotizaciones enviadas"),
    ("🟢 VENTAS CERRADAS",               "=COUNTA(TBL_Ventas[ID Venta])",       "=B7/B4",           "=B8/B7",  "Ventas confirmadas"),
    ("⭐ CLIENTES RECURRENTES",          "=COUNTIF(TBL_Clientes[Segmento],\"VIP\")+COUNTIF(TBL_Clientes[Segmento],\"Frecuente\")", "=B8/B7", "—", "VIP + Frecuentes"),
]

for r_offset, (etapa, cant, pct, conv, obs) in enumerate(etapas):
    row = 4 + r_offset
    colors = ["1565C0","F57F17","E65100","2E7D32","6A1B9A"]
    bgs = ["E3F2FD","FFF9C4","FFE0B2","C8E6C9","F3E5F5"]
    ws13.cell(row=row, column=1, value=etapa).fill = fill(bgs[r_offset])
    ws13.cell(row=row, column=1).font = font(bold=True, size=10, color=colors[r_offset])
    ws13.cell(row=row, column=1).alignment = align("left","center")
    ws13.cell(row=row, column=1).border = border_thin()

    ws13.cell(row=row, column=2, value=cant).number_format = "#,##0"
    ws13.cell(row=row, column=2).alignment = align("center","center")
    ws13.cell(row=row, column=2).fill = fill(bgs[r_offset])
    ws13.cell(row=row, column=2).font = font(bold=True, size=14, color=colors[r_offset])
    ws13.cell(row=row, column=2).border = border_med()

    ws13.cell(row=row, column=3, value=pct).number_format = "0.0%"
    ws13.cell(row=row, column=3).alignment = align("center","center")
    ws13.cell(row=row, column=3).fill = fill(bgs[r_offset])
    ws13.cell(row=row, column=3).border = border_thin()

    ws13.cell(row=row, column=4, value=conv if r_offset < 4 else "—")
    if r_offset < 4:
        ws13.cell(row=row, column=4).number_format = "0.0%"
    ws13.cell(row=row, column=4).alignment = align("center","center")
    ws13.cell(row=row, column=4).fill = fill(bgs[r_offset])
    ws13.cell(row=row, column=4).border = border_thin()

    ws13.cell(row=row, column=5, value=obs).fill = fill(bgs[r_offset])
    ws13.cell(row=row, column=5).font = font(italic=True, size=9, color=C_DGRAY)
    ws13.cell(row=row, column=5).alignment = align("left","center")
    ws13.cell(row=row, column=5).border = border_thin()
    ws13.row_dimensions[row].height = 28

for col, w in zip("ABCDE", [35, 14, 14, 18, 30]):
    ws13.column_dimensions[col].width = w

# Summary KPIs row
row = 10
section_header(ws13, row, 1, 5, "▶  RESUMEN FINANCIERO DEL EMBUDO")
ws13.merge_cells(f"A{row+1}:E{row+1}")
ws13[f"A{row+1}"].value = "=\" Facturación total: $\"&TEXT(SUM(TBL_Ventas[Facturación]),\"#,##0\")&\" | Ganancia total: $\"&TEXT(SUM(TBL_Ventas[Ganancia]),\"#,##0\")&\" | Ticket prom.: $\"&TEXT(AVERAGE(TBL_Ventas[Facturación]),\"#,##0\")"
ws13[f"A{row+1}"].fill = fill(C_DARK)
ws13[f"A{row+1}"].font = font(bold=True, color=C_GOLD, size=11)
ws13[f"A{row+1}"].alignment = align("center","center")
ws13.row_dimensions[row+1].height = 30
print("Sheet 13 done")

# ═══════════════════════════════════════════════════════
# SHEET 14 — P&L MENSUAL
# ═══════════════════════════════════════════════════════
ws14 = wb.create_sheet("14_PNL_MENSUAL")
tab_color(ws14, "1A237E")
ws14.sheet_view.showGridLines = False
ws14.merge_cells("A1:H1")
c = ws14["A1"]
c.value = "📊  ESTADO DE RESULTADOS — P&L MENSUAL"
c.fill = fill("1A237E"); c.font = font(bold=True, color=C_WHITE, size=13)
c.alignment = align("center","center"); ws14.row_dimensions[1].height = 28

# Month headers
meses = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic","TOTAL"]
ws14.cell(row=2, column=1, value="CONCEPTO").fill = fill("1A237E")
ws14.cell(row=2, column=1).font = font(bold=True, color=C_WHITE, size=9)
ws14.cell(row=2, column=1).alignment = align("center","center")
ws14.cell(row=2, column=1).border = border_thin()
ws14.column_dimensions["A"].width = 32

for m_idx, mes in enumerate(meses, 2):
    cell = ws14.cell(row=2, column=m_idx, value=mes)
    cell.fill = fill("283593" if mes != "TOTAL" else C_ACCENT)
    cell.font = font(bold=True, color=C_WHITE, size=9)
    cell.alignment = align("center","center"); cell.border = border_thin()
    ws14.column_dimensions[get_column_letter(m_idx)].width = 11
ws14.row_dimensions[2].height = 20

# P&L structure
pnl_rows = [
    # (label, bg, fg, bold, formula_type, indent)
    ("INGRESOS", "1A237E", C_GOLD, True, "section", 0),
    ("(+) Ventas brutas", "E8EAF6", C_DARK, False, "ventas", 1),
    ("(-) Descuentos", "E8EAF6", C_DARK, False, "descuentos", 1),
    ("(=) Facturación neta", "C5CAE9", C_DARK, True, "fact_neta", 0),
    ("", "", "", False, "blank", 0),
    ("COSTOS DE VENTAS", "1A237E", C_GOLD, True, "section", 0),
    ("(-) Costo de productos", "FFEBEE", "B71C1C", False, "costo_prod", 1),
    ("(-) Comisiones plataformas", "FFEBEE", "B71C1C", False, "comisiones", 1),
    ("(-) Envíos", "FFEBEE", "B71C1C", False, "envios", 1),
    ("(=) MARGEN BRUTO", "E8F5E9", "1B5E20", True, "margen_bruto", 0),
    ("MARGEN BRUTO %", "E8F5E9", "1B5E20", False, "margen_bruto_pct", 0),
    ("", "", "", False, "blank", 0),
    ("GASTOS OPERATIVOS", "1A237E", C_GOLD, True, "section", 0),
    ("(-) Publicidad Meta/Google", "FFF3E0", "E65100", False, "publicidad", 1),
    ("(-) Software / SaaS", "FFF3E0", "E65100", False, "software", 1),
    ("(-) Electricidad", "FFF3E0", "E65100", False, "electricidad", 1),
    ("(-) Filamento / Materiales", "FFF3E0", "E65100", False, "filamento", 1),
    ("(-) DTF / Sublimación", "FFF3E0", "E65100", False, "dtf", 1),
    ("(-) Packaging", "FFF3E0", "E65100", False, "packaging", 1),
    ("(-) Otros gastos", "FFF3E0", "E65100", False, "otros", 1),
    ("(=) TOTAL GASTOS OPERATIVOS", "FFCCBC", "BF360C", True, "total_gop", 0),
    ("", "", "", False, "blank", 0),
    ("RESULTADO", "1A237E", C_GOLD, True, "section", 0),
    ("(=) RESULTADO OPERATIVO", "E8F5E9", "1B5E20", True, "res_op", 0),
    ("MARGEN OPERATIVO %", "E8F5E9", "1B5E20", False, "margen_op_pct", 0),
    ("(-) Impuestos estimados (10%)", "FFEBEE", "B71C1C", False, "impuestos", 1),
    ("(=) UTILIDAD NETA", "C8E6C9", "1B5E20", True, "utilidad_neta", 0),
    ("MARGEN NETO %", "DCEDC8", "33691E", False, "margen_neto_pct", 0),
]

# Manual values for sample data (May only — rest blank)
pnl_values = {
    "ventas":      [0,0,0,0,10120,0,0,0,0,0,0,0],
    "descuentos":  [0,0,0,0,150, 0,0,0,0,0,0,0],
    "costo_prod":  [0,0,0,0,6821,0,0,0,0,0,0,0],
    "comisiones":  [0,0,0,0,353, 0,0,0,0,0,0,0],
    "envios":      [0,0,0,0,840, 0,0,0,0,0,0,0],
    "publicidad":  [0,0,0,0,4000,0,0,0,0,0,0,0],
    "software":    [0,0,0,0,630, 0,0,0,0,0,0,0],
    "electricidad":[0,0,0,0,1200,0,0,0,0,0,0,0],
    "filamento":   [0,0,0,0,2450,0,0,0,0,0,0,0],
    "dtf":         [0,0,0,0,800, 0,0,0,0,0,0,0],
    "packaging":   [0,0,0,0,650, 0,0,0,0,0,0,0],
    "otros":       [0,0,0,0,500, 0,0,0,0,0,0,0],
}

# Row mapping
row_map = {}
cur_row = 3
for item in pnl_rows:
    label, bg, fg, bold, ftype, indent = item
    if ftype == "blank":
        ws14.row_dimensions[cur_row].height = 8
        cur_row += 1; continue
    if ftype == "section":
        ws14.merge_cells(f"A{cur_row}:N{cur_row}")
        cell = ws14.cell(row=cur_row, column=1, value=label)
        cell.fill = fill(bg); cell.font = font(bold=True, color=fg, size=9)
        cell.alignment = align("left","center"); cell.border = border_thin()
        ws14.row_dimensions[cur_row].height = 20; cur_row += 1; continue

    row_map[ftype] = cur_row
    # Label
    cell = ws14.cell(row=cur_row, column=1, value=("  " * indent) + label)
    cell.fill = fill(bg); cell.font = font(bold=bold, color=fg, size=9)
    cell.alignment = align("left","center"); cell.border = border_thin()

    # Month values
    vals = pnl_values.get(ftype, [0]*12)
    for m_idx, val in enumerate(vals):
        col = m_idx + 2
        c = ws14.cell(row=cur_row, column=col)
        c.value = val if val != 0 else ""
        c.number_format = "#,##0"
        c.fill = fill(bg)
        c.font = font(bold=bold, color=fg, size=9)
        c.alignment = align("right","center"); c.border = border_thin()

    # Total column (N = col 14)
    tot_cell = ws14.cell(row=cur_row, column=14)
    tot_cell.value = f"=SUM(B{cur_row}:M{cur_row})"
    tot_cell.number_format = "#,##0"
    tot_cell.fill = fill("37474F" if bold else bg)
    tot_cell.font = font(bold=True, color=C_WHITE if bold else fg, size=9)
    tot_cell.alignment = align("right","center"); tot_cell.border = border_med()

    ws14.row_dimensions[cur_row].height = 18
    cur_row += 1

# Now add computed rows using row_map
def set_formula_row(row, formula_template, bg, fg, bold, fmt="#,##0"):
    for col in range(2, 14):
        col_l = get_column_letter(col)
        f = formula_template.replace("COL", col_l)
        c = ws14.cell(row=row, column=col, value=f)
        c.number_format = fmt; c.fill = fill(bg)
        c.font = font(bold=bold, color=fg, size=9)
        c.alignment = align("right","center"); c.border = border_thin()

if "fact_neta" in row_map and "ventas" in row_map and "descuentos" in row_map:
    r = row_map["fact_neta"]
    for col in range(2, 15):
        l = get_column_letter(col)
        rv = row_map["ventas"]; rd = row_map["descuentos"]
        ws14.cell(row=r, column=col).value = f"={l}{rv}-{l}{rd}"
        ws14.cell(row=r, column=col).number_format = "#,##0"
        ws14.cell(row=r, column=col).fill = fill("C5CAE9")
        ws14.cell(row=r, column=col).font = font(bold=True, color=C_DARK, size=9)
        ws14.cell(row=r, column=col).alignment = align("right","center")
        ws14.cell(row=r, column=col).border = border_thin()

if "margen_bruto" in row_map:
    r = row_map["margen_bruto"]
    rfn = row_map.get("fact_neta", 0)
    rcp = row_map.get("costo_prod", 0)
    rco = row_map.get("comisiones", 0)
    ren = row_map.get("envios", 0)
    for col in range(2, 15):
        l = get_column_letter(col)
        ws14.cell(row=r, column=col).value = f"={l}{rfn}-{l}{rcp}-{l}{rco}-{l}{ren}"
        ws14.cell(row=r, column=col).number_format = "#,##0"
        ws14.cell(row=r, column=col).fill = fill("E8F5E9")
        ws14.cell(row=r, column=col).font = font(bold=True, color="1B5E20", size=9)
        ws14.cell(row=r, column=col).alignment = align("right","center")
        ws14.cell(row=r, column=col).border = border_thin()

if "margen_bruto_pct" in row_map:
    r = row_map["margen_bruto_pct"]
    rfn = row_map.get("fact_neta", 0)
    rmb = row_map.get("margen_bruto", 0)
    for col in range(2, 15):
        l = get_column_letter(col)
        ws14.cell(row=r, column=col).value = f"=IF({l}{rfn}>0,{l}{rmb}/{l}{rfn}*100,0)"
        ws14.cell(row=r, column=col).number_format = '0.0"%"'
        ws14.cell(row=r, column=col).fill = fill("E8F5E9")
        ws14.cell(row=r, column=col).font = font(size=9, color="1B5E20")
        ws14.cell(row=r, column=col).alignment = align("right","center")
        ws14.cell(row=r, column=col).border = border_thin()

if "total_gop" in row_map:
    r = row_map["total_gop"]
    fields = ["publicidad","software","electricidad","filamento","dtf","packaging","otros"]
    for col in range(2, 15):
        l = get_column_letter(col)
        parts = "+".join([f"{l}{row_map[f]}" for f in fields if f in row_map])
        ws14.cell(row=r, column=col).value = f"={parts}"
        ws14.cell(row=r, column=col).number_format = "#,##0"
        ws14.cell(row=r, column=col).fill = fill("FFCCBC")
        ws14.cell(row=r, column=col).font = font(bold=True, color="BF360C", size=9)
        ws14.cell(row=r, column=col).alignment = align("right","center")
        ws14.cell(row=r, column=col).border = border_thin()

if "res_op" in row_map:
    r = row_map["res_op"]
    rmb = row_map.get("margen_bruto", 0)
    rtg = row_map.get("total_gop", 0)
    for col in range(2, 15):
        l = get_column_letter(col)
        ws14.cell(row=r, column=col).value = f"={l}{rmb}-{l}{rtg}"
        ws14.cell(row=r, column=col).number_format = "#,##0"
        ws14.cell(row=r, column=col).fill = fill("E8F5E9")
        ws14.cell(row=r, column=col).font = font(bold=True, color="1B5E20", size=10)
        ws14.cell(row=r, column=col).alignment = align("right","center")
        ws14.cell(row=r, column=col).border = border_med()

if "impuestos" in row_map and "res_op" in row_map:
    r = row_map["impuestos"]
    rro = row_map["res_op"]
    for col in range(2, 15):
        l = get_column_letter(col)
        ws14.cell(row=r, column=col).value = f"=IF({l}{rro}>0,{l}{rro}*0.10,0)"
        ws14.cell(row=r, column=col).number_format = "#,##0"
        ws14.cell(row=r, column=col).fill = fill("FFEBEE")
        ws14.cell(row=r, column=col).font = font(size=9, color="B71C1C")
        ws14.cell(row=r, column=col).alignment = align("right","center")
        ws14.cell(row=r, column=col).border = border_thin()

if "utilidad_neta" in row_map:
    r = row_map["utilidad_neta"]
    rro = row_map.get("res_op", 0)
    rim = row_map.get("impuestos", 0)
    for col in range(2, 15):
        l = get_column_letter(col)
        ws14.cell(row=r, column=col).value = f"={l}{rro}-{l}{rim}"
        ws14.cell(row=r, column=col).number_format = "#,##0"
        ws14.cell(row=r, column=col).fill = fill("C8E6C9")
        ws14.cell(row=r, column=col).font = font(bold=True, color="1B5E20", size=11)
        ws14.cell(row=r, column=col).alignment = align("right","center")
        ws14.cell(row=r, column=col).border = border_med()

if "margen_neto_pct" in row_map and "utilidad_neta" in row_map and "fact_neta" in row_map:
    r = row_map["margen_neto_pct"]
    run = row_map["utilidad_neta"]; rfn = row_map["fact_neta"]
    for col in range(2, 15):
        l = get_column_letter(col)
        ws14.cell(row=r, column=col).value = f"=IF({l}{rfn}>0,{l}{run}/{l}{rfn}*100,0)"
        ws14.cell(row=r, column=col).number_format = '0.0"%"'
        ws14.cell(row=r, column=col).fill = fill("DCEDC8")
        ws14.cell(row=r, column=col).font = font(bold=True, size=9, color="33691E")
        ws14.cell(row=r, column=col).alignment = align("right","center")
        ws14.cell(row=r, column=col).border = border_thin()

ws14.freeze_panes = "B3"
ws14.column_dimensions["N"].width = 13
print("Sheet 14 done")

# ═══════════════════════════════════════════════════════
# SHEET 15 — DASHBOARD EJECUTIVO
# ═══════════════════════════════════════════════════════
ws15 = wb.create_sheet("15_DASHBOARD_EJE")
tab_color(ws15, C_ACCENT)
ws15.sheet_view.showGridLines = False

ws15.merge_cells("A1:N1")
c = ws15["A1"]
c.value = "🏎️  EL RINCÓN 3D LAB — DASHBOARD EJECUTIVO"
c.fill = fill(C_ACCENT); c.font = font(bold=True, color=C_WHITE, size=16)
c.alignment = align("center","center"); ws15.row_dimensions[1].height = 36

ws15.merge_cells("A2:N2")
ws15["A2"].value = "=\" Período: \"&TEXT(TODAY(),\"MMMM YYYY\")&\" | Última actualización: \"&TEXT(NOW(),\"DD/MM/YYYY HH:MM\")"
ws15["A2"].fill = fill(C_DARK); ws15["A2"].font = font(color=C_GRAY, size=9, italic=True)
ws15["A2"].alignment = align("center","center"); ws15.row_dimensions[2].height = 16

# KPI Row 1 — Financial
section_header(ws15, 3, 1, 14, "💰  FINANCIERO")

kpi_data_r4 = [
    ("VENTAS MES", "=SUMPRODUCT((MONTH(TBL_Ventas[Fecha])=MONTH(TODAY()))*(YEAR(TBL_Ventas[Fecha])=YEAR(TODAY()))*TBL_Ventas[Facturación])", "1565C0"),
    ("VENTAS AÑO", "=SUMPRODUCT((YEAR(TBL_Ventas[Fecha])=YEAR(TODAY()))*TBL_Ventas[Facturación])", "0D47A1"),
    ("UTILIDAD NETA MES", "=SUMPRODUCT((MONTH(TBL_Ventas[Fecha])=MONTH(TODAY()))*(YEAR(TBL_Ventas[Fecha])=YEAR(TODAY()))*TBL_Ventas[Ganancia])", "2E7D32"),
    ("FACTURACIÓN TOTAL", "=SUM(TBL_Ventas[Facturación])", "37474F"),
    ("MARGEN PROMEDIO %", "=IF(SUM(TBL_Ventas[Facturación])>0,SUM(TBL_Ventas[Ganancia])/SUM(TBL_Ventas[Facturación])*100,0)", "4A148C"),
    ("TICKET PROMEDIO", "=IFERROR(AVERAGE(TBL_Ventas[Facturación]),0)", "006064"),
    ("TOTAL PEDIDOS", "=COUNTA(TBL_Ventas[ID Venta])", "BF360C"),
]

# Place KPIs in 2 rows per block, 7 cols
col_widths_dash = {
    "A":14,"B":14,"C":14,"D":14,"E":14,"F":14,"G":14,
    "H":14,"I":14,"J":14,"K":14,"L":14,"M":14,"N":14
}
for col, w in col_widths_dash.items():
    ws15.column_dimensions[col].width = w

row = 4
for col_idx, (label, formula, bg) in enumerate(kpi_data_r4, 1):
    lc = ws15.cell(row=row, column=col_idx, value=label)
    lc.fill = fill(C_DARK); lc.font = font(size=8, color=C_GRAY)
    lc.alignment = align("center","center"); lc.border = border_thin()
    ws15.row_dimensions[row].height = 16

    vc = ws15.cell(row=row+1, column=col_idx, value=formula)
    vc.fill = fill(bg); vc.font = font(bold=True, color=C_WHITE, size=16)
    vc.alignment = align("center","center"); vc.border = border_med()
    fmt = "0.0" if "%" in label else "#,##0"
    vc.number_format = fmt
    ws15.row_dimensions[row+1].height = 38

# KPI Row 2 — Commercial
section_header(ws15, 7, 1, 14, "🎯  COMERCIAL & MARKETING")
kpi_data_r8 = [
    ("CLIENTES NUEVOS MES", "=SUMPRODUCT((MONTH(TBL_Clientes[1ra Compra])=MONTH(TODAY()))*(YEAR(TBL_Clientes[1ra Compra])=YEAR(TODAY())))", "1B5E20"),
    ("CLIENTES VIP", "=COUNTIF(TBL_Clientes[Segmento],\"VIP\")", C_ACCENT),
    ("CLIENTES RECURRENTES", "=COUNTIF(TBL_Clientes[Segmento],\"Frecuente\")+COUNTIF(TBL_Clientes[Segmento],\"VIP\")", "F57F17"),
    ("LEADS MES", "=SUMPRODUCT((MONTH(TBL_Leads[Fecha])=MONTH(TODAY()))*(YEAR(TBL_Leads[Fecha])=YEAR(TODAY())))", "1565C0"),
    ("CONV. LEAD→VENTA %", "=IFERROR(COUNTA(TBL_Ventas[ID Venta])/COUNTA(TBL_Leads[ID Lead])*100,0)", "4A148C"),
    ("INVERSIÓN PUBLI MES", "=SUMPRODUCT((MONTH(TBL_Publicidad[Fecha])=MONTH(TODAY()))*(YEAR(TBL_Publicidad[Fecha])=YEAR(TODAY()))*TBL_Publicidad[Inversión])", "E65100"),
    ("ROAS PROMEDIO", "=IFERROR(SUM(TBL_Publicidad[Facturación])/SUM(TBL_Publicidad[Inversión]),0)", C_GOLD),
]
row = 8
for col_idx, (label, formula, bg) in enumerate(kpi_data_r8, 1):
    lc = ws15.cell(row=row, column=col_idx, value=label)
    lc.fill = fill(C_DARK); lc.font = font(size=8, color=C_GRAY)
    lc.alignment = align("center","center"); lc.border = border_thin()
    ws15.row_dimensions[row].height = 16

    vc = ws15.cell(row=row+1, column=col_idx, value=formula)
    vc.fill = fill(bg); vc.font = font(bold=True, color=C_WHITE if bg != C_GOLD else C_DARK, size=16)
    vc.alignment = align("center","center"); vc.border = border_med()
    fmt = "0.0" if "%" in label or "ROAS" in label else "#,##0"
    vc.number_format = fmt
    ws15.row_dimensions[row+1].height = 38

# KPI Row 3 — PitBot & Stock
section_header(ws15, 11, 1, 14, "🤖  PITBOT & OPERACIONES")
kpi_data_r12 = [
    ("CONVERSACIONES PITBOT", "=COUNTA(TBL_PitBot[ID Conv.])", "37474F"),
    ("VENTAS POR PITBOT", "=COUNTIF(TBL_PitBot[Resultado],\"Venta\")", "2E7D32"),
    ("CONV. PITBOT %", "=IFERROR(COUNTIF(TBL_PitBot[Resultado],\"Venta\")/COUNTA(TBL_PitBot[ID Conv.])*100,0)", "1B5E20"),
    ("COTIZACIONES ENVIADAS", "=COUNTA(TBL_Cotizaciones[ID Cotización])", "0D47A1"),
    ("COTIZ. ACEPTADAS %", "=IFERROR(COUNTIF(TBL_Cotizaciones[Estado],\"Aceptada\")/COUNTA(TBL_Cotizaciones[ID Cotización])*100,0)", "1565C0"),
    ("PRODUCTOS EN ALERTA", "=COUNTIF(TBL_Inventario[Alerta],\"STOCK BAJO\")+COUNTIF(TBL_Inventario[Alerta],\"CRÍTICO\")", C_ACCENT),
    ("PRODUCTOS ACTIVOS", "=COUNTIF(TBL_Productos[Activo],\"SI\")", "2E7D32"),
]
row = 12
for col_idx, (label, formula, bg) in enumerate(kpi_data_r12, 1):
    lc = ws15.cell(row=row, column=col_idx, value=label)
    lc.fill = fill(C_DARK); lc.font = font(size=8, color=C_GRAY)
    lc.alignment = align("center","center"); lc.border = border_thin()
    ws15.row_dimensions[row].height = 16

    vc = ws15.cell(row=row+1, column=col_idx, value=formula)
    vc.fill = fill(bg); vc.font = font(bold=True, color=C_WHITE, size=16)
    vc.alignment = align("center","center"); vc.border = border_med()
    fmt = "0.0" if "%" in label else "#,##0"
    vc.number_format = fmt
    ws15.row_dimensions[row+1].height = 38

# Top Products / Channels table
section_header(ws15, 15, 1, 7, "🏆  TOP PRODUCTOS POR FACTURACIÓN")
section_header(ws15, 15, 8, 14, "📡  FACTURACIÓN POR CANAL")

top_prod_headers = ["Producto","Ventas","Facturación","Ganancia","Margen %"]
top_chan_headers = ["Canal","Pedidos","Facturación","Ganancia","Margen %"]

for c_idx, h in enumerate(top_prod_headers, 1):
    cell = ws15.cell(row=16, column=c_idx, value=h)
    cell.fill = fill(C_MID); cell.font = font(bold=True, color=C_WHITE, size=8)
    cell.alignment = align("center","center"); cell.border = border_thin()

for c_idx, h in enumerate(top_chan_headers, 8):
    cell = ws15.cell(row=16, column=c_idx, value=h)
    cell.fill = fill(C_MID); cell.font = font(bold=True, color=C_WHITE, size=8)
    cell.alignment = align("center","center"); cell.border = border_thin()

# Note about PIVOT — manual or PQ needed for full ranking
note_row = 17
ws15.merge_cells(f"A{note_row}:G{note_row+6}")
ws15[f"A{note_row}"].value = "⚡ Conectar con Tabla Dinámica o Power Query\nbasada en TBL_Ventas + TBL_DetalleVentas\npara ranking automático de productos.\n\nUsar: Insertar > Tabla Dinámica > desde TBL_DetalleVentas\nAgrupar por: Producto | Valores: SUMA Ganancia, SUMA Facturación"
ws15[f"A{note_row}"].fill = fill("E8EAF6")
ws15[f"A{note_row}"].font = font(size=9, italic=True, color="1A237E")
ws15[f"A{note_row}"].alignment = align("left","top", wrap=True)
ws15[f"A{note_row}"].border = border_thin()

canales = ["Instagram","WhatsApp","Facebook","WooCommerce","Mercado Libre","Evento","Total"]
for i, canal in enumerate(canales):
    r = note_row + i
    for col in range(8, 13):
        c = ws15.cell(row=r, column=col)
        c.fill = fill(C_ROW_ALT)
        c.border = border_thin()
        c.font = font(size=9)
    ws15.cell(row=r, column=8, value=canal).alignment = align("left","center")
    ws15.cell(row=r, column=9, value=f'=COUNTIF(TBL_Ventas[Canal],H{r})').number_format="#,##0"
    ws15.cell(row=r, column=9).alignment = align("right","center")
    ws15.cell(row=r, column=10, value=f'=IFERROR(SUMIF(TBL_Ventas[Canal],H{r},TBL_Ventas[Facturación]),0)').number_format="#,##0"
    ws15.cell(row=r, column=10).alignment = align("right","center")
    ws15.row_dimensions[r].height = 16

print("Sheet 15 done")

# ═══════════════════════════════════════════════════════
# SHEETS 16-20 — remaining dashboards (summary cards)
# ═══════════════════════════════════════════════════════
dash_sheets = [
    ("16_DASH_PRODUCTOS", "📦  DASHBOARD PRODUCTOS", "1565C0", "E3F2FD",
     [
        ("TOP PRODUCTOS POR FACTURACIÓN","Usar Tabla Dinámica → TBL_DetalleVentas\nCampos: Producto (filas), SUMA(Subtotal) (valores)\nFiltro: Canal, Fecha"),
        ("TOP PRODUCTOS POR GANANCIA","Agregar columna calculada Ganancia en TBL_DetalleVentas\nTabla Dinámica con SUMA(Ganancia) por Producto"),
        ("TOP PRODUCTOS POR MARGEN %","Campo calculado: Margen = Ganancia/Subtotal\nOrdenar descendente"),
        ("TOP PRODUCTOS POR UNIDADES","SUMA(Cantidad) por Producto"),
        ("PRODUCTOS SIN MOVIMIENTO","Cruzar TBL_Inventario con TBL_DetalleVentas\nProductos en inventario que no aparecen en ventas"),
     ]),
    ("17_DASH_CLIENTES", "👤  DASHBOARD CLIENTES — CRM", "2E7D32", "E8F5E9",
     [
        ("TOP 10 CLIENTES POR FACTURACIÓN","GRANDE.K(TBL_Clientes[Facturación total], 1..10)\nFórmula INDEX/MATCH para nombre"),
        ("TOP 10 CLIENTES POR GANANCIA","Ordenar TBL_Clientes por Ganancia generada desc"),
        ("DISTRIBUCIÓN POR SEGMENTO","COUNTIF por Segmento: VIP / Frecuente / Ocasional / Dormido\nGráfico Donut"),
        ("RFM — ANÁLISIS DE VALOR","Scores R, F, M calculados en TBL_Clientes\nHeatmap de segmentos RFM"),
        ("ANÁLISIS DE RETENCIÓN","Clientes que compraron este mes vs. mes anterior\nFórmula COUNTIFS con fechas"),
     ]),
    ("18_DASH_GEO", "🗺️  DASHBOARD GEOGRÁFICO", "00695C", "E0F7FA",
     [
        ("VENTAS POR DEPARTAMENTO","=SUMIF(TBL_Ventas[Canal via JOIN Ciudad],...)\nTabla Dinámica TBL_Clientes + TBL_Ventas"),
        ("TOP CIUDADES","SUMIF con ciudad del cliente"),
        ("MAPA DE CALOR UY","Exportar tabla a Looker Studio / Power BI\npara mapa coroplético de Uruguay"),
        ("COSTO LOGÍSTICO POR ZONA","SUMIF(TBL_Envios[Departamento],...,TBL_Envios[Costo])"),
        ("RENTABILIDAD POR REGIÓN","Facturación - Costo Envío por departamento"),
     ]),
    ("19_DASH_MARKETING", "📣  DASHBOARD MARKETING", "E65100", "FFF3E0",
     [
        ("INVERSIÓN vs RETORNO","Gráfico Doble Eje: Inversión (barras) + ROAS (línea)\nDatos desde TBL_Publicidad"),
        ("CPL — COSTO POR LEAD","=SUMIF(Canal,TBL_Publicidad[Inversión])/COUNTIF(TBL_Leads[Canal],Canal)"),
        ("CAC — COSTO ADQUISICIÓN CLIENTE","=SUM(TBL_Publicidad[Inversión])/COUNTA(TBL_Clientes[ID Cliente])"),
        ("ROAS POR CAMPAÑA","=TBL_Publicidad[Facturación atribuida]/TBL_Publicidad[Inversión]\nFormato condicional: rojo<2, amarillo<4, verde≥4"),
        ("CTR Y CONVERSIÓN EMBUDO ADS","Alcance→Click→Consulta→Lead→Venta\nMini embudo por campaña"),
     ]),
    ("20_DASH_PITBOT", "🤖  DASHBOARD PITBOT — IA METRICS", "E31837", "FFEBEE",
     [
        ("CONVERSACIONES POR CANAL","COUNTIF(TBL_PitBot[Canal],canal)\nGráfico barras horizontales"),
        ("TASA DE RESOLUCIÓN","COUNTIF(Resultado,\"Resuelto\")/COUNTA(Resultado)"),
        ("TASA DE CONVERSIÓN A VENTA","COUNTIF(Resultado,\"Venta\")/COUNTA(Resultado)"),
        ("TIEMPO MEDIO DE RESPUESTA","=AVERAGE(TBL_PitBot[Tiempo Resp.(min)])"),
        ("CATEGORÍAS MÁS CONSULTADAS","COUNTIF por Categoría — Gráfico Donut"),
        ("SATISFACCIÓN PROMEDIO","=AVERAGEIF(TBL_PitBot[Satisfacción],\">0\")"),
        ("FACTURACIÓN ATRIBUIDA PITBOT","SUMIF(TBL_Ventas[ID Venta], ventas con origen PitBot)"),
     ]),
]

for sheet_name, title, bg_title, bg_card, cards in dash_sheets:
    ws = wb.create_sheet(sheet_name)
    tab_color(ws, bg_title)
    ws.sheet_view.showGridLines = False

    max_col = len(cards) * 2 + 1
    merge_end = get_column_letter(min(max_col, 14))
    ws.merge_cells(f"A1:{merge_end}1")
    c = ws["A1"]
    c.value = title; c.fill = fill(bg_title)
    c.font = font(bold=True, color=C_WHITE, size=14)
    c.alignment = align("center","center"); ws.row_dimensions[1].height = 32

    ws.merge_cells(f"A2:{merge_end}2")
    ws["A2"].value = "Conectar con tablas maestras · Usar Tablas Dinámicas o Power Query · Exportable a Looker Studio / Power BI"
    ws["A2"].fill = fill(C_DARK); ws["A2"].font = font(italic=True, color=C_GRAY, size=9)
    ws["A2"].alignment = align("center","center"); ws.row_dimensions[2].height = 16

    # Card grid
    card_row = 4
    for card_idx, (card_title, card_body) in enumerate(cards):
        col = (card_idx % 3) * 4 + 1
        if card_idx > 0 and card_idx % 3 == 0:
            card_row += 12

        col_l = get_column_letter(col)
        col_end = get_column_letter(col + 3)
        ws.merge_cells(f"{col_l}{card_row}:{col_end}{card_row}")
        ws[f"{col_l}{card_row}"].value = card_title
        ws[f"{col_l}{card_row}"].fill = fill(bg_title)
        ws[f"{col_l}{card_row}"].font = font(bold=True, color=C_WHITE, size=9)
        ws[f"{col_l}{card_row}"].alignment = align("center","center")
        ws[f"{col_l}{card_row}"].border = border_thin()
        ws.row_dimensions[card_row].height = 20

        ws.merge_cells(f"{col_l}{card_row+1}:{col_end}{card_row+9}")
        body_cell = ws[f"{col_l}{card_row+1}"]
        body_cell.value = card_body
        body_cell.fill = fill(bg_card)
        body_cell.font = font(size=9, color="263238")
        body_cell.alignment = align("left","top", wrap=True)
        body_cell.border = border_thin()
        for r in range(card_row+1, card_row+10):
            ws.row_dimensions[r].height = 14

    for col_idx in range(1, 14):
        ws.column_dimensions[get_column_letter(col_idx)].width = 14

    print(f"Sheet {sheet_name} done")

wb.save("/home/user/pitlane-fastf1-api/ERP_ElRincon3DLab.xlsx")
print("\n✅ WORKBOOK COMPLETO GUARDADO: ERP_ElRincon3DLab.xlsx")
