"""
EL RINCÓN 3D LAB — ERP Excel Builder
Part 2: Sheets 06-12 (Ventas, Detalle, Gastos, Publicidad, Envíos, Inventario, PitBot)
"""
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

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

def style_title(ws, ref, text, bg, fg=C_WHITE, size=13):
    c = ws[ref]; c.value = text
    c.fill = fill(bg); c.font = font(bold=True, color=fg, size=size)
    c.alignment = align("center", "center")

def add_headers(ws, row, headers, bg):
    for c_idx, (h, w) in enumerate(headers, 1):
        cell = ws.cell(row=row, column=c_idx, value=h)
        cell.fill = fill(bg)
        cell.font = font(bold=True, color=C_WHITE, size=9)
        cell.alignment = align("center", "center")
        cell.border = border_thin()
        ws.column_dimensions[get_column_letter(c_idx)].width = w
    ws.row_dimensions[row].height = 22

def add_table(ws, ref, name, style="TableStyleMedium2"):
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True)
    ws.add_table(t)

def freeze(ws, cell): ws.freeze_panes = cell
def tab_color(ws, c): ws.sheet_properties.tabColor = c

def data_row(ws, r_idx, data, num_cols=None, money_cols=None, center_cols=None, date_cols=None):
    bg = C_ROW_ALT if r_idx % 2 == 0 else C_ROW_EVEN
    for c_idx, val in enumerate(data, 1):
        cell = ws.cell(row=r_idx, column=c_idx, value=val)
        cell.border = border_thin()
        cell.fill = fill(bg)
        cell.font = font(size=9)
        if date_cols and c_idx in date_cols:
            cell.number_format = "DD/MM/YYYY"
            cell.alignment = align("center", "center")
        elif money_cols and c_idx in money_cols:
            cell.number_format = "#,##0"
            cell.alignment = align("right", "center")
        elif center_cols and c_idx in center_cols:
            cell.alignment = align("center", "center")
        else:
            cell.alignment = align("left", "center")
    ws.row_dimensions[r_idx].height = 16

wb = load_workbook("/home/user/pitlane-fastf1-api/erp_temp_part1.xlsx")

# ═══════════════════════════════════════════════════════
# SHEET 06 — VENTAS
# ═══════════════════════════════════════════════════════
ws6 = wb.create_sheet("06_VENTAS")
tab_color(ws6, "0D47A1")
ws6.sheet_view.showGridLines = False
ws6.merge_cells("A1:P1")
style_title(ws6, "A1", "🛒  VENTAS — REGISTRO MAESTRO", "0D47A1")
ws6.row_dimensions[1].height = 28

h6 = [
    ("ID Venta",9),("Fecha",12),("ID Cliente",10),("Nombre Cliente",20),
    ("Canal",13),("Método Pago",14),("Estado Pago",13),("Estado Pedido",14),
    ("Estado Envío",13),("Subtotal",12),("Descuento",12),("Facturación",13),
    ("Costo Total",12),("Ganancia",12),("Margen %",10),("Notas",25),
]
add_headers(ws6, 2, h6, "0D47A1")

sample_ventas = [
    (1,"2025-05-01",1,"Martín García","Instagram","Mercado Pago","Pagado","Entregado","Entregado",450,0,450,230,220,"",""),
    (2,"2025-05-03",5,"Carlos López","WhatsApp","Transferencia","Pagado","En proceso","Pendiente",380,0,380,195,185,"",""),
    (3,"2025-05-05",1,"Martín García","WooCommerce","Mercado Pago","Pagado","Entregado","Entregado",2200,100,2100,1310,790,"","Descuento fidelidad"),
    (4,"2025-05-08",6,"Sofía Pérez","WhatsApp","Efectivo","Pagado","Entregado","Retiró","1800",0,1800,845,955,"","Lámpara LED F1"),
    (5,"2025-05-10",2,"Laura Rodríguez","Instagram","Mercado Pago","Pagado","Enviado","En tránsito",1100,0,1100,635,465,"","Remera F1"),
    (6,"2025-05-12",7,"Pablo González","WooCommerce","Mercado Pago","Pagado","Entregado","Entregado",90,0,90,31,59,"","Llavero F1"),
    (7,"2025-05-15",5,"Carlos López","Instagram","Transferencia","Pagado","En proceso","Pendiente",550,50,500,285,215,"","Mini Garaje F1"),
    (8,"2025-05-18",3,"Diego Fernández","Facebook","Mercado Pago","Pendiente","Pendiente","Pendiente",1400,0,1400,760,640,"","Mate F1"),
    (9,"2025-05-20",1,"Martín García","WooCommerce","Mercado Pago","Pagado","Entregado","Entregado",1200,0,1200,640,560,"","Gorra F1"),
    (10,"2025-05-22",6,"Sofía Pérez","WhatsApp","Transferencia","Pagado","Enviado","En tránsito",3200,0,3200,1410,1790,"","Neón LED F1"),
]

ep_colors = {"Pagado":"C8E6C9","Pendiente":"FFF9C4","Cancelado":"FFCDD2"}
ep_fonts = {"Pagado":"1B5E20","Pendiente":"F57F17","Cancelado":"B71C1C"}

for r_idx, rd in enumerate(sample_ventas, 3):
    bg = C_ROW_ALT if r_idx % 2 == 0 else C_ROW_EVEN
    for c_idx, val in enumerate(rd, 1):
        cell = ws6.cell(row=r_idx, column=c_idx, value=val if not isinstance(val, str) or val else val)
        cell.border = border_thin()
        cell.font = font(size=9)
        if c_idx == 2:
            cell.number_format = "DD/MM/YYYY"; cell.alignment = align("center","center"); cell.fill = fill(bg)
        elif c_idx in [10,11,12,13,14]:
            try:
                cell.value = float(str(val).replace(",",""))
            except:
                cell.value = 0
            cell.number_format = "#,##0"; cell.alignment = align("right","center"); cell.fill = fill(bg)
        elif c_idx == 15:
            cell.value = f"=IF(L{r_idx}>0,(L{r_idx}-M{r_idx})/L{r_idx}*100,0)"
            cell.number_format = '0.0'; cell.alignment = align("right","center"); cell.fill = fill(bg)
        elif c_idx == 7:
            cell.fill = fill(ep_colors.get(val, bg))
            cell.font = font(bold=True, size=9, color=ep_fonts.get(val, C_BLACK))
            cell.alignment = align("center","center")
        else:
            cell.fill = fill(bg); cell.alignment = align("left" if c_idx in [4,16] else "center","center")
    ws6.row_dimensions[r_idx].height = 16

for dv_data, col_ref in [
    ('"Instagram,WhatsApp,Facebook,WooCommerce,Mercado Libre,Evento,Tienda Física,Distribuidor,Otro"',"E3:E100000"),
    ('"Mercado Pago,Transferencia,Efectivo,Tarjeta,Otro"',"F3:F100000"),
    ('"Pagado,Pendiente,Parcial,Cancelado,Reembolsado"',"G3:G100000"),
    ('"Pendiente,En proceso,Listo,Enviado,Entregado,Cancelado"',"H3:H100000"),
    ('"Pendiente,En tránsito,Entregado,Retiró,Sin envío,Devuelto"',"I3:I100000"),
]:
    dv = DataValidation(type="list", formula1=dv_data, allow_blank=True)
    ws6.add_data_validation(dv); dv.add(col_ref)

add_table(ws6, f"A2:P{2+len(sample_ventas)}", "TBL_Ventas", "TableStyleMedium2")
freeze(ws6, "D3")
print("Sheet 06 done")

# ═══════════════════════════════════════════════════════
# SHEET 07 — DETALLE VENTAS
# ═══════════════════════════════════════════════════════
ws7 = wb.create_sheet("07_DETALLE_VENTAS")
tab_color(ws7, "1565C0")
ws7.sheet_view.showGridLines = False
ws7.merge_cells("A1:I1")
style_title(ws7, "A1", "📋  DETALLE DE VENTAS — LÍNEAS DE PEDIDO", "1565C0")
ws7.row_dimensions[1].height = 28

h7 = [("ID Detalle",10),("ID Venta",9),("SKU",14),("Producto",28),
      ("Cantidad",10),("Precio Unit.",12),("Subtotal",12),("Costo Unit.",12),("Ganancia",12)]
add_headers(ws7, 2, h7, "1565C0")

sample_det = [
    (1,1,"IMP3D-001","Alerón F1 Escala 1:18",1,450,450,230,220),
    (2,2,"IMP3D-002","Circuito Monaco Mini",1,380,380,195,185),
    (3,3,"INDUM-001","Hoodie F1",1,2200,2200,1310,890),
    (4,4,"ILUM-001","Lámpara LED F1",1,1800,1800,845,955),
    (5,5,"INDUM-002","Remera F1",1,1100,1100,635,465),
    (6,6,"IMP3D-004","Llavero F1",1,90,90,31,59),
    (7,7,"IMP3D-003","Mini Garaje F1",1,550,550,285,265),
    (8,8,"TERM-001","Mate F1 Personalizado",1,1400,1400,760,640),
    (9,9,"INDUM-003","Gorra F1",1,1200,1200,640,560),
    (10,10,"ILUM-002","Neón LED F1",1,3200,3200,1410,1790),
    (11,3,"IMP3D-001","Alerón F1 Escala 1:18",0,0,0,0,0),  # multi-line example placeholder
]

for r_idx, rd in enumerate(sample_det, 3):
    data_row(ws7, r_idx, rd, money_cols={6,7,8,9}, center_cols={1,2,5})
ws7.row_dimensions[r_idx].height = 16

# Subtotal formula
for r in range(3, 3+len(sample_det)):
    ws7.cell(row=r, column=7).value = f"=E{r}*F{r}"
    ws7.cell(row=r, column=9).value = f"=G{r}-(E{r}*H{r})"
    for col in [7, 9]:
        ws7.cell(row=r, column=col).number_format = "#,##0"
        ws7.cell(row=r, column=col).alignment = align("right", "center")

add_table(ws7, f"A2:I{2+len(sample_det)}", "TBL_DetalleVentas", "TableStyleMedium2")
freeze(ws7, "D3")
print("Sheet 07 done")

# ═══════════════════════════════════════════════════════
# SHEET 08 — GASTOS
# ═══════════════════════════════════════════════════════
ws8 = wb.create_sheet("08_GASTOS")
tab_color(ws8, "B71C1C")
ws8.sheet_view.showGridLines = False
ws8.merge_cells("A1:I1")
style_title(ws8, "A1", "💸  GASTOS — REGISTRO COMPLETO", "B71C1C")
ws8.row_dimensions[1].height = 28

h8 = [("ID",8),("Fecha",12),("Categoría",18),("Subcategoría",18),
      ("Proveedor",20),("Descripción",30),("Monto",12),("Método Pago",14),("Comprobante",15)]
add_headers(ws8, 2, h8, "B71C1C")

sample_gastos = [
    (1,"2025-05-01","Filamento","PLA","FilaPrint","Filamento PLA 1kg x3 rollos",1800,"Transferencia","FAC-001"),
    (2,"2025-05-02","Publicidad","Meta Ads","Meta","Campaña Mayo Instagram",2500,"Tarjeta","AUTO"),
    (3,"2025-05-03","Packaging","Cajas","PackUY","Cajas + papel kraft 100u",650,"Efectivo","FAC-002"),
    (4,"2025-05-05","Software","SaaS","Canva Pro","Suscripción mensual",350,"Tarjeta","REC-001"),
    (5,"2025-05-07","DTF","Impresión","DTF Uruguay","Sublimación hoodies x10",800,"Transferencia","FAC-003"),
    (6,"2025-05-10","Envíos","DAC","DAC","Envíos mayo semana 1",480,"Efectivo","REC-002"),
    (7,"2025-05-12","Equipamiento","Consumibles","3D Store","Nozzle 0.4mm x5 + rodamientos",420,"Transferencia","FAC-004"),
    (8,"2025-05-15","Servicios","Electricidad","UTE","Factura eléctrica mayo",1200,"Débito","FAC-005"),
    (9,"2025-05-18","Publicidad","Meta Ads","Meta","Campaña Mayo Facebook",1500,"Tarjeta","AUTO"),
    (10,"2025-05-20","Filamento","PETG","FilaPrint","Filamento PETG negro 1kg",650,"Transferencia","FAC-006"),
    (11,"2025-05-22","Software","SaaS","WooCommerce","Plugin WooCommerce",280,"Tarjeta","REC-003"),
    (12,"2025-05-25","Envíos","DAC","DAC","Envíos mayo semana 2",360,"Efectivo","REC-004"),
]

cat_gasto_colors = {
    "Publicidad":"E3F2FD","Software":"F3E5F5","Filamento":"E8F5E9",
    "DTF":"FFF8E1","Packaging":"FBE9E7","Envíos":"E0F7FA",
    "Equipamiento":"F9FBE7","Servicios":"FCE4EC","Impuestos":"FFEBEE","Otros":"F5F5F5"
}

for r_idx, rd in enumerate(sample_gastos, 3):
    bg = cat_gasto_colors.get(rd[2], C_ROW_EVEN)
    for c_idx, val in enumerate(rd, 1):
        cell = ws8.cell(row=r_idx, column=c_idx, value=val)
        cell.border = border_thin(); cell.font = font(size=9)
        if c_idx == 2:
            cell.number_format = "DD/MM/YYYY"; cell.alignment = align("center","center"); cell.fill = fill(bg)
        elif c_idx == 7:
            cell.number_format = "#,##0"; cell.alignment = align("right","center"); cell.fill = fill(bg)
        else:
            cell.fill = fill(bg)
            cell.alignment = align("left" if c_idx in [5,6,9] else "center","center")
    ws8.row_dimensions[r_idx].height = 16

for dv_data, col_ref in [
    ('"Publicidad,Software,Filamento,DTF,Packaging,Envíos,Equipamiento,Servicios,Impuestos,Combustible,Herramientas,Otros"',
     "C3:C100000"),
    ('"Mercado Pago,Transferencia,Efectivo,Tarjeta,Débito,Otro"', "H3:H100000"),
]:
    dv = DataValidation(type="list", formula1=dv_data, allow_blank=True)
    ws8.add_data_validation(dv); dv.add(col_ref)

add_table(ws8, f"A2:I{2+len(sample_gastos)}", "TBL_Gastos", "TableStyleMedium3")
freeze(ws8, "C3")
print("Sheet 08 done")

# ═══════════════════════════════════════════════════════
# SHEET 09 — PUBLICIDAD
# ═══════════════════════════════════════════════════════
ws9 = wb.create_sheet("09_PUBLICIDAD")
tab_color(ws9, "1565C0")
ws9.sheet_view.showGridLines = False
ws9.merge_cells("A1:N1")
style_title(ws9, "A1", "📣  PUBLICIDAD — META ADS & GOOGLE ADS", "1565C0")
ws9.row_dimensions[1].height = 28

h9 = [
    ("ID",7),("Fecha",12),("Campaña",25),("Objetivo",16),("Canal",12),
    ("Inversión",12),("Alcance",12),("Clicks",10),("CTR %",8),
    ("Consultas",10),("Leads",9),("Ventas",9),("Facturación",13),("ROAS",9),
]
add_headers(ws9, 2, h9, "1565C0")

sample_pub = [
    (1,"2025-05-01","Alerones F1 Mayo","Conversiones","Instagram",1200,8500,320,3.76,28,15,6,2700,2.25),
    (2,"2025-05-01","Hoodies F1 Mayo","Alcance","Facebook",800,12000,180,1.50,12,8,3,6600,8.25),
    (3,"2025-05-08","Lanzamiento Neones","Conversiones","Instagram",1500,10200,450,4.41,35,22,9,28800,19.20),
    (4,"2025-05-15","Regalos Día del Padre","Conversiones","Instagram",2000,15800,620,3.92,48,30,14,18200,9.10),
    (5,"2025-05-15","Regalos Día del Padre","Conversiones","Facebook",1000,9500,280,2.95,20,12,5,6500,6.50),
    (6,"2025-05-22","Mates F1 Flash","Tráfico","Instagram",500,4200,150,3.57,18,10,4,5600,11.20),
]

for r_idx, rd in enumerate(sample_pub, 3):
    bg = C_ROW_ALT if r_idx % 2 == 0 else C_ROW_EVEN
    for c_idx, val in enumerate(rd, 1):
        cell = ws9.cell(row=r_idx, column=c_idx, value=val)
        cell.border = border_thin(); cell.fill = fill(bg); cell.font = font(size=9)
        if c_idx == 2:
            cell.number_format = "DD/MM/YYYY"; cell.alignment = align("center","center")
        elif c_idx in [6, 13]:
            cell.number_format = "#,##0"; cell.alignment = align("right","center")
        elif c_idx in [7, 8, 10, 11, 12]:
            cell.number_format = "#,##0"; cell.alignment = align("right","center")
        elif c_idx == 9:
            cell.number_format = "0.00"; cell.alignment = align("right","center")
        elif c_idx == 14:
            # ROAS color coding
            cell.number_format = "0.00"; cell.alignment = align("right","center")
            if isinstance(val, (int,float)):
                if val >= 5: cell.fill = fill("C8E6C9")
                elif val >= 2: cell.fill = fill("FFF9C4")
                else: cell.fill = fill("FFCDD2")
        else:
            cell.alignment = align("left" if c_idx in [3,4] else "center","center")
    ws9.row_dimensions[r_idx].height = 16

# Auto-calc CTR and ROAS
for r in range(3, 3+len(sample_pub)):
    ws9.cell(row=r, column=9).value = f"=IF(G{r}>0,H{r}/G{r}*100,0)"
    ws9.cell(row=r, column=14).value = f"=IF(F{r}>0,M{r}/F{r},0)"
    for col in [9, 14]:
        ws9.cell(row=r, column=col).number_format = "0.00"
        ws9.cell(row=r, column=col).alignment = align("right","center")

for dv_data, col_ref in [
    ('"Conversiones,Alcance,Tráfico,Interacción,Retargeting,Leads,Branding"', "D3:D100000"),
    ('"Instagram,Facebook,Google,TikTok,Email,WhatsApp,Otro"', "E3:E100000"),
]:
    dv = DataValidation(type="list", formula1=dv_data, allow_blank=True)
    ws9.add_data_validation(dv); dv.add(col_ref)

add_table(ws9, f"A2:N{2+len(sample_pub)}", "TBL_Publicidad", "TableStyleMedium2")
freeze(ws9, "C3")
print("Sheet 09 done")

# ═══════════════════════════════════════════════════════
# SHEET 10 — ENVÍOS
# ═══════════════════════════════════════════════════════
ws10 = wb.create_sheet("10_ENVIOS")
tab_color(ws10, "00695C")
ws10.sheet_view.showGridLines = False
ws10.merge_cells("A1:K1")
style_title(ws10, "A1", "🚚  ENVÍOS — SEGUIMIENTO LOGÍSTICO", "00695C")
ws10.row_dimensions[1].height = 28

h10 = [("ID Envío",9),("ID Venta",9),("Fecha",12),("Cliente",20),
       ("Ciudad",14),("Departamento",14),("Transportista",14),
       ("Nro. Seguimiento",20),("Costo",10),("Estado",14),("Fecha Entrega",13)]
add_headers(ws10, 2, h10, "00695C")

sample_env = [
    (1,1,"2025-05-02","Martín García","Montevideo","Montevideo","DAC","DAC-20250502-001",150,"Entregado","2025-05-04"),
    (2,2,"2025-05-04","Carlos López","Rivera","Rivera","OCA","OCA-20250504-001",220,"En tránsito",""),
    (3,3,"2025-05-06","Martín García","Montevideo","Montevideo","DAC","DAC-20250506-002",150,"Entregado","2025-05-08"),
    (4,5,"2025-05-11","Laura Rodríguez","Salto","Salto","OCA","OCA-20250511-001",280,"En tránsito",""),
    (5,7,"2025-05-16","Carlos López","Rivera","Rivera","DAC","DAC-20250516-001",220,"Pendiente",""),
    (6,10,"2025-05-23","Sofía Pérez","Maldonado","Maldonado","OCA","OCA-20250523-001",260,"En tránsito",""),
]

env_colors = {"Entregado":"C8E6C9","En tránsito":"BBDEFB","Pendiente":"FFF9C4","Devuelto":"FFCDD2"}
env_fonts = {"Entregado":"1B5E20","En tránsito":"0D47A1","Pendiente":"F57F17","Devuelto":"B71C1C"}

for r_idx, rd in enumerate(sample_env, 3):
    bg = C_ROW_ALT if r_idx % 2 == 0 else C_ROW_EVEN
    for c_idx, val in enumerate(rd, 1):
        cell = ws10.cell(row=r_idx, column=c_idx, value=val)
        cell.border = border_thin(); cell.font = font(size=9)
        if c_idx in [3, 11]:
            cell.number_format = "DD/MM/YYYY"; cell.alignment = align("center","center"); cell.fill = fill(bg)
        elif c_idx == 9:
            cell.number_format = "#,##0"; cell.alignment = align("right","center"); cell.fill = fill(bg)
        elif c_idx == 10:
            cell.fill = fill(env_colors.get(val, bg))
            cell.font = font(bold=True, size=9, color=env_fonts.get(val, "000000"))
            cell.alignment = align("center","center")
        else:
            cell.fill = fill(bg); cell.alignment = align("left" if c_idx in [4,5,6,8] else "center","center")
    ws10.row_dimensions[r_idx].height = 16

dv_trans = DataValidation(type="list", formula1='"DAC,OCA,Correo,Retiro personal,Otro"', allow_blank=True)
ws10.add_data_validation(dv_trans); dv_trans.add("G3:G100000")
dv_est_env = DataValidation(type="list", formula1='"Pendiente,En tránsito,Entregado,Devuelto,Sin envío"', allow_blank=True)
ws10.add_data_validation(dv_est_env); dv_est_env.add("J3:J100000")

add_table(ws10, f"A2:K{2+len(sample_env)}", "TBL_Envios", "TableStyleMedium7")
freeze(ws10, "D3")
print("Sheet 10 done")

# ═══════════════════════════════════════════════════════
# SHEET 11 — INVENTARIO
# ═══════════════════════════════════════════════════════
ws11 = wb.create_sheet("11_INVENTARIO")
tab_color(ws11, "4A148C")
ws11.sheet_view.showGridLines = False
ws11.merge_cells("A1:K1")
style_title(ws11, "A1", "📦  INVENTARIO — CONTROL DE STOCK", "4A148C")
ws11.row_dimensions[1].height = 28

h11 = [("SKU",14),("Producto",30),("Categoría",18),("Proveedor",18),
       ("Stock Actual",12),("Stock Mínimo",12),("Stock Ideal",11),
       ("Costo Unit.",12),("Valor Stock",12),("Alerta",14),("Fecha Act.",13)]
add_headers(ws11, 2, h11, "4A148C")

sample_inv = [
    ("IMP3D-001","Alerón F1 Escala 1:18","IMPRESIÓN 3D","Filamento PLA",12,5,20,230,2760,"","2025-05-28"),
    ("IMP3D-002","Circuito Monaco Mini","IMPRESIÓN 3D","Filamento PLA",8,5,15,195,1560,"","2025-05-28"),
    ("IMP3D-003","Mini Garaje F1","IMPRESIÓN 3D","Filamento PLA",3,5,10,285,855,"STOCK BAJO","2025-05-28"),
    ("IMP3D-004","Llavero F1","IMPRESIÓN 3D","Filamento PLA",25,10,40,31,775,"","2025-05-28"),
    ("IMP3D-005","Imán F1","IMPRESIÓN 3D","Filamento PLA",30,10,50,23,690,"","2025-05-28"),
    ("IMP3D-006","Portallaves F1","IMPRESIÓN 3D","Filamento PLA",15,8,25,65,975,"","2025-05-28"),
    ("IMP3D-007","Calendario F1 2025","IMPRESIÓN 3D","Filamento PLA",4,5,15,140,560,"STOCK BAJO","2025-05-28"),
    ("INDUM-001","Hoodie F1","INDUMENTARIA","Proveedor Textil",6,5,15,1310,7860,"","2025-05-28"),
    ("INDUM-002","Remera F1","INDUMENTARIA","Proveedor Textil",10,5,20,635,6350,"","2025-05-28"),
    ("INDUM-003","Gorra F1","INDUMENTARIA","Proveedor Textil",8,5,15,640,5120,"","2025-05-28"),
    ("TERM-001","Mate F1 Personalizado","TERMOS Y BEBIDAS","Proveedor Mates",5,5,12,760,3800,"","2025-05-28"),
    ("TERM-002","Vaso Térmico F1","TERMOS Y BEBIDAS","Proveedor Mates",2,5,10,815,1630,"CRÍTICO","2025-05-28"),
    ("ILUM-001","Lámpara LED F1","ILUMINACIÓN","Proveedor LED",4,3,8,845,3380,"","2025-05-28"),
    ("ILUM-002","Neón LED F1","ILUMINACIÓN","Proveedor LED",1,2,5,1410,1410,"CRÍTICO","2025-05-28"),
]

alerta_colors = {"":"E8F5E9","STOCK BAJO":"FFF9C4","CRÍTICO":"FFCDD2","REPOSICIÓN":"FFE0B2"}
alerta_fonts = {"":"1B5E20","STOCK BAJO":"F57F17","CRÍTICO":"B71C1C","REPOSICIÓN":"E65100"}

for r_idx, rd in enumerate(sample_inv, 3):
    bg = C_ROW_ALT if r_idx % 2 == 0 else C_ROW_EVEN
    for c_idx, val in enumerate(rd, 1):
        cell = ws11.cell(row=r_idx, column=c_idx, value=val)
        cell.border = border_thin(); cell.font = font(size=9)
        if c_idx == 11:
            cell.number_format = "DD/MM/YYYY"; cell.alignment = align("center","center"); cell.fill = fill(bg)
        elif c_idx in [5,6,7]:
            cell.number_format = "#,##0"; cell.alignment = align("right","center"); cell.fill = fill(bg)
        elif c_idx == 8:
            cell.number_format = "#,##0"; cell.alignment = align("right","center"); cell.fill = fill(bg)
        elif c_idx == 9:
            cell.value = f"=E{r_idx}*H{r_idx}"
            cell.number_format = "#,##0"; cell.alignment = align("right","center"); cell.fill = fill(bg)
        elif c_idx == 10:
            # Auto-alert formula
            cell.value = f'=IF(E{r_idx}<=0,"CRÍTICO",IF(E{r_idx}<F{r_idx},"CRÍTICO",IF(E{r_idx}=F{r_idx},"STOCK BAJO","")))'
            alerta_val = rd[9]
            cell.fill = fill(alerta_colors.get(alerta_val, bg))
            cell.font = font(bold=True, size=9, color=alerta_fonts.get(alerta_val, "000000"))
            cell.alignment = align("center","center")
        else:
            cell.fill = fill(bg)
            cell.alignment = align("left" if c_idx in [2,3,4] else "center","center")
    ws11.row_dimensions[r_idx].height = 16

add_table(ws11, f"A2:K{2+len(sample_inv)}", "TBL_Inventario", "TableStyleMedium9")
freeze(ws11, "C3")
print("Sheet 11 done")

# ═══════════════════════════════════════════════════════
# SHEET 12 — CONVERSACIONES PITBOT
# ═══════════════════════════════════════════════════════
ws12 = wb.create_sheet("12_PITBOT")
tab_color(ws12, "E65100")
ws12.sheet_view.showGridLines = False
ws12.merge_cells("A1:M1")
style_title(ws12, "A1", "🤖  PITBOT — CONVERSACIONES & MÉTRICAS IA", "E65100")
ws12.row_dimensions[1].height = 28

h12 = [
    ("ID Conv.",9),("Fecha",12),("Hora",9),("Canal",12),("Teléfono/User",16),
    ("Pregunta (resumen)",35),("Categoría",16),("Respuesta Enviada",35),
    ("Tiempo Resp.(min)",14),("Resultado",16),("Venta ID",9),
    ("Satisfacción",12),("Notas",20),
]
add_headers(ws12, 2, h12, "E65100")

sample_bot = [
    (1,"2025-05-10","09:15","WhatsApp","099111111","¿Tienen alerón del Ferrari SF-24?","Producto","Sí! Tenemos varios modelos. ¿Escala 1:18 o 1:43?",2,"Venta",1,"5","Cliente VIP"),
    (2,"2025-05-10","10:30","Instagram","@usuario2","¿Cuánto tarda la entrega?","Soporte","Entregamos en 3-5 días hábiles por DAC o OCA.",1,"Resuelto",None,"4",""),
    (3,"2025-05-11","11:00","WhatsApp","098222222","¿Hacen diseños personalizados de equipo?","Personalizado","¡Por supuesto! Mandame el logo y te hacemos presupuesto.",3,"Derivado humano",None,"5","Empresa grande"),
    (4,"2025-05-12","14:20","Facebook","@usuario4","¿Tienen stock del circuito Interlagos?","Producto","Actualmente no tenemos stock. ¿Te aviso cuando llegue?",2,"Resuelto",None,"4",""),
    (5,"2025-05-13","16:45","WhatsApp","097333333","¿Cómo pago? ¿Aceptan tarjeta?","Precio","Aceptamos MP, transferencia, efectivo y tarjeta.",1,"Resuelto",None,"5",""),
    (6,"2025-05-14","09:00","Instagram","@usuario6","¿Cuánto sale el neón de Red Bull?","Precio","El neón Red Bull sale $3200. ¿Te mando foto?",2,"Venta",10,"5",""),
    (7,"2025-05-15","11:30","WhatsApp","096444444","¿Envían a Rivera?","Soporte","Sí! Enviamos a todo Uruguay por OCA o DAC.",1,"Resuelto",None,"5",""),
    (8,"2025-05-16","15:00","WhatsApp","095555555","Mi pedido no llegó","Soporte","Disculpá! Te paso el número de seguimiento ahora.",5,"Derivado humano",None,"3","Pedido demorado"),
]

res_colors = {"Venta":"C8E6C9","Resuelto":"BBDEFB","Derivado humano":"FFF9C4","Sin respuesta":"FFCDD2"}
res_fonts = {"Venta":"1B5E20","Resuelto":"0D47A1","Derivado humano":"F57F17","Sin respuesta":"B71C1C"}
cat_colors = {"Producto":"E3F2FD","Precio":"F3E5F5","Soporte":"FBE9E7","Personalizado":"E8F5E9","Estado pedido":"FFF9C4","Otro":"F5F5F5"}

for r_idx, rd in enumerate(sample_bot, 3):
    bg = C_ROW_ALT if r_idx % 2 == 0 else C_ROW_EVEN
    for c_idx, val in enumerate(rd, 1):
        cell = ws12.cell(row=r_idx, column=c_idx, value=val if val is not None else "")
        cell.border = border_thin(); cell.font = font(size=9)
        if c_idx == 2:
            cell.number_format = "DD/MM/YYYY"; cell.alignment = align("center","center"); cell.fill = fill(bg)
        elif c_idx == 7:
            cell.fill = fill(cat_colors.get(val, bg)); cell.alignment = align("center","center")
        elif c_idx == 10:
            cell.fill = fill(res_colors.get(val, bg))
            cell.font = font(bold=True, size=9, color=res_fonts.get(val, "000000"))
            cell.alignment = align("center","center")
        elif c_idx in [6, 8, 13]:
            cell.fill = fill(bg); cell.alignment = align("left","center", wrap=True)
            ws12.row_dimensions[r_idx].height = 28
        else:
            cell.fill = fill(bg); cell.alignment = align("center","center")

for dv_data, col_ref in [
    ('"WhatsApp,Instagram,Facebook,WooCommerce,Email,Otro"', "D3:D100000"),
    ('"Producto,Precio,Estado pedido,Personalizado,Soporte,Otro"', "G3:G100000"),
    ('"Resuelto,Venta,Derivado humano,Sin respuesta"', "J3:J100000"),
    ('"1,2,3,4,5"', "L3:L100000"),
]:
    dv = DataValidation(type="list", formula1=dv_data, allow_blank=True)
    ws12.add_data_validation(dv); dv.add(col_ref)

add_table(ws12, f"A2:M{2+len(sample_bot)}", "TBL_PitBot", "TableStyleMedium3")
freeze(ws12, "D3")
print("Sheet 12 done")

wb.save("/home/user/pitlane-fastf1-api/erp_temp_part2.xlsx")
print("PART 2 SAVED — sheets 06-12")
