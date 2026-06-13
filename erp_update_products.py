"""
Actualiza 02_PRODUCTOS y 11_INVENTARIO del ERP con datos reales del CSV de WooCommerce.
"""
import csv, re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

CSV_PATH = "/root/.claude/uploads/45cbde52-e6db-589b-b4ea-d918a46228c4/1c1945bd-wcproductexport12620261781288607325.csv"
XLS_PATH = "/home/user/pitlane-fastf1-api/ERP_ElRincon3DLab.xlsx"

# ─── Estilos ──────────────────────────────────────────────
C_WHITE="FFFFFF"; C_LGRAY="F5F5F5"; C_BORDER="DDDDDD"
C_DARK="1A1A2E"; C_DGRAY="555555"; C_BLACK="0D0D0D"
C_ACCENT="E31837"; C_GOLD="FFD700"; C_TEAL="00BCD4"

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

def parse_price(s):
    if not s: return 0
    s = str(s).strip().replace(',', '.').replace(' ', '')
    try: return float(s)
    except: return 0

def slugify(s, max_len=8):
    s = s.upper()
    s = re.sub(r'[ÁÀÄÂ]','A', s); s = re.sub(r'[ÉÈËÊ]','E', s)
    s = re.sub(r'[ÍÌÏÎ]','I', s); s = re.sub(r'[ÓÒÖÔ]','O', s)
    s = re.sub(r'[ÚÙÜÛ]','U', s)
    s = re.sub(r'\s+','-', s.strip())
    s = re.sub(r'[^A-Z0-9-]','', s)
    return s[:max_len].rstrip('-')

# ─── Parsear CSV ──────────────────────────────────────────
with open(CSV_PATH, newline='', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    rows = list(reader)

# Normalizar key del ID (tiene BOM)
def get_id(r):
    for k in r:
        if 'ID' in k: return r[k]
    return ''

# Construir índice de padres por SKU e ID
parent_by_sku  = {}  # SKU → row
parent_by_id   = {}  # "id:XXXX" → row

for r in rows:
    if r['Tipo'] in ('simple', 'variable'):
        sku = r['SKU'].strip()
        pid = get_id(r).strip()
        if sku:  parent_by_sku[sku]  = r
        if pid:  parent_by_id[f"id:{pid}"] = r

# ─── Mapeo de categorías WC → ERP ─────────────────────────
def map_categoria(cats_str):
    cats_lower = cats_str.lower()
    # Prioridad alta: categorías específicas primero
    if 'lámpara' in cats_lower or 'lampara' in cats_lower or 'neón' in cats_lower or 'neon' in cats_lower:
        return "ILUMINACIÓN", "Lámparas & Neones"
    if 'matera' in cats_lower:
        return "TERMOS Y BEBIDAS", "Materas"
    if 'vaso' in cats_lower or 'térmico' in cats_lower or 'termico' in cats_lower:
        return "TERMOS Y BEBIDAS", "Vasos Térmicos"
    if 'taza' in cats_lower:
        return "TERMOS Y BEBIDAS", "Tazas"
    if 'jarra' in cats_lower:
        return "TERMOS Y BEBIDAS", "Jarras"
    if 'mate' in cats_lower and 'matera' not in cats_lower:
        return "TERMOS Y BEBIDAS", "Mates"
    if 'indumentaria' in cats_lower:
        return "INDUMENTARIA", "Hoodies & Ropa"
    if 'personal' in cats_lower:
        return "PERSONALIZADOS", "A medida"
    if 'aleron' in cats_lower or 'alerón' in cats_lower or 'réplica' in cats_lower or 'replica' in cats_lower:
        return "IMPRESIÓN 3D", "Alerones & Réplicas"
    if 'portallaves' in cats_lower:
        return "IMPRESIÓN 3D", "Portallaves & Garaje"
    if 'escritorio' in cats_lower:
        return "IMPRESIÓN 3D", "Escritorio & Organización"
    if 'decorac' in cats_lower or 'garaje' in cats_lower:
        return "IMPRESIÓN 3D", "Decoración"
    return "IMPRESIÓN 3D", "General"

# ─── Construir lista maestra de productos ─────────────────
# Cada entrada = una fila en el ERP
# (id_erp, sku, nombre, cat, subcat, tipo_wc, precio_normal, precio_rebajado, activo, atrib_nombre, atrib_valor, parent_sku)

products_erp = []
erp_id = 1

for r in rows:
    tipo = r['Tipo']
    sku  = r['SKU'].strip()
    nombre = r['Nombre'].strip()
    precio = parse_price(r['Precio normal'])
    precio_reb = parse_price(r['Precio rebajado'])
    cats = r['Categorías'].strip()
    pub  = r['Publicado'].strip()
    activo = 'SI' if pub == '1' else 'NO'
    atrib1n = r['Nombre del atributo 1'].strip()
    atrib1v = r['Valor(es) del atributo 1'].strip()
    superior = r['Superior'].strip()

    if tipo == 'simple':
        cat, subcat = map_categoria(cats)
        products_erp.append({
            'id': erp_id, 'sku': sku, 'nombre': nombre,
            'cat': cat, 'subcat': subcat, 'tipo': 'Simple',
            'precio': precio, 'precio_reb': precio_reb,
            'activo': activo, 'variante': '', 'parent': '',
            'parent_nombre': '', 'is_header': False,
        })
        erp_id += 1

    elif tipo == 'variable':
        # Guardar como header (no-price row)
        cat, subcat = map_categoria(cats)
        products_erp.append({
            'id': erp_id, 'sku': sku, 'nombre': nombre,
            'cat': cat, 'subcat': subcat, 'tipo': 'Variable (familia)',
            'precio': 0, 'precio_reb': 0,
            'activo': activo, 'variante': f"Variantes: {atrib1v}",
            'parent': '', 'parent_nombre': '', 'is_header': True,
        })
        erp_id += 1

    elif tipo == 'variation':
        # Buscar parent
        parent_row = None
        if superior:
            parent_row = parent_by_sku.get(superior) or parent_by_id.get(superior)

        if parent_row:
            parent_sku = parent_row['SKU'].strip()
            parent_nombre = parent_row['Nombre'].strip()
            cats_p = parent_row['Categorías'].strip()
            cat, subcat = map_categoria(cats_p)
        else:
            parent_sku = superior
            parent_nombre = superior
            cat, subcat = "IMPRESIÓN 3D", "General"

        # Generar SKU para la variante
        atrib_val = r['Nombre del atributo 1'].split('=')[-1].strip() if '=' in r['Nombre del atributo 1'] else atrib1v
        var_suffix = slugify(atrib_val, 10)
        var_sku = f"{parent_sku}-{var_suffix}" if parent_sku else f"VAR-{var_suffix}"

        # Nombre limpio de la variante
        if ' - ' in nombre:
            nombre_var = nombre  # ya tiene el formato "Producto - Variante"
        else:
            nombre_var = f"{parent_nombre} — {atrib_val}"

        # Si no tiene precio, usar el precio del hoodie estándar (1890)
        if precio == 0 and parent_sku == 'ER3D-HOODIE-001':
            precio = 1890

        products_erp.append({
            'id': erp_id, 'sku': var_sku, 'nombre': nombre_var,
            'cat': cat, 'subcat': subcat, 'tipo': 'Variante',
            'precio': precio, 'precio_reb': precio_reb,
            'activo': activo, 'variante': atrib_val,
            'parent': parent_sku, 'parent_nombre': parent_nombre,
            'is_header': False,
        })
        erp_id += 1

print(f"Total productos ERP: {len(products_erp)}")
print(f"  Headers (familias variable): {sum(1 for p in products_erp if p['is_header'])}")
print(f"  Simples: {sum(1 for p in products_erp if p['tipo']=='Simple')}")
print(f"  Variantes: {sum(1 for p in products_erp if p['tipo']=='Variante')}")

# ─── Cargar workbook ──────────────────────────────────────
wb = load_workbook(XLS_PATH)

# ═══════════════════════════════════════════════════════
# RECONSTRUIR 02_PRODUCTOS
# ═══════════════════════════════════════════════════════
# Eliminar la hoja vieja y crear nueva
if "02_PRODUCTOS" in wb.sheetnames:
    idx = wb.sheetnames.index("02_PRODUCTOS")
    del wb["02_PRODUCTOS"]
    ws2 = wb.create_sheet("02_PRODUCTOS", idx)
else:
    ws2 = wb.create_sheet("02_PRODUCTOS", 1)

ws2.sheet_properties.tabColor = "1565C0"
ws2.sheet_view.showGridLines = False

# Título
ws2.merge_cells("A1:T1")
c = ws2["A1"]
c.value = "📦  PRODUCTOS — BASE MAESTRA  |  Fuente: WooCommerce"
c.fill = fill("1565C0"); c.font = font(bold=True, color=C_WHITE, size=13)
c.alignment = align("center","center")
ws2.row_dimensions[1].height = 28

# Subtítulo
ws2.merge_cells("A2:T2")
ws2["A2"].value = "⚡ Costos de producción a completar manualmente · Precios cargados desde WooCommerce · Margen se calcula automático"
ws2["A2"].fill = fill(C_DARK); ws2["A2"].font = font(italic=True, color="FFF9C4", size=8)
ws2["A2"].alignment = align("center","center")
ws2.row_dimensions[2].height = 16

# Headers
headers_prod = [
    ("ID",6),("SKU",22),("Nombre",36),("Categoría",20),("Subcategoría",20),
    ("Tipo",16),("Variante",24),("Costo MP",11),("Costo Imp.",11),("Costo DTF",11),
    ("Costo Pack.",11),("Costo Envío",11),("⚙ Costo Total",13),("💰 Precio WEB",13),
    ("Precio Oferta",13),("💵 Ganancia",12),("📊 Margen %",10),
    ("Peso (g)",9),("Tiempo Imp.(h)",13),("Activo",8),
]

for c_idx, (h, w) in enumerate(headers_prod, 1):
    cell = ws2.cell(row=3, column=c_idx, value=h)
    cell.fill = fill("1565C0"); cell.font = font(bold=True, color=C_WHITE, size=8)
    cell.alignment = align("center","center"); cell.border = border_thin()
    ws2.column_dimensions[get_column_letter(c_idx)].width = w
ws2.row_dimensions[3].height = 22

# Colores por categoría
cat_bg = {
    "IMPRESIÓN 3D":    ("E3F2FD","1565C0"),
    "INDUMENTARIA":    ("FCE4EC","880E4F"),
    "TERMOS Y BEBIDAS":("E8F5E9","1B5E20"),
    "ILUMINACIÓN":     ("FFF8E1","E65100"),
    "PERSONALIZADOS":  ("F3E5F5","4A148C"),
}

# Color por tipo de fila
HEADER_BG  = "263238"  # gris oscuro = familia variable
SIMPLE_BG  = [("FFFFFF",""),("F8F9FA","")]  # alternado blanco/gris
VARIANT_BG = [("E3F2FD",""),("EEF2FF","")]  # alternado azul claro

simple_count = 0
variant_counts = {}  # parent_sku → count

cur_row = 4
for prod in products_erp:
    is_header = prod['is_header']
    tipo      = prod['tipo']

    if is_header:
        # Fila de familia variable → fondo oscuro
        ws2.merge_cells(f"A{cur_row}:T{cur_row}")
        c = ws2.cell(row=cur_row, column=1)
        c.value = f"▶  {prod['nombre'].upper()}  ·  SKU padre: {prod['sku']}  ·  {prod['cat']}  ·  {prod['variante']}"
        c.fill = fill(HEADER_BG)
        c.font = font(bold=True, color=C_GOLD, size=9)
        c.alignment = align("left","center")
        c.border = border_thin()
        ws2.row_dimensions[cur_row].height = 18
        cur_row += 1
        continue

    # Fila de datos
    r = cur_row
    _, cat_fg = cat_bg.get(prod['cat'], ("F5F5F5", C_DGRAY))

    if tipo == 'Simple':
        bg = "FFFFFF" if simple_count % 2 == 0 else "F8F9FA"
        simple_count += 1
    else:
        vcount = variant_counts.get(prod['parent'], 0)
        bg = "EEF5FF" if vcount % 2 == 0 else "F5F0FF"
        variant_counts[prod['parent']] = vcount + 1

    data = [
        prod['id'], prod['sku'], prod['nombre'],
        prod['cat'], prod['subcat'], prod['tipo'], prod['variante'],
        0, 0, 0, 0, 0,  # costos vacíos (cols 8-12)
        None,            # costo total (fórmula) col 13
        prod['precio'],  # precio web col 14
        prod['precio_reb'] if prod['precio_reb'] > 0 else None,  # precio oferta col 15
        None,            # ganancia (fórmula) col 16
        None,            # margen % (fórmula) col 17
        0, 0,            # peso, tiempo imp. cols 18-19
        prod['activo'],  # col 20
    ]

    for c_idx, val in enumerate(data, 1):
        cell = ws2.cell(row=r, column=c_idx, value=val)
        cell.border = border_thin()
        cell.font = font(size=8)

        if c_idx in [8,9,10,11,12]:  # costos
            cell.fill = fill("FFFDE7")
            cell.number_format = "#,##0"
            cell.alignment = align("right","center")
            cell.font = font(size=8, italic=True, color="9E9E9E")
        elif c_idx == 13:  # costo total
            cell.value = f"=H{r}+I{r}+J{r}+K{r}+L{r}"
            cell.fill = fill("FFF3E0")
            cell.number_format = "#,##0"
            cell.alignment = align("right","center")
            cell.font = font(bold=True, size=9, color="E65100")
        elif c_idx == 14:  # precio web
            cell.fill = fill("E8F5E9") if val and val > 0 else fill("FFEBEE")
            cell.number_format = "#,##0"
            cell.alignment = align("right","center")
            cell.font = font(bold=True, size=10, color="1B5E20" if val and val > 0 else "B71C1C")
        elif c_idx == 15:  # precio oferta
            cell.fill = fill("FFF9C4") if val else fill(bg)
            cell.number_format = "#,##0"
            cell.alignment = align("right","center")
            cell.font = font(size=9, color="F57F17") if val else font(size=8, color=C_DGRAY)
        elif c_idx == 16:  # ganancia
            cell.value = f"=IF(N{r}>0,N{r}-M{r},0)"
            cell.fill = fill("E8F5E9")
            cell.number_format = "#,##0"
            cell.alignment = align("right","center")
            cell.font = font(bold=True, size=9, color="2E7D32")
        elif c_idx == 17:  # margen %
            cell.value = f"=IF(N{r}>0,P{r}/N{r}*100,0)"
            cell.fill = fill("E8F5E9")
            cell.number_format = "0.0"
            cell.alignment = align("right","center")
        elif c_idx == 4:  # categoría
            _, fg = cat_bg.get(prod['cat'], ("F5F5F5", C_DGRAY))
            bg_cat = cat_bg.get(prod['cat'], ("F5F5F5","555555"))[0]
            cell.fill = fill(bg_cat); cell.font = font(bold=True, size=8, color=fg)
            cell.alignment = align("center","center")
        elif c_idx == 6:  # tipo
            type_colors = {"Simple":"C8E6C9","Variante":"BBDEFB","Variable (familia)":HEADER_BG}
            cell.fill = fill(type_colors.get(tipo, bg))
            cell.font = font(size=8, color="555555")
            cell.alignment = align("center","center")
        elif c_idx == 20:  # activo
            cell.fill = fill("C8E6C9") if val == "SI" else fill("FFCDD2")
            cell.font = font(bold=True, size=8,
                            color="1B5E20" if val=="SI" else "B71C1C")
            cell.alignment = align("center","center")
        else:
            cell.fill = fill(bg)
            cell.alignment = align("left" if c_idx in [2,3,5,7] else "center","center")

    ws2.row_dimensions[r].height = 16
    cur_row += 1

# Tabla Excel (desde fila 3 hasta la última con datos)
last_data_row = cur_row - 1
# Para la tabla, necesitamos solo las filas de datos (no los headers de sección que están merged)
# La tabla debe abarcar el rango completo pero los headers merged rompen la tabla
# → Usamos solo desde fila 3 hasta la última fila con SKU

add_table = True
try:
    tbl = Table(displayName="TBL_Productos", ref=f"A3:T{last_data_row}")
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=False)
    ws2.add_table(tbl)
except:
    print("Nota: tabla no se pudo agregar por filas merged de headers de sección")
    add_table = False

# Validaciones
dv_cat = DataValidation(type="list",
    formula1='"IMPRESIÓN 3D,INDUMENTARIA,TERMOS Y BEBIDAS,ILUMINACIÓN,PERSONALIZADOS"',
    allow_blank=True)
ws2.add_data_validation(dv_cat); dv_cat.add(f"D4:D{last_data_row}")

dv_act = DataValidation(type="list", formula1='"SI,NO"', allow_blank=False)
ws2.add_data_validation(dv_act); dv_act.add(f"T4:T{last_data_row}")

ws2.freeze_panes = "D4"

print(f"02_PRODUCTOS: {cur_row-4} filas de datos escritas (hasta fila {last_data_row})")

# ═══════════════════════════════════════════════════════
# RECONSTRUIR 11_INVENTARIO
# ═══════════════════════════════════════════════════════
if "11_INVENTARIO" in wb.sheetnames:
    idx = wb.sheetnames.index("11_INVENTARIO")
    del wb["11_INVENTARIO"]
    ws11 = wb.create_sheet("11_INVENTARIO", idx)
else:
    ws11 = wb.create_sheet("11_INVENTARIO")

ws11.sheet_properties.tabColor = "4A148C"
ws11.sheet_view.showGridLines = False

ws11.merge_cells("A1:L1")
c = ws11["A1"]
c.value = "📦  INVENTARIO — CONTROL DE STOCK  |  Actualizado desde WooCommerce"
c.fill = fill("4A148C"); c.font = font(bold=True, color=C_WHITE, size=13)
c.alignment = align("center","center"); ws11.row_dimensions[1].height = 28

headers_inv = [
    ("SKU",22),("Producto",36),("Categoría",20),("Subcategoría",18),
    ("Tipo",16),("Precio Venta",13),("Stock Actual",12),("Stock Mínimo",12),
    ("Stock Ideal",11),("Costo Unit.",12),("Valor Stock",12),("Alerta",14),
]
for c_idx, (h, w) in enumerate(headers_inv, 1):
    cell = ws11.cell(row=2, column=c_idx, value=h)
    cell.fill = fill("4A148C"); cell.font = font(bold=True, color=C_WHITE, size=9)
    cell.alignment = align("center","center"); cell.border = border_thin()
    ws11.column_dimensions[get_column_letter(c_idx)].width = w
ws11.row_dimensions[2].height = 22

# Solo productos con precio (excluir headers de familias)
stockable = [p for p in products_erp if not p['is_header'] and p['precio'] > 0]

alerta_colors = {"":"E8F5E9","STOCK BAJO":"FFF9C4","CRÍTICO":"FFCDD2","":"E8F5E9"}

inv_row = 3
for idx_s, prod in enumerate(stockable):
    bg = "F8F9FA" if idx_s % 2 == 0 else "FFFFFF"
    r = inv_row

    ws11.cell(row=r, column=1,  value=prod['sku']).fill = fill(bg)
    ws11.cell(row=r, column=2,  value=prod['nombre']).fill = fill(bg)
    ws11.cell(row=r, column=3,  value=prod['cat'])
    ws11.cell(row=r, column=4,  value=prod['subcat']).fill = fill(bg)
    ws11.cell(row=r, column=5,  value=prod['tipo']).fill = fill(bg)
    ws11.cell(row=r, column=6,  value=prod['precio'])
    ws11.cell(row=r, column=7,  value=0)   # Stock actual
    ws11.cell(row=r, column=8,  value=3)   # Stock mínimo por defecto
    ws11.cell(row=r, column=9,  value=10)  # Stock ideal por defecto
    ws11.cell(row=r, column=10, value=0)   # Costo unitario
    # Col K = Valor stock = G*J
    ws11.cell(row=r, column=11, value=f"=G{r}*J{r}")
    # Col L = Alerta automática
    ws11.cell(row=r, column=12, value=f'=IF(G{r}<=0,"CRÍTICO",IF(G{r}<H{r},"CRÍTICO",IF(G{r}=H{r},"STOCK BAJO","")))')

    # Estilos por columna
    for col in range(1, 13):
        cell = ws11.cell(row=r, column=col)
        cell.border = border_thin()
        cell.font = font(size=8)
        if col == 3:
            _, fg = cat_bg.get(prod['cat'], ("F5F5F5", C_DGRAY))
            bg_cat = cat_bg.get(prod['cat'], ("F5F5F5","555555"))[0]
            cell.fill = fill(bg_cat); cell.font = font(bold=True, size=8, color=fg)
            cell.alignment = align("center","center")
        elif col == 6:
            cell.fill = fill("E8F5E9")
            cell.number_format = "#,##0"
            cell.alignment = align("right","center")
            cell.font = font(bold=True, size=9, color="1B5E20")
        elif col in [7, 8, 9]:
            cell.fill = fill("FFFDE7")
            cell.number_format = "#,##0"
            cell.alignment = align("right","center")
            cell.font = font(italic=(col!=7), size=9, color="555555" if col != 7 else C_BLACK)
        elif col in [10, 11]:
            cell.fill = fill(bg)
            cell.number_format = "#,##0"
            cell.alignment = align("right","center")
        elif col == 12:
            # Alerta — color se aplica con formato condicional; por ahora fondo neutro
            cell.fill = fill("E8F5E9")
            cell.font = font(bold=True, size=8, color="1B5E20")
            cell.alignment = align("center","center")
        elif col in [1, 2, 4]:
            cell.fill = fill(bg)
            cell.alignment = align("left","center")
        else:
            cell.fill = fill(bg)
            cell.alignment = align("center","center")

    ws11.row_dimensions[r].height = 15
    inv_row += 1

# Tabla
last_inv = inv_row - 1
try:
    tbl_inv = Table(displayName="TBL_Inventario", ref=f"A2:L{last_inv}")
    tbl_inv.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=False)
    ws11.add_table(tbl_inv)
except Exception as e:
    print(f"Inventario tabla: {e}")

ws11.freeze_panes = "C3"

print(f"11_INVENTARIO: {inv_row-3} productos stockeables")

# ─── Guardar ──────────────────────────────────────────────
wb.save(XLS_PATH)
print(f"\n✅ Workbook actualizado: {XLS_PATH}")
print(f"   Productos ERP: {len(products_erp)} filas ({sum(1 for p in products_erp if p['is_header'])} familias + {sum(1 for p in products_erp if not p['is_header'])} SKUs)")
print(f"   Inventario: {inv_row-3} SKUs")
