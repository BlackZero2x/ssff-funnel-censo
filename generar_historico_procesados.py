import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings('ignore')

# ── PARAMETROS ────────────────────────────────────────────────────────
PERIODOS        = [2602, 2603, 2604]
LABELS          = {2602: 'Feb-26', 2603: 'Mar-26', 2604: 'Abr-26*'}
LINEAS_EMB      = ['EMBUTIDOS']
LINEAS_CON      = ['SEMIELABORADOS', 'PRECOCIDOS']   # ELABORADOS no existe en det por ruta
OUTPUT          = 'C:/proyectos/SSFF/historico_procesados_rutas.xlsx'

TABLAS_PATH     = 'C:/proyectos/SSFF/TABLAS_RUTAS.xlsx'
DET_PATH        = 'C:/proyectos/SSFF/export_data_ssff_rutas.csv'

# Factor proyeccion abril (23 dias transcurridos de 25)
DIAS_LAB_MES_ACTUAL    = 25
DIAS_LAB_TRANSCURRIDOS = 23

# ── CARGA ─────────────────────────────────────────────────────────────
print("Cargando datos...")
det = pd.read_csv(DET_PATH, low_memory=False)
det.columns = det.columns.str.strip()
det['mes']           = det['mes'].astype(int)
det['total_volumen'] = pd.to_numeric(det['total_volumen'], errors='coerce').fillna(0)

tablas = pd.read_excel(TABLAS_PATH, sheet_name='RUTA_ACTUAL')
rutas_oficiales = tablas['RUTA'].astype(str).str.strip().tolist()
for v in ['V001', 'V002']:
    if v not in rutas_oficiales:
        rutas_oficiales.append(v)

ruta_ffvv_map = {}
for r in rutas_oficiales:
    if r.startswith('KB'):   ruta_ffvv_map[r] = 'KB'
    elif r.startswith('M0'): ruta_ffvv_map[r] = 'M0'
    elif r.startswith('P0'): ruta_ffvv_map[r] = 'P0'
    elif r.startswith('V0') or r in ['V001','V002']: ruta_ffvv_map[r] = 'V0'
    else:                    ruta_ffvv_map[r] = 'F8'
FFVV_ORDER = ['F8', 'M0', 'KB', 'P0', 'V0']

# Factor de proyeccion para abril
real_abr_global_vol = det[(det['mes'] == 2604) & (det['categoria'] == 'PROCESADOS')]['total_volumen'].sum()
factor_proy = DIAS_LAB_MES_ACTUAL / DIAS_LAB_TRANSCURRIDOS
print(f"   Rutas: {len(rutas_oficiales)} | Factor proyeccion abril: {factor_proy:.4f}")

# ── PROCESAR: clasificar lineas ────────────────────────────────────────
proc = det[
    (det['mes'].isin(PERIODOS)) &
    (det['categoria'] == 'PROCESADOS')
].copy()

proc['GRUPO'] = np.where(
    proc['linea'].isin(LINEAS_EMB), 'EMBUTIDOS',
    np.where(proc['linea'].isin(LINEAS_CON), 'CONGELADOS', None)
)
proc = proc[proc['GRUPO'].notna()]

# Proyectar abril
mask_abr = proc['mes'] == 2604
proc.loc[mask_abr, 'total_volumen'] *= factor_proy

# ── AGREGAR: kg y cobertura por ruta x grupo x mes ────────────────────
# Kilos
kg_agg = (proc.groupby(['ccod_ruta', 'GRUPO', 'mes'])['total_volumen']
              .sum().reset_index())
kg_pivot = kg_agg.pivot_table(index='ccod_ruta', columns=['GRUPO','mes'],
                               values='total_volumen', aggfunc='sum', fill_value=0)
kg_pivot.columns = [f'{g}_{m}_kg' for g, m in kg_pivot.columns]

# Cobertura (clientes unicos)
cob_agg = (proc.groupby(['ccod_ruta', 'GRUPO', 'mes'])['ccod_cli']
               .nunique().reset_index())
cob_agg.columns = ['ccod_ruta', 'GRUPO', 'mes', 'clientes']
cob_pivot = cob_agg.pivot_table(index='ccod_ruta', columns=['GRUPO','mes'],
                                 values='clientes', aggfunc='sum', fill_value=0)
cob_pivot.columns = [f'{g}_{m}_cob' for g, m in cob_pivot.columns]

# Combinar y alinear a rutas oficiales
datos = pd.concat([kg_pivot, cob_pivot], axis=1).reindex(rutas_oficiales, fill_value=0)

# Asegurar que todas las columnas existen
for g in ['EMBUTIDOS', 'CONGELADOS']:
    for m in PERIODOS:
        for suf in ['kg', 'cob']:
            col = f'{g}_{m}_{suf}'
            if col not in datos.columns:
                datos[col] = 0

print("   Datos agregados OK")

# ── HELPERS EXCEL ─────────────────────────────────────────────────────
def fill(h):  return PatternFill('solid', fgColor=h)
def fnt(bold=False, color='FF000000', sz=10):
    return Font(name='Aptos Narrow', bold=bold, color=color, size=sz)
def aln(h='center'): return Alignment(horizontal=h, vertical='center', wrap_text=True)
def brd():
    s = Side(style='thin', color='FFCCCCCC')
    return Border(left=s, right=s, top=s, bottom=s)

C_HDR  = 'FF1F3864'
C_EMB  = 'FF2E75B6'   # azul medio — EMBUTIDOS
C_CON  = 'FF70AD47'   # verde     — CONGELADOS
C_GEN  = 'FFED7D31'   # naranja   — TOTAL
C_GRIS = 'FFF2F2F2'
C_BLC  = 'FFFFFFFF'
C_PESO = 'FFFFD966'   # amarillo  — proyectado

# ── CONSTRUIR EXCEL ───────────────────────────────────────────────────
print("Generando Excel...")
wb = Workbook()
wb.remove(wb.active)
ws = wb.create_sheet('HISTORICO_PROC')

# ── Estructura de columnas ────────────────────────────────────────────
# Fila 1: titulo
# Fila 2: grupos (FFVV | RUTA | --EMBUTIDOS-- x3 meses x2 metricas | --CONGELADOS-- | --TOTAL--)
# Fila 3: sub-cabeceras (mes x metrica)
# Fila 4: metrica (kg / cob)
# Fila 5+: datos

# Layout:
# Col A=1: FFVV
# Col B=2: RUTA
# EMBUTIDOS: cols 3..8  (3 meses x 2 metricas = 6 cols)
# CONGELADOS: cols 9..14
# TOTAL PROC: cols 15..20
# TOTAL RUTA kg: col 21, TOTAL RUTA cob: col 22

COL_FFVV = 1
COL_RUTA = 2
COL_EMB  = 3        # inicio EMBUTIDOS
COL_CON  = 9        # inicio CONGELADOS
COL_TOT  = 15       # inicio TOTAL PROCESADOS
COL_END  = 20       # ultima col de datos

# Anchos
ws.column_dimensions['A'].width = 5
ws.column_dimensions['B'].width = 8
for c in range(3, COL_END + 1):
    ws.column_dimensions[get_column_letter(c)].width = 10

# ── Fila 1: titulo
ws.merge_cells(f'A1:{get_column_letter(COL_END)}1')
ws['A1'] = 'HISTORICO REAL PROCESADOS — EMBUTIDOS y CONGELADOS — KG y COBERTURA POR RUTA (Ult. 3 meses)'
ws['A1'].font = fnt(bold=True, color='FFFFFFFF', sz=12)
ws['A1'].fill = fill(C_HDR)
ws['A1'].alignment = aln()
ws.row_dimensions[1].height = 26

# ── Fila 2: grupos
grupos = [
    (COL_EMB, COL_EMB+5, 'EMBUTIDOS', C_EMB),
    (COL_CON, COL_CON+5, 'CONGELADOS', C_CON),
    (COL_TOT, COL_TOT+5, 'TOTAL PROCESADOS', C_GEN),
]
for col_ini, col_fin, label, bg in grupos:
    ws.merge_cells(f'{get_column_letter(col_ini)}2:{get_column_letter(col_fin)}2')
    c = ws.cell(row=2, column=col_ini, value=label)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=10)
    c.fill = fill(bg); c.alignment = aln()
ws.cell(row=2, column=COL_FFVV, value='').fill = fill(C_HDR)
ws.cell(row=2, column=COL_RUTA, value='').fill = fill(C_HDR)
ws.row_dimensions[2].height = 20

# ── Fila 3: mes por grupo  (periodo label)
# ── Fila 4: kg / cob
for grp_col in [COL_EMB, COL_CON, COL_TOT]:
    for k, mes in enumerate(PERIODOS):
        col_kg  = grp_col + k * 2
        col_cob = grp_col + k * 2 + 1
        lbl = LABELS[mes]
        bg  = C_PESO if mes == 2604 else C_GRIS

        # Fila 3: periodo (merge kg+cob)
        ws.merge_cells(f'{get_column_letter(col_kg)}3:{get_column_letter(col_cob)}3')
        c3 = ws.cell(row=3, column=col_kg, value=lbl)
        c3.font = fnt(bold=True, sz=9); c3.fill = fill(bg); c3.alignment = aln(); c3.border = brd()

        # Fila 4: metrica
        for col, lbl4 in [(col_kg, 'kg'), (col_cob, 'cob')]:
            c4 = ws.cell(row=4, column=col, value=lbl4)
            c4.font = fnt(bold=True, sz=9); c4.fill = fill(bg); c4.alignment = aln(); c4.border = brd()

# Cabeceras FFVV y RUTA
for row in [3, 4]:
    for col, lbl in [(COL_FFVV, 'FFVV'), (COL_RUTA, 'RUTA')]:
        c = ws.cell(row=row, column=col, value=lbl if row == 3 else '')
        c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
        c.fill = fill(C_HDR); c.alignment = aln(); c.border = brd()
ws.row_dimensions[3].height = 22
ws.row_dimensions[4].height = 20

# ── Filas de datos ────────────────────────────────────────────────────
row = 5
subtotales_ffvv = {}   # para fila de subtotal por FFVV

for ffvv in FFVV_ORDER:
    rutas_ffvv = sorted([r for r in rutas_oficiales if ruta_ffvv_map[r] == ffvv])
    sub = {col: 0 for col in range(COL_EMB, COL_END + 1)}

    for ruta in rutas_ffvv:
        f = fill(C_GRIS if row % 2 == 0 else C_BLC)
        ws.cell(row=row, column=COL_FFVV, value=ffvv).fill = f
        c = ws.cell(row=row, column=COL_RUTA, value=ruta)
        c.fill = f; c.font = fnt(bold=True, sz=9)

        emb_kg_tot = emb_cob_tot = con_kg_tot = con_cob_tot = 0

        for k, mes in enumerate(PERIODOS):
            emb_kg  = int(round(datos.loc[ruta, f'EMBUTIDOS_{mes}_kg']))
            emb_cob = int(round(datos.loc[ruta, f'EMBUTIDOS_{mes}_cob']))
            con_kg  = int(round(datos.loc[ruta, f'CONGELADOS_{mes}_kg']))
            con_cob = int(round(datos.loc[ruta, f'CONGELADOS_{mes}_cob']))
            tot_kg  = emb_kg + con_kg
            tot_cob = emb_cob + con_cob   # max de los dos (cliente puede comprar ambas)

            bg = fill(C_PESO) if mes == 2604 else f

            for col, val in [
                (COL_EMB + k*2,     emb_kg),
                (COL_EMB + k*2 + 1, emb_cob),
                (COL_CON + k*2,     con_kg),
                (COL_CON + k*2 + 1, con_cob),
                (COL_TOT + k*2,     tot_kg),
                (COL_TOT + k*2 + 1, tot_cob),
            ]:
                c = ws.cell(row=row, column=col, value=val)
                c.number_format = '#,##0'; c.fill = bg
                c.font = fnt(sz=9); c.alignment = aln('right')
                sub[col] = sub.get(col, 0) + val

        row += 1

    # Fila subtotal FFVV
    ws.cell(row=row, column=COL_RUTA, value=f'SUBTOTAL {ffvv}').fill = fill(C_HDR)
    ws.cell(row=row, column=COL_RUTA).font = fnt(bold=True, color='FFFFFFFF', sz=9)
    ws.cell(row=row, column=COL_FFVV, value=ffvv).fill = fill(C_HDR)
    ws.cell(row=row, column=COL_FFVV).font = fnt(bold=True, color='FFFFFFFF', sz=9)
    for col in range(COL_EMB, COL_END + 1):
        c = ws.cell(row=row, column=col, value=sub.get(col, 0))
        c.number_format = '#,##0'; c.fill = fill(C_HDR)
        c.font = fnt(bold=True, color='FFFFFFFF', sz=9); c.alignment = aln('right')
    subtotales_ffvv[ffvv] = sub
    row += 1

# ── Fila TOTAL GENERAL ────────────────────────────────────────────────
ws.cell(row=row, column=COL_RUTA, value='TOTAL GENERAL').fill = fill(C_GEN)
ws.cell(row=row, column=COL_RUTA).font = fnt(bold=True, color='FFFFFFFF', sz=10)
ws.cell(row=row, column=COL_FFVV, value='').fill = fill(C_GEN)
for col in range(COL_EMB, COL_END + 1):
    total_col = sum(sub.get(col, 0) for sub in subtotales_ffvv.values())
    c = ws.cell(row=row, column=col, value=total_col)
    c.number_format = '#,##0'; c.fill = fill(C_GEN)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=10); c.alignment = aln('right')
row += 2

# ── Nota al pie ───────────────────────────────────────────────────────
ws.cell(row=row, column=COL_FFVV,
        value=f'* Abr-26 = proyectado lineal ({DIAS_LAB_TRANSCURRIDOS} dias laborales transcurridos de {DIAS_LAB_MES_ACTUAL}). '
              'CONGELADOS = SEMIELABORADOS + PRECOCIDOS. '
              'Cobertura = clientes unicos con al menos 1 compra en el mes.')
ws.cell(row=row, column=COL_FFVV).font = fnt(sz=8)
ws.merge_cells(f'{get_column_letter(COL_FFVV)}{row}:{get_column_letter(COL_END)}{row}')

ws.freeze_panes = 'C5'

wb.save(OUTPUT)
print(f"Guardado: {OUTPUT}")

# ── Verificacion rapida ───────────────────────────────────────────────
total_emb_kg = datos[[f'EMBUTIDOS_{m}_kg' for m in PERIODOS]].sum()
total_con_kg = datos[[f'CONGELADOS_{m}_kg' for m in PERIODOS]].sum()
print("\nTotales por mes:")
for m in PERIODOS:
    print(f"  {LABELS[m]}: EMB {total_emb_kg[f'EMBUTIDOS_{m}_kg']:,.0f} kg | "
          f"CON {total_con_kg[f'CONGELADOS_{m}_kg']:,.0f} kg")
print("Done.")
