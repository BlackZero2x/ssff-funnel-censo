import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings('ignore')

# ════════════════════════════════════════════════════════════════════════
# PARAMETROS  ← ajustar cada mes
# ════════════════════════════════════════════════════════════════════════
PERIODOS_PRE        = [2601, 2602]   # Ene-26 y Feb-26 (baseline pre-reestructuracion)
PERIODOS_POST       = [2603, 2604]   # Mar-26 y Abr-26* (post-reestructuracion 08/03/2026)
LABELS              = {2601:'Ene-26', 2602:'Feb-26', 2603:'Mar-26', 2604:'Abr-26*'}
DIAS_LAB_MES_ACTUAL    = 25
DIAS_LAB_TRANSCURRIDOS = 23
FACTOR_PROY         = DIAS_LAB_MES_ACTUAL / DIAS_LAB_TRANSCURRIDOS

CATS_VALIDAS = ['ACCESORIOS','ANDINA','CERDO','COLGATE','DERMODIS','DULFINA',
                'HIGIENE Y CUIDADO','HOMEPRO PERU','HUEVO','KIMBERLY','LA CORONA',
                'LA PATRONA','MEDIFARMA','PAVO','POLLO','PROCESADOS','RINTI',
                'TAMBOS PERU','VERDUM','YICHANG']
SSFF_CATS    = ['CERDO','HUEVO','PAVO','POLLO','PROCESADOS']
FFVV_ORDER   = ['F8','M0','KB','P0','V0']

DET_PATH    = 'C:/proyectos/SSFF/export_data_ssff_rutas.csv'
MAEST_PATH  = 'C:/proyectos/SSFF/maestro_clientes_SSFF.xlsx'
TABLAS_PATH = 'C:/proyectos/SSFF/TABLAS_RUTAS.xlsx'
OUTPUT      = 'C:/proyectos/SSFF/seguimiento_censo_SSFF.xlsx'

# ════════════════════════════════════════════════════════════════════════
# HELPERS EXCEL
# ════════════════════════════════════════════════════════════════════════
def fill(h):  return PatternFill('solid', fgColor=h)
def fnt(bold=False, color='FF000000', sz=10):
    return Font(name='Aptos Narrow', bold=bold, color=color, size=sz)
def aln(h='center'): return Alignment(horizontal=h, vertical='center', wrap_text=True)
def brd():
    s = Side(style='thin', color='FFCCCCCC')
    return Border(left=s, right=s, top=s, bottom=s)

C_HDR1 = 'FF1F3864'   # azul oscuro — cabecera principal
C_HDR2 = 'FF2E74B5'   # azul medio — cabecera secundaria
C_POS  = 'FFE2EFDA'   # verde claro — variacion positiva
C_NEG  = 'FFFFC7CE'   # rojo claro — variacion negativa
C_PESO = 'FFFFD966'   # amarillo   — abril proyectado
C_CENSO= 'FFD6E4BC'   # verde suave — universo censo
C_SSFF = 'FFFCE4D6'   # naranja claro — categorias SSFF
C_SUB  = 'FFBDD7EE'   # azul claro — subtotales
C_GRIS = 'FFF2F2F2'
C_BLC  = 'FFFFFFFF'
C_ORG  = 'FFED7D31'   # naranja — total general

def _set(ws, row, col, val, bold=False, fg='FF000000', bg=None, num=None,
         halign='center', sz=10, merge_to=None):
    c = ws.cell(row=row, column=col, value=val)
    c.font   = fnt(bold=bold, color=fg, sz=sz)
    c.alignment = aln(halign)
    c.border = brd()
    if bg:   c.fill = fill(bg)
    if num:  c.number_format = num
    if merge_to:
        ws.merge_cells(f'{get_column_letter(col)}{row}:{get_column_letter(merge_to)}{row}')
    return c

def _color_var(ws, row, col, val):
    """Aplica color verde/rojo segun si val es positivo o negativo."""
    c = ws.cell(row=row, column=col)
    if isinstance(val, float) and not np.isnan(val):
        c.fill = fill(C_POS) if val >= 0 else fill(C_NEG)


# ════════════════════════════════════════════════════════════════════════
# [1] CARGA DE DATOS
# ════════════════════════════════════════════════════════════════════════
print("[1] Cargando datos...")

det = pd.read_csv(DET_PATH, low_memory=False)
det.columns = det.columns.str.strip()
det['mes']          = det['mes'].astype(int)
det['ccod_cli']     = pd.to_numeric(det['ccod_cli'], errors='coerce').fillna(0).astype(int)
det['ccod_vend']    = pd.to_numeric(det['ccod_vend'], errors='coerce').fillna(0).astype(int)
det['total_monto']  = pd.to_numeric(det['total_monto'],  errors='coerce').fillna(0)
det['total_volumen']= pd.to_numeric(det['total_volumen'], errors='coerce').fillna(0)
det['pedidos']      = pd.to_numeric(det['pedidos'],      errors='coerce').fillna(0)
det = det[det['categoria'].isin(CATS_VALIDAS)]

# Version proyectada de abril (solo montos, NUNCA cobertura)
det_proy = det.copy()
mask_abr = det_proy['mes'] == 2604
det_proy.loc[mask_abr, 'total_monto']   *= FACTOR_PROY
det_proy.loc[mask_abr, 'total_volumen'] *= FACTOR_PROY

# Maestro de clientes
lc_raw = pd.read_excel(MAEST_PATH, sheet_name='lista_clientes')
lc_raw['codigo'] = pd.to_numeric(lc_raw['codigo'], errors='coerce').fillna(0).astype(int)
lc_raw['es_censo'] = lc_raw['censo'].notna()
lc = lc_raw.drop_duplicates('codigo', keep='first')   # 976 duplicados → keep first

lc_censo = lc[lc['es_censo']][['codigo','ruta','codVend','supervisor','cluster','giro']].copy()
lc_censo['ruta'] = lc_censo['ruta'].astype(str).str.strip()

# Maestro de rutas
tablas = pd.read_excel(TABLAS_PATH, sheet_name='RUTA_ACTUAL')
tablas['RUTA'] = tablas['RUTA'].astype(str).str.strip()
rutas_oficiales = tablas['RUTA'].tolist()
for v in ['V001','V002']:
    if v not in rutas_oficiales:
        rutas_oficiales.append(v)

ruta_ffvv_map = {}
for r in rutas_oficiales:
    if r.startswith('KB'):                          ruta_ffvv_map[r] = 'KB'
    elif r.startswith('M0'):                        ruta_ffvv_map[r] = 'M0'
    elif r.startswith('P0'):                        ruta_ffvv_map[r] = 'P0'
    elif r.startswith('V0') or r in ['V001','V002']:ruta_ffvv_map[r] = 'V0'
    else:                                           ruta_ffvv_map[r] = 'F8'

# Tabla de vendedores: ccod_vend → nombre, supervisor, ruta
tablas['CODIGO_VEND'] = pd.to_numeric(tablas['CODIGO_VEND'], errors='coerce').fillna(0).astype(int)
vend_info = (tablas[['CODIGO_VEND','VENDEDOR','SUPERVISOR','RUTA']]
             .rename(columns={'CODIGO_VEND':'ccod_vend','VENDEDOR':'vendedor',
                               'SUPERVISOR':'supervisor_vend','RUTA':'ruta_actual'})
             .set_index('ccod_vend'))

# Supervisor de cada ruta oficial
ruta_sup_map = tablas.set_index('RUTA')['SUPERVISOR'].to_dict()

print(f"   CSV: {len(det):,} filas | Censo: {len(lc_censo):,} clientes | Rutas: {len(rutas_oficiales)}")


# ════════════════════════════════════════════════════════════════════════
# [2] ANALISIS 1 — COBERTURA Y VENTAS DE CLIENTES CENSO
# ════════════════════════════════════════════════════════════════════════
print("[2] Calculando cobertura censo...")

censo_ids = set(lc_censo['codigo'].astype(int))
censo_ruta_map = lc_censo.set_index('codigo')['ruta'].to_dict()
censo_sup_map  = lc_censo.set_index('codigo')['supervisor'].to_dict()

# Universo: cuantos clientes censo tiene asignados cada ruta
censo_por_ruta = (lc_censo.groupby('ruta')['codigo'].count()
                  .reindex(rutas_oficiales, fill_value=0))

# Filtrar transacciones post-reestructuracion de clientes censo
det_post_all   = det[det['mes'].isin(PERIODOS_POST) & det['ccod_cli'].isin(censo_ids)].copy()
det_proy_post  = det_proy[det_proy['mes'].isin(PERIODOS_POST) & det_proy['ccod_cli'].isin(censo_ids)].copy()

# Asignar ruta actual (de lista_clientes) en lugar de la ruta historica del CSV
det_post_all['ruta_actual']  = det_post_all['ccod_cli'].map(censo_ruta_map)
det_proy_post['ruta_actual'] = det_proy_post['ccod_cli'].map(censo_ruta_map)

# Cobertura: clientes unicos por ruta x mes (datos sin proyectar)
cob_ruta_mes = (det_post_all.groupby(['ruta_actual','mes'])['ccod_cli']
                .nunique().unstack(fill_value=0)
                .reindex(rutas_oficiales, fill_value=0))
for m in PERIODOS_POST:
    if m not in cob_ruta_mes.columns: cob_ruta_mes[m] = 0

# Soles: suma proyectada por ruta x mes
sol_ruta_mes = (det_proy_post.groupby(['ruta_actual','mes'])['total_monto']
                .sum().unstack(fill_value=0)
                .reindex(rutas_oficiales, fill_value=0))
for m in PERIODOS_POST:
    if m not in sol_ruta_mes.columns: sol_ruta_mes[m] = 0

# Por categoria (para bloque inferior de la hoja)
cob_cat_mes = (det_post_all.groupby(['categoria','mes'])['ccod_cli']
               .nunique().unstack(fill_value=0)
               .reindex(CATS_VALIDAS, fill_value=0))
sol_cat_mes = (det_proy_post.groupby(['categoria','mes'])['total_monto']
               .sum().unstack(fill_value=0)
               .reindex(CATS_VALIDAS, fill_value=0))
for m in PERIODOS_POST:
    if m not in cob_cat_mes.columns: cob_cat_mes[m] = 0
    if m not in sol_cat_mes.columns: sol_cat_mes[m] = 0


# ════════════════════════════════════════════════════════════════════════
# [3] ANALISIS 2 — VENDEDORES PRE vs POST REESTRUCTURACION
# ════════════════════════════════════════════════════════════════════════
print("[3] Calculando rendimiento vendedores...")

det_pre_v  = det[det['mes'].isin(PERIODOS_PRE)]
det_post_v = det_proy[det_proy['mes'].isin(PERIODOS_POST)]

pre_agg = det_pre_v.groupby('ccod_vend').agg(
    cli_pre   =('ccod_cli',    'nunique'),
    soles_pre =('total_monto', 'sum')
)
post_agg = det_post_v.groupby('ccod_vend').agg(
    cli_post   =('ccod_cli',    'nunique'),
    soles_post =('total_monto', 'sum')
)

# Censo aportado por vendedor (conteo real, sin proyectar)
det_post_raw_censo_v = det[det['mes'].isin(PERIODOS_POST) & det['ccod_cli'].isin(censo_ids)]
censo_vend = det_post_raw_censo_v.groupby('ccod_vend').agg(
    censo_cli_post   =('ccod_cli',    'nunique'),
    censo_soles_post =('total_monto', 'sum')
)

vendor_df = (vend_info
             .join(pre_agg,    how='left')
             .join(post_agg,   how='left')
             .join(censo_vend, how='left')
             .fillna(0))

n_pre  = len(PERIODOS_PRE)
n_post = len(PERIODOS_POST)
vendor_df['cli_pre_avg']    = vendor_df['cli_pre']   / n_pre
vendor_df['cli_post_avg']   = vendor_df['cli_post']  / n_post
vendor_df['soles_pre_avg']  = vendor_df['soles_pre'] / n_pre
vendor_df['soles_post_avg'] = vendor_df['soles_post']/ n_post

vendor_df['var_cli']   = (vendor_df['cli_post_avg']   /
                          vendor_df['cli_pre_avg'].replace(0, np.nan)   - 1)
vendor_df['var_soles'] = (vendor_df['soles_post_avg'] /
                          vendor_df['soles_pre_avg'].replace(0, np.nan) - 1)
vendor_df['censo_pct_cli'] = (vendor_df['censo_cli_post'] /
                               vendor_df['cli_post'].replace(0, np.nan))

vendor_df = vendor_df.reset_index().sort_values(['supervisor_vend','vendedor'])


# ════════════════════════════════════════════════════════════════════════
# [4] ANALISIS 3 — PENETRACION DE CATEGORIAS EN CENSO
# ════════════════════════════════════════════════════════════════════════
print("[4] Calculando penetracion de categorias...")

cat_activos = (det_post_all.groupby('categoria')['ccod_cli']
               .nunique().reindex(CATS_VALIDAS, fill_value=0))

cat_df = pd.DataFrame(index=CATS_VALIDAS)
cat_df['n_censo_total']    = len(lc_censo)
cat_df['cli_activos']      = cat_activos
cat_df['cli_mar26']        = cob_cat_mes.get(2603, pd.Series(0, index=CATS_VALIDAS))
cat_df['cli_abr26']        = cob_cat_mes.get(2604, pd.Series(0, index=CATS_VALIDAS))
cat_df['pct_penetracion']  = cat_df['cli_activos'] / cat_df['n_censo_total']
cat_df['soles_mar26']      = sol_cat_mes.get(2603, pd.Series(0.0, index=CATS_VALIDAS))
cat_df['soles_abr26']      = sol_cat_mes.get(2604, pd.Series(0.0, index=CATS_VALIDAS))
cat_df['soles_total']      = cat_df['soles_mar26'] + cat_df['soles_abr26']
cat_df = cat_df.sort_values('soles_total', ascending=False)


# ════════════════════════════════════════════════════════════════════════
# [5] ANALISIS 4 — RESUMEN POR SUPERVISOR
# ════════════════════════════════════════════════════════════════════════
print("[5] Calculando resumen supervisores...")

sup_df = (vendor_df.groupby('supervisor_vend').agg(
    n_vendedores     =('ccod_vend',       'count'),
    cli_pre          =('cli_pre',         'sum'),
    cli_post         =('cli_post',        'sum'),
    soles_pre        =('soles_pre',       'sum'),
    soles_post       =('soles_post',      'sum'),
    censo_cli_post   =('censo_cli_post',  'sum'),
    censo_soles_post =('censo_soles_post','sum'),
))
sup_df['cli_pre_avg']    = sup_df['cli_pre']   / n_pre
sup_df['cli_post_avg']   = sup_df['cli_post']  / n_post
sup_df['soles_pre_avg']  = sup_df['soles_pre'] / n_pre
sup_df['soles_post_avg'] = sup_df['soles_post']/ n_post
sup_df['var_cli']   = sup_df['cli_post_avg']   / sup_df['cli_pre_avg'].replace(0, np.nan)   - 1
sup_df['var_soles'] = sup_df['soles_post_avg'] / sup_df['soles_pre_avg'].replace(0, np.nan) - 1
sup_df = sup_df.sort_values('var_soles', ascending=False).reset_index()


# ════════════════════════════════════════════════════════════════════════
# [6] GENERAR EXCEL
# ════════════════════════════════════════════════════════════════════════
print("[6] Generando Excel...")
wb = Workbook()
wb.remove(wb.active)


# ── HOJA 1: COBERTURA_CENSO ───────────────────────────────────────────
ws1 = wb.create_sheet('COBERTURA_CENSO')

# Layout de columnas:
# A=FFVV B=RUTA C=SUPERVISOR D=N_CENSO
# E=Mar26-COB F=Mar26-COB% G=Mar26-SOLES
# H=Abr26-COB I=Abr26-COB% J=Abr26-SOLES
C_END1 = 10

ws1.column_dimensions['A'].width = 6
ws1.column_dimensions['B'].width = 8
ws1.column_dimensions['C'].width = 20
ws1.column_dimensions['D'].width = 9
for c in range(5, C_END1+1):
    ws1.column_dimensions[get_column_letter(c)].width = 12

# Fila 1: titulo
ws1.merge_cells(f'A1:{get_column_letter(C_END1)}1')
ws1['A1'] = 'SEGUIMIENTO COBERTURA Y VENTAS — CLIENTES CENSO POST REESTRUCTURACION (08/03/2026)'
ws1['A1'].font    = fnt(bold=True, color='FFFFFFFF', sz=12)
ws1['A1'].fill    = fill(C_HDR1)
ws1['A1'].alignment = aln()
ws1.row_dimensions[1].height = 26

# Fila 2: cabeceras de periodo
_set(ws1, 2, 1, 'FFVV',       bold=True, fg='FFFFFFFF', bg=C_HDR1, merge_to=1)
_set(ws1, 2, 2, 'RUTA',       bold=True, fg='FFFFFFFF', bg=C_HDR1)
_set(ws1, 2, 3, 'SUPERVISOR', bold=True, fg='FFFFFFFF', bg=C_HDR1)
_set(ws1, 2, 4, 'N_CENSO',    bold=True, fg='FFFFFFFF', bg=C_HDR1)
ws1.merge_cells('E2:G2')
ws1['E2'] = LABELS[2603]
ws1['E2'].font = fnt(bold=True, sz=10); ws1['E2'].fill = fill(C_HDR2)
ws1['E2'].alignment = aln(); ws1['E2'].border = brd()
ws1.merge_cells('H2:J2')
ws1['H2'] = LABELS[2604]
ws1['H2'].font = fnt(bold=True, sz=10); ws1['H2'].fill = fill(C_PESO)
ws1['H2'].alignment = aln(); ws1['H2'].border = brd()
ws1.row_dimensions[2].height = 20

# Fila 3: sub-cabeceras metricas
for col, lbl, bg in [
    (1,'',C_HDR1),(2,'',C_HDR1),(3,'',C_HDR1),(4,'',C_HDR1),
    (5,'COB',C_HDR2),(6,'COB %',C_HDR2),(7,'SOLES',C_HDR2),
    (8,'COB',C_PESO),(9,'COB %',C_PESO),(10,'SOLES',C_PESO),
]:
    c = ws1.cell(row=3, column=col, value=lbl)
    c.font = fnt(bold=True, sz=9, color='FFFFFFFF' if bg in (C_HDR1,C_HDR2) else 'FF000000')
    c.fill = fill(bg); c.alignment = aln(); c.border = brd()
ws1.row_dimensions[3].height = 18

# Datos por ruta
row = 4
subtotales = {}

for ffvv in FFVV_ORDER:
    rutas_ffvv = sorted([r for r in rutas_oficiales if ruta_ffvv_map.get(r) == ffvv])
    sub = {5:0, 6:0.0, 7:0.0, 8:0, 9:0.0, 10:0.0}
    sub_n = 0

    for ruta in rutas_ffvv:
        bg_row = C_GRIS if row % 2 == 0 else C_BLC
        n_cen = int(censo_por_ruta.get(ruta, 0))
        sup   = ruta_sup_map.get(ruta, '')
        cob3  = int(cob_ruta_mes.loc[ruta, 2603]) if 2603 in cob_ruta_mes.columns else 0
        cob4  = int(cob_ruta_mes.loc[ruta, 2604]) if 2604 in cob_ruta_mes.columns else 0
        sol3  = float(sol_ruta_mes.loc[ruta, 2603]) if 2603 in sol_ruta_mes.columns else 0.0
        sol4  = float(sol_ruta_mes.loc[ruta, 2604]) if 2604 in sol_ruta_mes.columns else 0.0
        pct3  = cob3 / n_cen if n_cen > 0 else 0.0
        pct4  = cob4 / n_cen if n_cen > 0 else 0.0

        _set(ws1, row, 1, ffvv, bg=bg_row, sz=9)
        _set(ws1, row, 2, ruta, bold=True, bg=bg_row, sz=9)
        _set(ws1, row, 3, sup,  bg=bg_row, sz=9, halign='left')
        _set(ws1, row, 4, n_cen, bg=C_CENSO, sz=9, num='#,##0')
        _set(ws1, row, 5, cob3, bg=bg_row, sz=9, num='#,##0')
        _set(ws1, row, 6, pct3, bg=bg_row, sz=9, num='0.0%')
        _set(ws1, row, 7, sol3, bg=bg_row, sz=9, num='#,##0')
        _set(ws1, row, 8, cob4, bg=C_PESO, sz=9, num='#,##0')
        _set(ws1, row, 9, pct4, bg=C_PESO, sz=9, num='0.0%')
        _set(ws1, row, 10, sol4, bg=C_PESO, sz=9, num='#,##0')

        sub[5] += cob3; sub[7] += sol3
        sub[8] += cob4; sub[10]+= sol4
        sub_n  += n_cen
        row += 1

    # Subtotal FFVV
    _set(ws1, row, 1, ffvv,           bold=True, fg='FFFFFFFF', bg=C_SUB, sz=9)
    _set(ws1, row, 2, f'SUBTOTAL {ffvv}', bold=True, fg='FFFFFFFF', bg=C_SUB, sz=9)
    _set(ws1, row, 3, '',             bg=C_SUB)
    _set(ws1, row, 4, sub_n,          bold=True, bg=C_SUB, sz=9, num='#,##0')
    _set(ws1, row, 5, sub[5],         bold=True, bg=C_SUB, sz=9, num='#,##0')
    _set(ws1, row, 6, sub[5]/sub_n if sub_n else 0, bold=True, bg=C_SUB, sz=9, num='0.0%')
    _set(ws1, row, 7, sub[7],         bold=True, bg=C_SUB, sz=9, num='#,##0')
    _set(ws1, row, 8, sub[8],         bold=True, bg=C_SUB, sz=9, num='#,##0')
    _set(ws1, row, 9, sub[8]/sub_n if sub_n else 0, bold=True, bg=C_SUB, sz=9, num='0.0%')
    _set(ws1, row, 10, sub[10],       bold=True, bg=C_SUB, sz=9, num='#,##0')
    subtotales[ffvv] = (sub, sub_n)
    row += 1

# Total general
tot_n   = sum(v[1] for v in subtotales.values())
tot_3c  = sum(v[0][5]  for v in subtotales.values())
tot_3s  = sum(v[0][7]  for v in subtotales.values())
tot_4c  = sum(v[0][8]  for v in subtotales.values())
tot_4s  = sum(v[0][10] for v in subtotales.values())
_set(ws1, row, 1, '',            bg=C_ORG)
_set(ws1, row, 2, 'TOTAL GENERAL', bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10)
_set(ws1, row, 3, '',            bg=C_ORG)
_set(ws1, row, 4, tot_n,         bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
_set(ws1, row, 5, tot_3c,        bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
_set(ws1, row, 6, tot_3c/tot_n if tot_n else 0, bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='0.0%')
_set(ws1, row, 7, tot_3s,        bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
_set(ws1, row, 8, tot_4c,        bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
_set(ws1, row, 9, tot_4c/tot_n if tot_n else 0, bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='0.0%')
_set(ws1, row, 10, tot_4s,       bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
row += 3

# ── Bloque por categoria ──────────────────────────────────────────────
ws1.merge_cells(f'A{row}:{get_column_letter(C_END1)}{row}')
c = ws1.cell(row=row, column=1, value='DETALLE POR CATEGORIA — CLIENTES CENSO')
c.font = fnt(bold=True, color='FFFFFFFF', sz=11); c.fill = fill(C_HDR1); c.alignment = aln()
row += 1

# cabeceras categoria
for col, lbl, bg in [
    (1,'CATEGORIA',C_HDR1),(2,LABELS[2603]+' COB',C_HDR2),(3,LABELS[2603]+' COB%',C_HDR2),
    (4,LABELS[2603]+' SOLES',C_HDR2),(5,LABELS[2604]+' COB',C_PESO),
    (6,LABELS[2604]+' COB%',C_PESO),(7,LABELS[2604]+' SOLES',C_PESO),
]:
    c = ws1.cell(row=row, column=col, value=lbl)
    c.font = fnt(bold=True, sz=9, color='FFFFFFFF' if bg in (C_HDR1,C_HDR2) else 'FF000000')
    c.fill = fill(bg); c.alignment = aln(); c.border = brd()
row += 1

n_censo_total = len(lc_censo)
for cat in CATS_VALIDAS:
    bg_row = C_SSFF if cat in SSFF_CATS else (C_GRIS if row % 2 == 0 else C_BLC)
    c3 = int(cob_cat_mes.loc[cat, 2603]) if 2603 in cob_cat_mes.columns else 0
    c4 = int(cob_cat_mes.loc[cat, 2604]) if 2604 in cob_cat_mes.columns else 0
    s3 = float(sol_cat_mes.loc[cat, 2603]) if 2603 in sol_cat_mes.columns else 0.0
    s4 = float(sol_cat_mes.loc[cat, 2604]) if 2604 in sol_cat_mes.columns else 0.0
    _set(ws1, row, 1, cat,  bg=bg_row, sz=9, halign='left')
    _set(ws1, row, 2, c3,   bg=bg_row, sz=9, num='#,##0')
    _set(ws1, row, 3, c3/n_censo_total, bg=bg_row, sz=9, num='0.0%')
    _set(ws1, row, 4, s3,   bg=bg_row, sz=9, num='#,##0')
    _set(ws1, row, 5, c4,   bg=C_PESO, sz=9, num='#,##0')
    _set(ws1, row, 6, c4/n_censo_total, bg=C_PESO, sz=9, num='0.0%')
    _set(ws1, row, 7, s4,   bg=C_PESO, sz=9, num='#,##0')
    row += 1

# Nota al pie
row += 1
nota = (f'N_CENSO = clientes censo asignados a la ruta segun lista_clientes (compren o no). '
        f'COB% = clientes activos / N_CENSO. '
        f'* Abr-26 = proyectado lineal ({DIAS_LAB_TRANSCURRIDOS} dias de {DIAS_LAB_MES_ACTUAL}, '
        f'factor {FACTOR_PROY:.4f}). Cobertura NO proyectada.')
ws1.merge_cells(f'A{row}:{get_column_letter(C_END1)}{row}')
ws1.cell(row=row, column=1, value=nota).font = fnt(sz=8)
ws1.freeze_panes = 'E4'


# ── HOJA 2: VENDEDORES_PRE_POST ───────────────────────────────────────
ws2 = wb.create_sheet('VENDEDORES_PRE_POST')

# Columnas: A=SUP B=RUTA C=VENDEDOR D=COD E=CLI_PRE F=CLI_POST G=VAR_CLI%
#           H=SOL_PRE I=SOL_POST J=VAR_SOL% K=CENSO_CLI L=CENSO_SOL M=CENSO_PCT
COLS2 = ['SUPERVISOR','RUTA_ACTUAL','VENDEDOR','COD_VEND',
         f'CLI_PRE\n({LABELS[2601]}+{LABELS[2602]})\nPROM/MES',
         f'CLI_POST\n({LABELS[2603]}+{LABELS[2604]}*)\nPROM/MES',
         'VAR\nCLI %',
         f'SOLES_PRE\nPROM/MES',
         f'SOLES_POST\nPROM/MES',
         'VAR\nSOLES %',
         'CENSO\nCLI_POST',
         'CENSO\nSOLES_POST',
         'CENSO %\nDE CLIENTES']
WIDTHS2 = [20,10,22,9,12,12,9,14,14,9,11,14,12]

ws2.column_dimensions['A'].width = WIDTHS2[0]
for i, w in enumerate(WIDTHS2[1:], 2):
    ws2.column_dimensions[get_column_letter(i)].width = w

# Titulo
ws2.merge_cells(f'A1:{get_column_letter(len(COLS2))}1')
ws2['A1'] = f'RENDIMIENTO VENDEDORES — PRE ({LABELS[2601]}+{LABELS[2602]}) vs POST REESTRUCTURACION ({LABELS[2603]}+{LABELS[2604]}*)'
ws2['A1'].font = fnt(bold=True, color='FFFFFFFF', sz=12)
ws2['A1'].fill = fill(C_HDR1); ws2['A1'].alignment = aln()
ws2.row_dimensions[1].height = 26

for col, lbl in enumerate(COLS2, 1):
    c = ws2.cell(row=2, column=col, value=lbl)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
    c.fill = fill(C_HDR1); c.alignment = aln(); c.border = brd()
ws2.row_dimensions[2].height = 44

row = 3
sups_vistos = {}

for _, vrow in vendor_df.iterrows():
    sup = vrow['supervisor_vend']
    bg_row = C_GRIS if row % 2 == 0 else C_BLC

    if sup not in sups_vistos:
        sups_vistos[sup] = {'start': row, 'rows': []}
    sups_vistos[sup]['rows'].append(row)

    var_cli   = vrow['var_cli']   if not np.isnan(vrow['var_cli'])   else None
    var_soles = vrow['var_soles'] if not np.isnan(vrow['var_soles']) else None

    _set(ws2, row, 1,  sup,                       bg=bg_row, sz=9, halign='left')
    _set(ws2, row, 2,  vrow['ruta_actual'],        bg=bg_row, sz=9)
    _set(ws2, row, 3,  vrow['vendedor'],           bg=bg_row, sz=9, halign='left')
    _set(ws2, row, 4,  int(vrow['ccod_vend']),     bg=bg_row, sz=9)
    _set(ws2, row, 5,  vrow['cli_pre_avg'],        bg=bg_row, sz=9, num='#,##0.0')
    _set(ws2, row, 6,  vrow['cli_post_avg'],       bg=bg_row, sz=9, num='#,##0.0')
    _set(ws2, row, 7,  var_cli   if var_cli   is not None else 'N/D', bg=bg_row, sz=9, num='0.0%')
    _set(ws2, row, 8,  vrow['soles_pre_avg'],      bg=bg_row, sz=9, num='#,##0')
    _set(ws2, row, 9,  vrow['soles_post_avg'],     bg=bg_row, sz=9, num='#,##0')
    _set(ws2, row, 10, var_soles if var_soles is not None else 'N/D', bg=bg_row, sz=9, num='0.0%')
    _set(ws2, row, 11, int(vrow['censo_cli_post']),   bg=bg_row, sz=9, num='#,##0')
    _set(ws2, row, 12, vrow['censo_soles_post'],      bg=bg_row, sz=9, num='#,##0')
    censo_pct = vrow['censo_pct_cli'] if not np.isnan(vrow['censo_pct_cli']) else 0
    _set(ws2, row, 13, censo_pct, bg=bg_row, sz=9, num='0.0%')

    if var_cli is not None:   _color_var(ws2, row, 7,  var_cli)
    if var_soles is not None: _color_var(ws2, row, 10, var_soles)

    row += 1

# Subtotales por supervisor
for sup, info in sups_vistos.items():
    sub_rows = info['rows']
    # Calcular subtotales del supervisor
    mask = vendor_df['supervisor_vend'] == sup
    vd_s = vendor_df[mask]
    s_cli_pre   = vd_s['cli_pre'].sum()  / n_pre
    s_cli_post  = vd_s['cli_post'].sum() / n_post
    s_sol_pre   = vd_s['soles_pre'].sum()  / n_pre
    s_sol_post  = vd_s['soles_post'].sum() / n_post
    s_var_cli   = s_cli_post / s_cli_pre   - 1 if s_cli_pre   > 0 else np.nan
    s_var_sol   = s_sol_post / s_sol_pre   - 1 if s_sol_pre   > 0 else np.nan
    s_cen_cli   = int(vd_s['censo_cli_post'].sum())
    s_cen_sol   = vd_s['censo_soles_post'].sum()

    _set(ws2, row, 1,  sup,               bold=True, bg=C_SUB, sz=9, halign='left')
    _set(ws2, row, 2,  f'SUB {sup}',      bold=True, bg=C_SUB, sz=9)
    _set(ws2, row, 3,  f'({len(sub_rows)} vendedores)', bg=C_SUB, sz=9)
    _set(ws2, row, 4,  '',                bg=C_SUB)
    _set(ws2, row, 5,  s_cli_pre,         bold=True, bg=C_SUB, sz=9, num='#,##0.0')
    _set(ws2, row, 6,  s_cli_post,        bold=True, bg=C_SUB, sz=9, num='#,##0.0')
    _set(ws2, row, 7,  s_var_cli if not np.isnan(s_var_cli) else 'N/D', bold=True, bg=C_SUB, sz=9, num='0.0%')
    _set(ws2, row, 8,  s_sol_pre,         bold=True, bg=C_SUB, sz=9, num='#,##0')
    _set(ws2, row, 9,  s_sol_post,        bold=True, bg=C_SUB, sz=9, num='#,##0')
    _set(ws2, row, 10, s_var_sol if not np.isnan(s_var_sol) else 'N/D', bold=True, bg=C_SUB, sz=9, num='0.0%')
    _set(ws2, row, 11, s_cen_cli,         bold=True, bg=C_SUB, sz=9, num='#,##0')
    _set(ws2, row, 12, s_cen_sol,         bold=True, bg=C_SUB, sz=9, num='#,##0')
    _set(ws2, row, 13, '',                bg=C_SUB)
    if not np.isnan(s_var_cli): _color_var(ws2, row, 7,  s_var_cli)
    if not np.isnan(s_var_sol): _color_var(ws2, row, 10, s_var_sol)
    row += 1

# Total general
t_cli_pre  = vendor_df['cli_pre'].sum()  / n_pre
t_cli_post = vendor_df['cli_post'].sum() / n_post
t_sol_pre  = vendor_df['soles_pre'].sum()  / n_pre
t_sol_post = vendor_df['soles_post'].sum() / n_post
t_var_cli  = t_cli_post / t_cli_pre   - 1 if t_cli_pre  > 0 else np.nan
t_var_sol  = t_sol_post / t_sol_pre   - 1 if t_sol_pre  > 0 else np.nan
t_cen_cli  = int(vendor_df['censo_cli_post'].sum())
t_cen_sol  = vendor_df['censo_soles_post'].sum()

for col, val, num in [
    (1,'TOTAL GENERAL',None),(2,'',None),(3,f'({len(vendor_df)} vendedores)',None),(4,'',None),
    (5,t_cli_pre,'#,##0.0'),(6,t_cli_post,'#,##0.0'),
    (7,t_var_cli if not np.isnan(t_var_cli) else 'N/D','0.0%'),
    (8,t_sol_pre,'#,##0'),(9,t_sol_post,'#,##0'),
    (10,t_var_sol if not np.isnan(t_var_sol) else 'N/D','0.0%'),
    (11,t_cen_cli,'#,##0'),(12,t_cen_sol,'#,##0'),(13,'',None),
]:
    _set(ws2, row, col, val, bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num=num)
if not np.isnan(t_var_cli): _color_var(ws2, row, 7, t_var_cli)
if not np.isnan(t_var_sol): _color_var(ws2, row, 10, t_var_sol)
row += 2

nota2 = (f'Pre = promedio mensual {LABELS[2601]}+{LABELS[2602]}. '
         f'Post = promedio mensual {LABELS[2603]}+{LABELS[2604]}* (Abr-26 proyectado lineal, factor {FACTOR_PROY:.4f}). '
         f'Solo vendedores activos en RUTA_ACTUAL. N/D = sin actividad en periodo pre.')
ws2.merge_cells(f'A{row}:{get_column_letter(len(COLS2))}{row}')
ws2.cell(row=row, column=1, value=nota2).font = fnt(sz=8)
ws2.freeze_panes = 'E3'


# ── HOJA 3: PENETRACION_CATEGORIAS ───────────────────────────────────
ws3 = wb.create_sheet('PENETRACION_CATEGORIAS')

COLS3 = ['CATEGORIA','N_CENSO\nTOTAL','CLI\nACTIVOS',
         f'CLI\n{LABELS[2603]}','CLI\n{LABELS[2604]}*',
         'PCT\nPENETRACION',
         f'SOLES\n{LABELS[2603]}',f'SOLES\n{LABELS[2604]}*','SOLES\nTOTAL']
WIDTHS3 = [24,10,10,10,10,12,14,14,14]

for i, w in enumerate(WIDTHS3, 1):
    ws3.column_dimensions[get_column_letter(i)].width = w

ws3.merge_cells(f'A1:{get_column_letter(len(COLS3))}1')
ws3['A1'] = 'PENETRACION DE CATEGORIAS EN CLIENTES CENSO — POST REESTRUCTURACION (08/03/2026)'
ws3['A1'].font = fnt(bold=True, color='FFFFFFFF', sz=12)
ws3['A1'].fill = fill(C_HDR1); ws3['A1'].alignment = aln()
ws3.row_dimensions[1].height = 26

for col, lbl in enumerate(COLS3, 1):
    c = ws3.cell(row=2, column=col, value=lbl)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
    c.fill = fill(C_HDR1); c.alignment = aln(); c.border = brd()
ws3.row_dimensions[2].height = 36

row = 3
for cat, cdata in cat_df.iterrows():
    bg_row = C_SSFF if cat in SSFF_CATS else (C_GRIS if row % 2 == 0 else C_BLC)
    _set(ws3, row, 1, cat,                  bg=bg_row, sz=10, halign='left')
    _set(ws3, row, 2, int(cdata['n_censo_total']),  bg=bg_row, sz=10, num='#,##0')
    _set(ws3, row, 3, int(cdata['cli_activos']),    bg=bg_row, sz=10, num='#,##0')
    _set(ws3, row, 4, int(cdata['cli_mar26']),      bg=bg_row, sz=10, num='#,##0')
    _set(ws3, row, 5, int(cdata['cli_abr26']),      bg=C_PESO, sz=10, num='#,##0')
    _set(ws3, row, 6, cdata['pct_penetracion'],     bg=bg_row, sz=10, num='0.0%')
    _set(ws3, row, 7, cdata['soles_mar26'],         bg=bg_row, sz=10, num='#,##0')
    _set(ws3, row, 8, cdata['soles_abr26'],         bg=C_PESO, sz=10, num='#,##0')
    _set(ws3, row, 9, cdata['soles_total'],         bg=bg_row, sz=10, num='#,##0')
    row += 1

# Total
_set(ws3, row, 1, 'TOTAL',              bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10)
_set(ws3, row, 2, int(cat_df['n_censo_total'].iloc[0]), bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
_set(ws3, row, 3, int(cat_df['cli_activos'].sum()),     bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
_set(ws3, row, 4, int(cat_df['cli_mar26'].sum()),       bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
_set(ws3, row, 5, int(cat_df['cli_abr26'].sum()),       bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
_set(ws3, row, 6, cat_df['cli_activos'].max() / n_censo_total, bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='0.0%')
_set(ws3, row, 7, cat_df['soles_mar26'].sum(),          bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
_set(ws3, row, 8, cat_df['soles_abr26'].sum(),          bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
_set(ws3, row, 9, cat_df['soles_total'].sum(),          bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
row += 2

nota3 = (f'Activos = clientes censo con al menos 1 compra en {LABELS[2603]} o {LABELS[2604]}. '
         f'N_CENSO_TOTAL = {n_censo_total:,} (hoja lista_clientes, columna censo). '
         f'Categorias SSFF resaltadas en naranja. '
         f'* {LABELS[2604]} = proyectado (factor {FACTOR_PROY:.4f}), cobertura sin proyectar.')
ws3.merge_cells(f'A{row}:{get_column_letter(len(COLS3))}{row}')
ws3.cell(row=row, column=1, value=nota3).font = fnt(sz=8)
ws3.freeze_panes = 'B3'


# ── HOJA 4: RESUMEN_SUPERVISORES ─────────────────────────────────────
ws4 = wb.create_sheet('RESUMEN_SUPERVISORES')

COLS4 = ['SUPERVISOR','N_VEND',
         f'CLI_PRE\nPROM/MES',f'CLI_POST\nPROM/MES','VAR\nCLI %',
         f'SOLES_PRE\nPROM/MES',f'SOLES_POST\nPROM/MES','VAR\nSOLES %',
         'CENSO\nCLI_POST','CENSO\nSOLES_POST']
WIDTHS4 = [22,8,13,13,10,15,15,10,12,15]

for i, w in enumerate(WIDTHS4, 1):
    ws4.column_dimensions[get_column_letter(i)].width = w

ws4.merge_cells(f'A1:{get_column_letter(len(COLS4))}1')
ws4['A1'] = f'RESUMEN POR SUPERVISOR — PRE ({LABELS[2601]}+{LABELS[2602]}) vs POST REESTRUCTURACION'
ws4['A1'].font = fnt(bold=True, color='FFFFFFFF', sz=12)
ws4['A1'].fill = fill(C_HDR1); ws4['A1'].alignment = aln()
ws4.row_dimensions[1].height = 26

for col, lbl in enumerate(COLS4, 1):
    c = ws4.cell(row=2, column=col, value=lbl)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
    c.fill = fill(C_HDR1); c.alignment = aln(); c.border = brd()
ws4.row_dimensions[2].height = 36

row = 3
for _, srow in sup_df.iterrows():
    bg_row = C_GRIS if row % 2 == 0 else C_BLC
    var_cli   = srow['var_cli']   if not np.isnan(srow['var_cli'])   else None
    var_soles = srow['var_soles'] if not np.isnan(srow['var_soles']) else None

    _set(ws4, row, 1,  srow['supervisor_vend'],    bg=bg_row, sz=10, halign='left')
    _set(ws4, row, 2,  int(srow['n_vendedores']),  bg=bg_row, sz=10, num='#,##0')
    _set(ws4, row, 3,  srow['cli_pre_avg'],        bg=bg_row, sz=10, num='#,##0.0')
    _set(ws4, row, 4,  srow['cli_post_avg'],       bg=bg_row, sz=10, num='#,##0.0')
    _set(ws4, row, 5,  var_cli   if var_cli   is not None else 'N/D', bg=bg_row, sz=10, num='0.0%')
    _set(ws4, row, 6,  srow['soles_pre_avg'],      bg=bg_row, sz=10, num='#,##0')
    _set(ws4, row, 7,  srow['soles_post_avg'],     bg=bg_row, sz=10, num='#,##0')
    _set(ws4, row, 8,  var_soles if var_soles is not None else 'N/D', bg=bg_row, sz=10, num='0.0%')
    _set(ws4, row, 9,  int(srow['censo_cli_post']),   bg=bg_row, sz=10, num='#,##0')
    _set(ws4, row, 10, srow['censo_soles_post'],       bg=bg_row, sz=10, num='#,##0')

    if var_cli   is not None: _color_var(ws4, row, 5, var_cli)
    if var_soles is not None: _color_var(ws4, row, 8, var_soles)
    row += 1

# Total
_set(ws4, row, 1, 'TOTAL',           bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10)
_set(ws4, row, 2, int(sup_df['n_vendedores'].sum()), bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
_set(ws4, row, 3, sup_df['cli_pre_avg'].sum(),       bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0.0')
_set(ws4, row, 4, sup_df['cli_post_avg'].sum(),      bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0.0')
_set(ws4, row, 5, t_var_cli if not np.isnan(t_var_cli) else 'N/D', bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='0.0%')
_set(ws4, row, 6, sup_df['soles_pre_avg'].sum(),     bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
_set(ws4, row, 7, sup_df['soles_post_avg'].sum(),    bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
_set(ws4, row, 8, t_var_sol if not np.isnan(t_var_sol) else 'N/D', bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='0.0%')
_set(ws4, row, 9, int(sup_df['censo_cli_post'].sum()),  bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
_set(ws4, row, 10, sup_df['censo_soles_post'].sum(),    bold=True, fg='FFFFFFFF', bg=C_ORG, sz=10, num='#,##0')
if not np.isnan(t_var_cli): _color_var(ws4, row, 5, t_var_cli)
if not np.isnan(t_var_sol): _color_var(ws4, row, 8, t_var_sol)
row += 2

nota4 = (f'Ordenado por VAR SOLES% descendente. '
         f'Pre = {LABELS[2601]}+{LABELS[2602]} | Post = {LABELS[2603]}+{LABELS[2604]}* (proyectado). '
         f'Censo CLI_POST = clientes censo con compra real (sin proyectar).')
ws4.merge_cells(f'A{row}:{get_column_letter(len(COLS4))}{row}')
ws4.cell(row=row, column=1, value=nota4).font = fnt(sz=8)
ws4.freeze_panes = 'C3'


# ════════════════════════════════════════════════════════════════════════
# GUARDAR
# ════════════════════════════════════════════════════════════════════════
wb.save(OUTPUT)
print(f"\nGuardado: {OUTPUT}")

# ── Verificacion de datos ─────────────────────────────────────────────
cob_total_mar = int(cob_ruta_mes[2603].sum()) if 2603 in cob_ruta_mes.columns else 0
cob_total_abr = int(cob_ruta_mes[2604].sum()) if 2604 in cob_ruta_mes.columns else 0
rutas_con_censo = int((censo_por_ruta > 0).sum())
top_cat = cat_df.index[0] if len(cat_df) > 0 else '-'

print(f"\n{'='*55}")
print(f"  Clientes censo:          {len(lc_censo):,}")
print(f"  Rutas con censo > 0:     {rutas_con_censo} de {len(rutas_oficiales)}")
print(f"  Cobertura censo Mar-26:  {cob_total_mar:,} clientes unicos")
print(f"  Cobertura censo Abr-26:  {cob_total_abr:,} clientes unicos")
print(f"  Soles censo Mar-26:      S/ {float(sol_ruta_mes[2603].sum()):,.0f}" if 2603 in sol_ruta_mes.columns else "")
print(f"  Soles censo Abr-26*:     S/ {float(sol_ruta_mes[2604].sum()):,.0f}" if 2604 in sol_ruta_mes.columns else "")
print(f"  Vendedores en tabla:     {len(vendor_df)}")
print(f"  Categoria top soles:     {top_cat} (S/ {cat_df.loc[top_cat,'soles_total']:,.0f})")
print(f"{'='*55}")
print("Done.")
