import datetime
import os
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings('ignore')

# Leer .env del proyecto
def _leer_env(path='.env'):
    env = {}
    if os.path.exists(path):
        for line in open(path, encoding='utf-8'):
            line = line.strip()
            if line and not line.startswith('#') and '=' in line:
                k, v = line.split('=', 1)
                env[k.strip()] = v.strip()
    return env

_env = _leer_env('C:/proyectos/SSFF/.env')
SQL_SERVER   = _env.get('SQL_SERVER',   'AUREN22\\AUREN')
SQL_DATABASE = _env.get('SQL_DATABASE', 'eAuren')
SQL_USER     = _env.get('SQL_USER',     'eauren')
SQL_PASSWORD = _env.get('SQL_PASSWORD', 'eauren')

# ════════════════════════════════════════════════════════════════════════
# PARAMETROS  <- ajustar rutas si es necesario
# ════════════════════════════════════════════════════════════════════════
HIST_PATH      = 'C:/proyectos/SSFF/export_data_ssff_2511-2604_fecha.csv'
MAEST_PATH     = 'C:/proyectos/SSFF/maestro_clientes_SSFF.xlsx'
TABLAS_PATH    = 'C:/proyectos/SSFF/TABLAS_RUTAS.xlsx'
OUTPUT         = 'C:/proyectos/SSFF/funnel_censo_SSFF.xlsx'
ENCUESTAS_PATH = 'C:/proyectos/SSFF/encuestas_mayo2026.csv'   # <- actualizar cada mes

# Query SQL para mayo (mes en curso) -- se ejecuta en cada corrida
QUERY_MAYO = """
SELECT
     a.ccod_cli
    ,a.[mes]
    ,a.[ccod_ruta]
    ,a.ccod_vend
    ,b.[categoria]
    ,a.fecha
    ,SUM(a.[cantidad])  AS [pedidos]
    ,SUM(a.[volumen])   AS [total_volumen]
    ,SUM(a.[monto])     AS [total_monto]
FROM [eAuren].[dbo].[base_com] a
LEFT JOIN [comercial_productos] b
    ON a.[codProd] = b.[codProd]
WHERE a.mes = '2605'
  AND a.cstatus  != 'A'
  AND a.ctipo_vta = '0003'
  AND ccod_ruta  != '0000'
GROUP BY
     a.ccod_cli
    ,a.[mes]
    ,a.[ccod_ruta]
    ,a.ccod_vend
    ,b.[categoria]
    ,a.fecha
"""

HOY = datetime.date.today()

CATS_VALIDAS = ['ACCESORIOS','ANDINA','CERDO','COLGATE','DERMODIS','DULFINA',
                'HIGIENE Y CUIDADO','HOMEPRO PERU','HUEVO','KIMBERLY','LA CORONA',
                'LA PATRONA','MEDIFARMA','PAVO','POLLO','PROCESADOS','RINTI',
                'TAMBOS PERU','VERDUM','YICHANG']

# ════════════════════════════════════════════════════════════════════════
# PALETA DE COLORES
# ════════════════════════════════════════════════════════════════════════
C_AZUL_OSC = 'FF0E1C36'   # azul marino oscuro — cabeceras principales
C_AZUL_MED = 'FFAFCBFF'   # azul pastel        — cabeceras secundarias
C_AZUL_MED_T='FF0E1C36'  # azul marino — texto sobre azul pastel (oscuro sobre claro)
C_AZUL_CLA = 'FFD7F9FF'   # celeste muy claro  — proyectado / destacado
C_VERDE    = 'FFD7F9FF'   # celeste muy claro  — checkmark cobertura
C_VERDE_T  = 'FF0E1C36'   # azul marino        — texto sobre celeste
C_AMARILLO = 'FFD7F9FF'   # celeste muy claro  — mayo en curso
C_AMARILLO_T='FF0E1C36'  # azul marino        — texto etiquetas proyectado
C_ROJO_CLA = 'FFFFC7CE'   # rojo claro (semáforo — sin cambio)
C_ROJO_T   = 'FF9C0006'   # rojo oscuro texto  (semáforo — sin cambio)
C_GRIS     = 'FFF9FBF2'   # blanco verdoso     — filas alternas
C_BLANCO   = 'FFFFFFFF'
C_NARANJA  = 'FFFFEDE1'   # melocotón suave    — filas de total / destacados
C_NARANJA_T= 'FF0E1C36'   # azul marino        — texto sobre melocotón

# ════════════════════════════════════════════════════════════════════════
# HELPERS EXCEL
# ════════════════════════════════════════════════════════════════════════
def fill(h):  return PatternFill('solid', fgColor=h)
def fnt(bold=False, color='FF000000', sz=11):
    return Font(name='Aptos Narrow', bold=bold, color=color, size=sz)
def aln(h='center', wrap=False):
    return Alignment(horizontal=h, vertical='center', wrap_text=wrap)
def brd(color='FFCCCCCC'):
    s = Side(style='thin', color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def cset(ws, row, col, val=None, bold=False, fg='FF000000', bg=None,
         sz=10, halign='center', num=None, merge_to=None, wrap=False):
    c = ws.cell(row=row, column=col, value=val)
    c.font      = fnt(bold=bold, color=fg, sz=sz)
    c.alignment = aln(halign, wrap)
    c.border    = brd()
    if bg:  c.fill = fill(bg)
    if num: c.number_format = num
    if merge_to:
        ws.merge_cells(
            f'{get_column_letter(col)}{row}:{get_column_letter(merge_to)}{row}')
    return c

def titulo(ws, row, texto, ncols, bg=C_AZUL_OSC, sz=12):
    ws.merge_cells(f'A{row}:{get_column_letter(ncols)}{row}')
    c = ws['A' + str(row)]
    c.value = texto
    c.font      = fnt(bold=True, color=C_BLANCO, sz=sz)
    c.fill      = fill(bg)
    c.alignment = aln()
    ws.row_dimensions[row].height = 24

def subtitulo(ws, row, texto, ncols, bg=C_AZUL_MED, col_ini=1):
    ws.merge_cells(f'{get_column_letter(col_ini)}{row}:{get_column_letter(ncols)}{row}')
    c = ws.cell(row=row, column=col_ini, value=texto)
    c.font      = fnt(bold=True, color=C_BLANCO, sz=10)
    c.fill      = fill(bg)
    c.alignment = aln()
    ws.row_dimensions[row].height = 20

# ════════════════════════════════════════════════════════════════════════
# [1] CARGA DE DATOS
# ════════════════════════════════════════════════════════════════════════
print("[1] Cargando datos...")

def _leer_csv(path):
    df = pd.read_csv(path, low_memory=False)
    df.columns = df.columns.str.strip()
    df['ccod_cli']    = pd.to_numeric(df['ccod_cli'],    errors='coerce').fillna(0).astype(int)
    df['mes']         = pd.to_numeric(df['mes'],         errors='coerce').fillna(0).astype(int)
    df['ccod_vend']   = pd.to_numeric(df['ccod_vend'],   errors='coerce').fillna(0).astype(int)
    df['total_monto'] = pd.to_numeric(df['total_monto'], errors='coerce').fillna(0)
    df['pedidos']     = pd.to_numeric(df.get('pedidos', 0), errors='coerce').fillna(0)
    df['fecha']       = pd.to_datetime(df['fecha'],      errors='coerce')
    df = df[df['categoria'].isin(CATS_VALIDAS)]
    return df

def _leer_sql_mayo():
    try:
        import pyodbc
        conn_str = (
            f"DRIVER={{SQL Server}};"
            f"SERVER={SQL_SERVER};"
            f"DATABASE={SQL_DATABASE};"
            f"UID={SQL_USER};"
            f"PWD={SQL_PASSWORD};"
        )
        conn = pyodbc.connect(conn_str, timeout=30)
        df = pd.read_sql(QUERY_MAYO, conn)
        conn.close()
        df.columns = df.columns.str.strip()
        df['ccod_cli']    = pd.to_numeric(df['ccod_cli'],    errors='coerce').fillna(0).astype(int)
        df['mes']         = pd.to_numeric(df['mes'],         errors='coerce').fillna(0).astype(int)
        df['ccod_vend']   = pd.to_numeric(df['ccod_vend'],   errors='coerce').fillna(0).astype(int)
        df['total_monto'] = pd.to_numeric(df['total_monto'], errors='coerce').fillna(0)
        df['fecha']       = pd.to_datetime(df['fecha'],      errors='coerce')
        df = df[df['categoria'].isin(CATS_VALIDAS)]
        return df, True
    except Exception as e:
        print(f"   AVISO: No se pudo conectar a SQL Server -- {e}")
        print(f"   Continuando solo con historico hasta 2604.")
        return pd.DataFrame(), False

df_hist = _leer_csv(HIST_PATH)

print("   Consultando SQL Server para mayo 2026...")
df_mayo, tiene_mayo = _leer_sql_mayo()
if tiene_mayo and len(df_mayo) > 0:
    df = pd.concat([df_hist, df_mayo], ignore_index=True)
    print(f"   Mayo desde SQL: {len(df_mayo):,} filas")
else:
    df = df_hist
    tiene_mayo = False

df = df.dropna(subset=['fecha'])
print(f"   Total filas: {len(df):,}  |  Rango: {df['fecha'].min().date()} a {df['fecha'].max().date()}")

# Maestro clientes
lc_raw = pd.read_excel(MAEST_PATH, sheet_name='lista_clientes')
lc_raw['codigo'] = pd.to_numeric(lc_raw['codigo'], errors='coerce').fillna(0).astype(int)
lc_raw['es_censo'] = lc_raw['censo'].notna()
lc = lc_raw.drop_duplicates('codigo', keep='first').copy()
lc['ruta']     = lc['ruta'].astype(str).str.strip()
lc['cliente']  = lc['cliente'].fillna('').astype(str).str.strip()
lc['vendedor'] = lc['vendedor'].fillna('').astype(str).str.strip()
lc['supervisor']= lc['supervisor'].fillna('').astype(str).str.strip()
lc['giro']     = lc['giro'].fillna('').astype(str)

lc_censo    = lc[lc['es_censo']].copy()
lc_no_censo = lc[~lc['es_censo']].copy()

# Encuestas de visita (CSV actualizable mensualmente)
def _cargar_encuestas(path):
    if not os.path.exists(path):
        print(f"   [AVISO] No se encontró {path} — columna VALIDACION VISITA quedará vacía.")
        return pd.Series(dtype=str), pd.Series(dtype=str)
    for enc in ('utf-8-sig', 'latin-1', 'cp1252'):
        try:
            df_enc = pd.read_csv(path, encoding=enc, low_memory=False)
            break
        except UnicodeDecodeError:
            continue
    df_enc.columns = df_enc.columns.str.strip()
    df_enc['codCliente'] = pd.to_numeric(df_enc['codCliente'], errors='coerce').fillna(0).astype(int)
    # Tomar la validación más reciente por cliente (por si hay duplicados)
    if 'Fecha' in df_enc.columns:
        df_enc['Fecha'] = pd.to_datetime(df_enc['Fecha'], dayfirst=True, errors='coerce')
        df_enc = df_enc.sort_values('Fecha', ascending=False)
    col_val = 'Validacion de la visita'
    if col_val not in df_enc.columns:
        print(f"   [AVISO] Columna '{col_val}' no encontrada en encuestas.")
        return pd.Series(dtype=str), pd.Series(dtype=str)
    dedup = df_enc.drop_duplicates('codCliente', keep='first').set_index('codCliente')
    mapa_val  = dedup[col_val]
    mapa_fec  = dedup['Fecha'] if 'Fecha' in dedup.columns else pd.Series(dtype='datetime64[ns]')
    print(f"   Encuestas cargadas: {len(mapa_val):,} clientes con validación de visita.")
    return mapa_val, mapa_fec

enc_validacion, enc_fecha = _cargar_encuestas(ENCUESTAS_PATH)

# Maestro rutas
tablas = pd.read_excel(TABLAS_PATH, sheet_name='RUTA_ACTUAL')
tablas['RUTA'] = tablas['RUTA'].astype(str).str.strip()
ruta_sup_map  = tablas.set_index('RUTA')['SUPERVISOR'].to_dict()
ruta_vend_map = tablas.set_index('RUTA')['VENDEDOR'].to_dict()
ruta_nomv_map = tablas.set_index('RUTA')['NOMBRE LARGO'].to_dict()

def ruta_to_ffvv(r):
    if r.startswith('KB'):   return 'KB'
    if r.startswith('M0'):   return 'M0'
    if r.startswith('P0'):   return 'P0'
    if r in ('V001','V002'): return 'V0'
    if r.startswith('V0'):   return 'V0'
    return 'F8'

print(f"   Censo: {len(lc_censo):,} | No-censo: {len(lc_no_censo):,} | Total cartera: {len(lc):,}")


# ════════════════════════════════════════════════════════════════════════
# [2] CALCULOS FUNNEL
# ════════════════════════════════════════════════════════════════════════
print("[2] Calculando funnel censo...")

censo_ids    = set(lc_censo['codigo'])
no_censo_ids = set(lc_no_censo['codigo'])

PERIODOS_POST = [2603, 2604]
if tiene_mayo: PERIODOS_POST.append(2605)

df_post = df[df['mes'].isin(PERIODOS_POST)]
df_censo_post    = df_post[df_post['ccod_cli'].isin(censo_ids)]
df_no_censo_post = df_post[df_post['ccod_cli'].isin(no_censo_ids)]

def cob_mes(mes_val):
    sub = df_censo_post[df_censo_post['mes'] == mes_val]
    return sub['ccod_cli'].nunique(), sub['total_monto'].sum()

n_censo_total    = len(lc_censo)
cob_total        = df_censo_post['ccod_cli'].nunique()
sol_total        = df_censo_post['total_monto'].sum()
ped_total_censo  = df_censo_post['pedidos'].sum()
cob_mar, sol_mar = cob_mes(2603)
cob_abr, sol_abr = cob_mes(2604)
cob_may, sol_may = cob_mes(2605) if tiene_mayo else (None, None)

# Ticket no-censo por mes: soles / clientes coberturados no-censo de ese mes
def _ticket_nc_mes(mes_val):
    sub = df_no_censo_post[df_no_censo_post['mes'] == mes_val]
    sol = sub['total_monto'].sum()
    cob = sub['ccod_cli'].nunique()
    return sol / cob if cob else None

ticket_nc_mar = _ticket_nc_mes(2603)
ticket_nc_abr = _ticket_nc_mes(2604)
ticket_nc_may = _ticket_nc_mes(2605) if tiene_mayo else None

print(f"   Cobertura total post-reest.: {cob_total:,} ({cob_total/n_censo_total:.1%})")
print(f"   Mar-26: {cob_mar:,} | Abr-26: {cob_abr:,}" +
      (f" | May-26: {cob_may:,}" if tiene_mayo else ""))


# ════════════════════════════════════════════════════════════════════════
# [3] TABLA DETALLE CENSO (supervisor > vendedor > cliente)
# ════════════════════════════════════════════════════════════════════════
print("[3] Construyendo tabla detalle censo...")

# Soles por cliente x mes
sol_cli_mes = (df_censo_post.groupby(['ccod_cli','mes'])['total_monto']
               .sum().unstack(fill_value=0))
for m in PERIODOS_POST:
    if m not in sol_cli_mes.columns: sol_cli_mes[m] = 0.0

# Unir con maestro censo
det_censo = lc_censo[['codigo','cliente','ruta','supervisor','vendedor','giro','diaVisita']].copy()
det_censo = det_censo.join(sol_cli_mes, on='codigo', how='left').fillna(0)
det_censo['ffvv']      = det_censo['ruta'].apply(ruta_to_ffvv)
det_censo['nom_vend']  = det_censo['ruta'].map(ruta_nomv_map).fillna(det_censo['vendedor'])
det_censo['diaVisita'] = det_censo['diaVisita'].fillna('').astype(str)
det_censo = det_censo.sort_values(['supervisor','nom_vend','cliente'])


# ════════════════════════════════════════════════════════════════════════
# [4] ULTIMA COMPRA POR CLIENTE
# ════════════════════════════════════════════════════════════════════════
print("[4] Calculando ultima compra...")

def calcular_ultima(df_src, ids):
    compras = df_src[df_src['ccod_cli'].isin(ids)].copy()
    if compras.empty:
        return pd.DataFrame(columns=['ccod_cli','fecha_ultima','cat_ultima','monto_ultima','dias'])
    ult = (compras.sort_values(['ccod_cli','fecha','total_monto'],
                                ascending=[True, False, False])
                  .groupby('ccod_cli', as_index=False)
                  .first()
                  [['ccod_cli','fecha','categoria','total_monto']])
    ult.columns = ['ccod_cli','fecha_ultima','cat_ultima','monto_ultima']
    ult['dias'] = ult['fecha_ultima'].apply(
        lambda x: (HOY - x.date()).days if pd.notna(x) else 9999)
    return ult

ult_censo    = calcular_ultima(df, censo_ids)
ult_no_censo = calcular_ultima(df, no_censo_ids)

def enriquecer(lc_df, ult_df):
    base = lc_df[['codigo','cliente','ruta','supervisor','vendedor','giro','diaVisita']].copy()
    res  = base.merge(ult_df, left_on='codigo', right_on='ccod_cli', how='left')
    res['cat_ultima']   = res['cat_ultima'].fillna('SIN COMPRA')
    res['monto_ultima'] = res['monto_ultima'].fillna(0)
    res['dias']         = res['dias'].fillna(9999).astype(int)
    res['nom_vend']     = res['ruta'].map(ruta_nomv_map).fillna(res['vendedor'])
    res['diaVisita']    = res['diaVisita'].fillna('').astype(str)
    res = res.sort_values('dias', ascending=False)
    return res

ult_censo_enr    = enriquecer(lc_censo,    ult_censo)
ult_no_censo_enr = enriquecer(lc_no_censo, ult_no_censo)

# Soles por cliente x mes desde Ene-26 (para ULTIMA_COMPRA_NO_CENSO)
MESES_NC = [m for m in sorted(df['mes'].unique()) if m >= 2601]
sol_nc_mes = (df[df['ccod_cli'].isin(no_censo_ids) & df['mes'].isin(MESES_NC)]
              .groupby(['ccod_cli','mes'])['total_monto'].sum()
              .unstack(fill_value=0))
for m in MESES_NC:
    if m not in sol_nc_mes.columns: sol_nc_mes[m] = 0.0
NC_MES_LABELS = {2601:'Ene-26',2602:'Feb-26',2603:'Mar-26',
                 2604:'Abr-26',2605:'May-26',2606:'Jun-26'}

sin_compra_censo = (ult_censo_enr['cat_ultima'] == 'SIN COMPRA').sum()
print(f"   Censo sin ninguna compra post-reest.: {sin_compra_censo:,}")
print(f"   No-censo en tabla ultima compra: {len(ult_no_censo_enr):,}")


# ════════════════════════════════════════════════════════════════════════
# [5] CALCULOS RESUMEN SUPERVISORES Y VENDEDORES
# ════════════════════════════════════════════════════════════════════════
print("[5] Calculando resumenes por supervisor y vendedor...")

# Cobertura (n clientes unicos que compraron) por supervisor x mes
def _cob_sup(mes_val):
    sub = df_censo_post[df_censo_post['mes'] == mes_val].copy()
    sub = sub.merge(lc_censo[['codigo','supervisor']], left_on='ccod_cli', right_on='codigo', how='left')
    return sub.groupby('supervisor')['ccod_cli'].nunique()

# Soles por supervisor x mes
def _sol_sup(mes_val):
    sub = df_censo_post[df_censo_post['mes'] == mes_val].copy()
    sub = sub.merge(lc_censo[['codigo','supervisor']], left_on='ccod_cli', right_on='codigo', how='left')
    return sub.groupby('supervisor')['total_monto'].sum()

# Total cartera y censo por supervisor (desde lista_clientes)
total_car_sup = lc.groupby(lc['supervisor'])['codigo'].count().rename('total_cartera')
censo_car_sup = lc_censo.groupby('supervisor')['codigo'].count().rename('cartera_censo')

cob_sup_mar = _cob_sup(2603)
sol_sup_mar = _sol_sup(2603)
cob_sup_abr = _cob_sup(2604)
sol_sup_abr = _sol_sup(2604)
cob_sup_may = _cob_sup(2605) if tiene_mayo else pd.Series(dtype=float)
sol_sup_may = _sol_sup(2605) if tiene_mayo else pd.Series(dtype=float)

# Soles totales (toda la cartera, no solo censo) por supervisor x mes — para % soles
def _sol_sup_total(mes_val):
    sub = df_post[df_post['mes'] == mes_val].copy()
    sub = sub.merge(lc[['codigo','supervisor']], left_on='ccod_cli', right_on='codigo', how='left')
    return sub.groupby('supervisor')['total_monto'].sum()

sol_sup_total_mar = _sol_sup_total(2603)
sol_sup_total_abr = _sol_sup_total(2604)
sol_sup_total_may = _sol_sup_total(2605) if tiene_mayo else pd.Series(dtype=float)

# Construir DataFrame resumen supervisores
sup_resumen = pd.DataFrame({'cartera_censo': censo_car_sup})
sup_resumen = sup_resumen.join(total_car_sup, how='left').fillna(0)
sup_resumen['cob_mar']       = cob_sup_mar.reindex(sup_resumen.index).fillna(0).astype(int)
sup_resumen['sol_mar']       = sol_sup_mar.reindex(sup_resumen.index).fillna(0)
sup_resumen['sol_tot_mar']   = sol_sup_total_mar.reindex(sup_resumen.index).fillna(0)
sup_resumen['cob_abr']       = cob_sup_abr.reindex(sup_resumen.index).fillna(0).astype(int)
sup_resumen['sol_abr']       = sol_sup_abr.reindex(sup_resumen.index).fillna(0)
sup_resumen['sol_tot_abr']   = sol_sup_total_abr.reindex(sup_resumen.index).fillna(0)
if tiene_mayo:
    sup_resumen['cob_may']     = cob_sup_may.reindex(sup_resumen.index).fillna(0).astype(int)
    sup_resumen['sol_may']     = sol_sup_may.reindex(sup_resumen.index).fillna(0)
    sup_resumen['sol_tot_may'] = sol_sup_total_may.reindex(sup_resumen.index).fillna(0)
sup_resumen = sup_resumen.sort_index()  # orden temporal; se reordena por no_cob mas abajo

# Resumen por vendedor (ruta)
det_vend = det_censo.groupby(['ruta','supervisor','nom_vend','ffvv']).agg(
    cartera_censo=('codigo','count'),
    **{str(m): (m,'sum') for m in PERIODOS_POST}
).reset_index()

# Total cartera por ruta (censo + no-censo)
total_car_ruta = lc.groupby('ruta')['codigo'].count().rename('total_cartera').reset_index()
det_vend = det_vend.merge(total_car_ruta, on='ruta', how='left').fillna({'total_cartera': 0})
det_vend['total_cartera'] = det_vend['total_cartera'].astype(int)

# Soles totales por ruta x mes (toda la cartera) — para % soles
def _sol_ruta_total(mes_val):
    sub = df_post[df_post['mes'] == mes_val].copy()
    sub = sub.merge(lc[['codigo','ruta']], left_on='ccod_cli', right_on='codigo', how='left')
    return sub.groupby('ruta')['total_monto'].sum().rename(f'sol_tot_{mes_val}')

# Cobertura por vendedor (ruta) x mes: n clientes unicos con venta
def _cob_ruta(mes_val):
    sub = df_censo_post[df_censo_post['mes'] == mes_val].copy()
    sub = sub.merge(lc_censo[['codigo','ruta']], left_on='ccod_cli', right_on='codigo', how='left')
    return sub.groupby('ruta')['ccod_cli'].nunique().rename(f'cob_{mes_val}')

cob_ruta_mar = _cob_ruta(2603)
cob_ruta_abr = _cob_ruta(2604)
sol_ruta_tot_mar = _sol_ruta_total(2603)
sol_ruta_tot_abr = _sol_ruta_total(2604)
det_vend = det_vend.join(cob_ruta_mar,     on='ruta').fillna({'cob_2603': 0})
det_vend = det_vend.join(cob_ruta_abr,     on='ruta').fillna({'cob_2604': 0})
det_vend = det_vend.join(sol_ruta_tot_mar, on='ruta').fillna({'sol_tot_2603': 0})
det_vend = det_vend.join(sol_ruta_tot_abr, on='ruta').fillna({'sol_tot_2604': 0})
if tiene_mayo:
    cob_ruta_may     = _cob_ruta(2605)
    sol_ruta_tot_may = _sol_ruta_total(2605)
    det_vend = det_vend.join(cob_ruta_may,     on='ruta').fillna({'cob_2605': 0})
    det_vend = det_vend.join(sol_ruta_tot_may, on='ruta').fillna({'sol_tot_2605': 0})
# Calcular no_cob por supervisor y ruta (para ordenar tablas)
_cob_sup_total_s  = (df_censo_post
                     .merge(lc_censo[['codigo','supervisor']], left_on='ccod_cli', right_on='codigo', how='left')
                     .groupby('supervisor')['ccod_cli'].nunique())
sup_resumen['no_cob'] = sup_resumen['cartera_censo'] - _cob_sup_total_s.reindex(sup_resumen.index).fillna(0).astype(int)
sup_resumen = sup_resumen.sort_values('no_cob', ascending=False)

_cob_ruta_total_s = (df_censo_post
                     .merge(lc_censo[['codigo','ruta']], left_on='ccod_cli', right_on='codigo', how='left')
                     .groupby('ruta')['ccod_cli'].nunique())
det_vend['no_cob'] = det_vend['cartera_censo'] - det_vend['ruta'].map(_cob_ruta_total_s).fillna(0).astype(int)
det_vend = det_vend.sort_values('no_cob', ascending=False)


# ════════════════════════════════════════════════════════════════════════
# [6] TABLA CATEGORIAS x MES (clientes censo)
# ════════════════════════════════════════════════════════════════════════
cat_mes_df = (df_censo_post.groupby(['categoria','mes'])['total_monto']
              .sum().unstack(fill_value=0).reset_index())
for m in PERIODOS_POST:
    if m not in cat_mes_df.columns:
        cat_mes_df[m] = 0.0
cat_mes_df['_total'] = cat_mes_df[[m for m in PERIODOS_POST]].sum(axis=1)
cat_mes_df = cat_mes_df.sort_values('_total', ascending=False).drop(columns='_total')


# ════════════════════════════════════════════════════════════════════════
# FACTOR DE PROYECCION MAYO (necesario para todas las hojas)
# ════════════════════════════════════════════════════════════════════════
# Feriados nacionales Perú 2026 (si caen domingo no afectan, ya lo maneja la lógica)
FERIADOS_2026 = {
    datetime.date(2026,  1,  1),  # Año Nuevo
    datetime.date(2026,  4,  9),  # Jueves Santo
    datetime.date(2026,  4, 10),  # Viernes Santo
    datetime.date(2026,  5,  1),  # Día del Trabajo
    datetime.date(2026,  6, 29),  # San Pedro y San Pablo
    datetime.date(2026,  7, 26),  # No laborable Fiestas Patrias
    datetime.date(2026,  7, 28),  # Fiestas Patrias
    datetime.date(2026,  7, 29),  # Fiestas Patrias
    datetime.date(2026,  8,  6),  # Batalla de Junín
    datetime.date(2026,  8, 30),  # Santa Rosa de Lima
    datetime.date(2026, 10,  8),  # Combate de Angamos
    datetime.date(2026, 11,  1),  # Todos los Santos
    datetime.date(2026, 12,  8),  # Inmaculada Concepción
    datetime.date(2026, 12, 25),  # Navidad
}

def _dias_habiles(fecha_ini, fecha_fin):
    """Cuenta días hábiles entre fecha_ini y fecha_fin (ambos inclusive),
    excluyendo solo domingos y feriados (sábado es hábil)."""
    feriados_no_dom = {f for f in FERIADOS_2026 if f.weekday() != 6}
    count = 0
    d = fecha_ini
    while d <= fecha_fin:
        if d.weekday() != 6 and d not in feriados_no_dom:
            count += 1
        d += datetime.timedelta(days=1)
    return count

if tiene_mayo:
    _ini_mayo     = datetime.date(2026, 5, 1)
    _fin_mayo     = datetime.date(2026, 5, 31)
    _dias_trans   = _dias_habiles(_ini_mayo, HOY)
    _dias_hab_mes = _dias_habiles(_ini_mayo, _fin_mayo)
    FACTOR_PROY_MAY = _dias_hab_mes / _dias_trans if _dias_trans > 0 else 1.0
    print(f"   Proyección mayo: {_dias_trans} días háb. transcurridos (al {HOY}) / {_dias_hab_mes} días háb. totales  (factor {FACTOR_PROY_MAY:.2f}x)")
else:
    FACTOR_PROY_MAY = None

# ════════════════════════════════════════════════════════════════════════
# [7] GENERAR EXCEL
# ════════════════════════════════════════════════════════════════════════
print("[7] Generando Excel...")

wb = Workbook()
wb.remove(wb.active)

FECHA_HOY_STR = HOY.strftime('%d/%m/%Y')

# helper para porcentaje seguro (evita div/0)
def _pct(num, den):
    return num / den if den else 0.0


# ── HOJA 1: FUNNEL_CENSO ─────────────────────────────────────────────
ws1 = wb.create_sheet('FUNNEL_CENSO')

# Ocultar líneas de cuadrícula
ws1.sheet_view.showGridLines = False

# Anchos de columna — diseño vertical, tablas completas en cols A-P
ws1.column_dimensions['A'].width = 3.0    # margen izq
ws1.column_dimensions['B'].width = 28.0   # SUPERVISOR / SEGMENTO / ETAPA
ws1.column_dimensions['C'].width = 14.0   # CLIENTES / COB
ws1.column_dimensions['D'].width = 10.0   # %
ws1.column_dimensions['E'].width = 15.0   # SOLES
ws1.column_dimensions['F'].width = 14.0   # TICKET PROM CENSO
ws1.column_dimensions['G'].width = 14.0   # TICKET PROM NO CENSO
ws1.column_dimensions['H'].width = 4.0    # margen separador
ws1.column_dimensions['I'].width = 20.0   # CATEGORIA
ws1.column_dimensions['J'].width = 15.0   # MARZO soles
ws1.column_dimensions['K'].width = 15.0   # ABRIL soles
ws1.column_dimensions['L'].width = 15.0   # MAYO soles
ws1.column_dimensions['M'].width = 15.0   # PROY MAYO (cat) / CARTERA CENSO (tablas sup/vend)

# Anchos para tabla supervisores/vendedores (cols B-W)
ws1.column_dimensions['N'].width = 9.0    # MAR COB
ws1.column_dimensions['O'].width = 14.0   # MAR SOLES
ws1.column_dimensions['P'].width = 9.0    # MAR % SOLES
ws1.column_dimensions['Q'].width = 9.0    # ABR COB
ws1.column_dimensions['R'].width = 14.0   # ABR SOLES
ws1.column_dimensions['S'].width = 9.0    # ABR % SOLES
ws1.column_dimensions['T'].width = 9.0    # MAY COB
ws1.column_dimensions['U'].width = 14.0   # MAY SOLES
ws1.column_dimensions['V'].width = 9.0    # MAY % SOLES
ws1.column_dimensions['W'].width = 10.0   # NO COB #
ws1.column_dimensions['X'].width = 10.0   # NO COB %

# Numero de columnas totales de la hoja (para titulos)
NCOLS_TOT = 24   # A..X (incluye cols ticket F,G y categorias desplazadas a I-M)

# ─── TITULO PRINCIPAL ─────────────────────────────────────────────────
ws1.merge_cells(f'A1:{get_column_letter(NCOLS_TOT)}1')
c = ws1['A1']
c.value     = f'FUNNEL SEGUIMIENTO CARTERA CENSO  --  Actualizado: {FECHA_HOY_STR}'
c.font      = fnt(bold=True, color=C_BLANCO, sz=14)
c.fill      = fill(C_AZUL_OSC)
c.alignment = aln()
ws1.row_dimensions[1].height = 32

# ─── BLOQUE A: Distribucion de cartera (cols B-E, filas 3-7) ──────────
ws1.row_dimensions[2].height = 8   # espacio

ws1.merge_cells('B3:E3')
c = ws1['B3']
c.value     = 'DISTRIBUCION DE CARTERA'
c.font      = fnt(bold=True, color=C_AZUL_MED_T, sz=14)
c.fill      = fill(C_AZUL_MED)
c.alignment = aln()
ws1.row_dimensions[3].height = 28

for col, lbl in [(2,'SEGMENTO'),(3,'CLIENTES'),(4,'%'),(5,'')]:
    cset(ws1, 4, col, lbl, bold=True, fg=C_BLANCO, bg=C_AZUL_OSC, sz=11)
ws1.row_dimensions[4].height = 20

datos_seg = [
    ('CENSO',    len(lc_censo),    _pct(len(lc_censo),    len(lc)), C_AZUL_CLA, 'FF000000'),
    ('NO CENSO', len(lc_no_censo), _pct(len(lc_no_censo), len(lc)), C_GRIS,     'FF000000'),
    ('TOTAL',    len(lc),          1.0,                              C_NARANJA,  C_NARANJA_T),
]
for i, (seg, n, pct, bg, fg) in enumerate(datos_seg, 5):
    bold = (bg == C_NARANJA)
    cset(ws1, i, 2, seg, bold=bold, fg=fg, bg=bg, sz=11, halign='left')
    cset(ws1, i, 3, n,   bold=bold, fg=fg, bg=bg, sz=11, num='#,##0')
    cset(ws1, i, 4, pct, bold=bold, fg=fg, bg=bg, sz=11, num='0.0%')
    cset(ws1, i, 5, '',  bg=bg)
ws1.row_dimensions[7].height = 14  # TOTAL un poco mas alto

# ─── BLOQUE B: Funnel de cobertura (cols B-E, filas 9+) ───────────────
ws1.row_dimensions[8].height = 10  # espacio

ws1.merge_cells('B9:E9')
c = ws1['B9']
c.value     = 'FUNNEL DE COBERTURA  --  CLIENTES CENSO POST-REESTRUCTURACION (08/03/2026)'
c.font      = fnt(bold=True, color=C_AZUL_MED_T, sz=10)
c.fill      = fill(C_AZUL_MED)
c.alignment = aln()
ws1.row_dimensions[9].height = 28

# Cabeceras funnel
hdrs_funnel = ['ETAPA', 'CLIENTES', '% SOBRE CENSO', 'SOLES (S/)', 'TICKET PROM\nCENSO', 'TICKET PROM\nNO CENSO']
for i, lbl in enumerate(hdrs_funnel):
    cset(ws1, 10, i+2, lbl, bold=True, fg=C_BLANCO, bg=C_AZUL_OSC, sz=11, wrap=(i>=4))
ws1.row_dimensions[10].height = 28

sin_cob = n_censo_total - cob_total

# Ticket promedio: total_monto / clientes coberturados
def _ticket(sol, cob_n): return sol / cob_n if cob_n else None

# funnel_rows: (etapa, cob, pct, sol, bg, bold, ticket_nc_val)
# ticket_nc_val: valor del ticket no-censo para esa fila (None = no mostrar)
funnel_rows = [
    ('UNIVERSO — Total clientes censo',
     n_censo_total, 1.0, None, C_BLANCO, True, None),
    (f'COBERTURADOS — Con compra Mar{"-May" if tiene_mayo else "-Abr"} 2026',
     cob_total, _pct(cob_total,n_censo_total), sol_total, C_AZUL_CLA, True, None),
    ('COBERTURADOS MAR-26 — Marzo 2026',
     cob_mar, _pct(cob_mar,n_censo_total), sol_mar, C_BLANCO, False, ticket_nc_mar),
    ('COBERTURADOS ABR-26 — Abril 2026',
     cob_abr, _pct(cob_abr,n_censo_total), sol_abr, C_BLANCO, False, ticket_nc_abr),
]
if tiene_mayo:
    funnel_rows.append((
        'COBERTURADOS MAY-26 — Mayo 2026 (en curso)',
        cob_may, _pct(cob_may,n_censo_total), sol_may, C_AMARILLO, False, ticket_nc_may))
else:
    funnel_rows.append((
        'COBERTURADOS MAY-26 — Sin datos',
        None, None, None, C_AMARILLO, False, None))
funnel_rows.append((
    'SIN COBERTURA — Sin compra post-reestructuracion',
    sin_cob, _pct(sin_cob,n_censo_total), 0, C_ROJO_CLA, True, None))

fila_f = 11
for etapa, cob, pct, sol, bg, bold_row, tnc in funnel_rows:
    is_sin_cob = bg == C_ROJO_CLA
    fg_txt = C_ROJO_T if is_sin_cob else 'FF000000'
    # ticket censo = soles de esta etapa / clientes coberturados de esta etapa
    ticket_c = _ticket(sol, cob) if (sol is not None and not is_sin_cob) else None
    cset(ws1, fila_f, 2, etapa, bold=bold_row, fg=fg_txt, bg=bg, sz=11, halign='left')
    if cob is not None:
        cset(ws1, fila_f, 3, cob,                           bold=bold_row, fg=fg_txt, bg=bg, sz=11, num='#,##0')
        cset(ws1, fila_f, 4, pct,                           fg=fg_txt, bg=bg, sz=11, num='0.0%')
        cset(ws1, fila_f, 5, sol if sol is not None else 0, fg=fg_txt, bg=bg, sz=11, num='#,##0')
        cset(ws1, fila_f, 6, ticket_c if ticket_c else '',  fg=fg_txt, bg=bg, sz=11, num='#,##0')
        cset(ws1, fila_f, 7, tnc if tnc else '',            fg=fg_txt, bg=C_AZUL_CLA, sz=11, num='#,##0')
    else:
        cset(ws1, fila_f, 3, 'S/D', bg=bg, sz=11)
        cset(ws1, fila_f, 4, '',    bg=bg)
        cset(ws1, fila_f, 5, '',    bg=bg)
        cset(ws1, fila_f, 6, '',    bg=bg)
        cset(ws1, fila_f, 7, '',    bg=bg)
    ws1.row_dimensions[fila_f].height = 18
    fila_f += 1

# ─── BLOQUE C: Tabla categorias x meses (cols I-M, filas 3+) ─────────
_cat_titulo_fin = 'M' if tiene_mayo and FACTOR_PROY_MAY else 'L'
ws1.merge_cells(f'I3:{_cat_titulo_fin}3')
c = ws1['I3']
c.value     = 'SOLES CENSO POR CATEGORIA'
c.font      = fnt(bold=True, color=C_AZUL_MED_T, sz=14)
c.fill      = fill(C_AZUL_MED)
c.alignment = aln()

hdrs_cat = [('I','CATEGORIA'),('J','MARZO'),('K','ABRIL'),('L','MAYO')]
if tiene_mayo and FACTOR_PROY_MAY:
    hdrs_cat.append(('M','PROY\nMAYO'))
for cl, lbl in hdrs_cat:
    c = ws1[f'{cl}4']
    c.value     = lbl
    c.font      = fnt(bold=True, color=C_BLANCO, sz=11)
    c.fill      = fill(C_AZUL_OSC)
    c.alignment = aln(wrap=True)
    c.border    = brd()
ws1.row_dimensions[4].height = 20

cat_row = 5
for _, cr in cat_mes_df.iterrows():
    bg_c = C_GRIS if cat_row % 2 == 0 else C_BLANCO
    cset(ws1, cat_row, 9,  cr['categoria'],       bg=bg_c, sz=11, halign='left')
    cset(ws1, cat_row, 10, float(cr.get(2603,0)), bg=bg_c, sz=11, num='#,##0')
    cset(ws1, cat_row, 11, float(cr.get(2604,0)), bg=bg_c, sz=11, num='#,##0')
    if tiene_mayo:
        sol_cat_may = float(cr.get(2605,0))
        cset(ws1, cat_row, 12, sol_cat_may, bg=C_AMARILLO, sz=11, num='#,##0')
        if FACTOR_PROY_MAY:
            cset(ws1, cat_row, 13, sol_cat_may * FACTOR_PROY_MAY, bg=C_AZUL_CLA, sz=11, num='#,##0')
    else:
        cset(ws1, cat_row, 12, 'S/D', bg=C_AMARILLO, sz=11)
    ws1.row_dimensions[cat_row].height = 16
    cat_row += 1

# Fila total categorias
sol_may_proy_total = (sol_may * FACTOR_PROY_MAY) if (tiene_mayo and FACTOR_PROY_MAY) else None
cset(ws1, cat_row, 9,  'TOTAL', bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11)
cset(ws1, cat_row, 10, sol_mar, bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='#,##0')
cset(ws1, cat_row, 11, sol_abr, bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='#,##0')
cset(ws1, cat_row, 12, sol_may if tiene_mayo else 0, bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='#,##0')
if sol_may_proy_total:
    cset(ws1, cat_row, 13, sol_may_proy_total, bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='#,##0')
ws1.row_dimensions[cat_row].height = 16

# Calcular fila de inicio para los bloques C/D (debajo del mas largo entre funnel y cat)
fila_tablas = max(fila_f, cat_row) + 2

# ─── BLOQUE D: Resumen por supervisor ─────────────────────────────────
ws1.row_dimensions[fila_tablas - 1].height = 10  # espacio

# Subtitulo supervisor — ocupa cols B-V
ws1.merge_cells(f'B{fila_tablas}:{get_column_letter(NCOLS_TOT)}{fila_tablas}')
c = ws1.cell(row=fila_tablas, column=2, value='RESUMEN POR SUPERVISOR')
c.font      = fnt(bold=True, color=C_AZUL_MED_T, sz=14)
c.fill      = fill(C_AZUL_MED)
c.alignment = aln()
ws1.row_dimensions[fila_tablas].height = 28
fila_tablas += 1

# Cabeceras supervisor
# Columnas: B=SUPERVISOR, C=TOTAL CART, D=CART CENSO,
#   E=margen, K=CARTERA CENSO, L=MAR COB, M=MAR SOLES, N=MAR %SOL,
#   O=ABR COB, P=ABR SOLES, Q=ABR %SOL, [R=MAY COB, S=MAY SOLES, T=MAY %SOL,]
#   U=NO COB #, V=NO COB %
# Simplificado: B..V empezando col 2
hdrs_s = ['SUPERVISOR','TOTAL\nCARTERA','CARTERA\nCENSO',
          'MAR-26\nCOB','MAR-26\nSOLES','MAR-26\n% SOLES',
          'ABR-26\nCOB','ABR-26\nSOLES','ABR-26\n% SOLES']
if tiene_mayo:
    hdrs_s += ['MAY-26\nCOB','MAY-26\nSOLES','MAY-26\n% SOLES']
hdrs_s += ['NO COB\n#','NO COB\n%']

for i, lbl in enumerate(hdrs_s):
    cset(ws1, fila_tablas, i+2, lbl, bold=True, fg=C_BLANCO, bg=C_AZUL_OSC, sz=10)
ws1.row_dimensions[fila_tablas].height = 32
fila_tablas += 1

# Soles totales de toda la cartera por supervisor (para denominador del %)
sol_tot_mar_all = df_post[df_post['mes']==2603]['total_monto'].sum()
sol_tot_abr_all = df_post[df_post['mes']==2604]['total_monto'].sum()
sol_tot_may_all = df_post[df_post['mes']==2605]['total_monto'].sum() if tiene_mayo else 0

for sup, row_s in sup_resumen.iterrows():
    if not sup or str(sup).strip() == '': continue
    censo_n    = int(row_s['cartera_censo'])
    no_cob_n   = int(row_s['no_cob'])
    tot_c      = int(row_s['total_cartera'])
    sol_m      = float(row_s['sol_mar'])
    tot_m      = float(row_s['sol_tot_mar'])
    sol_a      = float(row_s['sol_abr'])
    tot_a      = float(row_s['sol_tot_abr'])
    c_idx = 2
    cset(ws1, fila_tablas, c_idx,   sup,      bg=C_AZUL_CLA, sz=11, halign='left'); c_idx+=1
    cset(ws1, fila_tablas, c_idx,   tot_c,    bg=C_AZUL_CLA, sz=11, num='#,##0');   c_idx+=1
    cset(ws1, fila_tablas, c_idx,   censo_n,  bg=C_AZUL_CLA, sz=11, num='#,##0');   c_idx+=1
    cset(ws1, fila_tablas, c_idx,   int(row_s['cob_mar']), bg=C_AZUL_CLA, sz=11, num='#,##0'); c_idx+=1
    cset(ws1, fila_tablas, c_idx,   sol_m,    bg=C_AZUL_CLA, sz=11, num='#,##0');   c_idx+=1
    cset(ws1, fila_tablas, c_idx,   _pct(sol_m,tot_m), bg=C_AZUL_CLA, sz=11, num='0.0%'); c_idx+=1
    cset(ws1, fila_tablas, c_idx,   int(row_s['cob_abr']), bg=C_AZUL_CLA, sz=11, num='#,##0'); c_idx+=1
    cset(ws1, fila_tablas, c_idx,   sol_a,    bg=C_AZUL_CLA, sz=11, num='#,##0');   c_idx+=1
    cset(ws1, fila_tablas, c_idx,   _pct(sol_a,tot_a), bg=C_AZUL_CLA, sz=11, num='0.0%'); c_idx+=1
    if tiene_mayo:
        sol_my = float(row_s.get('sol_may',0))
        tot_my = float(row_s.get('sol_tot_may',0))
        cset(ws1, fila_tablas, c_idx, int(row_s.get('cob_may',0)), bg=C_AZUL_CLA, sz=11, num='#,##0'); c_idx+=1
        cset(ws1, fila_tablas, c_idx, sol_my,  bg=C_AMARILLO, sz=11, num='#,##0');  c_idx+=1
        cset(ws1, fila_tablas, c_idx, _pct(sol_my,tot_my), bg=C_AMARILLO, sz=11, num='0.0%'); c_idx+=1
    cset(ws1, fila_tablas, c_idx, no_cob_n,             bg=C_ROJO_CLA, fg=C_ROJO_T, sz=11, num='#,##0'); c_idx+=1
    cset(ws1, fila_tablas, c_idx, _pct(no_cob_n,censo_n), bg=C_ROJO_CLA, fg=C_ROJO_T, sz=11, num='0.0%')
    ws1.row_dimensions[fila_tablas].height = 16
    fila_tablas += 1

# Fila total general supervisores
c_idx = 2
tot_cob_mar   = int(sup_resumen['cob_mar'].sum())
tot_sol_mar   = float(sup_resumen['sol_mar'].sum())
tot_solT_mar  = float(sup_resumen['sol_tot_mar'].sum())
tot_cob_abr   = int(sup_resumen['cob_abr'].sum())
tot_sol_abr   = float(sup_resumen['sol_abr'].sum())
tot_solT_abr  = float(sup_resumen['sol_tot_abr'].sum())
tot_censo_n   = int(sup_resumen['cartera_censo'].sum())
tot_car_n     = int(sup_resumen['total_cartera'].sum())
tot_no_cob    = sin_cob   # match exacto con el funnel principal
cset(ws1, fila_tablas, c_idx, 'TOTAL GENERAL', bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, halign='left'); c_idx+=1
cset(ws1, fila_tablas, c_idx, tot_car_n,  bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='#,##0');  c_idx+=1
cset(ws1, fila_tablas, c_idx, tot_censo_n,bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='#,##0');  c_idx+=1
cset(ws1, fila_tablas, c_idx, tot_cob_mar,bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='#,##0');  c_idx+=1
cset(ws1, fila_tablas, c_idx, tot_sol_mar,bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='#,##0');  c_idx+=1
cset(ws1, fila_tablas, c_idx, _pct(tot_sol_mar,tot_solT_mar), bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='0.0%'); c_idx+=1
cset(ws1, fila_tablas, c_idx, tot_cob_abr,bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='#,##0');  c_idx+=1
cset(ws1, fila_tablas, c_idx, tot_sol_abr,bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='#,##0');  c_idx+=1
cset(ws1, fila_tablas, c_idx, _pct(tot_sol_abr,tot_solT_abr), bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='0.0%'); c_idx+=1
if tiene_mayo:
    tot_cob_may  = int(sup_resumen['cob_may'].sum())
    tot_sol_may  = float(sup_resumen['sol_may'].sum())
    tot_solT_may = float(sup_resumen['sol_tot_may'].sum())
    cset(ws1, fila_tablas, c_idx, tot_cob_may, bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='#,##0');  c_idx+=1
    cset(ws1, fila_tablas, c_idx, tot_sol_may, bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='#,##0');  c_idx+=1
    cset(ws1, fila_tablas, c_idx, _pct(tot_sol_may,tot_solT_may), bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='0.0%'); c_idx+=1
cset(ws1, fila_tablas, c_idx, tot_no_cob,   bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='#,##0'); c_idx+=1
cset(ws1, fila_tablas, c_idx, _pct(tot_no_cob,tot_censo_n), bold=True, fg=C_NARANJA_T, bg=C_NARANJA, sz=11, num='0.0%')
ws1.row_dimensions[fila_tablas].height = 18
fila_tablas += 2  # espacio

# ─── BLOQUE E: Resumen por vendedor ───────────────────────────────────
ws1.merge_cells(f'B{fila_tablas}:{get_column_letter(NCOLS_TOT)}{fila_tablas}')
c = ws1.cell(row=fila_tablas, column=2, value='RESUMEN POR VENDEDOR')
c.font      = fnt(bold=True, color=C_AZUL_MED_T, sz=14)
c.fill      = fill(C_AZUL_MED)
c.alignment = aln()
ws1.row_dimensions[fila_tablas].height = 28
fila_tablas += 1

hdrs_v = ['VENDEDOR','SUPERVISOR','RUTA','TOTAL\nCARTERA','CARTERA\nCENSO',
          'MAR-26\nCOB','MAR-26\nSOLES','MAR-26\n% SOLES',
          'ABR-26\nCOB','ABR-26\nSOLES','ABR-26\n% SOLES']
if tiene_mayo:
    hdrs_v += ['MAY-26\nCOB','MAY-26\nSOLES','MAY-26\n% SOLES']
hdrs_v += ['NO COB\n#','NO COB\n%']

for i, lbl in enumerate(hdrs_v):
    cset(ws1, fila_tablas, i+2, lbl, bold=True, fg=C_BLANCO, bg=C_AZUL_OSC, sz=10)
ws1.row_dimensions[fila_tablas].height = 32
fila_tablas += 1

for _, vr in det_vend.iterrows():
    censo_n  = int(vr['cartera_censo'])
    no_cob_n = int(vr['no_cob'])
    cob_m    = int(vr.get('cob_2603', 0))
    sol_m    = float(vr.get('2603', vr.get(2603, 0)))
    tot_m    = float(vr.get('sol_tot_2603', 0))
    cob_a    = int(vr.get('cob_2604', 0))
    sol_a    = float(vr.get('2604', vr.get(2604, 0)))
    tot_a    = float(vr.get('sol_tot_2604', 0))
    bg_row   = C_BLANCO if fila_tablas % 2 == 0 else C_GRIS
    c_idx = 2
    cset(ws1, fila_tablas, c_idx, vr['nom_vend'],  bg=bg_row, sz=11, halign='left'); c_idx+=1
    cset(ws1, fila_tablas, c_idx, vr['supervisor'],bg=bg_row, sz=11, halign='left'); c_idx+=1
    cset(ws1, fila_tablas, c_idx, vr['ruta'],      bg=bg_row, sz=11);              c_idx+=1
    cset(ws1, fila_tablas, c_idx, int(vr['total_cartera']), bg=bg_row, sz=11, num='#,##0'); c_idx+=1
    cset(ws1, fila_tablas, c_idx, censo_n,  bg=bg_row, sz=11, num='#,##0'); c_idx+=1
    cset(ws1, fila_tablas, c_idx, cob_m,    bg=bg_row, sz=11, num='#,##0'); c_idx+=1
    cset(ws1, fila_tablas, c_idx, sol_m,    bg=bg_row, sz=11, num='#,##0'); c_idx+=1
    cset(ws1, fila_tablas, c_idx, _pct(sol_m,tot_m), bg=bg_row, sz=11, num='0.0%'); c_idx+=1
    cset(ws1, fila_tablas, c_idx, cob_a,    bg=bg_row, sz=11, num='#,##0'); c_idx+=1
    cset(ws1, fila_tablas, c_idx, sol_a,    bg=bg_row, sz=11, num='#,##0'); c_idx+=1
    cset(ws1, fila_tablas, c_idx, _pct(sol_a,tot_a), bg=bg_row, sz=11, num='0.0%'); c_idx+=1
    if tiene_mayo:
        cob_my  = int(vr.get('cob_2605', 0))
        sol_my  = float(vr.get('2605', vr.get(2605, 0)))
        tot_my  = float(vr.get('sol_tot_2605', 0))
        cset(ws1, fila_tablas, c_idx, cob_my,  bg=C_AMARILLO, sz=11, num='#,##0'); c_idx+=1
        cset(ws1, fila_tablas, c_idx, sol_my,  bg=C_AMARILLO, sz=11, num='#,##0'); c_idx+=1
        cset(ws1, fila_tablas, c_idx, _pct(sol_my,tot_my), bg=C_AMARILLO, sz=11, num='0.0%'); c_idx+=1
    cset(ws1, fila_tablas, c_idx, no_cob_n,             bg=C_ROJO_CLA, fg=C_ROJO_T, sz=11, num='#,##0'); c_idx+=1
    cset(ws1, fila_tablas, c_idx, _pct(no_cob_n,censo_n), bg=C_ROJO_CLA, fg=C_ROJO_T, sz=11, num='0.0%')
    ws1.row_dimensions[fila_tablas].height = 16
    fila_tablas += 1


# ── HOJA 2: DETALLE_CLIENTE ───────────────────────────────────────────
ws2 = wb.create_sheet('DETALLE_CLIENTE')

ws2.column_dimensions['A'].width = 5.0    # FFVV
ws2.column_dimensions['B'].width = 22.0   # SUPERVISOR
ws2.column_dimensions['C'].width = 40.0   # VENDEDOR
ws2.column_dimensions['D'].width = 6.0    # RUTA
ws2.column_dimensions['E'].width = 40.0   # CLIENTE
ws2.column_dimensions['F'].width = 8.0    # DIA VISITA
ws2.column_dimensions['G'].width = 10.0   # CODIGO
ws2.column_dimensions['H'].width = 8.0    # N_CENSO
for cl in ('I','J','K','L','M'):
    ws2.column_dimensions[cl].width = 12.0

# Cabecera directa en fila 1 — sin columnas COB, agrega DIA VISITA después de CLIENTE
hdrs_det = ['FFVV','SUPERVISOR','VENDEDOR','RUTA','CLIENTE','DIA VISITA','CODIGO','N_CENSO',
            'MAR-26 SOLES','ABR-26 SOLES']
if tiene_mayo:
    hdrs_det += ['MAY-26 SOLES']
hdrs_det += ['NO COB']
NCOLS_DET = len(hdrs_det)
for col, lbl in enumerate(hdrs_det, 1):
    cset(ws2, 1, col, lbl, bold=True, fg=C_AZUL_MED_T, bg=C_AZUL_MED, sz=9)
ws2.row_dimensions[1].height = 20

# Precalcular set de clientes coberturados (cualquier mes post-reest)
cob_ids = set(df_censo_post['ccod_cli'].unique())

fila_det = 2
for _, cr in det_censo.iterrows():
    bg_row  = C_GRIS if fila_det % 2 == 0 else C_BLANCO
    es_cob  = int(cr['codigo']) in cob_ids
    cset(ws2, fila_det, 1, cr['ffvv'],           bg=bg_row, sz=9)
    cset(ws2, fila_det, 2, cr['supervisor'],      bg=bg_row, sz=9, halign='left')
    cset(ws2, fila_det, 3, cr['nom_vend'],        bg=bg_row, sz=9, halign='left')
    cset(ws2, fila_det, 4, cr['ruta'],            bg=bg_row, sz=9)
    cset(ws2, fila_det, 5, cr['cliente'],         bg=bg_row, sz=9, halign='left')
    cset(ws2, fila_det, 6, cr['diaVisita'],       bg=bg_row, sz=9)
    cset(ws2, fila_det, 7, int(cr['codigo']),     bg=bg_row, sz=9)
    cset(ws2, fila_det, 8, 1,                     bg=C_AZUL_CLA, sz=9)

    col = 9
    for m in PERIODOS_POST:
        val  = float(cr.get(m, 0))
        bg_m = C_VERDE if val > 0 else bg_row
        cset(ws2, fila_det, col, val if val > 0 else 0, bg=bg_m, sz=9, num='#,##0')
        col += 1

    # Columna NO COB: 0 si ya compro en algun mes, 1 si nunca
    bg_nc = C_ROJO_CLA if not es_cob else bg_row
    fg_nc = C_ROJO_T   if not es_cob else 'FF000000'
    cset(ws2, fila_det, col, 0 if es_cob else 1, bold=(not es_cob), fg=fg_nc, bg=bg_nc, sz=9)

    fila_det += 1

ws2.auto_filter.ref = f'A1:{get_column_letter(NCOLS_DET)}{fila_det-1}'
ws2.freeze_panes    = 'F2'


# ── HOJA 3: ULTIMA_COMPRA_CENSO ───────────────────────────────────────
ws3 = wb.create_sheet('ULTIMA_COMPRA_CENSO')

hdrs3 = ['SUPERVISOR','VENDEDOR','RUTA','CODIGO','CLIENTE','DIA VISITA',
         'GIRO','C/ COMPRA','ULTIMA CATEGORIA','ULT FECHA COMPRA','MONTO (S/)','VALIDACION VISITA','ULT FECHA VISITA','DIAS SIN COMPRA']
WCOLS3 = [17.43, 32.57, 7.0, 9.0, 30.0, 8.0, 20.0, 9.0, 18.0, 13.0, 12.0, 22.0, 13.0, 14.0]
for i, w in enumerate(WCOLS3, 1):
    ws3.column_dimensions[get_column_letter(i)].width = w

titulo(ws3, 1, f'ULTIMA COMPRA -- CLIENTES CENSO  |  Referencia: {FECHA_HOY_STR}', len(hdrs3))
for col, lbl in enumerate(hdrs3, 1):
    cset(ws3, 2, col, lbl, bold=True, fg=C_AZUL_MED_T, bg=C_AZUL_MED, sz=9)
ws3.row_dimensions[2].height = 22

row3 = 3
for _, cr in ult_censo_enr.iterrows():
    dias     = int(cr['dias'])
    sin_comp = cr['cat_ultima'] == 'SIN COMPRA'
    if sin_comp:
        bg_row = C_ROJO_CLA; bg_dias = C_ROJO_CLA; fg_dias = C_ROJO_T
    elif dias <= 30:
        bg_row = C_BLANCO if row3 % 2 else C_GRIS
        bg_dias = C_VERDE;   fg_dias = C_VERDE_T
    elif dias <= 60:
        bg_row = C_BLANCO if row3 % 2 else C_GRIS
        bg_dias = C_AMARILLO; fg_dias = C_AMARILLO_T
    else:
        bg_row = C_BLANCO if row3 % 2 else C_GRIS
        bg_dias = C_ROJO_CLA; fg_dias = C_ROJO_T

    fecha_str = cr['fecha_ultima'].strftime('%d/%m/%Y') if pd.notna(cr.get('fecha_ultima')) else ''
    cod3      = int(cr['codigo'])
    val_vis3  = enc_validacion.get(cod3, '')
    if pd.isna(val_vis3): val_vis3 = ''
    fec_vis3_raw = enc_fecha.get(cod3, None)
    fec_vis3  = fec_vis3_raw.strftime('%d/%m/%Y') if pd.notna(fec_vis3_raw) else ''
    cat_txt   = '' if sin_comp else cr['cat_ultima']   # vacío en vez de 'SIN COMPRA'
    compra_si = 'NO' if sin_comp else 'SI'
    bg_compra = C_ROJO_CLA if sin_comp else C_VERDE
    fg_compra = C_ROJO_T   if sin_comp else C_VERDE_T
    cset(ws3, row3, 1,  cr['supervisor'],  bg=bg_row, sz=9, halign='left')
    cset(ws3, row3, 2,  cr['nom_vend'],    bg=bg_row, sz=9, halign='left')
    cset(ws3, row3, 3,  cr['ruta'],        bg=bg_row, sz=9)
    cset(ws3, row3, 4,  cod3,              bg=bg_row, sz=9)
    cset(ws3, row3, 5,  cr['cliente'],     bg=bg_row, sz=9, halign='left')
    cset(ws3, row3, 6,  cr['diaVisita'],   bg=bg_row, sz=9)
    cset(ws3, row3, 7,  cr['giro'],        bg=bg_row, sz=9, halign='left')
    cset(ws3, row3, 8,  compra_si,         bg=bg_compra, fg=fg_compra, bold=True, sz=9)
    cset(ws3, row3, 9,  cat_txt,           bg=bg_row, sz=9)
    cset(ws3, row3, 10, fecha_str,         bg=bg_row, sz=9)
    cset(ws3, row3, 11, float(cr['monto_ultima']), bg=bg_row, sz=9, num='#,##0')
    cset(ws3, row3, 12, val_vis3,          bg=bg_row, sz=9, halign='left')
    cset(ws3, row3, 13, fec_vis3,          bg=bg_row, sz=9)
    cset(ws3, row3, 14, dias if not sin_comp else '',
         bg=bg_dias, fg=fg_dias, bold=(not sin_comp), sz=9)
    row3 += 1

ws3.auto_filter.ref = f'A2:{get_column_letter(len(hdrs3))}{row3-1}'
ws3.freeze_panes    = 'A3'

row3 += 1
nota3 = (f'Ordenado por dias sin compra (descendente). '
         f'Verde <=30 dias | Amarillo 31-60 dias | Rojo >60 dias | SIN COMPRA = sin actividad post 08/03/2026. '
         f'Fecha de referencia: {FECHA_HOY_STR}.')
ws3.merge_cells(f'A{row3}:{get_column_letter(len(hdrs3))}{row3}')
ws3.cell(row=row3, column=1, value=nota3).font = fnt(sz=8)


# ── HOJA 4: ULTIMA_COMPRA_NO_CENSO ────────────────────────────────────
ws4 = wb.create_sheet('ULTIMA_COMPRA_NO_CENSO')

_nc_mes_hdrs = [NC_MES_LABELS.get(m, str(m)) for m in MESES_NC]
hdrs4 = ['SUPERVISOR','VENDEDOR','RUTA','CODIGO','CLIENTE','DIA VISITA',
         'GIRO','ULTIMA CATEGORIA','ULT FECHA COMPRA','MONTO (S/)'] + _nc_mes_hdrs + ['VALIDACION VISITA','ULT FECHA VISITA','DIAS SIN COMPRA']
WCOLS4_BASE = [17.43, 32.57, 7.0, 9.0, 30.0, 8.0, 20.0, 18.0, 13.0, 12.0]
WCOLS4_MES  = [11.0] * len(MESES_NC)
WCOLS4_FIN  = [22.0, 13.0, 14.0]
WCOLS4 = WCOLS4_BASE + WCOLS4_MES + WCOLS4_FIN
for i, w in enumerate(WCOLS4, 1):
    ws4.column_dimensions[get_column_letter(i)].width = w

titulo(ws4, 1, f'ULTIMA COMPRA -- CLIENTES NO CENSO  |  Referencia: {FECHA_HOY_STR}', len(hdrs4))
for col, lbl in enumerate(hdrs4, 1):
    cset(ws4, 2, col, lbl, bold=True, fg=C_AZUL_MED_T, bg=C_AZUL_MED, sz=9)
ws4.row_dimensions[2].height = 22

with_sales = set(df['ccod_cli'].unique())
ult_nc_filt = ult_no_censo_enr[
    ult_no_censo_enr['codigo'].isin(with_sales) |
    (ult_no_censo_enr['cat_ultima'] != 'SIN COMPRA')
]

row4 = 3
for _, cr in ult_nc_filt.iterrows():
    dias     = int(cr['dias'])
    sin_comp = cr['cat_ultima'] == 'SIN COMPRA'
    if sin_comp:
        bg_dias = C_ROJO_CLA; fg_dias = C_ROJO_T
    elif dias <= 30:
        bg_dias = C_VERDE;    fg_dias = C_VERDE_T
    elif dias <= 60:
        bg_dias = C_AMARILLO; fg_dias = C_AMARILLO_T
    else:
        bg_dias = C_ROJO_CLA; fg_dias = C_ROJO_T

    bg_row    = C_BLANCO if row4 % 2 else C_GRIS
    fecha_str = cr['fecha_ultima'].strftime('%d/%m/%Y') if pd.notna(cr.get('fecha_ultima')) else ''
    cod4      = int(cr['codigo'])
    val_vis4  = enc_validacion.get(cod4, '')
    if pd.isna(val_vis4): val_vis4 = ''
    fec_vis4_raw = enc_fecha.get(cod4, None)
    fec_vis4  = fec_vis4_raw.strftime('%d/%m/%Y') if pd.notna(fec_vis4_raw) else ''
    cset(ws4, row4, 1,  cr['supervisor'],  bg=bg_row, sz=9, halign='left')
    cset(ws4, row4, 2,  cr['nom_vend'],    bg=bg_row, sz=9, halign='left')
    cset(ws4, row4, 3,  cr['ruta'],        bg=bg_row, sz=9)
    cset(ws4, row4, 4,  cod4,              bg=bg_row, sz=9)
    cset(ws4, row4, 5,  cr['cliente'],     bg=bg_row, sz=9, halign='left')
    cset(ws4, row4, 6,  cr['diaVisita'],   bg=bg_row, sz=9)
    cset(ws4, row4, 7,  cr['giro'],        bg=bg_row, sz=9, halign='left')
    cset(ws4, row4, 8,  cr['cat_ultima'],  bg=bg_row, sz=9)
    cset(ws4, row4, 9,  fecha_str,         bg=bg_row, sz=9)
    cset(ws4, row4, 10, float(cr['monto_ultima']), bg=bg_row, sz=9, num='#,##0')
    col4 = 11
    for m in MESES_NC:
        sol_m4 = float(sol_nc_mes.loc[cod4, m]) if cod4 in sol_nc_mes.index and m in sol_nc_mes.columns else 0.0
        bg_m4  = C_VERDE if sol_m4 > 0 else bg_row
        cset(ws4, row4, col4, sol_m4 if sol_m4 > 0 else '', bg=bg_m4, sz=9, num='#,##0')
        col4 += 1
    cset(ws4, row4, col4,   val_vis4, bg=bg_row, sz=9, halign='left'); col4+=1
    cset(ws4, row4, col4,   fec_vis4, bg=bg_row, sz=9);                col4+=1
    cset(ws4, row4, col4,   dias if not sin_comp else 'SIN COMPRA',
         bg=bg_dias, fg=fg_dias, bold=True, sz=9)
    row4 += 1

ws4.auto_filter.ref = f'A2:{get_column_letter(len(hdrs4))}{row4-1}'
ws4.freeze_panes    = 'A3'

row4 += 1
nota4 = (f'Solo clientes no-censo con actividad en el historico disponible. '
         f'Verde <=30 dias | Amarillo 31-60 dias | Rojo >60 dias. '
         f'Fecha de referencia: {FECHA_HOY_STR}.')
ws4.merge_cells(f'A{row4}:{get_column_letter(len(hdrs4))}{row4}')
ws4.cell(row=row4, column=1, value=nota4).font = fnt(sz=8)


# ════════════════════════════════════════════════════════════════════════
# HOJA 5: HISTORICO POR SUPERVISOR Y VENDEDOR
# ════════════════════════════════════════════════════════════════════════
ws5 = wb.create_sheet('HISTORICO')

# Cruzar df completo con maestro para obtener supervisor y ruta actuales
_lc_map = lc[['codigo','supervisor','ruta','vendedor']].copy()
_lc_map.columns = ['ccod_cli','supervisor','ruta','nom_vend']

df_hist_full = df.merge(_lc_map, on='ccod_cli', how='left')
df_hist_full = df_hist_full[df_hist_full['supervisor'].notna() & (df_hist_full['supervisor'] != '')]

MESES_HIST = sorted(df_hist_full['mes'].unique())
MESES_LABELS = {
    2511:'Nov-25', 2512:'Dic-25', 2601:'Ene-26', 2602:'Feb-26',
    2603:'Mar-26', 2604:'Abr-26', 2605:'May-26',
}

def _hist_agg(df_in, grupo):
    rows = []
    grupos_vals = df_in[grupo].dropna().unique()
    for g in sorted(grupos_vals):
        sub = df_in[df_in[grupo] == g]
        row = {grupo: g}
        for m in MESES_HIST:
            sm = sub[sub['mes'] == m]
            row[f'cob_{m}']  = sm['ccod_cli'].nunique()
            row[f'sol_{m}']  = sm['total_monto'].sum()
        rows.append(row)
    return pd.DataFrame(rows)

hist_sup  = _hist_agg(df_hist_full, 'supervisor')
hist_vend = _hist_agg(df_hist_full[['ccod_cli','mes','total_monto','ruta','supervisor','nom_vend']]
                      .rename(columns={'nom_vend':'_vend'}), '_vend')

# Recalcular vendedor con más campos
def _hist_vend_agg(df_in):
    rows = []
    for (sup, ruta, vend), sub in df_in.groupby(['supervisor','ruta','nom_vend']):
        row = {'supervisor': sup, 'ruta': ruta, 'nom_vend': vend}
        for m in MESES_HIST:
            sm = sub[sub['mes'] == m]
            row[f'cob_{m}'] = sm['ccod_cli'].nunique()
            row[f'sol_{m}'] = sm['total_monto'].sum()
        rows.append(row)
    return pd.DataFrame(rows)

hist_vend = _hist_vend_agg(df_hist_full)

# ── layout hoja HISTORICO ────────────────────────────────────────────
NCM = len(MESES_HIST)   # número de meses

def _escribir_bloque_hist(ws, fila_ini, df_bloque, id_cols, id_labels, id_widths, titulo_txt, col_ini=1):
    # ncols: id_cols + por mes (COB+SOLES+VAR%) + proyección mayo si aplica
    _extra = 1 if FACTOR_PROY_MAY and 2605 in MESES_HIST else 0
    ncols_total = len(id_cols) + NCM * 3 + _extra
    # Título fusionado desde col_ini
    col_fin_titulo = col_ini + ncols_total - 1
    ws.merge_cells(f'{get_column_letter(col_ini)}{fila_ini}:{get_column_letter(col_fin_titulo)}{fila_ini}')
    ct = ws[f'{get_column_letter(col_ini)}{fila_ini}']
    ct.value = titulo_txt
    ct.font  = fnt(bold=True, color=C_BLANCO, sz=12)
    ct.fill  = fill(C_AZUL_OSC)
    ct.alignment = aln()
    ws.row_dimensions[fila_ini].height = 24

    fila_hdr = fila_ini + 1
    col = col_ini
    for lbl in id_labels:
        cset(ws, fila_hdr, col, lbl, bold=True, fg=C_BLANCO, bg=C_AZUL_OSC, sz=9)
        col += 1
    for i, m in enumerate(MESES_HIST):
        lbl_m = MESES_LABELS.get(m, str(m))
        cset(ws, fila_hdr, col,   f'{lbl_m}\nCOB',   bold=True, fg=C_AZUL_MED_T, bg=C_AZUL_MED, sz=9)
        cset(ws, fila_hdr, col+1, f'{lbl_m}\nSOLES', bold=True, fg=C_AZUL_MED_T, bg=C_AZUL_MED, sz=9)
        if i > 0:
            cset(ws, fila_hdr, col+2, f'{lbl_m}\nVAR%', bold=True, fg=C_BLANCO, bg=C_AZUL_OSC, sz=9)
        else:
            cset(ws, fila_hdr, col+2, '', bold=False, bg=C_AZUL_OSC, sz=9)
        col += 3
        if m == 2605 and FACTOR_PROY_MAY:
            cset(ws, fila_hdr, col, 'MAY-26\nPROY', bold=True, fg=C_BLANCO, bg=C_AMARILLO_T, sz=9)
            col += 1
    ws.row_dimensions[fila_hdr].height = 24

    fila = fila_hdr + 1
    for _, r in df_bloque.iterrows():
        bg_row = C_BLANCO if fila % 2 else C_GRIS
        col = col_ini
        for ic in id_cols:
            cset(ws, fila, col, r.get(ic, ''), bg=bg_row, sz=9, halign='left')
            col += 1
        for i, m in enumerate(MESES_HIST):
            cob  = int(r.get(f'cob_{m}', 0))
            sol  = float(r.get(f'sol_{m}', 0))
            # Para mayo: VAR% se calcula con el proyectado vs mes anterior real
            sol_para_var = (sol * FACTOR_PROY_MAY) if (m == 2605 and FACTOR_PROY_MAY) else sol
            if i > 0:
                m_prev   = MESES_HIST[i-1]
                sol_prev = float(r.get(f'sol_{m_prev}', 0))
                var = (sol_para_var - sol_prev) / sol_prev if sol_prev else None
            else:
                var = None
            cset(ws, fila, col,   cob if cob else '', bg=bg_row, sz=9, num='#,##0')
            cset(ws, fila, col+1, sol if sol else '', bg=bg_row, sz=9, num='#,##0')
            if var is not None:
                bg_v = C_VERDE if var >= 0 else C_ROJO_CLA
                fg_v = C_VERDE_T if var >= 0 else C_ROJO_T
                cset(ws, fila, col+2, var, bg=bg_v, fg=fg_v, sz=9, num='0.0%', bold=True)
            else:
                cset(ws, fila, col+2, '', bg=bg_row, sz=9)
            col += 3
            if m == 2605 and FACTOR_PROY_MAY:
                proy = sol * FACTOR_PROY_MAY
                cset(ws, fila, col, proy if proy else '', bg=C_AMARILLO, sz=9, num='#,##0')
                col += 1
        fila += 1

    # Fila TOTAL
    bg_tot = C_NARANJA
    col = col_ini
    for idx_ic, ic in enumerate(id_cols):
        cset(ws, fila, col, 'TOTAL' if idx_ic == 0 else '', bold=True, bg=bg_tot, fg=C_NARANJA_T, sz=9)
        col += 1
    for i, m in enumerate(MESES_HIST):
        sol_tot = float(df_bloque[f'sol_{m}'].sum())
        cob_tot = int((df_bloque[f'cob_{m}'] > 0).sum())
        sol_tot_para_var = (sol_tot * FACTOR_PROY_MAY) if (m == 2605 and FACTOR_PROY_MAY) else sol_tot
        if i > 0:
            m_prev   = MESES_HIST[i-1]
            sol_prev = float(df_bloque[f'sol_{m_prev}'].sum())
            var_tot  = (sol_tot_para_var - sol_prev) / sol_prev if sol_prev else None
        else:
            var_tot = None
        cset(ws, fila, col,   cob_tot if cob_tot else '', bold=True, bg=bg_tot, fg=C_NARANJA_T, sz=9, num='#,##0')
        cset(ws, fila, col+1, sol_tot if sol_tot else '', bold=True, bg=bg_tot, fg=C_NARANJA_T, sz=9, num='#,##0')
        if var_tot is not None:
            cset(ws, fila, col+2, var_tot, bold=True, bg=bg_tot, fg=C_NARANJA_T, sz=9, num='0.0%')
        else:
            cset(ws, fila, col+2, '', bg=bg_tot, sz=9)
        col += 3
        if m == 2605 and FACTOR_PROY_MAY:
            proy_tot = sol_tot * FACTOR_PROY_MAY
            cset(ws, fila, col, proy_tot, bold=True, bg=bg_tot, fg=C_NARANJA_T, sz=9, num='#,##0')
            col += 1
    fila += 1

    return fila + 1

# Anchos fijos de columna para HISTORICO
# Sup: empieza en col C (3) → C=supervisor, desde F datos de mes
# Vend: empieza en col A (1) → A=supervisor, B=ruta, C=vendedor, desde D datos de mes
_extra_col = 1 if FACTOR_PROY_MAY and 2605 in MESES_HIST else 0
COL_INI_SUP  = 3   # supervisores en col C
COL_INI_VEND = 1   # vendedores en col A

ws5.column_dimensions['A'].width = 32   # supervisor (bloque vendedor)
ws5.column_dimensions['B'].width = 7    # ruta (bloque vendedor)
ws5.column_dimensions['C'].width = 35   # vendedor / supervisor (bloque sup)
# Desde col D (4): datos de mes para bloque vendedor (3 id cols → datos desde col 4)
# Desde col F (6): datos de mes para bloque supervisor (col C=supervisor → datos desde col 4 también)
# Usamos el máximo: bloque vend col_ini=1+3=4, bloque sup col_ini=3+1=4 → ambos desde col 4
_inicio_datos = 4
ncols_datos   = NCM * 3 + _extra_col
for i in range(_inicio_datos, _inicio_datos + ncols_datos + 1):
    cl = get_column_letter(i)
    idx_within = (i - _inicio_datos) % 3
    ws5.column_dimensions[cl].width = [9, 12, 8][idx_within]
    if _extra_col and i == _inicio_datos + ncols_datos - 1:
        ws5.column_dimensions[cl].width = 12

fila_cur = 1
fila_cur = _escribir_bloque_hist(
    ws5, fila_cur,
    hist_sup,
    id_cols   = ['supervisor'],
    id_labels = ['SUPERVISOR'],
    id_widths = [32],
    titulo_txt= f'HISTORICO VENTAS POR SUPERVISOR  |  {MESES_LABELS.get(MESES_HIST[0],"?")} → {MESES_LABELS.get(MESES_HIST[-1],"?")}',
    col_ini   = COL_INI_SUP,
)

fila_cur = _escribir_bloque_hist(
    ws5, fila_cur,
    hist_vend,
    id_cols   = ['supervisor','ruta','nom_vend'],
    id_labels = ['SUPERVISOR','RUTA','VENDEDOR'],
    id_widths = [32, 7, 35],
    titulo_txt= f'HISTORICO VENTAS POR VENDEDOR  |  {MESES_LABELS.get(MESES_HIST[0],"?")} → {MESES_LABELS.get(MESES_HIST[-1],"?")}',
    col_ini   = COL_INI_VEND,
)

ws5.freeze_panes = 'D3'

# ════════════════════════════════════════════════════════════════════════
# GUARDAR LIBRO PRINCIPAL
# ════════════════════════════════════════════════════════════════════════
wb.save(OUTPUT)
print(f"\nGuardado: {OUTPUT}")


# ════════════════════════════════════════════════════════════════════════
# FILETEADO — un libro por supervisor (4 hojas, sin FUNNEL_CENSO)
# ════════════════════════════════════════════════════════════════════════
def _escribir_libro_supervisor(sup_nombre):
    """Genera un workbook con las 4 hojas filtradas para un supervisor."""

    # ── Filtros de datos ──────────────────────────────────────────────
    det_s      = det_censo[det_censo['supervisor'] == sup_nombre].copy()
    ult_c_s    = ult_censo_enr[ult_censo_enr['supervisor'] == sup_nombre].copy()
    ult_nc_s   = ult_no_censo_enr[ult_no_censo_enr['supervisor'] == sup_nombre].copy()
    nc_filt_s  = ult_nc_filt[ult_nc_filt['supervisor'] == sup_nombre].copy()
    hist_full_s= df_hist_full[df_hist_full['supervisor'] == sup_nombre].copy()

    if det_s.empty and ult_c_s.empty and ult_nc_s.empty:
        return  # supervisor sin datos — saltar

    wb_s = Workbook()
    wb_s.remove(wb_s.active)

    # ── HOJA 1: DETALLE_CLIENTE ───────────────────────────────────────
    ws = wb_s.create_sheet('DETALLE_CLIENTE')
    ws.column_dimensions['A'].width = 5.0
    ws.column_dimensions['B'].width = 22.0
    ws.column_dimensions['C'].width = 40.0
    ws.column_dimensions['D'].width = 6.0
    ws.column_dimensions['E'].width = 40.0
    ws.column_dimensions['F'].width = 8.0
    ws.column_dimensions['G'].width = 10.0
    ws.column_dimensions['H'].width = 8.0
    for cl in ('I','J','K','L','M'):
        ws.column_dimensions[cl].width = 12.0

    hdrs = ['FFVV','SUPERVISOR','VENDEDOR','RUTA','CLIENTE','DIA VISITA','CODIGO','N_CENSO',
            'MAR-26 SOLES','ABR-26 SOLES']
    if tiene_mayo:
        hdrs += ['MAY-26 SOLES']
    hdrs += ['NO COB']
    for col, lbl in enumerate(hdrs, 1):
        cset(ws, 1, col, lbl, bold=True, fg=C_AZUL_MED_T, bg=C_AZUL_MED, sz=9)
    ws.row_dimensions[1].height = 20

    cob_ids_s = set(df_censo_post[df_censo_post['ccod_cli'].isin(det_s['codigo'])]['ccod_cli'])
    fila = 2
    for _, cr in det_s.iterrows():
        bg_row = C_GRIS if fila % 2 == 0 else C_BLANCO
        es_cob = int(cr['codigo']) in cob_ids_s
        cset(ws, fila, 1, cr['ffvv'],       bg=bg_row, sz=9)
        cset(ws, fila, 2, cr['supervisor'], bg=bg_row, sz=9, halign='left')
        cset(ws, fila, 3, cr['nom_vend'],   bg=bg_row, sz=9, halign='left')
        cset(ws, fila, 4, cr['ruta'],       bg=bg_row, sz=9)
        cset(ws, fila, 5, cr['cliente'],    bg=bg_row, sz=9, halign='left')
        cset(ws, fila, 6, cr['diaVisita'],  bg=bg_row, sz=9)
        cset(ws, fila, 7, int(cr['codigo']),bg=bg_row, sz=9)
        cset(ws, fila, 8, 1,               bg=C_AZUL_CLA, sz=9)
        col = 9
        for m in PERIODOS_POST:
            val = float(cr.get(m, 0))
            cset(ws, fila, col, val if val > 0 else 0, bg=C_VERDE if val > 0 else bg_row, sz=9, num='#,##0')
            col += 1
        if tiene_mayo:
            val = float(cr.get(2605, 0))
            cset(ws, fila, col, val if val > 0 else 0, bg=C_VERDE if val > 0 else bg_row, sz=9, num='#,##0')
            col += 1
        bg_nc = C_ROJO_CLA if not es_cob else bg_row
        fg_nc = C_ROJO_T   if not es_cob else 'FF000000'
        cset(ws, fila, col, 0 if es_cob else 1, bold=(not es_cob), fg=fg_nc, bg=bg_nc, sz=9)
        fila += 1

    ws.auto_filter.ref = f'A1:{get_column_letter(len(hdrs))}{fila-1}'
    ws.freeze_panes    = 'F2'

    # ── HOJA 2: ULTIMA_COMPRA_CENSO ──────────────────────────────────
    ws = wb_s.create_sheet('ULTIMA_COMPRA_CENSO')
    hdrs3 = ['SUPERVISOR','VENDEDOR','RUTA','CODIGO','CLIENTE','DIA VISITA',
             'GIRO','C/ COMPRA','ULTIMA CATEGORIA','ULT FECHA COMPRA','MONTO (S/)',
             'VALIDACION VISITA','ULT FECHA VISITA','DIAS SIN COMPRA']
    WCOLS3 = [17.43,32.57,7.0,9.0,30.0,8.0,20.0,9.0,18.0,13.0,12.0,22.0,13.0,14.0]
    for i, w in enumerate(WCOLS3, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    titulo(ws, 1, f'ULTIMA COMPRA -- CLIENTES CENSO  |  Referencia: {FECHA_HOY_STR}', len(hdrs3))
    for col, lbl in enumerate(hdrs3, 1):
        cset(ws, 2, col, lbl, bold=True, fg=C_AZUL_MED_T, bg=C_AZUL_MED, sz=9)
    ws.row_dimensions[2].height = 22

    row = 3
    for _, cr in ult_c_s.iterrows():
        dias     = int(cr['dias'])
        sin_comp = cr['cat_ultima'] == 'SIN COMPRA'
        if sin_comp:
            bg_row = C_ROJO_CLA; bg_dias = C_ROJO_CLA; fg_dias = C_ROJO_T
        elif dias <= 30:
            bg_row = C_BLANCO if row % 2 else C_GRIS
            bg_dias = C_VERDE;    fg_dias = C_VERDE_T
        elif dias <= 60:
            bg_row = C_BLANCO if row % 2 else C_GRIS
            bg_dias = C_AMARILLO; fg_dias = C_AMARILLO_T
        else:
            bg_row = C_BLANCO if row % 2 else C_GRIS
            bg_dias = C_ROJO_CLA; fg_dias = C_ROJO_T
        fecha_str = cr['fecha_ultima'].strftime('%d/%m/%Y') if pd.notna(cr.get('fecha_ultima')) else ''
        cod       = int(cr['codigo'])
        val_vis   = enc_validacion.get(cod, ''); val_vis = '' if pd.isna(val_vis) else val_vis
        fec_vis_r = enc_fecha.get(cod, None)
        fec_vis   = fec_vis_r.strftime('%d/%m/%Y') if pd.notna(fec_vis_r) else ''
        cat_txt   = '' if sin_comp else cr['cat_ultima']
        compra_si = 'NO' if sin_comp else 'SI'
        bg_comp   = C_ROJO_CLA if sin_comp else C_VERDE
        fg_comp   = C_ROJO_T   if sin_comp else C_VERDE_T
        cset(ws, row,  1, cr['supervisor'],          bg=bg_row, sz=9, halign='left')
        cset(ws, row,  2, cr['nom_vend'],            bg=bg_row, sz=9, halign='left')
        cset(ws, row,  3, cr['ruta'],                bg=bg_row, sz=9)
        cset(ws, row,  4, cod,                       bg=bg_row, sz=9)
        cset(ws, row,  5, cr['cliente'],             bg=bg_row, sz=9, halign='left')
        cset(ws, row,  6, cr['diaVisita'],           bg=bg_row, sz=9)
        cset(ws, row,  7, cr['giro'],                bg=bg_row, sz=9, halign='left')
        cset(ws, row,  8, compra_si,                 bg=bg_comp, fg=fg_comp, bold=True, sz=9)
        cset(ws, row,  9, cat_txt,                   bg=bg_row, sz=9)
        cset(ws, row, 10, fecha_str,                 bg=bg_row, sz=9)
        cset(ws, row, 11, float(cr['monto_ultima']), bg=bg_row, sz=9, num='#,##0')
        cset(ws, row, 12, val_vis,                   bg=bg_row, sz=9, halign='left')
        cset(ws, row, 13, fec_vis,                   bg=bg_row, sz=9)
        cset(ws, row, 14, dias if not sin_comp else '', bg=bg_dias, fg=fg_dias, bold=(not sin_comp), sz=9)
        row += 1
    ws.auto_filter.ref = f'A2:{get_column_letter(len(hdrs3))}{row-1}'
    ws.freeze_panes    = 'A3'

    # ── HOJA 3: ULTIMA_COMPRA_NO_CENSO ───────────────────────────────
    ws = wb_s.create_sheet('ULTIMA_COMPRA_NO_CENSO')
    _nc_mes_hdrs = [NC_MES_LABELS.get(m, str(m)) for m in MESES_NC]
    hdrs4 = ['SUPERVISOR','VENDEDOR','RUTA','CODIGO','CLIENTE','DIA VISITA',
             'GIRO','ULTIMA CATEGORIA','ULT FECHA COMPRA','MONTO (S/)'] + _nc_mes_hdrs + \
            ['VALIDACION VISITA','ULT FECHA VISITA','DIAS SIN COMPRA']
    WCOLS4 = [17.43,32.57,7.0,9.0,30.0,8.0,20.0,18.0,13.0,12.0] + \
             [11.0]*len(MESES_NC) + [22.0,13.0,14.0]
    for i, w in enumerate(WCOLS4, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    titulo(ws, 1, f'ULTIMA COMPRA -- CLIENTES NO CENSO  |  Referencia: {FECHA_HOY_STR}', len(hdrs4))
    for col, lbl in enumerate(hdrs4, 1):
        cset(ws, 2, col, lbl, bold=True, fg=C_AZUL_MED_T, bg=C_AZUL_MED, sz=9)
    ws.row_dimensions[2].height = 22

    row = 3
    for _, cr in nc_filt_s.iterrows():
        dias     = int(cr['dias'])
        sin_comp = cr['cat_ultima'] == 'SIN COMPRA'
        if sin_comp:
            bg_dias = C_ROJO_CLA; fg_dias = C_ROJO_T
        elif dias <= 30:
            bg_dias = C_VERDE;    fg_dias = C_VERDE_T
        elif dias <= 60:
            bg_dias = C_AMARILLO; fg_dias = C_AMARILLO_T
        else:
            bg_dias = C_ROJO_CLA; fg_dias = C_ROJO_T
        bg_row    = C_BLANCO if row % 2 else C_GRIS
        fecha_str = cr['fecha_ultima'].strftime('%d/%m/%Y') if pd.notna(cr.get('fecha_ultima')) else ''
        cod       = int(cr['codigo'])
        val_vis   = enc_validacion.get(cod, ''); val_vis = '' if pd.isna(val_vis) else val_vis
        fec_vis_r = enc_fecha.get(cod, None)
        fec_vis   = fec_vis_r.strftime('%d/%m/%Y') if pd.notna(fec_vis_r) else ''
        cset(ws, row,  1, cr['supervisor'],          bg=bg_row, sz=9, halign='left')
        cset(ws, row,  2, cr['nom_vend'],            bg=bg_row, sz=9, halign='left')
        cset(ws, row,  3, cr['ruta'],                bg=bg_row, sz=9)
        cset(ws, row,  4, cod,                       bg=bg_row, sz=9)
        cset(ws, row,  5, cr['cliente'],             bg=bg_row, sz=9, halign='left')
        cset(ws, row,  6, cr['diaVisita'],           bg=bg_row, sz=9)
        cset(ws, row,  7, cr['giro'],                bg=bg_row, sz=9, halign='left')
        cset(ws, row,  8, cr['cat_ultima'],          bg=bg_row, sz=9)
        cset(ws, row,  9, fecha_str,                 bg=bg_row, sz=9)
        cset(ws, row, 10, float(cr['monto_ultima']), bg=bg_row, sz=9, num='#,##0')
        col4 = 11
        for m in MESES_NC:
            sol_m = float(sol_nc_mes.loc[cod, m]) if cod in sol_nc_mes.index and m in sol_nc_mes.columns else 0.0
            cset(ws, row, col4, sol_m if sol_m > 0 else '', bg=C_VERDE if sol_m > 0 else bg_row, sz=9, num='#,##0')
            col4 += 1
        cset(ws, row, col4,   val_vis, bg=bg_row, sz=9, halign='left'); col4+=1
        cset(ws, row, col4,   fec_vis, bg=bg_row, sz=9);                col4+=1
        cset(ws, row, col4,   dias if not sin_comp else 'SIN COMPRA', bg=bg_dias, fg=fg_dias, bold=True, sz=9)
        row += 1
    ws.auto_filter.ref = f'A2:{get_column_letter(len(hdrs4))}{row-1}'
    ws.freeze_panes    = 'A3'

    # ── HOJA 4: HISTORICO ────────────────────────────────────────────
    ws = wb_s.create_sheet('HISTORICO')
    hist_sup_s  = _hist_agg(hist_full_s, 'supervisor')
    hist_vend_s = _hist_vend_agg(hist_full_s)

    ws.column_dimensions['A'].width = 32
    ws.column_dimensions['B'].width = 7
    ws.column_dimensions['C'].width = 35
    _ini_d = 4
    _nd    = NCM * 3 + _extra_col
    for i in range(_ini_d, _ini_d + _nd + 1):
        cl = get_column_letter(i)
        ws.column_dimensions[cl].width = [9, 12, 8][(i - _ini_d) % 3]
        if _extra_col and i == _ini_d + _nd - 1:
            ws.column_dimensions[cl].width = 12

    fila_s = 1
    fila_s = _escribir_bloque_hist(ws, fila_s, hist_sup_s,
        id_cols=['supervisor'], id_labels=['SUPERVISOR'], id_widths=[32],
        titulo_txt=f'HISTORICO VENTAS POR SUPERVISOR  |  {MESES_LABELS.get(MESES_HIST[0],"?")} → {MESES_LABELS.get(MESES_HIST[-1],"?")}',
        col_ini=COL_INI_SUP)
    fila_s = _escribir_bloque_hist(ws, fila_s, hist_vend_s,
        id_cols=['supervisor','ruta','nom_vend'], id_labels=['SUPERVISOR','RUTA','VENDEDOR'], id_widths=[32,7,35],
        titulo_txt=f'HISTORICO VENTAS POR VENDEDOR  |  {MESES_LABELS.get(MESES_HIST[0],"?")} → {MESES_LABELS.get(MESES_HIST[-1],"?")}',
        col_ini=COL_INI_VEND)
    ws.freeze_panes = 'D3'

    # ── Guardar ───────────────────────────────────────────────────────
    # Nombre seguro: reemplazar caracteres inválidos en nombres de archivo
    nombre_safe = sup_nombre.replace('/', '-').replace('\\', '-').replace(':', '-')
    out_sup = os.path.join(os.path.dirname(OUTPUT), f'funnel_censo_{nombre_safe}.xlsx')
    wb_s.save(out_sup)
    return out_sup


# Obtener lista de supervisores únicos (del maestro de clientes)
supervisores_unicos = sorted(lc['supervisor'].dropna().unique())
supervisores_unicos = [s for s in supervisores_unicos if str(s).strip() != '']

print(f"\n[8] Generando libros por supervisor ({len(supervisores_unicos)} supervisores)...")
dir_output = os.path.dirname(OUTPUT)
for sup in supervisores_unicos:
    ruta_sup = _escribir_libro_supervisor(sup)
    if ruta_sup:
        print(f"   OK {os.path.basename(ruta_sup)}")
print(f"   Fileteado completado.")


# ════════════════════════════════════════════════════════════════════════
# GUARDAR (print final)
print(f"{'='*55}")
print(f"  Cartera total:           {len(lc):,}")
print(f"  Clientes censo:          {len(lc_censo):,}")
print(f"  Clientes no-censo:       {len(lc_no_censo):,}")
print(f"  Coberturados total:      {cob_total:,} ({cob_total/n_censo_total:.1%})")
print(f"  Coberturados Mar-26:     {cob_mar:,} ({cob_mar/n_censo_total:.1%})")
print(f"  Coberturados Abr-26:     {cob_abr:,} ({cob_abr/n_censo_total:.1%})")
if tiene_mayo:
    print(f"  Coberturados May-26:     {cob_may:,} ({cob_may/n_censo_total:.1%})")
print(f"  Censo SIN compra:        {sin_compra_censo:,} ({sin_compra_censo/n_censo_total:.1%})")
print(f"  Fecha referencia:        {FECHA_HOY_STR}")
print(f"{'='*55}")
print("Done.")
