"""
Genera el reporte FUNNEL_PREVENTA_<dia>_DD_MM_YYYY.xlsx
con el estado de efectividad diaria de preventa por cliente.

Uso:
    python generar_funnel_preventa.py <dia_semana>
    dia_semana: 1=Lunes, 2=Martes, 3=Miércoles, 4=Jueves, 5=Viernes, 6=Sábado, 7=Domingo
"""
# ── Librería estándar ──────────────────────────────────────────────────────────
import argparse
import datetime
import os
import sys
import warnings

# ── Terceros ───────────────────────────────────────────────────────────────────
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

# ── Credenciales SQL ──────────────────────────────────────────────────────────
def _leer_env(path='.env'):
    env = {}
    if os.path.exists(path):
        with open(path, encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    k, v = line.split('=', 1)
                    env[k.strip()] = v.strip()
    return env

_env = _leer_env('C:/proyectos/SSFF/.env')

def _env_req(clave):
    valor = _env.get(clave)
    if not valor:
        raise EnvironmentError(f"Variable '{clave}' no encontrada en .env")
    return valor

SQL_SERVER   = _env_req('SQL_SERVER')
SQL_DATABASE = _env_req('SQL_DATABASE')
SQL_USER     = _env_req('SQL_USER')
SQL_PASSWORD = _env_req('SQL_PASSWORD')

# ── Parámetros ────────────────────────────────────────────────────────────────
DIAS_SEMANA = {1: 'Lunes', 2: 'Martes', 3: 'Miércoles',
               4: 'Jueves', 5: 'Viernes', 6: 'Sábado', 7: 'Domingo'}

MESES_ES = {1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
            7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'}

HOY            = datetime.date.today()
HOY_STR        = HOY.strftime('%Y%m%d')      # para el SP
MES_ACTUAL_INT = int(HOY.strftime('%y%m'))   # ej: 2605
COL_MES_ACTUAL = f'Venta {MESES_ES[HOY.month]} S/'

BASE_DIR      = 'C:/proyectos/SSFF'
TABLAS_PATH   = f'{BASE_DIR}/TABLAS_RUTAS.xlsx'
VTA_HIST_PATH = f'{BASE_DIR}/files/vtas_soles_resumen_2604-2607.csv'

MESES_HIST    = ['Abr-26 S/', 'May-26 S/', 'Jun-26 S/', 'Jul-26 S/']
MESES_INT     = [2604, 2605, 2606, 2607]    # correspondencia con CSV histórico

UMBRAL_FRECUENCIA = 3
UMBRAL_TICKET     = 200

# ── Conexión SQL ──────────────────────────────────────────────────────────────
def conectar_sql():
    import pyodbc
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={SQL_SERVER};"
        f"DATABASE={SQL_DATABASE};"
        f"UID={SQL_USER};"
        f"PWD={SQL_PASSWORD};"
        f"TrustServerCertificate=yes;"
    )
    return pyodbc.connect(conn_str)

# ════════════════════════════════════════════════════════════════════════════════
# 1. CARGA DE DATOS
# ════════════════════════════════════════════════════════════════════════════════

def cargar_efectividad(conn, dia_semana: int) -> pd.DataFrame:
    """Fuente 1: SP de efectividad del día."""
    print(f"   Ejecutando SP efectividad ({HOY_STR}, día {dia_semana})...")
    query = f"EXEC [eAuren].[dbo].[spPreventa_efectividad_dia] '{HOY_STR}', '{dia_semana}'"
    df = pd.read_sql(query, conn)
    print(f"   >> {len(df):,} clientes en ruta hoy")
    return df


def cargar_ventas_historico() -> pd.DataFrame:
    """Fuente 2a: CSV histórico 2601-2604, pivotea a columnas de mes."""
    print("   Cargando histórico de ventas CSV...")
    df = pd.read_csv(VTA_HIST_PATH, dtype={'ccod_cli': str, 'mes': int, 'total_monto': float})
    # normalizar código cliente: quitar ceros a la izquierda para unificar con fuente1
    df['ccod_cli'] = df['ccod_cli'].str.lstrip('0')
    df = df[df['mes'].isin(MESES_INT)]
    pivot = df.pivot_table(index='ccod_cli', columns='mes', values='total_monto', aggfunc='sum')
    # puede que no todos los meses estén presentes en el CSV
    cols_presentes = [m for m in MESES_INT if m in pivot.columns]
    pivot.columns = [MESES_HIST[MESES_INT.index(c)] for c in cols_presentes]
    for col in MESES_HIST:
        if col not in pivot.columns:
            pivot[col] = np.nan
    pivot = pivot[MESES_HIST].reset_index()
    return pivot


def cargar_ventas_mayo(conn) -> pd.DataFrame:
    """Fuente 2b: ventas del mes actual desde SQL."""
    print(f"   Cargando ventas mayo {MES_ACTUAL_INT} desde SQL...")
    query = f"""
    SELECT
         a.ccod_cli
        ,a.[mes]
        ,SUM(a.[monto]) AS [total_monto]
    FROM [eAuren].[dbo].[base_com] a
    WHERE a.mes = '{MES_ACTUAL_INT}'
      AND a.cstatus != 'A'
      AND a.ctipo_vta = '0003'
      AND ccod_ruta != '0000'
    GROUP BY a.ccod_cli, a.[mes]
    """
    df = pd.read_sql(query, conn)
    # normalizar código: quitar ceros a la izquierda
    df['ccod_cli'] = df['ccod_cli'].astype(str).str.lstrip('0')
    df = df.rename(columns={'total_monto': COL_MES_ACTUAL})
    return df[['ccod_cli', COL_MES_ACTUAL]]


def cargar_motivos_no_preventa(conn) -> pd.DataFrame:
    """Fuente 3: estado (descripción motivo) y distancia del día."""
    print("   Cargando motivos de no preventa...")
    query = """
    SELECT codCliente, estado, distancia
    FROM dbo.viewMotivoNoPreVentaTM
    WHERE CAST(fecha AS DATE) = CAST(GETDATE() AS DATE)
    """
    df = pd.read_sql(query, conn)
    # normalizar key igual que fuente1
    df['codCliente'] = df['codCliente'].astype(str).str.lstrip('0')
    df = df.rename(columns={'estado': 'MOTIVO NO COMPRA', 'distancia': 'DISTANCIA (km)'})
    # si un cliente tiene varios registros, quedarse con el más reciente (último)
    df = df.drop_duplicates(subset='codCliente', keep='last')
    return df


def cargar_tablas_rutas() -> pd.DataFrame:
    """Maestro de rutas: columna ZONA2 por ruta."""
    df = pd.read_excel(TABLAS_PATH, sheet_name='RUTA_ACTUAL', dtype={'RUTA': str})
    return df[['RUTA', 'ZONA2']].rename(columns={'RUTA': 'ruta_key', 'ZONA2': 'Origen Lima'})


def cargar_distribucion_ffvv(conn) -> pd.DataFrame:
    """RUTA → VENDEDOR/SUPERVISOR vigente desde [eAuren].[dbo].[viewSFffvv].

    Esta vista la mantiene Sistemas actualizada permanentemente — se usa para
    sobrescribir supervisor/vendedor del SP de efectividad, que puede quedar
    desactualizado tras cambios de cartera.
    """
    df = pd.read_sql(
        "SELECT ruta, vendedorCorto, supervisor FROM [eAuren].[dbo].[viewSFffvv]", conn
    )
    df['ruta'] = df['ruta'].astype(str).str.strip()
    return df.rename(columns={'vendedorCorto': 'vendedor_sql', 'supervisor': 'supervisor_sql'})

# ════════════════════════════════════════════════════════════════════════════════
# 2. LÓGICA DE CLUSTERS
# ════════════════════════════════════════════════════════════════════════════════

def calcular_clusters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    activos = df[MESES_HIST].fillna(0)
    df['meses_activos'] = (activos > 0).sum(axis=1)
    df['total_hist']    = activos.sum(axis=1)
    df['prom_mensual']  = df['total_hist'] / df['meses_activos'].replace(0, np.nan)

    df['Coberturado'] = df['meses_activos'].apply(lambda x: 'FOCO' if x >= 1 else 'NO')

    def _seg(row):
        ma = row['meses_activos']
        p  = row['prom_mensual']
        if ma >= 1:
            frecuente = ma >= UMBRAL_FRECUENCIA
            alto      = p  >= UMBRAL_TICKET
            if frecuente and alto: return 'A - Activo frecuente alto'
            if frecuente:          return 'B - Activo frecuente bajo'
            if alto:               return 'C - Activo esporadico alto'
            return                        'D - Activo esporadico bajo'
        censo = str(row['Censo']).strip().upper() if 'Censo' in row.index else ''
        if censo == 'SI':
            zona = row['Origen Lima'] if 'Origen Lima' in row.index else ''
            return 'F - Censo Casco' if zona == 'Casco' else 'E - Censo Lima Este'
        return 'G - Inactivo cartera'

    df['Segmento'] = df.apply(_seg, axis=1)
    return df

# ════════════════════════════════════════════════════════════════════════════════
# 3. ARMADO DEL DATAFRAME FINAL
# ════════════════════════════════════════════════════════════════════════════════

def construir_base(df_efec, df_hist, df_mayo, df_motivos, df_rutas, df_ffvv) -> pd.DataFrame:
    df_efec = df_efec.copy()

    # Quedarse solo con clientes cuyo día de visita programado (diaPvta) incluye
    # el día de hoy — el SP también trae clientes que compraron fuera de su día
    # programado (fueraRuta='S'), los cuales no corresponden al funnel del día.
    df_efec = df_efec[df_efec['fueraRuta'] != 'S'].copy()

    # Hora máxima de pedido por cliente: un mismo codCliente puede aparecer más
    # de una vez (visitado/vendido por más de un vendedor el mismo día) — se toma
    # la hora más tardía entre todas sus filas y se aplica a cada una de ellas.
    hora_max_cliente = df_efec.groupby('codCliente')['horaTP'].transform('max')
    df_efec['HORA MAXIMA PEDIDO'] = hora_max_cliente

    # Sobrescribir vendedor/supervisor con la distribución vigente de viewSFffvv
    # (el SP de efectividad puede quedar desactualizado tras cambios de cartera).
    df_efec['ruta'] = df_efec['ruta'].astype(str).str.strip()
    df_efec = df_efec.merge(df_ffvv, left_on='ruta', right_on='ruta', how='left')
    df_efec['vendedor']   = df_efec['vendedor_sql'].fillna(df_efec['vendedor'])
    df_efec['supervisor'] = df_efec['supervisor_sql'].fillna(df_efec['supervisor'])
    df_efec = df_efec.drop(columns=['vendedor_sql', 'supervisor_sql'])

    # Normalizar key de cliente: quitar ceros a la izquierda para unificar con CSV/SQL
    df_efec['_key'] = df_efec['codCliente'].astype(str).str.lstrip('0')

    # Join fuente 2a (histórico pivotado)
    df = df_efec.merge(df_hist, left_on='_key', right_on='ccod_cli', how='left')

    # Join fuente 2b (venta mayo)
    df = df.merge(df_mayo, left_on='_key', right_on='ccod_cli', how='left')
    df[COL_MES_ACTUAL] = df[COL_MES_ACTUAL].fillna(0)

    # Join fuente 3 (motivos: estado y distancia)
    df = df.merge(df_motivos, left_on='_key', right_on='codCliente', how='left',
                  suffixes=('', '_motivo'))

    # Join ZONA2 desde maestro de rutas (key: columna 'ruta' del SP)
    df_rutas['ruta_key'] = df_rutas['ruta_key'].astype(str)
    df = df.merge(df_rutas, left_on='ruta', right_on='ruta_key', how='left')
    # V001/V002 (Verdum) no están en el maestro → asignar manualmente
    mask_verdum = df['ruta'].isin(['V001', 'V002'])
    df.loc[mask_verdum, 'Origen Lima'] = 'Verdum'

    # ── Censo: 'S' → 'Si', cualquier otra cosa → 'No' ────────────────────────
    df['Censo'] = df['censo'].fillna('').astype(str).str.strip().str.upper()
    df['Censo'] = df['Censo'].apply(lambda x: 'Si' if x == 'S' else 'No')

    # ── Calcular clusters (necesita Censo y Origen Lima ya presentes) ─────────
    df = calcular_clusters(df)

    # ── Monto compra desde fuente1 ────────────────────────────────────────────
    df['MONTO COMPRA S/'] = pd.to_numeric(df['monto'], errors='coerce').fillna(0)

    # ── Estado cliente ────────────────────────────────────────────────────────
    def _estado_cli(row):
        if row['MONTO COMPRA S/'] >= 1:
            return 'CON COMPRA'
        motivo = str(row['MOTIVO NO COMPRA']).strip()
        if motivo and motivo.lower() not in ('nan', 'none', ''):
            return 'MOTIVO NO COMPRA'
        return 'SIN VISITA'

    df['ESTADO_CLIENTE'] = df.apply(_estado_cli, axis=1)

    # ── >150 metros ───────────────────────────────────────────────────────────
    df['DISTANCIA (km)'] = pd.to_numeric(df['DISTANCIA (km)'], errors='coerce')
    df['>150 METROS']    = df['DISTANCIA (km)'].apply(
        lambda x: 'SI' if pd.notna(x) and x >= 0.15 else ''
    )

    # ── Renombrar columnas de fuente1 a nombres de presentación ──────────────
    df = df.rename(columns={
        'supervisor': 'Supervisor',
        'ruta':       'Ruta',
        'vendedor':   'Vendedor',
        'diaPvta':    'Día Visita',
        'codCliente': 'Código',
        'cliente':    'Cliente',
        'distrito':   'Distrito_cliente',
    })

    # ── Seleccionar y ordenar columnas finales ────────────────────────────────
    cols_finales = [
        'Supervisor', 'Ruta', 'Vendedor', 'Día Visita',
        'Código', 'Cliente', 'Distrito_cliente', 'Censo',
        'Origen Lima',
        *MESES_HIST, COL_MES_ACTUAL,
        'Coberturado', 'ESTADO_CLIENTE', 'MONTO COMPRA S/',
        'MOTIVO NO COMPRA', 'HORA MAXIMA PEDIDO', 'DISTANCIA (km)', '>150 METROS',
    ]
    cols_presentes = [c for c in cols_finales if c in df.columns]
    df = df[cols_presentes]

    sort_cols = [c for c in ['Supervisor', 'Ruta', 'Cliente'] if c in df.columns]
    df = df.sort_values(sort_cols).reset_index(drop=True)

    return df


def _detectar_col(df, candidatos):
    """Devuelve el primer nombre de columna que exista en el DataFrame."""
    for c in candidatos:
        if c in df.columns:
            return c
    return None

# ════════════════════════════════════════════════════════════════════════════════
# 4. HELPERS OPENPYXL
# ════════════════════════════════════════════════════════════════════════════════

def fill(hex_color):
    return PatternFill('solid', fgColor=hex_color)

def fnt(bold=False, size=10, color='000000', name='Aptos Narrow', italic=False):
    return Font(bold=bold, size=size, color=color, name=name, italic=italic)

def aln(h='center', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def brd(thin=True):
    s = Side(style='thin') if thin else Side(style=None)
    return Border(left=s, right=s, top=s, bottom=s)

# Paleta de colores
C_AZUL_OSC  = '006D77'   # título principal
C_AZUL_MED  = '83C5BE'   # encabezados de columna
C_AZUL_CLAR = 'EDF6F9'   # filas alternas
C_GRIS_CLAR = 'F5F5F5'   # filas pares base
C_BLANCO    = 'FFFFFF'

COL_WIDTHS = {
    'Supervisor': 22, 'Ruta': 8, 'Vendedor': 22, 'Día Visita': 10,
    'Código': 10, 'Cliente': 35, 'Distrito_cliente': 18, 'Censo': 7,
    'Origen Lima': 14,
    **{m: 12 for m in MESES_HIST}, COL_MES_ACTUAL: 14,
    'Coberturado': 12, 'ESTADO_CLIENTE': 18, 'MONTO COMPRA S/': 15,
    'MOTIVO NO COMPRA': 25, 'HORA MAXIMA PEDIDO': 16,
    'DISTANCIA (km)': 14, '>150 METROS': 12,
}

COLOR_ESTADO = {
    'CON COMPRA':      'E2EFDA',
    'MOTIVO NO COMPRA':'FCE4D6',
    'SIN VISITA':      'FFF2CC',
}

COLOR_SEGMENTO = {
    'A - Activo frecuente alto':  'C6EFCE',
    'B - Activo frecuente bajo':  'DDEBF7',
    'C - Activo esporadico alto': 'EBF1DE',
    'D - Activo esporadico bajo': 'FFF2CC',
    'E - Censo Lima Este':        'FCE4D6',
    'F - Censo Casco':            'F4CCCC',
    'G - Inactivo cartera':       'D9D9D9',
}

# ════════════════════════════════════════════════════════════════════════════════
# 5. ESCRITURA DEL EXCEL
# ════════════════════════════════════════════════════════════════════════════════

def escribir_hoja_funnel(ws, df: pd.DataFrame, titulo: str):
    cols = list(df.columns)
    n_cols = len(cols)

    # Fila 1: título combinado
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    c = ws.cell(row=1, column=1, value=titulo)
    c.font      = fnt(bold=True, size=13, color=C_BLANCO)
    c.fill      = fill(C_AZUL_OSC)
    c.alignment = aln('center')

    # Fila 2: encabezados
    for j, col in enumerate(cols, 1):
        c = ws.cell(row=2, column=j, value=col)
        c.font      = fnt(bold=True, size=9, color='006D77')
        c.fill      = fill(C_AZUL_MED)
        c.alignment = aln('center', wrap=True)
        c.border    = brd()

    # Ancho de columnas
    for j, col in enumerate(cols, 1):
        ancho = COL_WIDTHS.get(col, 15)
        ws.column_dimensions[get_column_letter(j)].width = ancho

    ws.row_dimensions[2].height = 30

    # Índices de columnas clave para colorear
    idx_estado   = cols.index('ESTADO_CLIENTE') + 1 if 'ESTADO_CLIENTE'   in cols else None
    idx_segmento = cols.index('Segmento')        + 1 if 'Segmento'         in cols else None
    idx_cobert   = cols.index('Coberturado')     + 1 if 'Coberturado'      in cols else None

    cols_soles = [j+1 for j, c in enumerate(cols)
                  if c in (*MESES_HIST, COL_MES_ACTUAL, 'MONTO COMPRA S/')]
    cols_dist  = [j+1 for j, c in enumerate(cols) if c == 'DISTANCIA (km)']
    cols_hora  = [j+1 for j, c in enumerate(cols) if c == 'HORA MAXIMA PEDIDO']

    idx_cliente = cols.index('Cliente') + 1 if 'Cliente' in cols else -1

    # Datos — usar iloc para evitar problemas con nombres que tienen espacios/slash
    for i, (_, fila) in enumerate(df.iterrows(), 3):
        fila_par = (i % 2 == 0)
        bg_base  = C_AZUL_CLAR if fila_par else C_BLANCO

        estado   = fila.get('ESTADO_CLIENTE', '')
        segmento = fila.get('Segmento', '')

        for j, col in enumerate(cols, 1):
            val = fila[col]
            if isinstance(val, float) and pd.isna(val):
                val = None
            elif hasattr(val, 'item'):   # numpy scalar → Python nativo
                val = val.item()

            c = ws.cell(row=i, column=j, value=val)
            c.border    = brd()
            c.alignment = aln('left') if j == idx_cliente else aln('center')

            # Color de fondo por columna
            if j == idx_estado and estado in COLOR_ESTADO:
                c.fill = fill(COLOR_ESTADO[estado])
            elif j == idx_segmento and segmento in COLOR_SEGMENTO:
                c.fill = fill(COLOR_SEGMENTO[segmento])
            elif j == idx_cobert:
                c.fill = fill('C6EFCE') if val == 'FOCO' else fill('FCE4D6')
            else:
                c.fill = fill(bg_base)

            # Formato numérico
            if j in cols_soles:
                c.number_format = '#,##0.00'
                c.font = fnt(size=9)
            elif j in cols_dist:
                c.number_format = '0.000'
                c.font = fnt(size=9)
            elif j in cols_hora:
                c.number_format = 'hh:mm'
                c.font = fnt(size=9)
            else:
                c.font = fnt(size=9)

    # Congelar encabezados
    ws.freeze_panes = 'A3'

    # Autofilter
    ws.auto_filter.ref = f"A2:{get_column_letter(n_cols)}{len(df) + 2}"


def calcular_resumen_supervisor(df_total: pd.DataFrame) -> pd.DataFrame:
    """
    Resumen de avance de cobertura/preventa por Supervisor (hoja RESUMEN).

    Columnas (ver files/EJEMPLO5.png):
      CARTERA DIA VISITA     : cuenta de todos los clientes del día (df_total ya
                                viene filtrado al día de hoy).
      CLIENTES CON VTA DIA   : cuenta de clientes con venta HOY (MONTO COMPRA S/ > 0)
                                O con facturación acumulada en el mes actual
                                (Venta Agosto S/ > 0) — complemento exacto de
                                CLIENTES S/VTA MES ACT. (por ley de De Morgan, la
                                suma de ambas columnas siempre da CARTERA DIA VISITA).
      %COBERTURA DIA         : con_vta_dia / cartera.
      CLIENTES S/VTA MES ACT.: cuenta de clientes SIN venta hoy Y SIN facturación
                                acumulada en el mes actual (Venta Agosto S/ == 0
                                Y MONTO COMPRA S/ == 0). Un cliente que compró hoy
                                ya cuenta como cubierto aquí, aunque el acumulado
                                mensual todavía no lo refleje.
      %PDTE COBERTURA        : ese conteo / cartera.
      CLIENTES S/VTA ULT2MES : subconjunto del anterior — además tampoco tienen
                                facturación en el mes anterior (Jul-26 S/ == 0).
      %PDTE COBERTURA ULT2M  : ese conteo / cartera.
    """
    col_jul = 'Jul-26 S/'

    def _agg(g):
        # Jul-26 S/ y Venta Agosto S/ quedan NaN cuando el cliente no compró ese
        # mes — se tratan igual que 0 para estas métricas.
        hoy_vacio = g['MONTO COMPRA S/'].fillna(0) == 0
        ago_vacio = g[COL_MES_ACTUAL].fillna(0) == 0
        jul_vacio = g[col_jul].fillna(0) == 0

        cartera       = len(g)
        con_vta_dia   = (~ago_vacio | ~hoy_vacio).sum()
        sin_vta_mes   = (ago_vacio & hoy_vacio).sum()
        sin_vta_2mes  = (ago_vacio & jul_vacio & hoy_vacio).sum()
        return pd.Series({
            'CARTERA DIA VISITA':      cartera,
            'CLIENTES CON VTA DIA':    con_vta_dia,
            '%COBERTURA DIA':          con_vta_dia / cartera if cartera else 0,
            'CLIENTES S/VTA MES ACT.': sin_vta_mes,
            '%PDTE COBERTURA':         sin_vta_mes / cartera if cartera else 0,
            'CLIENTES S/VTA ULT2MES':  sin_vta_2mes,
            '%PDTE COBERTURA ULT2M':   sin_vta_2mes / cartera if cartera else 0,
        })

    resumen = df_total.groupby('Supervisor', sort=True).apply(_agg, include_groups=False)
    resumen = resumen.reset_index().sort_values('Supervisor')

    total = resumen[['CARTERA DIA VISITA', 'CLIENTES CON VTA DIA',
                      'CLIENTES S/VTA MES ACT.', 'CLIENTES S/VTA ULT2MES']].sum()
    fila_total = {
        'Supervisor':              'TOTAL AUREN',
        'CARTERA DIA VISITA':      total['CARTERA DIA VISITA'],
        'CLIENTES CON VTA DIA':    total['CLIENTES CON VTA DIA'],
        '%COBERTURA DIA':          total['CLIENTES CON VTA DIA'] / total['CARTERA DIA VISITA'] if total['CARTERA DIA VISITA'] else 0,
        'CLIENTES S/VTA MES ACT.': total['CLIENTES S/VTA MES ACT.'],
        '%PDTE COBERTURA':         total['CLIENTES S/VTA MES ACT.'] / total['CARTERA DIA VISITA'] if total['CARTERA DIA VISITA'] else 0,
        'CLIENTES S/VTA ULT2MES':  total['CLIENTES S/VTA ULT2MES'],
        '%PDTE COBERTURA ULT2M':   total['CLIENTES S/VTA ULT2MES'] / total['CARTERA DIA VISITA'] if total['CARTERA DIA VISITA'] else 0,
    }
    resumen = pd.concat([resumen, pd.DataFrame([fila_total])], ignore_index=True)
    return resumen


COLOR_RESUMEN_HEADER = {
    'Supervisor':              'DDEBF7',
    'CARTERA DIA VISITA':      'DDEBF7',
    'CLIENTES CON VTA DIA':    'E2EFDA',
    '%COBERTURA DIA':          'E2EFDA',
    'CLIENTES S/VTA MES ACT.': 'FBE0CE',
    '%PDTE COBERTURA':         'FBE0CE',
    'CLIENTES S/VTA ULT2MES':  'FFFF00',
    '%PDTE COBERTURA ULT2M':   'FFFF00',
}


def escribir_hoja_resumen(ws, df: pd.DataFrame, hora_corte: str):
    cols = list(df.columns)  # Supervisor + 6 columnas de métricas
    n_cols = len(cols)
    nombre_dia = DIAS_SEMANA[HOY.weekday() + 1]

    # Fila 1: título (A1:F1) + CORTE (G1) + hora (H1)
    ws.merge_cells('A1:F1')
    c = ws.cell(row=1, column=1,
                value=f"AVANCE DE COBERTURA/PREVENTA DIA DE VISITA - [{nombre_dia.upper()}] - [{HOY.strftime('%d/%m/%Y')}]")
    c.font      = fnt(bold=True, size=12, color='000000', italic=True)
    c.fill      = fill('F8CBD9')
    c.alignment = aln('left')

    # H1 (hora) define el formato; G1 (etiqueta CORTE) copia el mismo formato
    fmt_hora = {'font': fnt(bold=True, size=13, color='000000'), 'fill': fill('FFFF99'), 'alignment': aln('center')}

    c_lbl = ws.cell(row=1, column=7, value='CORTE')  # G1
    c_lbl.font      = fmt_hora['font']
    c_lbl.fill      = fmt_hora['fill']
    c_lbl.alignment = fmt_hora['alignment']

    c_hora = ws.cell(row=1, column=8, value=hora_corte)  # H1
    c_hora.font      = fmt_hora['font']
    c_hora.fill      = fmt_hora['fill']
    c_hora.alignment = fmt_hora['alignment']

    # Fila 2: encabezados
    for j, col in enumerate(cols, 1):
        texto = col.upper() if col == 'Supervisor' else col
        c = ws.cell(row=2, column=j, value=texto)
        c.font      = fnt(bold=True, size=9, color='000000')
        c.fill      = fill(COLOR_RESUMEN_HEADER.get(col, C_AZUL_MED))
        c.alignment = aln('center', wrap=True)
        c.border    = brd()
    ws.row_dimensions[2].height = 45

    ws.column_dimensions['A'].width = 20
    for j in range(2, n_cols + 1):
        ws.column_dimensions[get_column_letter(j)].width = 11

    cols_pct = [j + 1 for j, c in enumerate(cols) if c.startswith('%')]

    # Filas de datos (supervisores + TOTAL AUREN al final)
    n = len(df)
    for i, (_, fila) in enumerate(df.iterrows(), 3):
        es_total = fila['Supervisor'] == 'TOTAL AUREN'
        for j, col in enumerate(cols, 1):
            val = fila[col]
            if hasattr(val, 'item'):
                val = val.item()
            c = ws.cell(row=i, column=j, value=val)
            c.border    = brd()
            c.alignment = aln('left' if j == 1 else 'center')
            c.font      = fnt(bold=es_total, size=10)
            if es_total:
                c.fill = fill('D9D9D9')
            if j in cols_pct:
                c.number_format = '0%'

    ws.freeze_panes = 'A3'


def generar_excel(df_total: pd.DataFrame, nombre_archivo: str, hora_corte: str):
    wb = Workbook()

    ws_resumen = wb.active
    ws_resumen.title = 'RESUMEN'
    df_resumen = calcular_resumen_supervisor(df_total)
    print(f"   Escribiendo hoja RESUMEN ({len(df_resumen):,} filas)...")
    escribir_hoja_resumen(ws_resumen, df_resumen, hora_corte)

    ws_total = wb.create_sheet(title='FUNNEL TOTAL')

    supervisores = df_total['Supervisor'].dropna().unique() if 'Supervisor' in df_total.columns else []

    print(f"   Escribiendo hoja FUNNEL TOTAL ({len(df_total):,} filas)...")
    escribir_hoja_funnel(ws_total, df_total, f'FUNNEL PREVENTA — {HOY.strftime("%d/%m/%Y")}')

    for sup in sorted(supervisores):
        df_sup = df_total[df_total['Supervisor'] == sup].reset_index(drop=True)
        nombre_hoja = str(sup)[:31]  # límite Excel
        ws = wb.create_sheet(title=nombre_hoja)
        escribir_hoja_funnel(ws, df_sup, f'FUNNEL PREVENTA — {sup} — {HOY.strftime("%d/%m/%Y")}')
        print(f"   Hoja '{nombre_hoja}': {len(df_sup):,} clientes")

    wb.save(nombre_archivo)
    print(f"\n   Guardado: {nombre_archivo}")


# ════════════════════════════════════════════════════════════════════════════════
# 6. MAIN
# ════════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description='Funnel Preventa diario')
    parser.add_argument('dia_semana', type=int, nargs='?', default=None,
                         help='1=Lunes … 7=Domingo (default: hoy)')
    parser.add_argument('--corte', default='9AM',
                         help="Etiqueta del corte horario mostrada en la hoja RESUMEN (default: '9AM')")
    args = parser.parse_args()

    if args.dia_semana is not None:
        dia_semana = args.dia_semana
        if dia_semana not in DIAS_SEMANA:
            print(f"Error: dia_semana debe estar entre 1 y 7 (recibido: {dia_semana})")
            sys.exit(1)
    else:
        # weekday(): 0=lunes … 6=domingo  →  sumar 1 para el convenio del SP
        dia_semana = HOY.weekday() + 1

    nombre_dia = DIAS_SEMANA[dia_semana]
    fecha_str  = HOY.strftime('%d_%m_%Y')
    output     = f'{BASE_DIR}/FUNNEL_PREVENTA_{nombre_dia}_{fecha_str}.xlsx'

    print(f"\n{'='*60}")
    print(f"  FUNNEL PREVENTA — {nombre_dia} {HOY.strftime('%d/%m/%Y')}")
    print(f"{'='*60}")

    conn = conectar_sql()
    try:
        print("\n[1] Cargando datos...")
        df_efec    = cargar_efectividad(conn, dia_semana)
        df_hist    = cargar_ventas_historico()
        df_mayo    = cargar_ventas_mayo(conn)
        df_motivos = cargar_motivos_no_preventa(conn)
        df_rutas   = cargar_tablas_rutas()
        df_ffvv    = cargar_distribucion_ffvv(conn)

        print("\n[2] Construyendo base consolidada...")
        df_base = construir_base(df_efec, df_hist, df_mayo, df_motivos, df_rutas, df_ffvv)

        print(f"\n   Columnas del SP efectividad: {list(df_efec.columns)}")
        print(f"   Total clientes: {len(df_base):,}")
        print(f"   Distribución ESTADO_CLIENTE:\n{df_base['ESTADO_CLIENTE'].value_counts().to_string()}")

        print("\n[3] Generando Excel...")
        generar_excel(df_base, output, args.corte)

    finally:
        conn.close()

    print("\nProceso completado OK.")


if __name__ == '__main__':
    main()
