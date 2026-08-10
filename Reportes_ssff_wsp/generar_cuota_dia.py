"""
generar_cuota_dia.py

Recalcula CUOTA_DIA por ruta usando la formula del Gerente Comercial:

    CUOTA_DIA = (Nueva Cuota - Avance del mes) / Dias habiles restantes

Fuentes:
  - CUOTAS_SSFF.xlsx (hoja SOLES)    → Nueva Cuota por RUTA (columna del mes vigente)
  - base_com (SQL)                    → Avance del mes por RUTA (sin anulados)
  - TABLAS_RUTAS.xlsx                 → RUTA → VENDEDOR, SUPERVISOR, PREFIX_RUTA
  - Dias habiles restantes            → lunes-sabado desde mañana hasta fin de mes,
                                         excluyendo feriados

Genera: CuotaDia_RECALCULADA.xlsx (para revisión)

Uso:
    python generar_cuota_dia.py
"""

import calendar
import datetime
import os
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Paths (mismos que generar_corte_ventas.py) ───────────────────────────────
BASE_DIR    = 'C:/proyectos/SSFF'
OUT_DIR     = f'{BASE_DIR}/Reportes_ssff_wsp'
CUOTA_PATH  = f'{BASE_DIR}/files/CUOTAS_SSFF.xlsx'
TABLAS_PATH = f'{BASE_DIR}/TABLAS_RUTAS.xlsx'

FERIADOS = {datetime.date(2026, 7, 29)}   # 23/07 y 28/07 son laborables (confirmado por Jesus Asencios)


# ── Credenciales SQL ──────────────────────────────────────────────────────────

def _leer_env(path=f'{BASE_DIR}/.env'):
    env = {}
    if os.path.exists(path):
        with open(path, encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    k, v = line.split('=', 1)
                    env[k.strip()] = v.strip()
    return env

_env = _leer_env()

def conectar_sql():
    import pyodbc
    e = _env
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={e['SQL_SERVER']};DATABASE={e['SQL_DATABASE']};"
        f"UID={e['SQL_USER']};PWD={e['SQL_PASSWORD']};"
        f"TrustServerCertificate=yes;"
    )
    return pyodbc.connect(conn_str)


# ── Días hábiles restantes ────────────────────────────────────────────────────

def dias_habiles_restantes() -> list:
    """Días hábiles desde HOY hasta el fin del mes vigente inclusive.
    Lunes a sábado (weekday 0-5), excluyendo feriados en FERIADOS.
    Domingo (weekday 6) no es hábil."""
    hoy = datetime.date.today()
    ultimo_dia = calendar.monthrange(hoy.year, hoy.month)[1]
    fin_mes = datetime.date(hoy.year, hoy.month, ultimo_dia)
    dias = []
    d = hoy
    while d <= fin_mes:
        if d.weekday() < 6 and d not in FERIADOS:   # lun-sab y no feriado
            dias.append(d)
        d += datetime.timedelta(days=1)
    return dias


# ── Carga de datos ────────────────────────────────────────────────────────────

def cargar_avance_sql() -> pd.DataFrame:
    """Avance del mes vigente por ruta desde base_com."""
    mes_actual = datetime.date.today().strftime('%y%m')
    query = f"""
        SELECT
               [mes]
              ,[ccod_ruta]
              ,SUM([monto]) AS AVANCE
          FROM [eAuren].[dbo].[base_com]
         WHERE mes='{mes_actual}' AND cstatus != 'A'
         GROUP BY [mes], [ccod_ruta]
    """
    print(f"   Consultando avance {mes_actual} en SQL...")
    conn = conectar_sql()
    try:
        df = pd.read_sql(query, conn)
    finally:
        conn.close()
    df['ccod_ruta'] = df['ccod_ruta'].astype(str).str.strip()
    df = df.rename(columns={'ccod_ruta': 'RUTA', 'AVANCE': 'AVANCE_MES'})
    print(f"   {len(df)} rutas con avance en {mes_actual}")
    return df[['RUTA', 'AVANCE_MES']]


def _col_mes_vigente(columnas) -> object:
    """Encuentra la columna de fecha (datetime) del mes vigente entre los headers de una hoja."""
    hoy = datetime.date.today()
    for col in columnas:
        if isinstance(col, datetime.datetime) and col.year == hoy.year and col.month == hoy.month:
            return col
    raise ValueError(f"No se encontró columna del mes vigente ({hoy.year}-{hoy.month:02d}) en {columnas}")


def cargar_cuota() -> pd.DataFrame:
    """Nueva Cuota mensual (soles) por ruta desde CUOTAS_SSFF.xlsx, hoja SOLES."""
    df = pd.read_excel(CUOTA_PATH, sheet_name='SOLES', dtype={'RUTA': str})
    df['RUTA'] = df['RUTA'].str.strip()
    col_mes = _col_mes_vigente(df.columns)
    return df[['RUTA', col_mes]].rename(columns={col_mes: 'CUOTA_MES'})


# Caché en memoria del maestro de rutas (Excel + VENDEDOR/SUPERVISOR vigente de SQL)
_df_rutas_maestro_cache: 'pd.DataFrame | None' = None


def _cargar_distribucion_ffvv_sql() -> pd.DataFrame:
    """RUTA → VENDEDOR/SUPERVISOR vigente desde [eAuren].[dbo].[viewSFffvv].

    Esta vista la mantiene Sistemas actualizada permanentemente, a diferencia de
    TABLAS_RUTAS.xlsx (mantenimiento manual mensual, históricamente desactualizado).
    """
    conn = conectar_sql()
    try:
        df = pd.read_sql(
            "SELECT ruta, vendedorCorto, supervisor FROM [eAuren].[dbo].[viewSFffvv]", conn
        )
    finally:
        conn.close()
    df['ruta'] = df['ruta'].astype(str).str.strip()
    return df.rename(columns={'ruta': 'RUTA', 'vendedorCorto': 'VENDEDOR', 'supervisor': 'SUPERVISOR'})


def _cargar_rutas_maestro() -> pd.DataFrame:
    """TABLAS_RUTAS.xlsx con VENDEDOR/SUPERVISOR sobrescritos por la vista SQL vigente."""
    global _df_rutas_maestro_cache
    if _df_rutas_maestro_cache is None:
        df = pd.read_excel(TABLAS_PATH, sheet_name='RUTA_ACTUAL', dtype={'RUTA': str})
        df['RUTA'] = df['RUTA'].str.strip()

        df_ffvv = _cargar_distribucion_ffvv_sql()
        df = df.drop(columns=['VENDEDOR', 'SUPERVISOR']).merge(df_ffvv, on='RUTA', how='left')
        faltantes = df[df['VENDEDOR'].isna()]['RUTA'].tolist()
        if faltantes:
            print(f"   [WARN] Rutas sin distribución en viewSFffvv: {faltantes}")

        _df_rutas_maestro_cache = df
    return _df_rutas_maestro_cache


def cargar_rutas() -> pd.DataFrame:
    """RUTA → VENDEDOR, SUPERVISOR, FFVV (PREFIX_RUTA)."""
    df = _cargar_rutas_maestro()
    df = df.rename(columns={'PREFIX_RUTA': 'FFVV'})
    return df[['RUTA', 'FFVV', 'VENDEDOR', 'SUPERVISOR']]


# ── Estilos ───────────────────────────────────────────────────────────────────

def fill(color): return PatternFill('solid', fgColor=color)
def fnt(bold=False, size=11, color='000000'):
    return Font(name='Aptos Narrow', bold=bold, size=size, color=color)
def aln(h='center', v='center'):
    return Alignment(horizontal=h, vertical=v, wrap_text=False)
def brd_all():
    s = Side(style='thin')
    return Border(left=s, right=s, top=s, bottom=s)

C_TIT   = '1F3864'   # azul marino (mismo que ZONAL)
C_HDR   = 'DAE3F3'   # azul claro
C_ALRT  = 'FFD9D9'   # rosa claro → CUOTA_DIA negativa (cuota ya cumplida)
C_GRIS  = 'D9D9D9'
C_BLANC = 'FFFFFF'
FMT_NUM = '#,##0'
FMT_PCT = '0%'


# ── Escritura Excel ───────────────────────────────────────────────────────────

def escribir_excel(df: pd.DataFrame, dias_hab: list, out_path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = 'CUOTA_DIA'
    ws.sheet_view.showGridLines = False

    hoy       = datetime.date.today()
    hoy_str   = hoy.strftime('%d/%m/%Y')
    n_dias    = len(dias_hab)
    ult_dia   = dias_hab[-1].strftime('%d/%m') if dias_hab else '-'
    primer    = dias_hab[0].strftime('%d/%m') if dias_hab else '-'
    feriados_str = ', '.join(f.strftime('%d/%m') for f in sorted(FERIADOS))

    # ── Fila 1: título
    headers_title = (
        f'CUOTA DIA RECALCULADA — {hoy.strftime("%B %Y")}   |   '
        f'Generado: {hoy_str}   |   '
        f'Dias habiles restantes: {n_dias} ({primer} al {ult_dia}, excl. {feriados_str})'
    )
    ws.merge_cells('A1:K1')
    c = ws.cell(1, 1, headers_title)
    c.font   = fnt(bold=True, size=12, color='FFFFFF')
    c.fill   = fill(C_TIT)
    c.alignment = aln('left')

    # ── Fila 2: encabezados
    cols = ['FFVV', 'RUTA', 'VENDEDOR', 'SUPERVISOR',
            'CUOTA MES', 'AVANCE MES', 'SALDO', 'DIAS HAB. REST.', 'CUOTA DIA']
    for j, txt in enumerate(cols, 1):
        c = ws.cell(2, j, txt)
        c.font      = fnt(bold=True)
        c.fill      = fill(C_HDR)
        c.alignment = aln()
        c.border    = brd_all()

    # ── Filas de datos
    for i, row in enumerate(df.itertuples(), 3):
        es_par = (i % 2 == 0)
        color_f = C_GRIS if es_par else C_BLANC
        saldo       = row.CUOTA_MES - row.AVANCE_MES
        cuota_nueva = row.CUOTA_DIA_NUEVA
        color_saldo = C_ALRT if saldo < 0 else color_f
        color_nueva = C_ALRT if cuota_nueva < 0 else color_f

        vals = [
            (1, row.FFVV,        color_f,      'General'),
            (2, row.RUTA,        color_f,      'General'),
            (3, row.VENDEDOR,    color_f,      'General'),
            (4, row.SUPERVISOR,  color_f,      'General'),
            (5, row.CUOTA_MES,   color_f,      FMT_NUM),
            (6, row.AVANCE_MES,  color_f,      FMT_NUM),
            (7, saldo,           color_saldo,  FMT_NUM),
            (8, n_dias,          color_f,      'General'),
            (9, cuota_nueva,     color_nueva,  FMT_NUM),
        ]
        for j, val, color_c, fmt in vals:
            c = ws.cell(i, j, round(val, 2) if isinstance(val, float) else val)
            c.font          = fnt()
            c.fill          = fill(color_c)
            c.alignment     = aln('left') if j in (2, 3, 4) else aln()
            c.number_format = fmt

    # ── Fila TOTAL
    r_tot = len(df) + 3
    last_data = r_tot - 1
    n_cols = 9
    for j in range(1, n_cols + 1):
        c = ws.cell(r_tot, j)
        c.fill      = fill(C_TIT)
        c.font      = fnt(bold=True, color='FFFFFF')
        c.border    = brd_all()
        c.alignment = aln()
    ws.cell(r_tot, 1, 'TOTAL').alignment = aln()
    for j, formula in [
        (5, f'=SUM(E3:E{last_data})'),
        (6, f'=SUM(F3:F{last_data})'),
        (7, f'=E{r_tot}-F{r_tot}'),
        (8, n_dias),
        (9, f'=G{r_tot}/H{r_tot}'),
    ]:
        c = ws.cell(r_tot, j, formula)
        c.number_format = FMT_NUM

    # ── Anchos
    anchos = {'A': 7, 'B': 8, 'C': 22, 'D': 22,
              'E': 14, 'F': 14, 'G': 13, 'H': 9, 'I': 16}
    for col, w in anchos.items():
        ws.column_dimensions[col].width = w
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 32
    ws.freeze_panes = 'A3'

    wb.save(out_path)
    print(f'\n   Guardado: {out_path}')


# ── API pública: usada por generar_corte_ventas.py ───────────────────────────

def cargar_cuota_pedidos_categoria() -> pd.DataFrame:
    """Lee CUOTAS_SSFF.xlsx hoja PEDIDOS y devuelve DataFrame con columnas:
    PREFIX_RUTA, Categoria, CUOTA_PEDIDOS.

    El valor de CUOTA_PEDIDOS es la cuota DIARIA directa (no hay que dividir
    entre días hábiles — el archivo ya tiene la cuota por día).
    Combina con TABLAS_RUTAS para validar que los PREFIX_RUTA coincidan.
    """
    df = pd.read_excel(CUOTA_PATH, sheet_name='PEDIDOS')
    col_mes = _col_mes_vigente(df.columns)
    col_cat = df.columns[1]   # segunda columna: 'Categoría' (puede tener tilde)
    df = df[['PREFIX_RUTA', col_cat, col_mes]].rename(
        columns={col_cat: 'Categoria', col_mes: 'CUOTA_PEDIDOS'}
    )
    df['PREFIX_RUTA'] = df['PREFIX_RUTA'].astype(str).str.strip()
    df['Categoria']   = df['Categoria'].astype(str).str.strip().str.upper()
    df_rutas = _cargar_rutas_maestro()
    prefix_validos = set(df_rutas['PREFIX_RUTA'].astype(str).str.strip().dropna().unique())
    sin_match = set(df['PREFIX_RUTA'].dropna().unique()) - prefix_validos
    if sin_match:
        print(f"   [WARN] PREFIX_RUTA en PEDIDOS sin match en TABLAS_RUTAS: {sin_match}")
    return df


def calcular_cuota_pedidos_supervisor() -> dict:
    """Devuelve {SUPERVISOR: cuota_pedidos_dia_total} sumando todas las
    categorías de su PREFIX_RUTA directamente desde CUOTAS_SSFF.xlsx hoja PEDIDOS.

    El valor en PEDIDOS ya es la cuota diaria — no se divide entre días hábiles.
    Combina con TABLAS_RUTAS por PREFIX_RUTA para mapear a SUPERVISOR.
    Un supervisor con múltiples PREFIX_RUTA (ej. JUAN CARLOS ARAY: KB+P0) suma ambos.
    """
    df_ped = cargar_cuota_pedidos_categoria()
    cuota_por_prefix = df_ped.groupby('PREFIX_RUTA')['CUOTA_PEDIDOS'].sum().to_dict()

    df_rutas = _cargar_rutas_maestro().copy()
    df_rutas['PREFIX_RUTA'] = df_rutas['PREFIX_RUTA'].astype(str).str.strip()
    df_rutas['SUPERVISOR']  = df_rutas['SUPERVISOR'].fillna('SIN ASIGNAR')

    resultado = {}
    for _, row in df_rutas[['SUPERVISOR', 'PREFIX_RUTA']].drop_duplicates().iterrows():
        sup    = row['SUPERVISOR']
        prefix = row['PREFIX_RUTA']
        cuota  = cuota_por_prefix.get(prefix, 0.0)
        resultado[sup] = resultado.get(sup, 0.0) + cuota

    return resultado


_DIAS_ES = {
    0: 'lunes', 1: 'martes', 2: 'miércoles',
    3: 'jueves', 4: 'viernes', 5: 'sábado', 6: 'domingo',
}


def cargar_cuota_dia_tabla() -> dict:
    """Lee CUOTAS_SSFF.xlsx hoja CUOTA_DIA → {nombre_dia: monto}.
    Normaliza a minúsculas para comparar con _DIAS_ES."""
    df = pd.read_excel(CUOTA_PATH, sheet_name='CUOTA_DIA')
    df.columns = ['DIA', 'CUOTA_DIA']
    df['DIA'] = df['DIA'].astype(str).str.strip().str.lower()
    return df.set_index('DIA')['CUOTA_DIA'].to_dict()


def calcular_cuota_dia_vendedor() -> dict:
    """Devuelve {RUTA: cuota_dia} distribuyendo la CUOTA_DIA del día según el
    peso porcentual de cada ruta sobre la cuota mensual total (hoja SOLES).

    CUOTA_DIA_ruta = peso_ruta × CUOTA_DIA_dia_semana
    peso_ruta      = CUOTA_MES_ruta / SUM(CUOTA_MES_todas_las_rutas)
    """
    df_cuota = cargar_cuota()

    cuota_total = df_cuota['CUOTA_MES'].sum()
    if cuota_total == 0:
        return {}

    cuota_dia_tabla = cargar_cuota_dia_tabla()
    hoy = datetime.date.today()
    nombre_dia = _DIAS_ES[hoy.weekday()]
    cuota_dia_hoy = cuota_dia_tabla.get(nombre_dia, 0.0)
    print(f"   Cuota del día ({nombre_dia}): S/ {cuota_dia_hoy:,.0f}")

    df_cuota['PESO'] = df_cuota['CUOTA_MES'] / cuota_total
    df_cuota['CUOTA_DIA'] = df_cuota['PESO'] * cuota_dia_hoy

    return df_cuota.set_index('RUTA')['CUOTA_DIA'].to_dict()


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    hoy = datetime.date.today()
    mes_nombre = hoy.strftime('%B %Y')
    ultimo_dia = calendar.monthrange(hoy.year, hoy.month)[1]

    print('=' * 65)
    print(f'  CUOTA DIA RECALCULADA — {mes_nombre}')
    print('=' * 65)

    # Días hábiles restantes
    dias_hab = dias_habiles_restantes()
    print(f'\n[1] Dias habiles restantes desde {hoy + datetime.timedelta(1)} hasta {hoy.month:02d}/{ultimo_dia}:')
    print(f'    {len(dias_hab)} dias (excluye domingos + feriados: {sorted(FERIADOS)})')
    for d in dias_hab:
        print(f'    {d.strftime("%A %d/%m")}')

    if len(dias_hab) == 0:
        print('\n[ERROR] No quedan dias habiles en el mes. Abortando.')
        return

    # Cargar cuota mensual
    print('\n[2] Cargando cuota mensual (CUOTAS_SSFF.xlsx, hoja SOLES)...')
    df_cuota = cargar_cuota()
    print(f'    {len(df_cuota)} rutas con cuota')

    # Cargar avance SQL
    print('\n[3] Cargando avance del mes desde SQL...')
    df_avance = cargar_avance_sql()

    # Cargar maestro de rutas
    print('\n[4] Cargando maestro de rutas...')
    df_rutas = cargar_rutas()

    # Merge
    print('\n[5] Calculando...')
    df = df_cuota.merge(df_avance, on='RUTA', how='left')
    df['AVANCE_MES'] = df['AVANCE_MES'].fillna(0.0)
    df = df.merge(df_rutas, on='RUTA', how='left')
    df['VENDEDOR']   = df['VENDEDOR'].fillna('SIN ASIGNAR')
    df['SUPERVISOR'] = df['SUPERVISOR'].fillna('SIN ASIGNAR')
    df['FFVV']       = df['FFVV'].fillna('?')

    n = len(dias_hab)
    df['CUOTA_DIA_NUEVA'] = (df['CUOTA_MES'] - df['AVANCE_MES']) / n

    # Ordenar: FFVV → SUPERVISOR → RUTA
    orden_ffvv = {'F8': 0, 'M0': 1, 'K0': 2, 'P0': 3, 'V0': 4}
    df['_ord'] = df['FFVV'].map(orden_ffvv).fillna(9)
    df = df.sort_values(['_ord', 'SUPERVISOR', 'RUTA']).drop(columns='_ord')

    # Rutas con avance pero sin cuota (informativo)
    rutas_sin_cuota = set(df_avance['RUTA']) - set(df_cuota['RUTA'])
    if rutas_sin_cuota:
        print(f'    [WARN] {len(rutas_sin_cuota)} rutas con avance pero SIN cuota asignada: {rutas_sin_cuota}')

    # Resumen
    total_cuota   = df['CUOTA_MES'].sum()
    total_avance  = df['AVANCE_MES'].sum()
    total_saldo   = total_cuota - total_avance
    total_cuota_d = total_saldo / n
    pct_avance    = total_avance / total_cuota if total_cuota else 0

    print(f'\n    Cuota mes total:    S/ {total_cuota:>12,.0f}')
    print(f'    Avance acumulado:   S/ {total_avance:>12,.0f}  ({pct_avance:.1%})')
    print(f'    Saldo restante:     S/ {total_saldo:>12,.0f}')
    print(f'    Cuota dia nueva:    S/ {total_cuota_d:>12,.0f}  (÷ {n} dias habiles)')

    # Generar Excel
    print('\n[6] Generando Excel...')
    out = f'{OUT_DIR}/CuotaDia_{hoy.strftime("%Y%m")}_RECALCULADA.xlsx'
    escribir_excel(df, dias_hab, out)

    print('\nOK — Revisar y confirmar antes de actualizar CUOTAS_SSFF.xlsx')


if __name__ == '__main__':
    main()
