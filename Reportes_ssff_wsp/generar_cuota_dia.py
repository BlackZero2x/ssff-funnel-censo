"""
generar_cuota_dia.py

Recalcula CUOTA_DIA por ruta usando la formula del Gerente Comercial:

    CUOTA_DIA = (Nueva Cuota - Avance del mes) / Dias habiles restantes

Fuentes:
  - CuotaJunioV2.xlsx     → Nueva Cuota por RUTA
  - base_com (SQL)         → Avance junio por RUTA (mes='2606', sin anulados)
  - TABLAS_RUTAS.xlsx      → RUTA → VENDEDOR, SUPERVISOR
  - Dias habiles restantes → lunes-viernes desde mañana hasta 30/06,
                             excluyendo 29/06 (feriado no laborable)

Genera: CuotaDia_Junio_RECALCULADA.xlsx (para revisión)

Uso:
    python generar_cuota_dia.py
"""

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
CUOTA_PATH  = f'{BASE_DIR}/files/CuotaJunioV2.xlsx'
TABLAS_PATH = f'{BASE_DIR}/TABLAS_RUTAS.xlsx'

FERIADOS = {datetime.date(2026, 6, 29)}   # San Pedro y San Pablo


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
    """Días hábiles desde HOY hasta el 30/06/2026 inclusive.
    Lunes a sábado (weekday 0-5), excluyendo feriados en FERIADOS.
    Domingo (weekday 6) no es hábil."""
    hoy = datetime.date.today()
    fin_mes = datetime.date(2026, 6, 30)
    dias = []
    d = hoy
    while d <= fin_mes:
        if d.weekday() < 6 and d not in FERIADOS:   # lun-sab y no feriado
            dias.append(d)
        d += datetime.timedelta(days=1)
    return dias


# ── Carga de datos ────────────────────────────────────────────────────────────

def cargar_avance_sql() -> pd.DataFrame:
    """Avance junio por ruta desde base_com."""
    query = """
        SELECT
               [mes]
              ,[ccod_ruta]
              ,SUM([monto]) AS AVANCE
          FROM [eAuren].[dbo].[base_com]
         WHERE mes='2606' AND cstatus != 'A'
         GROUP BY [mes], [ccod_ruta]
    """
    print("   Consultando avance junio en SQL...")
    conn = conectar_sql()
    try:
        df = pd.read_sql(query, conn)
    finally:
        conn.close()
    df['ccod_ruta'] = df['ccod_ruta'].astype(str).str.strip()
    df = df.rename(columns={'ccod_ruta': 'RUTA', 'AVANCE': 'AVANCE_MES'})
    print(f"   {len(df)} rutas con avance en junio")
    return df[['RUTA', 'AVANCE_MES']]


def cargar_cuota() -> pd.DataFrame:
    """Nueva Cuota mensual por ruta desde CuotaJunioV2.xlsx."""
    df = pd.read_excel(CUOTA_PATH, dtype={'RUTA': str})
    df['RUTA'] = df['RUTA'].str.strip()
    return df[['RUTA', 'Nueva Cuota']].rename(columns={'Nueva Cuota': 'CUOTA_MES'})


def cargar_rutas() -> pd.DataFrame:
    """RUTA → VENDEDOR, SUPERVISOR, FFVV desde TABLAS_RUTAS."""
    df = pd.read_excel(TABLAS_PATH, sheet_name='RUTA_ACTUAL', dtype={'RUTA': str})
    df['RUTA'] = df['RUTA'].str.strip()
    cols = ['RUTA', 'VENDEDOR', 'SUPERVISOR']
    # FFVV se infiere de la ruta (F8, M0, K0, P0, V0)
    df['FFVV'] = df['RUTA'].str[:1].map(
        {'F': 'F8', 'M': 'M0', 'K': 'K0', 'P': 'P0', 'V': 'V0'}
    ).fillna('?')
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

    hoy_str   = datetime.date.today().strftime('%d/%m/%Y')
    n_dias    = len(dias_hab)
    ult_dia   = dias_hab[-1].strftime('%d/%m') if dias_hab else '-'
    primer    = dias_hab[0].strftime('%d/%m') if dias_hab else '-'

    # ── Fila 1: título
    headers_title = (
        f'CUOTA DIA RECALCULADA — Junio 2026   |   '
        f'Generado: {hoy_str}   |   '
        f'Dias habiles restantes: {n_dias} ({primer} al {ult_dia}, excl. 29/06)'
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

def calcular_cuota_dia_vendedor() -> dict:
    """Devuelve {RUTA: cuota_dia} con la fórmula recalculada del Gerente.

    CUOTA_DIA = (Nueva Cuota - Avance del mes) / Dias habiles restantes
    Dias habiles: lun-sab desde hoy hasta 30/06, excl. 29/06 feriado.
    Si el saldo es negativo (cuota ya superada), cuota_dia = 0.
    """
    dias_hab = dias_habiles_restantes()
    n = len(dias_hab)
    if n == 0:
        return {}

    df_cuota  = cargar_cuota()
    df_avance = cargar_avance_sql()

    df = df_cuota.merge(df_avance, on='RUTA', how='left')
    df['AVANCE_MES'] = df['AVANCE_MES'].fillna(0.0)
    df['CUOTA_DIA']  = (df['CUOTA_MES'] - df['AVANCE_MES']) / n
    df['CUOTA_DIA']  = df['CUOTA_DIA'].clip(lower=0)   # no mostrar negativos

    return df.set_index('RUTA')['CUOTA_DIA'].to_dict()


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print('=' * 65)
    print('  CUOTA DIA RECALCULADA — Junio 2026')
    print('=' * 65)

    # Días hábiles restantes
    dias_hab = dias_habiles_restantes()
    hoy = datetime.date.today()
    print(f'\n[1] Dias habiles restantes desde {hoy + datetime.timedelta(1)} hasta 30/06:')
    print(f'    {len(dias_hab)} dias (excluye fines de semana + 29/06 feriado)')
    for d in dias_hab:
        print(f'    {d.strftime("%A %d/%m")}')

    if len(dias_hab) == 0:
        print('\n[ERROR] No quedan dias habiles en el mes. Abortando.')
        return

    # Cargar cuota mensual
    print('\n[2] Cargando cuota mensual (CuotaJunioV2.xlsx)...')
    df_cuota = cargar_cuota()
    print(f'    {len(df_cuota)} rutas con cuota')

    # Cargar avance SQL
    print('\n[3] Cargando avance junio desde SQL...')
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
    out = f'{OUT_DIR}/CuotaDia_Junio_RECALCULADA.xlsx'
    escribir_excel(df, dias_hab, out)

    print('\nOK — Revisar y confirmar antes de actualizar CuotaJunioV2.xlsx')


if __name__ == '__main__':
    main()
