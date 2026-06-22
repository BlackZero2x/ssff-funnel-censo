"""
generar_corte_ventas.py

Reporte de CORTE DE VENTAS del día por hora, con 3 vistas:
  - GENERAL    (A2:H8)
  - ZONAL      (K2:U12)  ← por ZONA2 (merge ruta↔TABLA_RUTAS)
  - SUPERVISOR (X2:AH14)

Cada tabla compara D (hoy) vs D-7 vs D-14 (mismo día de semana), con corte
acumulado hasta la hora de ejecución (8AM..6PM).

Reutiliza la conexión SQL y el SP spPreventa_efectividad_dia de
generar_funnel_preventa.py (misma fuente de datos; campo horaTP + monto).

Uso:
    python generar_corte_ventas.py            # corte = hora actual
    python generar_corte_ventas.py --hora 10  # forzar corte 10AM (pruebas)
"""
# ── Librería estándar ──────────────────────────────────────────────────────────
import argparse
import datetime
import os
import pickle
import warnings
from pathlib import Path

# ── Terceros ───────────────────────────────────────────────────────────────────
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import IconSetRule
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

# ════════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN Y CREDENCIALES (mismo patrón que generar_funnel_preventa.py)
# ════════════════════════════════════════════════════════════════════════════════

BASE_DIR    = 'C:/proyectos/SSFF'
OUT_DIR     = f'{BASE_DIR}/Reportes_ssff_wsp'
SCRIPT_DIR  = Path(__file__).parent.absolute()
CACHE_DIR   = f'{OUT_DIR}/cache'
TABLAS_PATH = f'{BASE_DIR}/TABLAS_RUTAS.xlsx'

DIAS_SEMANA = {1: 'Lunes', 2: 'Martes', 3: 'Miércoles',
               4: 'Jueves', 5: 'Viernes', 6: 'Sábado', 7: 'Domingo'}

# Etiquetas de corte horario: 8AM..6PM. Índice = hora clamp(8..18).
HORA_LBL = {8: '8AM', 9: '9AM', 10: '10AM', 11: '11AM', 12: '12PM',
            13: '1PM', 14: '2PM', 15: '3PM', 16: '4PM', 17: '5PM', 18: '6PM'}

# Orden fijo de zonas (filas K6:K13 del EJEMPLO2)
ZONAS_ORDEN = ['Lima Este', 'Casco', 'May. SSFF', 'May. SSFF Casco', 'Exclusivo',
               'Casa Reposo', 'Rinti', 'Verdum']

CUOTA_PATH = f'{BASE_DIR}/files/CuotaJunioV2.xlsx'

UMBRAL_ICON = 0.03   # ±3% → zona de alerta (!)


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

def _env_req(clave):
    valor = _env.get(clave)
    if not valor:
        raise EnvironmentError(f"Variable '{clave}' no encontrada en .env")
    return valor


def conectar_sql():
    import pyodbc
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={_env_req('SQL_SERVER')};"
        f"DATABASE={_env_req('SQL_DATABASE')};"
        f"UID={_env_req('SQL_USER')};"
        f"PWD={_env_req('SQL_PASSWORD')};"
        f"TrustServerCertificate=yes;"
    )
    return pyodbc.connect(conn_str)


# ════════════════════════════════════════════════════════════════════════════════
# 1. CARGA DE DATOS (con caché para días pasados)
# ════════════════════════════════════════════════════════════════════════════════

def cargar_preventa_dia(conn, fecha: datetime.date) -> pd.DataFrame:
    """Ejecuta el SP de efectividad para una fecha y devuelve solo pedidos con compra.
    Mantiene columnas mínimas: ruta, supervisor, monto, horaTP.

    Nota: Si fecha==hoy, ejecuta sin parametros (SP devuelve datos de hoy por default).
    Si fecha es otra, pasa fecha y dia_semana."""
    hoy = datetime.date.today()

    if fecha == hoy:
        # Para hoy: ejecutar sin parametros (SP default a hoy)
        query = "EXEC [eAuren].[dbo].[spPreventa_efectividad_dia]"
    else:
        # Para otras fechas: pasar parametros
        dia_semana = fecha.weekday() + 1
        query = (f"EXEC [eAuren].[dbo].[spPreventa_efectividad_dia] "
                 f"'{fecha.strftime('%Y%m%d')}', '{dia_semana}'")

    df = pd.read_sql(query, conn)
    df = df[df['monto'].notna()].copy()
    df['monto'] = pd.to_numeric(df['monto'], errors='coerce')
    df['horaTP'] = pd.to_datetime(df['horaTP'], errors='coerce')
    df = df[df['horaTP'].notna() & df['monto'].notna()]
    return df[['ruta', 'supervisor', 'monto', 'horaTP']]


def cargar_dia_con_cache(conn, fecha: datetime.date, etiqueta: str) -> pd.DataFrame:
    """Días pasados (D-7, D-14): se consultan 1 sola vez al día y se cachean.
    El nombre de la caché lleva la fecha de HOY para invalidarla cada jornada."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    hoy_str = datetime.date.today().strftime('%Y%m%d')
    cache_file = f"{CACHE_DIR}/corte_{etiqueta}_{hoy_str}.pkl"

    if os.path.exists(cache_file):
        with open(cache_file, 'rb') as f:
            df = pickle.load(f)
        print(f"   [{etiqueta}] {fecha} — desde caché ({len(df):,} pedidos)")
        return df

    print(f"   [{etiqueta}] {fecha} — consultando SQL...")
    df = cargar_preventa_dia(conn, fecha)
    with open(cache_file, 'wb') as f:
        pickle.dump(df, f)
    print(f"      >> {len(df):,} pedidos cacheados")
    return df


def cargar_mapa_zonas() -> pd.Series:
    """ruta → ZONA2 desde TABLA_RUTAS (hoja RUTA_ACTUAL).
    V001/V002 (Verdum) no están en maestro → se fuerzan."""
    df = pd.read_excel(TABLAS_PATH, sheet_name='RUTA_ACTUAL', dtype={'RUTA': str})
    mapa = df.set_index('RUTA')['ZONA2'].to_dict()
    for r in ('V001', 'V002'):
        mapa.setdefault(r, 'Verdum')
    return mapa


def calcular_cuota_supervisor(cuota_vend: dict, ruta_por_vendedor: dict) -> dict:
    """Calcula cuota por supervisor como suma de vendedores.

    Esto asegura que SUPERVISOR sea suma de sus VENDEDORES (granularidad).

    Args:
        cuota_vend: {RUTA: cuota_dia} desde calcular_cuota_dia_vendedor()
        ruta_por_vendedor: {VENDEDOR: [rutas]}

    Retorna: {supervisor: cuota_dia_total}
    """
    df_rutas = pd.read_excel(TABLAS_PATH, sheet_name='RUTA_ACTUAL', dtype={'RUTA': str})

    # Crear mapeo VENDEDOR → SUPERVISOR
    vend_sup = df_rutas[['VENDEDOR', 'SUPERVISOR']].drop_duplicates().set_index('VENDEDOR')['SUPERVISOR'].to_dict()

    # Sumar cuota por vendedor, luego agrupar por supervisor
    cuota_vendedor = {}
    for vendedor, rutas in ruta_por_vendedor.items():
        cuota_vendedor[vendedor] = sum(cuota_vend.get(rt, 0.0) for rt in rutas)

    # Agrupar vendedor → supervisor
    cuota_sup = {}
    for vendedor, cuota in cuota_vendedor.items():
        supervisor = vend_sup.get(vendedor, 'SIN ASIGNAR')
        cuota_sup[supervisor] = cuota_sup.get(supervisor, 0.0) + cuota

    return cuota_sup


def calcular_cuota_zonal(cuota_vend: dict, ruta_por_vendedor: dict) -> dict:
    """Calcula cuota por ZONAL sumando desde cuota_vend.

    Args:
        cuota_vend: {RUTA: cuota_dia} desde calcular_cuota_dia_vendedor()
        ruta_por_vendedor: {VENDEDOR: [rutas]}

    Retorna: {zonal: cuota_dia_total}
    """
    df_rutas = pd.read_excel(TABLAS_PATH, sheet_name='RUTA_ACTUAL', dtype={'RUTA': str})

    # Crear mapeo RUTA → ZONA2
    ruta_zona = df_rutas.set_index('RUTA')['ZONA2'].to_dict()

    # Sumar cuota por ZONA2
    cuota_zonal = {}
    for ruta, cuota_dia in cuota_vend.items():
        zona = ruta_zona.get(ruta, '(sin zona)')
        cuota_zonal[zona] = cuota_zonal.get(zona, 0.0) + cuota_dia

    return cuota_zonal


# ════════════════════════════════════════════════════════════════════════════════
# 2. TRANSFORMACIÓN: etiqueta de hora + corte acumulado
# ════════════════════════════════════════════════════════════════════════════════

def agregar_hora_lbl(df: pd.DataFrame) -> pd.DataFrame:
    """Campo hora_h = clamp(HOUR(horaTP), 8, 18). Madrugada→8, noche→18."""
    df = df.copy()
    h = df['horaTP'].dt.hour.clip(lower=8, upper=18)
    df['hora_h'] = h
    return df


def filtrar_corte(df: pd.DataFrame, hora_corte: int) -> pd.DataFrame:
    """Acumulado: todos los pedidos con hora_h <= hora_corte."""
    return df[df['hora_h'] <= hora_corte]


def indicadores(serie_monto: pd.Series) -> dict:
    """Pedidos = recuento, Soles = suma, Ticket Prom. = promedio."""
    return {
        'Pedidos': int(serie_monto.count()),
        'Soles':   float(serie_monto.sum()),
        'Ticket':  float(serie_monto.mean()) if len(serie_monto) else 0.0,
    }


# ════════════════════════════════════════════════════════════════════════════════
# 3. HELPERS DE ESTILO (colores resueltos del EJEMPLO1.xlsx)
# ════════════════════════════════════════════════════════════════════════════════

FUENTE = 'Aptos Narrow'

# Colores resueltos (theme+tint → RGB)
C_TIT_GEN  = '5B9BD5'   # título GENERAL (accent5)
C_TIT_ZON  = '1F3864'   # título ZONAL (azul marino)
C_TIT_SUP  = '006C50'   # título SUPERVISOR (verde oscuro)
C_CORTE    = 'FFFFCC'   # celdas Corte/hora
C_HDR_GEN  = 'DEEBF7'   # header fila4 GENERAL
C_DIF_GEN  = '9DC3E6'   # header dif GENERAL
C_HDR_ZON  = 'DAE3F3'   # header fila4 ZONAL (azul claro)
C_DIF_ZON  = '1F3864'   # header dif ZONAL (azul marino, texto blanco)
C_HDR_SUP  = 'CCFFCC'   # header fila4 SUPERVISOR (verde claro)
C_DIF_SUP  = '006C50'   # header dif SUPERVISOR (texto blanco)
C_DATA_GEN = 'F2F2F2'   # relleno datos GENERAL
C_DATA_ZS  = 'D9D9D9'   # relleno datos ZONAL/SUPERVISOR
C_BLANCO   = 'FFFFFF'
C_NEGRO    = '000000'
C_VEN_TIT  = '403151'   # título tabla VENDEDOR (THEME:7 tint=-0.5, morado oscuro)
C_VEN_TOT  = '403151'   # TOTAL tabla VENDEDOR
C_VEN_HDR  = 'DDD4E9'   # banda fecha VENDEDOR (THEME:7 tint=+0.8, lila claro)

# Formato moneda soles (idéntico al EJEMPLO en columnas Diferencia)
FMT_SOLES = '_-"S/"\\ * #,##0.00_-;\\-"S/"\\ * #,##0.00_-;_-"S/"\\ * "-"??_-;_-@_-'
FMT_NUM_SIN_DEC = '#,##0'  # Soles: sin decimales, coma miles
FMT_TICKET = '0.00'        # Ticket Prom.: 2 decimales
FMT_DIF_NUM = '#,##0'      # Diferencia (números): sin decimales
FMT_PCT   = '0%'


def fill(hex_color):
    return PatternFill('solid', fgColor=hex_color)

def fnt(bold=False, italic=False, size=11, color=C_NEGRO):
    return Font(name=FUENTE, bold=bold, italic=italic, size=size, color=color)

def aln(h='center', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def brd(left=False, right=False, top=False, bottom=False):
    s = Side(style='thin')
    n = Side(style=None)
    return Border(left=s if left else n, right=s if right else n,
                  top=s if top else n, bottom=s if bottom else n)

def brd_all():
    s = Side(style='thin')
    return Border(left=s, right=s, top=s, bottom=s)


# ════════════════════════════════════════════════════════════════════════════════
# 4. ESCRITURA DE TABLAS
# ════════════════════════════════════════════════════════════════════════════════

def _set(ws, r, c, val, **kw):
    cell = ws.cell(row=r, column=c, value=val)
    if 'font' in kw:   cell.font = kw['font']
    if 'fill' in kw:   cell.fill = kw['fill']
    if 'align' in kw:  cell.alignment = kw['align']
    if 'border' in kw: cell.border = kw['border']
    if 'fmt' in kw:    cell.number_format = kw['fmt']
    return cell


def _cabecera_corte(ws, col_corte, col_hora, hora_lbl):
    """Celdas 'Corte' | 'NAM' en esquina superior derecha de cada tabla."""
    _set(ws, 2, col_corte, 'Corte',
         font=fnt(bold=True, italic=True, size=12), fill=fill(C_CORTE), align=aln())
    _set(ws, 2, col_hora, hora_lbl,
         font=fnt(bold=True, italic=True, size=12), fill=fill(C_CORTE), align=aln())


def _icon_rule():
    """Set de 3 símbolos en círculo (✓ ! ✗) con umbrales ±3%.
    Excel '3Symbols' (circled): icono0=rojo✗, icono1=amarillo!, icono2=verde✓.
    Umbrales en porcentaje: <-3% rojo, -3%..+3% amarillo, >+3% verde."""
    return IconSetRule(
        icon_style='3Symbols',
        type='num',
        values=[-9999, -UMBRAL_ICON, UMBRAL_ICON],
        showValue=True,
    )


# ── HELPER DE BORDES (nivel módulo, usado por todas las tablas) ─────────────────

def _apply_border(cell, left=False, right=False, top=False, bottom=False):
    """Mezcla los lados solicitados con los ya existentes en la celda."""
    _thin = Side(style='thin')
    b = cell.border
    cell.border = Border(
        left   = _thin if left   else b.left,
        right  = _thin if right  else b.right,
        top    = _thin if top    else b.top,
        bottom = _thin if bottom else b.bottom,
    )


# ── TABLA GENERAL (A2:H8) ───────────────────────────────────────────────────────

def escribir_general(ws, datos, fechas, nombre_dia, hora_lbl):
    """datos: dict {indicador: {'d14':v,'d7':v,'d':v}} para Pedidos/Soles/Ticket."""
    # Título: texto en A2 (BLANCO) + relleno de color en A2:F2 (sin merge, como el EJEMPLO)
    for cc in range(1, 7):
        _set(ws, 2, cc, 'RESUMEN GENERAL' if cc == 1 else None,
             font=fnt(bold=True, italic=True, size=14, color='FFFFFF'),
             fill=fill(C_TIT_GEN), align=aln('left') if cc == 1 else aln())
    _cabecera_corte(ws, 7, 8, hora_lbl)

    # Fila 4: bandas superiores
    ws.merge_cells('B4:D4')
    _set(ws, 4, 2, f'{nombre_dia} ', font=fnt(bold=True, italic=True), fill=fill(C_HDR_GEN), align=aln())
    ws.merge_cells('E4:F4')
    _set(ws, 4, 5, '[D-7] vs. [D]', font=fnt(bold=True, italic=True), fill=fill(C_DIF_GEN), align=aln())
    ws.merge_cells('G4:H4')
    _set(ws, 4, 7, '[D-14] vs. [D]', font=fnt(bold=True, italic=True), fill=fill(C_DIF_GEN), align=aln())

    # Fila 5: encabezados de columna — texto NEGRO
    hdr = ['Indicador', f'({fechas["d14"]}) [D-14]', f'({fechas["d7"]}) [D-7]',
           f'({fechas["d"]}) [D]', 'Diferencia', '%Dif.', 'Diferencia', '%Dif.']
    for j, txt in enumerate(hdr, 1):
        b = (j == 1)
        _set(ws, 5, j, txt,
             font=fnt(bold=b, italic=True, color=C_NEGRO), align=aln() if j > 1 else aln('left'),
             border=brd_all())

    # Filas 6-8: Pedidos, Soles, Ticket Prom. — DAMERO (par=gris, impar=blanco)
    indic = [('Pedidos', 'Pedidos', None), ('Soles', 'Soles', FMT_NUM_SIN_DEC),
             ('Ticket Prom.', 'Ticket', FMT_TICKET)]
    for i, (lbl, key, fmt) in enumerate(indic):
        r = 6 + i
        es_par = (r % 2 == 0)  # fila 6,8=par=gris; fila 7=impar=blanco
        color_fila = C_DATA_GEN if es_par else C_BLANCO
        _set(ws, r, 1, lbl, font=fnt(italic=True), fill=fill(color_fila), align=aln('left'))
        v14, v7, vd = datos[key]['d14'], datos[key]['d7'], datos[key]['d']
        for j, val in zip((2, 3, 4), (v14, v7, vd)):
            _set(ws, r, j, round(val, 2) if key == 'Pedidos' else val,
                 font=fnt(), fill=fill(color_fila), align=aln(),
                 border=brd(left=(j == 2), right=(j == 4)), fmt=fmt or 'General')
        # Diferencias como fórmulas vivas (idéntico al EJEMPLO: =+D-C, =E/C ...) — con DAMERO
        # Formato Dif.: #,##0 (sin decimales)
        fmt_dif = FMT_DIF_NUM if key != 'Pedidos' else 'General'
        _set(ws, r, 5, f'=+D{r}-C{r}', font=fnt(), fill=fill(color_fila), align=aln(),
             border=brd(left=True), fmt=fmt_dif)
        _set(ws, r, 6, f'=IFERROR(E{r}/C{r},0)', font=fnt(), fill=fill(color_fila),
             align=aln('right'), border=brd(right=True), fmt=FMT_PCT)
        _set(ws, r, 7, f'=+D{r}-B{r}', font=fnt(), fill=fill(color_fila), align=aln(),
             border=brd(left=True), fmt=fmt_dif)
        _set(ws, r, 8, f'=IFERROR(G{r}/B{r},0)', font=fnt(), fill=fill(color_fila),
             align=aln('right'), border=brd(right=True), fmt=FMT_PCT)

    # C5:C8 — left+right border
    for rr in range(5, 9):
        _apply_border(ws.cell(rr, 3), left=True, right=True)

    # Iconos en %Dif (F6:F8 y H6:H8)
    ws.conditional_formatting.add('F6:F8', _icon_rule())
    ws.conditional_formatting.add('H6:H8', _icon_rule())

    # Anchos (B,C,D = 14; Dif. = 14, %Dif. = 7.57)
    anchos = {'A': 12.57, 'B': 14.0, 'C': 14.0, 'D': 14.0,
              'E': 14.0, 'F': 7.57, 'G': 14.0, 'H': 7.57}
    for col, w in anchos.items():
        ws.column_dimensions[col].width = w
    ws.row_dimensions[2].height = 18.75


# ── TABLA POR CATEGORÍA (ZONAL / SUPERVISOR) ────────────────────────────────────

def escribir_categoria(ws, col0, titulo, tit_color, hdr_color, dif_color, dif_font,
                       data_fill, etiqueta_col, filas, datos, fechas, nombre_dia,
                       hora_lbl, n_filas_fijo, cuota_sup=None, cuota_zonal=None):
    """Genérico para ZONAL (col0=11/K) y SUPERVISOR (col0=24/X).
    Para SUPERVISOR y ZONAL: agrega columnas Cuota_Día y %Avance.
    datos: dict {fila_label: {'d14':(ped,sol),'d7':(..),'d':(..)}}."""
    c = col0
    L = get_column_letter

    # Estructuras: SUPERVISOR y ZONAL ahora ambas tienen Cuota_Día y %Avance
    es_supervisor = (etiqueta_col == 'Supervisor')
    es_zonal = (etiqueta_col == 'ZONAL')

    if es_supervisor or es_zonal:
        # [X][Y Z][AA AB][AC AD][AE AF][AG AH] (SUPERVISOR)
        # [K][L M][N O][P Q][R S][T U] (ZONAL con cuota)
        # Label, P14 S14, P7 S7, Pd Sd, Cuota %Avance, %Dif7, %Dif14
        c_lbl = c
        c_p14, c_s14 = c+1, c+2
        c_p7,  c_s7  = c+3, c+4
        c_pd,  c_sd  = c+5, c+6
        c_cuota = c+7
        c_pct_avance = c+8
        c_pct7 = c+9      # SOLO % para [D-7] vs [D]
        c_pct14 = c+10    # SOLO % para [D-14] vs [D]
    else:
        # Caso no usado actualmente
        c_lbl = c
        c_p14, c_s14 = c+1, c+2
        c_p7,  c_s7  = c+3, c+4
        c_pd,  c_sd  = c+5, c+6
        c_dif7, c_pct7 = c+7, c+8
        c_dif14, c_pct14 = c+9, c+10

    # Título: texto en col0 + relleno
    # SUPERVISOR: [X][Y Z][AA AB][AC AD][AE AF] = 9 columnas (c a c+8), Corte en AG-AH (c+9, c+10)
    # ZONAL: [K][L M][N O][P Q][R S] = 9 columnas (c a c+8), Corte en T-U (c+9, c+10)
    relleno_fin = c + 9  # siempre hasta col+9 para ambas
    for cc in range(c, relleno_fin):
        _set(ws, 2, cc, titulo if cc == c else None,
             font=fnt(bold=True, italic=True, size=14, color=C_BLANCO),
             fill=fill(tit_color), align=aln('left') if cc == c else aln())
    # Corte en últimas 2 columnas
    _cabecera_corte(ws, relleno_fin, relleno_fin + 1, hora_lbl)

    # Fila 4: bandas de fecha por día (cada una merge de 2 cols) + extras
    for (cc, txt) in [(c_p14, f'{nombre_dia} ({fechas["d14"]}) [D-14]'),
                      (c_p7,  f'{nombre_dia} ({fechas["d7"]}) [D-7]'),
                      (c_pd,  f'{nombre_dia} ({fechas["d"]}) [D]')]:
        ws.merge_cells(start_row=4, start_column=cc, end_row=4, end_column=cc+1)
        _set(ws, 4, cc, txt, font=fnt(italic=True, size=10), fill=fill(hdr_color), align=aln())

    # Para SUPERVISOR y ZONAL: agregar "Seguimiento del día" (merge AE4:AF4 o R4:S4)
    if es_supervisor or es_zonal:
        ws.merge_cells(start_row=4, start_column=c_cuota, end_row=4, end_column=c_pct_avance)
        # Color de texto: verde para SUPERVISOR, marino para ZONAL
        color_texto = '006C50' if es_supervisor else '1F3864'
        _set(ws, 4, c_cuota, 'Seguimiento del día',
             font=fnt(bold=True, italic=True, size=10, color=color_texto),
             fill=fill('FFFFFF'), align=aln())
        # % Diferencia: merge AG4:AH4 (SUPERVISOR) o T4:U4 (ZONAL), SIN fill
        ws.merge_cells(start_row=4, start_column=c_pct7, end_row=4, end_column=c_pct14)
        _set(ws, 4, c_pct7, '% Diferencia',
             font=fnt(bold=True, italic=True, size=10, color=color_texto),
             fill=fill('FFFFFF'), align=aln())

    # Fila 5: encabezados
    if es_supervisor or es_zonal:
        # SUPERVISOR y ZONAL: SOLO % en diferencias, sin columnas "Diferencia" numérica
        headers = [(c_lbl, 'Supervisor' if es_supervisor else 'ZONAL'), (c_p14, 'Pedidos'), (c_s14, 'Soles'),
                   (c_p7, 'Pedidos'), (c_s7, 'Soles'), (c_pd, 'Pedidos'), (c_sd, 'Soles'),
                   (c_cuota, 'Cuota_Día'), (c_pct_avance, '%Avance'),
                   (c_pct7, '[D-7] vs. [D]'), (c_pct14, '[D-14] vs. [D]')]
    else:
        headers = [(c_lbl, etiqueta_col), (c_p14, 'Pedidos'), (c_s14, 'Soles'),
                   (c_p7, 'Pedidos'), (c_s7, 'Soles'), (c_pd, 'Pedidos'), (c_sd, 'Soles'),
                   (c_dif7, 'Diferencia'), (c_pct7, '%Dif.'),
                   (c_dif14, 'Diferencia'), (c_pct14, '%Dif.')]
    for cc, txt in headers:
        _set(ws, 5, cc, txt, font=fnt(bold=True), align=aln(), border=brd_all())

    # Filas de datos — DAMERO (par=gris, impar=blanco)
    for i, fila_lbl in enumerate(filas):
        r = 6 + i
        es_par = (r % 2 == 0)
        color_fila_cat = data_fill if es_par else C_BLANCO
        d = datos.get(fila_lbl, {'d14': (0, 0.0), 'd7': (0, 0.0), 'd': (0, 0.0)})
        p14, s14 = d['d14']; p7, s7 = d['d7']; pd_, sd = d['d']
        _set(ws, r, c_lbl, fila_lbl, font=fnt(), fill=fill(color_fila_cat), align=aln('left'),
             border=brd(left=True))
        vals = [(c_p14, p14, 'General'), (c_s14, round(s14, 2), FMT_NUM_SIN_DEC),
                (c_p7, p7, 'General'),   (c_s7, round(s7, 2), FMT_NUM_SIN_DEC),
                (c_pd, pd_, 'General'),  (c_sd, round(sd, 2), FMT_NUM_SIN_DEC)]  # Soles sin decimales
        for cc, val, fmt in vals:
            _set(ws, r, cc, val, font=fnt(), fill=fill(color_fila_cat), align=aln(), fmt=fmt)

        # Para SUPERVISOR: agregar Cuota_Día y %Avance
        if es_supervisor and cuota_sup:
            cuota_dia = cuota_sup.get(fila_lbl, 0.0)
            _set(ws, r, c_cuota, cuota_dia, font=fnt(), fill=fill(color_fila_cat),
                 align=aln(), fmt=FMT_NUM_SIN_DEC)
            # %Avance = IFERROR(SD/CUOTA, "-")
            sd_c = L(c_sd)
            cuota_c = L(c_cuota)
            _set(ws, r, c_pct_avance, f'=IFERROR({sd_c}{r}/{cuota_c}{r},"-")',
                 font=fnt(), fill=fill(color_fila_cat), align=aln(), fmt=FMT_PCT)

            # %Dif7 = IFERROR((SD-S7)/S7,"-")  %Dif14 = IFERROR((SD-S14)/S14,"")
            s7_c, s14_c = L(c_s7), L(c_s14)
            _set(ws, r, c_pct7,
                 f'=IFERROR(({sd_c}{r}-{s7_c}{r})/{s7_c}{r},"-")',
                 font=fnt(), fill=fill(color_fila_cat), align=aln(), fmt=FMT_PCT)
            _set(ws, r, c_pct14,
                 f'=IFERROR(({sd_c}{r}-{s14_c}{r})/{s14_c}{r},"")',
                 font=fnt(), fill=fill(color_fila_cat), align=aln(), fmt=FMT_PCT)
        # Para ZONAL: agregar Cuota_Día y %Avance
        elif es_zonal and cuota_zonal:
            cuota_dia = cuota_zonal.get(fila_lbl, 0.0)
            _set(ws, r, c_cuota, cuota_dia, font=fnt(), fill=fill(color_fila_cat),
                 align=aln(), fmt=FMT_NUM_SIN_DEC)
            # %Avance = IFERROR(SD/CUOTA, "-")
            sd_c = L(c_sd)
            cuota_c = L(c_cuota)
            _set(ws, r, c_pct_avance, f'=IFERROR({sd_c}{r}/{cuota_c}{r},"-")',
                 font=fnt(), fill=fill(color_fila_cat), align=aln(), fmt=FMT_PCT)

            # %Dif7 = IFERROR((SD-S7)/S7,"-")  %Dif14 = IFERROR((SD-S14)/S14,"")
            s7_c, s14_c = L(c_s7), L(c_s14)
            _set(ws, r, c_pct7,
                 f'=IFERROR(({sd_c}{r}-{s7_c}{r})/{s7_c}{r},"-")',
                 font=fnt(), fill=fill(color_fila_cat), align=aln(), fmt=FMT_PCT)
            _set(ws, r, c_pct14,
                 f'=IFERROR(({sd_c}{r}-{s14_c}{r})/{s14_c}{r},"")',
                 font=fnt(), fill=fill(color_fila_cat), align=aln(), fmt=FMT_PCT)

    last_row = 6 + len(filas) - 1
    r_tot = last_row + 1

    # ── Fila TOTAL ─────────────────────────────────────────────────────────────
    _set(ws, r_tot, c_lbl, 'TOTAL',
         font=fnt(bold=True, color=C_BLANCO), fill=fill(tit_color),
         align=aln('left'), border=brd_all())
    for cc in (c_p14, c_s14, c_p7, c_s7, c_pd, c_sd):
        col_l = L(cc)
        _set(ws, r_tot, cc, f'=SUM({col_l}6:{col_l}{last_row})',
             font=fnt(bold=True, color=C_BLANCO), fill=fill(tit_color),
             align=aln(), fmt=FMT_NUM_SIN_DEC, border=brd_all())
    if es_supervisor or es_zonal:
        # Cuota total (SUPERVISOR y ZONAL)
        cuota_c = L(c_cuota)
        _set(ws, r_tot, c_cuota, f'=SUM({cuota_c}6:{cuota_c}{last_row})',
             font=fnt(bold=True, color=C_BLANCO), fill=fill(tit_color),
             align=aln(), fmt=FMT_NUM_SIN_DEC, border=brd_all())
        # %Avance total
        sd_c = L(c_sd)
        _set(ws, r_tot, c_pct_avance,
             f'=IFERROR({sd_c}{r_tot}/{cuota_c}{r_tot},"-")',
             font=fnt(bold=True, color=C_BLANCO), fill=fill(tit_color),
             align=aln(), fmt=FMT_PCT, border=brd_all())
        # %Dif7 y %Dif14 totales
        s7_c, s14_c = L(c_s7), L(c_s14)
        _set(ws, r_tot, c_pct7,
             f'=IFERROR(({sd_c}{r_tot}-{s7_c}{r_tot})/{s7_c}{r_tot},"-")',
             font=fnt(bold=True, color=C_BLANCO), fill=fill(tit_color),
             align=aln(), fmt=FMT_PCT, border=brd_all())
        _set(ws, r_tot, c_pct14,
             f'=IFERROR(({sd_c}{r_tot}-{s14_c}{r_tot})/{s14_c}{r_tot},"")',
             font=fnt(bold=True, color=C_BLANCO), fill=fill(tit_color),
             align=aln(), fmt=FMT_PCT, border=brd_all())

    # Iconos en %Dif — incluye la fila TOTAL
    ws.conditional_formatting.add(f'{L(c_pct7)}6:{L(c_pct7)}{r_tot}', _icon_rule())
    ws.conditional_formatting.add(f'{L(c_pct14)}6:{L(c_pct14)}{r_tot}', _icon_rule())

    # APLICAR BORDES ESPECÍFICOS (r_tot incluye la fila TOTAL)
    if es_supervisor or es_zonal:
        # 1) Encabezado — outside border
        for cc in range(c_lbl, c_pct14 + 1):
            _apply_border(ws.cell(5, cc), left=True, right=True, top=True, bottom=True)
        # 2) left border en columna etiqueta (filas de datos + TOTAL)
        for rr in range(6, r_tot + 1):
            _apply_border(ws.cell(rr, c_lbl), left=True)
        # 3) Right borders en columnas de cierre de par, filas 4..r_tot
        for rr in range(4, r_tot + 1):
            _apply_border(ws.cell(rr, c_s14),        right=True)
            _apply_border(ws.cell(rr, c_s7),         right=True)
            _apply_border(ws.cell(rr, c_sd),         right=True)
            _apply_border(ws.cell(rr, c_pct_avance), right=True)
            _apply_border(ws.cell(rr, c_pct14),      right=True)
        # 4) right border en columna label filas 4..r_tot
        for rr in range(4, r_tot + 1):
            _apply_border(ws.cell(rr, c_lbl), right=True)

    # Anchos (SUPERVISOR y ZONAL comparten estructura de columnas)
    if es_supervisor or es_zonal:
        ws.column_dimensions[L(c_lbl)].width = 21.71 if es_supervisor else 16.0
        for cc in (c_p14, c_s14, c_p7, c_s7, c_pd, c_sd):
            ws.column_dimensions[L(cc)].width = 9.71
        ws.column_dimensions[L(c_cuota)].width = 11.0
        ws.column_dimensions[L(c_pct_avance)].width = 9.0
        ws.column_dimensions[L(c_pct7)].width = 14.0
        ws.column_dimensions[L(c_pct14)].width = 14.0


# ── HOJA VENDEDOR ───────────────────────────────────────────────────────────────

def enriquecer_con_vendedor(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega columna 'vendedor' y 'supervisor_rutas' haciendo join con TABLAS_RUTAS."""
    df_rutas = pd.read_excel(TABLAS_PATH, sheet_name='RUTA_ACTUAL', dtype={'RUTA': str})
    df_rutas = df_rutas[['RUTA', 'VENDEDOR', 'SUPERVISOR']].copy()
    df_rutas.columns = ['ruta', 'vendedor', 'supervisor_rutas']
    df = df.merge(df_rutas, on='ruta', how='left')
    df['vendedor'] = df['vendedor'].fillna('SIN ASIGNAR')
    df['supervisor_rutas'] = df['supervisor_rutas'].fillna(df['supervisor'])
    return df


def cargar_cuota_vendedor() -> dict:
    """{RUTA: cuota_dia} desde CuotaJunioV2.xlsx."""
    df_c = pd.read_excel(CUOTA_PATH, dtype={'RUTA': str})
    return df_c.set_index('RUTA')['CUOTA_DIA'].to_dict()


def agregar_por_sup_vendedor(df_d14, df_d7, df_d):
    """Agrega por (supervisor_rutas, vendedor) para D-14, D-7 y D.

    Retorna: {supervisor: {vendedor: {'d14':(ped,sol,hora_primer,hora_ultimo),'d7':..,'d':..}}}
    """
    resultado = {}
    for etiq, df in [('d14', df_d14), ('d7', df_d7), ('d', df_d)]:
        if df is None or df.empty:
            continue
        grp = (df.groupby(['supervisor_rutas', 'vendedor'])
                 .agg(pedidos=('monto', 'count'),
                      soles=('monto', 'sum'),
                      hora_primer=('horaTP', 'min'),
                      hora_ultimo=('horaTP', 'max'))
                 .reset_index())
        for _, row in grp.iterrows():
            sup  = row['supervisor_rutas']
            vend = row['vendedor']
            resultado.setdefault(sup, {}).setdefault(
                vend, {'d14': (0, 0.0, None, None), 'd7': (0, 0.0, None, None), 'd': (0, 0.0, None, None)})

            # Extraer horas como string HH:MM (solo para la fecha actual 'd')
            hora_p = None
            hora_u = None
            if etiq == 'd':
                if pd.notna(row['hora_primer']):
                    hora_p = pd.Timestamp(row['hora_primer']).strftime('%H:%M')
                if pd.notna(row['hora_ultimo']):
                    hora_u = pd.Timestamp(row['hora_ultimo']).strftime('%H:%M')

            resultado[sup][vend][etiq] = (int(row['pedidos']), float(row['soles']), hora_p, hora_u)
    return resultado


def escribir_hoja_vendedor(ws_v, datos_sup_vend, cuota_vend,
                           ruta_por_vendedor, fechas, nombre_dia, hora_lbl):
    """Una hoja con tablas de vendedores, una por supervisor, separadas por 4 filas."""
    ws_v.sheet_view.showGridLines = False
    L = get_column_letter

    # Columnas fijas (A..M)
    COL_LBL  = 1   # A  Vendedor / RUTA
    COL_P14, COL_S14 = 2, 3   # B C
    COL_P7,  COL_S7  = 4, 5   # D E
    COL_PD,  COL_SD  = 6, 7   # F G
    COL_CUO  = 8              # H
    COL_PCT  = 9              # I  %Avance
    COL_DIF7 = 10             # J  [D-7] vs [D]
    COL_DIF14= 11             # K  [D-14] vs [D]
    COL_H1ER = 12             # L  1er. Ped.
    COL_HULT = 13             # M  Últ. Ped.

    # Anchos de columna (una sola vez, aplican a toda la hoja)
    ws_v.column_dimensions[L(COL_LBL)].width   = 27.5
    ws_v.column_dimensions[L(COL_P14)].width   = 10.0
    ws_v.column_dimensions[L(COL_S14)].width   = 10.0
    ws_v.column_dimensions[L(COL_P7)].width    = 10.0
    ws_v.column_dimensions[L(COL_S7)].width    = 10.0
    ws_v.column_dimensions[L(COL_PD)].width    = 8.0
    ws_v.column_dimensions[L(COL_SD)].width    = 10.0
    ws_v.column_dimensions[L(COL_CUO)].width   = 11.0
    ws_v.column_dimensions[L(COL_PCT)].width   = 9.0
    ws_v.column_dimensions[L(COL_DIF7)].width  = 14.0
    ws_v.column_dimensions[L(COL_DIF14)].width = 14.0
    ws_v.column_dimensions[L(COL_H1ER)].width  = 10.0
    ws_v.column_dimensions[L(COL_HULT)].width  = 10.0

    # Obtener todos los supervisores y vendedores del maestro TABLAS_RUTAS
    df_rutas = pd.read_excel(TABLAS_PATH, sheet_name='RUTA_ACTUAL', dtype={'RUTA': str})
    vendedores_por_sup_maestro = (df_rutas.groupby('SUPERVISOR')['VENDEDOR']
                                  .apply(lambda x: sorted(x.unique()))
                                  .to_dict())

    supervisores = sorted(vendedores_por_sup_maestro.keys())
    fila_inicio = 1   # fila donde empieza la primera tabla

    for sup in supervisores:
        datos_sup = datos_sup_vend.get(sup, {})
        # Usar TODOS los vendedores del supervisor, aunque no tengan ventas
        vendedores = vendedores_por_sup_maestro.get(sup, [])
        n_vend = len(vendedores)

        r0 = fila_inicio           # fila 1: título
        r3 = r0 + 2                # fila 3: bandas de fecha
        r4 = r0 + 3                # fila 4: encabezados
        r_data_ini = r0 + 4        # fila 5: primer vendedor
        r_data_fin = r_data_ini + n_vend - 1
        r_tot = r_data_fin + 1

        # ── Fila 1: título supervisor + Corte/hora (L1:M1 amarillo, J1:K1 morado)
        for cc in range(COL_LBL, COL_DIF7):  # A..I → color supervisor morado
            _set(ws_v, r0, cc,
                 sup if cc == COL_LBL else None,
                 font=fnt(bold=True, italic=True, size=13, color=C_BLANCO),
                 fill=fill(C_VEN_TIT),
                 align=aln('left') if cc == COL_LBL else aln())
        # J y K: morado oscuro #403151 (hueco)
        for cc in (COL_DIF7, COL_DIF14):
            _set(ws_v, r0, cc, None,
                 fill=fill(C_VEN_TIT), align=aln())
        # L y M: Corte y Hora (color amarillo #FFFFCC)
        _set(ws_v, r0, COL_H1ER,  'Corte',
             font=fnt(bold=True, italic=True, size=12), fill=fill(C_CORTE), align=aln())
        _set(ws_v, r0, COL_HULT, hora_lbl,
             font=fnt(bold=True, italic=True, size=12), fill=fill(C_CORTE), align=aln())

        # ── Fila 2: vacía (separación visual)
        # (sin contenido)

        # ── Fila 3: bandas de fecha (merge de pares)
        for (cc, txt) in [
            (COL_P14, f'{nombre_dia} ({fechas["d14"]}) [D-14]'),
            (COL_P7,  f'{nombre_dia} ({fechas["d7"]}) [D-7]'),
            (COL_PD,  f'{nombre_dia} ({fechas["d"]}) [D]'),
        ]:
            ws_v.merge_cells(start_row=r3, start_column=cc, end_row=r3, end_column=cc+1)
            _set(ws_v, r3, cc, txt,
                 font=fnt(italic=True, size=10),
                 fill=fill(C_VEN_HDR), align=aln())
        # "Seguimiento del dia" — merge H3:I3
        ws_v.merge_cells(start_row=r3, start_column=COL_CUO, end_row=r3, end_column=COL_PCT)
        _set(ws_v, r3, COL_CUO, 'Seguimiento del dia',
             font=fnt(bold=True, italic=True, size=10, color=C_VEN_TIT),
             fill=fill(C_BLANCO), align=aln())
        # "% Diferencia" — merge J3:K3
        ws_v.merge_cells(start_row=r3, start_column=COL_DIF7, end_row=r3, end_column=COL_DIF14)
        _set(ws_v, r3, COL_DIF7, '% Diferencia',
             font=fnt(bold=True, italic=True, size=10, color=C_VEN_TIT),
             fill=fill(C_BLANCO), align=aln())
        # "Hora de:" — merge L3:M3
        ws_v.merge_cells(start_row=r3, start_column=COL_H1ER, end_row=r3, end_column=COL_HULT)
        _set(ws_v, r3, COL_H1ER, 'Hora de:',
             font=fnt(bold=True, italic=True, size=10, color=C_VEN_TIT),
             fill=fill(C_BLANCO), align=aln())
        # Aplicar bordes a fila 3 en J,K,L,M
        for cc in (COL_DIF7, COL_DIF14, COL_H1ER, COL_HULT):
            _apply_border(ws_v.cell(r3, cc), right=True)

        # ── Fila 4: encabezados
        headers4 = [
            (COL_LBL,   'Vendedor / RUTA'),
            (COL_P14,   'Pedidos'), (COL_S14, 'Soles'),
            (COL_P7,    'Pedidos'), (COL_S7,  'Soles'),
            (COL_PD,    'Pedidos'), (COL_SD,  'Soles'),
            (COL_CUO,   'Cuota_Dia'), (COL_PCT, '%Avance'),
            (COL_DIF7,  '[D-7] vs. [D]'), (COL_DIF14, '[D-14] vs. [D]'),
            (COL_H1ER,  '1er. Ped.'), (COL_HULT, 'Últ. Ped.'),
        ]
        for cc, txt in headers4:
            # J4 y K4: encabezados con fondo blanco
            if cc in (COL_DIF7, COL_DIF14):
                _set(ws_v, r4, cc, txt,
                     font=fnt(bold=True), fill=fill(C_BLANCO), align=aln(), border=brd_all())
            else:
                _set(ws_v, r4, cc, txt,
                     font=fnt(bold=True), align=aln(), border=brd_all())

        # ── Filas de datos
        for i, vend in enumerate(vendedores):
            r = r_data_ini + i
            es_par = (r % 2 == 0)
            color_f = C_DATA_ZS if es_par else C_BLANCO

            d = datos_sup.get(vend, {'d14': (0, 0.0, None, None), 'd7': (0, 0.0, None, None), 'd': (0, 0.0, None, None)})
            p14, s14, _, _ = d['d14']
            p7,  s7,  _, _ = d['d7']
            pd_, sd, h_primer, h_ultimo  = d['d']

            # Concatenar NOMBRE - RUTA (todas las rutas del vendedor para este sup)
            rutas = ruta_por_vendedor.get(vend, [])
            ruta_lbl = rutas[0] if len(rutas) == 1 else ('/'.join(rutas) if rutas else '')
            celda_lbl = f'{vend} - {ruta_lbl}' if ruta_lbl else vend

            _set(ws_v, r, COL_LBL, celda_lbl,
                 font=fnt(), fill=fill(color_f), align=aln('left'),
                 border=brd(left=True, right=True))

            for cc, val, fmt in [
                (COL_P14, p14,         'General'),
                (COL_S14, round(s14,2), FMT_NUM_SIN_DEC),
                (COL_P7,  p7,          'General'),
                (COL_S7,  round(s7,2),  FMT_NUM_SIN_DEC),
                (COL_PD,  pd_,         'General'),
                (COL_SD,  round(sd,2),  FMT_NUM_SIN_DEC),
            ]:
                _set(ws_v, r, cc, val, font=fnt(), fill=fill(color_f),
                     align=aln(), fmt=fmt)

            # Cuota_Dia (suma de cuotas de todas las rutas del vendedor)
            cuota_dia = sum(cuota_vend.get(rt, 0.0) for rt in rutas)
            _set(ws_v, r, COL_CUO, round(cuota_dia, 2),
                 font=fnt(), fill=fill(color_f), align=aln(), fmt=FMT_NUM_SIN_DEC)

            # %Avance = G / H
            sd_c, cuo_c = L(COL_SD), L(COL_CUO)
            _set(ws_v, r, COL_PCT,
                 f'=IFERROR({sd_c}{r}/{cuo_c}{r},"-")',
                 font=fnt(), fill=fill(color_f), align=aln(), fmt=FMT_PCT,
                 border=brd(right=True))

            # %Dif [D-7] y [D-14] — con damero
            s7_c, s14_c = L(COL_S7), L(COL_S14)
            _set(ws_v, r, COL_DIF7,
                 f'=IFERROR(({sd_c}{r}-{s7_c}{r})/{s7_c}{r},"-")',
                 font=fnt(), fill=fill(color_f), align=aln(), fmt=FMT_PCT)
            _set(ws_v, r, COL_DIF14,
                 f'=IFERROR(({sd_c}{r}-{s14_c}{r})/{s14_c}{r},"-")',
                 font=fnt(), fill=fill(color_f), align=aln(),fmt=FMT_PCT,
                 border=brd(right=True))

            # 1er. Ped. y Últ. Ped. (solo para la fecha actual 'd') — con damero
            _set(ws_v, r, COL_H1ER, h_primer if h_primer else "-",
                 font=fnt(), fill=fill(color_f), align=aln())
            _set(ws_v, r, COL_HULT, h_ultimo if h_ultimo else "-",
                 font=fnt(), fill=fill(color_f), align=aln())

        # ── Fila TOTAL
        _set(ws_v, r_tot, COL_LBL, 'TOTAL',
             font=fnt(bold=True, color=C_BLANCO), fill=fill(C_VEN_TOT),
             align=aln('left'), border=brd_all())
        for cc in (COL_P14, COL_S14, COL_P7, COL_S7, COL_PD, COL_SD, COL_CUO):
            col_l = L(cc)
            _set(ws_v, r_tot, cc,
                 f'=SUM({col_l}{r_data_ini}:{col_l}{r_data_fin})',
                 font=fnt(bold=True, color=C_BLANCO), fill=fill(C_VEN_TOT),
                 align=aln(), fmt=FMT_NUM_SIN_DEC, border=brd_all())
        sd_c, cuo_c = L(COL_SD), L(COL_CUO)
        _set(ws_v, r_tot, COL_PCT,
             f'=IFERROR({sd_c}{r_tot}/{cuo_c}{r_tot},"-")',
             font=fnt(bold=True, color=C_BLANCO), fill=fill(C_VEN_TOT),
             align=aln(), fmt=FMT_PCT, border=brd_all())
        s7_c, s14_c = L(COL_S7), L(COL_S14)
        # J y K en TOTAL: morados
        _set(ws_v, r_tot, COL_DIF7,
             f'=IFERROR(({sd_c}{r_tot}-{s7_c}{r_tot})/{s7_c}{r_tot},"-")',
             font=fnt(bold=True, color=C_BLANCO), fill=fill(C_VEN_TIT),
             align=aln(), fmt=FMT_PCT, border=brd_all())
        _set(ws_v, r_tot, COL_DIF14,
             f'=IFERROR(({sd_c}{r_tot}-{s14_c}{r_tot})/{s14_c}{r_tot},"-")',
             font=fnt(bold=True, color=C_BLANCO), fill=fill(C_VEN_TIT),
             align=aln(), fmt=FMT_PCT, border=brd_all())
        # L y M en TOTAL: vacías (morado)
        for cc in (COL_H1ER, COL_HULT):
            _set(ws_v, r_tot, cc, "-",
                 font=fnt(bold=True, color=C_BLANCO), fill=fill(C_VEN_TOT),
                 align=aln(), border=brd_all())

        # ── Bordes: left/right en columna A, right en C,E,G,I,K,M (por pares + final)
        for rr in range(r3, r_tot + 1):
            _apply_border(ws_v.cell(rr, COL_LBL), left=True, right=True)
        for rr in range(r3, r_tot + 1):
            for cc in (COL_S14, COL_S7, COL_SD, COL_PCT, COL_DIF14, COL_HULT):
                _apply_border(ws_v.cell(rr, cc), right=True)
        # Columna M es el final: agregar right border adicional
        for rr in range(r3, r_tot + 1):
            _apply_border(ws_v.cell(rr, COL_HULT), right=True)

        # ── Formato condicional iconos en %Dif (datos + TOTAL)
        ws_v.conditional_formatting.add(
            f'{L(COL_DIF7)}{r_data_ini}:{L(COL_DIF7)}{r_tot}', _icon_rule())
        ws_v.conditional_formatting.add(
            f'{L(COL_DIF14)}{r_data_ini}:{L(COL_DIF14)}{r_tot}', _icon_rule())

        # ── Avanzar al siguiente bloque: TOTAL + 4 filas vacías + 1 de inicio
        fila_inicio = r_tot + 5


# ════════════════════════════════════════════════════════════════════════════════
# 5. AGREGACIÓN DE DATOS POR TABLA
# ════════════════════════════════════════════════════════════════════════════════

def agregar_general(df_d14, df_d7, df_d):
    out = {}
    for key, name in [('Pedidos', 'Pedidos'), ('Soles', 'Soles'), ('Ticket', 'Ticket')]:
        out[key] = {
            'd14': indicadores(df_d14['monto'])[key],
            'd7':  indicadores(df_d7['monto'])[key],
            'd':   indicadores(df_d['monto'])[key],
        }
    return out


def agregar_por_columna(df_d14, df_d7, df_d, col, orden=None):
    """Devuelve {valor_col: {'d14':(ped,sol),'d7':(..),'d':(..)}}."""
    def grp(df):
        g = df.groupby(col)['monto'].agg(['count', 'sum'])
        return {idx: (int(row['count']), float(row['sum'])) for idx, row in g.iterrows()}
    g14, g7, gd = grp(df_d14), grp(df_d7), grp(df_d)
    claves = orden if orden else sorted(set(g14) | set(g7) | set(gd))
    out = {}
    for k in claves:
        out[k] = {'d14': g14.get(k, (0, 0.0)),
                  'd7':  g7.get(k, (0, 0.0)),
                  'd':   gd.get(k, (0, 0.0))}
    return out


# ════════════════════════════════════════════════════════════════════════════════
# 6. MAIN
# ════════════════════════════════════════════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser(
        description='Genera corte de ventas del día con comparativos D-7 y D-14',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  python generar_corte_ventas.py                    # hoy, hora actual
  python generar_corte_ventas.py --hora 10          # hoy, corte 10AM
  python generar_corte_ventas.py --fecha 2026-06-17 --hora 14  # 17/06 a las 2PM
  python generar_corte_ventas.py --dias-atras 1 --hora 13      # ayer a la 1PM
  python generar_corte_ventas.py --hora 10 --solo-excel        # solo generar Excel, sin WhatsApp
        """
    )
    ap.add_argument('--hora', type=int, default=None,
                    help='Hora de corte 8..18 (default: hora actual)')
    ap.add_argument('--fecha', type=str, default=None,
                    help='Fecha en formato YYYY-MM-DD (default: hoy)')
    ap.add_argument('--dias-atras', type=int, default=None,
                    help='Generar N días hacia atrás (ej: 1=ayer, 2=anteayer)')
    ap.add_argument('--solo-excel', action='store_true',
                    help='Generar solo el archivo Excel sin enviar a WhatsApp')
    args = ap.parse_args()

    # Resolver fecha
    if args.fecha:
        try:
            hoy = datetime.datetime.strptime(args.fecha, '%Y-%m-%d').date()
        except ValueError:
            print(f"Error: formato de fecha inválido '{args.fecha}'. Use YYYY-MM-DD")
            exit(1)
    elif args.dias_atras is not None:
        hoy = datetime.date.today() - datetime.timedelta(days=args.dias_atras)
    else:
        hoy = datetime.date.today()

    # Resolver hora
    if args.hora is not None:
        hora_corte = min(max(args.hora, 8), 18)
    else:
        # Si es fecha pasada, usar 18 (fin del día); si es hoy, usar hora actual
        if hoy < datetime.date.today():
            hora_corte = 18  # día pasado: asumir corte fin de día
        else:
            hora_corte = min(max(datetime.datetime.now().hour, 8), 18)
    hora_lbl = HORA_LBL[hora_corte]
    nombre_dia = DIAS_SEMANA[hoy.weekday() + 1]

    d7  = hoy - datetime.timedelta(days=7)
    d14 = hoy - datetime.timedelta(days=14)
    fechas = {'d': hoy.strftime('%d/%m'),
              'd7': d7.strftime('%d/%m'),
              'd14': d14.strftime('%d/%m')}

    print('=' * 60)
    print(f'  CORTE DE VENTAS — {nombre_dia} {hoy.strftime("%d/%m/%Y")} — Corte {hora_lbl}')
    print('=' * 60)

    conn = conectar_sql()
    try:
        print('\n[1] Cargando datos...')
        # Reintentar hasta 5 veces para hoy (SP tarda ~5s, datos pueden no estar listos)
        df_d = None
        max_intentos = 5
        for intento in range(1, max_intentos + 1):
            df_d = cargar_preventa_dia(conn, hoy)
            if len(df_d) > 0:
                print(f"   [Intento {intento}] Datos de hoy OK - {len(df_d):,} pedidos")
                break
            if intento < max_intentos:
                espera = 3
                print(f"   [Intento {intento}/{max_intentos}] Sin datos aun, esperando {espera}s...")
                import time
                time.sleep(espera)
            else:
                print(f"   [Intento {intento}/{max_intentos}] Sin datos despues de {(max_intentos-1)*3}s de espera")

        print(f"   [D] {hoy} — {len(df_d):,} pedidos (consulta en vivo)")

        # Validacion: si no hay datos de hoy, error
        if len(df_d) == 0:
            print('\n[ERROR] No hay datos de ventas para hoy en la BD.')
            print('   Posibles causas:')
            print('   - Los datos aun no han sido procesados')
            print('   - Problema en la BD o en el SP')
            print('   - Problema de conectividad')
            print('\n   Contactar al Data Engineer si el problema persiste.')
            exit(1)

        df_d7  = cargar_dia_con_cache(conn, d7, 'd7')
        df_d14 = cargar_dia_con_cache(conn, d14, 'd14')
    finally:
        conn.close()

    mapa_zona = cargar_mapa_zonas()

    print('\n[2] Etiquetando horas y aplicando corte acumulado...')
    dfs = {}
    for nm, df in [('d', df_d), ('d7', df_d7), ('d14', df_d14)]:
        df = agregar_hora_lbl(df)
        df = filtrar_corte(df, hora_corte)
        df['ZONA2'] = df['ruta'].map(mapa_zona).fillna('(sin zona)')
        dfs[nm] = df
        print(f"   {nm}: {len(df):,} pedidos hasta {hora_lbl}")

    print('\n[3] Agregando indicadores...')
    g_general = agregar_general(dfs['d14'], dfs['d7'], dfs['d'])
    g_zonal   = agregar_por_columna(dfs['d14'], dfs['d7'], dfs['d'], 'ZONA2', ZONAS_ORDEN)
    # Incluir TODOS los supervisores del maestro + los que tienen ventas
    # (igual como se hace con vendedores en la hoja VENDEDOR)
    df_rutas_maestro = pd.read_excel(TABLAS_PATH, sheet_name='RUTA_ACTUAL', dtype={'RUTA': str})
    sup_maestro = set(df_rutas_maestro['SUPERVISOR'].dropna().unique())
    sup_con_ventas = (set(dfs['d14']['supervisor'].dropna()) |
                      set(dfs['d7']['supervisor'].dropna()) |
                      set(dfs['d']['supervisor'].dropna()))
    sup_orden = sorted(sup_maestro | sup_con_ventas)
    g_super = agregar_por_columna(dfs['d14'], dfs['d7'], dfs['d'], 'supervisor', sup_orden)

    # Datos para hoja VENDEDOR
    print('   Enriqueciendo con vendedor...')
    dfs_v = {}
    for nm, df in [('d', dfs['d']), ('d7', dfs['d7']), ('d14', dfs['d14'])]:
        dfs_v[nm] = enriquecer_con_vendedor(df)

    datos_sup_vend = agregar_por_sup_vendedor(dfs_v['d14'], dfs_v['d7'], dfs_v['d'])
    from generar_cuota_dia import calcular_cuota_dia_vendedor
    cuota_vend = calcular_cuota_dia_vendedor()
    # Mapa vendedor → lista de rutas (para cuota y etiqueta)
    df_rutas_raw = pd.read_excel(TABLAS_PATH, sheet_name='RUTA_ACTUAL', dtype={'RUTA': str})
    ruta_por_vendedor = (df_rutas_raw.groupby('VENDEDOR')['RUTA']
                         .apply(list).to_dict())

    # Cuota SUPERVISOR = suma de vendedores (para match con granularidad VENDEDOR)
    cuota_sup = calcular_cuota_supervisor(cuota_vend, ruta_por_vendedor)
    # Cuota ZONAL = suma por zona
    cuota_zonal = calcular_cuota_zonal(cuota_vend, ruta_por_vendedor)

    # Validación: Detectar si hay datos en cero (limpieza de BD por data center)
    soles_total_hoy = dfs['d']['monto'].sum() if len(dfs['d']) > 0 else 0.0
    if soles_total_hoy == 0:
        print(f'\n[WARN] DATOS EN CERO detectados')
        print(f'   Total de soles hoy: S/ {soles_total_hoy:.2f}')
        print(f'   El Data Center posiblemente está en limpieza de pedidos')
        print(f'   Se creará flag para reintento automático en 15 minutos')
        # Crear flag para que ejecutar_corte_orquestado.py lo detecte
        flag_file = SCRIPT_DIR / f"corte_validacion_{hora_corte}.flag"
        with open(flag_file, 'w') as f:
            f.write(f"SOLES_TOTAL:{soles_total_hoy}")
        print(f'   Flag creado para reintentar a las {datetime.datetime.now() + datetime.timedelta(seconds=900)}')

    print('\n[4] Generando Excel...')
    wb = Workbook()
    ws = wb.active
    ws.title = 'CORTE VENTAS'
    ws.sheet_view.showGridLines = False

    escribir_general(ws, g_general, fechas, nombre_dia, hora_lbl)
    escribir_categoria(ws, 11, 'RESUMEN POR ZONAL/ORIGEN', C_TIT_ZON, C_HDR_ZON,
                       C_DIF_ZON, C_BLANCO, C_DATA_ZS, 'ZONAL', ZONAS_ORDEN,
                       g_zonal, fechas, nombre_dia, hora_lbl, 7, cuota_sup=None, cuota_zonal=cuota_zonal)
    escribir_categoria(ws, 24, 'RESUMEN POR SUPERVISOR', C_TIT_SUP, C_HDR_SUP,
                       C_DIF_SUP, C_BLANCO, C_DATA_ZS, 'Supervisor', sup_orden,
                       g_super, fechas, nombre_dia, hora_lbl, len(sup_orden), cuota_sup=cuota_sup, cuota_zonal=None)

    # Hoja VENDEDOR
    ws_v = wb.create_sheet(title='VENDEDOR')
    escribir_hoja_vendedor(ws_v, datos_sup_vend, cuota_vend,
                           ruta_por_vendedor, fechas, nombre_dia, hora_lbl)

    out = f'{OUT_DIR}/CORTE_VENTAS_{hoy.strftime("%Y%m%d")}.xlsx'
    wb.save(out)
    print(f'\n   Guardado: {out}')
    print('\nProceso completado OK.')


if __name__ == '__main__':
    main()
