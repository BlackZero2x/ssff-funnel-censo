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


def cargar_cuota_supervisor() -> dict:
    """Carga cuota diaria por supervisor desde CuotaJunioV2.xlsx.

    Flujo:
    1. Lee CuotaJunioV2.xlsx (RUTA → CUOTA_DIA)
    2. Lee TABLAS_RUTAS.xlsx (RUTA → SUPERVISOR)
    3. Suma CUOTA_DIA por SUPERVISOR

    Retorna: {supervisor: cuota_dia_total}
    """
    df_cuota = pd.read_excel(CUOTA_PATH, dtype={'RUTA': str})
    df_rutas = pd.read_excel(TABLAS_PATH, sheet_name='RUTA_ACTUAL', dtype={'RUTA': str})

    # Merge: RUTA → CUOTA_DIA + SUPERVISOR
    df_merge = df_cuota[['RUTA', 'CUOTA_DIA']].merge(
        df_rutas[['RUTA', 'SUPERVISOR']],
        on='RUTA',
        how='left'
    )

    # Agrupar por SUPERVISOR y sumar
    cuota_sup = df_merge.groupby('SUPERVISOR')['CUOTA_DIA'].sum().to_dict()
    return cuota_sup


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
C_TIT_ZON  = '44546A'   # título ZONAL (dk2)
C_TIT_SUP  = '006C50'   # título SUPERVISOR (rgb directo)
C_CORTE    = 'FFFFCC'   # celdas Corte/hora
C_HDR_GEN  = 'DEEBF7'   # header fila4 GENERAL
C_DIF_GEN  = '9DC3E6'   # header dif GENERAL
C_HDR_ZON  = 'DAE3F3'   # header fila4 ZONAL
C_DIF_ZON  = '335693'   # header dif ZONAL (texto blanco)
C_HDR_SUP  = 'A0FFE6'   # header fila4 SUPERVISOR
C_DIF_SUP  = '006C50'   # header dif SUPERVISOR (texto blanco)
C_DATA_GEN = 'F2F2F2'   # relleno datos GENERAL
C_DATA_ZS  = 'D9D9D9'   # relleno datos ZONAL/SUPERVISOR
C_BLANCO   = 'FFFFFF'
C_NEGRO    = '000000'

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
                       hora_lbl, n_filas_fijo, cuota_sup=None):
    """Genérico para ZONAL (col0=11/K) y SUPERVISOR (col0=24/X).
    Para SUPERVISOR: agrega columnas Cuota_Día y %Avance, SOLO % en difs.
    datos: dict {fila_label: {'d14':(ped,sol),'d7':(..),'d':(..)}}."""
    c = col0
    L = get_column_letter

    # Para SUPERVISOR: estructura incluye Cuota_Día (AE) y %Avance (AF)
    es_supervisor = (etiqueta_col == 'Supervisor')

    if es_supervisor:
        # [X][Y Z][AA AB][AC AD][AE AF][AG AH]
        # Label, P14 S14, P7 S7, Pd Sd, Cuota %Avance, %Dif7, %Dif14
        # IMPORTANTE: SOLO columnas de % en diferencias (sin Diferencia numérica)
        c_lbl = c
        c_p14, c_s14 = c+1, c+2
        c_p7,  c_s7  = c+3, c+4
        c_pd,  c_sd  = c+5, c+6
        c_cuota = c+7
        c_pct_avance = c+8
        c_pct7 = c+9      # SOLO % para [D-7] vs [D]
        c_pct14 = c+10    # SOLO % para [D-14] vs [D]
    else:
        # ZONAL: [K][L M][N O][P Q][R S][T U]
        # Label, P14 S14, P7 S7, Pd Sd, Dif7 %7, Dif14 %14
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

    # Para SUPERVISOR: agregar "Seguimiento del día" (merge AE4:AF4)
    # SIN fill, font color verde (006C50, mismo que dif_color)
    if es_supervisor:
        ws.merge_cells(start_row=4, start_column=c_cuota, end_row=4, end_column=c_pct_avance)
        _set(ws, 4, c_cuota, 'Seguimiento del día',
             font=fnt(bold=True, italic=True, size=10, color='006C50'),
             fill=fill('FFFFFF'), align=aln())
        # % Diferencia: merge AG4:AH4, SIN fill, font verde
        ws.merge_cells(start_row=4, start_column=c_pct7, end_row=4, end_column=c_pct14)
        _set(ws, 4, c_pct7, '% Diferencia',
             font=fnt(bold=True, italic=True, size=10, color='006C50'),
             fill=fill('FFFFFF'), align=aln())
    else:
        # ZONAL: bandas de diferencia con color
        ws.merge_cells(start_row=4, start_column=c_dif7, end_row=4, end_column=c_pct7)
        _set(ws, 4, c_dif7, '[D-7] vs. [D]', font=fnt(italic=True, size=10, color=dif_font),
             fill=fill(dif_color), align=aln())
        ws.merge_cells(start_row=4, start_column=c_dif14, end_row=4, end_column=c_pct14)
        _set(ws, 4, c_dif14, '[D-14] vs. [D]', font=fnt(italic=True, size=10, color=dif_font),
             fill=fill(dif_color), align=aln())

    # Fila 5: encabezados
    if es_supervisor:
        # SUPERVISOR: SOLO % en diferencias, sin columnas "Diferencia" numérica
        headers = [(c_lbl, 'Supervisor'), (c_p14, 'Pedidos'), (c_s14, 'Soles'),
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
            # %Avance = IFERROR(AD/AE, "-")
            sd_c = L(c_sd)
            cuota_c = L(c_cuota)
            _set(ws, r, c_pct_avance, f'=IFERROR({sd_c}{r}/{cuota_c}{r},"-")',
                 font=fnt(), fill=fill(color_fila_cat), align=aln(), fmt=FMT_PCT)

            # Para SUPERVISOR: SOLO % en diferencias (sin columnas numérica)
            # %Dif7 = (Pd-Ps7)/Ps7, %Dif14 = (Pd-Ps14)/Ps14
            s7_c, s14_c = L(c_s7), L(c_s14)
            _set(ws, r, c_pct7, f'=({sd_c}{r}-{s7_c}{r})/{s7_c}{r}',
                 font=fnt(), fill=fill(color_fila_cat), align=aln(), fmt=FMT_PCT)
            _set(ws, r, c_pct14, f'=({sd_c}{r}-{s14_c}{r})/{s14_c}{r}',
                 font=fnt(), fill=fill(color_fila_cat), align=aln(), fmt=FMT_PCT)
        else:
            # Para ZONAL: Dif + % en ambas columnas
            sd_c, s7_c, s14_c = L(c_sd), L(c_s7), L(c_s14)
            dif7_c, dif14_c = L(c_dif7), L(c_dif14)
            _set(ws, r, c_dif7, f'=({sd_c}{r}-{s7_c}{r})', font=fnt(), fill=fill(color_fila_cat),
                 align=aln(), fmt=FMT_DIF_NUM)
            _set(ws, r, c_pct7, f'=IFERROR({dif7_c}{r}/{s7_c}{r},0)', font=fnt(), fill=fill(color_fila_cat),
                 align=aln('right'), fmt=FMT_PCT)
            _set(ws, r, c_dif14, f'=({sd_c}{r}-{s14_c}{r})', font=fnt(), fill=fill(color_fila_cat),
                 align=aln(), fmt=FMT_DIF_NUM)
            _set(ws, r, c_pct14, f'=IFERROR({dif14_c}{r}/{s14_c}{r},0)', font=fnt(), fill=fill(color_fila_cat),
                 align=aln('right'), fmt=FMT_PCT)

    # Iconos en %Dif (ambas tablas tienen iconos en las columnas de %)
    last_row = 6 + len(filas) - 1
    ws.conditional_formatting.add(f'{L(c_pct7)}6:{L(c_pct7)}{last_row}', _icon_rule())
    ws.conditional_formatting.add(f'{L(c_pct14)}6:{L(c_pct14)}{last_row}', _icon_rule())

    # APLICAR BORDES ESPECÍFICOS
    if es_supervisor:
        # SUPERVISOR
        # 1) X5:AH5 — outside border en fila encabezado
        for cc in range(c_lbl, c_pct14 + 1):
            _apply_border(ws.cell(5, cc), left=True, right=True, top=True, bottom=True)
        # 2) X6:X{last_row} — left border en columna etiqueta (filas de datos)
        for rr in range(6, last_row + 1):
            _apply_border(ws.cell(rr, c_lbl), left=True)
        # 3) Right borders en columnas de cierre de par, filas 4..last_row
        for rr in range(4, last_row + 1):
            _apply_border(ws.cell(rr, c_s14),       right=True)  # Z
            _apply_border(ws.cell(rr, c_s7),        right=True)  # AB
            _apply_border(ws.cell(rr, c_sd),        right=True)  # AD
            _apply_border(ws.cell(rr, c_pct_avance),right=True)  # AF
            _apply_border(ws.cell(rr, c_pct14),     right=True)  # AH
        # 4) X3:X{last_row} — right border en columna label filas 3..last_row
        for rr in range(3, last_row + 1):
            _apply_border(ws.cell(rr, c_lbl), right=True)
    else:
        # ZONAL
        # 1) K5:U5 — outside border en fila encabezado
        for cc in range(c_lbl, c_pct14 + 1):
            _apply_border(ws.cell(5, cc), left=True, right=True, top=True, bottom=True)
        # 2) K6:K{last_row} — left border en columna etiqueta (filas de datos)
        for rr in range(6, last_row + 1):
            _apply_border(ws.cell(rr, c_lbl), left=True)
        # 3) Right borders en columnas de cierre de par, filas 4..last_row
        right_border_cols = [c_lbl, c_s14, c_s7, c_sd, c_pct7, c_pct14]
        for cc in right_border_cols:
            for rr in range(4, last_row + 1):
                _apply_border(ws.cell(rr, cc), right=True)

    # Anchos
    if es_supervisor:
        # SUPERVISOR: [X][Y Z][AA AB][AC AD][AE AF][AG AH]
        ws.column_dimensions[L(c_lbl)].width = 21.71
        for cc in (c_p14, c_s14, c_p7, c_s7, c_pd, c_sd):
            ws.column_dimensions[L(cc)].width = 9.71
        ws.column_dimensions[L(c_cuota)].width = 11.0
        ws.column_dimensions[L(c_pct_avance)].width = 9.0
        ws.column_dimensions[L(c_pct7)].width = 14.0
        ws.column_dimensions[L(c_pct14)].width = 14.0
    else:
        # ZONAL
        ws.column_dimensions[L(c_lbl)].width = 16.0           # cambio: 13.0 → 16.0
        for cc in (c_p14, c_s14, c_p7, c_s7, c_pd, c_sd):
            ws.column_dimensions[L(cc)].width = 9.71
        ws.column_dimensions[L(c_dif7)].width = 14.0
        ws.column_dimensions[L(c_pct7)].width = 7.71
        ws.column_dimensions[L(c_dif14)].width = 14.0
        ws.column_dimensions[L(c_pct14)].width = 7.71


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
    # Orden de supervisores: unión de los 3 días, alfabético
    sup_orden = sorted(set(dfs['d14']['supervisor'].dropna()) |
                       set(dfs['d7']['supervisor'].dropna()) |
                       set(dfs['d']['supervisor'].dropna()))
    g_super = agregar_por_columna(dfs['d14'], dfs['d7'], dfs['d'], 'supervisor', sup_orden)

    # Cargar cuota diaria por supervisor
    cuota_sup = cargar_cuota_supervisor()

    print('\n[4] Generando Excel...')
    wb = Workbook()
    ws = wb.active
    ws.title = 'CORTE VENTAS'
    ws.sheet_view.showGridLines = False

    escribir_general(ws, g_general, fechas, nombre_dia, hora_lbl)
    escribir_categoria(ws, 11, 'RESUMEN POR ZONAL/ORIGEN', C_TIT_ZON, C_HDR_ZON,
                       C_DIF_ZON, C_BLANCO, C_DATA_ZS, 'ZONAL', ZONAS_ORDEN,
                       g_zonal, fechas, nombre_dia, hora_lbl, 7, cuota_sup=None)
    escribir_categoria(ws, 24, 'RESUMEN POR SUPERVISOR', C_TIT_SUP, C_HDR_SUP,
                       C_DIF_SUP, C_BLANCO, C_DATA_ZS, 'Supervisor', sup_orden,
                       g_super, fechas, nombre_dia, hora_lbl, len(sup_orden), cuota_sup=cuota_sup)

    out = f'{OUT_DIR}/CORTE_VENTAS_{nombre_dia}_{hoy.strftime("%d_%m_%Y")}_{hora_lbl}.xlsx'
    wb.save(out)
    print(f'\n   Guardado: {out}')

    # Envío automático a WhatsApp con comportamiento antibang (si no está --solo-excel)
    if not args.solo_excel:
        print('\n[5] Enviando reporte a WhatsApp...')
        try:
            from wa_sender_antibang import WABangSafeSender
            import json

            # Cargar configuración
            config_path = f'{OUT_DIR}/config.json'
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                sender = WABangSafeSender(wa_server_url="http://localhost:8002")

                # Verificar si puede enviar
                if not sender._puede_enviar():
                    print('   [WARN]  Cuenta bloqueada, saltando envío')
                else:
                    grupo_id = config.get('groups', {}).get('Canal_SSFF_2026_Gestion', '')
                    if not grupo_id:
                        print('   [WARN]  ID de grupo no configurado en config.json')
                    else:
                        # Mensaje automático
                        mensaje = f"📊 Corte de Ventas — {nombre_dia} {hoy.strftime('%d/%m/%Y')}\n⏰ Corte: {hora_lbl}\n[OK] Reporte generado"

                        if sender.send_to_group(grupo_id, mensaje, imagen_path=out):
                            print(f'   [OK] Enviado a Canal SSFF')
                        else:
                            print(f'   [ERROR] Error al enviar a WhatsApp')
            else:
                print(f'   [WARN]  config.json no encontrado (saltando envío automático)')

        except ImportError:
            print('   [WARN]  wa_sender_antibang no disponible (saltando envío)')
        except Exception as e:
            print(f'   [ERROR] Error en envío: {e}')
    else:
        print('\n[5] Omitiendo envío a WhatsApp (--solo-excel activo)')

    print('\nProceso completado OK.')


if __name__ == '__main__':
    main()
