"""
generar_reporte_pedidos.py

Genera un Excel con corte de pedidos por supervisor/vendedor vs. cuota por categoría.
Una hoja única "PEDIDOS" con tablas apiladas (una por supervisor, 3 filas de separación).

Estructura de cada tabla (basada en CORTE_PEDIDOS_EJEMPLO.xlsx):
  Fila +0: [SUPERVISOR teal] ................. [Corte][9AM amarillo]
  Fila +1: vacía
  Fila +2: [ ] [PROCESADOS merge B:C] [ANDINA merge D:E] ... (categorías italic, fondo teal)
  Fila +3: [Vendedor/RUTA] [Cuota][Pedidos] x10  (bold, borde thin all)
  Fila +4..: vendedores (damero blanco/#D9D9D9, borde A y borde derecho de cada cat)
  Última: TOTAL (fondo teal, bold, borde thin all)

Color teal = #255663 (theme 8 + tint -0.5 → accent5 #4BACC6 oscurecido)
Cuota: CUOTAS_CATEGORIAS_JUN2026.xlsx cruzada por ZONA3.
Pedidos: cobertura del SP sp_comxp_PedidosDiaResumen (avance='categorias').

Semáforo en celda Pedidos:
  <=70% cuota → #FFC7CE (rojo)
  >70% y <90% → #FFEB9C (amarillo)
  >=90%       → #C6EFCE (verde)
  cuota = 0   → sin color (fondo del damero)

Uso:
  python generar_reporte_pedidos.py             # hora actual
  python generar_reporte_pedidos.py --hora 9    # forzar etiqueta "9AM"
"""

import argparse
import datetime
import warnings
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════════════════

BASE_DIR   = Path('C:/proyectos/SSFF')
SCRIPT_DIR = Path(__file__).parent.absolute()
FILES_DIR  = BASE_DIR / 'files'

TABLAS_RUTAS_PATH = BASE_DIR / 'TABLAS_RUTAS.xlsx'
CUOTAS_PATH       = FILES_DIR / 'CUOTAS_SSFF.xlsx'

CATEGORIAS = [
    'PROCESADOS', 'ANDINA', 'KIMBERLY', 'COLGATE',
    'RINTI', 'VERDUM', 'HUEVO', 'HOMEPRO PERU', 'LA PATRONA', 'TAMBOS PERU',
]


def _etiqueta_hora(hora: int) -> str:
    if hora < 12:
        return f'{hora}AM'
    elif hora == 12:
        return '12PM'
    else:
        return f'{hora - 12}PM'


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS DE ESTILO
# ══════════════════════════════════════════════════════════════════════════════

def _fill(hex6: str) -> PatternFill:
    """PatternFill sólido a partir de 6 dígitos hex (sin #)."""
    return PatternFill('solid', fgColor=f'FF{hex6}')

FUENTE = 'Aptos Narrow'

def _fnt(bold=False, italic=False, size=11, color='000000') -> Font:
    return Font(name=FUENTE, bold=bold, italic=italic, size=size, color=f'FF{color}')

def _aln(h='center', v='center') -> Alignment:
    return Alignment(horizontal=h, vertical=v)

def _side(style=None) -> Side:
    return Side(style=style, color='FF000000') if style else Side(style=None)

def _brd(left=None, right=None, top=None, bottom=None) -> Border:
    return Border(
        left=_side(left), right=_side(right),
        top=_side(top),   bottom=_side(bottom),
    )

_BRD_ALL   = _brd('thin', 'thin', 'thin', 'thin')
_BRD_LR    = _brd('thin', 'thin', None, None)   # solo izq+der (filas de datos col A)
_BRD_RIGHT = _brd(None, 'thin', None, None)      # solo der (límite de cada categoría)
_BRD_NONE  = _brd()

# Colores
_TEAL     = '255663'   # encabezado supervisor y TOTAL y categorías
_AMARILLO = 'FFFFCC'   # celda "Corte" / hora
_BLANCO   = 'FFFFFF'
_GRIS     = 'D9D9D9'
_ROJO     = 'FFC7CE'
_NARANJA  = 'FFEB9C'
_VERDE    = 'C6EFCE'

# Fills pre-calculados
FILL_TEAL    = _fill(_TEAL)
FILL_HORA    = _fill(_AMARILLO)
FILL_BLANCO  = _fill(_BLANCO)
FILL_GRIS    = _fill(_GRIS)
FILL_ROJO    = _fill(_ROJO)
FILL_NARANJA = _fill(_NARANJA)
FILL_VERDE   = _fill(_VERDE)
FILL_VACIO   = PatternFill(fill_type=None)   # sin relleno (fila vacía)

# Fonts
FONT_SUP   = _fnt(bold=True,  italic=True,  size=14, color='FFFFFF')
FONT_HORA  = _fnt(bold=True,  italic=True,  size=14, color='000000')
FONT_CAT   = _fnt(bold=False, italic=True,  size=11, color='FFFFFF')
FONT_HDR   = _fnt(bold=True,  italic=False, size=11, color='000000')
FONT_VEND  = _fnt(bold=False, italic=False, size=11, color='000000')
FONT_TOT   = _fnt(bold=True,  italic=False, size=11, color='FFFFFF')


# ══════════════════════════════════════════════════════════════════════════════
# CONEXIÓN SQL
# ══════════════════════════════════════════════════════════════════════════════

def _leer_env(path: Path) -> dict:
    env = {}
    if path.exists():
        with open(path, encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    k, v = line.split('=', 1)
                    env[k.strip()] = v.strip()
    return env


def conectar_sql():
    import pyodbc
    env = _leer_env(BASE_DIR / '.env')
    return pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={env['SQL_SERVER']};DATABASE={env['SQL_DATABASE']};"
        f"UID={env['SQL_USER']};PWD={env['SQL_PASSWORD']};"
        f"TrustServerCertificate=yes;"
    )


# ══════════════════════════════════════════════════════════════════════════════
# CARGA DE DATOS
# ══════════════════════════════════════════════════════════════════════════════

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


def cargar_rutas() -> pd.DataFrame:
    df = pd.read_excel(TABLAS_RUTAS_PATH, sheet_name='RUTA_ACTUAL', dtype={'RUTA': str})
    df = df[df['ZONA3'] != 'OMITIR'].copy()
    df['RUTA'] = df['RUTA'].str.strip()
    df['PREFIX_RUTA'] = df['PREFIX_RUTA'].astype(str).str.strip()

    df_ffvv = _cargar_distribucion_ffvv_sql()
    df = df.drop(columns=['VENDEDOR', 'SUPERVISOR']).merge(df_ffvv, on='RUTA', how='left')
    faltantes = df[df['VENDEDOR'].isna()]['RUTA'].tolist()
    if faltantes:
        print(f"   [WARN] Rutas sin distribución en viewSFffvv: {faltantes}")

    return df


def cargar_cuotas() -> dict:
    """Lee CUOTAS_SSFF.xlsx hoja PEDIDOS y devuelve {(PREFIX_RUTA, Categoria): cuota_dia}.
    El valor ya es la cuota diaria directa — no requiere división."""
    df = pd.read_excel(CUOTAS_PATH, sheet_name='PEDIDOS')
    # Detectar columna del mes vigente (primera columna de tipo datetime)
    import datetime as _dt
    hoy = _dt.date.today()
    col_mes = None
    for col in df.columns:
        if isinstance(col, _dt.datetime) and col.year == hoy.year and col.month == hoy.month:
            col_mes = col
            break
    if col_mes is None:
        # Fallback: usar la última columna numérica disponible
        col_mes = [c for c in df.columns if c not in ('PREFIX_RUTA', df.columns[1])][-1]
        print(f'   [WARN] No se encontró columna del mes vigente en PEDIDOS; usando: {col_mes}')
    col_cat = df.columns[1]
    df = df[['PREFIX_RUTA', col_cat, col_mes]].rename(
        columns={col_cat: 'CAT', col_mes: 'CUOTA'}
    )
    df['PREFIX_RUTA'] = df['PREFIX_RUTA'].astype(str).str.strip()
    df['CAT']         = df['CAT'].astype(str).str.strip().str.upper()
    df['CUOTA']       = pd.to_numeric(df['CUOTA'], errors='coerce').fillna(0).astype(int)
    return {(r['PREFIX_RUTA'], r['CAT']): r['CUOTA'] for _, r in df.iterrows()}


def cargar_pedidos(fecha: datetime.date) -> pd.DataFrame:
    """Devuelve DataFrame {ruta, categoria, cobertura} del SP (avance='categorias')."""
    conn = conectar_sql()
    cur  = conn.cursor()
    cur.execute('EXEC [eAuren].[dbo].[sp_comxp_PedidosDiaResumen] ?, ?', '0003', fecha)

    frames = []
    while True:
        if cur.description:
            cols = [d[0] for d in cur.description]
            if 'avance' in cols:
                rows = cur.fetchall()
                frames.append(pd.DataFrame([list(r) for r in rows], columns=cols))
            else:
                cur.fetchall()
        if not cur.nextset():
            break
    conn.close()

    if not frames:
        return pd.DataFrame(columns=['ruta', 'categoria', 'cobertura'])

    df = pd.concat(frames, ignore_index=True)
    df = df[df['avance'] == 'categorias'].copy()
    df['ruta']      = df['ruta'].astype(str).str.strip()
    df['categoria'] = df['categoria'].astype(str).str.strip()
    df['cobertura'] = pd.to_numeric(df['cobertura'], errors='coerce').fillna(0).astype(int)
    return df[['ruta', 'categoria', 'cobertura']]


# ══════════════════════════════════════════════════════════════════════════════
# CONSTRUCCIÓN DEL EXCEL
# ══════════════════════════════════════════════════════════════════════════════

def _semaforo(pedidos: int, cuota: int) -> PatternFill:
    """Devuelve el fill del semáforo o None si cuota=0."""
    if cuota <= 0:
        return None
    r = pedidos / cuota
    if r <= 0.70:
        return FILL_ROJO
    elif r < 0.90:
        return FILL_NARANJA
    else:
        return FILL_VERDE


def _set(cell, value=None, fill=None, font=None, aln=None, brd=None):
    """Aplica propiedades a una celda de forma compacta."""
    if value is not None:
        cell.value = value
    if fill is not None:
        cell.fill = fill
    if font is not None:
        cell.font = font
    if aln is not None:
        cell.alignment = aln
    if brd is not None:
        cell.border = brd


def _escribir_tabla(ws, f0: int, supervisor: str, vendedores: list,
                    cuota_zona: dict, pedidos_pivot: pd.DataFrame,
                    etiqueta: str) -> int:
    """
    Escribe la tabla de un supervisor a partir de la fila f0.
    Retorna la fila siguiente (f0 + altura_tabla + 1).

    Layout (relativo a f0):
      f0+0  → título supervisor (teal, bold+italic)
      f0+1  → fila vacía
      f0+2  → nombres de categorías (merge 2 celdas, teal, italic, blanco)
      f0+3  → sub-encabezados Cuota/Pedidos (borde thin all)
      f0+4  → primer vendedor
      ...
      f0+4+n → fila TOTAL (teal, bold, borde thin all)
    """
    n_cat  = len(CATEGORIAS)
    n_cols = 1 + 2 * n_cat          # col A + 2 por cada categoría
    col_U  = get_column_letter(n_cols)  # última columna (U si 10 cats)

    # ── F0: título supervisor ──────────────────────────────────────────────────
    f = f0
    for col in range(1, n_cols - 1):   # A hasta penúltima: fondo teal
        c = ws.cell(row=f, column=col)
        c.fill = FILL_TEAL
    # Nombre supervisor en A
    _set(ws.cell(row=f, column=1, value=supervisor),
         font=FONT_SUP, aln=_aln('left'), fill=FILL_TEAL)
    # "Corte" en penúltima columna, hora en última
    _set(ws.cell(row=f, column=n_cols - 1, value='Corte'),
         fill=FILL_HORA, font=FONT_HORA, aln=_aln('center'))
    _set(ws.cell(row=f, column=n_cols, value=etiqueta),
         fill=FILL_HORA, font=FONT_HORA, aln=_aln('center'))

    # ── F0+1: fila vacía ───────────────────────────────────────────────────────
    # No escribimos nada; las celdas quedan en blanco por defecto.

    # ── F0+2: nombres de categorías (con merge) ────────────────────────────────
    f = f0 + 2
    # Columna A: sin relleno, borde izq+der para cerrar visualmente la tabla
    _set(ws.cell(row=f, column=1),
         fill=FILL_BLANCO, brd=_brd('thin', 'thin', None, None))
    for i, cat in enumerate(CATEGORIAS):
        col_q = 2 + i * 2
        col_p = 3 + i * 2
        # Merge sobre Cuota+Pedidos
        ws.merge_cells(start_row=f, start_column=col_q,
                       end_row=f,   end_column=col_p)
        c = ws.cell(row=f, column=col_q, value=cat)
        _set(c, fill=FILL_TEAL, font=FONT_CAT, aln=_aln('center'),
             brd=_brd(None, 'thin', None, None))

    # ── F0+3: sub-encabezados Cuota / Pedidos ─────────────────────────────────
    f = f0 + 3
    _set(ws.cell(row=f, column=1, value='Vendedor / RUTA'),
         fill=FILL_BLANCO, font=FONT_HDR, aln=_aln('center'), brd=_BRD_ALL)
    for i in range(n_cat):
        col_q = 2 + i * 2
        col_p = 3 + i * 2
        _set(ws.cell(row=f, column=col_q, value='Cuota'),
             fill=FILL_BLANCO, font=FONT_HDR, aln=_aln('center'), brd=_BRD_ALL)
        _set(ws.cell(row=f, column=col_p, value='Pedidos'),
             fill=FILL_BLANCO, font=FONT_HDR, aln=_aln('center'), brd=_BRD_ALL)

    # ── Filas de vendedores ────────────────────────────────────────────────────
    f_vend = f0 + 4
    tot_q  = {cat: 0 for cat in CATEGORIAS}
    tot_p  = {cat: 0 for cat in CATEGORIAS}

    for idx, vend in enumerate(vendedores):
        nombre      = f"{vend['VENDEDOR']} - {vend['RUTA']}"
        prefix_ruta = vend['PREFIX_RUTA']
        ruta        = vend['RUTA']
        fondo       = FILL_BLANCO if idx % 2 == 0 else FILL_GRIS

        # Columna A: borde izq+der
        _set(ws.cell(row=f_vend, column=1, value=nombre),
             fill=fondo, font=FONT_VEND, aln=_aln('left'),
             brd=_brd('thin', 'thin', None, None))

        for i, cat in enumerate(CATEGORIAS):
            col_q = 2 + i * 2
            col_p = 3 + i * 2

            cuota   = cuota_zona.get((prefix_ruta, cat), 0)
            pedidos = int(pedidos_pivot.at[ruta, cat]) \
                      if ruta in pedidos_pivot.index and cat in pedidos_pivot.columns else 0

            tot_q[cat] += cuota
            tot_p[cat] += pedidos

            # Cuota: sin borde (interior de la cat)
            _set(ws.cell(row=f_vend, column=col_q, value=cuota),
                 fill=fondo, font=FONT_VEND, aln=_aln('center'), brd=_BRD_NONE)

            # Pedidos: borde derecho (límite de la categoría)
            sem = _semaforo(pedidos, cuota)
            _set(ws.cell(row=f_vend, column=col_p, value=pedidos),
                 fill=sem if sem else fondo, font=FONT_VEND, aln=_aln('center'),
                 brd=_brd(None, 'thin', None, None))

        f_vend += 1

    # ── Fila TOTAL ─────────────────────────────────────────────────────────────
    _set(ws.cell(row=f_vend, column=1, value='TOTAL'),
         fill=FILL_TEAL, font=FONT_TOT, aln=_aln('left'), brd=_BRD_ALL)
    for i, cat in enumerate(CATEGORIAS):
        col_q = 2 + i * 2
        col_p = 3 + i * 2
        _set(ws.cell(row=f_vend, column=col_q, value=tot_q[cat]),
             fill=FILL_TEAL, font=FONT_TOT, aln=_aln('center'), brd=_BRD_ALL)
        _set(ws.cell(row=f_vend, column=col_p, value=tot_p[cat]),
             fill=FILL_TEAL, font=FONT_TOT, aln=_aln('center'), brd=_BRD_ALL)

    return f_vend + 1   # siguiente fila disponible


def generar_excel(fecha: datetime.date, hora: int) -> str:
    etiqueta = _etiqueta_hora(hora)

    print('[1/3] Cargando rutas...')
    df_rutas = cargar_rutas()

    print('[2/3] Cargando cuotas...')
    cuota_zona = cargar_cuotas()

    print(f'[3/3] Consultando pedidos del SP ({fecha})...')
    df_pedidos = cargar_pedidos(fecha)

    pedidos_pivot = df_pedidos.pivot_table(
        index='ruta', columns='categoria', values='cobertura',
        aggfunc='sum', fill_value=0
    )

    supervisores = df_rutas['SUPERVISOR'].drop_duplicates().tolist()

    wb = Workbook()
    ws = wb.active
    ws.title = 'PEDIDOS'
    ws.sheet_view.showGridLines = False

    # Anchos de columna
    ws.column_dimensions['A'].width = 28
    for i in range(len(CATEGORIAS)):
        ws.column_dimensions[get_column_letter(2 + i * 2)].width = 8
        ws.column_dimensions[get_column_letter(3 + i * 2)].width = 8

    fila = 1
    for sup in supervisores:
        vendedores = df_rutas[df_rutas['SUPERVISOR'] == sup][
            ['VENDEDOR', 'RUTA', 'ZONA3', 'PREFIX_RUTA']
        ].to_dict('records')
        if not vendedores:
            continue
        fila_sig = _escribir_tabla(ws, fila, sup, vendedores,
                                   cuota_zona, pedidos_pivot, etiqueta)
        fila = fila_sig + 3   # 3 filas de separación entre tablas

    nombre  = f"CORTE_PEDIDOS_{fecha.strftime('%Y%m%d')}.xlsx"
    ruta_out = SCRIPT_DIR / nombre
    wb.save(str(ruta_out))
    print(f'\n[OK] Archivo generado: {ruta_out}')
    return str(ruta_out)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description='Genera reporte de pedidos por supervisor')
    parser.add_argument('--hora',  type=int, default=None,
                        help='Hora (0-23). Omitir = hora actual.')
    parser.add_argument('--fecha', type=str, default=None,
                        help='Fecha YYYY-MM-DD. Omitir = hoy.')
    args = parser.parse_args()

    ahora = datetime.datetime.now()
    fecha = datetime.date.fromisoformat(args.fecha) if args.fecha else ahora.date()
    hora  = args.hora if args.hora is not None else ahora.hour

    print(f"\n{'='*70}")
    print(f'[PEDIDOS] Corte {_etiqueta_hora(hora)} — {fecha.strftime("%d/%m/%Y")}')
    print(f"{'='*70}\n")

    generar_excel(fecha, hora)


if __name__ == '__main__':
    main()
