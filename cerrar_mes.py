"""
cerrar_mes.py
-------------
Ejecutar al inicio de cada mes (ej: el 01/06, el 01/07, etc.).

Qué hace:
  1. Detecta automáticamente el mes a cerrar (el mes calendario anterior).
  2. Consulta SQL Server para traer ese mes completo y definitivo.
  3. Guarda el snapshot como export_data_ssff_YYMM_fecha.csv.
  4. Busca el CSV histórico más reciente en disco, fusiona, y genera
     el nuevo export_data_ssff_NNNN-YYMM_fecha.csv para el siguiente mes.

Uso:
  python cerrar_mes.py           → cierra el mes anterior automáticamente
  python cerrar_mes.py 2605      → cierra el mes indicado explícitamente
"""
import os
import sys
import glob
import datetime
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

BASE_DIR = 'C:/proyectos/SSFF'

# ── Determinar mes a cerrar ───────────────────────────────────────────
if len(sys.argv) > 1:
    _arg = sys.argv[1].strip()
    if len(_arg) == 4 and _arg.isdigit():
        MES_CIERRE_INT = int(_arg)
        MES_CIERRE_STR = _arg
    else:
        print(f"ERROR: argumento inválido '{_arg}'. Usa formato YYMM (ej: 2605).")
        raise SystemExit(1)
else:
    # Mes anterior al actual
    _hoy      = datetime.date.today()
    _primer   = _hoy.replace(day=1)
    _mes_ant  = _primer - datetime.timedelta(days=1)
    MES_CIERRE_STR = _mes_ant.strftime('%y%m')
    MES_CIERRE_INT = int(MES_CIERRE_STR)

# Nombre legible (ej: "Mayo 2026")
_anio = 2000 + int(MES_CIERRE_STR[:2])
_mes  = int(MES_CIERRE_STR[2:])
_NOMBRES = {1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',
            7:'Julio',8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'}
MES_NOMBRE = f"{_NOMBRES[_mes]} {_anio}"

print(f"\n{'═'*60}")
print(f"  CIERRE DE MES: {MES_NOMBRE}  ({MES_CIERRE_STR})")
print(f"{'═'*60}\n")

# ── Rutas de archivos ─────────────────────────────────────────────────
OUT_SNAPSHOT = f'{BASE_DIR}/export_data_ssff_{MES_CIERRE_STR}_fecha.csv'

# Histórico más reciente disponible (excluye snapshots de un solo mes)
_candidatos = sorted(
    glob.glob(f'{BASE_DIR}/export_data_ssff_????-????_fecha.csv'),
    reverse=True
)
if not _candidatos:
    print(f"ERROR: no se encontró ningún export_data_ssff_NNNN-NNNN_fecha.csv en {BASE_DIR}")
    raise SystemExit(1)

HIST_ACTUAL = _candidatos[0]
# Extraer el mes final del nombre del histórico actual (ej: 2511-2604 → 2604)
_sufijo     = os.path.basename(HIST_ACTUAL)          # export_data_ssff_2511-2604_fecha.csv
_rango      = _sufijo.replace('export_data_ssff_','').replace('_fecha.csv','')  # 2511-2604
_mes_ini    = _rango.split('-')[0]                   # 2511
OUT_NUEVO   = f'{BASE_DIR}/export_data_ssff_{_mes_ini}-{MES_CIERRE_STR}_fecha.csv'

print(f"  Histórico base:  {os.path.basename(HIST_ACTUAL)}")
print(f"  Snapshot salida: {os.path.basename(OUT_SNAPSHOT)}")
print(f"  Nuevo histórico: {os.path.basename(OUT_NUEVO)}\n")

# ── Leer .env ─────────────────────────────────────────────────────────
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

def _conn_str():
    return (
        f"DRIVER={{SQL Server}};"
        f"SERVER={_env['SQL_SERVER']};"
        f"DATABASE={_env['SQL_DATABASE']};"
        f"UID={_env['SQL_USER']};"
        f"PWD={_env['SQL_PASSWORD']};"
    )

QUERY_CIERRE = f"""
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
WHERE a.mes = '{MES_CIERRE_STR}'
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

# ── [1] Snapshot del mes desde SQL ───────────────────────────────────
print(f"[1] Consultando SQL Server — {MES_NOMBRE} completo...")
try:
    import pyodbc
    with pyodbc.connect(_conn_str(), timeout=60) as conn:
        df_mes = pd.read_sql(QUERY_CIERRE, conn)
    df_mes.columns        = df_mes.columns.str.strip()
    df_mes['ccod_cli']    = pd.to_numeric(df_mes['ccod_cli'],    errors='coerce').fillna(0).astype(int)
    df_mes['mes']         = pd.to_numeric(df_mes['mes'],         errors='coerce').fillna(0).astype(int)
    df_mes['ccod_vend']   = pd.to_numeric(df_mes['ccod_vend'],   errors='coerce').fillna(0).astype(int)
    df_mes['total_monto'] = pd.to_numeric(df_mes['total_monto'], errors='coerce').fillna(0)
    print(f"   Filas:           {len(df_mes):,}")
    print(f"   Clientes únicos: {df_mes['ccod_cli'].nunique():,}")
    print(f"   Total soles:     S/ {df_mes['total_monto'].sum():,.0f}")
except Exception as e:
    print(f"ERROR conectando a SQL Server: {e}")
    raise SystemExit(1)

# ── [2] Guardar snapshot ──────────────────────────────────────────────
print(f"\n[2] Guardando snapshot → {os.path.basename(OUT_SNAPSHOT)}")
df_mes.to_csv(OUT_SNAPSHOT, index=False, encoding='utf-8-sig')
print(f"   OK — {os.path.getsize(OUT_SNAPSHOT)/1024/1024:.1f} MB")

# ── [3] Fusionar con histórico ────────────────────────────────────────
print(f"\n[3] Fusionando con {os.path.basename(HIST_ACTUAL)}...")
df_hist = pd.read_csv(
    HIST_ACTUAL, low_memory=False,
    dtype={'ccod_cli': 'int32', 'mes': 'int32', 'ccod_vend': 'int32',
           'ccod_ruta': str, 'categoria': str, 'fecha': str}
)
# Eliminar filas del mes cerrado si ya existían en el histórico (idempotente)
df_hist = df_hist[df_hist['mes'] != MES_CIERRE_INT]
df_nuevo = pd.concat([df_hist, df_mes], ignore_index=True)
df_nuevo = df_nuevo.sort_values(['mes', 'ccod_cli', 'fecha']).reset_index(drop=True)

print(f"   Histórico base:  {len(df_hist):,} filas")
print(f"   + {MES_NOMBRE}:  {len(df_mes):,} filas")
print(f"   = Total nuevo:   {len(df_nuevo):,} filas")
print(f"   Rango meses:     {df_nuevo['mes'].min()} → {df_nuevo['mes'].max()}")

# ── [4] Guardar nuevo histórico ───────────────────────────────────────
print(f"\n[4] Guardando nuevo histórico → {os.path.basename(OUT_NUEVO)}")
df_nuevo.to_csv(OUT_NUEVO, index=False, encoding='utf-8-sig')
print(f"   OK — {os.path.getsize(OUT_NUEVO)/1024/1024:.1f} MB")

# ── Resumen ───────────────────────────────────────────────────────────
_mes_sig_int = MES_CIERRE_INT + 1
if _mes_sig_int % 100 == 13:   # diciembre → enero del año siguiente
    _mes_sig_int = (_mes_sig_int // 100 + 1) * 100 + 1
print(f"""
╔══════════════════════════════════════════════════════════════════╗
║  CIERRE COMPLETADO: {MES_NOMBRE:<44}║
╠══════════════════════════════════════════════════════════════════╣
║  Snapshot:        {os.path.basename(OUT_SNAPSHOT):<45}║
║  Nuevo histórico: {os.path.basename(OUT_NUEVO):<45}║
╠══════════════════════════════════════════════════════════════════╣
║  generar_funnel_censo.py se actualiza automáticamente.           ║
║  Solo verificar:                                                 ║
║  · CUOTAS_MAY_PATH → archivo de cuotas del nuevo mes            ║
║  · FERIADOS del nuevo mes ({_mes_sig_int}) si aplica                    ║
╚══════════════════════════════════════════════════════════════════╝
""")
