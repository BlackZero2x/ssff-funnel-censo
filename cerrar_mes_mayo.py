"""
cerrar_mes_mayo.py
------------------
Ejecutar antes de generar la cuota de junio (puede ser con mayo aún abierto).

Qué hace:
  1. Consulta SQL Server mayo 2026 (parcial o completo).
  2. Proyecta al cierre de mes si los datos son parciales (según días laborales transcurridos).
  3. Actualiza export_data_ssff_rutas.csv  (detalle por cliente/ruta — usa generar_cuota.py)
  4. Actualiza export_data_ssff.csv        (global por periodo/categoria/linea — usa generar_cuota.py)
  5. Guarda snapshot de mayo: export_data_ssff_2605_fecha.csv

Proyección: factor = DIAS_LAB_MES / días_con_venta_en_SQL
  Si los datos ya son del mes completo (día hábil MAX = último del mes), factor = 1.
"""
import os
import datetime
import numpy as np
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

# ── Config ────────────────────────────────────────────────────────────
BASE_DIR    = 'C:/proyectos/SSFF'
HIST_RUTAS  = f'{BASE_DIR}/export_data_ssff_rutas.csv'   # detalle por ruta (usa generar_cuota.py)
HIST_GLOBAL = f'{BASE_DIR}/export_data_ssff.csv'         # global (usa generar_cuota.py)
OUT_SNAP    = f'{BASE_DIR}/export_data_ssff_2605_fecha.csv'  # snapshot mayo

DIAS_LAB_MAYO = 25   # días laborales totales de mayo 2026

# Query nivel cliente × ruta × categoria × linea × fecha
QUERY_MAYO = """
SELECT
     a.ccod_cli
    ,a.[mes]
    ,a.[ccod_ruta]
    ,a.ccod_vend
    ,b.[categoria]
    ,b.[linea]
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
  AND a.ccod_ruta != '0000'
GROUP BY
     a.ccod_cli
    ,a.[mes]
    ,a.[ccod_ruta]
    ,a.ccod_vend
    ,b.[categoria]
    ,b.[linea]
    ,a.fecha
"""

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

# ── [1] Consultar mayo desde SQL ──────────────────────────────────────
print("[1] Consultando SQL Server — mayo 2026...")
try:
    import pyodbc
    with pyodbc.connect(_conn_str(), timeout=60) as conn:
        df_mayo = pd.read_sql(QUERY_MAYO, conn)
    df_mayo.columns = df_mayo.columns.str.strip()
    for col in ['ccod_cli', 'mes', 'ccod_vend']:
        df_mayo[col] = pd.to_numeric(df_mayo[col], errors='coerce').fillna(0).astype(int)
    for col in ['pedidos', 'total_volumen', 'total_monto']:
        df_mayo[col] = pd.to_numeric(df_mayo[col], errors='coerce').fillna(0)
    df_mayo['ccod_ruta'] = df_mayo['ccod_ruta'].astype(str).str.strip()
    df_mayo['categoria'] = df_mayo['categoria'].astype(str).str.strip()
    df_mayo['linea']     = df_mayo['linea'].astype(str).str.strip()
    print(f"   Filas: {len(df_mayo):,}  |  Clientes únicos: {df_mayo['ccod_cli'].nunique():,}  |  "
          f"Total soles: S/ {df_mayo['total_monto'].sum():,.0f}")
except Exception as e:
    print(f"ERROR conectando a SQL Server: {e}")
    raise SystemExit(1)

# ── [2] Calcular factor de proyección ─────────────────────────────────
print("[2] Calculando factor de proyección al cierre de mayo...")
fechas_mayo = pd.to_datetime(df_mayo['fecha'], errors='coerce').dropna()
if len(fechas_mayo) > 0:
    fecha_max = fechas_mayo.max()
    # Contar días distintos con ventas como proxy de días laborales transcurridos
    dias_con_venta = df_mayo.assign(_f=pd.to_datetime(df_mayo['fecha'], errors='coerce'))['_f'].dt.date.nunique()
else:
    fecha_max = None
    dias_con_venta = DIAS_LAB_MAYO

factor_proy = DIAS_LAB_MAYO / dias_con_venta if dias_con_venta < DIAS_LAB_MAYO else 1.0
print(f"   Fecha máx en SQL: {fecha_max}  |  Días con venta: {dias_con_venta}  |  "
      f"Días lab. mayo: {DIAS_LAB_MAYO}  |  Factor proyección: {factor_proy:.4f}")

if factor_proy > 1.0:
    print(f"   → Proyectando mayo parcial × {factor_proy:.4f} para estimar cierre de mes")
    df_mayo['total_monto']   = (df_mayo['total_monto']   * factor_proy).round(4)
    df_mayo['total_volumen'] = (df_mayo['total_volumen'] * factor_proy).round(4)
    df_mayo['pedidos']       = (df_mayo['pedidos']       * factor_proy).round(4)
else:
    print("   → Mes completo, sin proyección")

print(f"   Total soles proyectado: S/ {df_mayo['total_monto'].sum():,.0f}")

# ── [3] Guardar snapshot mayo ─────────────────────────────────────────
print(f"[3] Guardando snapshot → {OUT_SNAP}")
df_mayo.to_csv(OUT_SNAP, index=False, encoding='utf-8-sig')
print(f"   OK — {os.path.getsize(OUT_SNAP)/1024/1024:.1f} MB")

# ── [4] Actualizar CSV de rutas (export_data_ssff_rutas.csv) ──────────
print(f"[4] Actualizando CSV de rutas ({HIST_RUTAS})...")
df_rutas_ant = pd.read_csv(
    HIST_RUTAS, low_memory=False,
    dtype={'ccod_cli': 'int32', 'mes': 'int32', 'ccod_ruta': str,
           'categoria': str, 'linea': str}
)
df_rutas_ant = df_rutas_ant[df_rutas_ant['mes'] != 2605]  # quitar mayo si ya existía
# Alinear columnas al formato del CSV existente (sin fecha, sin sublinea/producto)
cols_rutas = list(df_rutas_ant.columns)
df_mayo_rutas = df_mayo[['ccod_cli','mes','ccod_ruta','ccod_vend',
                          'categoria','linea','pedidos','total_volumen','total_monto']].copy()
# Solo mantener columnas que existen en el CSV original
cols_comunes = [c for c in cols_rutas if c in df_mayo_rutas.columns]
df_mayo_rutas = df_mayo_rutas[cols_comunes]
df_rutas_nuevo = pd.concat([df_rutas_ant[cols_comunes], df_mayo_rutas], ignore_index=True)
df_rutas_nuevo = df_rutas_nuevo.sort_values(['mes','ccod_ruta','ccod_cli']).reset_index(drop=True)
df_rutas_nuevo.to_csv(HIST_RUTAS, index=False, encoding='utf-8-sig')
print(f"   Anterior: {len(df_rutas_ant):,} filas  +  Mayo: {len(df_mayo_rutas):,} = {len(df_rutas_nuevo):,} total")
print(f"   Rango meses: {df_rutas_nuevo['mes'].min()} → {df_rutas_nuevo['mes'].max()}")
print(f"   OK — {os.path.getsize(HIST_RUTAS)/1024/1024:.1f} MB")

# ── [5] Actualizar CSV global (export_data_ssff.csv) ──────────────────
print(f"[5] Actualizando CSV global ({HIST_GLOBAL})...")
df_global_ant = pd.read_csv(
    HIST_GLOBAL, header=None,
    names=['PERIODO','CATEGORIA','LINEA','PEDIDOS','VOLUMEN','MONTO']
)
df_global_ant = df_global_ant[df_global_ant['PERIODO'] != 2605]

# Agregar mayo al nivel PERIODO × CATEGORIA × LINEA
df_mayo_global = (df_mayo
    .groupby(['mes','categoria','linea'])[['pedidos','total_volumen','total_monto']]
    .sum().reset_index()
    .rename(columns={'mes':'PERIODO','categoria':'CATEGORIA','linea':'LINEA',
                     'pedidos':'PEDIDOS','total_volumen':'VOLUMEN','total_monto':'MONTO'})
)
df_global_nuevo = pd.concat([df_global_ant, df_mayo_global], ignore_index=True)
df_global_nuevo = df_global_nuevo.sort_values(['PERIODO','CATEGORIA','LINEA']).reset_index(drop=True)
df_global_nuevo.to_csv(HIST_GLOBAL, index=False, header=False, encoding='utf-8-sig')
print(f"   Anterior: {len(df_global_ant):,} filas  +  Mayo: {len(df_mayo_global):,} = {len(df_global_nuevo):,} total")
print(f"   OK — {os.path.getsize(HIST_GLOBAL)/1024/1024:.1f} MB")

print(f"""
╔══════════════════════════════════════════════════════════════════╗
║  MAYO INCORPORADO  —  {datetime.date.today()}  (factor: {factor_proy:.4f})        ║
╠══════════════════════════════════════════════════════════════════╣
║  export_data_ssff_rutas.csv  → actualizado con mayo 2605         ║
║  export_data_ssff.csv        → actualizado con mayo 2605         ║
║  Snapshot: export_data_ssff_2605_fecha.csv                       ║
╠══════════════════════════════════════════════════════════════════╣
║  SIGUIENTE PASO:                                                 ║
║    uv run python generar_cuota.py --periodo 202606               ║
╚══════════════════════════════════════════════════════════════════╝
""")
