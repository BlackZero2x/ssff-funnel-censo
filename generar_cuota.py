import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import warnings
import argparse
from datetime import date
warnings.filterwarnings('ignore')

# ════════════════════════════════════════════════════════════════════════
# ARGUMENTO DE PERIODO  -- python generar_cuota.py --periodo 202606
# ════════════════════════════════════════════════════════════════════════
_parser = argparse.ArgumentParser(description='Genera cuota mensual SSFF')
_parser.add_argument('--periodo', type=str, default=None,
                     help='Periodo de cuota en formato YYYYMM (ej: 202606). '
                          'Si se omite se usa el mes siguiente al actual.')
_args = _parser.parse_args()

if _args.periodo:
    _periodo_cuota = int(_args.periodo)
else:
    _hoy = date.today()
    _sig = _hoy.month % 12 + 1
    _anio = _hoy.year + (1 if _hoy.month == 12 else 0)
    _periodo_cuota = int(f'{_anio}{_sig:02d}')

# Derivar periodo anterior (mes de referencia para proyeccion)
_anio_c  = _periodo_cuota // 100
_mes_c   = _periodo_cuota % 100
_mes_ref = _mes_c - 1 if _mes_c > 1 else 12
_anio_r  = _anio_c if _mes_c > 1 else _anio_c - 1
_periodo_ref = int(f'{_anio_r}{_mes_ref:02d}')

# Nombres de mes para archivos y textos
_MESES_ES = {1:'ene',2:'feb',3:'mar',4:'abr',5:'may',6:'jun',
             7:'jul',8:'ago',9:'sep',10:'oct',11:'nov',12:'dic'}
_MESES_ES_LARGO = {1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',
                   7:'Julio',8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'}

_nombre_cuota = f'{_MESES_ES[_mes_c]}{_anio_c}'   # ej: jun2026
_nombre_ref   = f'{_MESES_ES[_mes_ref]}{_anio_r}'  # ej: may2026
_label_cuota  = f'{_MESES_ES_LARGO[_mes_c]} {_anio_c}'   # ej: Junio 2026
_label_ref    = f'{_MESES_ES_LARGO[_mes_ref]} {_anio_r}'  # ej: Mayo 2026

# Periodos Tipo 1: los 2 meses anteriores al mes de referencia (febrero excluido si es atipico)
_mes_t1b = _mes_ref - 1 if _mes_ref > 1 else 12
_anio_t1b = _anio_r if _mes_ref > 1 else _anio_r - 1
_periodo_t1b = int(f'{_anio_t1b}{_mes_t1b:02d}')

# Periodos en formato YYMM (4 dígitos) para comparar con el CSV (export_data_ssff.csv usa YYMM)
_periodo_ref_yymm  = int(f'{_anio_r % 100:02d}{_mes_ref:02d}')
_periodo_t1b_yymm  = int(f'{_anio_t1b % 100:02d}{_mes_t1b:02d}')

print(f"Periodo de cuota: {_periodo_cuota}  |  Referencia: {_periodo_ref} ({_periodo_ref_yymm})  |  T1: [{_periodo_t1b_yymm}, {_periodo_ref_yymm}]")

# ════════════════════════════════════════════════════════════════════════
# PARAMETROS DE ENTRADA  ← ajustar cada mes
# ════════════════════════════════════════════════════════════════════════
DIAS_LAB_MES_ACTUAL    = 25      # días laborales del mes de referencia
DIAS_LAB_TRANSCURRIDOS = 23      # días laborales transcurridos al corte del CSV
DIAS_LAB_MES_CUOTA     = 25      # días laborales del mes para el que se genera cuota
CRECIMIENTO            = 0.05    # crecimiento esperado (5%)
CUOTA_GERENCIA         = None    # cuota fijada por gerencia (None = usar proyectado*crecimiento)

# ════════════════════════════════════════════════════════════════════════
# CONSTANTES DEL PROYECTO
# ════════════════════════════════════════════════════════════════════════
SSFF_CATS         = ['CERDO', 'HUEVO', 'PAVO', 'POLLO', 'PROCESADOS']
LINEAS_EMBUTIDOS  = ['EMBUTIDOS']
LINEAS_CONGELADOS = ['ELABORADOS', 'SEMIELABORADOS', 'PRECOCIDOS']
PERIODOS_T1       = [_periodo_t1b_yymm, _periodo_ref_yymm]
FFVV_ORDER        = ['F8', 'M0', 'KB', 'P0', 'V0']
FFVV_KB           = 'KB'

CATS_VALIDAS = ['ACCESORIOS','ANDINA','CERDO','COLGATE','DERMODIS','DULFINA',
                'HIGIENE Y CUIDADO','HOMEPRO PERU','HUEVO','KIMBERLY','LA CORONA',
                'LA PATRONA','MEDIFARMA','PAVO','POLLO','PROCESADOS','RINTI',
                'TAMBOS PERU','VERDUM','YICHANG']

SSFF_ORDEN  = sorted(SSFF_CATS)
RESTO_ORDEN = sorted([c for c in CATS_VALIDAS if c not in SSFF_CATS])
CATS_ORDEN  = SSFF_ORDEN + RESTO_ORDEN

OUTPUT_CUOTAS = f'C:/proyectos/SSFF/cuotas_ssff_{_nombre_cuota}.xlsx'
OUTPUT_BBDD   = f'C:/proyectos/SSFF/CUOTAS_BBDD_{_periodo_cuota}.xlsx'

# ── CARGA ──────────────────────────────────────────────────────────────
print("[1] Cargando datos...")

hist = pd.read_csv('C:/proyectos/SSFF/export_data_ssff.csv', header=None,
                   names=['PERIODO','CATEGORIA','LINEA','PEDIDOS','VOLUMEN','MONTO'])
hist['PERIODO'] = hist['PERIODO'].astype(int)

det = pd.read_csv('C:/proyectos/SSFF/export_data_ssff_rutas.csv', low_memory=False)
det.columns = det.columns.str.strip()
det['mes']           = det['mes'].astype(int)
det['total_monto']   = pd.to_numeric(det['total_monto'],   errors='coerce').fillna(0)
det['total_volumen'] = pd.to_numeric(det['total_volumen'], errors='coerce').fillna(0)
det['pedidos']       = pd.to_numeric(det['pedidos'],       errors='coerce').fillna(0)
det = det[det['categoria'].isin(CATS_VALIDAS)]

tablas = pd.read_excel('C:/proyectos/SSFF/TABLAS_RUTAS.xlsx', sheet_name='RUTA_ACTUAL')
rutas_oficiales = tablas['RUTA'].astype(str).str.strip().tolist()
for v in ['V001', 'V002']:
    if v not in rutas_oficiales:
        rutas_oficiales.append(v)

cartera = pd.read_excel('C:/proyectos/SSFF/cartera_simplificada.xlsx')
cartera = cartera.rename(columns={'codigo': 'ccod_cli', 'ruta': 'ccod_ruta'})
cartera['es_censo'] = cartera['censo'].notna()

# Cuotas del mes anterior (para hoja VERSUS) — soporta formato nuevo (CUOTAS_SOLES, header=3) y antiguo (CUOTAS, header=4)
_ref_file = f'C:/proyectos/SSFF/cuotas_ssff_{_nombre_ref}.xlsx'
try:
    cuotas_abr_raw = pd.read_excel(_ref_file, sheet_name='CUOTAS_SOLES', header=3)
    _col_ffvv_ref = 'FFVV'
except Exception:
    cuotas_abr_raw = pd.read_excel(_ref_file, sheet_name='CUOTAS', header=4)
    _col_ffvv_ref = 'RUT'
cuotas_abr_raw = cuotas_abr_raw[cuotas_abr_raw['RUTA'].notna() & cuotas_abr_raw[_col_ffvv_ref].notna()].copy()
cuotas_abr_raw = cuotas_abr_raw[~cuotas_abr_raw['RUTA'].astype(str).str.contains('TOTAL', na=False)]
cuotas_abr_raw = cuotas_abr_raw[~cuotas_abr_raw[_col_ffvv_ref].astype(str).str.startswith('SUBTOTAL')]
cuotas_abr_raw = cuotas_abr_raw.set_index('RUTA')
CATS_ABR = [c for c in CATS_ORDEN if c in cuotas_abr_raw.columns]

# Cuotas oficiales SSFF mayo (fijadas por gerencia comercial)
oficial_raw = pd.read_excel('C:/proyectos/SSFF/cuota_oficial_ssff.xlsx')
oficial_raw.columns = oficial_raw.columns.str.strip().str.lower()
# Estructura: mes, categoria, linea, kilos, soles, cobertura
# Totales por categoria (soles y kilos)
oficial_soles = oficial_raw.groupby('categoria')['soles'].sum()
oficial_kilos = oficial_raw.groupby('categoria')['kilos'].sum()
# Totales por linea (para EMBUTIDOS y CONGELADOS)
oficial_linea_kilos = oficial_raw.groupby('linea')['kilos'].sum()
oficial_linea_cob   = oficial_raw.groupby('linea')['cobertura'].sum()
# Dicts listos para usar en pasos 6 y 7
vol_oficial = {'EMBUTIDOS': float(oficial_linea_kilos.get('EMBUTIDOS', 0)),
               'CONGELADOS': float(oficial_linea_kilos.get('CONGELADOS', 0))}
cob_oficial = {'EMBUTIDOS': float(oficial_linea_cob.get('EMBUTIDOS', 0)),
               'CONGELADOS': float(oficial_linea_cob.get('CONGELADOS', 0))}

CUOTA_MINIMA_RUTA = 100_000   # piso absoluto por ruta (todas las categorias)

print(f"   Historico global: {len(hist):,} | Detalle rutas: {len(det):,} | Rutas: {len(rutas_oficiales)}")
print(f"   Cuotas oficiales SSFF cargadas: {oficial_soles.to_dict()}")

# Ajuste de periodos si el mes de referencia aún no existe en el CSV
_periodos_csv = sorted(hist['PERIODO'].unique())
_ultimo_csv   = _periodos_csv[-1]
if _periodo_ref_yymm not in _periodos_csv:
    print(f"   AVISO: periodo {_periodo_ref_yymm} no encontrado en CSV. Usando {_ultimo_csv} como referencia.")
    _desplazamiento    = _periodo_ref_yymm - _ultimo_csv  # cuántos meses adelante está
    _periodo_ref_yymm  = _ultimo_csv
    _periodo_t1b_yymm  = sorted(_periodos_csv)[-2] if len(_periodos_csv) >= 2 else _ultimo_csv
    PERIODOS_T1        = [_periodo_t1b_yymm, _periodo_ref_yymm]
    # Recalcular labels para que los textos sean correctos
    _anio_r2 = 2000 + int(str(_periodo_ref_yymm)[:2])
    _mes_ref2 = int(str(_periodo_ref_yymm)[2:])
    _label_ref   = f'{_MESES_ES_LARGO[_mes_ref2]} {_anio_r2}'
    _nombre_ref  = f'{_MESES_ES[_mes_ref2]}{_anio_r2}'
    print(f"   Periodos T1 ajustados: {PERIODOS_T1}  |  Referencia: {_label_ref}")

# ── PROYECCION LINEAL DEL MES DE REFERENCIA ────────────────────────────
print(f"\n[2] Proyeccion lineal de {_label_ref} y objetivo {_label_cuota}...")
real_abril_global = hist[hist['PERIODO'] == _periodo_ref_yymm]['MONTO'].sum()
# Proyeccion lineal: (real / dias_transcurridos) * dias_totales_mes
PROJ_ABRIL      = (real_abril_global / DIAS_LAB_TRANSCURRIDOS) * DIAS_LAB_MES_ACTUAL
factor_proy     = PROJ_ABRIL / real_abril_global
CUOTA_TOTAL_OBJ = CUOTA_GERENCIA if CUOTA_GERENCIA is not None else PROJ_ABRIL * (1 + CRECIMIENTO)

print(f"   Real CSV {_label_ref} ({DIAS_LAB_TRANSCURRIDOS} dias): S/ {real_abril_global:,.0f}")
print(f"   Proyectado lineal ({DIAS_LAB_MES_ACTUAL} dias): S/ {PROJ_ABRIL:,.0f}  (factor: {factor_proy:.4f})")
if CUOTA_GERENCIA:
    print(f"   Cuota {_label_cuota} (fijada por gerencia): S/ {CUOTA_TOTAL_OBJ:,.0f}")
else:
    print(f"   Cuota {_label_cuota} (+{CRECIMIENTO*100:.0f}%): S/ {CUOTA_TOTAL_OBJ:,.0f}")

# Proyectar abril en datasets
det_proy = det.copy()
mask_abr = det_proy['mes'] == _periodo_ref_yymm
det_proy.loc[mask_abr, 'total_monto']   *= factor_proy
det_proy.loc[mask_abr, 'total_volumen'] *= factor_proy
det_proy.loc[mask_abr, 'pedidos']       *= factor_proy

hist_proy = hist.copy()
mask_abr_h = hist_proy['PERIODO'] == _periodo_ref_yymm
hist_proy.loc[mask_abr_h, 'MONTO']   *= factor_proy
hist_proy.loc[mask_abr_h, 'VOLUMEN'] *= factor_proy

# Proyectado lineal por ruta (para criterio minimo y columna extra en CUOTAS_SOLES)
abr_real_ruta = (det[det['mes'] == _periodo_ref_yymm]
                 .groupby('ccod_ruta')['total_monto'].sum())
proj_ruta = {}
for ruta in rutas_oficiales:
    real_r = abr_real_ruta.get(ruta, 0)
    proj_ruta[ruta] = (real_r / DIAS_LAB_TRANSCURRIDOS) * DIAS_LAB_MES_ACTUAL

# ── CUOTA TIPO 1 POR CATEGORIA (global) ───────────────────────────────
print("\n[3] Cuota Tipo 1 por categoria (global)...")
base_global = hist_proy[
    (hist_proy['PERIODO'].isin(PERIODOS_T1)) &
    (hist_proy['CATEGORIA'].isin(CATS_VALIDAS))
]

def tipo1(s): return (s.mean() + s.max()) / 2

def cuota_cat_global(df, col):
    pivot = df.groupby(['CATEGORIA', 'PERIODO'])[col].sum().unstack(fill_value=0)
    for p in PERIODOS_T1:
        if p not in pivot.columns:
            pivot[p] = 0
    return pivot[PERIODOS_T1].apply(tipo1, axis=1)

cuota_monto_cat = cuota_cat_global(base_global, 'MONTO')
cuota_vol_cat   = cuota_cat_global(base_global, 'VOLUMEN')

for cat in CATS_VALIDAS:
    if cat not in cuota_monto_cat.index:
        cuota_monto_cat[cat] = 0.0
        cuota_vol_cat[cat]   = 0.0

# Fijar cuotas oficiales SSFF (reemplazan el Tipo 1 para esas categorias)
ssff_oficial_total = 0.0
for cat in SSFF_CATS:
    if cat in oficial_soles.index:
        cuota_monto_cat[cat] = float(oficial_soles[cat])
        ssff_oficial_total  += float(oficial_soles[cat])

# Ajustar el resto (no-SSFF) para alcanzar CUOTA_TOTAL_OBJ
resto_tipo1_sum = sum(cuota_monto_cat.get(c, 0) for c in CATS_VALIDAS if c not in SSFF_CATS)
objetivo_resto  = CUOTA_TOTAL_OBJ - ssff_oficial_total
factor_ajuste   = objetivo_resto / resto_tipo1_sum if resto_tipo1_sum > 0 else 1.0
for cat in CATS_VALIDAS:
    if cat not in SSFF_CATS:
        cuota_monto_cat[cat] = cuota_monto_cat.get(cat, 0) * factor_ajuste
        cuota_vol_cat[cat]   = cuota_vol_cat.get(cat, 0)   * factor_ajuste
# Volumen SSFF: usar kilos oficiales para PROCESADOS; el resto de SSFF no tiene vol objetivo
for cat in SSFF_CATS:
    if cat in oficial_kilos.index:
        cuota_vol_cat[cat] = float(oficial_kilos[cat])
    else:
        cuota_vol_cat[cat] = cuota_vol_cat.get(cat, 0)

cuota_total_calculada = sum(cuota_monto_cat.get(c, 0) for c in CATS_VALIDAS)
print(f"   SSFF oficial: S/ {ssff_oficial_total:,.0f} | Resto ajustado (x{factor_ajuste:.4f}): S/ {objetivo_resto:,.0f}")
print(f"   Total cuota calculada: S/ {cuota_total_calculada:,.0f}  |  Objetivo: S/ {CUOTA_TOTAL_OBJ:,.0f}")

# ── TIPO 1 POR RUTA ────────────────────────────────────────────────────
print("\n[4] Tipo 1 por ruta x categoria...")
base_det = det_proy[det_proy['mes'].isin(PERIODOS_T1)]

def tipo1_ruta_pivot(df, col):
    agg = df.groupby(['ccod_ruta', 'categoria', 'mes'])[col].sum().unstack(fill_value=0)
    for p in PERIODOS_T1:
        if p not in agg.columns:
            agg[p] = 0
    return agg[PERIODOS_T1].apply(tipo1, axis=1).unstack(fill_value=0)

tipo1_ruta_monto = tipo1_ruta_pivot(base_det, 'total_monto')
tipo1_ruta_vol   = tipo1_ruta_pivot(base_det, 'total_volumen')

# Cobertura: clientes unicos
cob_raw = (base_det.groupby(['ccod_ruta', 'categoria', 'mes'])['ccod_cli']
           .nunique().unstack(fill_value=0))
for p in PERIODOS_T1:
    if p not in cob_raw.columns:
        cob_raw[p] = 0
tipo1_ruta_cob = cob_raw[PERIODOS_T1].apply(tipo1, axis=1).unstack(fill_value=0)

# Completar rutas y categorias faltantes
for df_t in [tipo1_ruta_monto, tipo1_ruta_vol, tipo1_ruta_cob]:
    for ruta in rutas_oficiales:
        if ruta not in df_t.index:
            df_t.loc[ruta] = 0.0
    for cat in CATS_VALIDAS:
        if cat not in df_t.columns:
            df_t[cat] = 0.0

tipo1_ruta_monto = tipo1_ruta_monto.loc[rutas_oficiales, CATS_VALIDAS].copy()
tipo1_ruta_vol   = tipo1_ruta_vol.loc[rutas_oficiales, CATS_VALIDAS].copy()
tipo1_ruta_cob   = tipo1_ruta_cob.loc[rutas_oficiales, CATS_VALIDAS].copy()

# ── PESOS Y RESTRICCION KB ─────────────────────────────────────────────
print("\n[5] Pesos y restriccion KB en SSFF...")
rutas_kb = [r for r in rutas_oficiales if r.startswith(FFVV_KB)]

for cat in SSFF_CATS:
    tipo1_ruta_monto.loc[rutas_kb, cat] = 0.0
    tipo1_ruta_vol.loc[rutas_kb, cat]   = 0.0
    tipo1_ruta_cob.loc[rutas_kb, cat]   = 0.0

def calcular_pesos(df_t1):
    tots = df_t1.sum(axis=0)
    return df_t1.div(tots, axis=1).fillna(0)

pesos_monto = calcular_pesos(tipo1_ruta_monto)
pesos_vol   = calcular_pesos(tipo1_ruta_vol)
pesos_cob   = calcular_pesos(tipo1_ruta_cob)

# ── DISTRIBUCION POR RUTA ──────────────────────────────────────────────
print("\n[6] Distribuyendo cuota...")

dist_monto = pd.DataFrame(index=rutas_oficiales, columns=CATS_VALIDAS, dtype=float)
dist_vol   = pd.DataFrame(index=rutas_oficiales, columns=CATS_VALIDAS, dtype=float)
dist_cob   = pd.DataFrame(index=rutas_oficiales, columns=CATS_VALIDAS, dtype=float)

cob_global_cat = tipo1_ruta_cob.sum(axis=0)
# Cobertura oficial PROCESADOS = suma de lineas oficiales
cob_oficial_procesados = sum(cob_oficial.values())

for cat in CATS_VALIDAS:
    dist_monto[cat] = (pesos_monto[cat] * cuota_monto_cat.get(cat, 0)).round(0)
    dist_vol[cat]   = (pesos_vol[cat]   * cuota_vol_cat.get(cat, 0)).round(0)
    # Para PROCESADOS usar total oficial; resto usa Tipo 1
    if cat == 'PROCESADOS':
        dist_cob[cat] = (pesos_cob[cat] * cob_oficial_procesados).round(0)
    else:
        dist_cob[cat] = (pesos_cob[cat] * cob_global_cat.get(cat, 0)).round(0)

# FFVV map
ruta_ffvv_map = {}
for r in rutas_oficiales:
    if r.startswith('KB'):
        ruta_ffvv_map[r] = 'KB'
    elif r.startswith('M0'):
        ruta_ffvv_map[r] = 'M0'
    elif r.startswith('P0'):
        ruta_ffvv_map[r] = 'P0'
    elif r.startswith('V0') or r in ['V001', 'V002']:
        ruta_ffvv_map[r] = 'V0'
    else:
        ruta_ffvv_map[r] = 'F8'

dist_monto['FFVV']    = pd.Series(ruta_ffvv_map)
dist_monto['GENERAL'] = dist_monto[CATS_VALIDAS].sum(axis=1)
dist_cob['FFVV']      = pd.Series(ruta_ffvv_map)
dist_cob['GENERAL']   = dist_cob[CATS_VALIDAS].sum(axis=1)

print(f"   Total monto antes ajuste: S/ {dist_monto['GENERAL'].sum():,.0f}  |  Objetivo: S/ {CUOTA_TOTAL_OBJ:,.0f}")

# ── CRITERIO MINIMO 100K por ruta (sin redistribucion, total sube libremente) ──
print("\n[6b] Aplicando piso minimo S/ 100K por ruta...")
exceso_total = 0.0
rutas_con_floor = []
for ruta in rutas_oficiales:
    cuota_r = dist_monto.loc[ruta, 'GENERAL']
    if cuota_r < CUOTA_MINIMA_RUTA:
        exceso_total += CUOTA_MINIMA_RUTA - cuota_r
        rutas_con_floor.append(ruta)
        if cuota_r > 0:
            # Escalar proporcionalmente todas las categorias de la ruta
            factor_r = CUOTA_MINIMA_RUTA / cuota_r
            for cat in CATS_VALIDAS:
                dist_monto.loc[ruta, cat] = round(dist_monto.loc[ruta, cat] * factor_r, 0)
        else:
            # Sin historial: distribuir 100K según peso global de cada categoria
            total_global = cuota_monto_cat.sum()
            for cat in CATS_VALIDAS:
                peso_cat = cuota_monto_cat.get(cat, 0) / total_global if total_global > 0 else 1/len(CATS_VALIDAS)
                dist_monto.loc[ruta, cat] = round(CUOTA_MINIMA_RUTA * peso_cat, 0)
        dist_monto.loc[ruta, 'GENERAL'] = dist_monto.loc[ruta, CATS_VALIDAS].sum()

dist_monto['GENERAL'] = dist_monto[CATS_VALIDAS].sum(axis=1)
totales_ffvv     = dist_monto.groupby('FFVV')[CATS_VALIDAS + ['GENERAL']].sum()
totales_ffvv_cob = dist_cob.groupby('FFVV')[CATS_VALIDAS + ['GENERAL']].sum()

cuota_final_total = dist_monto['GENERAL'].sum()
print(f"   Rutas con floor aplicado: {len(rutas_con_floor)} | Monto elevado: S/ {exceso_total:,.0f}")
print(f"   Total monto final: S/ {cuota_final_total:,.0f}  |  Objetivo base: S/ {CUOTA_TOTAL_OBJ:,.0f}")

# ── VOLUMEN Y COBERTURA EMBUTIDOS / CONGELADOS (F8 + M0) ──────────────
print("\n[7] Volumen PROCESADOS (F8 + M0)...")
rutas_f8 = sorted([r for r in rutas_oficiales if ruta_ffvv_map[r] in ('F8', 'M0')])

# Pesos por ruta desde Tipo 1 (F8 + M0, PROCESADOS)
base_proc_r = det_proy[
    (det_proy['mes'].isin(PERIODOS_T1)) &
    (det_proy['categoria'] == 'PROCESADOS') &
    (det_proy['ccod_ruta'].isin(rutas_f8))
].copy()
base_proc_r['GRUPO'] = np.where(
    base_proc_r['linea'].isin(LINEAS_EMBUTIDOS), 'EMBUTIDOS',
    np.where(base_proc_r['linea'].isin(LINEAS_CONGELADOS), 'CONGELADOS', 'OTROS')
)

# Tipo 1 volumen por ruta x grupo
agg_vp = (base_proc_r[base_proc_r['GRUPO'] != 'OTROS']
          .groupby(['ccod_ruta', 'GRUPO', 'mes'])['total_volumen']
          .sum().unstack(fill_value=0))
for p in PERIODOS_T1:
    if p not in agg_vp.columns:
        agg_vp[p] = 0
t1_vol_proc = agg_vp[PERIODOS_T1].apply(tipo1, axis=1).unstack(fill_value=0)
for r in rutas_f8:
    if r not in t1_vol_proc.index:
        t1_vol_proc.loc[r] = 0.0
t1_vol_proc = t1_vol_proc.loc[rutas_f8].fillna(0)

# Tipo 1 cobertura (clientes unicos) por ruta x grupo
agg_cp = (base_proc_r[base_proc_r['GRUPO'] != 'OTROS']
          .groupby(['ccod_ruta', 'GRUPO', 'mes'])['ccod_cli']
          .nunique().unstack(fill_value=0))
for p in PERIODOS_T1:
    if p not in agg_cp.columns:
        agg_cp[p] = 0
t1_cob_proc = agg_cp[PERIODOS_T1].apply(tipo1, axis=1).unstack(fill_value=0)
for r in rutas_f8:
    if r not in t1_cob_proc.index:
        t1_cob_proc.loc[r] = 0.0
t1_cob_proc = t1_cob_proc.loc[rutas_f8].fillna(0)

dist_vol_linea = pd.DataFrame(index=rutas_f8, dtype=float)
dist_cob_linea = pd.DataFrame(index=rutas_f8, dtype=float)
for grupo in ['EMBUTIDOS', 'CONGELADOS']:
    # Volumen
    if grupo in t1_vol_proc.columns:
        tot_v = t1_vol_proc[grupo].sum()
        peso_v = t1_vol_proc[grupo] / tot_v if tot_v > 0 else pd.Series(1/len(rutas_f8), index=rutas_f8)
    else:
        peso_v = pd.Series(1/len(rutas_f8), index=rutas_f8)
    dist_vol_linea[grupo] = (peso_v * vol_oficial[grupo]).round(0)
    # Cobertura
    if grupo in t1_cob_proc.columns:
        tot_c = t1_cob_proc[grupo].sum()
        peso_c = t1_cob_proc[grupo] / tot_c if tot_c > 0 else pd.Series(1/len(rutas_f8), index=rutas_f8)
    else:
        peso_c = pd.Series(1/len(rutas_f8), index=rutas_f8)
    dist_cob_linea[grupo] = (peso_c * cob_oficial[grupo]).round(0)

dist_vol_linea['TOTAL'] = (dist_vol_linea['EMBUTIDOS'] + dist_vol_linea['CONGELADOS']).round(0)
dist_cob_linea['TOTAL'] = (dist_cob_linea['EMBUTIDOS'] + dist_cob_linea['CONGELADOS']).round(0)

print(f"   EMBUTIDOS: {dist_vol_linea['EMBUTIDOS'].sum():,.0f} kg (cob {dist_cob_linea['EMBUTIDOS'].sum():,.0f}) | "
      f"CONGELADOS: {dist_vol_linea['CONGELADOS'].sum():,.0f} kg (cob {dist_cob_linea['CONGELADOS'].sum():,.0f})")

# ── HISTORICO POR CATEGORIA ────────────────────────────────────────────
print("\n[8] Preparando historico por categoria...")
hist_cats = hist[hist['CATEGORIA'].isin(CATS_VALIDAS)]
hist_pivot = (hist_cats.groupby(['PERIODO', 'CATEGORIA'])['MONTO']
              .sum().unstack(fill_value=0))
hist_pivot['TOTAL'] = hist_pivot.sum(axis=1)
for cat in CATS_ORDEN:
    if cat not in hist_pivot.columns:
        hist_pivot[cat] = 0
hist_pivot = hist_pivot[CATS_ORDEN + ['TOTAL']]
hist_pivot = hist_pivot.sort_index()

hist_pct = hist_pivot.pct_change().fillna(0)

# Historico volumen EMBUTIDOS y CONGELADOS (PROCESADOS, todas las lineas)
hist_proc = hist[hist['CATEGORIA'] == 'PROCESADOS'].copy()
hist_proc['GRUPO'] = np.where(
    hist_proc['LINEA'].isin(LINEAS_EMBUTIDOS), 'EMBUTIDOS',
    np.where(hist_proc['LINEA'].isin(LINEAS_CONGELADOS), 'CONGELADOS', None)
)
hist_proc = hist_proc[hist_proc['GRUPO'].notna()]
hist_vol_pivot = (hist_proc.groupby(['PERIODO', 'GRUPO'])['VOLUMEN']
                  .sum().unstack(fill_value=0))
for g in ['EMBUTIDOS', 'CONGELADOS']:
    if g not in hist_vol_pivot.columns:
        hist_vol_pivot[g] = 0
hist_vol_pivot = hist_vol_pivot[['EMBUTIDOS', 'CONGELADOS']]
hist_vol_pivot['TOTAL'] = hist_vol_pivot.sum(axis=1)
hist_vol_pivot = hist_vol_pivot.sort_index()

# Proyectado mes de referencia por ruta (para tabla 3 HISTORICO)
det_mar = det[det['mes'] == _periodo_t1b_yymm]
det_abr = det[det['mes'] == _periodo_ref_yymm].copy()
det_abr['total_monto'] *= factor_proy

mar_ruta = (det_mar.groupby(['ccod_ruta', 'categoria'])['total_monto']
            .sum().unstack(fill_value=0))
abr_ruta = (det_abr.groupby(['ccod_ruta', 'categoria'])['total_monto']
            .sum().unstack(fill_value=0))

for df_r in [mar_ruta, abr_ruta]:
    for r in rutas_oficiales:
        if r not in df_r.index:
            df_r.loc[r] = 0.0
    for cat in CATS_ORDEN:
        if cat not in df_r.columns:
            df_r[cat] = 0.0

mar_ruta = mar_ruta.loc[rutas_oficiales, CATS_ORDEN].copy()
abr_ruta = abr_ruta.loc[rutas_oficiales, CATS_ORDEN].copy()
mar_ruta['TOTAL'] = mar_ruta.sum(axis=1)
abr_ruta['TOTAL'] = abr_ruta.sum(axis=1)

# Cuota mayo por categoria (para fila final en tabla 1 HISTORICO)
cuota_mayo_cat = {cat: int(round(cuota_monto_cat.get(cat, 0))) for cat in CATS_ORDEN}
cuota_mayo_total = int(round(cuota_final_total))

# ════════════════════════════════════════════════════════════════════════
# HELPERS EXCEL
# ════════════════════════════════════════════════════════════════════════
def fill(h):  return PatternFill('solid', fgColor=h)
def fnt(bold=False, color='FF000000', sz=10): return Font(name='Aptos Narrow', bold=bold, color=color, size=sz)
def aln(h='center'): return Alignment(horizontal=h, vertical='center', wrap_text=True)
def brd():
    s = Side(style='thin', color='FFCCCCCC')
    return Border(left=s, right=s, top=s, bottom=s)

C_HDR   = 'FF1F3864'
C_FFVV  = 'FF2F5496'
C_TOTAL = 'FFDAE3F3'
C_PESO  = 'FFFFD966'
C_SSFF  = 'FF70AD47'
C_GEN   = 'FFED7D31'
C_BLC   = 'FFFFFFFF'
C_GRIS  = 'FFF2F2F2'
C_ROJO  = 'FFFFC7CE'
C_VERDE = 'FFC6EFCE'
C_PROJ  = 'FFD9EAD3'   # verde claro para columna proyectado
C_VAR   = 'FFFCE5CD'   # naranja claro para columna variacion


def escribir_hoja_cuota(wb, nombre_hoja, titulo, dist_df, totales_ffvv_df,
                        cuota_cat_series, total_obj, metrica='#,##0',
                        con_proyectado=False):
    """Genera una hoja estilo CUOTAS para monto o cobertura.
    con_proyectado=True agrega columnas PROYECTADO y VAR% con formulas Excel.
    """
    ws = wb.create_sheet(nombre_hoja)
    ncats = len(CATS_ORDEN)
    # col base: B=2, C=3, D..W=4..4+ncats-1, X=4+ncats (GENERAL)
    # Si con_proyectado: Y=4+ncats+1 (PROYECTADO), Z=4+ncats+2 (VAR%)
    col_gen  = 4 + ncats          # columna indice de GENERAL
    col_proj = col_gen + 1        # columna PROYECTADO
    col_var  = col_gen + 2        # columna VAR%
    last_col_idx = col_var if con_proyectado else col_gen
    last_col = get_column_letter(last_col_idx)

    # ── Fila 1: titulo
    ws.merge_cells(f'B1:{last_col}1')
    ws['B1'] = titulo
    ws['B1'].font = fnt(bold=True, color='FFFFFFFF', sz=12)
    ws['B1'].fill = fill(C_HDR)
    ws['B1'].alignment = aln()
    ws.row_dimensions[1].height = 24

    # ── Fila 2: PESO % (con formulas que suman las filas de datos)
    ws['B2'] = 'PESO %'
    ws['B2'].fill = fill(C_PESO)
    ws['B2'].font = fnt(bold=True, sz=9)
    for i, cat in enumerate(CATS_ORDEN):
        c = ws.cell(row=2, column=4+i)
        # Formula: columna_cat_total / columna_GENERAL_total  (fila 3)
        col_ltr = get_column_letter(4+i)
        gen_ltr = get_column_letter(col_gen)
        c.value = f'={col_ltr}3/{gen_ltr}3'
        c.number_format = '0.00%'
        c.fill = fill(C_PESO)
        c.alignment = aln()
        c.font = fnt(sz=9)
    # Celda PESO en columna GENERAL = 100%
    c_gen2 = ws.cell(row=2, column=col_gen)
    c_gen2.value = 1.0
    c_gen2.number_format = '0.00%'
    c_gen2.fill = fill(C_PESO)
    c_gen2.font = fnt(bold=True, sz=9)
    if con_proyectado:
        ws.cell(row=2, column=col_proj).fill = fill(C_PROJ)
        ws.cell(row=2, column=col_var).fill = fill(C_VAR)

    # ── Fila 3: TOTAL GLOBAL (con formulas SUM de las filas de datos)
    ws['B3'] = 'TOTAL GLOBAL'
    ws['B3'].fill = fill(C_TOTAL)
    ws['B3'].font = fnt(bold=True, sz=9)

    # Necesitamos saber el rango de filas de datos para las formulas
    # Filas de datos comienzan en fila 5, terminan en fila 5 + len(rutas) - 1
    fila_data_ini = 5
    fila_data_fin = fila_data_ini + len(rutas_oficiales) - 1

    for i, cat in enumerate(CATS_ORDEN):
        c = ws.cell(row=3, column=4+i)
        col_ltr = get_column_letter(4+i)
        c.value = f'=SUM({col_ltr}{fila_data_ini}:{col_ltr}{fila_data_fin})'
        c.number_format = metrica
        c.fill = fill(C_TOTAL)
        c.font = fnt(bold=True, sz=9)
        c.alignment = aln('right')
    gen_ltr = get_column_letter(col_gen)
    c_gen3 = ws.cell(row=3, column=col_gen)
    c_gen3.value = f'=SUM({gen_ltr}{fila_data_ini}:{gen_ltr}{fila_data_fin})'
    c_gen3.number_format = metrica
    c_gen3.fill = fill(C_GEN)
    c_gen3.font = fnt(bold=True, sz=9)
    if con_proyectado:
        proj_ltr = get_column_letter(col_proj)
        var_ltr  = get_column_letter(col_var)
        c_proj3 = ws.cell(row=3, column=col_proj)
        c_proj3.value = f'=SUM({proj_ltr}{fila_data_ini}:{proj_ltr}{fila_data_fin})'
        c_proj3.number_format = metrica
        c_proj3.fill = fill(C_PROJ)
        c_proj3.font = fnt(bold=True, sz=9)
        c_var3 = ws.cell(row=3, column=col_var)
        c_var3.value = f'={gen_ltr}3/{proj_ltr}3-1'
        c_var3.number_format = '0.00%'
        c_var3.fill = fill(C_VAR)
        c_var3.font = fnt(bold=True, sz=9)

    # ── Fila 4: cabeceras
    hdrs = ['FFVV', 'RUTA'] + CATS_ORDEN + ['GENERAL']
    if con_proyectado:
        hdrs += ['PROYECTADO', 'VAR%']
    for j, h in enumerate(hdrs):
        c = ws.cell(row=4, column=2+j)
        c.value = h
        if h in SSFF_CATS:
            bg = C_SSFF
        elif h == 'GENERAL':
            bg = C_GEN
        elif h in ('PROYECTADO', 'VAR%'):
            bg = C_PROJ if h == 'PROYECTADO' else C_VAR
        else:
            bg = C_HDR
        c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
        c.fill = fill(bg)
        c.alignment = aln()
        c.border = brd()
    ws.row_dimensions[4].height = 28
    ws.column_dimensions['B'].width = 5
    ws.column_dimensions['C'].width = 7
    for i in range(ncats + 1):
        ws.column_dimensions[get_column_letter(4+i)].width = 12
    if con_proyectado:
        ws.column_dimensions[get_column_letter(col_proj)].width = 13
        ws.column_dimensions[get_column_letter(col_var)].width = 9

    # ── Filas de datos por FFVV
    row = fila_data_ini
    for ffvv in FFVV_ORDER:
        rutas_ffvv = sorted([r for r in rutas_oficiales if ruta_ffvv_map[r] == ffvv])
        for ruta in rutas_ffvv:
            f = fill(C_GRIS if row % 2 == 0 else C_BLC)
            ws.cell(row=row, column=2, value=ffvv).fill = f
            c = ws.cell(row=row, column=3, value=ruta)
            c.fill = f
            c.font = fnt(bold=True, sz=9)
            # Categorias
            for i, cat in enumerate(CATS_ORDEN):
                c = ws.cell(row=row, column=4+i,
                            value=int(round(dist_df.loc[ruta, cat])))
                c.number_format = metrica
                c.fill = f
                c.alignment = aln('right')
                c.font = fnt(sz=9)
            # GENERAL con formula SUM de la fila
            gen_ltr = get_column_letter(col_gen)
            first_cat_ltr = get_column_letter(4)
            last_cat_ltr  = get_column_letter(4 + ncats - 1)
            c = ws.cell(row=row, column=col_gen)
            c.value = f'=SUM({first_cat_ltr}{row}:{last_cat_ltr}{row})'
            c.number_format = metrica
            c.fill = fill(C_GEN)
            c.font = fnt(bold=True, sz=9)
            c.alignment = aln('right')
            # PROYECTADO y VAR%
            if con_proyectado:
                proj_val = int(round(proj_ruta.get(ruta, 0)))
                gen_ltr  = get_column_letter(col_gen)
                proj_ltr = get_column_letter(col_proj)
                c_proj = ws.cell(row=row, column=col_proj, value=proj_val)
                c_proj.number_format = metrica
                c_proj.fill = fill(C_PROJ)
                c_proj.font = fnt(sz=9)
                c_proj.alignment = aln('right')
                c_var = ws.cell(row=row, column=col_var)
                c_var.value = f'={gen_ltr}{row}/{proj_ltr}{row}-1'
                c_var.number_format = '0.00%'
                c_var.fill = fill(C_VAR)
                c_var.font = fnt(sz=9)
                c_var.alignment = aln('right')
            row += 1

    # ── Tabla SUMA POR FFVV
    row += 1
    ws.cell(row=row-1, column=2, value='SUMA POR FFVV').font = fnt(bold=True)
    ffvv_start_row = row
    for ffvv in FFVV_ORDER:
        if ffvv not in totales_ffvv_df.index:
            continue
        ws.cell(row=row, column=3, value=ffvv).fill = fill(C_FFVV)
        ws.cell(row=row, column=3).font = fnt(bold=True, color='FFFFFFFF', sz=9)
        # Rutas de esta FFVV para formula SUMIF
        rutas_ffvv = sorted([r for r in rutas_oficiales if ruta_ffvv_map[r] == ffvv])
        # Rango de filas de esta FFVV en la tabla de datos
        ffvv_rows = [fila_data_ini + rutas_oficiales.index(r)
                     if r in rutas_oficiales else None for r in rutas_ffvv]
        ffvv_rows = [x for x in ffvv_rows if x is not None]
        # Usamos formula SUM de las filas especificas de esa FFVV
        for i, cat in enumerate(CATS_ORDEN):
            col_ltr = get_column_letter(4+i)
            refs = ','.join(f'{col_ltr}{fr}' for fr in ffvv_rows)
            c = ws.cell(row=row, column=4+i)
            c.value = f'=SUM({refs})' if refs else 0
            c.number_format = metrica
            c.fill = fill(C_FFVV)
            c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
            c.alignment = aln('right')
        # GENERAL de la FFVV
        gen_ltr = get_column_letter(col_gen)
        refs_gen = ','.join(f'{gen_ltr}{fr}' for fr in ffvv_rows)
        c = ws.cell(row=row, column=col_gen)
        c.value = f'=SUM({refs_gen})' if refs_gen else 0
        c.number_format = metrica
        c.fill = fill(C_GEN)
        c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
        c.alignment = aln('right')
        # PROYECTADO de la FFVV
        if con_proyectado:
            proj_ltr = get_column_letter(col_proj)
            refs_proj = ','.join(f'{proj_ltr}{fr}' for fr in ffvv_rows)
            c = ws.cell(row=row, column=col_proj)
            c.value = f'=SUM({refs_proj})' if refs_proj else 0
            c.number_format = metrica
            c.fill = fill(C_PROJ)
            c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
            c.alignment = aln('right')
            var_ltr  = get_column_letter(col_var)
            gen_ltr  = get_column_letter(col_gen)
            c = ws.cell(row=row, column=col_var)
            c.value = f'={gen_ltr}{row}/{proj_ltr}{row}-1'
            c.number_format = '0.00%'
            c.fill = fill(C_VAR)
            c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
            c.alignment = aln('right')
        row += 1

    # Fila TOTAL (suma de toda la tabla de datos)
    ws.cell(row=row, column=3, value='TOTAL').fill = fill(C_HDR)
    ws.cell(row=row, column=3).font = fnt(bold=True, color='FFFFFFFF', sz=10)
    for i in range(ncats):
        col_ltr = get_column_letter(4+i)
        c = ws.cell(row=row, column=4+i)
        c.value = f'=SUM({col_ltr}{fila_data_ini}:{col_ltr}{fila_data_fin})'
        c.number_format = metrica
        c.fill = fill(C_HDR)
        c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
        c.alignment = aln('right')
    gen_ltr = get_column_letter(col_gen)
    c = ws.cell(row=row, column=col_gen)
    c.value = f'=SUM({gen_ltr}{fila_data_ini}:{gen_ltr}{fila_data_fin})'
    c.number_format = metrica
    c.fill = fill(C_HDR)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=10)
    c.alignment = aln('right')
    if con_proyectado:
        proj_ltr = get_column_letter(col_proj)
        var_ltr  = get_column_letter(col_var)
        c = ws.cell(row=row, column=col_proj)
        c.value = f'=SUM({proj_ltr}{fila_data_ini}:{proj_ltr}{fila_data_fin})'
        c.number_format = metrica
        c.fill = fill(C_HDR)
        c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
        c.alignment = aln('right')
        c = ws.cell(row=row, column=col_var)
        c.value = f'={gen_ltr}{row}/{proj_ltr}{row}-1'
        c.number_format = '0.00%'
        c.fill = fill(C_HDR)
        c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
        c.alignment = aln('right')

    ws.freeze_panes = 'D5'
    return ws


# ════════════════════════════════════════════════════════════════════════
# GENERAR EXCEL
# ════════════════════════════════════════════════════════════════════════
print("\n[9] Generando Excel...")
wb = Workbook()
wb.remove(wb.active)

# ── HOJA PROCESO ──────────────────────────────────────────────────────
print("   > PROCESO")
ws_proc = wb.create_sheet('PROCESO')
ws_proc.column_dimensions['A'].width = 28
ws_proc.column_dimensions['B'].width = 55

def proc_titulo(ws, r, texto):
    ws.merge_cells(f'A{r}:B{r}')
    c = ws.cell(row=r, column=1, value=texto)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=11)
    c.fill = fill(C_HDR)
    c.alignment = aln('left')
    ws.row_dimensions[r].height = 20
    return r + 1

def proc_subtitulo(ws, r, texto):
    ws.merge_cells(f'A{r}:B{r}')
    c = ws.cell(row=r, column=1, value=texto)
    c.font = fnt(bold=True, sz=10)
    c.fill = fill(C_TOTAL)
    c.alignment = aln('left')
    return r + 1

def proc_fila(ws, r, campo, valor, bold_val=False):
    c1 = ws.cell(row=r, column=1, value=campo)
    c1.font = fnt(bold=True, sz=9)
    c1.fill = fill(C_GRIS if r % 2 == 0 else C_BLC)
    c1.alignment = Alignment(horizontal='left', vertical='center')
    c2 = ws.cell(row=r, column=2, value=valor)
    c2.font = fnt(bold=bold_val, sz=9)
    c2.fill = fill(C_GRIS if r % 2 == 0 else C_BLC)
    c2.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws.row_dimensions[r].height = 18
    return r + 1

r = 1
r = proc_titulo(ws_proc, r, f'PROCESO DE GENERACION DE CUOTA — {_label_cuota.upper()}')
r += 1

r = proc_subtitulo(ws_proc, r, '1. OBJETIVO GENERAL')
r = proc_fila(ws_proc, r, 'Periodo de cuota', f'{_label_cuota} ({_periodo_cuota})')
r = proc_fila(ws_proc, r, 'Objetivo total (soles)', f'S/ {cuota_final_total:,.0f}  (sin IGV)', True)
r = proc_fila(ws_proc, r, 'Objetivo volumen SSFF — EMBUTIDOS',
              f'{dist_vol_linea["EMBUTIDOS"].sum():,.0f} kg', True)
r = proc_fila(ws_proc, r, 'Objetivo volumen SSFF — CONGELADOS',
              f'{dist_vol_linea["CONGELADOS"].sum():,.0f} kg', True)
r = proc_fila(ws_proc, r, 'Objetivo volumen SSFF — TOTAL PROCESADOS',
              f'{dist_vol_linea["TOTAL"].sum():,.0f} kg', True)
r += 1

r = proc_subtitulo(ws_proc, r, '2. PARAMETROS UTILIZADOS')
r = proc_fila(ws_proc, r, f'Dias lab. mes actual ({_label_ref})', str(DIAS_LAB_MES_ACTUAL))
r = proc_fila(ws_proc, r, 'Dias lab. transcurridos (corte CSV)', str(DIAS_LAB_TRANSCURRIDOS))
r = proc_fila(ws_proc, r, f'Dias lab. mes de cuota ({_label_cuota})', str(DIAS_LAB_MES_CUOTA))
r = proc_fila(ws_proc, r, 'Crecimiento esperado', f'{CRECIMIENTO*100:.1f}%')
r = proc_fila(ws_proc, r, f'Real acumulado {_label_ref} (CSV)',
              f'S/ {real_abril_global:,.0f}')
r = proc_fila(ws_proc, r, f'Proyectado lineal {_label_ref}',
              f'S/ {PROJ_ABRIL:,.0f}  =  ({real_abril_global:,.0f} / {DIAS_LAB_TRANSCURRIDOS}) x {DIAS_LAB_MES_ACTUAL}')
if CUOTA_GERENCIA:
    r = proc_fila(ws_proc, r, 'Cuota total objetivo',
                  f'S/ {CUOTA_TOTAL_OBJ:,.0f}  (fijada por Gerencia)', True)
else:
    r = proc_fila(ws_proc, r, 'Cuota total objetivo',
                  f'S/ {CUOTA_TOTAL_OBJ:,.0f}  =  proyectado x (1 + {CRECIMIENTO*100:.0f}%)')
r += 1

r = proc_subtitulo(ws_proc, r, '3. METODOLOGIA — TIPO 1 (Promedio/Maximo)')
r = proc_fila(ws_proc, r, 'Formula Tipo 1',
              'CUOTA = ( PROMEDIO(periodos) + MAX(periodos) ) / 2')
periodos_str = ', '.join(str(p) for p in PERIODOS_T1[:-1]) + f', {PERIODOS_T1[-1]}* (* = proyectado lineal)'
r = proc_fila(ws_proc, r, 'Periodos usados', periodos_str)
r = proc_fila(ws_proc, r, 'Nota metodologica',
              'Febrero 2026 (2602) excluido: mes atipicamente bajo por efectos estacionales.')
r = proc_fila(ws_proc, r, 'Aplicacion global',
              'Tipo 1 calculado sobre historico agregado por categoria. '
              'Se ajusta proporcionalmente para que la suma iguale el objetivo total.')
r = proc_fila(ws_proc, r, 'Aplicacion por ruta',
              'Tipo 1 calculado por ruta x categoria desde export_data_ssff_rutas.csv '
              '(1.29M registros, nivel cliente x ruta x categoria). '
              'El peso de cada ruta = Tipo1_ruta / sum(Tipo1_todas_rutas) por categoria.')
r = proc_fila(ws_proc, r, 'Cobertura (clientes unicos)',
              'Tipo 1 aplicado sobre conteo de ccod_cli unicos por ruta x categoria '
              'en los mismos 3 periodos.')
r += 1

r = proc_subtitulo(ws_proc, r, '4. REGLAS DE DISTRIBUCION')
r = proc_fila(ws_proc, r, 'Rutas oficiales',
              f'{len(rutas_oficiales)} rutas: 55 F8, 14 M0, 7 KB, 1 P0, 2 V0 '
              '(77 de TABLAS_RUTAS + V001 + V002)')
r = proc_fila(ws_proc, r, 'Restriccion FFVV KB',
              'Las rutas KB NO reciben cuota en categorias SSFF '
              '(CERDO, HUEVO, PAVO, POLLO, PROCESADOS). '
              'La cuota de esas categorias se redistribuye entre F8, M0, P0 y V0.')
r = proc_fila(ws_proc, r, 'Volumen PROCESADOS',
              'Solo aplica a las 55 rutas F8. Distribuido por linea: '
              'EMBUTIDOS y CONGELADOS (ELABORADOS + SEMIELABORADOS).')
r = proc_fila(ws_proc, r, 'Criterio minimo por ruta',
              f'Piso absoluto S/ {CUOTA_MINIMA_RUTA:,.0f} por ruta (GENERAL). '
              'Rutas bajo el piso se elevan escalando sus categorias proporcionalmente. '
              'El total general sube libremente; no se redistribuye el exceso.')
r = proc_fila(ws_proc, r, 'Rutas con floor aplicado',
              f'{len(rutas_con_floor)} rutas  |  Exceso redistribuido: S/ {exceso_total:,.0f}')
r = proc_fila(ws_proc, r, 'Cuotas oficiales SSFF',
              'Los soles de CERDO, HUEVO, PAVO, POLLO, PROCESADOS y los kg/cobertura '
              'de EMBUTIDOS y CONGELADOS provienen del archivo cuota_oficial_ssff.xlsx. '
              'Las categorias restantes se ajustan proporcionalmente para alcanzar el objetivo total.')
r += 1

r = proc_subtitulo(ws_proc, r, '5. FUENTES DE DATOS')
r = proc_fila(ws_proc, r, 'Historico global',
              'export_data_ssff.csv — 1,149 registros. Nivel: periodo x categoria x linea.')
r = proc_fila(ws_proc, r, 'Historico por ruta',
              'export_data_ssff_rutas.csv — 1.29M registros. '
              'Nivel: cliente x ruta x categoria x linea x mes (2501-2604).')
r = proc_fila(ws_proc, r, 'Rutas oficiales',
              'TABLAS_RUTAS.xlsx, hoja RUTA_ACTUAL.')
r = proc_fila(ws_proc, r, 'Cartera de clientes',
              'cartera_simplificada.xlsx — 24,715 clientes, incluye 5,930 del censo '
              '(reestructuracion 08/03/2026).')
r += 1

r = proc_subtitulo(ws_proc, r, '6. ARCHIVOS GENERADOS')
r = proc_fila(ws_proc, r, 'Presentacion Gerencia', f'cuotas_ssff_{_nombre_cuota}.xlsx')
r = proc_fila(ws_proc, r, '  > Hoja CUOTAS_SOLES', f'Cuota en soles por ruta x categoria + proyectado {_label_ref}')
r = proc_fila(ws_proc, r, '  > Hoja CUOTA_COBERTURA', 'Cuota de clientes unicos por ruta x categoria')
r = proc_fila(ws_proc, r, '  > Hoja CUOTA_VOL_SSFF',
              'Cuota kg y cobertura EMBUTIDOS y CONGELADOS — FFVV F8 y M0. '
              'CONGELADOS = ELABORADOS + SEMIELABORADOS + PRECOCIDOS. '
              'Cobertura TOTAL calza con PROCESADOS en hoja CUOTA_COBERTURA.')
r = proc_fila(ws_proc, r, '  > Hoja HISTORICO', f'Historico mensual, variacion % y tabla {_MESES_ES_LARGO[_mes_t1b]} vs {_label_ref} por ruta')
r = proc_fila(ws_proc, r, '  > Hoja VERSUS',
              f'Comparativo por ruta: cuota {_label_ref} vs proyectado {_label_ref} vs cuota {_label_cuota}. '
              'Resumen por categoria: proyectado vs cuota para estimar cierre.')
r = proc_fila(ws_proc, r, 'BBDD Sistemas', f'CUOTAS_BBDD_{_periodo_cuota}.xlsx')
r = proc_fila(ws_proc, r, f'  > Hoja {_periodo_cuota}',
              'Estructura relacional: 79 filas GENERAL + 1,580 filas LINEA = 1,659 filas')

# ── HOJA CUOTAS_SOLES ─────────────────────────────────────────────────
print("   > CUOTAS_SOLES")
escribir_hoja_cuota(
    wb, 'CUOTAS_SOLES',
    f'CUOTA VENTAS MAYO 2026 — SOLES (S/ sin IGV)  |  Objetivo: S/ {cuota_final_total:,.0f}',
    dist_monto, totales_ffvv,
    cuota_monto_cat, cuota_final_total, '#,##0',
    con_proyectado=True
)

# ── HOJA CUOTA_COBERTURA ──────────────────────────────────────────────
print("   > CUOTA_COBERTURA")
cuota_cob_total_cat = dist_cob[CATS_VALIDAS].sum()
total_cob_obj = cuota_cob_total_cat.sum()

escribir_hoja_cuota(
    wb, 'CUOTA_COBERTURA',
    'CUOTA COBERTURA MAYO 2026 — CLIENTES UNICOS POR CATEGORIA',
    dist_cob, totales_ffvv_cob,
    cuota_cob_total_cat, total_cob_obj, '#,##0',
    con_proyectado=False
)

# ── HOJA CUOTA_VOL_SSFF ───────────────────────────────────────────────
# Columnas: FFVV | RUTA | EMB kg | EMB cob | CON kg | CON cob | TOTAL kg | TOTAL cob
# col B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9
print("   > CUOTA_VOL_SSFF")
ws_vol = wb.create_sheet('CUOTA_VOL_SSFF')
ws_vol.merge_cells('B1:I1')
ws_vol['B1'] = 'CUOTA VOLUMEN MAYO 2026 — PROCESADOS: EMBUTIDOS y CONGELADOS (kg + cobertura) — FFVV F8 y M0'
ws_vol['B1'].font = fnt(bold=True, color='FFFFFFFF', sz=12)
ws_vol['B1'].fill = fill(C_SSFF)
ws_vol['B1'].alignment = aln()
ws_vol.row_dimensions[1].height = 24

# Fila 2: totales
ws_vol['B2'] = 'TOTAL F8'
ws_vol['B2'].fill = fill(C_TOTAL); ws_vol['B2'].font = fnt(bold=True, sz=9)
vol_totals = [('EMBUTIDOS', 'kg', dist_vol_linea['EMBUTIDOS'].sum(), C_TOTAL),
              ('EMBUTIDOS', 'cob', dist_cob_linea['EMBUTIDOS'].sum(), C_TOTAL),
              ('CONGELADOS', 'kg', dist_vol_linea['CONGELADOS'].sum(), C_TOTAL),
              ('CONGELADOS', 'cob', dist_cob_linea['CONGELADOS'].sum(), C_TOTAL),
              ('TOTAL', 'kg', dist_vol_linea['TOTAL'].sum(), C_GEN),
              ('TOTAL', 'cob', dist_cob_linea['TOTAL'].sum(), C_GEN)]
for k, (_, _, val, bg) in enumerate(vol_totals):
    c = ws_vol.cell(row=2, column=4+k, value=int(round(val)))
    c.number_format = '#,##0'; c.fill = fill(bg); c.font = fnt(bold=True, sz=9); c.alignment = aln('right')

# Fila 3: cabeceras
hdrs_vol = ['FFVV', 'RUTA', 'EMB kg', 'EMB cob', 'CON kg', 'CON cob', 'TOTAL kg', 'TOTAL cob']
for j, h in enumerate(hdrs_vol):
    c = ws_vol.cell(row=3, column=2+j, value=h)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
    c.fill = fill(C_SSFF); c.alignment = aln(); c.border = brd()
ws_vol.row_dimensions[3].height = 26
ws_vol.column_dimensions['B'].width = 5
ws_vol.column_dimensions['C'].width = 8
for k in range(6):
    ws_vol.column_dimensions[get_column_letter(4+k)].width = 12

r = 4
for ruta in rutas_f8:
    f = fill(C_GRIS if r % 2 == 0 else C_BLC)
    ws_vol.cell(row=r, column=2, value='F8').fill = f
    c = ws_vol.cell(row=r, column=3, value=ruta); c.fill = f; c.font = fnt(bold=True, sz=9)
    row_vals = [
        (dist_vol_linea.loc[ruta, 'EMBUTIDOS'],  f,        '#,##0'),
        (dist_cob_linea.loc[ruta, 'EMBUTIDOS'],  f,        '#,##0'),
        (dist_vol_linea.loc[ruta, 'CONGELADOS'], f,        '#,##0'),
        (dist_cob_linea.loc[ruta, 'CONGELADOS'], f,        '#,##0'),
        (dist_vol_linea.loc[ruta, 'TOTAL'],      fill(C_GEN), '#,##0'),
        (dist_cob_linea.loc[ruta, 'TOTAL'],      fill(C_GEN), '#,##0'),
    ]
    for k, (val, bg, fmt) in enumerate(row_vals):
        c = ws_vol.cell(row=r, column=4+k, value=int(round(val)))
        c.number_format = fmt; c.fill = bg; c.alignment = aln('right')
        c.font = fnt(bold=(k >= 4), sz=9)
    r += 1

r += 1
ws_vol.cell(row=r, column=3, value='TOTAL F8').fill = fill(C_HDR)
ws_vol.cell(row=r, column=3).font = fnt(bold=True, color='FFFFFFFF')
total_vals = [dist_vol_linea['EMBUTIDOS'].sum(), dist_cob_linea['EMBUTIDOS'].sum(),
              dist_vol_linea['CONGELADOS'].sum(), dist_cob_linea['CONGELADOS'].sum(),
              dist_vol_linea['TOTAL'].sum(), dist_cob_linea['TOTAL'].sum()]
for k, val in enumerate(total_vals):
    c = ws_vol.cell(row=r, column=4+k, value=int(round(val)))
    c.number_format = '#,##0'; c.fill = fill(C_HDR)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=9); c.alignment = aln('right')
ws_vol.freeze_panes = 'D4'

# Nota de coherencia con hoja CUOTA_COBERTURA
r += 2
ws_vol.cell(row=r, column=2,
    value='Nota: TOTAL cob (EMB + CON) debe coincidir con cobertura PROCESADOS en hoja CUOTA_COBERTURA')
ws_vol.cell(row=r, column=2).font = fnt(sz=8)

# ── HOJA HISTORICO ────────────────────────────────────────────────────
print("   > HISTORICO")
ws_h = wb.create_sheet('HISTORICO')
ws_h.column_dimensions['A'].width = 10
for i in range(len(CATS_ORDEN) + 2):
    ws_h.column_dimensions[get_column_letter(2+i)].width = 12

periodos_disp = sorted(hist_pivot.index.tolist())
ncats_h = len(CATS_ORDEN)

# ─ Tabla 1: Soles por periodo x categoria
r = 1
ws_h.merge_cells(f'A{r}:{get_column_letter(2+ncats_h)}{r}')
ws_h.cell(row=r, column=1).value = 'HISTORICO VENTAS POR CATEGORIA — SOLES (S/ sin IGV)'
ws_h.cell(row=r, column=1).font = fnt(bold=True, color='FFFFFFFF', sz=12)
ws_h.cell(row=r, column=1).fill = fill(C_HDR)
ws_h.cell(row=r, column=1).alignment = aln()
ws_h.row_dimensions[r].height = 22
r += 1

ws_h.cell(row=r, column=1).value = 'PERIODO'
ws_h.cell(row=r, column=1).fill = fill(C_HDR)
ws_h.cell(row=r, column=1).font = fnt(bold=True, color='FFFFFFFF', sz=9)
ws_h.cell(row=r, column=1).alignment = aln()
for i, cat in enumerate(CATS_ORDEN):
    c = ws_h.cell(row=r, column=2+i)
    c.value = cat
    c.fill = fill(C_SSFF if cat in SSFF_CATS else C_HDR)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
    c.alignment = aln()
    c.border = brd()
ws_h.cell(row=r, column=2+ncats_h).value = 'TOTAL'
ws_h.cell(row=r, column=2+ncats_h).fill = fill(C_GEN)
ws_h.cell(row=r, column=2+ncats_h).font = fnt(bold=True, color='FFFFFFFF', sz=9)
ws_h.cell(row=r, column=2+ncats_h).alignment = aln()
ws_h.row_dimensions[r].height = 28
r += 1

for idx, periodo in enumerate(periodos_disp):
    f = fill(C_GRIS if idx % 2 == 0 else C_BLC)
    ws_h.cell(row=r, column=1, value=str(periodo)).fill = f
    ws_h.cell(row=r, column=1).font = fnt(bold=True, sz=9)
    for i, cat in enumerate(CATS_ORDEN):
        c = ws_h.cell(row=r, column=2+i, value=int(round(hist_pivot.loc[periodo, cat])))
        c.number_format = '#,##0'
        c.fill = f
        c.alignment = aln('right')
        c.font = fnt(sz=9)
    c = ws_h.cell(row=r, column=2+ncats_h, value=int(round(hist_pivot.loc[periodo, 'TOTAL'])))
    c.number_format = '#,##0'
    c.fill = fill(C_GEN)
    c.font = fnt(bold=True, sz=9)
    c.alignment = aln('right')
    r += 1

# Fila proyectado abril
f = fill(C_PESO)
ws_h.cell(row=r, column=1, value=f'{_periodo_ref_yymm}*').fill = f
ws_h.cell(row=r, column=1).font = fnt(bold=True, sz=9)
proj_cat_h = hist_proy[hist_proy['PERIODO'] == _periodo_ref_yymm].groupby('CATEGORIA')['MONTO'].sum()
for i, cat in enumerate(CATS_ORDEN):
    c = ws_h.cell(row=r, column=2+i, value=int(round(proj_cat_h.get(cat, 0))))
    c.number_format = '#,##0'
    c.fill = f
    c.alignment = aln('right')
    c.font = fnt(sz=9)
c = ws_h.cell(row=r, column=2+ncats_h, value=int(round(proj_cat_h.sum())))
c.number_format = '#,##0'
c.fill = fill(C_GEN)
c.font = fnt(bold=True, sz=9)
c.alignment = aln('right')
r += 1

# Fila cuota del mes de cuota
f = fill(C_VERDE)
ws_h.cell(row=r, column=1, value=f'{_periodo_cuota} CUOTA').fill = f
ws_h.cell(row=r, column=1).font = fnt(bold=True, sz=9)
for i, cat in enumerate(CATS_ORDEN):
    c = ws_h.cell(row=r, column=2+i, value=cuota_mayo_cat.get(cat, 0))
    c.number_format = '#,##0'
    c.fill = f
    c.alignment = aln('right')
    c.font = fnt(bold=True, sz=9)
c = ws_h.cell(row=r, column=2+ncats_h, value=cuota_mayo_total)
c.number_format = '#,##0'
c.fill = fill(C_SSFF)
c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
c.alignment = aln('right')
r += 1

ws_h.cell(row=r, column=1, value=f'* {_periodo_ref_yymm} = proyectado lineal  |  {_periodo_cuota} CUOTA = cuota asignada {_label_cuota}')
ws_h.cell(row=r, column=1).font = fnt(sz=8)
r += 2

# ─ Tabla 2: % variacion
ws_h.merge_cells(f'A{r}:{get_column_letter(2+ncats_h)}{r}')
ws_h.cell(row=r, column=1).value = 'VARIACION % MENSUAL POR CATEGORIA'
ws_h.cell(row=r, column=1).font = fnt(bold=True, color='FFFFFFFF', sz=12)
ws_h.cell(row=r, column=1).fill = fill(C_FFVV)
ws_h.cell(row=r, column=1).alignment = aln()
ws_h.row_dimensions[r].height = 22
r += 1

ws_h.cell(row=r, column=1, value='PERIODO').fill = fill(C_HDR)
ws_h.cell(row=r, column=1).font = fnt(bold=True, color='FFFFFFFF', sz=9)
for i, cat in enumerate(CATS_ORDEN):
    c = ws_h.cell(row=r, column=2+i)
    c.value = cat
    c.fill = fill(C_SSFF if cat in SSFF_CATS else C_HDR)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
    c.alignment = aln()
    c.border = brd()
ws_h.cell(row=r, column=2+ncats_h, value='TOTAL').fill = fill(C_GEN)
ws_h.cell(row=r, column=2+ncats_h).font = fnt(bold=True, color='FFFFFFFF', sz=9)
ws_h.row_dimensions[r].height = 28
r += 1

periodos_pct = periodos_disp[1:]
for idx, periodo in enumerate(periodos_pct):
    prev = periodos_disp[periodos_disp.index(periodo) - 1]
    f = fill(C_GRIS if idx % 2 == 0 else C_BLC)
    ws_h.cell(row=r, column=1, value=str(periodo)).fill = f
    ws_h.cell(row=r, column=1).font = fnt(bold=True, sz=9)
    for i, cat in enumerate(CATS_ORDEN):
        v_act  = hist_pivot.loc[periodo, cat]
        v_prev = hist_pivot.loc[prev, cat]
        pct = (v_act - v_prev) / v_prev if v_prev != 0 else 0
        c = ws_h.cell(row=r, column=2+i, value=pct)
        c.number_format = '0.00%'
        c.fill = fill(C_VERDE) if pct >= 0 else fill(C_ROJO)
        c.alignment = aln('right')
        c.font = fnt(sz=9)
    tot_act  = hist_pivot.loc[periodo, 'TOTAL']
    tot_prev = hist_pivot.loc[prev, 'TOTAL']
    pct_tot  = (tot_act - tot_prev) / tot_prev if tot_prev != 0 else 0
    c = ws_h.cell(row=r, column=2+ncats_h, value=pct_tot)
    c.number_format = '0.00%'
    c.fill = fill(C_VERDE) if pct_tot >= 0 else fill(C_ROJO)
    c.font = fnt(bold=True, sz=9)
    c.alignment = aln('right')
    r += 1
r += 2

# ─ Tabla 3: Rutas — Cierre Marzo vs Proyectado Abril
ws_h.merge_cells(f'A{r}:{get_column_letter(2+ncats_h+3)}{r}')
ws_h.cell(row=r, column=1).value = 'CIERRE MARZO 2026 vs PROYECTADO ABRIL 2026 — POR RUTA'
ws_h.cell(row=r, column=1).font = fnt(bold=True, color='FFFFFFFF', sz=12)
ws_h.cell(row=r, column=1).fill = fill(C_HDR)
ws_h.cell(row=r, column=1).alignment = aln()
ws_h.row_dimensions[r].height = 22
r += 1

ws_h.cell(row=r, column=1, value='RUTA').fill = fill(C_HDR)
ws_h.cell(row=r, column=1).font = fnt(bold=True, color='FFFFFFFF', sz=9)
for i, cat in enumerate(CATS_ORDEN):
    c = ws_h.cell(row=r, column=2+i)
    c.value = cat
    c.fill = fill(C_SSFF if cat in SSFF_CATS else C_HDR)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
    c.alignment = aln()
    c.border = brd()
extra_cols = ['TOTAL MAR', 'PROY ABR*', '% VAR']
for i, h in enumerate(extra_cols):
    c = ws_h.cell(row=r, column=2+ncats_h+i)
    c.value = h
    c.fill = fill(C_GEN)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
    c.alignment = aln()
    c.border = brd()
ws_h.row_dimensions[r].height = 28
ws_h.column_dimensions[get_column_letter(2+ncats_h)].width = 12
ws_h.column_dimensions[get_column_letter(2+ncats_h+1)].width = 12
ws_h.column_dimensions[get_column_letter(2+ncats_h+2)].width = 10
r += 1

for idx, ruta in enumerate(rutas_oficiales):
    f = fill(C_GRIS if idx % 2 == 0 else C_BLC)
    ws_h.cell(row=r, column=1, value=ruta).fill = f
    ws_h.cell(row=r, column=1).font = fnt(bold=True, sz=9)
    for i, cat in enumerate(CATS_ORDEN):
        c = ws_h.cell(row=r, column=2+i, value=int(round(mar_ruta.loc[ruta, cat])))
        c.number_format = '#,##0'
        c.fill = f
        c.alignment = aln('right')
        c.font = fnt(sz=9)
    tot_mar = int(round(mar_ruta.loc[ruta, 'TOTAL']))
    tot_abr = int(round(abr_ruta.loc[ruta, 'TOTAL']))
    pct_var  = (tot_abr - tot_mar) / tot_mar if tot_mar != 0 else 0
    c = ws_h.cell(row=r, column=2+ncats_h, value=tot_mar)
    c.number_format = '#,##0'
    c.fill = fill(C_GEN)
    c.font = fnt(bold=True, sz=9)
    c.alignment = aln('right')
    c = ws_h.cell(row=r, column=2+ncats_h+1, value=tot_abr)
    c.number_format = '#,##0'
    c.fill = fill(C_GEN)
    c.font = fnt(bold=True, sz=9)
    c.alignment = aln('right')
    c = ws_h.cell(row=r, column=2+ncats_h+2, value=pct_var)
    c.number_format = '0.00%'
    c.fill = fill(C_VERDE) if pct_var >= 0 else fill(C_ROJO)
    c.font = fnt(sz=9)
    c.alignment = aln('right')
    r += 1

# Total de la tabla de rutas
ws_h.cell(row=r, column=1, value='TOTAL').fill = fill(C_HDR)
ws_h.cell(row=r, column=1).font = fnt(bold=True, color='FFFFFFFF', sz=9)
for i, cat in enumerate(CATS_ORDEN):
    c = ws_h.cell(row=r, column=2+i, value=int(round(mar_ruta[cat].sum())))
    c.number_format = '#,##0'
    c.fill = fill(C_HDR)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
    c.alignment = aln('right')
tot_mar_total = int(round(mar_ruta['TOTAL'].sum()))
tot_abr_total = int(round(abr_ruta['TOTAL'].sum()))
pct_var_total = (tot_abr_total - tot_mar_total) / tot_mar_total if tot_mar_total != 0 else 0
for col_val, val, fmt in [
    (2+ncats_h,   tot_mar_total, '#,##0'),
    (2+ncats_h+1, tot_abr_total, '#,##0'),
    (2+ncats_h+2, pct_var_total, '0.00%'),
]:
    c = ws_h.cell(row=r, column=col_val, value=val)
    c.number_format = fmt
    c.fill = fill(C_HDR)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
    c.alignment = aln('right')

r += 3

# ─ Tabla 4: Historico volumen EMBUTIDOS y CONGELADOS
ws_h.merge_cells(f'A{r}:D{r}')
ws_h.cell(row=r, column=1).value = 'HISTORICO VOLUMEN PROCESADOS — EMBUTIDOS y CONGELADOS (kg)'
ws_h.cell(row=r, column=1).font = fnt(bold=True, color='FFFFFFFF', sz=12)
ws_h.cell(row=r, column=1).fill = fill(C_SSFF)
ws_h.cell(row=r, column=1).alignment = aln()
ws_h.row_dimensions[r].height = 22
r += 1

# Header
for j, h in enumerate(['PERIODO', 'EMBUTIDOS', 'CONGELADOS', 'TOTAL']):
    c = ws_h.cell(row=r, column=1+j, value=h)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
    c.fill = fill(C_SSFF if j > 0 else C_HDR)
    c.alignment = aln()
    c.border = brd()
ws_h.row_dimensions[r].height = 26
r += 1

periodos_vol = sorted(hist_vol_pivot.index.tolist())
for idx, periodo in enumerate(periodos_vol):
    f = fill(C_GRIS if idx % 2 == 0 else C_BLC)
    ws_h.cell(row=r, column=1, value=str(periodo)).fill = f
    ws_h.cell(row=r, column=1).font = fnt(bold=True, sz=9)
    for j, g in enumerate(['EMBUTIDOS', 'CONGELADOS', 'TOTAL']):
        c = ws_h.cell(row=r, column=2+j, value=int(round(hist_vol_pivot.loc[periodo, g])))
        c.number_format = '#,##0'
        c.fill = f if g != 'TOTAL' else fill(C_GEN)
        c.alignment = aln('right')
        c.font = fnt(bold=(g == 'TOTAL'), sz=9)
    r += 1

# Fila proyectado mes de referencia (volumen)
det_abr_vol = det[det['mes'] == _periodo_ref_yymm].copy()
det_abr_vol['total_volumen'] *= factor_proy
det_abr_vol['GRUPO'] = np.where(
    det_abr_vol['linea'].isin(LINEAS_EMBUTIDOS), 'EMBUTIDOS',
    np.where(det_abr_vol['linea'].isin(LINEAS_CONGELADOS), 'CONGELADOS', None)
)
det_abr_vol = det_abr_vol[det_abr_vol['GRUPO'].notna()]
proj_vol_abr = det_abr_vol.groupby('GRUPO')['total_volumen'].sum()
emb_proy  = int(round(proj_vol_abr.get('EMBUTIDOS', 0)))
con_proy  = int(round(proj_vol_abr.get('CONGELADOS', 0)))
tot_v_proy = emb_proy + con_proy

f = fill(C_PESO)
ws_h.cell(row=r, column=1, value=f'{_periodo_ref_yymm}*').fill = f
ws_h.cell(row=r, column=1).font = fnt(bold=True, sz=9)
for j, val in enumerate([emb_proy, con_proy, tot_v_proy]):
    c = ws_h.cell(row=r, column=2+j, value=val)
    c.number_format = '#,##0'
    c.fill = f if j < 2 else fill(C_GEN)
    c.alignment = aln('right')
    c.font = fnt(bold=(j == 2), sz=9)
r += 1

# Fila cuota volumen
emb_cuota  = int(round(dist_vol_linea['EMBUTIDOS'].sum()))
con_cuota  = int(round(dist_vol_linea['CONGELADOS'].sum()))
tot_v_cuota = int(round(dist_vol_linea['TOTAL'].sum()))
f = fill(C_VERDE)
ws_h.cell(row=r, column=1, value=f'{_periodo_cuota} CUOTA').fill = f
ws_h.cell(row=r, column=1).font = fnt(bold=True, sz=9)
for j, val in enumerate([emb_cuota, con_cuota, tot_v_cuota]):
    c = ws_h.cell(row=r, column=2+j, value=val)
    c.number_format = '#,##0'
    c.fill = f if j < 2 else fill(C_SSFF)
    c.alignment = aln('right')
    c.font = fnt(bold=True, sz=9, color='FF000000' if j < 2 else 'FFFFFFFF')
r += 1

ws_h.cell(row=r, column=1,
          value=f'* {_periodo_ref_yymm} = proyectado lineal  |  {_periodo_cuota} CUOTA = cuota asignada {_label_cuota}').font = fnt(sz=8)
r += 1

ws_h.freeze_panes = 'B3'

# ── HOJA VERSUS ──────────────────────────────────────────────────────
print("   > VERSUS")
ws_v = wb.create_sheet('VERSUS')

# Proyectado abril por ruta (total general)
proj_ruta_ser = pd.Series(proj_ruta)

# Cuota mes anterior por ruta (del archivo cuotas_ssff del mes de referencia)
cuota_abr_ruta = {}
_cats_abr_num = [c for c in CATS_ABR if c in cuotas_abr_raw.columns]
for ruta in rutas_oficiales:
    if ruta in cuotas_abr_raw.index:
        if 'GENERAL' in cuotas_abr_raw.columns:
            val = pd.to_numeric(cuotas_abr_raw.loc[ruta, 'GENERAL'], errors='coerce')
        else:
            val = float('nan')
        # Si GENERAL es NaN o 0 (formula no evaluada), recalcular desde categorias
        if pd.isna(val) or val == 0:
            val = sum(pd.to_numeric(cuotas_abr_raw.loc[ruta, c], errors='coerce') or 0
                      for c in _cats_abr_num)
        cuota_abr_ruta[ruta] = float(val) if not pd.isna(val) else 0.0
    else:
        cuota_abr_ruta[ruta] = 0.0

# Cuota mayo por ruta
cuota_may_ruta = dist_monto['GENERAL'].to_dict()

# Tabla rutas: cols A-G (1-7). Tabla categorias: cols J-P (10-16), 2 cols de separacion (H, I).
COL_R  = 1   # inicio tabla rutas
COL_C  = 10  # inicio tabla categorias (J)

# Calcular proyectado mes de referencia por categoria (desde det con factor)
proj_cat_abr = (det_proy[det_proy['mes'] == _periodo_ref_yymm]
                .groupby('categoria')['total_monto'].sum())
# Cuota abril por categoria (desde el archivo)
cuota_abr_cat = {}
for cat in CATS_ORDEN:
    if cat in cuotas_abr_raw.columns:
        cuota_abr_cat[cat] = float(cuotas_abr_raw[cat].sum())
    else:
        cuota_abr_cat[cat] = 0.0

# ─ Titulo fila 1: ambas tablas
ws_v.merge_cells(f'{get_column_letter(COL_R)}1:{get_column_letter(COL_R+6)}1')
ws_v.cell(row=1, column=COL_R).value = 'VERSUS: CUOTA ABRIL vs PROYECTADO ABRIL vs CUOTA MAYO 2026'
ws_v.cell(row=1, column=COL_R).font = fnt(bold=True, color='FFFFFFFF', sz=12)
ws_v.cell(row=1, column=COL_R).fill = fill(C_HDR)
ws_v.cell(row=1, column=COL_R).alignment = aln()
ws_v.row_dimensions[1].height = 24

ws_v.merge_cells(f'{get_column_letter(COL_C)}1:{get_column_letter(COL_C+6)}1')
ws_v.cell(row=1, column=COL_C).value = 'RESUMEN POR CATEGORIA — PROYECTADO ABRIL vs CUOTA MAYO'
ws_v.cell(row=1, column=COL_C).font = fnt(bold=True, color='FFFFFFFF', sz=11)
ws_v.cell(row=1, column=COL_C).fill = fill(C_FFVV)
ws_v.cell(row=1, column=COL_C).alignment = aln()

# ─ Encabezados tabla rutas (fila 2)
hdrs_v = ['FFVV', 'RUTA', 'CUOTA ABR', 'PROY ABR*', '% CUM ABR', 'CUOTA MAYO', '% vs PROY']
col_bg_v = [C_HDR, C_HDR, C_PESO, C_PROJ, C_VAR, C_GEN, C_VAR]
for j, (h, bg) in enumerate(zip(hdrs_v, col_bg_v)):
    c = ws_v.cell(row=2, column=COL_R+j, value=h)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
    c.fill = fill(bg); c.alignment = aln(); c.border = brd()

# ─ Encabezados tabla categorias (fila 2)
hdrs_cat = ['CATEGORIA', 'PROY ABR*', 'CUOTA MAYO', '% vs PROY', 'CUOTA ABR', '% CUM ABR', 'DIFERENCIA']
col_bg_cat = [C_HDR, C_PROJ, C_GEN, C_VAR, C_PESO, C_VAR, C_TOTAL]
for j, (h, bg) in enumerate(zip(hdrs_cat, col_bg_cat)):
    c = ws_v.cell(row=2, column=COL_C+j, value=h)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
    c.fill = fill(bg); c.alignment = aln(); c.border = brd()
ws_v.row_dimensions[2].height = 26

# Anchos de columna
anchos_v = [5, 7, 13, 13, 10, 13, 10]
for j, w in enumerate(anchos_v):
    ws_v.column_dimensions[get_column_letter(COL_R+j)].width = w
# separacion cols H, I
ws_v.column_dimensions['H'].width = 2
ws_v.column_dimensions['I'].width = 2
anchos_cat = [15, 13, 13, 10, 13, 10, 13]
for j, w in enumerate(anchos_cat):
    ws_v.column_dimensions[get_column_letter(COL_C+j)].width = w

# ─ Filas datos tabla rutas (fila 3 en adelante)
rv = 3
for ffvv in FFVV_ORDER:
    rutas_ffvv = sorted([r for r in rutas_oficiales if ruta_ffvv_map[r] == ffvv])
    for ruta in rutas_ffvv:
        f = fill(C_GRIS if rv % 2 == 0 else C_BLC)
        c_abr  = cuota_abr_ruta.get(ruta, 0)
        p_abr  = proj_ruta.get(ruta, 0)
        c_may  = cuota_may_ruta.get(ruta, 0)
        var_abr  = (p_abr / c_abr - 1) if c_abr > 0 else 0
        var_mayo = (c_may / p_abr - 1) if p_abr > 0 else 0

        ws_v.cell(row=rv, column=COL_R,   value=ffvv).fill = f
        ws_v.cell(row=rv, column=COL_R+1, value=ruta).fill = f
        ws_v.cell(row=rv, column=COL_R+1).font = fnt(bold=True, sz=9)

        c = ws_v.cell(row=rv, column=COL_R+2, value=int(round(c_abr)))
        c.number_format = '#,##0'; c.fill = fill(C_PESO); c.alignment = aln('right'); c.font = fnt(sz=9)
        c = ws_v.cell(row=rv, column=COL_R+3, value=int(round(p_abr)))
        c.number_format = '#,##0'; c.fill = fill(C_PROJ); c.alignment = aln('right'); c.font = fnt(sz=9)
        c = ws_v.cell(row=rv, column=COL_R+4, value=var_abr)
        c.number_format = '0.00%'; c.fill = fill(C_VERDE if var_abr >= 0 else C_ROJO)
        c.alignment = aln('right'); c.font = fnt(sz=9)
        c = ws_v.cell(row=rv, column=COL_R+5, value=int(round(c_may)))
        c.number_format = '#,##0'; c.fill = fill(C_GEN); c.alignment = aln('right'); c.font = fnt(bold=True, sz=9)
        c = ws_v.cell(row=rv, column=COL_R+6, value=var_mayo)
        c.number_format = '0.00%'; c.fill = fill(C_VERDE if var_mayo >= 0 else C_ROJO)
        c.alignment = aln('right'); c.font = fnt(sz=9)
        rv += 1

# Fila TOTAL tabla rutas
fila_total_v = rv
tot_cabr  = sum(cuota_abr_ruta.get(r, 0) for r in rutas_oficiales)
tot_pabr  = sum(proj_ruta.get(r, 0) for r in rutas_oficiales)
tot_cmay  = sum(cuota_may_ruta.get(r, 0) for r in rutas_oficiales)
var_abr_t = (tot_pabr / tot_cabr - 1) if tot_cabr > 0 else 0
var_may_t = (tot_cmay / tot_pabr - 1) if tot_pabr > 0 else 0
datos_total_r = [('TOTAL', None), (None, None),
                 (int(round(tot_cabr)), '#,##0'), (int(round(tot_pabr)), '#,##0'),
                 (var_abr_t, '0.00%'), (int(round(tot_cmay)), '#,##0'), (var_may_t, '0.00%')]
for j, (val, fmt) in enumerate(datos_total_r):
    c = ws_v.cell(row=fila_total_v, column=COL_R+j, value=val)
    c.fill = fill(C_HDR); c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
    if fmt: c.number_format = fmt
    c.alignment = aln('right') if j > 0 else aln('left')

# ─ Filas datos tabla categorias (fila 3 en adelante, mismas filas que rutas)
# Calcula fila de inicio categorias = fila 3, independiente de cuantas rutas haya
rc = 3
for idx, cat in enumerate(CATS_ORDEN):
    f = fill(C_SSFF if cat in SSFF_CATS else (C_GRIS if idx % 2 == 0 else C_BLC))
    font_cat = fnt(bold=(cat in SSFF_CATS), sz=9)
    p_abr_c = proj_cat_abr.get(cat, 0)
    c_may_c = cuota_monto_cat.get(cat, 0)
    c_abr_c = cuota_abr_cat.get(cat, 0)
    var_p   = (c_may_c / p_abr_c - 1) if p_abr_c > 0 else 0
    var_a   = (p_abr_c / c_abr_c - 1) if c_abr_c > 0 else 0
    diff    = c_may_c - p_abr_c

    vals = [cat, int(round(p_abr_c)), int(round(c_may_c)), var_p,
            int(round(c_abr_c)), var_a, int(round(diff))]
    fmts = [None, '#,##0', '#,##0', '0.00%', '#,##0', '0.00%', '#,##0']
    for j, (val, fmt) in enumerate(zip(vals, fmts)):
        c = ws_v.cell(row=rc, column=COL_C+j, value=val)
        c.fill = fill(C_VERDE if (j == 3 and var_p >= 0) else
                      C_ROJO  if (j == 3 and var_p < 0)  else
                      C_VERDE if (j == 5 and var_a >= 0) else
                      C_ROJO  if (j == 5 and var_a < 0)  else
                      f.fgColor.rgb if hasattr(f, 'fgColor') else C_BLC)
        c.fill = f  # base
        if j == 3: c.fill = fill(C_VERDE if var_p >= 0 else C_ROJO)
        if j == 5: c.fill = fill(C_VERDE if var_a >= 0 else C_ROJO)
        c.font = font_cat
        if fmt: c.number_format = fmt
        c.alignment = aln('right') if j > 0 else aln('left')
    rc += 1

# Fila TOTAL categorias
tot_p  = sum(proj_cat_abr.get(c, 0) for c in CATS_ORDEN)
tot_cm = sum(cuota_monto_cat.get(c, 0) for c in CATS_ORDEN)
tot_ca = sum(cuota_abr_cat.get(c, 0) for c in CATS_ORDEN)
var_pt = (tot_cm / tot_p - 1) if tot_p > 0 else 0
var_at = (tot_p / tot_ca - 1) if tot_ca > 0 else 0
diff_t = tot_cm - tot_p
datos_total_c = [('TOTAL', None), (int(round(tot_p)), '#,##0'), (int(round(tot_cm)), '#,##0'),
                 (var_pt, '0.00%'), (int(round(tot_ca)), '#,##0'), (var_at, '0.00%'),
                 (int(round(diff_t)), '#,##0')]
for j, (val, fmt) in enumerate(datos_total_c):
    c = ws_v.cell(row=rc, column=COL_C+j, value=val)
    c.fill = fill(C_HDR); c.font = fnt(bold=True, color='FFFFFFFF', sz=9)
    if fmt: c.number_format = fmt
    c.alignment = aln('right') if j > 0 else aln('left')
rc += 2

# Nota al pie (debajo de ambas tablas, en la col de categorias)
nota_row = max(fila_total_v, rc) + 1
ws_v.cell(row=nota_row, column=COL_R,
          value='* Proyectado lineal = (real acumulado / dias lab. transcurridos) x dias lab. totales del mes')
ws_v.cell(row=nota_row, column=COL_R).font = fnt(sz=8)
ws_v.freeze_panes = 'A3'

# Guardar presentacion
wb.save(OUTPUT_CUOTAS)
print(f"   Guardado: {OUTPUT_CUOTAS}")

# ════════════════════════════════════════════════════════════════════════
# GENERAR LIBRO BBDD
# ════════════════════════════════════════════════════════════════════════
print("\n[10] Generando BBDD...")
PERIODO_BBDD = str(_periodo_cuota)

wb_bbdd = Workbook()
ws_bbdd = wb_bbdd.active
ws_bbdd.title = PERIODO_BBDD

hdrs_bbdd = ['periodo','ruta','tipo','linea','producto','soles','kilos','cobertura','temporal']
for j, h in enumerate(hdrs_bbdd):
    c = ws_bbdd.cell(row=1, column=1+j, value=h)
    c.font = fnt(bold=True, color='FFFFFFFF', sz=10)
    c.fill = fill(C_HDR)
    c.alignment = aln()

bbdd_rows = []

linea_por_ruta = {ruta: 0 for ruta in rutas_oficiales}
linea_rows = []
for cat in CATS_ORDEN:
    for ruta in rutas_oficiales:
        soles_val = int(round(float(dist_monto.loc[ruta, cat])))
        kilos_val = int(round(float(dist_vol_linea.loc[ruta, 'TOTAL']))) \
                    if (cat == 'PROCESADOS' and ruta_ffvv_map.get(ruta) == 'F8'
                        and ruta in dist_vol_linea.index) \
                    else 0
        cob_val = int(round(float(dist_cob.loc[ruta, cat])))
        linea_por_ruta[ruta] += soles_val
        linea_rows.append([PERIODO_BBDD, ruta, 'LINEA', cat, cat,
                           soles_val, kilos_val, cob_val, 0])

for ruta in rutas_oficiales:
    bbdd_rows.append([PERIODO_BBDD, ruta, 'GENERAL', 'GENERAL', 'GENERAL',
                      linea_por_ruta[ruta], 0, 0, 0])
bbdd_rows.extend(linea_rows)

for i, row_data in enumerate(bbdd_rows, start=2):
    f = fill(C_GRIS if i % 2 == 0 else C_BLC)
    for j, val in enumerate(row_data):
        c = ws_bbdd.cell(row=i, column=1+j, value=val)
        c.fill = f
        c.font = fnt(sz=9)
        if j >= 5:
            c.number_format = '#,##0'
            c.alignment = aln('right')

anchos = [10, 8, 10, 22, 22, 12, 12, 12, 10]
for j, w in enumerate(anchos):
    ws_bbdd.column_dimensions[get_column_letter(1+j)].width = w

wb_bbdd.save(OUTPUT_BBDD)
print(f"   Guardado: {OUTPUT_BBDD}")

# ════════════════════════════════════════════════════════════════════════
# VERIFICACION
# ════════════════════════════════════════════════════════════════════════
print(f"\n{'='*60}")
print("VERIFICACION FINAL")
print(f"{'='*60}")
print(f"  Cuota {_label_cuota}:  S/ {dist_monto['GENERAL'].sum():,.0f}")
print(f"  Objetivo (proj+5%):  S/ {CUOTA_TOTAL_OBJ:,.0f}")
print(f"  KB en SSFF:          S/ {dist_monto.loc[rutas_kb, SSFF_CATS].sum().sum():.0f}  (debe ser 0)")
print(f"  Rutas VOL_SSFF:      {len(rutas_f8)} (F8 + M0)")
sum_gen = sum(r[5] for r in bbdd_rows if r[2] == 'GENERAL')
sum_lin = sum(r[5] for r in bbdd_rows if r[2] == 'LINEA')
print(f"  BBDD GENERAL==LINEA: {abs(sum_gen - sum_lin) < 1}  ({sum_gen:,} vs {sum_lin:,})")
print(f"  Rutas con floor:     {len(rutas_con_floor)}")
print(f"  Hojas Excel:         {[ws.title for ws in wb.worksheets]}")
print(f"\nDone.")
