"""
Genera CUOTAS_BBDD_<periodo>.xlsx a partir de cuotas_ssff_<mes><anio>.xlsx (editado manualmente).
Formato: periodo | ruta | tipo | linea | producto | soles | kilos | cobertura | temporal
- kilos: solo PROCESADOS (EMBUTIDOS + CONGELADOS desde CUOTA_VOL_SSFF)
- cobertura: todas las categorías desde CUOTA_COBERTURA
Uso: python generar_bbdd.py --periodo 202606
"""
import pandas as pd
from openpyxl import Workbook
import argparse
from datetime import date

_parser = argparse.ArgumentParser(description='Genera BBDD de cuota mensual SSFF')
_parser.add_argument('--periodo', type=str, default=None,
                     help='Periodo de cuota en formato YYYYMM (ej: 202606)')
_args = _parser.parse_args()

if _args.periodo:
    PERIODO = int(_args.periodo)
else:
    _hoy = date.today()
    _sig = _hoy.month % 12 + 1
    _anio = _hoy.year + (1 if _hoy.month == 12 else 0)
    PERIODO = int(f'{_anio}{_sig:02d}')

_MESES_ES = {1:'ene',2:'feb',3:'mar',4:'abr',5:'may',6:'jun',
             7:'jul',8:'ago',9:'sep',10:'oct',11:'nov',12:'dic'}
_mes_c  = PERIODO % 100
_anio_c = PERIODO // 100
INPUT  = f'C:/proyectos/SSFF/cuotas_ssff_{_MESES_ES[_mes_c]}{_anio_c}.xlsx'
OUTPUT = f'C:/proyectos/SSFF/CUOTAS_BBDD_{PERIODO}.xlsx'
print(f"Periodo: {PERIODO}  |  Input: {INPUT}")

CATS_VALIDAS = [
    'ACCESORIOS','ANDINA','CERDO','COLGATE','DERMODIS','DULFINA',
    'HIGIENE Y CUIDADO','HOMEPRO PERU','HUEVO','KIMBERLY','LA CORONA',
    'LA PATRONA','MEDIFARMA','PAVO','POLLO','PROCESADOS','RINTI',
    'TAMBOS PERU','VERDUM','YICHANG'
]

# ── Leer soles (CUOTAS_SOLES) ─────────────────────────────────────────
df_sol = pd.read_excel(INPUT, sheet_name='CUOTAS_SOLES', header=3)
df_sol = df_sol[df_sol['RUTA'].notna() & df_sol['FFVV'].notna()].copy()
df_sol = df_sol[~df_sol['FFVV'].astype(str).str.startswith('SUBTOTAL')]
df_sol = df_sol[~df_sol['RUTA'].astype(str).str.contains('TOTAL', na=False)]
df_sol = df_sol.set_index('RUTA')
cats_presentes = [c for c in CATS_VALIDAS if c in df_sol.columns]
for c in cats_presentes:
    df_sol[c] = pd.to_numeric(df_sol[c], errors='coerce').fillna(0)

# ── Leer kilos PROCESADOS (CUOTA_VOL_SSFF) ────────────────────────────
df_vol = pd.read_excel(INPUT, sheet_name='CUOTA_VOL_SSFF', header=2)
df_vol = df_vol[df_vol['RUTA'].notna() & df_vol['FFVV'].notna()].copy()
df_vol = df_vol[~df_vol['RUTA'].astype(str).str.contains('TOTAL|SUBTOTAL', na=False)]
df_vol = df_vol.set_index('RUTA')
for c in ['EMBUTIDOS - VOL', 'CONGELADOS VOL']:
    df_vol[c] = pd.to_numeric(df_vol[c], errors='coerce').fillna(0)
# kilos PROCESADOS = EMB + CON
kilos_proc = (df_vol['EMBUTIDOS - VOL'] + df_vol['CONGELADOS VOL']).round(0).astype(int)

# ── Leer cobertura (CUOTA_COBERTURA) ──────────────────────────────────
df_cob = pd.read_excel(INPUT, sheet_name='CUOTA_COBERTURA', header=3)
df_cob = df_cob[df_cob['RUTA'].notna() & df_cob['FFVV'].notna()].copy()
df_cob = df_cob[~df_cob['FFVV'].astype(str).str.startswith('SUBTOTAL')]
df_cob = df_cob[~df_cob['RUTA'].astype(str).str.contains('TOTAL', na=False)]
df_cob = df_cob.set_index('RUTA')
for c in cats_presentes:
    if c in df_cob.columns:
        df_cob[c] = pd.to_numeric(df_cob[c], errors='coerce').fillna(0)

print(f"Rutas soles: {len(df_sol)} | Rutas vol: {len(df_vol)} | Rutas cob: {len(df_cob)}")
print(f"Categorias: {len(cats_presentes)}")

# ── Construir filas BBDD ───────────────────────────────────────────────
filas = []

# Filas GENERAL: total soles por ruta, kilos PROCESADOS totales, cobertura GENERAL
for ruta in df_sol.index:
    total_soles = int(round(df_sol.loc[ruta, cats_presentes].sum()))
    kilos_r = int(kilos_proc.get(ruta, 0))
    cob_r = int(df_cob.loc[ruta, 'GENERAL']) if ruta in df_cob.index and 'GENERAL' in df_cob.columns else None
    filas.append({
        'periodo':   PERIODO,
        'ruta':      ruta,
        'tipo':      'GENERAL',
        'linea':     'GENERAL',
        'producto':  'GENERAL',
        'soles':     total_soles,
        'kilos':     kilos_r,
        'cobertura': cob_r,
        'temporal':  None,
    })

# Filas LINEA: soles + kilos (solo PROCESADOS) + cobertura por categoría
for cat in cats_presentes:
    for ruta in df_sol.index:
        soles_cat = int(round(df_sol.loc[ruta, cat]))
        # kilos solo para PROCESADOS y solo si la ruta está en la hoja de volumen
        if cat == 'PROCESADOS':
            kilos_cat = int(kilos_proc.get(ruta, 0))
        else:
            kilos_cat = 0
        # cobertura desde hoja CUOTA_COBERTURA
        if ruta in df_cob.index and cat in df_cob.columns:
            cob_cat = int(df_cob.loc[ruta, cat])
        else:
            cob_cat = None
        filas.append({
            'periodo':   PERIODO,
            'ruta':      ruta,
            'tipo':      'LINEA',
            'linea':     cat,
            'producto':  cat,
            'soles':     soles_cat,
            'kilos':     kilos_cat,
            'cobertura': cob_cat,
            'temporal':  None,
        })

bbdd = pd.DataFrame(filas, columns=['periodo','ruta','tipo','linea','producto','soles','kilos','cobertura','temporal'])

# ── Verificación ───────────────────────────────────────────────────────
gen = bbdd[bbdd['tipo']=='GENERAL'].groupby('ruta')['soles'].sum()
lin = bbdd[bbdd['tipo']=='LINEA'].groupby('ruta')['soles'].sum()
dif_soles = (gen - lin).abs().max()
total_soles = bbdd[bbdd['tipo']=='GENERAL']['soles'].sum()
total_kilos = bbdd[(bbdd['tipo']=='LINEA') & (bbdd['linea']=='PROCESADOS')]['kilos'].sum()
print(f"Total filas: {len(bbdd)} (GENERAL: {(bbdd['tipo']=='GENERAL').sum()} | LINEA: {(bbdd['tipo']=='LINEA').sum()})")
print(f"Total soles:  S/ {total_soles:,.0f}")
print(f"Total kilos PROCESADOS: {total_kilos:,.0f} kg")
print(f"GENERAL == LINEA (soles): {dif_soles == 0}  (diferencia max: {dif_soles})")

# ── Guardar ────────────────────────────────────────────────────────────
wb = Workbook()
ws = wb.active
ws.title = str(PERIODO)

ws.append(['periodo','ruta','tipo','linea','producto','soles','kilos','cobertura','temporal'])
for _, row in bbdd.iterrows():
    ws.append([
        row['periodo'], row['ruta'], row['tipo'], row['linea'], row['producto'],
        row['soles'], row['kilos'], row['cobertura'], row['temporal']
    ])

wb.save(OUTPUT)
print(f"Guardado: {OUTPUT}")
