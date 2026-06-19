"""
Análisis: Cuota Operativa vs Cuota Concurso (EMBUTIDOS + CONGELADOS) — Mayo 2026

Fuentes:
  - Cuota operativa: CUOTAS_CONGELADOS_EMBUTIDOS_MAYO2026.xlsx  (desglose por ruta)
  - Cuota concurso:  cuotasdecliente/CUOTAS_MAYO2026_SSFF.xlsx   (concurso Auren)
  - Cuota soles:     CUOTAS_BBDD_202605.xlsx                     (soles operativos)
  - Cuota otros:     cuotasdecliente/CUOTAS_OTROSCLIENTES_MAYO2026.xlsx

Salida: ANALISIS_CUOTAS_MAYO2026.xlsx
"""

import os
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

DIR = r'C:\proyectos\SSFF'

# ── Colores ────────────────────────────────────────────────────────────────────
C_AZL_OSC  = PatternFill('solid', fgColor='1F3864')  # azul oscuro — títulos
C_AZL_MED  = PatternFill('solid', fgColor='2F5496')  # azul medio — cabeceras
C_AZL_CLA  = PatternFill('solid', fgColor='D6E4F0')  # azul claro — fondo alterno
C_NARANJA  = PatternFill('solid', fgColor='F4B942')  # naranja — totales
C_VERDE    = PatternFill('solid', fgColor='C6EFCE')  # verde — SI conservar
C_ROJO     = PatternFill('solid', fgColor='FFC7CE')  # rojo — NO conservar
C_AMARILLO = PatternFill('solid', fgColor='FFEB9C')  # amarillo — diferencia notable
C_GRIS     = PatternFill('solid', fgColor='F2F2F2')  # gris — alterno

THIN = Side(style='thin', color='CCCCCC')
BRD  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def fill(clr):  return clr
def fnt(bold=False, color='000000', size=10, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic, name='Calibri')
def aln(h='left', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def cset(ws, row, col, val, fill=None, font=None, align=None, fmt=None, border=None):
    c = ws.cell(row=row, column=col, value=val)
    if fill:   c.fill   = fill
    if font:   c.font   = font
    if align:  c.alignment = align
    if fmt:    c.number_format = fmt
    if border: c.border = border
    return c

# ── [1] CARGAR DATOS ──────────────────────────────────────────────────────────

# Cuota operativa desglosada (EMBUTIDOS + CONGELADOS)
df_op_raw = pd.read_excel(
    os.path.join(DIR, 'CUOTAS_CONGELADOS_EMBUTIDOS_MAYO2026.xlsx'), header=1
)
df_op_raw.columns = ['idx','ffvv','ruta','emb_vol_op','emb_cob_op',
                     'cong_vol_op','cong_cob_op','total_vol_op','total_cob_op']
df_op = df_op_raw[df_op_raw['ruta'].astype(str).str.match(r'^[FM]\d+$')].copy()
df_op['emb_vol_op']  = pd.to_numeric(df_op['emb_vol_op'],  errors='coerce').fillna(0)
df_op['emb_cob_op']  = pd.to_numeric(df_op['emb_cob_op'],  errors='coerce').fillna(0)
df_op['cong_vol_op'] = pd.to_numeric(df_op['cong_vol_op'], errors='coerce').fillna(0)
df_op['cong_cob_op'] = pd.to_numeric(df_op['cong_cob_op'], errors='coerce').fillna(0)
df_op = df_op[['ffvv','ruta','emb_vol_op','emb_cob_op','cong_vol_op','cong_cob_op']].reset_index(drop=True)

# Cuota concurso Auren (EMBUTIDOS + CONGELADOS + cobertura total)
df_con_raw = pd.read_excel(
    os.path.join(DIR, 'cuotasdecliente', 'CUOTAS_MAYO2026_SSFF.xlsx'), header=8
)
df_con_raw.columns = ['ruta','emb_vol_con','cong_vol_con','cartera_ef_abr',
                      'x7','cuota_cob_con','ratio','cob_mes_abr',
                      'cuota_cob_empacado','cong_granel_vol']
df_con = df_con_raw[df_con_raw['ruta'].astype(str).str.match(r'^[FM]\d+$')].copy()
for col in ['emb_vol_con','cong_vol_con','cuota_cob_con','cartera_ef_abr',
            'cob_mes_abr','cuota_cob_empacado','cong_granel_vol']:
    df_con[col] = pd.to_numeric(df_con[col], errors='coerce').fillna(0)
df_con = df_con[['ruta','emb_vol_con','cong_vol_con','cuota_cob_con',
                 'cartera_ef_abr','cob_mes_abr','cuota_cob_empacado','cong_granel_vol']].reset_index(drop=True)

# Cuota operativa en soles (PROCESADOS total)
df_bbdd = pd.read_excel(os.path.join(DIR, 'CUOTAS_BBDD_202605.xlsx'))
df_soles = df_bbdd[(df_bbdd['tipo']=='LINEA') & (df_bbdd['linea']=='PROCESADOS')][
    ['ruta','soles','kilos','cobertura']
].copy()
df_soles.columns = ['ruta','soles_op','kilos_op_total','cob_op_total']

# ── [2] MERGE ─────────────────────────────────────────────────────────────────
df = df_op.merge(df_con, on='ruta', how='outer').merge(df_soles, on='ruta', how='left')
df['ffvv'] = df['ffvv'].fillna(df['ruta'].str[:2].str.upper())
df = df.sort_values(['ffvv','ruta']).reset_index(drop=True)

# Métricas de comparación
df['emb_vol_dif']   = df['emb_vol_con']   - df['emb_vol_op']
df['emb_vol_var']   = np.where(df['emb_vol_op'] > 0,
                               df['emb_vol_dif'] / df['emb_vol_op'], np.nan)
df['cong_vol_dif']  = df['cong_vol_con']  - df['cong_vol_op']
df['cong_vol_var']  = np.where(df['cong_vol_op'] > 0,
                               df['cong_vol_dif'] / df['cong_vol_op'], np.nan)
df['total_vol_op']  = df['emb_vol_op']  + df['cong_vol_op']
df['total_vol_con'] = df['emb_vol_con'] + df['cong_vol_con']
df['total_vol_dif'] = df['total_vol_con'] - df['total_vol_op']
df['total_vol_var'] = np.where(df['total_vol_op'] > 0,
                               df['total_vol_dif'] / df['total_vol_op'], np.nan)

# ── Lógica de recomendación ──────────────────────────────────────────────────
# Criterios para CONSERVAR cuota operativa:
#  1. La cuota operativa (total vol) es >= concurso  → más exigente, conservar
#  2. La diferencia es <= 5% hacia abajo  → diferencia pequeña, indistinta
# Criterios para NO conservar (usar concurso):
#  1. Concurso es > 10% mayor que operativa → concurso más exigente, reemplazar
#  2. Concurso tiene mayor cobertura exigida y operativa tiene ratio bajo

def recomendar(row):
    var = row['total_vol_var']
    if pd.isna(var):
        return 'REVISAR', 'Sin datos suficientes para comparar'
    if var <= 0.05:  # concurso hasta 5% mayor → conservar operativa
        if var <= 0:
            return 'SI', f'Cuota operativa ≥ concurso (+{abs(var):.1%} más exigente)'
        else:
            return 'SI', f'Diferencia mínima ({var:.1%}); cuota operativa es suficientemente exigente'
    elif var <= 0.15:  # concurso 5-15% mayor → zona amarilla
        return 'EVALUAR', f'Concurso exige {var:.1%} más volumen; revisar capacidad de ruta'
    else:  # concurso > 15% mayor
        return 'NO', f'Concurso exige {var:.1%} más volumen; cuota operativa subestima la ruta'

rec = df.apply(recomendar, axis=1, result_type='expand')
df['recomendacion'] = rec[0]
df['razon']         = rec[1]

# ── [3] GENERAR EXCEL ─────────────────────────────────────────────────────────
wb = Workbook()

# ═══════════════════════════════════════════════════════════════════════════════
# HOJA 1 — VERSUS DETALLE
# ═══════════════════════════════════════════════════════════════════════════════
ws = wb.active
ws.title = 'VERSUS DETALLE'

# Anchos de columna
WCOLS = {
    1: 6,   # FFVV
    2: 7,   # RUTA
    3: 14,  # EMB VOL OP
    4: 14,  # EMB VOL CON
    5: 11,  # EMB DIF
    6: 9,   # EMB VAR%
    7: 14,  # CONG VOL OP
    8: 14,  # CONG VOL CON
    9: 11,  # CONG DIF
    10: 9,  # CONG VAR%
    11: 14, # TOTAL VOL OP
    12: 14, # TOTAL VOL CON
    13: 11, # TOTAL DIF
    14: 9,  # TOTAL VAR%
    15: 12, # SOLES OP
    16: 12, # COB OP TOTAL
    17: 12, # COB CON (concurso)
    18: 12, # RECOMENDACION
    19: 55, # RAZÓN
}
for col, w in WCOLS.items():
    ws.column_dimensions[get_column_letter(col)].width = w

# Título principal
ws.merge_cells('A1:S1')
cset(ws, 1, 1, 'ANÁLISIS CUOTA OPERATIVA vs CUOTA CONCURSO — PROCESADOS (EMBUTIDOS + CONGELADOS) — MAYO 2026',
     fill=C_AZL_OSC, font=fnt(bold=True, color='FFFFFF', size=12),
     align=aln('center'))
ws.row_dimensions[1].height = 22

# Subtítulo de secciones (fila 2)
ws.merge_cells('C2:F2')
cset(ws, 2, 3, 'EMBUTIDOS (kg)', fill=C_AZL_MED,
     font=fnt(bold=True, color='FFFFFF'), align=aln('center'))
ws.merge_cells('G2:J2')
cset(ws, 2, 7, 'CONGELADOS (kg)', fill=C_AZL_MED,
     font=fnt(bold=True, color='FFFFFF'), align=aln('center'))
ws.merge_cells('K2:N2')
cset(ws, 2, 11, 'TOTAL PROCESADOS (kg)', fill=C_AZL_MED,
     font=fnt(bold=True, color='FFFFFF'), align=aln('center'))
ws.merge_cells('O2:P2')
cset(ws, 2, 15, 'CUOTA OPERATIVA', fill=C_AZL_MED,
     font=fnt(bold=True, color='FFFFFF'), align=aln('center'))
ws.merge_cells('Q2:Q2')
cset(ws, 2, 17, 'COB. CONCURSO', fill=C_AZL_MED,
     font=fnt(bold=True, color='FFFFFF'), align=aln('center'))

# Cabeceras (fila 3)
hdrs = [
    'FFVV', 'RUTA',
    'VOL OP', 'VOL CONCURSO', 'DIFERENCIA', 'VAR%',
    'VOL OP', 'VOL CONCURSO', 'DIFERENCIA', 'VAR%',
    'VOL OP', 'VOL CONCURSO', 'DIFERENCIA', 'VAR%',
    'SOLES OP', 'COB. TOTAL OP',
    'COB. CONCURSO',
    'RECOMENDACIÓN', 'RAZÓN',
]
ws.row_dimensions[3].height = 32
for i, h in enumerate(hdrs, 1):
    cset(ws, 3, i, h, fill=C_AZL_MED,
         font=fnt(bold=True, color='FFFFFF', size=9),
         align=aln('center', wrap=True), border=BRD)

# Datos
ws.freeze_panes = 'A4'
for r, row in df.iterrows():
    excel_row = r + 4
    alt = C_AZL_CLA if r % 2 == 0 else None

    # Color fila según recomendación
    if row['recomendacion'] == 'SI':
        rec_fill = C_VERDE
    elif row['recomendacion'] == 'NO':
        rec_fill = C_ROJO
    elif row['recomendacion'] == 'EVALUAR':
        rec_fill = C_AMARILLO
    else:
        rec_fill = C_GRIS

    datos = [
        (row['ffvv'],         None),
        (row['ruta'],         None),
        (row['emb_vol_op'],   '#,##0'),
        (row['emb_vol_con'],  '#,##0'),
        (row['emb_vol_dif'],  '+#,##0;-#,##0;0'),
        (row['emb_vol_var'],  '0.0%;[Red]-0.0%'),
        (row['cong_vol_op'],  '#,##0'),
        (row['cong_vol_con'], '#,##0'),
        (row['cong_vol_dif'], '+#,##0;-#,##0;0'),
        (row['cong_vol_var'], '0.0%;[Red]-0.0%'),
        (row['total_vol_op'], '#,##0'),
        (row['total_vol_con'],'#,##0'),
        (row['total_vol_dif'],'+#,##0;-#,##0;0'),
        (row['total_vol_var'],'0.0%;[Red]-0.0%'),
        (row['soles_op'],     'S/ #,##0'),
        (row['cob_op_total'], '#,##0'),
        (row['cuota_cob_con'],'#,##0'),
        (row['recomendacion'],None),
        (row['razon'],        None),
    ]

    for c, (val, fmt) in enumerate(datos, 1):
        bg = alt
        if c == 18:  # col recomendación
            bg = rec_fill
        v = None if (isinstance(val, float) and np.isnan(val)) else val
        cset(ws, excel_row, c, v,
             fill=bg, font=fnt(size=9),
             align=aln('center' if c <= 17 else 'left'),
             fmt=fmt, border=BRD)

# Fila de totales
tot_row = len(df) + 4
ws.merge_cells(f'A{tot_row}:B{tot_row}')
cset(ws, tot_row, 1, 'TOTAL GENERAL', fill=C_NARANJA,
     font=fnt(bold=True, size=9), align=aln('center'), border=BRD)

totales = [
    (df['emb_vol_op'].sum(),   '#,##0', 3),
    (df['emb_vol_con'].sum(),  '#,##0', 4),
    (df['emb_vol_dif'].sum(),  '+#,##0;-#,##0;0', 5),
    (df['emb_vol_var'].mean(), '0.0%', 6),
    (df['cong_vol_op'].sum(),  '#,##0', 7),
    (df['cong_vol_con'].sum(), '#,##0', 8),
    (df['cong_vol_dif'].sum(), '+#,##0;-#,##0;0', 9),
    (df['cong_vol_var'].mean(),'0.0%', 10),
    (df['total_vol_op'].sum(), '#,##0', 11),
    (df['total_vol_con'].sum(),'#,##0', 12),
    (df['total_vol_dif'].sum(),'+#,##0;-#,##0;0', 13),
    (df['total_vol_var'].mean(),'0.0%', 14),
    (df['soles_op'].sum(),     'S/ #,##0', 15),
    (df['cob_op_total'].sum(), '#,##0', 16),
    (df['cuota_cob_con'].sum(),'#,##0', 17),
]
for val, fmt, col in totales:
    v = None if (isinstance(val, float) and np.isnan(val)) else val
    cset(ws, tot_row, col, v, fill=C_NARANJA,
         font=fnt(bold=True, size=9), align=aln('center'),
         fmt=fmt, border=BRD)

# ═══════════════════════════════════════════════════════════════════════════════
# HOJA 2 — RESUMEN POR FFVV
# ═══════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet('RESUMEN FFVV')

WCOLS2 = {1:8, 2:15, 3:15, 4:11, 5:9, 6:15, 7:15, 8:11, 9:9,
          10:15, 11:15, 12:11, 13:9, 14:10, 15:10, 16:8}
for col, w in WCOLS2.items():
    ws2.column_dimensions[get_column_letter(col)].width = w

ws2.merge_cells('A1:P1')
cset(ws2, 1, 1, 'RESUMEN POR FUERZA DE VENTAS (FFVV) — MAYO 2026',
     fill=C_AZL_OSC, font=fnt(bold=True, color='FFFFFF', size=12),
     align=aln('center'))
ws2.row_dimensions[1].height = 22

ws2.merge_cells('B2:E2')
cset(ws2, 2, 2, 'EMBUTIDOS (kg)', fill=C_AZL_MED,
     font=fnt(bold=True, color='FFFFFF'), align=aln('center'))
ws2.merge_cells('F2:I2')
cset(ws2, 2, 6, 'CONGELADOS (kg)', fill=C_AZL_MED,
     font=fnt(bold=True, color='FFFFFF'), align=aln('center'))
ws2.merge_cells('J2:M2')
cset(ws2, 2, 10, 'TOTAL PROCESADOS (kg)', fill=C_AZL_MED,
     font=fnt(bold=True, color='FFFFFF'), align=aln('center'))

hdrs2 = ['FFVV',
         'VOL OP','VOL CONCURSO','DIF','VAR%',
         'VOL OP','VOL CONCURSO','DIF','VAR%',
         'VOL OP','VOL CONCURSO','DIF','VAR%',
         'SOLES OP','COB OP','COB CONCURSO']
ws2.row_dimensions[3].height = 28
for i, h in enumerate(hdrs2, 1):
    cset(ws2, 3, i, h, fill=C_AZL_MED,
         font=fnt(bold=True, color='FFFFFF', size=9),
         align=aln('center', wrap=True), border=BRD)

grp = df.groupby('ffvv').agg(
    emb_vol_op=('emb_vol_op','sum'),
    emb_vol_con=('emb_vol_con','sum'),
    cong_vol_op=('cong_vol_op','sum'),
    cong_vol_con=('cong_vol_con','sum'),
    soles_op=('soles_op','sum'),
    cob_op_total=('cob_op_total','sum'),
    cuota_cob_con=('cuota_cob_con','sum'),
).reset_index()
grp['total_vol_op']  = grp['emb_vol_op']  + grp['cong_vol_op']
grp['total_vol_con'] = grp['emb_vol_con'] + grp['cong_vol_con']
grp['emb_dif']   = grp['emb_vol_con']   - grp['emb_vol_op']
grp['emb_var']   = grp['emb_dif']   / grp['emb_vol_op']
grp['cong_dif']  = grp['cong_vol_con']  - grp['cong_vol_op']
grp['cong_var']  = grp['cong_dif']  / grp['cong_vol_op']
grp['total_dif'] = grp['total_vol_con'] - grp['total_vol_op']
grp['total_var'] = grp['total_dif'] / grp['total_vol_op']

ws2.freeze_panes = 'A4'
for r, row in grp.iterrows():
    excel_row = r + 4
    alt = C_AZL_CLA if r % 2 == 0 else None
    datos2 = [
        (row['ffvv'],          None),
        (row['emb_vol_op'],    '#,##0'),
        (row['emb_vol_con'],   '#,##0'),
        (row['emb_dif'],       '+#,##0;-#,##0;0'),
        (row['emb_var'],       '0.0%;[Red]-0.0%'),
        (row['cong_vol_op'],   '#,##0'),
        (row['cong_vol_con'],  '#,##0'),
        (row['cong_dif'],      '+#,##0;-#,##0;0'),
        (row['cong_var'],      '0.0%;[Red]-0.0%'),
        (row['total_vol_op'],  '#,##0'),
        (row['total_vol_con'], '#,##0'),
        (row['total_dif'],     '+#,##0;-#,##0;0'),
        (row['total_var'],     '0.0%;[Red]-0.0%'),
        (row['soles_op'],      'S/ #,##0'),
        (row['cob_op_total'],  '#,##0'),
        (row['cuota_cob_con'], '#,##0'),
    ]
    for c, (val, fmt) in enumerate(datos2, 1):
        v = None if (isinstance(val, float) and np.isnan(val)) else val
        cset(ws2, excel_row, c, v,
             fill=alt, font=fnt(size=9), align=aln('center'),
             fmt=fmt, border=BRD)

# Total FFVV
tot2 = len(grp) + 4
ws2.merge_cells(f'A{tot2}:A{tot2}')
cset(ws2, tot2, 1, 'TOTAL', fill=C_NARANJA,
     font=fnt(bold=True, size=9), align=aln('center'), border=BRD)
totales2 = [
    (grp['emb_vol_op'].sum(),    '#,##0', 2),
    (grp['emb_vol_con'].sum(),   '#,##0', 3),
    (grp['emb_dif'].sum(),       '+#,##0;-#,##0;0', 4),
    (grp['emb_var'].mean(),      '0.0%', 5),
    (grp['cong_vol_op'].sum(),   '#,##0', 6),
    (grp['cong_vol_con'].sum(),  '#,##0', 7),
    (grp['cong_dif'].sum(),      '+#,##0;-#,##0;0', 8),
    (grp['cong_var'].mean(),     '0.0%', 9),
    (grp['total_vol_op'].sum(),  '#,##0', 10),
    (grp['total_vol_con'].sum(), '#,##0', 11),
    (grp['total_dif'].sum(),     '+#,##0;-#,##0;0', 12),
    (grp['total_var'].mean(),    '0.0%', 13),
    (grp['soles_op'].sum(),      'S/ #,##0', 14),
    (grp['cob_op_total'].sum(),  '#,##0', 15),
    (grp['cuota_cob_con'].sum(), '#,##0', 16),
]
for val, fmt, col in totales2:
    v = None if (isinstance(val, float) and np.isnan(val)) else val
    cset(ws2, tot2, col, v, fill=C_NARANJA,
         font=fnt(bold=True, size=9), align=aln('center'),
         fmt=fmt, border=BRD)

# ═══════════════════════════════════════════════════════════════════════════════
# HOJA 3 — OTROS CLIENTES
# ═══════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet('OTROS CLIENTES')
ws3.column_dimensions['A'].width = 30
ws3.column_dimensions['B'].width = 20
ws3.column_dimensions['C'].width = 45

ws3.merge_cells('A1:C1')
cset(ws3, 1, 1, 'CUOTAS OTROS CLIENTES — MAYO 2026',
     fill=C_AZL_OSC, font=fnt(bold=True, color='FFFFFF', size=12),
     align=aln('center'))
ws3.row_dimensions[1].height = 22

clientes = {
    'ANDINA': [
        ('CREMAS Y CONSERVAS', 2933),
        ('HORNOS', 34774),
        ('JUGOS Y NECTARES', 12083),
        ('LECHES', 267948),
        ('YOGURTS', None),
        ('DANLAC', 229493),
        ('YOLEIT', 65231),
        ('DERIVADOS LACTEOS', 4103),
        ('TOTAL', 616565),
    ],
    'COLGATE': [
        ('TOTAL GLOBAL', 360000),
    ],
    'RINTI': [
        ('SOLES TOTAL RINTI', 730000),
        ('COBERTURAS SACOS THOR', 230),
        ('SOLES HÚMEDOS (PATES Y POUCH)', 145000),
    ],
}

row3 = 2
for cliente, lineas in clientes.items():
    cset(ws3, row3, 1, cliente, fill=C_AZL_MED,
         font=fnt(bold=True, color='FFFFFF'), align=aln('center'), border=BRD)
    cset(ws3, row3, 2, 'CUOTA (S/)', fill=C_AZL_MED,
         font=fnt(bold=True, color='FFFFFF'), align=aln('center'), border=BRD)
    cset(ws3, row3, 3, 'NOTA', fill=C_AZL_MED,
         font=fnt(bold=True, color='FFFFFF'), align=aln('center'), border=BRD)
    row3 += 1
    for i, (linea, valor) in enumerate(lineas):
        alt = C_AZL_CLA if i % 2 == 0 else None
        es_total = linea.startswith('TOTAL')
        bg = C_NARANJA if es_total else alt
        fn = fnt(bold=es_total, size=9)
        nota = 'Sin cuota asignada' if valor is None else ''
        if linea == 'COBERTURAS SACOS THOR':
            nota = 'Unidad: clientes únicos (no soles)'
        cset(ws3, row3, 1, linea, fill=bg, font=fn, align=aln(), border=BRD)
        cset(ws3, row3, 2, valor, fill=bg, font=fn, align=aln('center'),
             fmt='S/ #,##0' if valor and valor > 1000 else '#,##0', border=BRD)
        cset(ws3, row3, 3, nota, fill=bg, font=fnt(size=9, italic=True),
             align=aln(), border=BRD)
        row3 += 1
    row3 += 1  # espacio entre clientes

# ═══════════════════════════════════════════════════════════════════════════════
# HOJA 4 — CRITERIOS
# ═══════════════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet('CRITERIOS')
ws4.column_dimensions['A'].width = 18
ws4.column_dimensions['B'].width = 65

ws4.merge_cells('A1:B1')
cset(ws4, 1, 1, 'CRITERIOS DE RECOMENDACIÓN — CONSERVAR CUOTA OPERATIVA',
     fill=C_AZL_OSC, font=fnt(bold=True, color='FFFFFF', size=12),
     align=aln('center'))
ws4.row_dimensions[1].height = 22

criterios = [
    ('RESULTADO', 'CONDICIÓN', 'DESCRIPCIÓN', True, C_AZL_MED),
    ('SI', 'Cuota operativa ≥ cuota concurso (VAR% ≤ 0%)',
     'La cuota propia ya es más exigente. El concurso no aporta exigencia adicional.', False, C_VERDE),
    ('SI', 'Cuota concurso supera operativa en ≤ 5%',
     'Diferencia mínima, dentro del margen de redondeo y proyección. Mantener operativa es suficiente.', False, C_VERDE),
    ('EVALUAR', 'Cuota concurso supera operativa en 5-15%',
     'Diferencia moderada. Revisar capacidad histórica de la ruta antes de adoptar el concurso.', False, C_AMARILLO),
    ('NO', 'Cuota concurso supera operativa en > 15%',
     'El concurso exige significativamente más. La cuota operativa subestima la ruta y debe actualizarse.', False, C_ROJO),
    ('REVISAR', 'Datos insuficientes',
     'Ruta presente en una sola fuente. Requiere validación manual.', False, C_GRIS),
]

ws4.column_dimensions['A'].width = 12
ws4.column_dimensions['B'].width = 45
ws4.column_dimensions['C'].width = 70
for i, (col_a, col_b, col_c, is_hdr, bg) in enumerate(criterios, 2):
    fn = fnt(bold=is_hdr, color='FFFFFF' if is_hdr else '000000', size=9)
    ws4.row_dimensions[i].height = 22
    cset(ws4, i, 1, col_a, fill=bg, font=fn, align=aln('center'), border=BRD)
    cset(ws4, i, 2, col_b, fill=bg, font=fn, align=aln(), border=BRD)
    cset(ws4, i, 3, col_c, fill=bg, font=fnt(size=9, italic=not is_hdr),
         align=aln(wrap=True), border=BRD)

# Nota metodológica
ws4.merge_cells('A9:C9')
cset(ws4, 9, 1,
     'Nota: La cuota operativa (CUOTAS_BBDD_202605) se calculó con Tipo 1 (promedio/máximo MAR+ABR proyectado). '
     'La cuota concurso (CUOTAS_MAYO2026_SSFF) se basa en cartera efectiva de abril x factor de activación.',
     font=fnt(size=9, italic=True, color='595959'),
     align=aln(wrap=True))
ws4.row_dimensions[9].height = 35

# ── Guardar ───────────────────────────────────────────────────────────────────
out_path = os.path.join(DIR, 'ANALISIS_CUOTAS_MAYO2026.xlsx')
wb.save(out_path)
print(f'OK Guardado: {out_path}')

# Resumen en consola
print('\n=== RESUMEN ===')
print(df['recomendacion'].value_counts().to_string())
print(f'\nTotal rutas: {len(df)}')
print(f'Vol operativo PROCESADOS:  {df["total_vol_op"].sum():,.0f} kg')
print(f'Vol concurso  PROCESADOS:  {df["total_vol_con"].sum():,.0f} kg')
print(f'Diferencia:                {df["total_vol_dif"].sum():+,.0f} kg ({df["total_vol_dif"].sum()/df["total_vol_op"].sum():.1%})')
print(f'Soles operativos:          S/ {df["soles_op"].sum():,.0f}')
