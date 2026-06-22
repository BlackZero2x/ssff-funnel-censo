#!/usr/bin/env python3
"""
generar_corte_html.py

Generador de cortes de ventas en HTML (versión experimental).
Usa los MISMOS datos que generar_corte_ventas.py de producción.

Uso:
    python generar_corte_html.py --hora 10
    python generar_corte_html.py --hora 14 --fecha 2026-06-20
"""

import argparse
import datetime
import os
import sys
from pathlib import Path

import pandas as pd
import numpy as np
from jinja2 import Template, Environment, FileSystemLoader

# Importar de producción (mismo código)
sys.path.insert(0, str(Path(__file__).parent.parent / "Reportes_ssff_wsp"))
from generar_corte_ventas import (
    conectar_sql, cargar_preventa_dia, cargar_dia_con_cache,
    cargar_mapa_zonas, calcular_cuota_supervisor, calcular_cuota_zonal,
    enriquecer_con_vendedor, agregar_hora_lbl, filtrar_corte, indicadores,
    agregar_general, agregar_por_columna, agregar_por_sup_vendedor,
    DIAS_SEMANA, ZONAS_ORDEN, TABLAS_PATH
)
from generar_cuota_dia import calcular_cuota_dia_vendedor

# Configuración
BASE_DIR = Path(__file__).parent
TEMPLATE_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output" / "cortes_html"
LOG_DIR = BASE_DIR / "logs"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)

# Colores (mismo que producción)
COLORES = {
    'tit_general': '#5B9BD5',
    'tit_zonal': '#1F3864',
    'tit_supervisor': '#006C50',
    'hdr_general': '#DEEBF7',
    'hdr_zonal': '#DAE3F3',
    'hdr_supervisor': '#CCFFCC',
    'data_gris': '#D9D9D9',
    'blanco': '#FFFFFF',
}

def generar_html_tabla(titulo, color_titulo, datos_filas, columnas, fecha_info):
    """Genera HTML de una tabla con estilos modernos."""
    html = f"""
    <div class="tabla-contenedor">
        <div class="tabla-titulo" style="background-color: {color_titulo};">
            <h2>{titulo}</h2>
            <p>{fecha_info}</p>
        </div>

        <table class="tabla-datos">
            <thead style="background-color: {COLORES['hdr_zonal']};">
                <tr>
                    {''.join(f'<th>{col}</th>' for col in columnas)}
                </tr>
            </thead>
            <tbody>
    """

    es_par = False
    for fila in datos_filas:
        color_fila = COLORES['data_gris'] if es_par else COLORES['blanco']
        html += f'<tr style="background-color: {color_fila};">'
        for valor in fila:
            # Formatear números
            if isinstance(valor, (int, float)):
                if 'Avance' in str(columnas) or '%' in str(columnas):
                    valor_str = f"{valor:.1%}" if isinstance(valor, float) else str(valor)
                elif 'Soles' in str(columnas) or 'S/' in str(columnas):
                    valor_str = f"S/ {valor:,.2f}" if isinstance(valor, float) else str(valor)
                else:
                    valor_str = f"{valor:,.0f}" if isinstance(valor, float) else str(valor)
            else:
                valor_str = str(valor) if valor else "-"

            html += f'<td>{valor_str}</td>'
        html += '</tr>'
        es_par = not es_par

    html += """
            </tbody>
        </table>
    </div>
    """
    return html


def generar_html_completo(titulo_documento, html_general, html_zonal, html_supervisor):
    """Envuelve todo en HTML con CSS."""
    return f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{titulo_documento}</title>
        <style>
            * {{
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }}

            body {{
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                background-color: #f5f5f5;
                padding: 20px;
                color: #333;
            }}

            .contenedor {{
                max-width: 1400px;
                margin: 0 auto;
                background-color: white;
                border-radius: 8px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.1);
                padding: 30px;
            }}

            .encabezado {{
                text-align: center;
                margin-bottom: 30px;
                border-bottom: 3px solid #1F3864;
                padding-bottom: 20px;
            }}

            .encabezado h1 {{
                color: #1F3864;
                font-size: 28px;
                margin-bottom: 5px;
            }}

            .encabezado p {{
                color: #666;
                font-size: 14px;
            }}

            .tablas {{
                display: grid;
                grid-template-columns: 1fr 1fr 1fr;
                gap: 20px;
                margin-top: 20px;
            }}

            @media (max-width: 1200px) {{
                .tablas {{
                    grid-template-columns: 1fr 1fr;
                }}
            }}

            @media (max-width: 768px) {{
                .tablas {{
                    grid-template-columns: 1fr;
                }}
            }}

            .tabla-contenedor {{
                background: white;
                border: 1px solid #ddd;
                border-radius: 4px;
                overflow: hidden;
                box-shadow: 0 1px 4px rgba(0,0,0,0.05);
            }}

            .tabla-titulo {{
                color: white;
                padding: 15px;
                text-align: center;
            }}

            .tabla-titulo h2 {{
                font-size: 16px;
                font-weight: 600;
                margin: 0;
            }}

            .tabla-titulo p {{
                font-size: 12px;
                margin: 5px 0 0 0;
                opacity: 0.9;
            }}

            table.tabla-datos {{
                width: 100%;
                border-collapse: collapse;
                font-size: 13px;
            }}

            table.tabla-datos thead th {{
                padding: 10px 8px;
                text-align: center;
                font-weight: 600;
                color: white;
                border-bottom: 2px solid #333;
            }}

            table.tabla-datos tbody td {{
                padding: 8px;
                text-align: right;
                border-bottom: 1px solid #eee;
            }}

            table.tabla-datos tbody td:first-child {{
                text-align: left;
                font-weight: 500;
            }}

            table.tabla-datos tbody tr:hover {{
                background-color: #f9f9f9 !important;
            }}

            .pie {{
                text-align: center;
                margin-top: 30px;
                padding-top: 20px;
                border-top: 1px solid #ddd;
                font-size: 12px;
                color: #999;
            }}

            @media print {{
                body {{
                    background: white;
                    padding: 0;
                }}
                .contenedor {{
                    box-shadow: none;
                    padding: 10px;
                }}
                .tablas {{
                    gap: 10px;
                }}
            }}
        </style>
    </head>
    <body>
        <div class="contenedor">
            <div class="encabezado">
                <h1>{titulo_documento}</h1>
                <p>Generado: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</p>
            </div>

            <div class="tablas">
                {html_general}
                {html_zonal}
                {html_supervisor}
            </div>

            <div class="pie">
                <p>Corte experimental de pruebas HTML — No es documento oficial</p>
            </div>
        </div>
    </body>
    </html>
    """


def main():
    parser = argparse.ArgumentParser(description='Genera corte de ventas en HTML')
    parser.add_argument('--hora', type=int, required=True, help='Hora (8-18)')
    parser.add_argument('--fecha', type=str, default=None, help='Fecha YYYY-MM-DD')
    args = parser.parse_args()

    # Resolver fecha
    if args.fecha:
        hoy = datetime.datetime.strptime(args.fecha, '%Y-%m-%d').date()
    else:
        hoy = datetime.date.today()

    hora = args.hora
    hora_lbl = {8: '8AM', 9: '9AM', 10: '10AM', 11: '11AM', 12: '12PM',
                13: '1PM', 14: '2PM', 15: '3PM', 16: '4PM', 17: '5PM', 18: '6PM'}.get(hora, f'{hora}H')
    nombre_dia = DIAS_SEMANA[hoy.weekday() + 1]

    print(f"\n{'='*70}")
    print(f"  GENERAR CORTE HTML — {nombre_dia} {hoy.strftime('%d/%m/%Y')} — {hora_lbl}")
    print(f"{'='*70}")

    # Cargar datos (MISMO PROCESO QUE PRODUCCIÓN)
    print(f"\n[1] Cargando datos...")
    conn = conectar_sql()
    try:
        d7 = hoy - datetime.timedelta(days=7)
        d14 = hoy - datetime.timedelta(days=14)

        df_d = cargar_preventa_dia(conn, hoy)
        df_d7 = cargar_dia_con_cache(conn, d7, 'd7')
        df_d14 = cargar_dia_con_cache(conn, d14, 'd14')

        print(f"   [D] {hoy} — {len(df_d):,} pedidos")
        print(f"   [D-7] {d7} — {len(df_d7):,} pedidos")
        print(f"   [D-14] {d14} — {len(df_d14):,} pedidos")
    finally:
        conn.close()

    # Procesar datos
    print(f"\n[2] Procesando datos...")
    mapa_zona = cargar_mapa_zonas()
    dfs = {}
    for nm, df in [('d', df_d), ('d7', df_d7), ('d14', df_d14)]:
        df = agregar_hora_lbl(df)
        df = filtrar_corte(df, hora)
        df['ZONA2'] = df['ruta'].map(mapa_zona).fillna('(sin zona)')
        dfs[nm] = df

    # Agregar indicadores
    print(f"\n[3] Calculando indicadores...")
    g_general = agregar_general(dfs['d14'], dfs['d7'], dfs['d'])
    g_zonal = agregar_por_columna(dfs['d14'], dfs['d7'], dfs['d'], 'ZONA2', ZONAS_ORDEN)

    df_rutas = pd.read_excel(TABLAS_PATH, sheet_name='RUTA_ACTUAL', dtype={'RUTA': str})
    sup_maestro = set(df_rutas['SUPERVISOR'].dropna().unique())
    sup_con_ventas = (set(dfs['d14']['supervisor'].dropna()) |
                     set(dfs['d7']['supervisor'].dropna()) |
                     set(dfs['d']['supervisor'].dropna()))
    sup_orden = sorted(sup_maestro | sup_con_ventas)
    g_super = agregar_por_columna(dfs['d14'], dfs['d7'], dfs['d'], 'supervisor', sup_orden)

    # Cuotas
    dfs_v = {}
    for nm, df in [('d', dfs['d']), ('d7', dfs['d7']), ('d14', dfs['d14'])]:
        dfs_v[nm] = enriquecer_con_vendedor(df)

    cuota_vend = calcular_cuota_dia_vendedor()
    df_rutas_raw = pd.read_excel(TABLAS_PATH, sheet_name='RUTA_ACTUAL', dtype={'RUTA': str})
    ruta_por_vendedor = df_rutas_raw.groupby('VENDEDOR')['RUTA'].apply(list).to_dict()
    cuota_sup = calcular_cuota_supervisor(cuota_vend, ruta_por_vendedor)
    cuota_zonal = calcular_cuota_zonal(cuota_vend, ruta_por_vendedor)

    # Generar HTMLs
    print(f"\n[4] Generando HTML...")

    fechas = {'d': hoy.strftime('%d/%m'), 'd7': d7.strftime('%d/%m'), 'd14': d14.strftime('%d/%m')}

    # TABLA GENERAL
    filas_general = [
        ('Pedidos', g_general['Pedidos']['d14'], g_general['Pedidos']['d7'], g_general['Pedidos']['d']),
        ('Soles', g_general['Soles']['d14'], g_general['Soles']['d7'], g_general['Soles']['d']),
        ('Ticket Prom.', g_general['Ticket']['d14'], g_general['Ticket']['d7'], g_general['Ticket']['d']),
    ]
    cols_general = ['Indicador', f'D-14 ({fechas["d14"]})', f'D-7 ({fechas["d7"]})', f'D ({fechas["d"]})']
    html_general = generar_html_tabla('📊 GENERAL', COLORES['tit_general'], filas_general, cols_general, '')

    # TABLA ZONAL
    filas_zonal = []
    for zona in ZONAS_ORDEN:
        d = g_zonal.get(zona, {'d14': (0, 0), 'd7': (0, 0), 'd': (0, 0)})
        cuota = cuota_zonal.get(zona, 0)
        pedidos_d = d['d'][0]
        soles_d = d['d'][1]
        pct_avance = (soles_d / cuota * 100) if cuota > 0 else 0
        filas_zonal.append((zona, d['d14'][1], d['d7'][1], soles_d, cuota, pct_avance))

    cols_zonal = ['ZONAL', f'D-14 ({fechas["d14"]})', f'D-7 ({fechas["d7"]})', f'D ({fechas["d"]})', 'Cuota Día', '%Avance']
    html_zonal = generar_html_tabla('📍 ZONAL', COLORES['tit_zonal'], filas_zonal, cols_zonal, '')

    # TABLA SUPERVISOR
    filas_super = []
    for sup in sup_orden:
        d = g_super.get(sup, {'d14': (0, 0), 'd7': (0, 0), 'd': (0, 0)})
        cuota = cuota_sup.get(sup, 0)
        pedidos_d = d['d'][0]
        soles_d = d['d'][1]
        pct_avance = (soles_d / cuota * 100) if cuota > 0 else 0
        filas_super.append((sup, d['d14'][1], d['d7'][1], soles_d, cuota, pct_avance))

    cols_super = ['Supervisor', f'D-14 ({fechas["d14"]})', f'D-7 ({fechas["d7"]})', f'D ({fechas["d"]})', 'Cuota Día', '%Avance']
    html_super = generar_html_tabla('👤 SUPERVISOR', COLORES['tit_supervisor'], filas_super, cols_super, '')

    # HTML completo
    titulo = f"CORTE DE VENTAS — {nombre_dia} {hoy.strftime('%d/%m/%Y')} — {hora_lbl}"
    html_completo = generar_html_completo(titulo, html_general, html_zonal, html_super)

    # Guardar
    out_file = OUTPUT_DIR / f"corte_{hoy.strftime('%Y%m%d')}_{hora_lbl.lower().replace(' ', '')}.html"
    with open(out_file, 'w', encoding='utf-8') as f:
        f.write(html_completo)

    print(f"\n   Guardado: {out_file}")
    print(f"\n✅ Abre en navegador:")
    print(f"   start {out_file}")


if __name__ == "__main__":
    main()
