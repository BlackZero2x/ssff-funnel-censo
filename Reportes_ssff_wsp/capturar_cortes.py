#!/usr/bin/env python3
"""
capturar_cortes.py

Captura 3 tablas (GENERAL, ZONAL, SUPERVISOR) como imágenes PNG
y las envía a WhatsApp con etiqueta "CORTE XAM".

Integra:
- screenshot_safe.py (mutex para evitar conflictos con MOVISTAR)
- wa_sender_antibang.py (envío seguro sin baneo)
- config.json (destinos)

Uso:
    python capturar_cortes.py --hora 8 --destino test
    python capturar_cortes.py --hora 10 --destino canal
"""

import argparse
import logging
import json
import random
import sys
import time
from pathlib import Path
from datetime import datetime
import xlwings as xw

from screenshot_safe import ScreenshotManager, capturar_tabla_excel

# Importar wa_client.py (mismo patrón que el resto del proyecto)
sys.path.insert(0, str(Path(__file__).parent))
from wa_client import WhatsAppClient

# ════════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ════════════════════════════════════════════════════════════════════════════════

BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config.json"
IMAGENES_DIR = BASE_DIR / "cortes_imagenes"
IMAGENES_DIR.mkdir(exist_ok=True)

HORA_LBL = {
    8: '8AM', 9: '9AM', 10: '10AM', 11: '11AM', 12: '12PM',
    13: '1PM', 14: '2PM', 15: '3PM', 16: '4PM', 17: '5PM', 18: '6PM'
}

# Logger
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
_logger = logging.getLogger("capturar_cortes")


# ════════════════════════════════════════════════════════════════════════════════
# FUNCIONES
# ════════════════════════════════════════════════════════════════════════════════

def cargar_config() -> dict:
    """Carga configuración desde config.json"""
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        _logger.error(f"[ERROR] Error cargando config: {e}")
        raise


def detectar_rangos_vendedor(archivo_excel: str) -> list:
    """
    Lee la hoja VENDEDOR con openpyxl y detecta el rango A:K de cada tabla
    de supervisor buscando filas cuya celda A tiene fondo #403151 (título).

    Retorna lista de (nombre_supervisor, rango_excel) p.ej.:
        [("DIANA MADALENGOITIA", "A1:K15"), ("EDWIN VIELMA", "A20:K31"), ...]
    """
    import openpyxl
    try:
        wb = openpyxl.load_workbook(archivo_excel, data_only=True)
        if 'VENDEDOR' not in wb.sheetnames:
            return []
        ws = wb['VENDEDOR']
        max_row = ws.max_row

        COLOR_TIT = {'403151', 'FF403151', '00403151'}

        bloques = []
        i = 1
        while i <= max_row:
            c = ws.cell(i, 1)
            fg = c.fill.fgColor
            color = fg.rgb if fg.type == 'rgb' else ''
            if color.upper().lstrip('0') in {'403151'} or color.upper() in {'FF403151', '00403151'}:
                nombre_sup = str(c.value or '').strip()
                fila_ini = i
                # Buscar la fila TOTAL: siguiente fila con mismo color o "TOTAL" en col A
                fila_fin = fila_ini
                for j in range(i + 1, min(i + 30, max_row + 1)):
                    v = ws.cell(j, 1).value
                    if v and str(v).strip() == 'TOTAL':
                        fila_fin = j
                        break
                if fila_fin == fila_ini:
                    fila_fin = i  # tabla vacía, saltar
                rango = f'A{fila_ini}:K{fila_fin}'
                bloques.append((nombre_sup, rango))
                i = fila_fin + 1
            else:
                i += 1

        return bloques
    except Exception as e:
        _logger.error(f"[ERROR] detectar_rangos_vendedor: {e}")
        return []


def capturar_tablas(archivo_excel: str, hora: int) -> dict:
    """
    Captura las tablas del Excel como PNG:
      - Hoja CORTE VENTAS: GENERAL, ZONAL, SUPERVISOR (rangos fijos)
      - Hoja VENDEDOR: una imagen por supervisor (rangos detectados dinámicamente)

    Retorna:
        {
            'exito': bool,
            'general': ruta_png,
            'zonal': ruta_png,
            'supervisor': ruta_png,
            'vendedor': [(nombre_sup, ruta_png), ...]
        }
    """
    etiqueta_hora = HORA_LBL.get(hora, f"{hora}H")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    mgr = ScreenshotManager(f"SSFF_Corte_{etiqueta_hora}")
    if not mgr.adquirir_lock(timeout=30):
        _logger.error(f"[ERROR] Timeout esperando lock para {etiqueta_hora}")
        return {'exito': False}

    resultado = {'exito': False, 'general': None, 'zonal': None,
                 'supervisor': None, 'vendedor': []}

    try:
        app = xw.App(visible=True)
        libro = app.books.open(archivo_excel)

        # ── Hoja CORTE VENTAS: 3 tablas con rangos fijos ──────────────────────
        hoja_cv = libro.sheets['CORTE VENTAS']
        tablas_cv = [
            ('general',    'A2:H8'),
            ('zonal',      'K2:U14'),   # Hasta columna U, hasta fila TOTAL
            ('supervisor', 'X2:AH15'),  # X2:AH15 - tabla SUPERVISOR con todos los supervisores + TOTAL
        ]
        for nombre, rango in tablas_cv:
            ruta_png = IMAGENES_DIR / f'{nombre.upper()}_{etiqueta_hora}_{timestamp}.png'
            _logger.info(f"[CAPTURE] {nombre.upper()} ({rango})...")
            if capturar_tabla_excel(hoja_cv, rango, str(ruta_png), escala=2.5):
                resultado[nombre] = str(ruta_png)
                _logger.info(f"[OK] {nombre.upper()} guardado")
            else:
                _logger.error(f"[ERROR] Fallo {nombre.upper()}")
            time.sleep(0.5)

        # ── Hoja VENDEDOR: una captura por supervisor ──────────────────────────
        rangos_vend = detectar_rangos_vendedor(archivo_excel)
        if rangos_vend and 'VENDEDOR' in [s.name for s in libro.sheets]:
            hoja_v = libro.sheets['VENDEDOR']
            for nombre_sup, rango in rangos_vend:
                slug = nombre_sup.replace(' ', '_')[:20]
                ruta_png = IMAGENES_DIR / f'VEND_{slug}_{etiqueta_hora}_{timestamp}.png'
                _logger.info(f"[CAPTURE] VENDEDOR {nombre_sup} ({rango})...")
                if capturar_tabla_excel(hoja_v, rango, str(ruta_png), escala=2.5):
                    resultado['vendedor'].append((nombre_sup, str(ruta_png)))
                    _logger.info(f"[OK] {nombre_sup} guardado")
                else:
                    _logger.error(f"[ERROR] Fallo VENDEDOR {nombre_sup}")
                time.sleep(0.5)
        else:
            _logger.warning("[WARN] Hoja VENDEDOR no encontrada o sin tablas")

        libro.close()
        app.quit()

        resultado['exito'] = all(resultado[k] for k in ['general', 'zonal', 'supervisor'])
        if resultado['exito']:
            _logger.info(f"[OK] Captura completada — {len(resultado['vendedor'])} tablas vendedor")
        else:
            _logger.error("[ERROR] Algunas tablas principales fallaron")

    except Exception as e:
        _logger.error(f"[ERROR] Error durante captura: {e}")
    finally:
        mgr.liberar_lock()

    return resultado


def enviar_corte_whatsapp(imagenes: dict, destino: str, hora: int, config: dict) -> bool:
    """
    Envía 3 imágenes a WhatsApp con etiqueta "CORTE XAM".

    Args:
        imagenes: Dict con rutas {'general': ..., 'zonal': ..., 'supervisor': ...}
        destino: Clave en config['cortes_horarios']['destinos']
        hora: Hora (8-18)
        config: Config cargada

    Returns:
        True si envío exitoso
    """
    if not imagenes.get('exito'):
        _logger.error("[ERROR] No hay imágenes para enviar")
        return False

    etiqueta_hora = HORA_LBL.get(hora, f"{hora}H")
    msg_titulo = f"CORTE {etiqueta_hora}"

    try:
        destinos = config.get('cortes_horarios', {}).get('destinos', {})
        numero_destino = destinos.get(destino)

        if not numero_destino:
            _logger.error(f"[ERROR] Destino '{destino}' no configurado en cortes_horarios.destinos")
            return False

        # Usar wa_client.py (mismo patrón que el resto del proyecto)
        wa = WhatsAppClient()

        # Tablas principales: GENERAL, ZONAL, SUPERVISOR
        archivos_principales = [
            ('GENERAL',    imagenes.get('general')),
            ('ZONAL',      imagenes.get('zonal')),
            ('SUPERVISOR', imagenes.get('supervisor')),
        ]
        for nombre_tabla, ruta_img in archivos_principales:
            if not ruta_img or not Path(ruta_img).exists():
                _logger.warning(f"[WARN] Imagen no encontrada: {ruta_img}")
                continue
            caption = f"{msg_titulo} - {nombre_tabla}"
            _logger.info(f"[SEND] {nombre_tabla}...")
            r = wa.send_image(numero_destino, ruta_img, caption=caption)
            if not r.get('success'):
                _logger.error(f"[ERROR] Fallo {nombre_tabla}: {r.get('error')}")
                return False
            _logger.info(f"[OK] {nombre_tabla} enviado")
            time.sleep(random.uniform(3, 6))

        # Tablas vendedor: una por supervisor
        tablas_vendedor = imagenes.get('vendedor', [])
        if tablas_vendedor:
            _logger.info(f"[SEND] Enviando {len(tablas_vendedor)} tablas vendedor...")
            for nombre_sup, ruta_img in tablas_vendedor:
                if not ruta_img or not Path(ruta_img).exists():
                    _logger.warning(f"[WARN] Imagen vendedor no encontrada: {ruta_img}")
                    continue
                caption = f"{msg_titulo} - {nombre_sup}"
                _logger.info(f"[SEND] VENDEDOR {nombre_sup}...")
                r = wa.send_image(numero_destino, ruta_img, caption=caption)
                if not r.get('success'):
                    _logger.error(f"[ERROR] Fallo VENDEDOR {nombre_sup}: {r.get('error')}")
                    # No abortar — continuar con el siguiente supervisor
                else:
                    _logger.info(f"[OK] {nombre_sup} enviado")
                time.sleep(random.uniform(3, 6))

        _logger.info(f"[OK] Corte {etiqueta_hora} completo — enviado a {destino} ({numero_destino})")
        return True

    except Exception as e:
        _logger.error(f"[ERROR] Error enviando corte: {e}")
        return False


def enviar_mensaje_cuota_diaria(destino: str, config: dict) -> bool:
    """Envía mensaje de cuota del día después de enviar todas las capturas.
    Calcula cuota_dia = (Cuota_Junio - Avance_Hasta_Ayer) / Dias_Habiles_Restantes"""
    try:
        import pandas as pd
        import pyodbc
        import os
        from datetime import date, timedelta

        destinos = config.get('cortes_horarios', {}).get('destinos', {})
        numero_destino = destinos.get(destino)

        if not numero_destino:
            _logger.error(f"[ERROR] Destino '{destino}' no configurado para cuota")
            return False

        hoy = date.today()
        ayer = hoy - timedelta(days=1)

        # 1. Cargar cuota total de junio desde CuotaJunioV2.xlsx (columna 'Nueva Cuota')
        cuota_path = "C:/proyectos/SSFF/files/CuotaJunioV2.xlsx"
        df = pd.read_excel(cuota_path, dtype={'RUTA': str})
        cuota_junio = df['Nueva Cuota'].sum()

        # 2. Obtener avance hasta ayer desde SQL
        env_path = Path(__file__).parent / ".env"
        env = {}
        if env_path.exists():
            with open(env_path, encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#') and '=' in line:
                        k, v = line.split('=', 1)
                        env[k.strip()] = v.strip()

        conn_str = (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={env.get('SQL_SERVER')};"
            f"DATABASE={env.get('SQL_DATABASE')};"
            f"UID={env.get('SQL_USER')};"
            f"PWD={env.get('SQL_PASSWORD')};"
            f"TrustServerCertificate=yes;"
        )
        conn = pyodbc.connect(conn_str)

        # Query: sumar montos hasta ayer (no incluir hoy)
        query = f"""
        SELECT SUM([monto]) AS total FROM [eAuren].[dbo].[base_com]
        WHERE mes='2606' AND cstatus != 'A'
        AND CAST(fecha AS DATE) < '{ayer.strftime('%Y-%m-%d')}'
        """
        df_avance = pd.read_sql(query, conn)
        avance_ayer = float(df_avance['total'].iloc[0] or 0.0)
        conn.close()

        # 3. Calcular días hábiles restantes (lun-sab, sin 29/06 feriado)
        dias_habiles = []
        fecha_actual = hoy
        while fecha_actual <= date(2026, 6, 30):
            dia_semana = fecha_actual.weekday()
            if dia_semana < 6 and fecha_actual != date(2026, 6, 29):
                dias_habiles.append(fecha_actual)
            fecha_actual += timedelta(days=1)

        dias_habiles_restantes = len(dias_habiles)

        # 4. Calcular cuota_dia
        numerador = cuota_junio - avance_ayer
        cuota_dia = numerador / dias_habiles_restantes if dias_habiles_restantes > 0 else 0.0

        # 5. Enviar mensaje
        mensaje = (
            f"*Cuota del dia*\n\n"
            f"La cuota dia de hoy es  S/ {cuota_dia:,.0f}\n\n"
            f"@51944956042\n"
            f"@51924876915"
        )

        wa = WhatsAppClient()
        mentions = [
            "51944956042@c.us",  # Jesus Ascencios
            "51924876915@c.us"   # Mercedes Loaiza
        ]
        resultado = wa.send_mention(numero_destino, mensaje, mentions=mentions)

        if resultado.get('success'):
            _logger.info(f"[OK] Mensaje de cuota enviado: S/ {cuota_dia:,.0f}")
            return True
        else:
            _logger.error(f"[ERROR] Fallo al enviar cuota: {resultado.get('error')}")
            return False

    except Exception as e:
        _logger.error(f"[ERROR] Error enviando cuota: {e}")
        return False


def enviar_alerta_tecnica(destino: str, hora: int, config: dict, motivo: str = "") -> bool:
    """Envía mensaje de texto cuando no se pueden generar las capturas."""
    etiqueta_hora = HORA_LBL.get(hora, f"{hora}H")
    try:
        destinos = config.get('cortes_horarios', {}).get('destinos', {})
        numero_destino = destinos.get(destino)
        if not numero_destino:
            _logger.error(f"[ERROR] Destino '{destino}' no configurado")
            return False

        wa = WhatsAppClient()
        msg = (
            f"[AVISO TECNICO] Corte {etiqueta_hora}\n"
            f"Se presento un problema tecnico al generar el reporte.\n"
            f"La informacion sera enviada lo antes posible.\n"
        )
        if motivo:
            msg += f"Detalle: {motivo[:120]}"

        resultado = wa.send_message(numero_destino, msg)
        if resultado.get('success'):
            _logger.info(f"[OK] Alerta tecnica enviada a {destino}")
            return True
        else:
            _logger.error(f"[ERROR] No se pudo enviar alerta: {resultado.get('error')}")
            return False
    except Exception as e:
        _logger.error(f"[ERROR] Excepcion enviando alerta: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(
        description="Captura cortes horarios y envía a WhatsApp"
    )
    parser.add_argument('--hora', type=int, required=True,
                        help="Hora (8-18)")
    parser.add_argument('--destino', default='test',
                        help="Destino en config.json (test, canal, supervisores_f8, etc.)")
    parser.add_argument('--archivo', default=None,
                        help="Ruta a CORTE_VENTAS_*.xlsx (default: busca el más reciente)")
    parser.add_argument('--solo-imagenes', action='store_true',
                        help="Captura imágenes pero no envía WhatsApp")

    args = parser.parse_args()

    # Validar hora
    if not 8 <= args.hora <= 18:
        _logger.error("[ERROR] Hora debe estar entre 8 y 18")
        return

    config = cargar_config()

    # Buscar archivo Excel del día actual (nombre fijo por fecha)
    if not args.archivo:
        from datetime import date
        nombre_hoy = f"CORTE_VENTAS_{date.today().strftime('%Y%m%d')}.xlsx"
        args.archivo = str(BASE_DIR / nombre_hoy)
        _logger.info(f"[FILE] Usando: {args.archivo}")

    if not Path(args.archivo).exists():
        _logger.error(f"[ERROR] Archivo no encontrado: {args.archivo}")
        if not args.solo_imagenes:
            enviar_alerta_tecnica(args.destino, args.hora, config,
                                  "Archivo de reporte no disponible")
        return

    # Capturar tablas
    _logger.info(f"\n{'='*70}")
    _logger.info(f"[CAPTURE] CAPTURANDO CORTE {args.hora:02d}:00")
    _logger.info(f"{'='*70}\n")

    imagenes = capturar_tablas(args.archivo, args.hora)

    if not imagenes['exito']:
        _logger.error("[ERROR] Fallo en captura — enviando alerta tecnica")
        if not args.solo_imagenes:
            enviar_alerta_tecnica(args.destino, args.hora, config,
                                  "Error al capturar las tablas del reporte")
        return

    if args.solo_imagenes:
        _logger.info("[OK] Imagenes capturadas (--solo-imagenes activo)")
        return

    # Enviar a WhatsApp
    _logger.info(f"\n{'='*70}")
    _logger.info(f"[SEND] ENVIANDO A WHATSAPP")
    _logger.info(f"{'='*70}\n")

    if enviar_corte_whatsapp(imagenes, args.destino, args.hora, config):
        _logger.info(f"\n[OK] CORTE {args.hora:02d}:00 COMPLETADO")

        # Enviar cuota del día solo en el corte de 8AM (primera ejecución)
        if args.hora == 8:
            _logger.info(f"\n{'='*70}")
            _logger.info(f"[CUOTA] ENVIANDO RESUMEN DE CUOTA DEL DÍA")
            _logger.info(f"{'='*70}\n")
            enviar_mensaje_cuota_diaria(args.destino, config)
    else:
        _logger.error(f"\n[ERROR] Fallo enviando corte")


if __name__ == "__main__":
    main()
