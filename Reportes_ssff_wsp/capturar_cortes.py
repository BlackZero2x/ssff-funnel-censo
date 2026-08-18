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

sys.path.insert(0, r"C:\proyectos\shared")
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
LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

HORA_LBL = {
    8: '8AM', 9: '9AM', 10: '10AM', 11: '11AM', 12: '12PM',
    13: '1PM', 14: '2PM', 15: '3PM', 16: '4PM', 17: '5PM', 18: '6PM'
}


def _setup_logger() -> logging.Logger:
    """Logger a consola + archivo diario logs/cortes_YYYYMMDD.log."""
    fecha_str = datetime.now().strftime("%Y%m%d")
    log_file = LOG_DIR / f"cortes_{fecha_str}.log"

    logger = logging.getLogger("capturar_cortes")
    if logger.handlers:
        return logger  # ya configurado (evita duplicar handlers en reimports)
    logger.setLevel(logging.DEBUG)

    fmt = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    fh = logging.FileHandler(log_file, encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)

    logger.addHandler(fh)
    logger.addHandler(ch)

    # screenshot_safe.py (mutex EXCEL_COM) usa su propio logger sin handlers;
    # sin esto, sus mensajes [LOCK] (adquirido/esperando/timeout) se pierden
    # y no quedan registrados en el log diario — dificulta diagnosticar
    # colisiones con otros proyectos (ej. AVANCE_MOVISTAR).
    lock_logger = logging.getLogger("screenshot_safe")
    lock_logger.setLevel(logging.DEBUG)
    lock_logger.addHandler(fh)
    lock_logger.addHandler(ch)
    lock_logger.propagate = False

    return logger


_logger = _setup_logger()


# ════════════════════════════════════════════════════════════════════════════════
# FUNCIONES
# ════════════════════════════════════════════════════════════════════════════════

def _cargar_alertas() -> dict:
    """Lee el JSON de alertas de rezago generado por generar_corte_ventas.py."""
    from datetime import date
    alertas_path = BASE_DIR / f"CORTE_ALERTAS_{date.today().strftime('%Y%m%d')}.json"
    if not alertas_path.exists():
        return {}
    try:
        import json
        with open(alertas_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        _logger.warning(f"[WARN] No se pudo leer alertas: {e}")
        return {}


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
    Lee la hoja VENDEDOR con openpyxl y detecta el rango A:M de cada tabla
    de supervisor buscando filas cuya celda A tiene fondo #403151 (título).

    Retorna lista de (nombre_supervisor, rango_excel) p.ej.:
        [("DIANA MADALENGOITIA", "A1:M15"), ("EDWIN VIELMA", "A20:M31"), ...]
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
                # Buscar la fila TOTAL hasta el fin de la hoja
                fila_fin = fila_ini
                for j in range(i + 1, max_row + 1):
                    v = ws.cell(j, 1).value
                    if v and str(v).strip() == 'TOTAL':
                        fila_fin = j
                        break
                    # Otro encabezado de supervisor = nueva tabla, detenerse
                    fg2 = ws.cell(j, 1).fill.fgColor
                    color2 = fg2.rgb if fg2.type == 'rgb' else ''
                    if color2.upper() in COLOR_TIT and j != i:
                        break
                if fila_fin == fila_ini:
                    fila_fin = i  # tabla vacía, saltar
                rango = f'A{fila_ini}:M{fila_fin}'
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
    Captura las tablas del Excel como PNG (con 1 reintento automático si falla
    por error COM — p.ej. colisión con otro proceso que cerró/abrió Excel a la
    vez, como el taskkill de AVANCE_MOVISTAR).

    Retorna:
        {
            'exito': bool,
            'general': ruta_png,
            'zonal': ruta_png,
            'supervisor': ruta_png,
            'vendedor': [(nombre_sup, ruta_png), ...]
        }
    """
    ESPERA_REINTENTO_COM = 30  # segundos

    for intento in range(1, 3):
        resultado = _intentar_capturar_tablas(archivo_excel, hora)
        if resultado['exito']:
            return resultado
        if intento == 1:
            _logger.warning(
                f"[REINTENTO] Captura falló en intento 1 — reintentando en "
                f"{ESPERA_REINTENTO_COM}s (posible colisión Excel/COM con otro proceso)"
            )
            time.sleep(ESPERA_REINTENTO_COM)

    _logger.error("[ERROR] Captura falló tras 2 intentos")
    return resultado


def _intentar_capturar_tablas(archivo_excel: str, hora: int) -> dict:
    """Un intento de captura completo: adquiere lock, abre Excel, captura, libera lock."""
    etiqueta_hora = HORA_LBL.get(hora, f"{hora}H")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    mgr = ScreenshotManager(f"SSFF_Corte_{etiqueta_hora}")
    # 90s: MOVISTAR_AVANCE_TDS puede retener el lock hasta ~74s (medido en
    # logs del 18/08/2026) — con margen para no fallar el primer intento
    # frente a esa retención.
    if not mgr.adquirir_lock(timeout=90):
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
        # Si xw.App()/books.open() falló a medio camino, puede quedar un proceso
        # EXCEL.EXE huérfano bloqueando el reintento — se limpia explícitamente.
        try:
            if 'libro' in locals() and libro is not None:
                libro.close()
        except Exception:
            pass
        try:
            if 'app' in locals() and app is not None:
                app.quit()
        except Exception:
            pass
    finally:
        mgr.liberar_lock()

    return resultado


def enviar_corte_whatsapp(imagenes: dict, destino: str, hora: int, config: dict) -> bool:
    """
    Envía 3 imágenes a WhatsApp con etiqueta "CORTE XAM" o "CIERRE DE HOY" (6PM).

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

    # A las 6PM (hora 18): mostrar "CIERRE DE HOY - Día dd/mm"
    if hora == 18:
        from datetime import datetime
        DIAS_SEMANA = {0: 'Lunes', 1: 'Martes', 2: 'Miércoles', 3: 'Jueves',
                       4: 'Viernes', 5: 'Sábado', 6: 'Domingo'}
        hoy = datetime.now()
        nombre_dia = DIAS_SEMANA.get(hoy.weekday(), 'Día')
        fecha_str = hoy.strftime('%d/%m')
        msg_titulo = f"CIERRE DE HOY - {nombre_dia} {fecha_str}"
    else:
        msg_titulo = f"CORTE {etiqueta_hora}"

    try:
        destinos = config.get('cortes_horarios', {}).get('destinos', {})
        numero_destino = destinos.get(destino)

        if not numero_destino:
            _logger.error(f"[ERROR] Destino '{destino}' no configurado en cortes_horarios.destinos")
            return False

        wa = WhatsAppClient()
        alertas = _cargar_alertas()
        vendedor_grupos_norm = {k.strip().upper(): v
                                for k, v in config.get('vendedor_grupos', {}).items()}

        _logger.info(f"[INFO] Alertas cargadas: canal={'sí' if alertas.get('canal') else 'vacío'}, "
                     f"supervisores con msg={sum(1 for v in alertas.get('supervisores', {}).values() if v)}")
        _logger.info(f"[INFO] Grupos individuales configurados: {list(vendedor_grupos_norm.keys())}")

        # ── PASO 1: todas las imágenes al canal general ───────────────────────
        _logger.info(f"[PASO 1] Enviando imágenes al canal general ({numero_destino})...")
        for nombre_tabla, ruta_img in [
            ('GENERAL',    imagenes.get('general')),
            ('ZONAL',      imagenes.get('zonal')),
            ('SUPERVISOR', imagenes.get('supervisor')),
        ]:
            if not ruta_img or not Path(ruta_img).exists():
                _logger.warning(f"[WARN] Imagen no encontrada: {ruta_img}")
                continue
            r = wa.send_image(numero_destino, ruta_img, caption=f"{msg_titulo} - {nombre_tabla}")
            if not r.get('success'):
                _logger.error(f"[ERROR] Fallo {nombre_tabla}: {r.get('error')}")
                return False
            _logger.info(f"[OK] {nombre_tabla} → canal")
            time.sleep(random.uniform(3, 6))

        tablas_vendedor = imagenes.get('vendedor', [])
        _logger.info(f"[PASO 1] Tablas vendedor a enviar: {len(tablas_vendedor)}")
        for nombre_sup, ruta_img in tablas_vendedor:
            if not ruta_img or not Path(ruta_img).exists():
                _logger.warning(f"[WARN] Imagen vendedor no encontrada: {ruta_img}")
                continue
            r = wa.send_image(numero_destino, ruta_img, caption=f"{msg_titulo} - {nombre_sup}")
            if not r.get('success'):
                _logger.error(f"[ERROR] Fallo VENDEDOR {nombre_sup}: {r.get('error')}")
            else:
                _logger.info(f"[OK] VENDEDOR {nombre_sup} → canal")
            time.sleep(random.uniform(3, 6))

        # ── PASO 2: alerta de rezago de supervisores al canal general ─────────
        _logger.info(f"[PASO 2] Alerta rezago canal...")
        msg_canal = alertas.get('canal', '') if hora != 8 else ''
        if hora == 8:
            _logger.info(f"[PASO 2] Corte 8AM — alertas de rezago omitidas")
        if msg_canal:
            time.sleep(random.uniform(3, 6))
            r = wa.send_text(numero_destino, msg_canal)
            if r.get('success'):
                _logger.info(f"[OK] Alerta rezago → canal")
            else:
                _logger.error(f"[ERROR] Fallo alerta rezago canal: {r.get('error')}")
        else:
            _logger.info(f"[PASO 2] Sin alerta (ningún supervisor bajo 80% o archivo vacío)")

        # ── PASO 3: imagen + alerta a cada grupo individual de supervisor ─────
        _logger.info(f"[PASO 3] Enviando a grupos individuales ({len(vendedor_grupos_norm)} configurados)...")
        grupos_enviados = 0
        for nombre_sup, ruta_img in tablas_vendedor:
            nombre_sup_norm = nombre_sup.strip().upper()
            grupo_sup = vendedor_grupos_norm.get(nombre_sup_norm)
            if not grupo_sup:
                _logger.debug(f"[PASO 3] {nombre_sup}: sin grupo configurado, omitiendo")
                continue
            if not ruta_img or not Path(ruta_img).exists():
                _logger.warning(f"[WARN] Imagen no existe para grupo {nombre_sup}: {ruta_img}")
                continue

            r2 = wa.send_image(grupo_sup, ruta_img, caption=f"{msg_titulo} - {nombre_sup}")
            if not r2.get('success'):
                _logger.error(f"[ERROR] Fallo imagen grupo {nombre_sup}: {r2.get('error')}")
            else:
                _logger.info(f"[OK] VENDEDOR {nombre_sup} → grupo individual ({grupo_sup})")
                grupos_enviados += 1
            time.sleep(random.uniform(3, 6))

            msg_vend = ''
            if hora != 8:
                for k, v in alertas.get('supervisores', {}).items():
                    if k.strip().upper() == nombre_sup_norm:
                        msg_vend = v
                        break
            if msg_vend:
                r3 = wa.send_text(grupo_sup, msg_vend)
                if r3.get('success'):
                    _logger.info(f"[OK] Alerta rezago vendedores → grupo {nombre_sup}")
                else:
                    _logger.error(f"[ERROR] Fallo alerta rezago grupo {nombre_sup}: {r3.get('error')}")
                time.sleep(random.uniform(3, 6))
            else:
                _logger.info(f"[PASO 3] Sin alerta rezago para {nombre_sup} (ningún vendedor bajo 80%)")

        _logger.info(f"[PASO 3] Grupos enviados: {grupos_enviados}/{len(vendedor_grupos_norm)}")

        _logger.info(f"[OK] Corte {etiqueta_hora} completo — canal + grupos individuales")
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
        mes_actual = hoy.strftime('%y%m')  # '2606' jun, '2607' jul, etc.
        query = f"""
        SELECT SUM([monto]) AS total FROM [eAuren].[dbo].[base_com]
        WHERE mes='{mes_actual}' AND cstatus != 'A'
        AND CAST(fecha AS DATE) < '{ayer.strftime('%Y-%m-%d')}'
        """
        df_avance = pd.read_sql(query, conn)
        avance_ayer = float(df_avance['total'].iloc[0] or 0.0)
        conn.close()

        # 3. Calcular días hábiles restantes (lun-sab) hasta fin del mes actual
        import calendar
        ultimo_dia = calendar.monthrange(hoy.year, hoy.month)[1]
        fin_mes = date(hoy.year, hoy.month, ultimo_dia)
        dias_habiles = []
        fecha_actual = hoy
        while fecha_actual <= fin_mes:
            if fecha_actual.weekday() < 6:  # lun=0 … sab=5
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
        msg = f"⚠️ Corte {etiqueta_hora} — reporte en proceso, sera enviado a la brevedad."

        resultado = wa.send_text(numero_destino, msg)
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
