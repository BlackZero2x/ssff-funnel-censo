#!/usr/bin/env python3
"""
capturar_funnel.py

Captura la hoja RESUMEN de FUNNEL_PREVENTA_<dia>_<fecha>.xlsx como imagen PNG
y la envía a WhatsApp con etiqueta "FUNNEL PREVENTA <hora_corte>".

Reutiliza wa_client.py y config.json de Reportes_ssff_wsp (mismo canal
que los cortes de ventas horarios).

Uso:
    python capturar_funnel.py --corte 9AM --destino test
    python capturar_funnel.py --corte 11AM --destino canal
"""

import argparse
import json
import logging
import sys
import time
from datetime import date
from pathlib import Path
import xlwings as xw

sys.path.insert(0, r"C:\proyectos\shared")
from screenshot_safe import ScreenshotManager, capturar_tabla_excel

REPORTES_DIR = Path(r"C:\proyectos\SSFF\Reportes_ssff_wsp")
sys.path.insert(0, str(REPORTES_DIR))
from wa_client import WhatsAppClient

BASE_DIR     = Path(__file__).parent
CONFIG_PATH  = REPORTES_DIR / "config.json"
IMAGENES_DIR = BASE_DIR / "funnel_imagenes"
IMAGENES_DIR.mkdir(exist_ok=True)
LOG_DIR      = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

DIAS_SEMANA = {0: 'Lunes', 1: 'Martes', 2: 'Miércoles', 3: 'Jueves',
               4: 'Viernes', 5: 'Sábado', 6: 'Domingo'}

RANGO_RESUMEN = "A1:H12"  # título + header + 9 supervisores + TOTAL AUREN


def _setup_logger() -> logging.Logger:
    fecha_str = date.today().strftime("%Y%m%d")
    log_file  = LOG_DIR / f"funnel_{fecha_str}.log"

    logger = logging.getLogger("capturar_funnel")
    if logger.handlers:
        return logger
    logger.setLevel(logging.DEBUG)

    fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s",
                             datefmt="%Y-%m-%d %H:%M:%S")
    fh = logging.FileHandler(log_file, encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)

    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger


_logger = _setup_logger()


def cargar_config() -> dict:
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)


def buscar_archivo_hoy() -> Path | None:
    """Busca FUNNEL_PREVENTA_<dia>_<dd_mm_YYYY>.xlsx del día actual."""
    hoy = date.today()
    nombre_dia = DIAS_SEMANA[hoy.weekday()]
    fecha_str  = hoy.strftime('%d_%m_%Y')
    archivo = BASE_DIR / f"FUNNEL_PREVENTA_{nombre_dia}_{fecha_str}.xlsx"
    return archivo if archivo.exists() else None


def capturar_resumen(archivo_excel: Path, etiqueta_corte: str) -> str | None:
    """Captura la hoja RESUMEN como PNG. Retorna la ruta o None si falla."""
    slug = etiqueta_corte.replace(':', '')  # ':' no es válido en nombres de archivo Windows
    timestamp = date.today().strftime("%Y%m%d") + "_" + slug
    ruta_png  = IMAGENES_DIR / f'RESUMEN_{timestamp}.png'

    mgr = ScreenshotManager(f"SSFF_Funnel_{slug}")
    if not mgr.adquirir_lock(timeout=30):
        _logger.error(f"[ERROR] Timeout esperando lock para {etiqueta_corte}")
        return None

    app = None
    try:
        # El archivo puede estar recién guardado por openpyxl (proceso Python) —
        # dar un margen antes de que Excel/COM (proceso distinto) intente abrirlo,
        # con reintentos por si el handle tarda en liberarse.
        app = xw.App(visible=True)
        libro = None
        for intento in range(1, 4):
            try:
                libro = app.books.open(str(archivo_excel))
                break
            except Exception as e:
                if intento == 3:
                    raise
                _logger.warning(f"[WARN] Intento {intento}/3 abriendo Excel falló: {e} — reintentando en 2s")
                time.sleep(2)

        try:
            hoja = libro.sheets['RESUMEN']
            _logger.info(f"[CAPTURE] RESUMEN ({RANGO_RESUMEN})...")
            if capturar_tabla_excel(hoja, RANGO_RESUMEN, str(ruta_png), escala=2.5):
                _logger.info(f"[OK] RESUMEN guardado: {ruta_png}")
                return str(ruta_png)
            _logger.error("[ERROR] Fallo captura RESUMEN")
            return None
        finally:
            libro.close()
    except Exception as e:
        _logger.error(f"[ERROR] Error durante captura: {e}")
        return None
    finally:
        if app is not None:
            app.quit()
        mgr.liberar_lock()


def enviar_whatsapp(ruta_png: str, ruta_xlsx: Path, destino: str, etiqueta_corte: str, config: dict) -> bool:
    destinos = config.get('cortes_horarios', {}).get('destinos', {})
    numero_destino = destinos.get(destino)
    if not numero_destino:
        _logger.error(f"[ERROR] Destino '{destino}' no configurado en cortes_horarios.destinos")
        return False

    wa = WhatsAppClient()

    if etiqueta_corte == '5:30PM':
        caption_img = f"*CIERRE Funnel Preventa — {etiqueta_corte}*"
    else:
        caption_img = f"*Funnel Preventa Corte: {etiqueta_corte}*"

    r = wa.send_image(numero_destino, ruta_png, caption=caption_img)
    if not r.get('success'):
        _logger.error(f"[ERROR] Fallo envío de imagen: {r.get('error')}")
        return False
    _logger.info(f"[OK] Imagen FUNNEL PREVENTA {etiqueta_corte} → {destino}")

    time.sleep(3)

    r2 = wa.send_file(numero_destino, str(ruta_xlsx), caption="")
    if not r2.get('success'):
        _logger.error(f"[ERROR] Fallo envío de archivo: {r2.get('error')}")
        return False
    _logger.info(f"[OK] Archivo FUNNEL PREVENTA {etiqueta_corte} → {destino}")

    return True


def main():
    parser = argparse.ArgumentParser(description="Captura hoja RESUMEN del Funnel Preventa y envía a WhatsApp")
    parser.add_argument('--corte', required=True, help="Etiqueta del corte (9AM, 11AM, 1PM, 3PM, 5PM)")
    parser.add_argument('--destino', default='test', help="Destino en config.json (test, canal, ...)")
    parser.add_argument('--archivo', default=None, help="Ruta al FUNNEL_PREVENTA_*.xlsx (default: busca el de hoy)")
    parser.add_argument('--solo-imagen', action='store_true', help="Captura pero no envía WhatsApp")
    args = parser.parse_args()

    config = cargar_config()

    archivo = Path(args.archivo) if args.archivo else buscar_archivo_hoy()
    if not archivo or not archivo.exists():
        _logger.error(f"[ERROR] Archivo no encontrado: {archivo}")
        sys.exit(1)

    _logger.info(f"{'='*70}")
    _logger.info(f"[CAPTURE] FUNNEL PREVENTA — corte {args.corte} — {archivo.name}")
    _logger.info(f"{'='*70}")

    ruta_png = capturar_resumen(archivo, args.corte)
    if not ruta_png:
        _logger.error("[ERROR] No se pudo capturar RESUMEN")
        sys.exit(1)

    if args.solo_imagen:
        _logger.info("[OK] Imagen capturada (--solo-imagen activo)")
        return

    if not enviar_whatsapp(ruta_png, archivo, args.destino, args.corte, config):
        sys.exit(1)

    _logger.info(f"[OK] FUNNEL PREVENTA {args.corte} completado")


if __name__ == "__main__":
    main()
