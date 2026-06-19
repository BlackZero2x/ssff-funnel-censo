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


def capturar_tablas(archivo_excel: str, hora: int) -> dict:
    """
    Captura 3 tablas del Excel como PNG.

    Args:
        archivo_excel: Ruta al CORTE_VENTAS_*.xlsx
        hora: Hora (8-18)

    Returns:
        {
            'exito': bool,
            'general': 'path/a/GENERAL.png',
            'zonal': 'path/a/ZONAL.png',
            'supervisor': 'path/a/SUPERVISOR.png'
        }
    """
    etiqueta_hora = HORA_LBL.get(hora, f"{hora}H")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Adquirir lock exclusivo
    mgr = ScreenshotManager(f"SSFF_Corte_{etiqueta_hora}")

    if not mgr.adquirir_lock(timeout=30):
        _logger.error(f"[ERROR] Timeout esperando lock para {etiqueta_hora}")
        return {'exito': False}

    resultado = {'exito': False, 'general': None, 'zonal': None, 'supervisor': None}

    try:
        # Abrir archivo con xlwings
        app = xw.App(visible=True)
        libro = app.books.open(archivo_excel)
        hoja = libro.sheets[0]  # Primera hoja

        # Rangos de las 3 tablas
        tablas = {
            'general': ('A2:H8', f'GENERAL_{etiqueta_hora}_{timestamp}.png'),
            'zonal': ('K2:U12', f'ZONAL_{etiqueta_hora}_{timestamp}.png'),
            'supervisor': ('X2:AH14', f'SUPERVISOR_{etiqueta_hora}_{timestamp}.png'),
        }

        for nombre, (rango, archivo) in tablas.items():
            ruta_png = IMAGENES_DIR / archivo

            _logger.info(f"[CAPTURE] Capturando {nombre.upper()} ({rango})...")

            if capturar_tabla_excel(hoja, rango, str(ruta_png), escala=2.5):
                resultado[nombre] = str(ruta_png)
                _logger.info(f"[OK] {nombre.upper()} guardado: {ruta_png}")
            else:
                _logger.error(f"[ERROR] Fallo capturando {nombre.upper()}")

            time.sleep(0.5)  # Pausa entre capturas

        resultado['exito'] = all(resultado[k] for k in ['general', 'zonal', 'supervisor'])

        if resultado['exito']:
            _logger.info(f"[OK] Todas las tablas capturadas para {etiqueta_hora}")
        else:
            _logger.error(f"[ERROR] Algunas tablas fallaron para {etiqueta_hora}")

        libro.close()
        app.quit()

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

        archivos = [
            ('GENERAL',    imagenes.get('general')),
            ('ZONAL',      imagenes.get('zonal')),
            ('SUPERVISOR', imagenes.get('supervisor')),
        ]

        for nombre_tabla, ruta_img in archivos:
            if not ruta_img or not Path(ruta_img).exists():
                _logger.warning(f"[WARN] Imagen no encontrada: {ruta_img}")
                continue

            caption = f"{msg_titulo} - {nombre_tabla}"
            _logger.info(f"[SEND] Enviando {nombre_tabla} a {numero_destino}...")

            resultado = wa.send_image(numero_destino, ruta_img, caption=caption)

            if not resultado.get('success'):
                _logger.error(f"[ERROR] Fallo enviando {nombre_tabla}: {resultado.get('error')}")
                return False

            _logger.info(f"[OK] {nombre_tabla} enviado")
            time.sleep(random.uniform(3, 6))  # delay antibang entre imágenes

        _logger.info(f"[OK] Corte {etiqueta_hora} enviado a {destino} ({numero_destino})")
        return True

    except Exception as e:
        _logger.error(f"[ERROR] Error enviando corte: {e}")
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

    # Buscar archivo Excel si no se especifica
    if not args.archivo:
        excels = sorted(BASE_DIR.glob("CORTE_VENTAS_*.xlsx"), reverse=True)
        if not excels:
            _logger.error("[ERROR] No se encontró CORTE_VENTAS_*.xlsx")
            return
        args.archivo = str(excels[0])
        _logger.info(f"[FILE] Usando: {args.archivo}")

    # Validar que existe
    if not Path(args.archivo).exists():
        _logger.error(f"[ERROR] Archivo no encontrado: {args.archivo}")
        return

    # Capturar tablas
    _logger.info(f"\n{'='*70}")
    _logger.info(f"[CAPTURE] CAPTURANDO CORTE {args.hora:02d}:00")
    _logger.info(f"{'='*70}\n")

    imagenes = capturar_tablas(args.archivo, args.hora)

    if not imagenes['exito']:
        _logger.error("[ERROR] Fallo en captura")
        return

    if args.solo_imagenes:
        _logger.info("[OK] Imágenes capturadas (--solo-imagenes activo)")
        return

    # Enviar a WhatsApp
    _logger.info(f"\n{'='*70}")
    _logger.info(f"[SEND] ENVIANDO A WHATSAPP")
    _logger.info(f"{'='*70}\n")

    config = cargar_config()

    if enviar_corte_whatsapp(imagenes, args.destino, args.hora, config):
        _logger.info(f"\n[OK] CORTE {args.hora:02d}:00 COMPLETADO")
    else:
        _logger.error(f"\n[ERROR] Fallo enviando corte")


if __name__ == "__main__":
    main()
