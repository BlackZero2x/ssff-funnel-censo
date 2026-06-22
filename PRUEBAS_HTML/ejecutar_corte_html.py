#!/usr/bin/env python3
"""
ejecutar_corte_html.py

Orquestador que ejecuta generar + capturar secuencialmente.

Uso:
    python ejecutar_corte_html.py --hora 10
"""

import argparse
import subprocess
import sys
import logging
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent
LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

# Logging
fecha_log = datetime.now().strftime("%Y%m%d")
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / f"corte_html_{fecha_log}.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def main():
    parser = argparse.ArgumentParser(description='Orquestador de cortes HTML')
    parser.add_argument('--hora', type=int, required=True, help='Hora (8-18)')
    parser.add_argument('--fecha', type=str, default=None, help='Fecha YYYY-MM-DD')
    args = parser.parse_args()

    hora = args.hora
    fecha_arg = f" --fecha {args.fecha}" if args.fecha else ""

    logger.info(f"{'='*70}")
    logger.info(f"[ORQUESTADOR] Corte HTML {hora:02d}:00")
    logger.info(f"{'='*70}")

    # [1] Generar HTML
    logger.info(f"\n[1/2] GENERANDO HTML...")
    cmd_generar = f"python {BASE_DIR}/generar_corte_html.py --hora {hora}{fecha_arg}"

    resultado1 = subprocess.run(cmd_generar, shell=True)
    if resultado1.returncode != 0:
        logger.error(f"[FATAL] Fallo en generación")
        sys.exit(1)

    logger.info(f"[OK] HTML generado")

    # [2] Capturar
    logger.info(f"\n[2/2] CAPTURANDO SCREENSHOT...")
    cmd_capturar = f"python {BASE_DIR}/capturar_cortes_html.py --hora {hora}{fecha_arg}"

    resultado2 = subprocess.run(cmd_capturar, shell=True)
    if resultado2.returncode != 0:
        logger.error(f"[WARN] Fallo en captura (revisa si Playwright está instalado)")
        logger.info(f"   Ejecuta: pip install playwright && playwright install chromium")
        sys.exit(1)

    logger.info(f"[OK] Screenshot capturado")

    logger.info(f"\n{'='*70}")
    logger.info(f"[OK] CORTE HTML COMPLETADO ✓")
    logger.info(f"{'='*70}\n")


if __name__ == "__main__":
    main()
