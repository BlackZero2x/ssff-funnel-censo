#!/usr/bin/env python3
"""
ejecutar_funnel_wrapper.py

Wrapper usado por el Task Scheduler. Ejecuta generar_funnel_preventa.py y
capturar_funnel.py secuencialmente en los cortes: 9AM, 1PM, 4PM y 5:30PM (cierre).

Uso: python ejecutar_funnel_wrapper.py --corte 9AM
"""

import argparse
import logging
import subprocess
import sys
import time
from datetime import date, datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.absolute()
GENERAR    = SCRIPT_DIR / "generar_funnel_preventa.py"
CAPTURAR   = SCRIPT_DIR / "capturar_funnel.py"
LOG_DIR    = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

TIMEOUT_GENERAR    = 180  # 3 minutos
TIMEOUT_CAPTURAR   = 120  # 2 minutos
ESPERA_POST_GENERAR = 20  # segundos — margen para que el SO/antivirus suelte el .xlsx recién escrito

CORTES_VALIDOS = {'9AM', '1PM', '4PM', '5:30PM'}

# Fechas sin envio de funnel (feriados, mantenimiento, etc.)
FECHAS_SIN_CORTE = {date(2026, 7, 29)}


def setup_logger(corte: str) -> logging.Logger:
    fecha_str = datetime.now().strftime("%Y%m%d")
    log_file  = LOG_DIR / f"funnel_{fecha_str}.log"

    logger = logging.getLogger("funnel_wrapper")
    logger.setLevel(logging.DEBUG)
    logger.handlers = []

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


def main():
    parser = argparse.ArgumentParser(description="Ejecuta generar + capturar del Funnel Preventa para un corte")
    parser.add_argument('--corte', required=True, help="Etiqueta del corte (9AM, 11AM, 1PM, 3PM, 5PM)")
    parser.add_argument('--destino', default='test',
                         help="Destino en config.json (default: 'test' — número personal, nunca canal por defecto)")
    args   = parser.parse_args()
    corte  = args.corte
    logger = setup_logger(corte)

    if corte not in CORTES_VALIDOS:
        logger.error(f"[ERROR] Corte inválido: {corte}. Debe ser uno de {sorted(CORTES_VALIDOS)}")
        sys.exit(1)

    if date.today() in FECHAS_SIN_CORTE:
        logger.info(f"[SKIP] {date.today()} está en FECHAS_SIN_CORTE — no se genera ni envía el funnel {corte}")
        return

    logger.info(f"{'='*70}")
    logger.info(f"[FUNNEL] Ejecutando funnel preventa {corte}")
    logger.info(f"{'='*70}")

    python = sys.executable

    def _run(cmd, nombre, timeout):
        r = subprocess.run(cmd, cwd=SCRIPT_DIR, timeout=timeout,
                           stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                           text=True, encoding='utf-8', errors='replace')
        if r.stdout:
            for line in r.stdout.strip().splitlines():
                logger.info(f"  [{nombre}] {line}")
        if r.returncode != 0 and r.stderr:
            for line in r.stderr.strip().splitlines():
                logger.error(f"  [{nombre}] {line}")
        return r

    logger.info(f"[1/2] Generando Excel... (timeout={TIMEOUT_GENERAR}s)")
    try:
        r1 = _run([python, str(GENERAR), "--corte", corte], "GENERAR", TIMEOUT_GENERAR)
    except subprocess.TimeoutExpired:
        logger.error(f"[TIMEOUT] generar_funnel_preventa.py superó {TIMEOUT_GENERAR}s")
        sys.exit(1)
    except Exception as e:
        logger.error(f"[EXCEPTION] generar_funnel_preventa.py: {e}")
        sys.exit(1)

    if r1.returncode != 0:
        logger.error(f"[ERROR] generar_funnel_preventa.py falló (código {r1.returncode})")
        sys.exit(1)
    logger.info("[OK] Excel generado correctamente")

    # Margen antes de que Excel/COM intente abrir el archivo recién escrito —
    # el antivirus/Windows Defender puede retener brevemente el acceso exclusivo
    # a un .xlsx recién modificado, aunque ya no haya ningún proceso reteniéndolo.
    logger.info(f"[WAIT] Esperando {ESPERA_POST_GENERAR}s antes de capturar...")
    time.sleep(ESPERA_POST_GENERAR)

    logger.info(f"[2/2] Capturando y enviando a WhatsApp... (timeout={TIMEOUT_CAPTURAR}s)")
    try:
        r2 = _run([python, str(CAPTURAR), "--corte", corte, "--destino", args.destino],
                  "CAPTURAR", TIMEOUT_CAPTURAR)
    except subprocess.TimeoutExpired:
        logger.error(f"[TIMEOUT] capturar_funnel.py superó {TIMEOUT_CAPTURAR}s")
        sys.exit(1)
    except Exception as e:
        logger.error(f"[EXCEPTION] capturar_funnel.py: {e}")
        sys.exit(1)

    if r2.returncode != 0:
        logger.error(f"[ERROR] capturar_funnel.py falló (código {r2.returncode})")
        sys.exit(1)
    logger.info("[OK] Imagen capturada y enviada a WhatsApp")

    logger.info(f"{'='*70}")
    logger.info(f"[OK] Funnel {corte} completado exitosamente")
    logger.info(f"{'='*70}")


if __name__ == "__main__":
    main()
