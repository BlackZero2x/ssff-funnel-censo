#!/usr/bin/env python3
"""
ejecutar_corte_wrapper.py

Wrapper usado por el Task Scheduler. Ejecuta generar_corte_ventas.py y
capturar_cortes.py secuencialmente con timeout y logging a archivo diario.

Uso: python ejecutar_corte_wrapper.py --hora 13
"""

import argparse
import logging
import subprocess
import sys
import time
from datetime import date, datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.absolute()
GENERAR    = SCRIPT_DIR / "generar_corte_ventas.py"
CAPTURAR   = SCRIPT_DIR / "capturar_cortes.py"
LOG_DIR    = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

TIMEOUT_GENERAR  = 120   # 2 minutos
TIMEOUT_CAPTURAR = 480   # 8 minutos — canal + grupos individuales + alertas

REINTENTOS_DATOS_VACIOS = 3    # máximo de reintentos si los datos están en cero
ESPERA_REINTENTO        = 300  # 5 minutos entre reintentos

# Fechas sin envio de cortes (feriados, mantenimiento, etc.)
FECHAS_SIN_CORTE = {date(2026, 7, 29)}


def setup_logger(hora: int) -> logging.Logger:
    fecha_str = datetime.now().strftime("%Y%m%d")
    log_file  = LOG_DIR / f"cortes_{fecha_str}.log"

    logger = logging.getLogger("wrapper")
    logger.setLevel(logging.DEBUG)
    logger.handlers = []

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
    return logger


def main():
    parser = argparse.ArgumentParser(
        description="Ejecuta generar + capturar para una hora específica"
    )
    parser.add_argument('--hora', type=int, required=True, help='Hora (8-18)')
    args   = parser.parse_args()
    hora   = args.hora
    logger = setup_logger(hora)

    if hora < 8 or hora > 18:
        logger.error(f"[ERROR] Hora fuera de rango: {hora}. Debe estar entre 8 y 18.")
        sys.exit(1)

    if date.today() in FECHAS_SIN_CORTE:
        logger.info(f"[SKIP] {date.today()} está en FECHAS_SIN_CORTE — no se genera ni envía el corte {hora:02d}:00")
        return

    logger.info(f"{'='*70}")
    logger.info(f"[CORTE] Ejecutando corte {hora:02d}:00")
    logger.info(f"{'='*70}")

    python = sys.executable

    def _run(cmd, nombre, timeout):
        """Ejecuta un subproceso y vuelca su stderr al log si falla."""
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

    # [1] Generar Excel (con reintentos si datos están vacíos por limpieza de caché)
    for intento in range(1, REINTENTOS_DATOS_VACIOS + 2):
        logger.info(f"[1/2] Generando Excel... intento {intento} (timeout={TIMEOUT_GENERAR}s)")
        try:
            r1 = _run(
                [python, str(GENERAR), "--hora", str(hora), "--solo-excel"],
                "GENERAR", TIMEOUT_GENERAR
            )
        except subprocess.TimeoutExpired:
            logger.error(f"[TIMEOUT] generar_corte_ventas.py superó {TIMEOUT_GENERAR}s")
            sys.exit(1)
        except Exception as e:
            logger.error(f"[EXCEPTION] generar_corte_ventas.py: {e}")
            sys.exit(1)

        if r1.returncode == 0:
            logger.info(f"[OK] Excel generado correctamente")
            break
        elif r1.returncode == 2:
            # Datos vacíos — posible limpieza de caché del Data Center
            if intento <= REINTENTOS_DATOS_VACIOS:
                logger.warning(
                    f"[DATOS_VACIOS] Datos insuficientes en intento {intento}/{REINTENTOS_DATOS_VACIOS} "
                    f"— reintentando en {ESPERA_REINTENTO//60} min..."
                )
                time.sleep(ESPERA_REINTENTO)
            else:
                logger.error(
                    f"[DATOS_VACIOS] {REINTENTOS_DATOS_VACIOS} reintentos agotados — "
                    f"se abandona el corte {hora:02d}:00"
                )
                sys.exit(2)
        else:
            logger.error(f"[ERROR] generar_corte_ventas.py falló (código {r1.returncode})")
            sys.exit(1)

    # [2] Capturar y enviar
    logger.info(f"[2/2] Capturando y enviando a WhatsApp... (timeout={TIMEOUT_CAPTURAR}s)")
    try:
        r2 = _run(
            [python, str(CAPTURAR), "--hora", str(hora), "--destino", "canal"],
            "CAPTURAR", TIMEOUT_CAPTURAR
        )
        if r2.returncode != 0:
            logger.error(f"[ERROR] capturar_cortes.py falló (código {r2.returncode})")
            sys.exit(1)
        logger.info(f"[OK] Imágenes capturadas y enviadas a WhatsApp")
    except subprocess.TimeoutExpired:
        logger.error(f"[TIMEOUT] capturar_cortes.py superó {TIMEOUT_CAPTURAR}s — posible cuelgue en envío WA")
        sys.exit(1)
    except Exception as e:
        logger.error(f"[EXCEPTION] capturar_cortes.py: {e}")
        sys.exit(1)

    logger.info(f"{'='*70}")
    logger.info(f"[OK] Corte {hora:02d}:00 completado exitosamente")
    logger.info(f"{'='*70}")


if __name__ == "__main__":
    main()
