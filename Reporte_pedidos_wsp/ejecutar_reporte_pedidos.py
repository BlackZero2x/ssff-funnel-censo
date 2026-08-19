"""
ejecutar_reporte_pedidos.py

Orquestador: genera el Excel de pedidos y luego lo captura y envía a WhatsApp.
Programado por Windows Task Scheduler con programar_reporte_pedidos.py.

Horarios de producción: 9:20, 11:20, 13:40, 15:20, 17:20
La etiqueta del corte es la hora entera (9:20 → hora=9 → "9AM")

Uso:
    python ejecutar_reporte_pedidos.py --hora 9
"""

import argparse
import subprocess
import sys
import logging
from pathlib import Path
from datetime import date, datetime

SCRIPT_DIR = Path(__file__).parent.absolute()
GENERAR    = SCRIPT_DIR / 'generar_reporte_pedidos.py'
CAPTURAR   = SCRIPT_DIR / 'capturar_reporte_pedidos.py'
LOG_DIR    = SCRIPT_DIR / 'logs'
LOG_DIR.mkdir(exist_ok=True)

# Fechas sin envio de reporte (feriados, mantenimiento, etc.)
FECHAS_SIN_CORTE = {date(2026, 7, 29)}

# Etiquetas de hora (mismas que generar_reporte_pedidos.py)
def _etiqueta_hora(hora: int) -> str:
    if hora < 12:
        return f'{hora}AM'
    elif hora == 12:
        return '12PM'
    else:
        return f'{hora - 12}PM'


def setup_logger(hora: int) -> logging.Logger:
    fecha_str = datetime.now().strftime('%Y%m%d')
    log_file  = LOG_DIR / f'pedidos_{fecha_str}.log'

    logger = logging.getLogger('ejecutar_pedidos')
    logger.setLevel(logging.DEBUG)
    logger.handlers = []

    fh = logging.FileHandler(log_file, encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S'
    ))

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter('%(message)s'))

    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger


def main():
    parser = argparse.ArgumentParser(description='Orquestador de reporte de pedidos')
    parser.add_argument('--hora', type=int, required=True, help='Hora del corte (0-23)')
    args = parser.parse_args()

    hora   = args.hora
    logger = setup_logger(hora)
    etiq   = _etiqueta_hora(hora)

    if hora < 0 or hora > 23:
        logger.error(f'Hora fuera de rango: {hora}')
        sys.exit(1)

    if date.today() in FECHAS_SIN_CORTE:
        logger.info(f'[SKIP] {date.today()} está en FECHAS_SIN_CORTE — no se genera ni envía el reporte {etiq}')
        return

    logger.info(f'{"="*70}')
    logger.info(f'[PEDIDOS] Corte {etiq} — {datetime.now().strftime("%d/%m/%Y %H:%M")}')
    logger.info(f'{"="*70}')

    def _run(cmd, nombre):
        r = subprocess.run(cmd, cwd=SCRIPT_DIR,
                           stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                           encoding='utf-8', errors='replace')
        for line in (r.stdout or '').splitlines():
            logger.info(f'  {line}')
        if r.returncode != 0:
            logger.error(f'[FATAL] {nombre} falló (código {r.returncode})')
            sys.exit(1)
        return r

    # [1] Generar Excel
    logger.info(f'\n[1/2] GENERANDO EXCEL...')
    _run([sys.executable, str(GENERAR), '--hora', str(hora)], 'generar_reporte_pedidos.py')
    logger.info(f'[OK] Excel generado')

    # [2] Capturar y enviar
    logger.info(f'\n[2/2] CAPTURANDO Y ENVIANDO...')
    _run([sys.executable, str(CAPTURAR), '--hora', str(hora), '--destino', 'canal'],
         'capturar_reporte_pedidos.py')

    logger.info(f'\n{"="*70}')
    logger.info(f'[OK] CORTE PEDIDOS {etiq} COMPLETADO')
    logger.info(f'{"="*70}\n')


if __name__ == '__main__':
    main()
