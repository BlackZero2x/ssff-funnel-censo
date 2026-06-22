#!/usr/bin/env python3
"""
capturar_cortes_html.py

Captura screenshot de los HTMLs generados usando Playwright.

Uso:
    python capturar_cortes_html.py --hora 10
"""

import argparse
import datetime
import asyncio
from pathlib import Path
import logging

try:
    from playwright.async_api import async_playwright
except ImportError:
    print("Error: Playwright no instalado.")
    print("Ejecuta: pip install playwright && playwright install chromium")
    exit(1)

# Configuración
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"
CORTES_HTML_DIR = OUTPUT_DIR / "cortes_html"
CORTES_PNG_DIR = OUTPUT_DIR / "cortes_png"
LOG_DIR = BASE_DIR / "logs"

CORTES_PNG_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)

# Logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / f"capturador_{datetime.date.today().strftime('%Y%m%d')}.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


async def capturar_screenshot(html_file, output_file, table_name):
    """Captura screenshot de un HTML usando Playwright."""
    logger.info(f"Capturando {table_name}...")

    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()

        # Abrir el archivo local
        page_url = f"file:///{html_file.resolve().as_posix()}"
        await page.goto(page_url, wait_until="networkidle")

        # Esperar a que el contenido se renderice
        await asyncio.sleep(1)

        # Capturar screenshot completo
        await page.screenshot(path=output_file, full_page=True)

        await browser.close()

    logger.info(f"   ✓ Guardado: {output_file}")


async def main():
    parser = argparse.ArgumentParser(description='Captura screenshots de cortes HTML')
    parser.add_argument('--hora', type=int, required=True, help='Hora (8-18)')
    parser.add_argument('--fecha', type=str, default=None, help='Fecha YYYY-MM-DD')
    args = parser.parse_args()

    # Resolver fecha
    if args.fecha:
        hoy = datetime.datetime.strptime(args.fecha, '%Y-%m-%d').date()
    else:
        hoy = datetime.date.today()

    hora = args.hora
    hora_lbl = {8: '8am', 9: '9am', 10: '10am', 11: '11am', 12: '12pm',
                13: '1pm', 14: '2pm', 15: '3pm', 16: '4pm', 17: '5pm', 18: '6pm'}.get(hora, f'{hora}h')

    print(f"\n{'='*70}")
    print(f"  CAPTURAR CORTES HTML — {hoy.strftime('%d/%m/%Y')} — {hora_lbl.upper()}")
    print(f"{'='*70}")

    # Buscar archivo HTML
    html_file = CORTES_HTML_DIR / f"corte_{hoy.strftime('%Y%m%d')}_{hora_lbl}.html"

    if not html_file.exists():
        logger.error(f"❌ No encontrado: {html_file}")
        logger.error(f"   Ejecuta primero: python generar_corte_html.py --hora {hora}")
        return

    logger.info(f"✓ HTML encontrado: {html_file}")

    # Capturar screenshot
    timestamp = datetime.datetime.now().strftime('%H%M%S')
    output_file = CORTES_PNG_DIR / f"CORTE_HTML_{hora_lbl.upper()}_{hoy.strftime('%Y%m%d')}_{timestamp}.png"

    try:
        await capturar_screenshot(html_file, output_file, f"corte HTML {hora_lbl}")
        logger.info(f"\n✅ Captura completada")
    except Exception as e:
        logger.error(f"❌ Error capturando: {e}")


if __name__ == "__main__":
    asyncio.run(main())
