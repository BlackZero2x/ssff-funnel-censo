#!/usr/bin/env python3
"""
ejecutar_corte_wrapper.py

Script wrapper que ejecuta generar_corte_ventas.py y capturar_cortes.py
secuencialmente para una hora específica.

Uso: python ejecutar_corte_wrapper.py --hora 13
"""

import argparse
import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.absolute()
GENERAR = SCRIPT_DIR / "generar_corte_ventas.py"
CAPTURAR = SCRIPT_DIR / "capturar_cortes.py"

def main():
    parser = argparse.ArgumentParser(
        description="Ejecuta generar + capturar para una hora específica"
    )
    parser.add_argument('--hora', type=int, required=True, help='Hora (8-18)')
    args = parser.parse_args()

    hora = args.hora

    # Validar rango
    if hora < 8 or hora > 18:
        print(f"[ERROR] Hora fuera de rango: {hora}. Debe estar entre 8 y 18.")
        sys.exit(1)

    print(f"\n{'='*70}")
    print(f"[CORTE] Ejecutando corte {hora:02d}:00")
    print(f"{'='*70}\n")

    # [1] Generar Excel
    print(f"[1/2] Generando Excel...")
    cmd_generar = [
        "uv", "run", "python", str(GENERAR),
        "--hora", str(hora), "--solo-excel"
    ]

    resultado1 = subprocess.run(cmd_generar, cwd=SCRIPT_DIR)
    if resultado1.returncode != 0:
        print(f"\n[ERROR] generar_corte_ventas.py falló (código {resultado1.returncode})")
        sys.exit(1)

    # [2] Capturar y enviar
    print(f"\n[2/2] Capturando y enviando a WhatsApp...")
    cmd_capturar = [
        "uv", "run", "python", str(CAPTURAR),
        "--hora", str(hora), "--destino", "canal"
    ]

    resultado2 = subprocess.run(cmd_capturar, cwd=SCRIPT_DIR)
    if resultado2.returncode != 0:
        print(f"\n[ERROR] capturar_cortes.py falló (código {resultado2.returncode})")
        sys.exit(1)

    print(f"\n{'='*70}")
    print(f"[OK] Corte {hora:02d}:00 completado")
    print(f"{'='*70}\n")


if __name__ == "__main__":
    main()
