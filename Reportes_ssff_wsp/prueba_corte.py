#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
prueba_corte.py

Genera un corte 8AM y envia prueba a tu numero (+51975155264).

Uso:
    python prueba_corte.py

Este script:
1. Genera el Excel con generar_corte_ventas.py
2. Captura 3 tablas como PNG
3. Envia a tu WhatsApp personal para validar
"""

import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.absolute()

def run_cmd(cmd: str, descripcion: str) -> bool:
    """Ejecuta comando y retorna True si exitoso"""
    print(f"\n{'='*70}")
    print(f"[>] {descripcion}")
    print(f"{'='*70}\n")

    resultado = subprocess.run(cmd, shell=True)
    exito = resultado.returncode == 0

    if exito:
        print(f"\n[OK] {descripcion} — EXITOSO")
    else:
        print(f"\n[ERROR] {descripcion} — FALLIDO (codigo {resultado.returncode})")

    return exito


def main():
    print("\n" + "="*70)
    print("[TEST] PRUEBA DE CORTE SSFF 8AM")
    print("="*70)

    # Paso 1: Generar Excel
    if not run_cmd(
        f'cd "{SCRIPT_DIR}" && python generar_corte_ventas.py --hora 8',
        "Paso 1: Generar Excel (generar_corte_ventas.py --hora 8)"
    ):
        print("\n[ERROR] Fallo generando Excel. Abortando.")
        return

    # Paso 2: Capturar imagenes y enviar
    if not run_cmd(
        f'cd "{SCRIPT_DIR}" && python capturar_cortes.py --hora 8 --destino test',
        "Paso 2: Capturar y enviar (capturar_cortes.py --hora 8 --destino test)"
    ):
        print("\n[ERROR] Fallo capturando/enviando. Abortando.")
        return

    # Resumen
    print("\n" + "="*70)
    print("[OK] PRUEBA COMPLETADA")
    print("="*70)
    print("\nVerifica tu WhatsApp (+51975155264):")
    print("   - Deberias recibir 3 imagenes (GENERAL, ZONAL, SUPERVISOR)")
    print("   - Cada una con etiqueta 'CORTE 8AM'")
    print("   - Enviadas con delays aleatorios (3-8s) para evitar baneo")
    print("\nSi recibiste las 3 imagenes:")
    print("   1. [OK] Sistema de captura funciona")
    print("   2. [OK] Integracion WhatsApp funciona")
    print("   3. [OK] Ya puedes programar las 11 tareas automaticas")
    print("      -> python programar_cortes.py")
    print("\n" + "="*70 + "\n")


if __name__ == "__main__":
    main()
