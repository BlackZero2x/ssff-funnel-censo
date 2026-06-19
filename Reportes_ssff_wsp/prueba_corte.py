#!/usr/bin/env python3
"""
prueba_corte.py

Genera un corte 8AM y envía prueba a tu número (+51975155264).

Uso:
    python prueba_corte.py

Este script:
1. Genera el Excel con generar_corte_ventas.py
2. Captura 3 tablas como PNG
3. Envía a tu WhatsApp personal para validar
"""

import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.absolute()

def run_cmd(cmd: str, descripcion: str) -> bool:
    """Ejecuta comando y retorna True si exitoso"""
    print(f"\n{'='*70}")
    print(f"▶️  {descripcion}")
    print(f"{'='*70}\n")

    resultado = subprocess.run(cmd, shell=True)
    exito = resultado.returncode == 0

    if exito:
        print(f"\n✅ {descripcion} — EXITOSO")
    else:
        print(f"\n❌ {descripcion} — FALLIDO (código {resultado.returncode})")

    return exito


def main():
    print("\n" + "="*70)
    print("🧪 PRUEBA DE CORTE SSFF 8AM")
    print("="*70)

    # Paso 1: Generar Excel
    if not run_cmd(
        f'cd "{SCRIPT_DIR}" && python generar_corte_ventas.py --hora 8',
        "Paso 1: Generar Excel (generar_corte_ventas.py --hora 8)"
    ):
        print("\n❌ Fallo generando Excel. Abortando.")
        return

    # Paso 2: Capturar imágenes y enviar
    if not run_cmd(
        f'cd "{SCRIPT_DIR}" && python capturar_cortes.py --hora 8 --destino test',
        "Paso 2: Capturar y enviar (capturar_cortes.py --hora 8 --destino test)"
    ):
        print("\n❌ Fallo capturando/enviando. Abortando.")
        return

    # Resumen
    print("\n" + "="*70)
    print("✅ PRUEBA COMPLETADA")
    print("="*70)
    print("\n📋 Verifica tu WhatsApp (+51975155264):")
    print("   • Deberías recibir 3 imágenes (GENERAL, ZONAL, SUPERVISOR)")
    print("   • Cada una con etiqueta 'CORTE 8AM'")
    print("   • Enviadas con delays aleatorios (3-8s) para evitar baneo")
    print("\n💡 Si recibiste las 3 imágenes:")
    print("   1. ✅ Sistema de captura funciona")
    print("   2. ✅ Integración WhatsApp funciona")
    print("   3. ✅ Ya puedes programar las 11 tareas automáticas")
    print("      → python programar_cortes.py")
    print("\n" + "="*70 + "\n")


if __name__ == "__main__":
    main()
