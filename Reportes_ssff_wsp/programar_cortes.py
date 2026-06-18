#!/usr/bin/env python3
"""
programar_cortes.py

Programa ejecuciones automáticas de cortes horarios en Windows Task Scheduler.
Ejecuta: python programar_cortes.py

Crea tareas programadas para:
  8 AM  → Corte 8AM
  10 AM → Corte 10AM
  12 PM → Corte 12PM
  2 PM  → Corte 2PM
  4 PM  → Corte 4PM
  6 PM  → Corte 6PM

(Sin incluir 9AM, 11AM, 1PM, 3PM, 5PM para reducir volumen de envíos)
"""

import subprocess
import os
from pathlib import Path

# Directorio actual
SCRIPT_DIR = Path(__file__).parent.absolute()
SCRIPT_GENERAR = SCRIPT_DIR / "generar_corte_ventas.py"

# Horarios a programar (hora en formato 24h)
HORARIOS = {
    8: "8AM",
    10: "10AM",
    12: "12PM",
    14: "2PM",
    16: "4PM",
    18: "6PM",
}

def crear_tarea_windows(hora: int, etiqueta: str):
    """
    Crea una tarea en Windows Task Scheduler.

    Args:
        hora: Hora en formato 24h (8, 10, 12, 14, 16, 18)
        etiqueta: Etiqueta amigable (8AM, 10AM, etc.)
    """

    # Nombre de la tarea
    nombre_tarea = f"SSFF_Corte_{etiqueta}"

    # Comando a ejecutar
    cmd = f'python "{SCRIPT_GENERAR}" --hora {hora}'

    # Crear tarea con SCHTASKS
    schtasks_cmd = [
        "schtasks",
        "/create",
        "/tn", nombre_tarea,
        "/tr", cmd,
        "/sc", "daily",
        "/st", f"{hora:02d}:00:00",
        "/ru", "SYSTEM",
        "/f"  # Force (sobrescribe si existe)
    ]

    print(f"\n📌 Creando tarea: {nombre_tarea}")
    print(f"   Hora: {etiqueta}")
    print(f"   Comando: {cmd}")

    try:
        resultado = subprocess.run(
            schtasks_cmd,
            capture_output=True,
            text=True,
            check=False
        )

        if resultado.returncode == 0:
            print(f"   ✅ Tarea creada")
        else:
            print(f"   ⚠️  Error: {resultado.stderr}")
            return False

    except Exception as e:
        print(f"   ❌ Excepción: {e}")
        return False

    return True


def listar_tareas_existentes():
    """Lista tareas SSFF existentes"""
    print("\n📋 Tareas SSFF existentes:\n")

    try:
        resultado = subprocess.run(
            ["schtasks", "/query", "/fo", "list"],
            capture_output=True,
            text=True,
            check=False
        )

        for linea in resultado.stdout.split("\n"):
            if "SSFF_Corte" in linea:
                print(f"   • {linea.strip()}")

    except Exception as e:
        print(f"   ❌ Error listando tareas: {e}")


def eliminar_tareas_existentes():
    """Elimina todas las tareas SSFF existentes"""
    print("\n🗑️  Eliminando tareas SSFF anteriores...\n")

    for etiqueta in HORARIOS.values():
        nombre_tarea = f"SSFF_Corte_{etiqueta}"

        try:
            resultado = subprocess.run(
                ["schtasks", "/delete", "/tn", nombre_tarea, "/f"],
                capture_output=True,
                text=True,
                check=False
            )

            if resultado.returncode == 0:
                print(f"   ✅ Eliminada: {nombre_tarea}")
            # No mostrar error si no existe

        except Exception as e:
            print(f"   ⚠️  {nombre_tarea}: {e}")


def main():
    print("=" * 70)
    print("🕐 PROGRAMADOR DE CORTES SSFF")
    print("=" * 70)

    # Verificar permisos de admin
    if not _es_admin():
        print("\n⚠️  ADVERTENCIA: Se requieren permisos de Administrador")
        print("   Ejecuta nuevamente como Admin:")
        print("   • PowerShell: right-click → Run as Administrator")
        print("   • CMD: right-click → Run as Administrator")
        return

    # Verificar que el script existe
    if not SCRIPT_GENERAR.exists():
        print(f"\n❌ No encontrado: {SCRIPT_GENERAR}")
        print("   Verifica que estás en el directorio correcto")
        return

    print(f"\n✅ Script encontrado: {SCRIPT_GENERAR}")
    print(f"   Directorio: {SCRIPT_DIR}")

    # Listar tareas existentes
    listar_tareas_existentes()

    # Preguntar si eliminar anteriores
    print("\n¿Eliminar tareas SSFF anteriores? (s/n): ", end="")
    respuesta = input().lower()

    if respuesta == "s":
        eliminar_tareas_existentes()

    # Crear nuevas tareas
    print("\n" + "=" * 70)
    print("📅 CREANDO NUEVAS TAREAS")
    print("=" * 70)

    exitos = 0
    fallos = 0

    for hora, etiqueta in HORARIOS.items():
        if crear_tarea_windows(hora, etiqueta):
            exitos += 1
        else:
            fallos += 1

    # Resumen
    print("\n" + "=" * 70)
    print(f"📊 RESUMEN: {exitos} ✅ — {fallos} ❌")
    print("=" * 70)

    if fallos == 0:
        print("\n✅ TODAS LAS TAREAS CREADAS EXITOSAMENTE")
        print("\nEjecución automática:")
        for etiqueta in HORARIOS.values():
            print(f"   • {etiqueta}")
        print("\nVerifica en Task Scheduler:")
        print("   Control Panel → Administrative Tools → Task Scheduler")
        print("   Busca: SSFF_Corte_*")
    else:
        print(f"\n⚠️  Se crearon {exitos} tareas pero {fallos} fallaron")
        print("   Verifica permisos y reintenra")

    print("\n" + "=" * 70)


def _es_admin():
    """Verifica si el script se ejecuta como Admin"""
    try:
        import ctypes
        return ctypes.windll.shell.IsUserAnAdmin()
    except:
        return False


if __name__ == "__main__":
    main()
