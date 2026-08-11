#!/usr/bin/env python3
"""
programar_funnel.py

Programa ejecuciones automáticas del Funnel Preventa en Windows Task Scheduler.
Ejecuta: python programar_funnel.py (requiere permisos de Administrador)

Crea tareas para los 4 cortes definidos:
  9AM, 1PM, 4PM y 5:30PM (cierre del día)

Flujo por corte:
  [ejecutar_funnel_wrapper.py --corte X --destino canal]
    → genera_funnel_preventa.py (Excel con hoja RESUMEN)
    → capturar_funnel.py (imagen RESUMEN + archivo .xlsx a WhatsApp)
"""

import ctypes
import subprocess
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.absolute()
SCRIPT_WRAPPER = SCRIPT_DIR / "ejecutar_funnel_wrapper.py"

# hora:minuto (24h) -> etiqueta del corte
HORARIOS = {
    (9, 0):  "9AM",
    (13, 0): "1PM",
    (16, 0): "4PM",
    (17, 30): "5:30PM",
}


def _es_admin() -> bool:
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def crear_tarea_windows(hora: int, minuto: int, etiqueta: str) -> bool:
    """Crea una tarea en Windows Task Scheduler para un corte del Funnel Preventa."""
    slug = etiqueta.replace(':', '').replace(' ', '')
    nombre_tarea = f"SSFF_Funnel_{slug}"

    python_exe = "C:\\proyectos\\.venv\\Scripts\\python.exe"
    argumentos = f'"{SCRIPT_WRAPPER}" --corte {etiqueta} --destino canal'

    schtasks_cmd = [
        "schtasks",
        "/create",
        "/tn", nombre_tarea,
        "/tr", f'"{python_exe}" {argumentos}',
        "/sc", "weekly",
        "/d", "MON,TUE,WED,THU,FRI,SAT",
        "/st", f"{hora:02d}:{minuto:02d}:00",
        "/ru", "SYSTEM",
        "/f",
    ]

    print(f"\nCreando tarea: {nombre_tarea}")
    print(f"   Hora: {hora:02d}:{minuto:02d} ({etiqueta})")
    print(f"   Argumentos: {argumentos}")

    try:
        resultado = subprocess.run(schtasks_cmd, capture_output=True, text=True, check=False)
        if resultado.returncode == 0:
            print(f"   [OK] Tarea creada")
            return True
        print(f"   [ERROR] {resultado.stderr.strip()}")
        return False
    except Exception as e:
        print(f"   [EXCEPTION] {e}")
        return False


def listar_tareas_existentes():
    print("\nTareas SSFF_Funnel existentes:\n")
    try:
        resultado = subprocess.run(["schtasks", "/query", "/fo", "list"],
                                   capture_output=True, text=True, check=False)
        for linea in resultado.stdout.split("\n"):
            if "SSFF_Funnel" in linea:
                print(f"   {linea.strip()}")
    except Exception as e:
        print(f"   [ERROR] {e}")


def eliminar_tareas_existentes():
    print("\nEliminando tareas SSFF_Funnel anteriores...\n")
    for (hora, minuto), etiqueta in HORARIOS.items():
        slug = etiqueta.replace(':', '').replace(' ', '')
        nombre_tarea = f"SSFF_Funnel_{slug}"
        resultado = subprocess.run(["schtasks", "/delete", "/tn", nombre_tarea, "/f"],
                                   capture_output=True, text=True, check=False)
        if resultado.returncode == 0:
            print(f"   [OK] Eliminada: {nombre_tarea}")


def main():
    print("=" * 70)
    print("PROGRAMADOR DE CORTES — FUNNEL PREVENTA")
    print("=" * 70)

    if not _es_admin():
        print("\n[ADVERTENCIA] Se requieren permisos de Administrador")
        print("   Ejecuta nuevamente como Admin (PowerShell/CMD: Run as Administrator)")
        return

    if not SCRIPT_WRAPPER.exists():
        print(f"\n[ERROR] No encontrado: {SCRIPT_WRAPPER}")
        return

    print(f"\n[OK] Script encontrado: {SCRIPT_WRAPPER}")

    listar_tareas_existentes()

    print("\n¿Eliminar tareas SSFF_Funnel anteriores? (s/n): ", end="")
    if input().lower() == "s":
        eliminar_tareas_existentes()

    print("\n" + "=" * 70)
    print("CREANDO NUEVAS TAREAS")
    print("=" * 70)

    exitos = fallos = 0
    for (hora, minuto), etiqueta in HORARIOS.items():
        if crear_tarea_windows(hora, minuto, etiqueta):
            exitos += 1
        else:
            fallos += 1

    print("\n" + "=" * 70)
    print(f"RESUMEN: {exitos} OK — {fallos} fallos")
    print("=" * 70)


if __name__ == "__main__":
    main()
