"""
programar_reporte_pedidos.py

Crea tareas en Windows Task Scheduler para los cortes de pedidos.

Horarios: 9:20, 11:20, 13:40, 15:20, 17:20
La tarea llama a ejecutar_reporte_pedidos.py --hora N (hora entera).

Requiere ejecutar como Administrador.
"""

import ctypes
import subprocess
from pathlib import Path

SCRIPT_DIR  = Path(__file__).parent.absolute()
PYTHON_EXE  = Path('C:/proyectos/.venv/Scripts/python.exe')
ORQUESTADOR = SCRIPT_DIR / 'ejecutar_reporte_pedidos.py'
USUARIO     = 'AUREN\\developer7'

# (hora_inicio, minuto_inicio, hora_argumento, etiqueta)
HORARIOS = [
    (9,  20, 9,  '9AM'),
    (11, 20, 11, '11AM'),
    (13, 40, 13, '1PM'),
    (15, 20, 15, '3PM'),
    (17, 20, 17, '5PM'),
]


def _es_admin() -> bool:
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def crear_tarea(hora_ini: int, minuto_ini: int, hora_arg: int, etiqueta: str) -> bool:
    nombre     = f'SSFF_Pedidos_{etiqueta}'
    hora_str   = f'{hora_ini:02d}:{minuto_ini:02d}:00'
    argumentos = f'"{ORQUESTADOR}" --hora {hora_arg}'

    cmd = [
        'schtasks', '/create',
        '/tn', nombre,
        '/tr', f'"{PYTHON_EXE}" {argumentos}',
        '/sc', 'weekly',
        '/d', 'MON,TUE,WED,THU,FRI,SAT',
        '/st', hora_str,
        '/ru', USUARIO,
        '/sd', '19/06/2026',
        '/f',
    ]

    print(f'\n  Creando tarea: {nombre}')
    print(f'     Hora: {hora_ini:02d}:{minuto_ini:02d} → etiqueta {etiqueta}')
    print(f'     Programa: {PYTHON_EXE}')
    print(f'     Argumentos: {argumentos}')
    print(f'     Usuario: {USUARIO}')

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, check=False)
        if result.returncode == 0:
            print(f'     OK - Tarea creada')
            return True
        else:
            print(f'     ERROR: {result.stderr.strip()}')
            return False
    except Exception as e:
        print(f'     Excepción: {e}')
        return False


def eliminar_tareas():
    print('\n[DELETE] Eliminando tareas anteriores...')
    for _, _, _, etiqueta in HORARIOS:
        nombre = f'SSFF_Pedidos_{etiqueta}'
        result = subprocess.run(
            ['schtasks', '/delete', '/tn', nombre, '/f'],
            capture_output=True, text=True, check=False
        )
        if result.returncode == 0:
            print(f'  Eliminada: {nombre}')


def listar_tareas():
    print('\n[LIST] Tareas SSFF_Pedidos existentes:\n')
    result = subprocess.run(
        ['schtasks', '/query', '/fo', 'list'],
        capture_output=True, text=True, check=False
    )
    encontradas = [l.strip() for l in result.stdout.split('\n') if 'SSFF_Pedidos' in l]
    if encontradas:
        for linea in encontradas:
            print(f'   • {linea}')
    else:
        print('   (ninguna)')


def main():
    print('=' * 70)
    print('PROGRAMADOR DE REPORTE PEDIDOS SSFF')
    print('=' * 70)

    if not _es_admin():
        print('\nADVERTENCIA: Se requieren permisos de Administrador')
        print('   Ejecuta nuevamente como Admin:')
        print('   • PowerShell: right-click → Run as Administrator')
        print('   • CMD: right-click → Run as Administrator')
        return

    if not ORQUESTADOR.exists():
        print(f'\nERROR: No encontrado: {ORQUESTADOR}')
        print('   Verifica que estás en el directorio correcto')
        return

    print(f'\n   Script encontrado: {ORQUESTADOR}')
    print(f'   Directorio: {SCRIPT_DIR}')
    print(f'   Usuario de tarea: {USUARIO}')

    listar_tareas()

    resp = input('\n¿Eliminar tareas anteriores? (s/n): ').lower()
    if resp == 's':
        eliminar_tareas()

    print('\n' + '=' * 70)
    print('CREANDO NUEVAS TAREAS')
    print('=' * 70)

    exitos = sum(crear_tarea(h, m, a, e) for h, m, a, e in HORARIOS)
    fallos = len(HORARIOS) - exitos

    print(f'\n{"="*70}')
    print(f'RESUMEN: {exitos} OK — {fallos} errores')
    print('=' * 70)

    if fallos == 0:
        print('\nTODAS LAS TAREAS CREADAS EXITOSAMENTE')
        print('\nEjecución automática (lunes a sábado):')
        for h, m, _, e in HORARIOS:
            print(f'   • {h:02d}:{m:02d} → {e}')
        print('\nVerifica en Task Scheduler:')
        print('   Control Panel → Administrative Tools → Task Scheduler')
        print('   Busca: SSFF_Pedidos_*')
    else:
        print(f'\nSe crearon {exitos} tareas pero {fallos} fallaron')
        print('   Verifica permisos y reintenta')

    print('=' * 70)


if __name__ == '__main__':
    main()
