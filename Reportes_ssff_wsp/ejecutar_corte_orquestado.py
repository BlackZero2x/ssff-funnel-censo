#!/usr/bin/env python3
"""
ejecutar_corte_orquestado.py

Wrapper mejorado que ejecuta generar_corte_ventas.py y capturar_cortes.py
con manejo robusto de errores y reintentos inteligentes.

Características:
  1. Reintentos automáticos (3 intentos por defecto)
  2. Detección de datos en cero + reintentos en 15min
  3. Timeout en conexiones SQL
  4. Limpieza de memoria entre ejecuciones

Uso:
  python ejecutar_corte_orquestado.py --hora 13
"""

import argparse
import subprocess
import sys
import logging
import gc
import psutil
import time
from pathlib import Path
from datetime import datetime

SCRIPT_DIR = Path(__file__).parent.absolute()
GENERAR = SCRIPT_DIR / "generar_corte_ventas.py"
CAPTURAR = SCRIPT_DIR / "capturar_cortes.py"
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

# Importar módulo de alertas personales
try:
    from alertas_personales import AlertasPersonales
    ALERTAS_DISPONIBLES = True
except ImportError:
    ALERTAS_DISPONIBLES = False

# Configuración de reintentos
REINTENTOS_POR_DEFECTO = 3
ESPERA_ENTRE_REINTENTOS = 5  # segundos
ESPERA_DATOS_CERO = 900  # 15 minutos en segundos
TIMEOUT_GENERAR = 120  # 2 minutos


def setup_logger(hora: int):
    """Configura logging a archivo diario + consola."""
    fecha_str = datetime.now().strftime("%Y%m%d")
    log_file = LOG_DIR / f"cortes_{fecha_str}.log"

    logger = logging.getLogger("orquestado")
    logger.setLevel(logging.DEBUG)
    logger.handlers = []

    fh = logging.FileHandler(log_file, encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(
        "%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    ))

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter("%(message)s"))

    logger.addHandler(fh)
    logger.addHandler(ch)

    return logger


def limpiar_memoria():
    """Libera memoria no utilizada."""
    gc.collect()
    try:
        proceso = psutil.Process()
        mem_antes = proceso.memory_info().rss / 1024 / 1024  # MB
        mem_despues = proceso.memory_info().rss / 1024 / 1024
        return mem_antes, mem_despues
    except Exception:
        return None, None


def detectar_datos_cero(hora: int, logger) -> bool:
    """Detecta si el archivo generado tiene datos en cero."""
    desde_archivo = SCRIPT_DIR / f"corte_validacion_{hora}.flag"

    try:
        if desde_archivo.exists():
            with open(desde_archivo, 'r') as f:
                datos = f.read().strip()
                desde_archivo.unlink()

                if "SOLES_TOTAL:" in datos:
                    valor = float(datos.split(":")[-1])
                    if valor > 0:
                        logger.info(f"[OK] Datos válidos (S/ {valor:,.2f})")
                        return True
                    else:
                        logger.warning(f"[ALERTA] Datos en cero — reintentando en 15 min")
                        return False
    except Exception as e:
        logger.warning(f"[WARN] No se pudo validar datos: {e}")
        return True

    return True


def ejecutar_con_reintentos(cmd, logger, max_reintentos=3, timeout=None, es_generar=False):
    """Ejecuta comando con reintentos automáticos.

    Para datos en cero (es_generar=True):
      - Primer cero: Espera 15 min, reintenta una sola vez
      - Si sigue siendo cero: ABORT + alerta técnica
    """
    intentos_cero = 0
    max_intentos_cero = 1  # Solo 1 reintento por datos en cero

    for intento in range(1, max_reintentos + 1):
        logger.info(f"[Intento {intento}/{max_reintentos}] {cmd[1].split('/')[-1]}...")

        try:
            mem_antes, mem_despues = limpiar_memoria()
            if mem_antes:
                logger.debug(f"   Memoria: {mem_antes:.1f}MB → {mem_despues:.1f}MB")

            resultado = subprocess.run(
                cmd,
                cwd=SCRIPT_DIR,
                timeout=timeout,
                capture_output=False
            )

            if es_generar and resultado.returncode == 0:
                if not detectar_datos_cero(int(cmd[cmd.index("--hora") + 1]), logger):
                    # Datos en cero detectados
                    intentos_cero += 1

                    if intentos_cero > max_intentos_cero:
                        # Ya reintentamos una vez y SIGUE siendo cero
                        logger.error(f"\n[CRÍTICO] DATOS EN CERO PERSISTENTES")
                        logger.error(f"   Intento 1: Ceros")
                        logger.error(f"   Espera: 15 minutos")
                        logger.error(f"   Intento 2: SIGUE SIENDO CEROS")
                        logger.error(f"   → Aborting: Contactar a Data Engineer")
                        return False, -2  # Código especial: -2 = ceros persistentes

                    # Primer cero: esperar y reintentar
                    logger.warning(f"[ALERTA] Datos en cero. Esperando 15 minutos antes de reintentar...")
                    logger.warning(f"   Si sigue siendo cero, se abortará (requiere intervención manual)")
                    time.sleep(ESPERA_DATOS_CERO)
                    continue

            if resultado.returncode == 0:
                logger.info(f"[OK] Intento {intento} exitoso ✓")
                return True, resultado.returncode

            if intento < max_reintentos:
                logger.warning(f"[ERROR] Código {resultado.returncode}. Esperando {ESPERA_ENTRE_REINTENTOS}s...")
                time.sleep(ESPERA_ENTRE_REINTENTOS)

        except subprocess.TimeoutExpired:
            logger.error(f"[TIMEOUT] Tardó más de {timeout}s")
            if intento < max_reintentos:
                logger.warning(f"   Esperando {ESPERA_ENTRE_REINTENTOS}s antes de reintentar...")
                time.sleep(ESPERA_ENTRE_REINTENTOS)

        except Exception as e:
            logger.error(f"[EXCEPTION] {type(e).__name__}: {e}")
            if intento < max_reintentos:
                time.sleep(ESPERA_ENTRE_REINTENTOS)

    logger.error(f"[FAIL] Todos los {max_reintentos} intentos fallaron")
    return False, -1


def main():
    parser = argparse.ArgumentParser(
        description="Orquestador robusto de cortes de ventas"
    )
    parser.add_argument('--hora', type=int, required=True, help='Hora (8-18)')
    parser.add_argument('--max-reintentos', type=int, default=REINTENTOS_POR_DEFECTO,
                       help=f'Máximo de reintentos (default: {REINTENTOS_POR_DEFECTO})')
    args = parser.parse_args()

    hora = args.hora
    max_reintentos = args.max_reintentos
    logger = setup_logger(hora)

    if hora < 8 or hora > 18:
        logger.error(f"Hora fuera de rango: {hora}")
        sys.exit(1)

    logger.info(f"{'='*70}")
    logger.info(f"[ORQUESTADO] Corte {hora:02d}:00 — Reintentos: {max_reintentos}")
    logger.info(f"{'='*70}")

    # Inicializar alertas personales (solo si están disponibles)
    alertas = None
    if ALERTAS_DISPONIBLES:
        alertas = AlertasPersonales()
        if alertas.disponible:
            logger.info(f"[ALERTAS] Activas — Notificaciones a WhatsApp personal")
        else:
            logger.warning(f"[ALERTAS] No disponibles (config incompleta)")

    # [1] Generar Excel
    logger.info(f"\n[1/2] GENERANDO EXCEL...")
    cmd_generar = [
        sys.executable, str(GENERAR),
        "--hora", str(hora), "--solo-excel"
    ]

    exito_generar, cod_gen = ejecutar_con_reintentos(
        cmd_generar, logger, max_reintentos=max_reintentos,
        timeout=TIMEOUT_GENERAR, es_generar=True
    )

    if not exito_generar:
        if cod_gen == -2:
            # Datos en cero persistentes
            logger.error(f"\n[CRÍTICO] DATOS EN CERO PERSISTENTES")
            logger.error(f"   El Data Center posiblemente está con problema serio")
            logger.error(f"   O es horario sin ventas (madrugada/fin de día)")
            logger.error(f"   Requiere revisión manual del Data Engineer")

            # ENVIAR ALERTA A WHATSAPP PERSONAL
            if alertas and alertas.disponible:
                logger.info(f"\n[ALERTAS] Enviando notificación a WhatsApp personal...")
                fecha_str = datetime.now().strftime('%Y-%m-%d')
                alertas.enviar_alerta_datos_cero(hora=hora, fecha=fecha_str)

            sys.exit(2)  # Exit code diferente para alertas
        else:
            logger.error(f"\n[FATAL] No se pudo generar Excel")
            sys.exit(1)

    logger.info(f"[OK] Excel generado")

    # [2] Capturar y enviar
    logger.info(f"\n[2/2] CAPTURANDO Y ENVIANDO...")
    cmd_capturar = [
        sys.executable, str(CAPTURAR),
        "--hora", str(hora), "--destino", "canal"
    ]

    exito_capturar, cod_cap = ejecutar_con_reintentos(
        cmd_capturar, logger, max_reintentos=max_reintentos, timeout=120
    )

    if not exito_capturar:
        logger.error(f"\n[FATAL] No se pudo capturar/enviar")
        sys.exit(1)

    logger.info(f"[OK] Imágenes capturadas y enviadas")

    logger.info(f"\n{'='*70}")
    logger.info(f"[OK] CORTE {hora:02d}:00 COMPLETADO ✓")
    logger.info(f"{'='*70}\n")

    # OPCIONAL: Enviar alerta de éxito a WhatsApp personal
    # Descomenta la siguiente línea si deseas recibir notificaciones de éxito
    # if alertas and alertas.disponible:
    #     fecha_str = datetime.now().strftime('%Y-%m-%d')
    #     alertas.enviar_alerta_exito(hora=hora, fecha=fecha_str)


if __name__ == "__main__":
    main()
