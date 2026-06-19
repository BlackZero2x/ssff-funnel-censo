#!/usr/bin/env python3
"""
enviar_reporte_ssff.py

Envía notificaciones de cuotas SSFF al grupo WhatsApp
Reutiliza el servidor wa_server.js de MOVISTAR (puerto 8002)

Uso:
    python enviar_reporte_ssff.py --archivo cuotas_ssff_mayo2026.xlsx --mes "Mayo 2026"
    python enviar_reporte_ssff.py --test
"""

import argparse
import json
import requests
import logging
from datetime import datetime
from pathlib import Path
import random
import time

# ══════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════

CONFIG_PATH = Path(__file__).parent / "config.json"
LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / f"envios_{datetime.now().strftime('%Y%m%d')}.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════════
# FUNCIONES CORE
# ══════════════════════════════════════════════════════════════════

def cargar_config():
    """Carga config.json"""
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        logger.error(f"config.json no encontrado en {CONFIG_PATH}")
        exit(1)
    except json.JSONDecodeError as e:
        logger.error(f"Error al parsear config.json: {e}")
        exit(1)

def verificar_horario_seguro(config):
    """Verifica si es horario permitido para envío"""
    horarios = config["ssff_config"]["horarios_permitidos"]
    ahora = datetime.now()

    # ¿Es día laboral?
    if ahora.weekday() not in horarios["dias"]:
        return False, "No es día laboral (fuera de lunes-viernes)"

    # ¿Está en horario?
    hora_inicio = int(horarios["inicio"].split(":")[0])
    hora_fin = int(horarios["fin"].split(":")[0])

    if not (hora_inicio <= ahora.hour < hora_fin):
        return False, f"Fuera de horario permitido ({hora_inicio}h-{hora_fin}h)"

    return True, "OK"

def conectar_servidor(config, timeout=5):
    """Verifica que el servidor está activo"""
    host = config["server_info"]["host"]
    port = config["server_info"]["port"]
    url = f"http://{host}:{port}/health"

    try:
        response = requests.get(url, timeout=timeout)
        if response.status_code == 200:
            return True, "Servidor activo"
        else:
            return False, f"Servidor respondió con status {response.status_code}"
    except requests.exceptions.ConnectionError:
        return False, f"No se pudo conectar a {url}"
    except requests.exceptions.Timeout:
        return False, f"Timeout conectando a {url}"
    except Exception as e:
        return False, f"Error: {str(e)}"

def enviar_mensaje_whatsapp(grupo_id, mensaje, config):
    """Envía un mensaje al grupo mediante wa_client.py"""
    try:
        # Llamar a wa_client.py con subprocess
        import subprocess

        cmd = [
            "python",
            str(Path(__file__).parent / "wa_client.py"),
            "--send-to", grupo_id,
            "--message", mensaje
        ]

        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)

        if result.returncode == 0:
            return True, "Mensaje enviado"
        else:
            error_msg = result.stderr or result.stdout
            return False, f"Error en wa_client: {error_msg}"

    except subprocess.TimeoutExpired:
        return False, "Timeout enviando mensaje"
    except Exception as e:
        return False, f"Excepción: {str(e)}"

def registrar_envio(grupo_id, mensaje, estado, detalles=""):
    """Registra el envío en JSON para auditoría"""
    registro = {
        "timestamp": datetime.now().isoformat(),
        "grupo_id": grupo_id,
        "mensaje": mensaje[:80] + "..." if len(mensaje) > 80 else mensaje,
        "estado": estado,
        "detalles": detalles
    }

    log_file = LOG_DIR / "envios_registro.jsonl"
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(json.dumps(registro, ensure_ascii=False) + "\n")

def leer_archivo(path):
    """Lee el contenido de un archivo (para validar existencia)"""
    try:
        with open(path, "rb") as f:
            return True, f"Archivo {Path(path).name} ({Path(path).stat().st_size / 1024:.1f} KB)"
    except FileNotFoundError:
        return False, f"Archivo no encontrado: {path}"
    except Exception as e:
        return False, f"Error leyendo archivo: {str(e)}"

# ══════════════════════════════════════════════════════════════════
# FLUJOS PRINCIPALES
# ══════════════════════════════════════════════════════════════════

def enviar_reporte(archivo, mes, menciones=None):
    """
    Flujo principal: Genera mensaje y envía al grupo SSFF

    Args:
        archivo: Ruta al archivo Excel de cuotas
        mes: Nombre del mes (ej: "Mayo 2026")
        menciones: Lista de roles a mencionar ["supervisores", "jefe_proyecto", ...]
    """
    config = cargar_config()
    menciones = menciones or []

    logger.info("=" * 70)
    logger.info(f"🚀 Iniciando envío de reporte SSFF — {mes}")
    logger.info("=" * 70)

    # 1. Verificar horario
    es_seguro, razon = verificar_horario_seguro(config)
    if not es_seguro:
        logger.warning(f"⏰ {razon} — Envío posrgado")
        return False
    logger.info("✅ Horario seguro para envío")

    # 2. Verificar servidor
    servidor_ok, msg = conectar_servidor(config)
    if not servidor_ok:
        logger.error(f"❌ Servidor no disponible: {msg}")
        logger.error("   Asegúrate que 'node wa_server.js' está corriendo en puerto 8002")
        return False
    logger.info(f"✅ Servidor activo: {msg}")

    # 3. Validar archivo
    archivo_ok, msg = leer_archivo(archivo)
    if not archivo_ok:
        logger.error(f"❌ {msg}")
        return False
    logger.info(f"✅ {msg}")

    # 4. Construir mensaje
    grupo_id = config["groups"]["Canal_SSFF_2026_Gestion"]
    variantes = config["message_variants"]["reporte_generado"]
    mensaje_base = random.choice(variantes).format(mes=mes)

    # Agregar menciones si corresponde
    if menciones:
        menciones_str = []
        for rol in menciones:
            if rol == "supervisores" and config["mentions"]["supervisores"]:
                menciones_str.append(f"@{', @'.join(config['mentions']['supervisores'])}")
            elif rol == "jefe_proyecto" and config["mentions"]["jefe_proyecto"]["nombre"]:
                menciones_str.append(f"@{config['mentions']['jefe_proyecto']['nombre']}")
            elif rol == "gerente_comercial" and config["mentions"]["gerente_comercial"]["nombre"]:
                menciones_str.append(f"@{config['mentions']['gerente_comercial']['nombre']}")

        if menciones_str:
            mensaje_base = f"{', '.join(menciones_str)}\n\n{mensaje_base}"

    logger.info(f"📝 Mensaje a enviar:\n   {mensaje_base}")

    # 5. Enviar
    logger.info(f"📤 Enviando a grupo: {grupo_id}...")
    time.sleep(config["ssff_config"]["intervalo_minimo_segundos"])

    exito, detalles = enviar_mensaje_whatsapp(grupo_id, mensaje_base, config)

    if exito:
        logger.info(f"✅ {detalles}")
        registrar_envio(grupo_id, mensaje_base, "OK", f"Archivo: {Path(archivo).name}")
        logger.info("=" * 70)
        return True
    else:
        logger.error(f"❌ {detalles}")
        registrar_envio(grupo_id, mensaje_base, "ERROR", detalles)
        logger.error("=" * 70)
        return False

def prueba_conexion():
    """Verifica que todo esté listo para envíos"""
    config = cargar_config()

    logger.info("🔍 Verificando configuración...")
    logger.info(f"   Número WhatsApp: {config['my_number']}")
    logger.info(f"   Servidor: {config['server_info']['host']}:{config['server_info']['port']}")
    logger.info(f"   Grupo SSFF: {config['groups']['Canal_SSFF_2026_Gestion']}")

    servidor_ok, msg = conectar_servidor(config)
    if servidor_ok:
        logger.info(f"✅ {msg}")
        return True
    else:
        logger.error(f"❌ {msg}")
        logger.error("\n   Solución: Inicia el servidor en la carpeta MOVISTAR")
        logger.error("   $ cd C:\\proyectos\\AVANCE_MOVISTAR\\whatsapp_server")
        logger.error("   $ node wa_server.js")
        return False

# ══════════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Envía notificaciones de cuotas SSFF a WhatsApp",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  python enviar_reporte_ssff.py --test
  python enviar_reporte_ssff.py --archivo cuotas_ssff_mayo2026.xlsx --mes "Mayo 2026"
  python enviar_reporte_ssff.py --archivo cuotas.xlsx --mes "Mayo 2026" --menciones supervisores,jefe_proyecto
        """
    )

    parser.add_argument("--test", action="store_true", help="Verifica conexión al servidor")
    parser.add_argument("--archivo", type=str, help="Ruta al archivo Excel de cuotas")
    parser.add_argument("--mes", type=str, default="Cuotas mensuales", help="Nombre del mes a incluir en mensaje")
    parser.add_argument("--menciones", type=str, help="Roles a mencionar: supervisores,jefe_proyecto,gerente_comercial")

    args = parser.parse_args()

    if args.test:
        exito = prueba_conexion()
        exit(0 if exito else 1)

    if args.archivo:
        menciones = [m.strip() for m in args.menciones.split(",")] if args.menciones else []
        exito = enviar_reporte(args.archivo, args.mes, menciones)
        exit(0 if exito else 1)

    parser.print_help()

if __name__ == "__main__":
    main()
