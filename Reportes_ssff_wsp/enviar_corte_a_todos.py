#!/usr/bin/env python3
"""
enviar_corte_a_todos.py

Distribución inteligente de cortes a múltiples destinos con antibang.
Reutiliza wa_sender_antibang.py para evitar baneo.

Envía a:
  1. Canal oficial SSFF (reporte general)
  2. Grupos de supervisores (por FFVV)
  3. Jefe de proyecto (resumen ejecutivo)
  4. Gerente comercial (alertas)

Uso:
    python enviar_corte_a_todos.py --archivo CORTE_VENTAS_*.xlsx --hora 10AM
    python enviar_corte_a_todos.py --archivo CORTE_VENTAS_*.xlsx --hora 10AM --test
"""

import argparse
import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

# Importar el sender antibang
from wa_sender_antibang import WABangSafeSender

# ════════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ════════════════════════════════════════════════════════════════════════════════

CONFIG_PATH = Path(__file__).parent / "config.json"
LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / f"distribucion_{datetime.now().strftime('%Y%m%d')}.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


# ════════════════════════════════════════════════════════════════════════════════
# FUNCIONES
# ════════════════════════════════════════════════════════════════════════════════

def cargar_config() -> Dict:
    """Carga config.json"""
    if not CONFIG_PATH.exists():
        logger.error(f"config.json no encontrado en {CONFIG_PATH}")
        return {}

    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        logger.error(f"Error parseando config.json: {e}")
        return {}


def obtener_grupo_canal(config: Dict) -> Optional[str]:
    """Obtiene ID del canal oficial SSFF"""
    return config.get("groups", {}).get("Canal_SSFF_2026_Gestion")


def obtener_grupos_supervisores(config: Dict) -> Dict[str, str]:
    """
    Obtiene grupos de supervisores por FFVV.

    Estructura esperada en config.json:
    {
      "groups_supervisores": {
        "F8": "120363...@g.us",
        "M0": "120363...@g.us",
        "K0": "120363...@g.us",
        ...
      }
    }
    """
    return config.get("groups_supervisores", {})


def obtener_contacto_jefe(config: Dict) -> Optional[str]:
    """Obtiene número del jefe de proyecto (1:1)"""
    jefe = config.get("mentions", {}).get("jefe_proyecto", {})
    return jefe.get("numero")  # 51XXXXXXXXX@c.us (chat individual)


def obtener_contacto_gerente(config: Dict) -> Optional[str]:
    """Obtiene número del gerente comercial (1:1)"""
    gerente = config.get("mentions", {}).get("gerente_comercial", {})
    return gerente.get("numero")  # 51XXXXXXXXX@c.us


def construir_mensaje_canal(archivo: str, hora: str) -> str:
    """Mensaje para el canal oficial (todas las vistas)"""
    nombre_archivo = Path(archivo).name
    return f"""📊 CORTE DE VENTAS — {hora}

Reporte generado: {datetime.now().strftime('%H:%M')}
Archivo: {nombre_archivo}

3 vistas:
  • GENERAL: Métricas consolidadas
  • ZONAL: Desglose por zona/origen
  • SUPERVISOR: Detalle por cada supervisor

vs D-7 y D-14 (mismo día de semana)
"""


def construir_mensaje_supervisor(ffvv: str, hora: str) -> str:
    """Mensaje para supervisores de cada FFVV"""
    return f"""📈 CORTE {ffvv} — {hora}

Tu reporte está listo en el canal.
Revisa tu zona en la vista SUPERVISOR.

Comparativas:
  • vs Semana pasada (D-7)
  • vs Hace 2 semanas (D-14)
"""


def construir_mensaje_jefe(archivo: str, hora: str) -> str:
    """Mensaje para jefe de proyecto (resumen ejecutivo)"""
    return f"""📋 CORTE EJECUTIVO — {hora}

Todos los reportes generados.
Archivo adjunto con 3 vistas (GENERAL, ZONAL, SUPERVISOR).

Estado: ✅ OK
Timestamp: {datetime.now().strftime('%H:%M:%S')}
"""


def construir_mensaje_gerente(archivo: str, hora: str) -> str:
    """Mensaje para gerente (alertas de excepción)"""
    # Aquí podrías agregar lógica de alertas
    # por ahora es informativo
    return f"""🔔 ALERTA DE GERENCIA — {hora}

Nuevo corte disponible.
Archivo: {Path(archivo).name}

Revisar alertas en el canal si las hay.
"""


def enviar_distribucion(
    archivo: str,
    hora: str,
    config: Dict,
    test: bool = False
) -> Dict[str, bool]:
    """
    Envía el corte a todos los destinos con estrategia antibang.

    Orden:
    1. Canal oficial (todos lo ven)
    2. Grupos de supervisores (por FFVV)
    3. Jefe de proyecto (1:1)
    4. Gerente comercial (1:1)

    Args:
        archivo: Ruta al Excel generado
        hora: Etiqueta de corte (ej: "10AM")
        config: Diccionario con IDs
        test: Si True, solo simula (no envía)

    Returns:
        {destino: True/False, ...}
    """
    sender = WABangSafeSender(wa_server_url="http://localhost:8002")
    resultados = {}

    if not Path(archivo).exists():
        logger.error(f"❌ Archivo no encontrado: {archivo}")
        return {"archivo": False}

    logger.info("=" * 70)
    logger.info(f"🚀 Iniciando distribución — {hora}")
    logger.info("=" * 70)

    # ──────────────────────────────────────────────────────────────────────────────
    # 1. CANAL OFICIAL (todos)
    # ──────────────────────────────────────────────────────────────────────────────

    canal_id = obtener_grupo_canal(config)
    if canal_id:
        logger.info(f"\n[1] Enviando a Canal SSFF...")
        mensaje = construir_mensaje_canal(archivo, hora)

        if test:
            logger.info(f"   [TEST] Mensaje:\n{mensaje}")
            resultados["canal_oficial"] = True
        else:
            if sender.send_to_group(canal_id, mensaje, imagen_path=archivo):
                logger.info(f"   ✅ Enviado al canal oficial")
                resultados["canal_oficial"] = True
            else:
                logger.error(f"   ❌ Error enviando al canal")
                resultados["canal_oficial"] = False
    else:
        logger.warning("   ⚠️  ID del canal no configurado")
        resultados["canal_oficial"] = None

    # ──────────────────────────────────────────────────────────────────────────────
    # 2. GRUPOS DE SUPERVISORES (por FFVV)
    # ──────────────────────────────────────────────────────────────────────────────

    grupos_sup = obtener_grupos_supervisores(config)
    if grupos_sup:
        logger.info(f"\n[2] Enviando a {len(grupos_sup)} grupos de supervisores...")

        for ffvv, grupo_id in grupos_sup.items():
            mensaje = construir_mensaje_supervisor(ffvv, hora)

            if test:
                logger.info(f"   [TEST] {ffvv}: {mensaje[:50]}...")
                resultados[f"grupo_sup_{ffvv}"] = True
            else:
                if sender.send_to_group(grupo_id, mensaje, imagen_path=None):
                    logger.info(f"   ✅ {ffvv}")
                    resultados[f"grupo_sup_{ffvv}"] = True
                else:
                    logger.error(f"   ❌ {ffvv}")
                    resultados[f"grupo_sup_{ffvv}"] = False
    else:
        logger.warning("   ⚠️  Ningún grupo de supervisores configurado")

    # ──────────────────────────────────────────────────────────────────────────────
    # 3. JEFE DE PROYECTO (1:1)
    # ──────────────────────────────────────────────────────────────────────────────

    jefe_id = obtener_contacto_jefe(config)
    if jefe_id:
        logger.info(f"\n[3] Enviando a Jefe de Proyecto...")
        mensaje = construir_mensaje_jefe(archivo, hora)

        if test:
            logger.info(f"   [TEST] Mensaje:\n{mensaje}")
            resultados["jefe_proyecto"] = True
        else:
            if sender.send_to_group(jefe_id, mensaje, imagen_path=None):
                logger.info(f"   ✅ Enviado a jefe")
                resultados["jefe_proyecto"] = True
            else:
                logger.error(f"   ❌ Error enviando a jefe")
                resultados["jefe_proyecto"] = False
    else:
        logger.warning("   ⚠️  Contacto de jefe no configurado")
        resultados["jefe_proyecto"] = None

    # ──────────────────────────────────────────────────────────────────────────────
    # 4. GERENTE COMERCIAL (1:1)
    # ──────────────────────────────────────────────────────────────────────────────

    gerente_id = obtener_contacto_gerente(config)
    if gerente_id:
        logger.info(f"\n[4] Enviando a Gerente Comercial...")
        mensaje = construir_mensaje_gerente(archivo, hora)

        if test:
            logger.info(f"   [TEST] Mensaje:\n{mensaje}")
            resultados["gerente_comercial"] = True
        else:
            if sender.send_to_group(gerente_id, mensaje, imagen_path=None):
                logger.info(f"   ✅ Enviado a gerente")
                resultados["gerente_comercial"] = True
            else:
                logger.error(f"   ❌ Error enviando a gerente")
                resultados["gerente_comercial"] = False
    else:
        logger.warning("   ⚠️  Contacto de gerente no configurado")
        resultados["gerente_comercial"] = None

    # ──────────────────────────────────────────────────────────────────────────────
    # RESUMEN
    # ──────────────────────────────────────────────────────────────────────────────

    logger.info("\n" + "=" * 70)
    exitos = sum(1 for v in resultados.values() if v is True)
    fallos = sum(1 for v in resultados.values() if v is False)
    logger.info(f"📊 RESUMEN: {exitos} ✅ — {fallos} ❌")

    for destino, exito in resultados.items():
        estado = "✅" if exito is True else "❌" if exito is False else "⚠️"
        logger.info(f"   {estado} {destino}")

    logger.info("=" * 70)

    return resultados


# ════════════════════════════════════════════════════════════════════════════════
# CLI
# ════════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Distribuye cortes a múltiples destinos con antibang",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  python enviar_corte_a_todos.py --archivo CORTE_VENTAS_*.xlsx --hora 10AM
  python enviar_corte_a_todos.py --archivo CORTE_VENTAS_*.xlsx --hora 10AM --test
  python enviar_corte_a_todos.py --archivo CORTE_VENTAS_*.xlsx --hora 10AM --config config_custom.json
        """
    )

    parser.add_argument("--archivo", type=str, required=True, help="Ruta al Excel generado")
    parser.add_argument("--hora", type=str, default="Corte", help="Etiqueta de corte (ej: 10AM)")
    parser.add_argument("--config", type=str, default=str(CONFIG_PATH), help="Path a config.json")
    parser.add_argument("--test", action="store_true", help="Simular sin enviar realmente")

    args = parser.parse_args()

    # Cargar config
    config = cargar_config()
    if not config:
        logger.error("No se pudo cargar la configuración")
        exit(1)

    # Enviar
    resultados = enviar_distribucion(args.archivo, args.hora, config, test=args.test)

    # Exit code
    fallos = sum(1 for v in resultados.values() if v is False)
    exit(1 if fallos > 0 else 0)


if __name__ == "__main__":
    main()
