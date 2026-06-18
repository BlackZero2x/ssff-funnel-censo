#!/usr/bin/env python3
"""
test_antibang.py

Prueba rápida de la integración antibang.
Verifica:
  1. Servidor OPENWA activo
  2. Configuración cargada
  3. Envío individual (test)
  4. Estado de la cuenta
"""

import sys
import json
from pathlib import Path
from datetime import datetime

# Importar antibang
try:
    from wa_sender_antibang import WABangSafeSender
except ImportError:
    print("❌ ERROR: wa_sender_antibang.py no encontrado")
    sys.exit(1)

# ════════════════════════════════════════════════════════════════════════════════

def test_servidor():
    """Verifica que el servidor wa_server.js está activo"""
    print("\n[1/4] Verificando servidor OPENWA...")

    sender = WABangSafeSender(wa_server_url="http://localhost:8002")

    try:
        import requests
        response = requests.get("http://localhost:8002/health", timeout=5)
        if response.status_code == 200:
            print("   ✅ Servidor activo")
            return True
        else:
            print(f"   ❌ Servidor respondió con status {response.status_code}")
            return False
    except requests.exceptions.ConnectionError:
        print("   ❌ No se pudo conectar a http://localhost:8002")
        print("      ¿Está corriendo? → cd Reportes_ssff_wsp && node wa_server.js")
        return False
    except Exception as e:
        print(f"   ❌ Error: {e}")
        return False


def test_config():
    """Verifica que config.json está bien configurado"""
    print("\n[2/4] Verificando configuración...")

    config_path = Path("config.json")
    if not config_path.exists():
        print("   ❌ config.json no encontrado")
        print("      Copia config.example.json → config.json")
        return False, None

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)

        # Validaciones
        checks = {
            "my_number": config.get("my_number"),
            "Canal_SSFF_2026_Gestion": config.get("groups", {}).get("Canal_SSFF_2026_Gestion"),
            "jefe_proyecto": config.get("mentions", {}).get("jefe_proyecto", {}).get("numero"),
            "gerente_comercial": config.get("mentions", {}).get("gerente_comercial", {}).get("numero"),
        }

        print(f"   Configuración encontrada:")
        for campo, valor in checks.items():
            estado = "✅" if valor else "⚠️"
            print(f"      {estado} {campo}: {valor if valor else '(vacío)'}")

        # Requerimiento mínimo: al menos el canal oficial
        if not checks["Canal_SSFF_2026_Gestion"]:
            print("\n   ⚠️  ADVERTENCIA: Canal_SSFF_2026_Gestion no está configurado")
            print("      El envío automático no funcionará sin este grupo")

        return True, config

    except json.JSONDecodeError as e:
        print(f"   ❌ Error parseando config.json: {e}")
        return False, None
    except Exception as e:
        print(f"   ❌ Error: {e}")
        return False, None


def test_envio(config):
    """Prueba un envío simulado"""
    print("\n[3/4] Prueba de envío (TEST, no realmente)...")

    grupo_id = config.get("groups", {}).get("Canal_SSFF_2026_Gestion")
    if not grupo_id:
        print("   ⏭️  Saltando (grupo no configurado)")
        return None

    sender = WABangSafeSender(wa_server_url="http://localhost:8002")

    # Mensaje de prueba
    mensaje = f"""🧪 PRUEBA DE ANTIBANG

Timestamp: {datetime.now().strftime('%H:%M:%S')}
Estado: ✅ Sistema listo

Este mensaje se envió con:
  • Delays aleatorios (3-8s)
  • Batching inteligente
  • Rate limit detection
  • Auto-pause en bloqueo

Sin riesgo de baneo. 🛡️
"""

    print(f"   Enviando a: {grupo_id}")
    print(f"   Mensaje: {mensaje[:50]}...")

    try:
        exito = sender.send_to_group(grupo_id, mensaje, imagen_path=None)

        if exito:
            print("   ✅ Enviado exitosamente")
            return True
        else:
            print("   ❌ Error en envío")
            return False

    except Exception as e:
        print(f"   ❌ Excepción: {e}")
        return False


def test_estado(config):
    """Muestra el estado actual del sistema"""
    print("\n[4/4] Estado actual...")

    sender = WABangSafeSender(wa_server_url="http://localhost:8002")
    estado = sender.get_status()

    bloqueado = estado.get("bloqueado", False)
    if bloqueado:
        secs = int(estado.get("tiempo_resta_segundos", 0))
        horas = secs // 3600
        mins = (secs % 3600) // 60
        print(f"   🔴 CUENTA BLOQUEADA — Reintentar en {horas}h {mins}m")
    else:
        print("   ✅ Cuenta operativa")

    conteo = estado.get("conteo_global", {})
    if conteo:
        print(f"\n   Envíos en última hora:")
        for grupo, eventos in conteo.items():
            print(f"      • {grupo[:20]}...: {len(eventos)} mensajes")
    else:
        print("   (Sin envíos previos en esta sesión)")


# ════════════════════════════════════════════════════════════════════════════════

def main():
    print("=" * 70)
    print("🧪 TEST ANTIBANG SSFF")
    print("=" * 70)

    # Test 1: Servidor
    servidor_ok = test_servidor()
    if not servidor_ok:
        print("\n❌ No puedes continuar sin servidor")
        print("   Inicia: cd Reportes_ssff_wsp && node wa_server.js")
        sys.exit(1)

    # Test 2: Config
    config_ok, config = test_config()
    if not config_ok:
        print("\n❌ Configuración inválida")
        sys.exit(1)

    # Test 3: Envío
    envio_ok = test_envio(config)

    # Test 4: Estado
    test_estado(config)

    # Resumen
    print("\n" + "=" * 70)
    if envio_ok:
        print("✅ TODAS LAS PRUEBAS PASARON")
        print("\nAhora puedes ejecutar:")
        print("  python generar_corte_ventas.py --hora 10")
        print("  python enviar_corte_a_todos.py --archivo CORTE_*.xlsx --hora 10AM")
    elif envio_ok is None:
        print("⚠️  PRUEBAS PARCIALES (grupo no configurado)")
        print("\nPróximos pasos:")
        print("  1. Llenar config.json con IDs reales")
        print("  2. Volver a ejecutar test_antibang.py")
    else:
        print("❌ FALLÓ PRUEBA DE ENVÍO")
        print("   Verifica config.json y que el servidor está activo")

    print("=" * 70)


if __name__ == "__main__":
    main()
