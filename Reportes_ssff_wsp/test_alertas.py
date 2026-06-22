#!/usr/bin/env python3
"""
test_alertas.py

Script para testear el módulo de alertas personales.

Uso:
    python test_alertas.py                    # Enviar mensaje de prueba
    python test_alertas.py --tipo datos_cero # Simular alerta datos cero
    python test_alertas.py --tipo exito      # Simular alerta éxito
"""

import argparse
from alertas_personales import AlertasPersonales
from datetime import datetime


def main():
    parser = argparse.ArgumentParser(description='Test alertas personales')
    parser.add_argument('--tipo', default='test',
                       choices=['test', 'datos_cero', 'exito', 'error'],
                       help='Tipo de alerta a probar')
    args = parser.parse_args()

    print(f"\n{'='*70}")
    print(f"  TEST ALERTAS PERSONALES")
    print(f"{'='*70}\n")

    alertas = AlertasPersonales()

    if not alertas.disponible:
        print("❌ Alertas no disponibles")
        print("   Verificar config.json: my_number debe estar configurado")
        return

    print(f"✓ Alertas disponibles")
    print(f"  Destino: {alertas.mi_numero}\n")

    if args.tipo == 'test':
        print("[1/1] Enviando mensaje de prueba...")
        resultado = alertas.test_conexion()

    elif args.tipo == 'datos_cero':
        print("[1/1] Simulando alerta datos cero...")
        resultado = alertas.enviar_alerta_datos_cero(
            hora=10,
            fecha=datetime.now().strftime('%Y-%m-%d')
        )

    elif args.tipo == 'exito':
        print("[1/1] Simulando alerta éxito...")
        resultado = alertas.enviar_alerta_exito(
            hora=10,
            fecha=datetime.now().strftime('%Y-%m-%d'),
            total_soles=85500.50
        )

    elif args.tipo == 'error':
        print("[1/1] Simulando alerta error...")
        resultado = alertas.enviar_alerta_error(
            hora=10,
            fecha=datetime.now().strftime('%Y-%m-%d'),
            error="Timeout en SQL (120s)"
        )

    print(f"\n{'='*70}")
    if resultado:
        print(f"✅ Mensaje enviado correctamente")
    else:
        print(f"❌ Fallo al enviar mensaje")
    print(f"{'='*70}\n")


if __name__ == "__main__":
    main()
