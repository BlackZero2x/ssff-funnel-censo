#!/usr/bin/env python3
"""
alertas_personales.py

Módulo para enviar alertas a WhatsApp personal en casos especiales:
  • Corte con datos en cero (CRÍTICO)
  • Corte completado exitosamente (OK)

NO se ejecuta en cortes normales.

Uso interno:
    from alertas_personales import AlertasPersonales

    alertas = AlertasPersonales()
    alertas.enviar_alerta_datos_cero(hora=10, fecha="2026-06-22")
    alertas.enviar_alerta_exito(hora=10, fecha="2026-06-22")
"""

import json
import logging
from pathlib import Path
from datetime import datetime
from typing import Optional

logger = logging.getLogger(__name__)


class AlertasPersonales:
    """Gestor de alertas personales a WhatsApp."""

    def __init__(self, config_path: str = "config.json"):
        """Inicializa con configuración.

        Args:
            config_path: Ruta al config.json
        """
        self.config_path = Path(config_path)
        self.config = self._cargar_config()
        self.mi_numero = self.config.get("my_number")

        if not self.mi_numero:
            logger.warning("[ALERTAS] my_number no configurado en config.json")
            self.disponible = False
        else:
            logger.info(f"[ALERTAS] Inicializado — Destino personal: {self.mi_numero}")
            self.disponible = True

    def _cargar_config(self) -> dict:
        """Carga config.json."""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            logger.warning(f"[ALERTAS] config.json no encontrado en {self.config_path}")
            return {}
        except json.JSONDecodeError:
            logger.warning(f"[ALERTAS] config.json malformado")
            return {}

    def _enviar_whatsapp(self, mensaje: str) -> bool:
        """Envía mensaje a mi WhatsApp personal.

        Args:
            mensaje: Texto del mensaje

        Returns:
            True si se envió correctamente
        """
        if not self.disponible:
            logger.warning("[ALERTAS] No se puede enviar: config incompleta")
            return False

        try:
            from wa_sender_antibang import WABangSafeSender

            sender = WABangSafeSender(wa_server_url="http://localhost:8002")

            if not sender._puede_enviar():
                logger.warning("[ALERTAS] Cuenta WhatsApp bloqueada")
                return False

            # Enviar a número personal
            resultado = sender.send_to_contact(self.mi_numero, mensaje)

            if resultado:
                logger.info(f"[ALERTAS] ✓ Mensaje enviado a {self.mi_numero}")
                return True
            else:
                logger.error(f"[ALERTAS] ✗ Fallo al enviar a {self.mi_numero}")
                return False

        except Exception as e:
            logger.error(f"[ALERTAS] Excepción: {e}")
            return False

    def enviar_alerta_datos_cero(self, hora: int, fecha: str) -> bool:
        """Envía alerta CRÍTICA cuando detecta datos en cero.

        Args:
            hora: Hora del corte (8-18)
            fecha: Fecha (YYYY-MM-DD)

        Returns:
            True si se envió correctamente
        """
        hora_lbl = {8: '8AM', 9: '9AM', 10: '10AM', 11: '11AM', 12: '12PM',
                    13: '1PM', 14: '2PM', 15: '3PM', 16: '4PM', 17: '5PM', 18: '6PM'}.get(hora, f'{hora}H')

        mensaje = (
            f"🚨 *ALERTA CRÍTICA — DATOS EN CERO*\n\n"
            f"*Corte:* {hora_lbl}\n"
            f"*Fecha:* {fecha}\n\n"
            f"❌ SQL devolvió S/ 0.00 en todas las categorías.\n\n"
            f"*Posibles causas:*\n"
            f"• Data center en limpieza de pedidos\n"
            f"• Problema en la BD\n"
            f"• Desconexión de red\n\n"
            f"*Acción:*\n"
            f"Se reintentará en 15 minutos automáticamente.\n"
            f"Si persiste, contactar a Data Engineer."
        )

        logger.info(f"[ALERTAS] Enviando alerta datos cero ({hora_lbl})...")
        return self._enviar_whatsapp(mensaje)

    def enviar_alerta_exito(self, hora: int, fecha: str, total_soles: float = 0.0) -> bool:
        """Envía confirmación cuando corte se envía exitosamente.

        Solo se envía si lo deseas (no es automático).

        Args:
            hora: Hora del corte
            fecha: Fecha (YYYY-MM-DD)
            total_soles: Total de soles del corte (opcional)

        Returns:
            True si se envió correctamente
        """
        hora_lbl = {8: '8AM', 9: '9AM', 10: '10AM', 11: '11AM', 12: '12PM',
                    13: '1PM', 14: '2PM', 15: '3PM', 16: '4PM', 17: '5PM', 18: '6PM'}.get(hora, f'{hora}H')

        soles_str = f"S/ {total_soles:,.2f}" if total_soles > 0 else "✓"

        mensaje = (
            f"✅ *CORTE EXITOSO*\n\n"
            f"*Corte:* {hora_lbl}\n"
            f"*Fecha:* {fecha}\n"
            f"*Total:* {soles_str}\n\n"
            f"📊 Reporte generado y enviado a Canal SSFF.\n"
            f"🔔 Gerencia notificada."
        )

        logger.info(f"[ALERTAS] Enviando confirmación éxito ({hora_lbl})...")
        return self._enviar_whatsapp(mensaje)

    def enviar_alerta_error(self, hora: int, fecha: str, error: str) -> bool:
        """Envía alerta cuando hay error en generación.

        Args:
            hora: Hora del corte
            fecha: Fecha (YYYY-MM-DD)
            error: Descripción del error

        Returns:
            True si se envió correctamente
        """
        hora_lbl = {8: '8AM', 9: '9AM', 10: '10AM', 11: '11AM', 12: '12PM',
                    13: '1PM', 14: '2PM', 15: '3PM', 16: '4PM', 17: '5PM', 18: '6PM'}.get(hora, f'{hora}H')

        mensaje = (
            f"⚠️ *ERROR EN CORTE*\n\n"
            f"*Corte:* {hora_lbl}\n"
            f"*Fecha:* {fecha}\n\n"
            f"*Error:* {error}\n\n"
            f"📋 Revisar logs en:\n"
            f"`C:\\proyectos\\SSFF\\Reportes_ssff_wsp\\logs\\`\n\n"
            f"🔧 Contactar a TI si persiste."
        )

        logger.info(f"[ALERTAS] Enviando alerta error ({hora_lbl})...")
        return self._enviar_whatsapp(mensaje)

    def test_conexion(self) -> bool:
        """Envía mensaje de prueba para verificar conexión.

        Returns:
            True si se envió correctamente
        """
        mensaje = (
            f"🧪 *Mensaje de prueba — Alertas personales*\n\n"
            f"Timestamp: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n\n"
            f"✅ Si ves esto, las alertas están configuradas correctamente."
        )

        logger.info("[ALERTAS] Enviando mensaje de prueba...")
        return self._enviar_whatsapp(mensaje)
