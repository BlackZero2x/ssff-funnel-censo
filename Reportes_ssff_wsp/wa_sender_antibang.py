"""
wa_sender_antibang.py

Wrapper para OPENWA que implementa comportamiento "antibang":
- Delays aleatorios (evita patrones detectables)
- Distribución inteligente de envíos (batching)
- Monitoreo de rate limits
- Rotación de métodos (texto + imagen)
- Pausa automática si detecta bloqueo

Uso:
    from wa_sender_antibang import WABangSafeSender
    sender = WABangSafeSender()
    sender.send_to_group(group_id, "Mensaje", image_path=None)
"""

import time
import random
import logging
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, List, Dict
import requests
import json

logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] [%(levelname)s] %(message)s'
)
logger = logging.getLogger(__name__)


class WABangSafeSender:
    """
    Envía mensajes a WhatsApp con comportamiento "antibang".
    Evita detección de automatización manteniendo patrones humanos.
    """

    def __init__(self, wa_server_url: str = "http://localhost:8002"):
        """
        Args:
            wa_server_url: URL del servidor wa_server.js
        """
        self.wa_server_url = wa_server_url
        self.last_send_time = 0
        self.bloqueo_hasta = None
        self.conteo_por_grupo = {}  # track mensajes por grupo
        self.conteo_global = {}  # track por hora

        # Configuración de delays
        self.delay_min = 3  # segundos mínimo
        self.delay_max = 8  # segundos máximo
        self.batch_size = 5  # mensajes por lote antes de pausa larga
        self.batch_pause = 15  # segundos entre lotes
        self.grupo_pause = 2  # minutos entre volver a grupo

    def _esperar_antibang(self, tipo: str = "normal") -> None:
        """
        Espera con patrón aleatorio (parece humano).

        Args:
            tipo: "normal" (3-8s), "batch" (15-20s), "grupo" (2-3 min)
        """
        if tipo == "normal":
            delay = random.uniform(self.delay_min, self.delay_max)
        elif tipo == "batch":
            delay = random.uniform(self.batch_pause, self.batch_pause + 5)
        elif tipo == "grupo":
            delay = random.uniform(60 * 2, 60 * 3)  # 2-3 min
        else:
            delay = random.uniform(1, 2)

        logger.debug(f"Esperando {delay:.2f}s ({tipo})")
        time.sleep(delay)

    def _detectar_bloqueo(self, response: Dict) -> bool:
        """
        Detecta si WhatsApp bloqueó la cuenta.

        Códigos detectables:
        - 429: Too Many Requests (rate limit)
        - "Blocked": Cuenta bloqueada
        - "restricted": Funcionalidad restringida
        """
        if isinstance(response, dict):
            status = response.get("status", "")
            error = str(response.get("error", "")).lower()

            if "429" in str(response) or status == 429:
                logger.warning("⚠️ Rate limit detectado (429) - pausando 1h")
                self.bloqueo_hasta = datetime.now() + timedelta(hours=1)
                return True

            if "blocked" in error or "restricted" in error:
                logger.warning("🔴 Cuenta bloqueada - pausando 24h")
                self.bloqueo_hasta = datetime.now() + timedelta(hours=24)
                return True

        return False

    def _puede_enviar(self) -> bool:
        """Verifica si puede enviar (no está en bloqueo)."""
        if self.bloqueo_hasta and datetime.now() < self.bloqueo_hasta:
            tiempo_resta = (self.bloqueo_hasta - datetime.now()).total_seconds()
            horas = int(tiempo_resta // 3600)
            mins = int((tiempo_resta % 3600) // 60)
            logger.error(f"❌ Bloqueado. Reintentar en {horas}h {mins}m")
            return False
        return True

    def _contar_envios_hora(self, grupo_id: str) -> int:
        """Cuenta envíos en la última hora (para respetar límites)."""
        ahora = datetime.now()
        hace_una_hora = ahora - timedelta(hours=1)

        # Limpiar registros antiguos
        if grupo_id not in self.conteo_global:
            self.conteo_global[grupo_id] = []

        self.conteo_global[grupo_id] = [
            t for t in self.conteo_global[grupo_id]
            if t > hace_una_hora
        ]

        return len(self.conteo_global[grupo_id])

    def send_to_group(
        self,
        grupo_id: str,
        mensaje: str,
        imagen_path: Optional[str] = None,
        force_send: bool = False
    ) -> bool:
        """
        Envía mensaje a grupo con comportamiento antibang.

        Args:
            grupo_id: ID del grupo/canal
            mensaje: Texto del mensaje
            imagen_path: Ruta a imagen (opcional)
            force_send: Ignorar bloqueos (no recomendado)

        Returns:
            True si se envió, False si falló o fue bloqueado
        """
        if not force_send and not self._puede_enviar():
            return False

        # Validar límite horario (max 60 msg/hora por grupo)
        conteo = self._contar_envios_hora(grupo_id)
        if conteo >= 60:
            logger.warning(f"⚠️ Límite horario alcanzado para {grupo_id}")
            return False

        # Esperar antes de enviar (patrón aleatorio)
        self._esperar_antibang("normal")

        try:
            # Preparar endpoint
            if imagen_path and Path(imagen_path).exists():
                endpoint = f"{self.wa_server_url}/send-image"
                files = {"image": open(imagen_path, "rb")}
                data = {"chat_id": grupo_id, "caption": mensaje}
                response = requests.post(endpoint, files=files, data=data, timeout=30)
            else:
                endpoint = f"{self.wa_server_url}/send-text"
                data = {"chat_id": grupo_id, "message": mensaje}
                response = requests.post(endpoint, json=data, timeout=30)

            result = response.json()

            # Detectar bloqueos
            if self._detectar_bloqueo(result):
                return False

            # Registrar envío
            if grupo_id not in self.conteo_global:
                self.conteo_global[grupo_id] = []
            self.conteo_global[grupo_id].append(datetime.now())

            logger.info(f"✅ Enviado a {grupo_id}")

            # Pausa por lote (cada 5 mensajes)
            if len(self.conteo_global.get(grupo_id, [])) % self.batch_size == 0:
                logger.info(f"🔄 Pausa de lote ({self.batch_pause}s)")
                self._esperar_antibang("batch")

            return True

        except Exception as e:
            logger.error(f"❌ Error enviando a {grupo_id}: {e}")
            return False

    def send_batch_to_groups(
        self,
        grupos: Dict[str, Dict],
        delay_entre_grupos: int = 120
    ) -> Dict[str, bool]:
        """
        Envía a múltiples grupos con pausa inteligente entre ellos.

        Args:
            grupos: {
                "grupo_id_1": {"mensaje": "...", "imagen": "..."},
                "grupo_id_2": {"mensaje": "..."},
                ...
            }
            delay_entre_grupos: segundos de pausa entre cambiar de grupo

        Returns:
            {"grupo_id": True/False, ...}
        """
        resultados = {}

        for i, (grupo_id, config) in enumerate(grupos.items()):
            if not self._puede_enviar():
                logger.warning(f"⚠️ Saltando {grupo_id} (cuenta bloqueada)")
                resultados[grupo_id] = False
                continue

            # Enviar
            success = self.send_to_group(
                grupo_id,
                config.get("mensaje", ""),
                config.get("imagen")
            )
            resultados[grupo_id] = success

            # Pausa entre grupos (parece cambio de usuario)
            if i < len(grupos) - 1:
                espera = random.uniform(delay_entre_grupos - 30, delay_entre_grupos + 30)
                logger.info(f"⏳ Pausa entre grupos: {espera:.0f}s")
                time.sleep(espera)

        return resultados

    def get_status(self) -> Dict:
        """Devuelve estado actual del sender."""
        bloqueado = False
        tiempo_resta = None

        if self.bloqueo_hasta and datetime.now() < self.bloqueo_hasta:
            bloqueado = True
            tiempo_resta = (self.bloqueo_hasta - datetime.now()).total_seconds()

        return {
            "bloqueado": bloqueado,
            "tiempo_resta_segundos": tiempo_resta,
            "conteo_global": self.conteo_global
        }


# Ejemplo de uso
if __name__ == "__main__":
    sender = WABangSafeSender()

    # Envío individual
    sender.send_to_group(
        "120363278118591818@g.us",
        "Hola desde antibang",
        imagen_path=None
    )

    # Envío a múltiples grupos
    grupos = {
        "120363278118591818@g.us": {"mensaje": "Mensaje 1"},
        "120363399204738966@g.us": {"mensaje": "Mensaje 2"},
    }
    resultados = sender.send_batch_to_groups(grupos)
    print("Resultados:", resultados)

    # Ver estado
    print("Estado:", sender.get_status())
