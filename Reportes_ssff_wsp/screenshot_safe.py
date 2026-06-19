#!/usr/bin/env python3
"""
screenshot_safe.py

Sistema de captura de imágenes Excel con mutex para evitar conflictos
cuando múltiples proyectos capturan al mismo tiempo.

Usa un archivo de bloqueo compartido en C:\proyectos\locks\ para sincronizar
acceso a recursos de pantalla/COM entre SSFF, MOVISTAR y otros proyectos.

Uso:
    mgr = ScreenshotManager("SSFF_Corte_8AM")
    if mgr.adquirir_lock(timeout=30):
        try:
            capturar_excel(...)
        finally:
            mgr.liberar_lock()
"""

import os
import time
import logging
from pathlib import Path
from datetime import datetime

# ════════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ════════════════════════════════════════════════════════════════════════════════

LOCKS_DIR = Path("C:/proyectos/locks")
LOCKS_DIR.mkdir(parents=True, exist_ok=True)

# Logger
_logger = logging.getLogger("screenshot_safe")


# ════════════════════════════════════════════════════════════════════════════════
# CLASE: ScreenshotManager
# ════════════════════════════════════════════════════════════════════════════════

class ScreenshotManager:
    """
    Gestiona acceso exclusivo a captura de pantalla Excel.

    Usa un archivo lock para sincronizar múltiples proyectos.
    Si otro proyecto está capturando, espera hasta timeout.

    Ejemplo:
        mgr = ScreenshotManager("SSFF_Corte_8AM")
        if mgr.adquirir_lock(timeout=30):
            try:
                # ... capturar imagen ...
            finally:
                mgr.liberar_lock()
        else:
            print("Timeout esperando lock")
    """

    def __init__(self, identificador: str):
        """
        Args:
            identificador: Nombre único del proyecto/tarea (ej: "SSFF_Corte_8AM", "MOVISTAR_Avance")
        """
        self.id = identificador
        self.lock_file = LOCKS_DIR / f"{identificador}.lock"
        self.adquirido = False

    def adquirir_lock(self, timeout: int = 30, intervalo: float = 0.5) -> bool:
        """
        Intenta adquirir el lock.

        Si otro proceso lo tiene, espera hasta timeout.

        Args:
            timeout: Segundos máximos a esperar
            intervalo: Segundos entre intentos de polling

        Returns:
            True si se adquirió, False si timeout
        """
        inicio = time.time()

        while time.time() - inicio < timeout:
            try:
                # Crear archivo de lock en modo exclusivo (falla si existe)
                fd = os.open(str(self.lock_file), os.O_CREAT | os.O_EXCL | os.O_WRONLY)

                # Escribir PID y timestamp
                with os.fdopen(fd, 'w') as f:
                    f.write(f"PID={os.getpid()}\n")
                    f.write(f"START={datetime.now().isoformat()}\n")
                    f.write(f"ID={self.id}\n")

                self.adquirido = True
                _logger.info(f"✅ Lock adquirido: {self.id}")
                return True

            except FileExistsError:
                # Otro proceso tiene el lock
                edad = self._edad_lock()
                if edad > 120:  # Si lock tiene >2 min, probablemente es stale
                    _logger.warning(f"⚠️  Lock stale detectado ({edad}s): {self.lock_file}")
                    try:
                        self.lock_file.unlink()
                        continue
                    except:
                        pass

                # Esperar y reintentar
                _logger.debug(f"⏳ Esperando lock: {self.id} (edad={edad}s)")
                time.sleep(intervalo)

            except Exception as e:
                _logger.error(f"❌ Error adquiriendo lock: {e}")
                return False

        _logger.error(f"❌ Timeout adquiriendo lock: {self.id}")
        return False

    def liberar_lock(self) -> bool:
        """
        Libera el lock.

        Returns:
            True si se liberó, False si ya estaba libre o error
        """
        if not self.adquirido:
            return False

        try:
            self.lock_file.unlink()
            self.adquirido = False
            _logger.info(f"🔓 Lock liberado: {self.id}")
            return True
        except Exception as e:
            _logger.error(f"❌ Error liberando lock: {e}")
            return False

    def _edad_lock(self) -> int:
        """Retorna edad del lock en segundos (0 si no existe)"""
        try:
            return int(time.time() - self.lock_file.stat().st_mtime)
        except:
            return 0

    def __enter__(self):
        """Context manager: adquirir lock"""
        if not self.adquirir_lock():
            raise RuntimeError(f"No se pudo adquirir lock: {self.id}")
        return self

    def __exit__(self, *args):
        """Context manager: liberar lock"""
        self.liberar_lock()


# ════════════════════════════════════════════════════════════════════════════════
# FUNCIÓN: Capturar tabla Excel como imagen PNG
# ════════════════════════════════════════════════════════════════════════════════

def capturar_tabla_excel(worksheet, rango: str, ruta_png: str,
                         escala: float = 2.5, tiempo_espera: float = 1.0) -> bool:
    """
    Captura una tabla Excel como imagen PNG.

    Requiere:
    - Excel abierto (COM)
    - Rango visible en worksheet
    - PIL instalado

    Args:
        worksheet: Objeto xlwings.Sheet
        rango: Rango Excel (ej: "A1:H8")
        ruta_png: Ruta de salida PNG
        escala: Factor de escala (2.5 = 250% resolución)
        tiempo_espera: Segundos a esperar después de copiar

    Returns:
        True si se capturó exitosamente
    """
    try:
        from PIL import ImageGrab, Image
        import win32gui
        import ctypes
    except ImportError:
        _logger.error("❌ Requeridas: PIL, win32gui (pip install pillow pywin32)")
        return False

    try:
        # Seleccionar rango
        xl_range = worksheet.range(rango)

        # Traer Excel al frente
        try:
            app = worksheet.book.app
            app.api.Visible = True
            hwnd = [0]

            def _cb(h, _):
                if (win32gui.IsWindowVisible(h) and
                    "Microsoft Excel" in win32gui.GetWindowText(h)):
                    hwnd[0] = h
                    return False
                return True

            win32gui.EnumWindows(_cb, None)
            if hwnd[0]:
                ctypes.windll.user32.AllowSetForegroundWindow(
                    ctypes.windll.kernel32.GetCurrentProcessId()
                )
                win32gui.ShowWindow(hwnd[0], 9)  # SW_RESTORE
                win32gui.SetForegroundWindow(hwnd[0])

            time.sleep(tiempo_espera)
        except Exception as e:
            _logger.warning(f"⚠️  No se pudo traer Excel al frente: {e}")

        # CopyPicture (con reintentos)
        for intento in range(3):
            try:
                xl_range.api.CopyPicture(Appearance=1, Format=2)
                time.sleep(tiempo_espera)
                break
            except Exception as e:
                if intento < 2:
                    _logger.warning(f"⚠️  CopyPicture intento {intento+1} fallido: {e}")
                    time.sleep(1 + intento * 0.5)
                else:
                    raise

        # Pegar desde clipboard
        img = ImageGrab.grabclipboard()
        if not img:
            _logger.error(f"❌ Clipboard vacío (no hay imagen)")
            return False

        # Escalar para mejor resolución
        nuevo_ancho = int(img.width * escala)
        nuevo_alto = int(img.height * escala)
        img = img.resize((nuevo_ancho, nuevo_alto), resample=Image.LANCZOS)

        # Guardar PNG
        Path(ruta_png).parent.mkdir(parents=True, exist_ok=True)
        img.save(ruta_png, "PNG")

        _logger.info(f"✅ Imagen capturada: {ruta_png} ({nuevo_ancho}x{nuevo_alto})")
        return True

    except Exception as e:
        _logger.error(f"❌ Error capturando tabla: {e}")
        return False


if __name__ == "__main__":
    # Test de mutex
    logging.basicConfig(level=logging.INFO)

    print("🔒 Test ScreenshotManager")
    mgr = ScreenshotManager("TEST_SSFF_8AM")

    if mgr.adquirir_lock(timeout=5):
        print("✅ Lock adquirido")
        time.sleep(2)
        mgr.liberar_lock()
        print("✅ Lock liberado")
    else:
        print("❌ Timeout esperando lock")
