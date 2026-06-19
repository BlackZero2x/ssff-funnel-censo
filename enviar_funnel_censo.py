"""
Envía el funnel censo SSFF por correo Gmail vía API.

Flujo:
  1. Verifica si los archivos Excel son recientes (<=60 min); si no, los regenera.
  2. Captura dos rangos de FUNNEL_CENSO como imágenes inline (PNG).
  3. Autentica con Gmail usando OAuth2 (token guardado en proyecto Movistar).
  4. Envía el correo con imágenes en el cuerpo y los Excel como adjuntos.

Ejecución automática: Programador de Tareas de Windows, todos los días a las 6am.
Ejecución manual:     python enviar_funnel_censo.py
"""
# ── Librería estándar ──────────────────────────────────────────────────────────
import base64
import glob
import logging
import os
import subprocess
import sys
import time
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from urllib.parse import quote

# ── Configuración ─────────────────────────────────────────────────────────────
DIR_SSFF     = r'C:\proyectos\SSFF'       # directorio del proyecto SSFF
DIR_MOVISTAR = r'C:\proyectos\AVANCE_MOVISTAR'  # directorio con credenciales Google

ARCHIVO_PRINCIPAL   = os.path.join(DIR_SSFF, 'funnel_censo_SSFF.xlsx')
PATRON_SUPERVISORES = os.path.join(DIR_SSFF, 'funnel_censo_*.xlsx')

# Rangos del Excel que se capturan como imágenes inline en el correo
RANGOS_CAPTURA = [
    ('FUNNEL_CENSO', 'B3:E16',  'captura_funnel.png'),
    ('FUNNEL_CENSO', 'I3:M25',  'captura_categorias.png'),
]

# ── Variables de entorno ───────────────────────────────────────────────────────
def _cargar_dotenv(env_path, override=False):
    """Carga un archivo .env en os.environ. Con override=True, sobreescribe valores existentes."""
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding='utf-8').splitlines():
        line = line.strip()
        if not line or line.startswith('#') or '=' not in line:
            continue
        key, _, value = line.partition('=')
        key = key.strip()
        value = value.strip()
        if override:
            os.environ[key] = value
        else:
            os.environ.setdefault(key, value)

# SSFF primero (correos, etc.); Movistar segundo para proxy/Google — no pisa SSFF
_cargar_dotenv(Path(DIR_SSFF) / '.env',     override=True)
_cargar_dotenv(Path(DIR_MOVISTAR) / '.env', override=False)

# Destinatarios leídos del .env (separados por coma)
PARA = [m.strip() for m in os.environ.get('EMAIL_PARA', '').split(',') if m.strip()]
CC   = [m.strip() for m in os.environ.get('EMAIL_CC',   '').split(',') if m.strip()]

def _configurar_proxy():
    """
    Ajusta las variables de entorno de proxy para que las APIs de Google
    no pasen por el proxy corporativo (que no soporta TLS-sobre-TLS).
    - googleapis.com, google.com y gstatic.com van en NO_PROXY.
    - HTTPS_PROXY se elimina siempre (causa error SSL con el proxy HTTP corporativo).
    """
    # Dominios de Google excluidos del proxy
    no_proxy = 'localhost,127.0.0.1,.googleapis.com,.google.com,.gstatic.com'
    os.environ['NO_PROXY'] = no_proxy
    os.environ['no_proxy'] = no_proxy

    proxy_host = os.environ.get('HTTP_PROXY', '')
    user       = os.environ.get('PROXY_USER', '')
    password   = os.environ.get('PROXY_PASS', '')

    if proxy_host and user and password:
        # Codificar usuario/contraseña para URLs (maneja caracteres especiales)
        encoded_user = quote(user, safe='')
        encoded_pass = quote(password, safe='')
        base = proxy_host.replace('http://', '').replace('https://', '')
        proxy_url = f'http://{encoded_user}:{encoded_pass}@{base}'
    elif proxy_host:
        proxy_url = proxy_host
    else:
        proxy_url = ''

    if proxy_url:
        os.environ['HTTP_PROXY'] = proxy_url
        os.environ['http_proxy'] = proxy_url
    else:
        for var in ('HTTP_PROXY', 'http_proxy'):
            os.environ.pop(var, None)

    # Eliminar HTTPS_PROXY: el proxy corporativo solo habla HTTP, no HTTPS
    for var in ('HTTPS_PROXY', 'https_proxy'):
        os.environ.pop(var, None)

_configurar_proxy()

# ── Logging ───────────────────────────────────────────────────────────────────
# Formato: "08:27:01  INFO  mensaje" — visible en consola y en el archivo .log del Programador
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s  %(levelname)s  %(message)s',
    datefmt='%H:%M:%S',
)

# ── Autenticación Google (OAuth2) ─────────────────────────────────────────────
# Las credenciales y el token se comparten con el proyecto Movistar
sys.path.insert(0, os.path.join(DIR_MOVISTAR, 'modules', 'shared'))

import requests as _requests
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Scopes idénticos al token existente en Movistar — no cambiar sin borrar token.json
SCOPES = [
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets',
]
CREDENTIALS_PATH = os.path.join(DIR_MOVISTAR, 'credentials.json')
TOKEN_PATH       = os.path.join(DIR_MOVISTAR, 'token.json')


def _session_sin_proxy():
    """Crea una sesión requests que ignora el proxy corporativo (necesario para googleapis)."""
    s = _requests.Session()
    s.proxies = {'http': None, 'https': None}
    return s


def _autenticar_gmail():
    """
    Devuelve un servicio Gmail autenticado via OAuth2.
    - Si el token existe y es válido, lo reutiliza.
    - Si expiró, lo refresca sin pasar por el proxy corporativo.
    - Si no existe, abre el navegador para autorizar (solo la primera vez).
    """
    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            # Refrescar sin proxy para evitar el error SSL con googleapis
            creds.refresh(Request(session=_session_sin_proxy()))
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, 'w') as f:
            f.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)


# ── Captura de rangos Excel ────────────────────────────────────────────────────
def _capturar_rango(app, wb, sheet, rango, cap_path):
    """
    Exporta un rango como PNG usando Chart export.
    Este método no requiere clipboard ni ventana en primer plano, por lo que
    funciona tanto en sesión interactiva como en el Programador de Tareas.
    """
    try:
        xl_rng    = sheet.range(rango)
        xl_rng.api.Copy()
        time.sleep(1)
        charts    = sheet.api.ChartObjects()
        chart_obj = charts.Add(0, 0, xl_rng.width, xl_rng.height)
        chart     = chart_obj.Chart
        chart.Paste()
        time.sleep(1)
        chart.Export(os.path.normpath(cap_path))
        chart_obj.Delete()
        time.sleep(0.5)
        if os.path.exists(cap_path):
            logging.info(f'   Captura OK: {os.path.basename(cap_path)}')
            return True
        logging.warning(f'   Chart export no generó archivo para {rango}')
    except Exception as e:
        logging.error(f'   Error al capturar {rango}: {e}')
    return False


def _esperar_archivo_libre(ruta, timeout=60):
    """
    Espera hasta que ningún proceso tenga el archivo abierto (máx timeout seg).
    Usa msvcrt.locking para detectar bloqueos exclusivos igual que Excel COM.
    """
    import msvcrt
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            fd = os.open(ruta, os.O_RDWR | os.O_BINARY)
            try:
                msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)
                msvcrt.locking(fd, msvcrt.LK_UNLCK, 1)
                return True
            except OSError:
                pass
            finally:
                os.close(fd)
        except OSError:
            pass
        time.sleep(3)
    return False


def _capturar_imagenes():
    """
    Abre el Excel con xlwings y captura los rangos definidos como PNG.
    Usa visible=False para funcionar en sesiones sin escritorio interactivo
    (Programador de Tareas de Windows). Solo usa Chart export, que no requiere
    clipboard ni ventana en primer plano.
    """
    import xlwings as xw

    temp_dir = os.path.join(DIR_SSFF, 'temp_capturas')
    os.makedirs(temp_dir, exist_ok=True)

    # Esperar a que generar_funnel_censo.py libere el archivo
    if not _esperar_archivo_libre(ARCHIVO_PRINCIPAL, timeout=30):
        logging.error('El archivo principal sigue bloqueado después de 30s — se omite captura.')
        return []

    capturas = []  # lista de (nombre_archivo, ruta_png)
    app = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        wb  = app.books.open(os.path.normpath(ARCHIVO_PRINCIPAL),
                             update_links=False, read_only=True)

        # Esperar a que Excel termine cálculos
        for _ in range(60):
            if app.api.CalculationState == 0:
                break
            time.sleep(1)

        for hoja, rango, nombre_png in RANGOS_CAPTURA:
            cap_path = os.path.join(temp_dir, nombre_png)
            if os.path.exists(cap_path):
                os.remove(cap_path)
            sheet = wb.sheets[hoja]
            sheet.activate()
            ok = _capturar_rango(app, wb, sheet, rango, cap_path)
            if ok:
                capturas.append((nombre_png, cap_path))
            else:
                logging.warning(f'   No se pudo capturar {hoja}!{rango}')

        wb.close()
    except Exception as e:
        logging.error(f'Error en captura de Excel: {e}')
    finally:
        if app:
            try:
                app.quit()
            except Exception:
                pass

    return capturas


# ── Envío ──────────────────────────────────────────────────────────────────────
def _adjuntar_archivo(msg, ruta):
    with open(ruta, 'rb') as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(ruta)}"')
    msg.attach(part)


def _enviar(gmail_service, asunto, fecha_ref, capturas, adjuntos):
    # Mensaje multipart/related para soportar imágenes inline
    msg_outer = MIMEMultipart('mixed')
    msg_outer['to']      = ', '.join(PARA)
    msg_outer['cc']      = ', '.join(CC)
    msg_outer['subject'] = asunto

    # Parte relacionada (texto HTML + imágenes inline)
    msg_related = MIMEMultipart('related')

    # Construir HTML con imágenes inline usando Content-ID
    imgs_html = ''
    for nombre_png, _ in capturas:
        cid = nombre_png.replace('.', '_')
        imgs_html += f'<br><img src="cid:{cid}" style="max-width:100%;"><br>\n'

    html = f"""<html><body>
<p>Buen día,</p>
<p>Adjunto el reporte de gestión diaria de clientes censo SSFF al <b>{fecha_ref}</b>.<br>
Se incluye el archivo consolidado y los libros individuales por supervisor.</p>
{imgs_html}
<p>Saludos.</p>
</body></html>"""

    msg_related.attach(MIMEText(html, 'html', 'utf-8'))

    for nombre_png, cap_path in capturas:
        if os.path.exists(cap_path):
            with open(cap_path, 'rb') as f:
                img_part = MIMEImage(f.read(), _subtype='png')
            cid = nombre_png.replace('.', '_')
            img_part.add_header('Content-ID', f'<{cid}>')
            img_part.add_header('Content-Disposition', 'inline', filename=nombre_png)
            msg_related.attach(img_part)

    msg_outer.attach(msg_related)

    # Adjuntar archivos Excel
    for ruta in adjuntos:
        if os.path.exists(ruta):
            _adjuntar_archivo(msg_outer, ruta)
            logging.info(f'   Adjunto: {os.path.basename(ruta)}')
        else:
            logging.warning(f'   Archivo no encontrado, omitido: {ruta}')

    raw = base64.urlsafe_b64encode(msg_outer.as_bytes()).decode()
    result = gmail_service.users().messages().send(userId='me', body={'raw': raw}).execute()
    logging.info(f'Correo enviado. ID: {result["id"]}')
    return True


def _archivos_son_recientes(max_minutos=60):
    """
    Devuelve True si el archivo principal existe y fue generado hace <= max_minutos.
    Evita regenerar si ya se ejecutó recientemente (p.ej. si se reintenta el envío).
    """
    if not os.path.exists(ARCHIVO_PRINCIPAL):
        return False
    edad_min = (time.time() - os.path.getmtime(ARCHIVO_PRINCIPAL)) / 60
    if edad_min <= max_minutos:
        logging.info(f'   Archivos recientes ({edad_min:.1f} min) — se omite regeneración.')
        return True
    logging.info(f'   Archivos desactualizados ({edad_min:.1f} min) — se regenerarán.')
    return False


def _fecha_referencia():
    """Devuelve la fecha de modificación del archivo principal como string DD-MM-YYYY."""
    if os.path.exists(ARCHIVO_PRINCIPAL):
        mtime = os.path.getmtime(ARCHIVO_PRINCIPAL)
        return datetime.fromtimestamp(mtime).strftime('%d-%m-%Y')
    return datetime.now().strftime('%d-%m-%Y')


def main():
    logging.info('=' * 60)
    logging.info('ENVIO FUNNEL CENSO SSFF')
    logging.info('=' * 60)

    # [1] Generar archivos solo si no hay versión reciente (<=60 min)
    logging.info('[1] Verificando antigüedad de archivos...')
    if _archivos_son_recientes(max_minutos=60):
        logging.info('[1] Usando archivos existentes.')
    else:
        logging.info('[1] Ejecutando generar_funnel_censo.py...')
        resultado = subprocess.run(
            [sys.executable, os.path.join(DIR_SSFF, 'generar_funnel_censo.py')],
            capture_output=True, text=True, encoding='utf-8', errors='replace'
        )
        if resultado.returncode != 0:
            logging.error('Error al generar el funnel:')
            logging.error(resultado.stderr)
            sys.exit(1)
        logging.info('[1] Archivos generados correctamente.')
        # Pequeña pausa para que openpyxl libere el handle antes de que xlwings intente abrir el archivo
        time.sleep(5)

    # [2] Capturar imágenes de rangos Excel
    logging.info('[2] Capturando imágenes de FUNNEL_CENSO...')
    capturas = _capturar_imagenes()
    logging.info(f'[2] {len(capturas)}/{len(RANGOS_CAPTURA)} capturas obtenidas.')

    # [3] Autenticar Gmail
    logging.info('[3] Autenticando Gmail...')
    gmail = _autenticar_gmail()

    # [4] Armar adjuntos Excel
    archivos_sup = sorted(
        f for f in glob.glob(PATRON_SUPERVISORES)
        if os.path.basename(f) != 'funnel_censo_SSFF.xlsx'
    )
    adjuntos = [ARCHIVO_PRINCIPAL] + archivos_sup
    logging.info(f'[4] Adjuntos a enviar: {len(adjuntos)} archivos')

    # [5] Enviar
    fecha_ref = _fecha_referencia()
    asunto    = f'Clientes Censo SSFF - Gestion Diaria ({fecha_ref})'
    logging.info(f'[5] Enviando correo: {asunto}')
    _enviar(gmail, asunto, fecha_ref, capturas, adjuntos)

    logging.info('Done.')


if __name__ == '__main__':
    main()
