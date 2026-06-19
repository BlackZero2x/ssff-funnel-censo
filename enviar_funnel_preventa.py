"""
Genera y envía el FUNNEL_PREVENTA del día por correo y WhatsApp.

Uso: python enviar_funnel_preventa.py
     (sin argumentos — detecta el día automáticamente)

Destinatarios correo : jesus.asencios@auren.com.pe, mercedes.loaiza@auren.com.pe
Destinatario WA      : Jesús Ascencios (51944956042)
"""
import datetime
import logging
import os
import subprocess
import sys

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('C:/proyectos/SSFF/logs/enviar_funnel_preventa.log',
                            encoding='utf-8'),
    ]
)

# ── Parámetros ────────────────────────────────────────────────────────────────
BASE_DIR      = 'C:/proyectos/SSFF'
DIR_MOVISTAR  = 'C:/proyectos/AVANCE_MOVISTAR'

DIAS_SEMANA   = {1: 'Lunes', 2: 'Martes', 3: 'Miércoles',
                 4: 'Jueves', 5: 'Viernes', 6: 'Sábado', 7: 'Domingo'}

HOY           = datetime.date.today()
DIA_SEMANA    = HOY.weekday() + 1   # 1=Lunes … 7=Domingo
NOMBRE_DIA    = DIAS_SEMANA[DIA_SEMANA]
FECHA_STR     = HOY.strftime('%d_%m_%Y')
FECHA_LEGIBLE = HOY.strftime('%d/%m/%Y')

ARCHIVO_EXCEL = f'{BASE_DIR}/FUNNEL_PREVENTA_{NOMBRE_DIA}_{FECHA_STR}.xlsx'

DESTINATARIOS_CORREO = [
    'jesus.asencios@auren.com.pe',
    'mercedes.loaiza@auren.com.pe',
]

JESUS_WA_CONTACT = 'Jesús Ascencios'

ASUNTO = f'Funnel Preventa SSFF — {NOMBRE_DIA} {FECHA_LEGIBLE}'

CUERPO_CORREO = f"""Estimados,

Adjunto el Funnel de Preventa SSFF del {NOMBRE_DIA} {FECHA_LEGIBLE}.

El reporte se actualiza durante el día conforme los vendedores registran su gestión. Incluye estado de visita, monto de compra, motivos de no compra, historial de ventas y segmento por cliente, con una hoja por supervisor.

Saludos,
Sistema SSFF
"""


# ════════════════════════════════════════════════════════════════════════════════
# 1. GENERAR EL EXCEL
# ════════════════════════════════════════════════════════════════════════════════

def generar_excel():
    logging.info(f'[1] Generando FUNNEL_PREVENTA ({NOMBRE_DIA})...')
    result = subprocess.run(
        [sys.executable, f'{BASE_DIR}/generar_funnel_preventa.py'],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        logging.error(f'Error generando Excel:\n{result.stderr}')
        raise RuntimeError('Falló la generación del Excel')
    logging.info(result.stdout)
    if not os.path.exists(ARCHIVO_EXCEL):
        raise FileNotFoundError(f'No se encontró el archivo generado: {ARCHIVO_EXCEL}')
    logging.info(f'   Excel listo: {ARCHIVO_EXCEL}')

# ════════════════════════════════════════════════════════════════════════════════
# 2. ENVÍO POR CORREO
# ════════════════════════════════════════════════════════════════════════════════

def _autenticar_gmail():
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    SCOPES           = ['https://www.googleapis.com/auth/gmail.modify']
    TOKEN_PATH       = os.path.join(DIR_MOVISTAR, 'token.json')
    CREDENTIALS_PATH = os.path.join(DIR_MOVISTAR, 'credentials.json')

    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, 'w') as f:
            f.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)


def enviar_correo():
    import base64
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders

    logging.info('[2] Enviando correo...')
    gmail = _autenticar_gmail()

    msg = MIMEMultipart()
    msg['to']      = ', '.join(DESTINATARIOS_CORREO)
    msg['subject'] = ASUNTO
    msg.attach(MIMEText(CUERPO_CORREO, 'plain', 'utf-8'))

    with open(ARCHIVO_EXCEL, 'rb') as f:
        parte = MIMEBase('application', 'octet-stream')
        parte.set_payload(f.read())
        encoders.encode_base64(parte)
        parte.add_header('Content-Disposition',
                         f'attachment; filename="{os.path.basename(ARCHIVO_EXCEL)}"')
        msg.attach(parte)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    result = gmail.users().messages().send(userId='me', body={'raw': raw}).execute()
    logging.info(f'   Correo enviado. ID: {result["id"]}')

# ════════════════════════════════════════════════════════════════════════════════
# 3. ENVÍO POR WHATSAPP
# ════════════════════════════════════════════════════════════════════════════════

def enviar_whatsapp():
    """Envía el archivo Excel a Jesús vía servidor Node de WhatsApp (wa_server.js)."""
    import json
    import requests

    logging.info('[3] Enviando archivo por WhatsApp a Jesus (via wa_server.js)...')

    WA_URL = 'http://localhost:8002'
    JESUS_WA_ID = '51944956042@c.us'
    ruta_abs = os.path.abspath(ARCHIVO_EXCEL).replace('/', '\\')

    # Verificar que el servidor está listo
    try:
        health = requests.get(f'{WA_URL}/health', timeout=5).json()
        if health.get('status') != 'ready':
            logging.error('   Servidor WhatsApp no está listo (status != ready).')
            return
    except Exception as e:
        logging.error(f'   No se puede conectar al servidor WhatsApp: {e}')
        return

    # Enviar el archivo
    try:
        resp = requests.post(
            f'{WA_URL}/send-file',
            json={'to': JESUS_WA_ID, 'file_path': ruta_abs, 'caption': ''},
            timeout=60,
        )
        data = resp.json()
        if resp.status_code == 200 and data.get('success'):
            logging.info(f'   Archivo enviado por WhatsApp. ID: {data.get("messageId")}')
        else:
            logging.error(f'   Error del servidor WhatsApp: {data.get("error")}')
    except Exception as e:
        logging.error(f'   Error enviando WhatsApp: {e}')
        import traceback
        logging.error(traceback.format_exc())

# ════════════════════════════════════════════════════════════════════════════════
# 4. MAIN
# ════════════════════════════════════════════════════════════════════════════════

def main():
    os.makedirs(f'{BASE_DIR}/logs', exist_ok=True)
    logging.info('=' * 60)
    logging.info(f'ENVIAR FUNNEL PREVENTA — {NOMBRE_DIA} {FECHA_LEGIBLE}')
    logging.info('=' * 60)

    generar_excel()
    enviar_correo()
    enviar_whatsapp()

    logging.info('Proceso completado OK.')


if __name__ == '__main__':
    main()
