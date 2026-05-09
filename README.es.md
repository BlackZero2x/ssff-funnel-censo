# SSFF Funnel Censo — Reporte Diario de Seguimiento

Sistema automatizado que genera y envía un reporte Excel diario con el seguimiento de cobertura de ventas para clientes censo de San Fernando (SSFF) post-reestructuración (08/03/2026).

> **English version:** [README.md](README.md)

---

## Qué hace

1. **Extrae datos en vivo** desde SQL Server (histórico de ventas, maestro de clientes, motivos de no-preventa) y MySQL (encuestas de visita de campo).
2. **Genera un libro Excel** (`funnel_censo_SSFF.xlsx`) con 5 hojas:
   - `FUNNEL_CENSO` — funnel de cobertura, tabla de soles por categoría y resúmenes por supervisor/vendedor.
   - `DETALLE_CLIENTE` — listado completo de clientes censo con indicadores de cobertura por mes.
   - `ULTIMA_COMPRA_CENSO` — detalle de última compra de clientes censo, con datos de visita y TomaPedido.
   - `ULTIMA_COMPRA_NO_CENSO` — ídem para clientes no-censo con histórico mensual de soles.
   - `HISTORICO` — histórico de ventas por supervisor y vendedor (COB + SOLES + VAR%).
3. **Genera un libro por supervisor** (`funnel_censo_<supervisor>.xlsx`) con las mismas hojas filtradas.
4. **Envía el reporte por correo** (Gmail API, OAuth2) con dos capturas de rangos Excel como imágenes inline y todos los libros como adjuntos.

Se ejecuta automáticamente cada día a las 6 AM mediante el Programador de Tareas de Windows.

---

## Estructura del proyecto

```
SSFF/
├── generar_funnel_censo.py   # Script principal — extracción de datos + generación Excel
├── enviar_funnel_censo.py    # Script de envío — captura de imágenes + envío Gmail
├── requirements.txt          # Dependencias Python
├── .gitignore                # Exclusiones de Git (datos, credenciales, logs)
├── .env                      # Variables de entorno — NO se versiona
└── README.md / README.es.md  # Documentación
```

---

## Configuración

### 1. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 2. Configurar el archivo `.env`

Crear un archivo `.env` en la raíz del proyecto (nunca subirlo a Git):

```env
# SQL Server
SQL_SERVER=TU_SERVIDOR\INSTANCIA
SQL_DATABASE=eAuren
SQL_USER=tu_usuario
SQL_PASSWORD=tu_contraseña

# Destinatarios del correo
EMAIL_PARA=destinatario1@empresa.com,destinatario2@empresa.com
EMAIL_CC=gerente@empresa.com
```

### 3. Credenciales de Google

Colocar `credentials.json` y `token.json` en el directorio del proyecto Movistar (`C:\proyectos\AVANCE_MOVISTAR\`). La primera ejecución abrirá el navegador para autorizar el acceso a Gmail.

### 4. DSN de MySQL

Configurar un DSN de sistema llamado `linux_auditorias_recargas` apuntando al servidor MySQL con la base de datos de encuestas.

---

## Uso

```bash
# Solo generar los archivos Excel
python generar_funnel_censo.py

# Generar (si es necesario) + capturar + enviar correo
python enviar_funnel_censo.py
```

`enviar_funnel_censo.py` verifica si los archivos Excel tienen menos de 60 minutos de antigüedad. Si son recientes, omite la regeneración y los envía directamente.

---

## Fuentes de datos

| Fuente | Conexión | Datos |
|--------|----------|-------|
| SQL Server `eAuren` | pyodbc directo | Histórico de ventas, SP maestro de clientes, vista de motivos |
| MySQL `auditorias_recargas` | pyodbc DSN | Encuestas de visita de campo (encuesta 166) |
| CSV / Excel fallback | Archivo local | Se usa cuando SQL/MySQL no está disponible |

---

## Lógica de fecha HOY

Cuando el script se ejecuta antes de las 10 AM, usa **ayer** como fecha de referencia para evitar que el proyectado del día parcial subestime los resultados. A partir de las 10 AM usa la fecha de hoy.

---

## Tarea programada (Windows)

La tarea está registrada en el Programador de Tareas de Windows para ejecutarse diariamente a las 6 AM:

```
Programa: C:\...\python.exe
Argumentos: C:\proyectos\SSFF\enviar_funnel_censo.py
Log: C:\proyectos\SSFF\enviar_funnel_censo.log
```

---

## Stack tecnológico

- **Python 3.10+**
- **pandas / numpy** — procesamiento de datos
- **openpyxl** — generación de Excel con formato completo
- **pyodbc** — conexión a SQL Server y MySQL
- **google-auth / google-api-python-client** — OAuth2 para Gmail
- **xlwings / Pillow / pywin32** — captura de rangos Excel como PNG (solo Windows)
