# SSFF Census Funnel — Daily Tracking Report

Automated system that generates and delivers a daily Excel report tracking sales coverage for San Fernando (SSFF) census clients post-restructuring (08/03/2026).

> **Versión en español:** [README.es.md](README.es.md)

---

## What it does

1. **Pulls live data** from SQL Server (sales history, client master, visit reasons) and MySQL (field visit surveys).
2. **Generates an Excel workbook** (`funnel_censo_SSFF.xlsx`) with 5 sheets:
   - `FUNNEL_CENSO` — coverage funnel, category sales table, and supervisor/seller summaries.
   - `DETALLE_CLIENTE` — full census client list with coverage indicators per month.
   - `ULTIMA_COMPRA_CENSO` — last purchase detail for census clients, with visit and TomaPedido data.
   - `ULTIMA_COMPRA_NO_CENSO` — same for non-census clients with monthly sales history.
   - `HISTORICO` — sales history by supervisor and seller (COB + SOLES + VAR%).
3. **Generates one workbook per supervisor** (`funnel_censo_<supervisor>.xlsx`) with the same sheets filtered.
4. **Sends the report by email** (Gmail API, OAuth2) with two Excel range captures as inline images and all workbooks as attachments.

Runs automatically every day at 6 AM via Windows Task Scheduler.

---

## Project structure

```
SSFF/
├── generar_funnel_censo.py   # Main script — data extraction + Excel generation
├── enviar_funnel_censo.py    # Email sender — capture + Gmail send
├── requirements.txt          # Python dependencies
├── .gitignore                # Git exclusions (data files, credentials, logs)
├── .env                      # Environment variables — NOT versioned
└── README.md / README.es.md  # Documentation
```

---

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure `.env`

Create a `.env` file in the project root (never commit it):

```env
# SQL Server
SQL_SERVER=YOUR_SERVER\INSTANCE
SQL_DATABASE=eAuren
SQL_USER=your_user
SQL_PASSWORD=your_password

# Email recipients
EMAIL_PARA=recipient1@company.com,recipient2@company.com
EMAIL_CC=manager@company.com
```

### 3. Google credentials

Place `credentials.json` and `token.json` in the Movistar project directory (`C:\proyectos\AVANCE_MOVISTAR\`). The first run will open a browser to authorize Gmail access.

### 4. MySQL DSN

Configure a system DSN named `linux_auditorias_recargas` pointing to the MySQL server with the surveys database.

---

## Usage

```bash
# Generate Excel files only
python generar_funnel_censo.py

# Generate (if needed) + capture + send email
python enviar_funnel_censo.py
```

`enviar_funnel_censo.py` checks if the Excel files are less than 60 minutes old. If they are recent, it skips regeneration and sends them directly.

---

## Data sources

| Source | Connection | Data |
|--------|-----------|------|
| SQL Server `eAuren` | pyodbc direct | Sales history, client master SP, visit reasons view |
| MySQL `auditorias_recargas` | pyodbc DSN | Field visit surveys (encuesta 166) |
| CSV fallback | Local file | Used when SQL/MySQL is unavailable |

---

## HOY cutoff logic

When running before 10 AM, the script uses **yesterday** as the reference date to avoid a partial-day projection that would understate results. After 10 AM it uses today.

---

## Scheduled task (Windows)

The task is registered in Windows Task Scheduler to run daily at 6 AM:

```
Program: C:\...\python.exe
Arguments: C:\proyectos\SSFF\enviar_funnel_censo.py
Log: C:\proyectos\SSFF\enviar_funnel_censo.log
```

---

## Tech stack

- **Python 3.10+**
- **pandas / numpy** — data processing
- **openpyxl** — Excel generation with full formatting
- **pyodbc** — SQL Server and MySQL connections
- **google-auth / google-api-python-client** — Gmail OAuth2
- **xlwings / Pillow / pywin32** — Excel range capture as PNG (Windows only)
