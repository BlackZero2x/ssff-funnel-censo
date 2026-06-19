@echo off
REM ejecutar_cortes_diarios.bat
REM Ejecuta el corte de la hora actual
REM Se ejecuta desde Task Scheduler cada hora (8AM-6PM)

setlocal enabledelayedexpansion

set "SCRIPT_DIR=%~dp0"
set "GENERAR=%SCRIPT_DIR%generar_corte_ventas.py"
set "CAPTURAR=%SCRIPT_DIR%capturar_cortes.py"
set "LOG_DIR=%SCRIPT_DIR%logs"

REM Crear carpeta logs si no existe
if not exist "!LOG_DIR!" mkdir "!LOG_DIR!"

REM Archivo de log diario
for /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set "FECHA=%%c%%a%%b")
set "LOG_FILE=!LOG_DIR!\cortes_!FECHA!.log"

REM Obtener hora actual (sin PowerShell, solo cmd nativo)
for /f "tokens=1 delims=:" %%a in ('echo %time%') do (set "HORA_ACTUAL=%%a")

REM Convertir a formato 0-padded
if %HORA_ACTUAL% lss 10 (
    set "HORA_LBL=0%HORA_ACTUAL%:00"
) else (
    set "HORA_LBL=%HORA_ACTUAL%:00"
)

REM Validar que la hora esté en rango 8-18
if %HORA_ACTUAL% lss 8 (
    echo [%HORA_LBL%] Fuera de horario (antes de 8AM). Abortando.
    echo [%HORA_LBL%] Fuera de horario (antes de 8AM). Abortando. >> "!LOG_FILE!"
    exit /b 0
)

if %HORA_ACTUAL% gtr 18 (
    echo [%HORA_LBL%] Fuera de horario (despues de 6PM). Abortando.
    echo [%HORA_LBL%] Fuera de horario (despues de 6PM). Abortando. >> "!LOG_FILE!"
    exit /b 0
)

echo.
echo ============================================================
echo  CORTE SSFF — %HORA_LBL%
echo ============================================================
echo.

echo. >> "!LOG_FILE!"
echo ============================================================ >> "!LOG_FILE!"
echo  CORTE SSFF — %HORA_LBL% >> "!LOG_FILE!"
echo  Fecha: !FECHA! >> "!LOG_FILE!"
echo  Inicio: %date% %time% >> "!LOG_FILE!"
echo ============================================================ >> "!LOG_FILE!"

echo [%HORA_LBL%] Generando y capturando corte...
echo [%HORA_LBL%] Inicio: %date% %time% >> "!LOG_FILE!"

REM Generar Excel + capturar (ambos con uv run)
cd /d "!SCRIPT_DIR!"
uv run python "!GENERAR!" --hora %HORA_ACTUAL% --solo-excel >> "!LOG_FILE!" 2>&1
if !errorlevel! equ 0 (
    echo [%HORA_LBL%] Excel generado OK
    echo [%HORA_LBL%] Excel generado OK >> "!LOG_FILE!"
) else (
    echo [%HORA_LBL%] ERROR en generar_corte_ventas.py (code !errorlevel!)
    echo [%HORA_LBL%] ERROR en generar_corte_ventas.py (code !errorlevel!) >> "!LOG_FILE!"
)

uv run python "!CAPTURAR!" --hora %HORA_ACTUAL% --destino canal >> "!LOG_FILE!" 2>&1
if !errorlevel! equ 0 (
    echo [%HORA_LBL%] Corte enviado OK
    echo [%HORA_LBL%] Corte enviado OK >> "!LOG_FILE!"
) else (
    echo [%HORA_LBL%] ERROR en capturar_cortes.py (code !errorlevel!)
    echo [%HORA_LBL%] ERROR en capturar_cortes.py (code !errorlevel!) >> "!LOG_FILE!"
)
echo [%HORA_LBL%] Fin: %date% %time% >> "!LOG_FILE!"
echo.

echo ============================================================
echo  Completado — Proxima ejecucion en 1 hora
echo ============================================================
echo.

endlocal
