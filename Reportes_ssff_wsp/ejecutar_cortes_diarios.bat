@echo off
REM ejecutar_cortes_diarios.bat
REM Ejecuta cortes horarios de 8AM a 6PM en loop diario
REM Una sola tarea en Task Scheduler a las 8:00 AM

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

REM Horarios a ejecutar (8 a 18, en formato 24h)
set "HORARIOS=8 9 10 11 12 13 14 15 16 17 18"

echo. >> "!LOG_FILE!"
echo ============================================================ >> "!LOG_FILE!"
echo  CORTES SSFF — Ejecucion en Loop Diario (8AM-6PM) >> "!LOG_FILE!"
echo  Fecha: !FECHA! >> "!LOG_FILE!"
echo  Inicio: %date% %time% >> "!LOG_FILE!"
echo ============================================================ >> "!LOG_FILE!"
echo.
echo ============================================================
echo  CORTES SSFF — Ejecucion en Loop Diario (8AM-6PM)
echo ============================================================
echo.

for %%H in (%HORARIOS%) do (
    set "HORA=%%H"

    if %%H lss 10 (
        set "HORA_LBL=0%%H:00"
    ) else (
        set "HORA_LBL=%%H:00"
    )

    echo [!HORA_LBL!] Generando y capturando corte...
    echo [!HORA_LBL!] Inicio: %date% %time% >> "!LOG_FILE!"

    REM Generar Excel + capturar (ambos con uv run)
    cd /d "!SCRIPT_DIR!"
    uv run python "!GENERAR!" --hora %%H --solo-excel >> "!LOG_FILE!" 2>&1
    if !errorlevel! equ 0 (
        echo [!HORA_LBL!] Excel generado OK >> "!LOG_FILE!"
    ) else (
        echo [!HORA_LBL!] ERROR en generar_corte_ventas.py (code !errorlevel!) >> "!LOG_FILE!"
    )

    uv run python "!CAPTURAR!" --hora %%H --destino canal >> "!LOG_FILE!" 2>&1
    if !errorlevel! equ 0 (
        echo [!HORA_LBL!] Corte enviado OK >> "!LOG_FILE!"
    ) else (
        echo [!HORA_LBL!] ERROR en capturar_cortes.py (code !errorlevel!) >> "!LOG_FILE!"
    )
    echo [!HORA_LBL!] Fin: %date% %time% >> "!LOG_FILE!"
    echo. >> "!LOG_FILE!"

    REM Pausa entre cortes (evita saturar recursos)
    if %%H lss 18 (
        echo [!HORA_LBL!] Esperando 55 minutos hasta proximo corte...
        echo [!HORA_LBL!] Proximo corte a las !HORA_LBL! >> "!LOG_FILE!"
        timeout /t 3300 /nobreak
    )
)

echo.
echo ============================================================
echo  Ciclo diario completado. Proxima ejecucion manana a las 8AM
echo ============================================================
echo.

echo. >> "!LOG_FILE!"
echo ============================================================ >> "!LOG_FILE!"
echo  Ciclo completado — %date% %time% >> "!LOG_FILE!"
echo ============================================================ >> "!LOG_FILE!"
echo. >> "!LOG_FILE!"

goto :end

:end
endlocal
