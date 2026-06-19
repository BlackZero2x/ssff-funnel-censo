@echo off
REM ejecutar_corte_horario.bat
REM Ejecuta el corte de una hora específica (recibida como parámetro)
REM Uso: ejecutar_corte_horario.bat 13

setlocal enabledelayedexpansion

set "SCRIPT_DIR=%~dp0"
set "GENERAR=%SCRIPT_DIR%generar_corte_ventas.py"
set "CAPTURAR=%SCRIPT_DIR%capturar_cortes.py"
set "LOG_DIR=%SCRIPT_DIR%logs"

REM Crear carpeta logs si no existe
if not exist "!LOG_DIR!" mkdir "!LOG_DIR!"

REM Obtener hora del parámetro (8-18)
set "HORA=%1"
if "%HORA%"=="" (
    echo [ERROR] Uso: ejecutar_corte_horario.bat ^<hora^>
    echo Ejemplo: ejecutar_corte_horario.bat 13
    exit /b 1
)

REM Validar rango
if %HORA% lss 8 (
    echo [ERROR] Hora fuera de rango. Debe estar entre 8 y 18.
    exit /b 1
)
if %HORA% gtr 18 (
    echo [ERROR] Hora fuera de rango. Debe estar entre 8 y 18.
    exit /b 1
)

REM Archivo de log diario
for /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set "FECHA=%%c%%a%%b")
set "LOG_FILE=!LOG_DIR!\cortes_!FECHA!.log"

REM Convertir a formato 0-padded
if %HORA% lss 10 (
    set "HORA_LBL=0%HORA%:00"
) else (
    set "HORA_LBL=%HORA%:00"
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
uv run python "!GENERAR!" --hora %HORA% --solo-excel >> "!LOG_FILE!" 2>&1
if !errorlevel! equ 0 (
    echo [%HORA_LBL%] Excel generado OK
    echo [%HORA_LBL%] Excel generado OK >> "!LOG_FILE!"
) else (
    echo [%HORA_LBL%] ERROR en generar_corte_ventas.py (code !errorlevel!)
    echo [%HORA_LBL%] ERROR en generar_corte_ventas.py (code !errorlevel!) >> "!LOG_FILE!"
)

uv run python "!CAPTURAR!" --hora %HORA% --destino canal >> "!LOG_FILE!" 2>&1
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
echo  Completado
echo ============================================================
echo.

endlocal
