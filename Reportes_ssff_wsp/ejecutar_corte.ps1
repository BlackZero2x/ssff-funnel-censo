# ejecutar_corte.ps1
# Script PowerShell que ejecuta generar + capturar para una hora específica
# Uso: .\ejecutar_corte.ps1 -Hora 13

param(
    [Parameter(Mandatory=$true)]
    [int]$Hora
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Generar = "$ScriptDir\generar_corte_ventas.py"
$Capturar = "$ScriptDir\capturar_cortes.py"
$LogDir = "$ScriptDir\logs"

# Crear carpeta logs
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

# Validar rango
if ($Hora -lt 8 -or $Hora -gt 18) {
    Write-Host "[ERROR] Hora fuera de rango. Debe estar entre 8 y 18."
    exit 1
}

# Archivo log diario
$Fecha = Get-Date -Format "yyyyMMdd"
$LogFile = "$LogDir\cortes_$Fecha.log"

# Etiqueta de hora
$HoraLbl = @{
    8="08:00"; 9="09:00"; 10="10:00"; 11="11:00"; 12="12:00"
    13="13:00"; 14="14:00"; 15="15:00"; 16="16:00"; 17="17:00"; 18="18:00"
}[$Hora]

Write-Host ""
Write-Host "============================================================"
Write-Host "  CORTE SSFF — $HoraLbl"
Write-Host "============================================================"
Write-Host ""

Add-Content -Path $LogFile -Value ""
Add-Content -Path $LogFile -Value "============================================================"
Add-Content -Path $LogFile -Value "  CORTE SSFF — $HoraLbl"
Add-Content -Path $LogFile -Value "  Fecha: $Fecha"
Add-Content -Path $LogFile -Value "  Inicio: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss.ff')"
Add-Content -Path $LogFile -Value "============================================================"

Write-Host "[$HoraLbl] Generando y capturando corte..."
Add-Content -Path $LogFile -Value "[$HoraLbl] Inicio: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss.ff')"

# Generar Excel
Set-Location $ScriptDir
$output = & uv run python $Generar --hora $Hora --solo-excel 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "[$HoraLbl] Excel generado OK"
    Add-Content -Path $LogFile -Value "[$HoraLbl] Excel generado OK"
} else {
    Write-Host "[$HoraLbl] ERROR en generar_corte_ventas.py (code $LASTEXITCODE)"
    Add-Content -Path $LogFile -Value "[$HoraLbl] ERROR en generar_corte_ventas.py (code $LASTEXITCODE)"
    Add-Content -Path $LogFile -Value "$output"
}

# Capturar y enviar
$output = & uv run python $Capturar --hora $Hora --destino canal 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "[$HoraLbl] Corte enviado OK"
    Add-Content -Path $LogFile -Value "[$HoraLbl] Corte enviado OK"
} else {
    Write-Host "[$HoraLbl] ERROR en capturar_cortes.py (code $LASTEXITCODE)"
    Add-Content -Path $LogFile -Value "[$HoraLbl] ERROR en capturar_cortes.py (code $LASTEXITCODE)"
    Add-Content -Path $LogFile -Value "$output"
}

Add-Content -Path $LogFile -Value "[$HoraLbl] Fin: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss.ff')"
Write-Host ""
Write-Host "============================================================"
Write-Host "  Completado"
Write-Host "============================================================"
Write-Host ""
