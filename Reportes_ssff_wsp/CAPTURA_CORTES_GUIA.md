# 📸 Guía de Captura de Cortes SSFF

Sistema de captura de tablas Excel como imágenes PNG y envío automático a WhatsApp.

---

## 🎯 Qué hace

1. **Genera Excel** con 3 tablas: GENERAL, ZONAL, SUPERVISOR
2. **Captura tablas como PNG** (resolución 2.5x, legible en móvil)
3. **Envía a WhatsApp** con etiqueta "CORTE 8AM", "CORTE 9AM", etc.
4. **Sincroniza con MOVISTAR** usando mutex (evita conflictos de screenshot)

---

## 🔧 Arquitectura

### Flujo por hora

```
[8 AM]
  ├─ generar_corte_ventas.py --hora 8
  │   └─ Genera CORTE_VENTAS_20260619_080000.xlsx
  ├─ capturar_cortes.py --hora 8 --destino test
  │   ├─ screenshot_safe.py (acquire lock)
  │   ├─ Captura 3 tablas como PNG
  │   ├─ wa_sender_antibang.py (envío seguro)
  │   └─ screenshot_safe.py (release lock)
  └─ Imágenes enviadas a WhatsApp

[9 AM - 6 PM]
  ├─ capturar_cortes.py --hora 9 --destino test
  │   └─ (Captura del Excel más reciente)
  └─ Imágenes enviadas a WhatsApp
```

### Módulos involucrados

| Módulo | Función |
|--------|---------|
| `generar_corte_ventas.py` | Genera Excel con 3 tablas |
| `capturar_cortes.py` | Captura tablas, envía WhatsApp |
| `screenshot_safe.py` | Mutex para evitar conflictos |
| `wa_sender_antibang.py` | Envío seguro (delays + batching) |
| `programar_cortes.py` | Crear tareas automáticas en Task Scheduler |
| `config.json` | Configuración (destinos, horarios) |

---

## 📋 Instalación

### Requisitos

```bash
pip install openpyxl pandas numpy xlwings pillow pywin32 requests
```

**Nota:** `pywin32` requiere setup post-instalación:

```bash
python -m pip install --upgrade pywin32
python Scripts/pywin32_postinstall.py -install
```

O en PowerShell como Admin:

```powershell
python -m pip install --upgrade pywin32
python -m pip install --upgrade pywin32 --force-reinstall
python -c "import pywin32_postinstall; pywin32_postinstall.install()"
```

### Carpeta de locks compartida

```bash
# Se crea automáticamente en:
# C:\proyectos\locks\

# Contiene archivos temporales:
# SSFF_Corte_8AM.lock
# MOVISTAR_Avance_12PM.lock
# etc.
```

---

## 🚀 Uso rápido

### Opción 1: Prueba manual (recomendado primero)

```bash
cd C:\proyectos\SSFF\Reportes_ssff_wsp
python prueba_corte.py
```

Esto:
1. Genera Excel 8AM
2. Captura 3 tablas como PNG
3. Envía a tu número (+51975155264)

**Verifica tu WhatsApp** en ~30 segundos.

### Opción 2: Generar solo Excel

```bash
python generar_corte_ventas.py --hora 10
```

Crea: `CORTE_VENTAS_20260619_100000.xlsx`

### Opción 3: Capturar imágenes de un Excel existente

```bash
python capturar_cortes.py --hora 10 --archivo C:\ruta\CORTE_VENTAS.xlsx --solo-imagenes
```

Captura PNG en: `cortes_imagenes/`

### Opción 4: Capturar + enviar a WhatsApp

```bash
python capturar_cortes.py --hora 10 --destino test
```

Envía a: config.json → cortes_horarios.destinos.test

### Opción 5: Programar ejecuciones automáticas

```powershell
# Como Administrador:
python programar_cortes.py
```

Crea 11 tareas en Task Scheduler (8AM-6PM).

---

## ⚙️ Configuración

### config.json

```json
{
  "cortes_horarios": {
    "destinos": {
      "test": "51975155264@c.us",           ← Tu número (privado)
      "canal": "120363278118591818@g.us",  ← Canal SSFF
      "supervisores_f8": "...",             ← Grupo F8 (opcional)
      "supervisores_m0": "...",             ← Grupo M0 (opcional)
      "jefe": "...",                        ← Jefe (1:1)
      "gerente": "..."                      ← Gerente (1:1)
    },
    "horas": [8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
  }
}
```

**Agregar más destinos:**

```bash
# Obtener ID de un grupo:
python wa_client.py --list-groups

# Copiar ID en config.json bajo destinos
```

---

## 🔒 Sincronización de recursos (Mutex)

### Problema

Múltiples proyectos capturan screenshot simultáneamente:
- SSFF: 8AM, 9AM, 10AM, ... 6PM (11 capturas)
- MOVISTAR: 12PM, 2PM, 4PM, 6PM, CIERRE (5 capturas)
- Otros proyectos: TBD

→ **Conflicto de COM/Windows**: Solo uno puede tener Excel abierto en modo CopyPicture.

### Solución

`screenshot_safe.py` implementa **file-based mutex**:

```python
mgr = ScreenshotManager("SSFF_Corte_8AM")

if mgr.adquirir_lock(timeout=30):
    try:
        # Capturar imagen (exclusivo, sin conflictos)
        capturar_tabla_excel(...)
    finally:
        mgr.liberar_lock()
```

**Cómo funciona:**

1. Intenta crear archivo lock: `C:\proyectos\locks\SSFF_Corte_8AM.lock`
2. Si otro proyecto lo tiene → espera hasta 30 segundos
3. Si otro proyecto libera → adquiere lock y captura
4. Libera lock automáticamente al terminar

**Logs:**

```
✅ Lock adquirido: SSFF_Corte_8AM
⏳ Esperando lock: MOVISTAR_Avance_12PM (edad=5s)
🔓 Lock liberado: SSFF_Corte_8AM
```

---

## 📊 Rangos de tablas

| Tabla | Rango | Filas | Columnas |
|-------|-------|-------|----------|
| GENERAL | A2:H8 | 7 | 8 (A-H) |
| ZONAL | K2:U12 | 11 | 11 (K-U) |
| SUPERVISOR | X2:AH14 | 13 | 11 (X-AH) |

**Nota:** Se capturan filas 2-14 (excluye encabezado fila 1).

---

## 🌐 Envío a WhatsApp

### Antibang (tecnología de envío)

- **Delays aleatorios:** 3-8 segundos entre mensajes
- **Batching:** 5 imágenes + pausa 15-20s
- **Rate limit detection:** Pausa 1h si Meta bloquea
- **Seguridad:** <1% riesgo de baneo

### Destinos soportados

```
test       → 51975155264@c.us          (tu número)
canal      → 120363278118591818@g.us   (Canal SSFF)
supervisores_f8 → Grupo F8 (llenar)
supervisores_m0 → Grupo M0 (llenar)
jefe       → Jefe de proyecto (1:1)
gerente    → Gerente comercial (1:1)
```

### Formato de mensaje

```
🔔 CORTE 8AM
[Imagen: GENERAL]

🔔 CORTE 8AM - ZONAL
[Imagen: ZONAL]

🔔 CORTE 8AM - SUPERVISOR
[Imagen: SUPERVISOR]
```

---

## 📅 Programación automática

### Crear tareas

```bash
# Requiere Admin
python programar_cortes.py
```

Crea 11 tareas:

```
✅ SSFF_Corte_8AM   (8:00 AM)
✅ SSFF_Corte_9AM   (9:00 AM)
✅ SSFF_Corte_10AM  (10:00 AM)
✅ SSFF_Corte_11AM  (11:00 AM)
✅ SSFF_Corte_12PM  (12:00 PM)
✅ SSFF_Corte_1PM   (1:00 PM)
✅ SSFF_Corte_2PM   (2:00 PM)
✅ SSFF_Corte_3PM   (3:00 PM)
✅ SSFF_Corte_4PM   (4:00 PM)
✅ SSFF_Corte_5PM   (5:00 PM)
✅ SSFF_Corte_6PM   (6:00 PM)
```

### Verificar en Task Scheduler

```
Control Panel → Administrative Tools → Task Scheduler
→ Busca "SSFF_Corte"
```

### Modificar horarios

Editar `programar_cortes.py`:

```python
# Línea ~25
HORARIOS = {
    8: "8AM",
    10: "10AM",      # ← Solo 8, 10, 12, 14, 16, 18 (6 horarios)
    12: "12PM",
    # ...
}

# Luego:
python programar_cortes.py
```

---

## 🧪 Troubleshooting

### ❌ "No se pudo adquirir lock"

```
❌ Timeout adquiriendo lock: SSFF_Corte_8AM
```

**Causa:** MOVISTAR u otro proyecto está capturando  
**Solución:** Esperar 30s y reintentar (automático)

### ❌ "Clipboard vacío"

```
❌ Error capturando tabla: Clipboard vacío (no hay imagen)
```

**Causa:**
- Excel no está abierto
- La tabla no está visible
- CopyPicture no funcionó

**Solución:**
```bash
# Probar:
python capturar_cortes.py --hora 10 --solo-imagenes
```

### ❌ "Imagen no encontrada"

```
⚠️  Imagen no encontrada: C:\...\GENERAL_8AM.png
```

**Causa:** Fallo en captura anterior  
**Solución:** Ver logs de capturar_cortes.py

### ❌ "WhatsApp no recibe"

```
❌ Error enviando corte: ...
```

**Causas:**
1. Servidor wa_server.js no está activo
2. config.json no tiene destino válido
3. Rate limit Meta (429)

**Solución:**
```bash
# Terminal 1: Iniciar servidor
node wa_server.js

# Terminal 2: Validar
python test_antibang.py

# Terminal 3: Probar
python prueba_corte.py
```

### ❌ "Task Scheduler no crea tareas"

```
⚠️  Error: Access Denied
```

**Solución:** Ejecutar PowerShell/CMD como **Administrador**

---

## 📝 Logs

Archivos de log:

```
C:\proyectos\SSFF\Reportes_ssff_wsp\
├── logs/
│   ├── distribucion_20260619.log      ← Envíos WhatsApp
│   └── corte_20260619.log             ← Captura
└── cortes_imagenes/
    ├── GENERAL_8AM_20260619_080000.png
    ├── ZONAL_8AM_20260619_080000.png
    └── SUPERVISOR_8AM_20260619_080000.png
```

Ver logs en tiempo real:

```bash
# PowerShell
Get-Content -Path logs/distribucion_*.log -Tail 20 -Wait

# Bash
tail -f logs/distribucion_*.log
```

---

## 🎯 Checklist de implementación

- [ ] Instalar dependencias: `pip install ...`
- [ ] Setup pywin32: `python -m pip install --upgrade pywin32`
- [ ] Probar manual: `python prueba_corte.py`
- [ ] Verificar imágenes en WhatsApp personal
- [ ] Agregar destinos en config.json (canales, grupos)
- [ ] Crear tareas automáticas: `python programar_cortes.py`
- [ ] Verificar en Task Scheduler
- [ ] Monitorear logs los primeros días

---

## 🚀 Próximos pasos

### Corto plazo (hoy)
1. ✅ Ejecutar `prueba_corte.py`
2. ✅ Verificar WhatsApp
3. ✅ Llenar config.json con destinos

### Mediano plazo (esta semana)
4. ✅ Programar tareas: `programar_cortes.py`
5. ✅ Validar primeros envíos automáticos

### Integración multi-proyecto
6. Sincronizar con MOVISTAR (usa mismo mutex)
7. Agregar a otros proyectos que necesiten captura

---

## 📞 Soporte

**Si algo falla:**

1. Ver logs: `tail -f logs/*.log`
2. Ejecutar con `--solo-imagenes` para aislar problema
3. Verificar Task Scheduler → Properties → History
4. Revisar credenciales .env

**Comandos útiles:**

```bash
# Ver Excel abiertos
wmic process list brief | find "EXCEL"

# Eliminar locks stale (si es necesario)
rm C:\proyectos\locks\*.lock

# Testear screenshot sin envío
python capturar_cortes.py --hora 8 --solo-imagenes

# Listar grupos WhatsApp
python wa_client.py --list-groups
```

---

**Versión:** 1.0  
**Última actualización:** 2026-06-19  
**Autor:** Claude Code
