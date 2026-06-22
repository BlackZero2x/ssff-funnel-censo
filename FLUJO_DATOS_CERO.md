# 🔄 FLUJO DE MANEJO DE DATOS EN CERO — Opción A Implementada

**Estado:** ✅ IMPLEMENTADO  
**Fecha:** 22 de Junio de 2026

---

## 📊 DIAGRAMA DE FLUJO

```
┌─────────────────────────────────────────────────────────────────┐
│        [ejecutar_corte_orquestado.py --hora 11]                 │
└────────────────────────┬────────────────────────────────────────┘
                         │
                         ↓
        ┌────────────────────────────────┐
        │ [1] Generar Excel              │
        │ generar_corte_ventas.py        │
        │ Timeout: 2 minutos             │
        └───────────┬──────────────────┘
                    │
                    ↓
        ┌────────────────────────────────┐
        │ ¿Returncode = 0?               │
        └───────────┬──────────────────┘
                    │
        ┌───────────┴──────────┐
        │                      │
       NO                     SÍ
        │                      │
        ↓                      ↓
   [FAIL]              ┌──────────────────┐
   Reintenta           │ ¿Datos cero?     │
   Normal              │ (validar flag)   │
                       └────────┬─────────┘
                                │
                        ┌───────┴──────────┐
                        │                  │
                       NO                 SÍ
                        │                  │
                        ↓                  ↓
                     [OK!]          ┌────────────────────────────┐
                     Continúa       │ intentos_cero = 0?         │
                     a Capturar     └────────┬───────────────────┘
                                            │
                                    ┌───────┴──────────┐
                                    │                  │
                                   SÍ                 NO
                                    │                  │
                                    ↓                  ↓
                        ┌────────────────────┐  ┌─────────────────────┐
                        │ intentos_cero += 1 │  │ [CRÍTICO] ABORTAR   │
                        │ intentos_cero = 1  │  │ Datos en cero 2x    │
                        └────────┬───────────┘  │ Requiere intervención│
                                 │             │ exit(2)             │
                                 ↓             └─────────────────────┘
                        ┌────────────────────┐
                        │ Esperar 15 minutos │
                        │ (900 segundos)     │
                        └────────┬───────────┘
                                 │
                                 ↓
                        ┌────────────────────┐
                        │ Reintentar Intento │
                        │ (continue loop)    │
                        └────────┬───────────┘
                                 │
                                 ↓
                        ┌────────────────────┐
                        │ ¿Datos cero OTRA   │
                        │ vez?               │
                        └────────┬───────────┘
                                 │
                        ┌────────┴──────────┐
                        │                   │
                       SÍ                  NO
                        │                   │
                        ↓                   ↓
        ┌──────────────────────────┐   [OK!]
        │ [CRÍTICO] ABORTAR        │   Continúa
        │ Datos en cero PERSISTENTES│   a Capturar
        │ Contactar Data Engineer  │
        │ exit(2)                  │
        └──────────────────────────┘
```

---

## 🔢 ESCENARIOS Y COMPORTAMIENTO

### **Escenario 1: Datos Normales (S/ > 0)**
```
Intento 1: generar_corte_ventas.py ✓
           ¿Datos = 0? NO
           ✓ Continúa a Capturar
           ✓ Envía a WhatsApp
           
Resultado: SUCCESS (exit 0)
```

### **Escenario 2: Datos en Cero (Primer Intento)**
```
Intento 1: generar_corte_ventas.py ✓
           ¿Datos = 0? SÍ
           intentos_cero = 1
           
           Espera 15 minutos...
           
Intento 2 (después de 15 min): generar_corte_ventas.py ✓
                                ¿Datos = 0? ?
                                
           Opción A: SÍ → Datos en cero persistentes → ABORT (exit 2)
           Opción B: NO → OK, continúa a Capturar → SUCCESS (exit 0)
           
Resultado: SUCCESS (0) o CRÍTICO (2)
```

### **Escenario 3: Error de Conexión (Normal)**
```
Intento 1: generar_corte_ventas.py ✗ (timeout/conexión)
           Espera 5 segundos
           
Intento 2: generar_corte_ventas.py ✗ (timeout/conexión)
           Espera 5 segundos
           
Intento 3: generar_corte_ventas.py ✗ (timeout/conexión)
           
Resultado: FAIL (exit 1)
```

### **Escenario 4: Datos Cero + Error Normal**
```
Intento 1: generar_corte_ventas.py ✓
           ¿Datos = 0? SÍ
           intentos_cero = 1
           
           Espera 15 minutos...
           
Intento 2: generar_corte_ventas.py ✗ (timeout/conexión)
           Espera 5 segundos
           
Intento 3: generar_corte_ventas.py ✗ (timeout/conexión)
           
Resultado: FAIL (exit 1) — Los datos en cero ya se intentaron
```

---

## 📝 LOGS ESPERADOS

### **Caso 1: Éxito Normal**
```
2026-06-22 11:00:00 - [OK] CORTE 11:00 COMPLETADO ✓
```

### **Caso 2: Primer Cero → Éxito en Reintento**
```
2026-06-22 11:00:07 - [ALERTA] Datos en cero. Esperando 15 minutos antes de reintentar...
2026-06-22 11:00:07 - Si sigue siendo cero, se abortará (requiere intervención manual)

[... 15 minutos de espera ...]

2026-06-22 11:15:07 - [Intento 2] generar_corte_ventas.py...
2026-06-22 11:15:37 - [OK] Datos válidos (S/ 45,200.50)
2026-06-22 11:15:45 - [OK] CORTE 11:00 COMPLETADO ✓
```

### **Caso 3: Datos en Cero Persistentes → ABORT**
```
2026-06-22 11:00:07 - [ALERTA] Datos en cero. Esperando 15 minutos antes de reintentar...
2026-06-22 11:00:07 - Si sigue siendo cero, se abortará (requiere intervención manual)

[... 15 minutos de espera ...]

2026-06-22 11:15:07 - [Intento 2] generar_corte_ventas.py...
2026-06-22 11:15:37 - [ALERTA] Datos en cero — reintentando en 15 min
2026-06-22 11:15:37 - [CRÍTICO] DATOS EN CERO PERSISTENTES
2026-06-22 11:15:37 - Intento 1: Ceros
2026-06-22 11:15:37 - Espera: 15 minutos
2026-06-22 11:15:37 - Intento 2: SIGUE SIENDO CEROS
2026-06-22 11:15:37 - Aborting: Contactar a Data Engineer

exit(2) ← Código especial para ceros persistentes
```

---

## 🔍 CÓMO DIFERENCIAR LOS ERRORES

| Situación | Log Final | Exit Code | Acción |
|-----------|-----------|-----------|--------|
| Éxito normal | `[OK] CORTE completado ✓` | 0 | Ninguna |
| Fallo temporal (timeout, conexión) | `[FAIL] Todos los 3 intentos fallaron` | 1 | Reintentar manualmente |
| Datos en cero persistentes | `[CRÍTICO] DATOS EN CERO PERSISTENTES` | 2 | 🚨 Contactar Data Engineer |

---

## 💻 TESTING MANUAL

### **Test 1: Simular éxito normal**
```bash
python ejecutar_corte_orquestado.py --hora 10
# Verificar: exit 0, logs muestran [OK]
```

### **Test 2: Ver comportamiento con ceros**
```bash
# Modificar generar_corte_ventas.py temporalmente para devolver ceros
# Luego ejecutar:
python ejecutar_corte_orquestado.py --hora 11 --max-reintentos 3
# Verificar: Espera 15 min, reintenta, luego abort con exit 2
```

### **Test 3: Ver logs en tiempo real**
```bash
tail -f Reportes_ssff_wsp/logs/cortes_20260622.log
```

---

## 📊 VENTAJAS DE LA OPCIÓN A

✅ **No entra en loop infinito**  
✅ **Detecta problema real vs. problema temporal**  
✅ **Solo espera 15 min una sola vez**  
✅ **Fácil de monitorear con exit codes**  
✅ **Mensaje claro para Data Engineer**  

---

## 🔐 INTEGRACIÓN CON ALERTAS (Futuro)

Cuando implementes alertas, puedes usar los exit codes:

```bash
# En PowerShell Task Scheduler
python ejecutar_corte_orquestado.py --hora 11

if ($LASTEXITCODE -eq 2) {
    # Enviar SMS/email de alerta crítica
    Send-Alert "CRÍTICO: Datos en cero persistentes en corte 11AM"
}
elseif ($LASTEXITCODE -eq 1) {
    # Enviar SMS/email de alerta normal
    Send-Alert "WARN: Corte 11AM falló después de reintentos"
}
```

---

**Documento:** FLUJO_DATOS_CERO.md  
**Implementación:** ✅ COMPLETADA EN CÓDIGO  
**Próximo Paso:** Testear y validar en producción
