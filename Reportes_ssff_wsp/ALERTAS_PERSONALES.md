# 🚨 ALERTAS PERSONALES A WHATSAPP — Documentación

**Creado:** 22 de Junio de 2026  
**Estado:** ✅ INTEGRADO EN ejecutar_corte_orquestado.py

---

## 📌 **¿QUÉ ES?**

Sistema de notificaciones a tu WhatsApp personal (**no al grupo de gerencia**) para:

- 🚨 **Datos en Cero:** Cuando SQL devuelve S/ 0.00 (alerta crítica)
- ✅ **Éxito:** Cuando el corte se envía exitosamente (opcional)
- ❌ **Error:** Cuando hay timeout o problema en generación (opcional)

---

## 🔧 **CONFIGURACIÓN (Ya está hecha)**

Tu número personal ya está en `config.json`:

```json
{
  "my_number": "51975155264@c.us",
  ...
}
```

✅ **Sin cambios necesarios**

---

## 🚀 **USO AUTOMÁTICO**

### **Cuando se detectan datos en cero:**

El sistema **automáticamente**:
1. Detecta que SQL devolvió S/ 0.00
2. Intenta 1 reintento en 15 minutos
3. Si SIGUE siendo cero → **Envía alerta a tu WhatsApp**
4. Mensaje incluye:
   - Hora del corte
   - Fecha
   - Posibles causas
   - Próximas acciones

### **Ejemplo de mensaje (Datos en Cero):**
```
🚨 ALERTA CRÍTICA — DATOS EN CERO

Corte: 10AM
Fecha: 2026-06-22

❌ SQL devolvió S/ 0.00 en todas las categorías.

Posibles causas:
• Data center en limpieza de pedidos
• Problema en la BD
• Desconexión de red

Acción:
Se reintentará en 15 minutos automáticamente.
Si persiste, contactar a Data Engineer.
```

---

## ✅ **USO MANUAL (Opcional)**

### **Para avisar éxito:**

Descomenta la línea 248 en `ejecutar_corte_orquestado.py`:

```python
# Línea 248-249: Descomenta si deseas notificaciones de éxito
if alertas and alertas.disponible:
    fecha_str = datetime.now().strftime('%Y-%m-%d')
    alertas.enviar_alerta_exito(hora=hora, fecha=fecha_str)
```

### **Ejemplo de mensaje (Éxito):**
```
✅ CORTE EXITOSO

Corte: 10AM
Fecha: 2026-06-22
Total: S/ 85,500.50

📊 Reporte generado y enviado a Canal SSFF.
🔔 Gerencia notificada.
```

---

## 🧪 **TESTEAR LAS ALERTAS**

### **Test 1: Mensaje de prueba**

```bash
python test_alertas.py
```

Envía un mensaje simple a tu WhatsApp personal.

### **Test 2: Simular alerta datos cero**

```bash
python test_alertas.py --tipo datos_cero
```

Envía la alerta que recibirías si hay datos en cero.

### **Test 3: Simular alerta éxito**

```bash
python test_alertas.py --tipo exito
```

Envía la alerta de corte exitoso.

### **Test 4: Simular alerta error**

```bash
python test_alertas.py --tipo error
```

Envía la alerta si hay error en generación.

---

## 📊 **FLUJO DE ALERTAS (Datos en Cero)**

```
[ejecutar_corte_orquestado.py]
    ↓
[generar_corte_html.py]
    ↓
    ¿Datos = 0?
        ├─ NO → Continúa normalmente (sin alerta)
        └─ SÍ → detecta_datos_cero()
              ↓
              Espera 15 minutos
              ↓
              Reintenta
              ↓
              ¿Datos SIGUEN siendo 0?
                  ├─ NO → OK, continúa
                  └─ SÍ → ALERTA A WHATSAPP PERSONAL
                         (AlertasPersonales.enviar_alerta_datos_cero)
```

---

## 📝 **ARCHIVOS IMPLICADOS**

| Archivo | Función |
|---------|---------|
| **alertas_personales.py** | Módulo principal de alertas |
| **test_alertas.py** | Script para testear |
| **ejecutar_corte_orquestado.py** | Integración automática |
| **config.json** | Configuración (my_number) |

---

## 🔒 **SEGURIDAD**

✅ **Solo envía alertas en casos especiales:**
- Datos en cero (problema real)
- Error en generación (problema real)
- Éxito manual (opcional, si lo activas)

❌ **NO envía en cortes normales**
- Sin spam
- Solo notificaciones importantes

---

## 🎯 **CASOS DE USO**

| Caso | Alerta | A Quién | Acción |
|------|--------|---------|--------|
| **Corte normal OK** | ❌ No | — | Gerencia solo recibe PNG en grupo |
| **Datos en cero** | ✅ SÍ | Tu WhatsApp personal | Contactar Data Engineer |
| **Error en SQL** | ✅ SÍ | Tu WhatsApp personal | Revisar logs |
| **Éxito (si activas)** | ✅ SÍ | Tu WhatsApp personal | Confirmación informativa |

---

## ⚙️ **CÓMO ACTIVAR ALERTA DE ÉXITO**

Si deseas recibir confirmación cada vez que un corte se envía correctamente:

### 1. Abre `ejecutar_corte_orquestado.py`
### 2. Ve a línea ~248
### 3. Descomenta estas líneas:

```python
# OPCIONAL: Enviar alerta de éxito a WhatsApp personal
if alertas and alertas.disponible:
    fecha_str = datetime.now().strftime('%Y-%m-%d')
    alertas.enviar_alerta_exito(hora=hora, fecha=fecha_str)
```

### 4. Queda:

```python
# OPCIONAL: Enviar alerta de éxito a WhatsApp personal
if alertas and alertas.disponible:
    fecha_str = datetime.now().strftime('%Y-%m-%d')
    alertas.enviar_alerta_exito(hora=hora, fecha=fecha_str)
```

✅ **Listo — Recibirás alertas de éxito también**

---

## 📞 **SOPORTE**

Si no recibes alertas:

1. **Verifica `config.json`:** `my_number` debe estar presente
2. **Testea:** `python test_alertas.py`
3. **Revisa logs:** `Reportes_ssff_wsp/logs/`
4. **Verifica WhatsApp server:** `http://localhost:8002` debe estar corriendo

---

## 🚀 **RESUMEN**

✅ **Datos en cero → Alerta automática a tu WhatsApp**  
✅ **Fácil testear con test_alertas.py**  
✅ **Opcional: Activar alertas de éxito**  
✅ **Sin spam en cortes normales**  
✅ **Solo alertas críticas**

