# 🔒 Estrategia de Seguridad — Automatización WhatsApp

## El riesgo: Por qué WhatsApp banea

WhatsApp detecta patrones de bot:
- **Múltiples sesiones** del mismo número simultaneamente
- **Envíos masivos rápidos** sin intervalos
- **Horas anómalas** (envíos a las 3 AM, etc.)
- **Patrones repetitivos** (mismo mensaje a N contactos en segundos)
- **IP diferente** o comportamiento no humano

---

## ✅ Recomendaciones de seguridad

### 1. **Separar por número (CRÍTICO)**

**NUNCA uses el mismo número para:**
- ✅ MOVISTAR (automatización activa en puerto 8002)
- ✅ SSFF (automatización futura en puerto 8003)

**Solución:** Crea una **cuenta WhatsApp Business** separada para SSFF.

| Proyecto | Número | Puerto | Uso | Riesgo |
|----------|--------|--------|-----|--------|
| MOVISTAR | Tu número personal | 8002 | Envíos automáticos | 🔴 ALTO si unes con SSFF |
| SSFF | Número Business (nuevo) | 8003 | Envíos automáticos | 🟢 BAJO (aislado) |

### 2. **Horarios seguros de envío**

Envía mensajes **solo en horarios laborales**:
- ✅ Lunes-Viernes: 08:00 - 18:00
- ❌ Noches, fines de semana, festivos

```python
from datetime import datetime

def es_horario_seguro():
    ahora = datetime.now()
    return ahora.weekday() < 5 and 8 <= ahora.hour < 18
```

### 3. **Intervalos entre envíos (IMPORTANTE)**

**Regla de oro:** Nunca envíes 2 mensajes en menos de 3-5 segundos.

```python
import time

def enviar_con_intervalo(destinatarios, mensaje):
    for dest in destinatarios:
        enviar_mensaje(dest, mensaje)
        time.sleep(5)  # Espera 5 segundos entre envíos
```

### 4. **Límites diarios por número**

**Recomendado:**
- MOVISTAR: máx. 100-150 mensajes/día
- SSFF: máx. 20-30 mensajes/día (es un grupo, no masivo)

```python
MAX_MENSAJES_DIARIOS_SSFF = 30
```

### 5. **Envíos a grupos es más seguro que a individuales**

✅ **Envío a grupo:** Menos riesgo (1 mensaje a N personas)
❌ **Envío individual:** Mayor riesgo (N mensajes separados)

**SSFF está bien:** Solo envía al grupo "Canal SSFF 2026", no a individuos.

### 6. **Logging y monitoreo**

Registra cada envío para detectar anomalías:

```python
import json
from datetime import datetime

def log_envio(destino, mensaje, estado, timestamp=None):
    log_entry = {
        "timestamp": timestamp or datetime.now().isoformat(),
        "destino": destino,
        "mensaje": mensaje[:50] + "..." if len(mensaje) > 50 else mensaje,
        "estado": estado  # "OK", "ERROR", "BLOCKED"
    }
    
    with open("logs/wa_envios.json", "a") as f:
        f.write(json.dumps(log_entry) + "\n")
```

### 7. **Detectar signos de advertencia**

Si ves estos síntomas, **detén inmediatamente** la automatización:

- ⚠️ Mensajes que no se envían (pero no hay error explícito)
- ⚠️ Demoras anormales (>10 segundos por mensaje)
- ⚠️ El servidor desconecta solo
- ⚠️ Error 401 o 403 en respuestas

```python
SIGNOS_PELIGRO = [401, 403, 429]  # 429 = Too many requests

if response.status_code in SIGNOS_PELIGRO:
    log_envio(dest, msg, "BLOCKED")
    print("⚠️ ALERTA: Posible restricción de WhatsApp")
    break  # Detén el ciclo
```

---

## 🔧 Configuración segura para SSFF

### config.json — Agregar límites

```json
{
  "my_number": "51XXXXXXXXX@c.us",
  "grupos": {
    "Canal_SSFF_2026_Gestion": {
      "group_id": "120363278118591818@g.us",
      "max_mensajes_diarios": 30,
      "intervalo_minimo_segundos": 5,
      "horarios_permitidos": {
        "inicio": "08:00",
        "fin": "18:00",
        "dias": [0, 1, 2, 3, 4]
      }
    }
  }
}
```

### wa_client.py — Agregar validaciones

```python
def puede_enviar(config, grupo_id):
    """Verifica si es seguro enviar ahora"""
    
    # 1. ¿Es horario seguro?
    from datetime import datetime
    ahora = datetime.now()
    horario = config['grupos'][grupo_id]['horarios_permitidos']
    if ahora.weekday() not in horario['dias']:
        return False, "No es día laboral"
    
    hora_inicio = int(horario['inicio'].split(':')[0])
    hora_fin = int(horario['fin'].split(':')[0])
    if not (hora_inicio <= ahora.hour < hora_fin):
        return False, f"Fuera de horario ({hora_inicio}h-{hora_fin}h)"
    
    # 2. ¿Se alcanzó límite diario?
    enviados_hoy = contar_envios_hoy(grupo_id)
    max_diarios = config['grupos'][grupo_id]['max_mensajes_diarios']
    if enviados_hoy >= max_diarios:
        return False, f"Límite diario alcanzado ({max_diarios})"
    
    return True, "OK"

def enviar_seguro(grupo_id, mensaje):
    """Envía respetando límites"""
    config = cargar_config()
    
    puede, razon = puede_enviar(config, grupo_id)
    if not puede:
        print(f"❌ No se puede enviar: {razon}")
        log_envio(grupo_id, mensaje, "RECHAZADO", razon)
        return False
    
    try:
        response = enviar_mensaje(grupo_id, mensaje)
        log_envio(grupo_id, mensaje, "OK")
        return True
    except Exception as e:
        log_envio(grupo_id, mensaje, f"ERROR: {str(e)}")
        return False
```

---

## 📋 Checklist de implementación

- [ ] **Obtener número WhatsApp Business separado** para SSFF (no usar el de MOVISTAR)
- [ ] **Configurar `config.json`** con límites diarios y horarios
- [ ] **Agregar validaciones** en `wa_client.py` (horario, límite, intervalo)
- [ ] **Implementar logging** de todos los envíos
- [ ] **Configurar alertas** para síntomas de bloqueo
- [ ] **Documentar en generar_cuota.py** dónde se llama a WhatsApp
- [ ] **Mantener registro** de cuántos mensajes envían MOVISTAR + SSFF (< 200/día total)

---

## ⚡ Resumen rápido

**Para NO banear tu número:**

1. ✅ Usa número **diferente** para SSFF (Business)
2. ✅ Envía **solo en horario laboral** (8-18h, lunes-viernes)
3. ✅ Espera **5+ segundos** entre mensajes
4. ✅ Máximo **30 mensajes/día** por proyecto
5. ✅ Usa **grupos**, no envíos individuales
6. ✅ **Registra** todo en logs
7. ✅ **Monitorea** signos de alerta

---

**Próxima acción:** ¿Ya tienes acceso a una cuenta WhatsApp Business para SSFF, o necesitas crear una?
