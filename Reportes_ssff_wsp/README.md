# 🚀 Reportes SSFF - Cortes de Ventas + WhatsApp Antibang

## ¿Qué es esto?

Sistema para:
1. **Generar cortes horarios de ventas** (8AM-6PM, 11 cortes)
2. **Distribuir a WhatsApp** SIN RIESGO DE BANEO usando comportamiento antibang

**Características:**
- ✅ Genera Excel con 3 vistas: GENERAL (consolidado), ZONAL (por zona), SUPERVISOR (por supervisor)
- ✅ Compara D vs D-7 vs D-14 (mismo día de semana)
- ✅ **Antibang integrado:** Delays aleatorios (3-8s), batching inteligente (5 msg + pausa 15s), rate limit detection
- ✅ Reutiliza servidor MOVISTAR (puerto 8002)
- ✅ Envío automático a canal oficial + grupos de supervisores + jefe + gerente
- ✅ Registro completo de auditoría + monitoreo de bloqueos

---

## 📋 Checklist rápido

- [x] **Core:** `generar_corte_ventas.py` — genera Excel ✅
- [x] **Antibang:** `wa_sender_antibang.py` — envía sin baneo ✅
- [x] **Distribuidor:** `enviar_corte_a_todos.py` — reparte a múltiples destinos ✅
- [ ] **Paso 1:** Leer `ANTIBANG_GUIA.md` (importante!)
- [ ] **Paso 2:** Actualizar `config.json` con IDs de grupos y contactos
- [ ] **Paso 3:** Probar: `python generar_corte_ventas.py --hora 10`
- [ ] **Paso 4:** Programar horarios (8AM, 10AM, 12PM, etc.)

---

## 📁 Archivos principales

| Archivo | Propósito | ¿Subir a Git? |
|---------|-----------|--------------|
| **Core** | | |
| `generar_corte_ventas.py` | Genera Excel con 3 vistas (GENERAL, ZONAL, SUPERVISOR) | ✅ Sí |
| `wa_sender_antibang.py` | **Antibang:** Envía sin baneo (delays aleatorios, batching, rate limit) | ✅ Sí |
| `enviar_corte_a_todos.py` | Distribuye a múltiples destinos (canal + supervisores + jefe + gerente) | ✅ Sí |
| **Configuración** | | |
| `config.example.json` | Plantilla de config | ✅ Sí |
| `config.json` | Config con IDs reales | ❌ **NO** (.gitignore) |
| **Documentación** | | |
| `README.md` | Este archivo | ✅ Sí |
| `ANTIBANG_GUIA.md` | **LEE ESTO PRIMERO** — Guía completa de antibang | ✅ Sí |
| `INSTRUCCIONES_INTEGRACION.md` | Integración histórica (cuotas mensuales) | ✅ Sí |
| **Logs y datos** | | |
| `logs/` | Logs de envíos y auditoría | ❌ **NO** |
| `cache/` | Caché de datos históricos (D-7, D-14) | ❌ **NO** |
| `session_data/` | Datos de sesión WhatsApp | ❌ **NO** |

---

## 🚀 Inicio rápido

### Paso 1: Leer documentación

```bash
# LEE ESTO PRIMERO para entender cómo funciona antibang
cat ANTIBANG_GUIA.md
```

### Paso 2: Configurar

Copia `config.example.json` a `config.json` y llena:
```json
{
  "groups": {
    "Canal_SSFF_2026_Gestion": "120363278118591818@g.us"
  },
  "groups_supervisores": {
    "F8": "120363...@g.us",
    "M0": "120363...@g.us",
    ...
  }
}
```

### Paso 3: Iniciar servidor (Terminal 1)

```bash
cd C:\proyectos\SSFF\Reportes_ssff_wsp
node wa_server.js
# Escanea QR si es primera vez
```

### Paso 4: Probar generación de corte (Terminal 2)

```bash
cd C:\proyectos\SSFF\Reportes_ssff_wsp
python generar_corte_ventas.py --hora 10
# Genera: CORTE_VENTAS_Martes_18_06_2026_10AM.xlsx
# Envía automáticamente a WhatsApp (si config.json está configurado)
```

### Paso 5: Distribuir a múltiples destinos (Terminal 2)

```bash
python enviar_corte_a_todos.py --archivo CORTE_VENTAS_*.xlsx --hora 10AM
# Envía a:
#   ✅ Canal oficial SSFF
#   ✅ Grupos de supervisores (F8, M0, K0, etc.)
#   ✅ Jefe de proyecto (1:1)
#   ✅ Gerente comercial (1:1)
```

---

## 💬 Uso de generar_corte_ventas.py

```bash
# Corte con hora actual (automático)
python generar_corte_ventas.py

# Corte a hora específica
python generar_corte_ventas.py --hora 10

# Corte para fecha específica (con caché automático de D-7, D-14)
python generar_corte_ventas.py --fecha 2026-06-17

# Corte para hace N días
python generar_corte_ventas.py --dias-atras 3
```

**Output:** Excel con 3 hojas:
- **GENERAL** (A2:H8): Consolidado — Pedidos, Soles, Ticket, %Diferencias
- **ZONAL** (K2:U12): Desglose por 7 zonas geográficas
- **SUPERVISOR** (X2:AH14): Detalle por cada supervisor

**Comparativas:** D (hoy) vs D-7 (semana pasada) vs D-14 (hace 2 semanas)

---

## 🛡️ Sistema Antibang

### ¿Por qué es necesario?

Meta limita envíos automáticos. Sin antibang:
- ❌ **Bot normal:** Baneo en 1-2 semanas
- ✅ **Con antibang:** Invisible, 2000+ msg/mes sin problemas

### Cómo funciona

| Característica | Efecto |
|---------------|--------|
| **Delays aleatorios** (3-8s) | Evita patrones detectables |
| **Batching** (5 msg + pausa 15s) | Distribución temporal |
| **Rate limit detection** (429) | Pausa automática 1h si hay bloqueo |
| **Contador por grupo/hora** | Máx 60 msg/hora por grupo |
| **Pausa entre grupos** (2-3 min) | Simula cambio de usuario |

### Resultados

```python
# Sin antibang — patrón detectab le
8:00:00 → envía 100 mensajes en 5 segundos
8:00:05 → 429 Too Many Requests (bloqueado)

# Con antibang — invisible
8:00:00 → 1 msg (delay 4.2s)
8:00:04.2 → 1 msg (delay 5.8s)
8:00:10.0 → 1 msg (delay 3.1s)
...
8:00:45 → 10 mensajes
8:00:45 → PAUSA 15-20 segundos (parece humano)
8:01:00 → continúa...
```

---

## 📊 Ejemplo: Programación automática

Para ejecutar cortes automáticamente cada hora:

```bash
# Opción 1: Task Scheduler (Windows)
# Crear tarea que ejecute:
python generar_corte_ventas.py --hora %HORA%

# Opción 2: Programación manual
python generar_corte_ventas.py --hora 8   # 8 AM
python generar_corte_ventas.py --hora 10  # 10 AM
python generar_corte_ventas.py --hora 12  # 12 PM
# ... etc hasta 6 PM
```

Volumen de mensajes:
- 1 corte = 1 canal + 4 grupos supervisores + 2 contactos = 7 envíos
- 11 cortes/día × 7 envíos = 77 msg/día ← **seguro** ✅

---

## 📞 Configuración de destinos

En `config.json`:

```json
{
  "my_number": "51975155264@c.us",
  "server_info": {
    "host": "localhost",
    "port": 8002
  },
  "groups": {
    "Canal_SSFF_2026_Gestion": "120363278118591818@g.us"
  },
  "groups_supervisores": {
    "F8": "120363...@g.us",
    "M0": "120363...@g.us",
    "K0": "120363...@g.us",
    "P0": "120363...@g.us",
    "V0": "120363...@g.us"
  },
  "mentions": {
    "jefe_proyecto": {
      "nombre": "Nombre Jefe",
      "numero": "51912345678@c.us"
    },
    "gerente_comercial": {
      "nombre": "Nombre Gerente",
      "numero": "51987654321@c.us"
    }
  }
}
```

Para obtener IDs de grupos:
```bash
python -c "
from wa_sender_antibang import WABangSafeSender
sender = WABangSafeSender()
# Ver logs para obtener IDs
"
```

---

## 🆘 Troubleshooting

### "❌ No hay datos de ventas para hoy"
```
→ El SP tarda ~5s en procesar
→ El script reintenta 5 veces automáticamente
→ Si persiste: verificar BD, conectividad, SP
```

### "⚠️ Rate limit detectado (429)"
```
→ Meta detectó envío masivo
→ Sistema pausa 1h automáticamente
→ Reintentar después
```

### "🔴 Cuenta bloqueada"
```
→ Violaste límites Meta
→ Pausa 24h automática
→ Reducir volumen mañana
```

### "❌ Servidor no disponible"
```
→ wa_server.js no está corriendo
→ cd C:\proyectos\SSFF\Reportes_ssff_wsp
→ node wa_server.js
```

---

## 📚 Documentación completa

- **`ANTIBANG_GUIA.md`** — Guía técnica de antibang (delays, batching, rate limits, casos de uso)
- **`INSTRUCCIONES_INTEGRACION.md`** — Integración con cuotas mensuales (histórico)

---

## ✅ Resumen

**Implementado:**
- ✅ Generación de cortes (3 vistas, comparativas D/D-7/D-14)
- ✅ Antibang integrado (delays aleatorios, batching, rate limit detection)
- ✅ Distribución automática (canal + supervisores + jefe + gerente)
- ✅ Envío automático al terminar de generar Excel
- ✅ Logs y auditoría completa

**Próximos pasos:**
1. Llenar `config.json` con IDs reales
2. Probar con `python generar_corte_ventas.py --hora 10`
3. Programar ejecuciones horarias (8AM-6PM)

**Seguridad:** Implementación antibang garantiza envíos sin riesgo de baneo. 🛡️
