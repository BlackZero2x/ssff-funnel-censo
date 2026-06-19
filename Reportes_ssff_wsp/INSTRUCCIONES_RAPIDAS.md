# ⚡ Instrucciones Rápidas — Captura de Cortes SSFF

**Para empezar en 5 minutos.**

---

## 🚀 Paso 1: Instalar (2 min)

```bash
cd C:\proyectos\SSFF\Reportes_ssff_wsp

pip install openpyxl pandas numpy xlwings pillow pywin32 requests

python -m pip install --upgrade pywin32
```

---

## 🧪 Paso 2: Probar (3 min)

```bash
python prueba_corte.py
```

**Qué pasa:**
- ✅ Genera Excel 8AM
- ✅ Captura 3 tablas como PNG
- ✅ Envía a tu WhatsApp (+51975155264)

**Verifica en tu teléfono:** Deberías recibir 3 imágenes en ~30 segundos.

Si ✅ recibes las 3 imágenes → **Va a funcionar**

---

## ⚙️ Paso 3: Configurar (1 min)

Abre `config.json`:

```json
{
  "cortes_horarios": {
    "destinos": {
      "test": "51975155264@c.us",           ← YA ESTÁ TU NÚMERO
      "canal": "120363278118591818@g.us",  ← YA ESTÁ CANAL SSFF
      "supervisores_f8": "",                ← Llenar con ID real
      "supervisores_m0": ""                 ← Llenar con ID real
    }
  }
}
```

**Para obtener IDs de grupos:**

```bash
python wa_client.py --list-groups
```

Copia/pega en config.json.

---

## 📅 Paso 4: Programar automático (1 min)

```bash
# COMO ADMINISTRADOR (PowerShell/CMD as Admin)
python programar_cortes.py
```

Responde "s" si te pregunta sobre tareas anteriores.

**Resultado:** 11 tareas creadas en Task Scheduler (8AM-6PM)

✅ **Listo. Se ejecuta automáticamente todos los días.**

---

## 📋 Verifica que funciona

### En Task Scheduler

```
Control Panel → Administrative Tools → Task Scheduler
→ Busca "SSFF_Corte"
```

Debería haber 11 tareas:
```
✅ SSFF_Corte_8AM
✅ SSFF_Corte_9AM
✅ SSFF_Corte_10AM
... (11 total)
```

### Logs en tiempo real

```bash
# Ver envíos WhatsApp
tail -f logs/distribucion_*.log

# Ver capturas
tail -f logs/corte_*.log
```

---

## 🆘 Si algo falla

### ❌ "Clipboard vacío" o no captura

```bash
python capturar_cortes.py --hora 10 --solo-imagenes
```

Si falla, significa Excel no está abierto. Revisar que generar_corte_ventas.py generó el .xlsx

### ❌ "No recibe en WhatsApp"

```bash
# Verificar servidor
node wa_server.js

# En otra terminal
python test_antibang.py
```

Si falla test → problema con servidor WhatsApp

### ❌ "Lock timeout"

```
❌ Timeout adquiriendo lock: SSFF_Corte_8AM
```

MOVISTAR u otro proyecto está capturando. Espera 30s y reintenta (automático).

---

## 📊 Verificar volumen

- **Cortes/día:** 11 (8AM-6PM)
- **Imágenes/corte:** 3 (GENERAL, ZONAL, SUPERVISOR)
- **Total imágenes/día:** 33
- **Mensajes WhatsApp/día:** 33
- **Límite Meta:** ~300 (89% de margen) ✅

---

## 📞 Archivos importantes

| Archivo | Qué es | Cuándo leerlo |
|---------|--------|--------------|
| `prueba_corte.py` | Script de test rápido | Primero (ahora) |
| `config.json` | Tu configuración | Después de test |
| `CAPTURA_CORTES_GUIA.md` | Guía técnica completa | Si necesitas detalles |
| `RESUMEN_CAPTURA_CORTES.md` | Resumen ejecutivo | Para resumen |

---

## ✅ Checklist

- [ ] **Instalé dependencias** (`pip install ...`)
- [ ] **Ejecuté prueba** (`python prueba_corte.py`)
- [ ] **Recibí 3 imágenes en WhatsApp** ✅
- [ ] **Configuré destinos** (config.json)
- [ ] **Programé tareas** (`python programar_cortes.py` como Admin)
- [ ] **Verifiqué en Task Scheduler** (11 tareas visibles)
- [ ] **Monitorearé logs** los primeros días

---

## 🎉 Listo

**Ya está. Sistema automático 24/7.**

Cada día:
- ⏰ 8:00 AM → Genera Excel + Captura + Envía
- ⏰ 9:00 AM → Captura + Envía
- ⏰ ... (hasta 6 PM)
- ⏰ 6:00 PM → Captura + Envía

**Sin intervención manual.**

---

**Dudas?** Lee `CAPTURA_CORTES_GUIA.md` (guía completa con troubleshooting)

**Próximo?** Monitorea logs en `logs/distribucion_*.log` los primeros días.
