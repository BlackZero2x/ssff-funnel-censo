# 🚀 Integración con Servidor MOVISTAR

## Arquitectura

```
┌─────────────────────────────────────────────────────┐
│   generar_cuota.py (SSFF)                           │
│   ↓                                                  │
│   Genera cuotas Excel → enviar_reporte_ssff.py     │
└────────────────────┬────────────────────────────────┘
                     │
                     ↓
┌──────────────────────────────────────────────────────┐
│   wa_client.py (SSFF) → conecta a wa_server.js     │
│                        (Puerto 8002, MOVISTAR)      │
└──────────────────────┬───────────────────────────────┘
                       │
                       ↓
┌──────────────────────────────────────────────────────┐
│   WhatsApp — 51975155264 (AUREN)                    │
│   └─ Canal SSFF 2026 Gestión Auren I               │
│      └─ Notificación + Menciones a supervisores     │
└──────────────────────────────────────────────────────┘
```

---

## ✅ Requisitos previos

1. **Servidor MOVISTAR activo:**
   ```bash
   cd C:\proyectos\AVANCE_MOVISTAR\whatsapp_server
   node wa_server.js
   ```
   Debe estar conectado a WhatsApp (ya debería estar funcionando)

2. **Completar `DIRECTORIO_MENCIONES.md`** en SSFF:
   - Nombres de supervisores
   - Nombre Jefe de Proyecto
   - Nombre Gerente Comercial

---

## 📋 Paso 1: Configurar mentions en config.json

Edita `C:\proyectos\SSFF\Reportes_ssff_wsp\config.json`:

```json
{
  "mentions": {
    "supervisores": [
      "Juan Pérez",
      "Carlos López",
      "María García"
    ],
    "jefe_proyecto": {
      "nombre": "Roberto Martínez",
      "numero": "51912345678@c.us"
    },
    "gerente_comercial": {
      "nombre": "Patricia Rodríguez",
      "numero": "51987654321@c.us"
    }
  }
}
```

---

## 🧪 Paso 2: Probar la conexión

```bash
cd C:\proyectos\SSFF\Reportes_ssff_wsp
python enviar_reporte_ssff.py --test
```

Debería mostrar:
```
✅ Verificando configuración...
   Número WhatsApp: 51975155264@c.us
   Servidor: localhost:8002
   Grupo SSFF: 120363278118591818@g.us
✅ Servidor activo
```

Si dice "❌ Servidor no disponible":
- Verifica que `node wa_server.js` está corriendo en MOVISTAR
- Que el puerto 8002 esté libre

---

## 📤 Paso 3: Integración en generar_cuota.py

En tu script `generar_cuota.py`, al final después de generar el archivo Excel:

### Opción A: Envío simple (sin menciones)

```python
import subprocess

# Después de crear cuotas_ssff_mayo2026.xlsx
archivo_salida = "cuotas_ssff_mayo2026.xlsx"
mes = "Mayo 2026"

resultado = subprocess.run([
    "python",
    "Reportes_ssff_wsp/enviar_reporte_ssff.py",
    "--archivo", archivo_salida,
    "--mes", mes
], capture_output=True, text=True)

if resultado.returncode == 0:
    print("✅ Reporte enviado a WhatsApp")
else:
    print(f"❌ Error al enviar: {resultado.stderr}")
```

### Opción B: Envío con menciones

```python
import subprocess

archivo_salida = "cuotas_ssff_mayo2026.xlsx"
mes = "Mayo 2026"

resultado = subprocess.run([
    "python",
    "Reportes_ssff_wsp/enviar_reporte_ssff.py",
    "--archivo", archivo_salida,
    "--mes", mes,
    "--menciones", "supervisores,jefe_proyecto,gerente_comercial"
], capture_output=True, text=True)

if resultado.returncode == 0:
    print("✅ Reporte enviado con menciones")
else:
    print(f"❌ Error: {resultado.stderr}")
```

### Opción C: Envío condicional (solo si no hay errores)

```python
import subprocess
from pathlib import Path

def enviar_reporte_si_exito(archivo, mes, menciones=None):
    """Envía reporte solo si el archivo se generó correctamente"""
    
    # Validar que el archivo existe
    if not Path(archivo).exists():
        print(f"❌ Archivo {archivo} no encontrado")
        return False
    
    # Validar tamaño mínimo (> 10 KB)
    tamaño = Path(archivo).stat().st_size
    if tamaño < 10000:
        print(f"❌ Archivo muy pequeño ({tamaño} bytes). Posible error en generación")
        return False
    
    # Enviar
    cmd = [
        "python",
        "Reportes_ssff_wsp/enviar_reporte_ssff.py",
        "--archivo", archivo,
        "--mes", mes
    ]
    
    if menciones:
        cmd.extend(["--menciones", menciones])
    
    resultado = subprocess.run(cmd, capture_output=True, text=True)
    
    if resultado.returncode == 0:
        print("✅ Reporte enviado exitosamente")
        return True
    else:
        print(f"❌ Error al enviar: {resultado.stderr}")
        return False

# Uso
if __name__ == "__main__":
    # ... tu código de generación de cuotas ...
    
    # Al final:
    enviar_reporte_si_exito(
        "cuotas_ssff_mayo2026.xlsx",
        "Mayo 2026",
        "supervisores,jefe_proyecto"
    )
```

---

## 📊 Registros y auditoría

Todos los envíos se registran en:
- **Logs en tiempo real:** `Reportes_ssff_wsp/logs/envios_YYYYMMDD.log`
- **Registro JSON:** `Reportes_ssff_wsp/logs/envios_registro.jsonl`

Ejemplo de log:
```
[2026-06-17T14:30:15.123456] [INFO] 🚀 Iniciando envío de reporte SSFF — Mayo 2026
[2026-06-17T14:30:15.234567] [INFO] ✅ Horario seguro para envío
[2026-06-17T14:30:16.345678] [INFO] ✅ Servidor activo: Servidor activo
[2026-06-17T14:30:17.456789] [INFO] ✅ Archivo cuotas_ssff_mayo2026.xlsx (2.5 MB)
[2026-06-17T14:30:17.567890] [INFO] 📝 Mensaje a enviar...
[2026-06-17T14:30:20.678901] [INFO] ✅ Mensaje enviado
```

Para ver registro JSON:
```bash
type Reportes_ssff_wsp\logs\envios_registro.jsonl
```

---

## ⚠️ Horarios y límites de seguridad

El sistema respeta automáticamente:

| Aspecto | Configuración | Por qué |
|---------|---------------|--------|
| **Horario** | Lunes-viernes, 08:00-18:00 | Evita patrones robóticos |
| **Intervalo** | 3 segundos mín entre envíos | Parece humano |
| **Volumen** | Máx 30 mensajes/día | Bajo perfil |

Si intentas enviar fuera de horario:
```
⏰ No es día laboral — Envío pospuesto
```

---

## 🔧 Troubleshooting

### "Servidor no disponible"
```
→ Verifica que node wa_server.js está corriendo
→ cd C:\proyectos\AVANCE_MOVISTAR\whatsapp_server
→ node wa_server.js
```

### "Fuera de horario"
```
→ El envío está programado fuera de lunes-viernes 08:00-18:00
→ Edita ssff_config en config.json si necesitas otros horarios
```

### "wa_client.py no encontrado"
```
→ Verifica que estás en la carpeta correcta:
→ cd C:\proyectos\SSFF\Reportes_ssff_wsp
→ python enviar_reporte_ssff.py --test
```

---

## 📝 Ejemplo completo: generar_cuota.py con envío

```python
import pandas as pd
import subprocess
from datetime import datetime
from pathlib import Path

def main():
    print("📊 Iniciando generación de cuotas...")
    
    # ... tu lógica de cuotas ...
    
    # Generar Excel
    mes_nombre = "Mayo 2026"
    archivo_salida = "cuotas_ssff_mayo2026.xlsx"
    
    print(f"✍️  Generando {archivo_salida}...")
    # ... código de creación de Excel ...
    
    # Enviar notificación
    print(f"📤 Enviando notificación a WhatsApp...")
    resultado = subprocess.run([
        "python",
        "Reportes_ssff_wsp/enviar_reporte_ssff.py",
        "--archivo", archivo_salida,
        "--mes", mes_nombre,
        "--menciones", "supervisores,jefe_proyecto"
    ], capture_output=True, text=True)
    
    if resultado.returncode == 0:
        print("✅ Proceso completado exitosamente")
        print(f"   • Archivo: {archivo_salida}")
        print(f"   • Notificación enviada a Canal SSFF")
    else:
        print(f"⚠️  Error en envío: {resultado.stderr}")
        # El archivo se generó, pero no se envió la notificación

if __name__ == "__main__":
    main()
```

---

## 🎯 Resumen

- ✅ Todo usa el **mismo servidor MOVISTAR** (puerto 8002)
- ✅ **Seguro:** Respeta horarios, intervalos y límites
- ✅ **Monitoreado:** Todos los envíos se registran
- ✅ **Flexible:** Envía con o sin menciones
- ✅ **Integrable:** Una línea en `generar_cuota.py`

**Estado:** Listo para integración en generar_cuota.py
