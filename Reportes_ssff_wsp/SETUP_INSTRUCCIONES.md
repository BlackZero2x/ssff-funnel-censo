# Setup Reportes SSFF WhatsApp

## 1. Instalación de dependencias

```bash
cd C:\proyectos\SSFF\Reportes_ssff_wsp
npm install
```

## 2. Configuración inicial

### Paso 1: Obtener el ID del grupo WhatsApp

```bash
cd C:\proyectos\SSFF\Reportes_ssff_wsp
node wa_server.js
```

- En la terminal verás un código QR
- **Escanea con tu teléfono** (WhatsApp → Configuración → Dispositivos vinculados → Vincular un dispositivo)
- Espera a que se conecte (verás mensaje de confirmación en la terminal)

Una vez conectado, en **otra terminal**:

```bash
python wa_client.py --list-groups
```

Busca el grupo `"Canal SSFF 2026 Gestión Auren I"` y copia su ID completo.

### Paso 2: Completar config.json

1. Abre `config.json`
2. Reemplaza:
   - `"my_number"`: Tu número de WhatsApp en formato `51XXXXXXXXX@c.us`
   - `groups.Canal_SSFF_2026_Gestion.group_id`: El ID del grupo que obtuviste arriba

### Paso 3: Llenar el directorio de menciones

Abre `DIRECTORIO_MENCIONES.md` y completa:
- **Supervisores**: nombres y números de WhatsApp
- **Jefe de Proyecto**: nombre y número
- **Gerente Comercial**: nombre y número

Para obtener los números exactos de WhatsApp de cada persona:

```bash
python wa_client.py --list-contacts "[nombre]"
```

Ejemplo:
```bash
python wa_client.py --list-contacts "Juan"
```

---

## 3. Pruebas de conectividad

### ¿El servidor está corriendo?

```bash
python wa_client.py --health
```

Debe responder con estado `connected`.

### ¿El grupo existe?

```bash
python wa_client.py --list-groups
```

### ¿Puedes enviar un mensaje de prueba?

```bash
python wa_client.py --test
```

Enviará un mensaje a tu número de WhatsApp como prueba.

---

## 4. Integración con scripts de cuotas

Una vez que el servidor esté activo, los scripts SSFF pueden enviar notificaciones:

```python
from wa_client import enviar_mensaje, enviar_con_menciones

# Opción 1: Mensaje simple
enviar_mensaje(
    group_id="120363...",  # ID del grupo
    message="✅ Cuotas mayo 2026 generadas exitosamente"
)

# Opción 2: Mensaje con menciones de supervisores
enviar_con_menciones(
    group_id="120363...",
    message="📊 Cuotas generadas",
    menciones=["supervisores", "jefe_proyecto"]
)
```

---

## 5. Parada y mantenimiento

Para detener el servidor:
```bash
# En la terminal del servidor Node.js: Ctrl+C
```

Para reiniciar sesión (nuevo QR):
```bash
# Borra session_data/
rm -r session_data
node wa_server.js
# Escanea el nuevo QR
```

---

## Troubleshooting

### "Cannot find module 'whatsapp-web.js'"
→ Ejecuta `npm install` nuevamente

### "QR código no aparece"
→ Verifica que tienes Node.js 14+ instalado: `node --version`

### "Conexión rechazada a localhost:3000"
→ El servidor wa_server.js no está corriendo. Abre una terminal y ejecuta:
```bash
cd C:\proyectos\SSFF\Reportes_ssff_wsp
node wa_server.js
```

### "Mensaje no se envía"
→ Verifica:
1. El grupo_id en `config.json` es correcto
2. El servidor está conectado: `python wa_client.py --health`
3. Tu número de WhatsApp está activo en el grupo

---

## Archivos generados

```
Reportes_ssff_wsp/
├── wa_server.js              # Servidor Node.js (ejecutar primero)
├── wa_client.py              # Cliente Python (ejecutar después)
├── config.json               # Configuración (NUNCA subir a Git)
├── config.example.json       # Plantilla (sí subir a Git)
├── package.json              # Dependencias Node.js
├── DIRECTORIO_MENCIONES.md   # Directorio de personas (rellenar manualmente)
├── SETUP_INSTRUCCIONES.md    # Este archivo
├── session_data/             # Datos de sesión de WhatsApp (NO subir)
└── logs/                      # Logs del servidor (NO subir)
```

---

**Próximo paso:** Completa `DIRECTORIO_MENCIONES.md` y luego contacta con los supervisores para obtener sus números de WhatsApp.
