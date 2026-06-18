# ✅ Checklist de Setup — Sistema Antibang SSFF

Sigue estos pasos para activar el sistema de cortes automáticos.

---

## 📋 Checklist de configuración

### Fase 1: Preparación (5 min)

- [ ] **Leer documentación**
  ```bash
  # Abre estos archivos en orden:
  # 1. README.md (overview)
  # 2. ANTIBANG_GUIA.md (técnico)
  # 3. IMPLEMENTACION_ANTIBANG.md (resumen)
  ```

- [ ] **Iniciar servidor OPENWA** (Terminal 1)
  ```bash
  cd C:\proyectos\SSFF\Reportes_ssff_wsp
  node wa_server.js
  # Escanea QR si es primera vez
  # Verifica que sea el número 51975155264
  ```

- [ ] **Obtener IDs de grupos**
  ```bash
  # En otra terminal, ejecuta:
  python wa_client.py --list-groups
  # Copia los IDs para:
  #   - Canal SSFF 2026 Gestión Auren I
  #   - Grupos de supervisores (F8, M0, K0, P0, V0)
  ```

### Fase 2: Configuración (10 min)

- [ ] **Editar config.json**
  ```bash
  # Abre: C:\proyectos\SSFF\Reportes_ssff_wsp\config.json
  # Llena estos campos:
  ```

  ```json
  {
    "groups": {
      "Canal_SSFF_2026_Gestion": "120363XXXXXXXXX@g.us"  // ID del canal
    },
    "groups_supervisores": {
      "F8": "120363XXXXXXXXX@g.us",   // Grupo F8 (opcional)
      "M0": "120363XXXXXXXXX@g.us",   // Grupo M0 (opcional)
      "K0": "120363XXXXXXXXX@g.us",   // Grupo K0 (opcional)
      "P0": "120363XXXXXXXXX@g.us",   // Grupo P0 (opcional)
      "V0": "120363XXXXXXXXX@g.us"    // Grupo V0 (opcional)
    },
    "mentions": {
      "jefe_proyecto": {
        "nombre": "Nombre Jefe",
        "numero": "51912345678@c.us"   // Opcional
      },
      "gerente_comercial": {
        "nombre": "Nombre Gerente",
        "numero": "51987654321@c.us"   // Opcional
      }
    }
  }
  ```

- [ ] **Guardar config.json**
  ```bash
  # Verifica que no tenga errores JSON:
  python -m json.tool config.json
  # Si no muestra errores, está correcto
  ```

### Fase 3: Validación (5 min)

- [ ] **Ejecutar test**
  ```bash
  cd C:\proyectos\SSFF\Reportes_ssff_wsp
  python test_antibang.py
  
  # Debería mostrar:
  # ✅ Servidor activo
  # ✅ Configuración encontrada
  # ✅ Enviado exitosamente
  # ✅ Cuenta operativa
  ```

- [ ] **Probar generación de corte**
  ```bash
  python generar_corte_ventas.py --hora 10
  
  # Debería:
  # 1. Cargar datos de BD
  # 2. Generar Excel
  # 3. Enviar automáticamente a WhatsApp
  # 4. Guardar en: CORTE_VENTAS_*.xlsx
  ```

- [ ] **Verificar en WhatsApp**
  ```
  • Abre el grupo/canal "Canal SSFF 2026 Gestión Auren I"
  • Busca el mensaje de prueba
  • Verifica que tenga el Excel adjunto
  ```

### Fase 4: Programación automática (5 min)

**IMPORTANTE:** Requiere permisos de Administrador

- [ ] **Ejecutar programador**
  ```bash
  # COMO ADMINISTRADOR:
  # 1. Abre PowerShell/CMD como Admin
  # 2. Ejecuta:
  cd C:\proyectos\SSFF\Reportes_ssff_wsp
  python programar_cortes.py
  
  # Elige: (s/n) para eliminar tareas anteriores si existen
  
  # Debería crear 6 tareas:
  #   ✅ SSFF_Corte_8AM
  #   ✅ SSFF_Corte_10AM
  #   ✅ SSFF_Corte_12PM
  #   ✅ SSFF_Corte_2PM
  #   ✅ SSFF_Corte_4PM
  #   ✅ SSFF_Corte_6PM
  ```

- [ ] **Verificar tareas**
  ```bash
  # Abre Task Scheduler:
  # Control Panel → Administrative Tools → Task Scheduler
  # Busca: SSFF_Corte_*
  # Debería haber 6 tareas programadas
  ```

- [ ] **Testear ejecución automática** (opcional)
  ```bash
  # En Task Scheduler, haz clic derecho en una tarea
  # → Run
  # Verifica que se ejecute sin errores
  ```

---

## 🆘 Troubleshooting

### ❌ "Servidor no disponible"
```
→ Verifica que node wa_server.js está corriendo
→ cd C:\proyectos\SSFF\Reportes_ssff_wsp
→ node wa_server.js
→ Escanea el QR si es primera vez
```

### ❌ "config.json no encontrado"
```
→ Verifica que estás en: C:\proyectos\SSFF\Reportes_ssff_wsp
→ El archivo debe existir (ya viene con la carpeta)
```

### ❌ "JSON error"
```
→ Verificar formato JSON:
→ python -m json.tool config.json
→ Común: falta coma, comillas mal cerradas
```

### ❌ "No hay datos de ventas para hoy"
```
→ El SP tarda ~5s en procesar
→ El script reintenta 5 veces (15 segundos máximo)
→ Si persiste: contactar Data Engineer
```

### ❌ "Rate limit detectado (429)"
```
→ Meta detectó envío masivo
→ Sistema pausa 1 hora automáticamente
→ Reintentar después
```

### ⚠️ "Permisos de administrador requeridos"
```
→ El programador necesita ser Admin
→ Abre PowerShell/CMD como Admin:
→ Tecla Windows → "PowerShell" → right-click → "Run as Administrator"
```

---

## 📊 Estructura final

Después de configurar, tu carpeta debería verse así:

```
C:\proyectos\SSFF\Reportes_ssff_wsp\
├── generar_corte_ventas.py          ← Script principal
├── wa_sender_antibang.py            ← Core antibang
├── enviar_corte_a_todos.py          ← Distribuidor
├── test_antibang.py                 ← Validación
├── programar_cortes.py              ← Programador
├── config.json                      ← TU CONFIGURACIÓN (no subir a git)
├── config.example.json              ← Plantilla
├── README.md                        ← Overview
├── ANTIBANG_GUIA.md                 ← Documentación técnica
├── IMPLEMENTACION_ANTIBANG.md       ← Resumen
├── CHECKLIST_SETUP.md               ← Este archivo
├── logs/                            ← Se crea automáticamente
│   ├── envios_YYYYMMDD.log
│   └── distribucion_YYYYMMDD.log
└── cache/                           ← Se crea automáticamente (D-7, D-14)
    ├── cache_d7.pkl
    └── cache_d14.pkl
```

---

## 🚀 Una vez configurado

### Comando diario

```bash
# Para generar corte a una hora específica:
python generar_corte_ventas.py --hora 10

# O déjalo automático:
# Task Scheduler ejecutará las 6 tareas (8AM, 10AM, 12PM, 2PM, 4PM, 6PM)
```

### Monitoreo

```bash
# Ver logs en tiempo real:
tail -f logs/distribucion_*.log

# Ver últimos envíos:
cat logs/distribucion_*.log | tail -50
```

### Si necesitas detener

```bash
# Desactivar una tarea en Task Scheduler:
# right-click en tarea → Disable

# O ejecutar:
schtasks /delete /tn "SSFF_Corte_10AM" /f
```

---

## 📞 Soporte rápido

| Problema | Solución |
|----------|----------|
| Servidor no corre | `node wa_server.js` |
| Config con errores | `python -m json.tool config.json` |
| Prueba fallida | `python test_antibang.py` |
| Sin datos de BD | Esperar 15s (SP tarda ~5s) |
| Rate limit (429) | Esperar 1h (automático) |
| Tareas no se ejecutan | Verificar Task Scheduler, permisos Admin |

---

## ✅ Señales de que está funcionando

- ✅ `python test_antibang.py` muestra 4 checks verdes
- ✅ `python generar_corte_ventas.py --hora 10` genera Excel
- ✅ Excel aparece en WhatsApp en el grupo/canal
- ✅ Task Scheduler muestra 6 tareas SSFF_Corte_*
- ✅ Logs en `logs/` registran envíos

---

## 📅 Próximos pasos

### Corto plazo (hoy)
1. ✅ Llenar config.json
2. ✅ Ejecutar test_antibang.py
3. ✅ Probar generar_corte_ventas.py --hora 10

### Mediano plazo (esta semana)
4. ✅ Ejecutar programar_cortes.py
5. ✅ Verificar que las 6 tareas se ejecutan automáticamente
6. ✅ Monitorear logs los primeros días

### Largo plazo (después)
7. Ajustar horarios si es necesario
8. Agregar grupos/contactos nuevos
9. Implementar alertas (si venta cae >25%)

---

## 🎉 Una vez todo funcione

El sistema corre completamente automático:
- ⏰ 8 AM → Corte 8AM automático
- ⏰ 10 AM → Corte 10AM automático
- ⏰ 12 PM → Corte 12PM automático
- ⏰ 2 PM → Corte 2PM automático
- ⏰ 4 PM → Corte 4PM automático
- ⏰ 6 PM → Corte 6PM automático

**Sin intervención manual, sin riesgo de baneo.** 🛡️

---

**Tiempo total estimado:** 25 minutos  
**Dificultad:** Media (copiar IDs, llenar JSON)  
**Soporte:** Ver ANTIBANG_GUIA.md si hay dudas  

¿Preguntas? Revisa IMPLEMENTACION_ANTIBANG.md o README.md.
