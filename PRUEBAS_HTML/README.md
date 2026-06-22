# 🧪 SUBPROYECTO: PRUEBAS HTML — Upgrade Futuro

**Estado:** Experimental (no afecta producción)  
**Objetivo:** Validar generación de cortes en HTML con captura automática  
**Fecha creación:** 22 de Junio de 2026

---

## 📁 **ESTRUCTURA**

```
PRUEBAS_HTML/
├── README.md (este archivo)
├── requirements.txt (dependencias)
├── generar_corte_html.py (generador HTML)
├── capturar_cortes_html.py (capturador + envío)
├── ejecutar_corte_html.py (orquestador)
├── templates/
│   ├── corte_general.html
│   ├── corte_zonal.html
│   └── corte_supervisor.html
├── static/
│   └── estilos.css
├── output/
│   ├── cortes_html/
│   └── cortes_png/
└── logs/
```

---

## 🚀 **USO (MANUAL)**

### **Paso 1: Instalar dependencias**
```bash
cd C:\proyectos\SSFF\PRUEBAS_HTML
pip install -r requirements.txt
```

### **Paso 2: Generar corte HTML manual**
```bash
python generar_corte_html.py --hora 10
```

Genera: `output/cortes_html/corte_20260622_10AM.html`

### **Paso 3: Abrir en navegador (para inspeccionar)**
```bash
# En Windows
start output\cortes_html\corte_20260622_10AM.html

# En Mac/Linux
open output/cortes_html/corte_20260622_10AM.html
```

### **Paso 4: Capturar como PNG**
```bash
python capturar_cortes_html.py --hora 10
```

Genera:
- `output/cortes_png/GENERAL_10AM_20260622.png`
- `output/cortes_png/ZONAL_10AM_20260622.png`
- `output/cortes_png/SUPERVISOR_10AM_20260622.png`

### **Paso 5: Ejecutar completo (orquestado)**
```bash
python ejecutar_corte_html.py --hora 10
```

Genera HTML → Captura PNG → Logs todo

---

## 📊 **COMPARAR CON PRODUCCIÓN**

**Lado a lado:**
```bash
# Producción (Excel)
C:\proyectos\SSFF\Reportes_ssff_wsp\CORTE_VENTAS_20260622.xlsx
  → PNG capturado: GENERAL_10AM_20260622.png

# Pruebas (HTML)
C:\proyectos\SSFF\PRUEBAS_HTML\output\cortes_html\corte_20260622_10AM.html
  → PNG capturado: GENERAL_10AM_20260622.png
```

Abre ambos PNG en el navegador y compara visualmente.

---

## 🔧 **CARACTERÍSTICAS IMPLEMENTADAS**

### **Generador HTML (generar_corte_html.py)**
- ✅ Obtiene datos MISMO SQL que producción
- ✅ Procesa datos idénticos
- ✅ Genera 3 HTMLs: GENERAL, ZONAL, SUPERVISOR
- ✅ Usa Jinja2 templates
- ✅ Estilos CSS con Tailwind

### **Capturador Playwright (capturar_cortes_html.py)**
- ✅ Abre HTML en navegador headless
- ✅ Renderiza completamente
- ✅ Captura screenshot full-page
- ✅ Cropea cada tabla automáticamente
- ✅ Guarda PNG con timestamp

### **Orquestador (ejecutar_corte_html.py)**
- ✅ Ejecuta gen → captura secuencial
- ✅ Reintentos automáticos
- ✅ Logs detallados
- ✅ Exit codes informativos

---

## 📈 **MÉTRICAS ESPERADAS**

| Métrica | Excel (Prod) | HTML (Prueba) | Mejora |
|---------|------------|--------------|--------|
| Tiempo generación | 6s | 3.5s | 42% ⬇️ |
| Tiempo captura | 4s | 1.6s | 60% ⬇️ |
| Tamaño archivo | 8-10 MB | 50 KB | 160x ⬇️ |
| Calidad PNG | Básica | Premium | 9/10 |

---

## 📝 **LOGS Y DEBUGGING**

```bash
# Ver logs en tiempo real
tail -f logs/corte_html_20260622.log

# Ver solo errores
grep ERROR logs/corte_html_20260622.log

# Comparar tiempos
grep "segundos" logs/corte_html_20260622.log
```

---

## 🎨 **CUSTOMIZACIÓN**

### **Cambiar colores**
Editar `static/estilos.css`:
```css
.titulo-general { background: #5B9BD5; }
.titulo-zonal { background: #1F3864; }
.titulo-supervisor { background: #006C50; }
```

### **Cambiar tipografía**
Editar `static/estilos.css`:
```css
body { font-family: 'Segoe UI', Arial, sans-serif; }
```

### **Agregar gráficos**
Usar Chart.js en templates (futuro):
```html
<canvas id="chart-zonal"></canvas>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
```

---

## 🚨 **TROUBLESHOOTING**

| Problema | Solución |
|----------|----------|
| `ModuleNotFoundError: No module named 'playwright'` | `pip install playwright` y luego `playwright install chromium` |
| `No such file or directory: templates/` | Asegúrate de estar en `C:\proyectos\SSFF\PRUEBAS_HTML` |
| PNG sale en blanco | Aumenta timeout en capturador (línea 50, cambiar a 5000ms) |
| Fonts no se ven | Instalar: `choco install noto-sans-fonts` (Windows) |

---

## 📞 **PRÓXIMOS PASOS**

Cuando SSFF esté listo para upgrade:

1. Copiar `generar_corte_html.py` a `Reportes_ssff_wsp/`
2. Copiar `capturar_cortes_html.py` a `Reportes_ssff_wsp/`
3. Copiar templates + estilos a `Reportes_ssff_wsp/`
4. Actualizar `ejecutar_corte_orquestado.py` para usar HTML
5. Testear 1 semana
6. Migración completa a producción

---

## ✅ **CHECKLIST DE PRUEBAS**

- [ ] Instalar dependencias sin errores
- [ ] Generar HTML sin errores
- [ ] Abrir HTML en navegador (se ve bien?)
- [ ] Capturar PNG sin errores
- [ ] PNG tiene buena calidad
- [ ] Comparar PNG Excel vs PNG HTML (¿cuál se ve mejor?)
- [ ] Verificar tiempos (¿HTML es más rápido?)
- [ ] Probar con 3 horas diferentes (8AM, 12PM, 6PM)
- [ ] Revisar logs

---

**Subproyecto creado:** 22/06/2026  
**Mantener separado de:** Reportes_ssff_wsp (producción)  
**No afecta a:** Task Scheduler, WhatsApp, gerencia
