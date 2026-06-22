# 📁 ESTRUCTURA DEL SUBPROYECTO PRUEBAS_HTML

**Aislado completamente de producción**

---

## 🗂️ **ÁRBOL DE DIRECTORIOS**

```
PRUEBAS_HTML/
│
├── README.md                        (documentación completa)
├── INICIO_RAPIDO.md                 (guía rápida)
├── ESTRUCTURA_PROYECTO.md           (este archivo)
├── requirements.txt                 (dependencias)
├── .gitignore                       (no incluir en git)
│
├── generar_corte_html.py           (🔴 SCRIPT PRINCIPAL #1)
│   └─ Genera HTML con datos SQL
│   └─ Reutiliza código de producción
│   └─ Uso: python generar_corte_html.py --hora 10
│
├── capturar_cortes_html.py         (🔴 SCRIPT PRINCIPAL #2)
│   └─ Captura HTML → PNG
│   └─ Usa Playwright (headless browser)
│   └─ Uso: python capturar_cortes_html.py --hora 10
│
├── ejecutar_corte_html.py          (🔴 SCRIPT PRINCIPAL #3)
│   └─ Orquestador: gen + captura
│   └─ Ejecuta ambos scripts
│   └─ Uso: python ejecutar_corte_html.py --hora 10
│
├── templates/                       (para futuro)
│   ├── corte_general.html          (templates Jinja2)
│   ├── corte_zonal.html
│   └── corte_supervisor.html
│
├── static/                          (para futuro)
│   └── estilos.css                 (CSS compartido)
│
└── output/                          (📂 GENERADO AUTOMÁTICAMENTE)
    ├── cortes_html/
    │   └── corte_20260622_10am.html
    │   └── corte_20260622_2pm.html
    │   └── ...
    │
    └── cortes_png/
        └── CORTE_HTML_10AM_20260622_123456.png
        └── CORTE_HTML_2PM_20260622_123457.png
        └── ...

logs/                               (📂 GENERADO AUTOMÁTICAMENTE)
├── corte_html_20260622.log        (generador)
└── capturador_20260622.log        (capturador)
```

---

## 🔄 **FLUJO DE DATOS**

```
SQL Server (eAuren)
    ↓
generar_corte_html.py
    ├─ Conecta a SQL
    ├─ Obtiene datos (MISMO que producción)
    ├─ Procesa (MISMO que producción)
    ├─ Renderiza HTML (NUEVO)
    └─ Guarda: output/cortes_html/corte_*.html
        ↓
capturar_cortes_html.py
    ├─ Lee HTML local
    ├─ Abre en Playwright (headless Chromium)
    ├─ Espera renderizado
    ├─ Captura screenshot
    └─ Guarda: output/cortes_png/*.png
        ↓
    Tu visualización
    ├─ Abrir PNG en navegador
    ├─ Comparar con Excel
    └─ Validar mejoras
```

---

## 📝 **SCRIPTS EXPLICADOS**

### **generar_corte_html.py**

```python
Propósito:  Genera HTML con tablas GENERAL, ZONAL, SUPERVISOR

Entrada:    --hora 8..18 (obligatorio)
            --fecha YYYY-MM-DD (opcional)

Proceso:
  1. Conectar SQL (pyodbc)
  2. Cargar datos (preventa_dia)
  3. Procesar datos (IDÉNTICO a producción)
  4. Calcular métricas (IDÉNTICO a producción)
  5. Renderizar templates (NUEVO)
  6. Guardar HTML

Salida:     output/cortes_html/corte_YYYYMMDD_HHXM.html
            Tamaño: ~50KB
            Tiempo: ~3.5 segundos
```

### **capturar_cortes_html.py**

```python
Propósito:  Captura screenshot del HTML usando Playwright

Entrada:    --hora 8..18 (obligatorio)
            --fecha YYYY-MM-DD (opcional)

Proceso:
  1. Buscar archivo HTML generado
  2. Lanzar navegador Chromium (headless)
  3. Cargar HTML (file://)
  4. Esperar renderizado (networkidle)
  5. Capturar screenshot
  6. Guardar PNG

Salida:     output/cortes_png/CORTE_HTML_*.png
            Tamaño: ~200-400KB (depende complejidad)
            Tiempo: ~1.6 segundos
```

### **ejecutar_corte_html.py**

```python
Propósito:  Orquestador simple

Entrada:    --hora 8..18 (obligatorio)
            --fecha YYYY-MM-DD (opcional)

Proceso:
  1. Ejecuta: generar_corte_html.py
  2. Ejecuta: capturar_cortes_html.py
  3. Registra logs

Salida:     HTML + PNG + logs
            Tiempo total: ~5.1 segundos
```

---

## 🔐 **AISLAMIENTO DE PRODUCCIÓN**

```
PRODUCCIÓN (C:\proyectos\SSFF\Reportes_ssff_wsp\)
├── generar_corte_ventas.py      (Excel)
├── capturar_cortes.py            (xlwings + screenshot)
├── ejecutar_corte_orquestado.py  (orquestador Excel)
└── CORTE_VENTAS_*.xlsx          (archivos Excel)
    ✅ NO MODIFICADO

PRUEBAS (C:\proyectos\SSFF\PRUEBAS_HTML\)
├── generar_corte_html.py         (HTML nuevo)
├── capturar_cortes_html.py       (Playwright nuevo)
├── ejecutar_corte_html.py        (orquestador nuevo)
└── output/                        (NO comparte paths con prod)
    ✅ COMPLETAMENTE AISLADO

🔒 NO HAY CRUCE DE ARCHIVOS
🔒 NO HAY COMPETENCIA POR RECURSOS
🔒 NO AFECTA TASK SCHEDULER
🔒 NO AFECTA A GERENCIA
```

---

## 📊 **REUTILIZACIÓN DE CÓDIGO**

```
Código compartido (IMPORTADO de producción):
  ✓ conectar_sql()
  ✓ cargar_preventa_dia()
  ✓ cargar_dia_con_cache()
  ✓ cargar_mapa_zonas()
  ✓ calcular_cuota_supervisor()
  ✓ calcular_cuota_zonal()
  ✓ enriquecer_con_vendedor()
  ✓ agregar_hora_lbl()
  ✓ filtrar_corte()
  ✓ indicadores()
  ✓ agregar_general()
  ✓ agregar_por_columna()
  ✓ agregar_por_sup_vendedor()
  ✓ calcular_cuota_dia_vendedor()

Código NUEVO:
  ✗ Renderización HTML (templates)
  ✗ Captura Playwright
  ✗ Estilos CSS

Ventaja: Si hay bug en datos → TODOS se benefician del fix
```

---

## 🚀 **ROADMAP DE PRUEBAS**

```
Semana 1: Setup y exploración
  □ Instalar dependencias
  □ Generar HTML cortes
  □ Capturar PNGs
  □ Comparar visualmente con Excel

Semana 2: Validación
  □ Probar todas las horas (8AM-6PM)
  □ Probar múltiples días
  □ Comparar métricas (tiempo, tamaño, calidad)
  □ Documentar hallazgos

Semana 3: Preparación para prod
  □ Crear script de migración
  □ Testing en staging
  □ Preparar rollback plan

Cuando esté listo → Migración a producción
  □ Backup Excel (keep for 1 month)
  □ Switch: generar_corte_html en producción
  □ Monitor logs (7 días)
  □ Feedback a gerencia
```

---

## 📞 **CÓMO REPORTAR PROBLEMAS**

Si algo no funciona:

1. **Nota el error exacto:**
   ```
   "Error en capturar_cortes_html.py: [detalle]"
   ```

2. **Revisa logs:**
   ```
   tail logs/corte_html_20260622.log
   ```

3. **Reproduce el problema:**
   ```
   python generar_corte_html.py --hora 10
   python capturar_cortes_html.py --hora 10
   ```

4. **Reporta con contexto:**
   - ¿Qué comando ejecutaste?
   - ¿Cuál es el error exacto?
   - ¿Qué hora/fecha?
   - ¿Ya funcionaba antes?

---

**Subproyecto creado:** 22/06/2026  
**Estado:** Listo para exploración manual  
**Próximo paso:** ¡Ejecutar los scripts! 🚀
