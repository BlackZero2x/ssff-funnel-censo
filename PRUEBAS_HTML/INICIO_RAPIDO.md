# ⚡ GUÍA DE INICIO RÁPIDO — Pruebas HTML

**Tiempo de setup:** 10 minutos  
**Sistemas requeridos:** Python 3.8+, pip

---

## 🚀 **PASO 1: Setup (Una sola vez)**

### Abrir terminal en C:\proyectos\SSFF\PRUEBAS_HTML

```bash
cd C:\proyectos\SSFF\PRUEBAS_HTML
```

### Instalar dependencias

```bash
pip install -r requirements.txt
playwright install chromium
```

✅ **Listo cuando termine sin errores**

---

## 🧪 **PASO 2: Generar tu primer corte HTML**

```bash
python generar_corte_html.py --hora 10
```

**Esperado:**
```
==...==
  GENERAR CORTE HTML — Lunes 22/06/2026 — 10AM
==...==

[1] Cargando datos...
   [D] 2026-06-22 — 673 pedidos
   [D-7] 2026-06-15 — 708 pedidos
   [D-14] 2026-06-08 — 702 pedidos

[2] Procesando datos...

[3] Calculando indicadores...

[4] Generando HTML...
   Guardado: output\cortes_html\corte_20260622_10am.html

✅ Abre en navegador:
   start output\cortes_html\corte_20260622_10am.html
```

---

## 📸 **PASO 3: Capturar como PNG**

```bash
python capturar_cortes_html.py --hora 10
```

**Resultado:**
```
output\cortes_png\CORTE_HTML_10AM_20260622_123456.png
```

Abre la imagen y compara con Excel original.

---

## 🔄 **PASO 4: Versión Completa (Gen + Captura)**

Una sola línea para todo:

```bash
python ejecutar_corte_html.py --hora 10
```

Genera HTML → Captura PNG → Todo automático

---

## 📊 **PASO 5: Comparar Excel vs HTML**

### Opción A: Lado a lado

```bash
# Terminal 1: Ver producción (Excel)
dir C:\proyectos\SSFF\Reportes_ssff_wsp\cortes_imagenes\

# Terminal 2: Ver pruebas (HTML)
dir output\cortes_png\

# Abre ambas carpetas y compara visualmente
```

### Opción B: Abrir ambos navegadores

```bash
# Excel como PNG
start C:\proyectos\SSFF\Reportes_ssff_wsp\cortes_imagenes\GENERAL_10AM_*.png

# HTML directo
start output\cortes_html\corte_20260622_10am.html
```

---

## 🎨 **PASO 6: Customizar (Opcional)**

### Cambiar colores

Editar `generar_corte_html.py`, línea ~35:

```python
COLORES = {
    'tit_general': '#5B9BD5',    # ← Cambiar aquí
    'tit_zonal': '#1F3864',
    'tit_supervisor': '#006C50',
    # ...
}
```

### Cambiar tipografía

Buscar en `generar_corte_html.py` la sección `<style>`:

```css
body {
    font-family: 'Segoe UI', Arial;  # ← Cambiar aquí
}
```

---

## 🧐 **TROUBLESHOOTING**

### Error: "No module named 'playwright'"

```bash
pip install playwright
playwright install chromium
```

### Error: "No such file or directory: Reportes_ssff_wsp"

Asegúrate de estar en el path correcto:

```bash
cd C:\proyectos\SSFF\PRUEBAS_HTML
python generar_corte_html.py --hora 10
```

### PNG sale en blanco

Aumenta timeout en `capturar_cortes_html.py`:

```python
await asyncio.sleep(2)  # ← Cambiar de 1 a 2 o 3
```

### Fonts se ven diferentes

Instalar:

```bash
# Windows (PowerShell como Admin)
choco install noto-sans-fonts

# O descargar manualmente: https://fonts.google.com/noto
```

---

## 📈 **PREGUNTAS COMUNES**

**P: ¿Afecta a producción?**  
R: NO. Todo está en PRUEBAS_HTML separado.

**P: ¿Puedo dejar de usar Excel?**  
R: NO, todavía. Esto es experimental. Excel sigue en producción.

**P: ¿Qué hago con los PNGs generados?**  
R: Puedes dejarlos, visualizarlos, compararlos. Son solo para pruebas.

**P: ¿Por qué es más lento la primera vez?**  
R: Playwright descarga Chromium (~150MB). Solo la primera vez.

**P: ¿Puedo usar esto con MOVISTAR?**  
R: SÍ. Copiar el código y adaptar los datos.

---

## ✅ **CHECKLIST DE VALIDACIÓN**

- [ ] Instalaste requirements.txt
- [ ] Instalaste playwright + chromium
- [ ] Generaste HTML sin errores
- [ ] Abriste HTML en navegador (¿se ve bien?)
- [ ] Capturaste PNG sin errores
- [ ] PNG está en output/cortes_png/
- [ ] Comparaste Excel PNG vs HTML PNG
- [ ] Leíste los logs sin problemas

---

## 📞 **SIGUIENTE PASO**

Cuando el subproyecto esté validado:

1. Reporta hallazgos (qué se ve mejor, qué está roto, etc.)
2. Si todo está OK → listo para migración a producción
3. Si hay issues → iterar aquí sin riesgo

**¡Listo para empezar! 🚀**
