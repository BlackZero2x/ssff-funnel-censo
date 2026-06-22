# ✅ ANÁLISIS DE REQUIREMENTS — Estado Actual

**Fecha:** 22 de Junio de 2026  
**Entorno:** uv con venv compartido en C:\proyectos\.venv  
**Python:** 3.13.13

---

## 📊 **ESTADO DE LIBRERÍAS REQUERIDAS**

| Librería | Versión Instalada | Requerida | Estado |
|----------|------------------|-----------|--------|
| **pandas** | 3.0.3 | ≥2.0.0 | ✅ OK |
| **numpy** | 2.4.6 | ≥1.20.0 | ✅ OK |
| **pyodbc** | 5.3.0 | ≥5.0.0 | ✅ OK |
| **playwright** | 1.60.0 | ≥1.40.0 | ✅ OK |
| **pillow** | 12.2.0 | ≥10.0.0 | ✅ OK |
| **jinja2** | ❌ NO INSTALADO | ≥3.1.0 | ⚠️ FALTA |
| **psutil** | ❌ NO INSTALADO | ≥5.9.0 | ⚠️ FALTA |

---

## 🎯 **RESUMEN**

✅ **5 de 7 librerías instaladas y actualizadas**

❌ **2 librerías faltantes:**
- `jinja2` (para templates HTML)
- `psutil` (para monitoreo de memoria)

---

## 🚀 **SOLUCIÓN (2 OPCIONES)**

### **Opción A: Instalar lo que falta (RECOMENDADO)**

```bash
uv pip install jinja2 psutil
```

**Tiempo:** ~10 segundos  
**Riesgo:** CERO (no afecta otros proyectos, están todas en .venv compartida)

### **Opción B: Ejecutar sin psutil (RÁPIDO)**

Si **no quieres instalar psutil** (monitoreo de memoria es opcional):

1. Edita `PRUEBAS_HTML/generar_corte_html.py` línea 300
2. Comenta: `# mem_antes, mem_despues = limpiar_memoria()`
3. Ejecuta sin psutil

---

## 📝 **RECOMENDACIÓN**

**Usa Opción A** (instalar ambas):
- `jinja2` es CRÍTICO para generar HTML
- `psutil` es auxiliar (logs, pero no essential)
- Toma 10 segundos
- Sin riesgo

```bash
uv pip install jinja2 psutil
```

---

## 📋 **REQUIREMENTS.txt ACTUALIZADO**

```
pandas==3.0.3          (ya instalado)
numpy==2.4.6           (ya instalado)
pyodbc==5.3.0          (ya instalado)
playwright==1.60.0     (ya instalado)
pillow==12.2.0         (ya instalado)
jinja2==3.1.2          (⚠️ instalar)
psutil==5.9.5          (⚠️ instalar)
```

---

## ✅ **DESPUÉS DE INSTALAR**

Verifica:

```bash
uv pip list | grep -E "jinja2|psutil"
```

Deberías ver:
```
jinja2                    3.1.2
psutil                    5.9.5
```

---

## 🎬 **LUEGO PUEDES EJECUTAR**

```bash
cd C:\proyectos\SSFF\PRUEBAS_HTML
python generar_corte_html.py --hora 10
python capturar_cortes_html.py --hora 10
python ejecutar_corte_html.py --hora 10
```

✅ **TODO LISTO**
