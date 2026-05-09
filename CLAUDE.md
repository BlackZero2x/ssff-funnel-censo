# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

## Objetivo del proyecto
Automatizar la generación mensual de cuotas de ventas para el proyecto San Fernando (SSFF), distribuyendo objetivos por categoría, fuerza de ventas (FFVV) y ruta (vendedor).

---

## Alcance
- **Qué hace:** extrae, compara y analiza el histórico de ventas para generar un archivo Excel con cuotas realistas por categoría y por ruta (vendedor).
- **Qué NO hace:** no genera reportería de SKU, KPIs individuales ni dashboards detallados.
- **Público objetivo:** Gerencia, Jefes y Supervisores.

---

## Comandos de ejecución

```bash
# Paso 1 — Generar cuotas de presentación (main script)
python generar_cuota_mayo.py

# Paso 2 — Convertir a formato BD relacional
python generar_bbdd.py

# Paso 3 — Histórico detallado de PROCESADOS (opcional, independiente)
python generar_historico_procesados.py
```

Los tres scripts son independientes entre sí, salvo que `generar_bbdd.py` lee el output de `generar_cuota_mayo.py` (`cuotas_ssff_may2026.xlsx`).

---

## Stack técnico
- **Lenguaje:** Python + SQL (conexión a SQL Server del proyecto SSFF)
- **Salida:** archivos Excel (.xlsx) generados con `openpyxl`
- **Librerías clave:** `pandas`, `numpy`, `openpyxl`
- **Skills de mensajería:** usar las skills del proyecto `C:\proyectos\AVANCE_MOVISTAR\` para notificaciones

---

## Arquitectura de scripts

### `generar_cuota_mayo.py` — Script principal (1.500 líneas aprox.)

Flujo secuencial:

```
[1] Carga de datos
    ├─ export_data_ssff.csv          → histórico global (1.149 registros)
    ├─ export_data_ssff_rutas.csv    → detalle por ruta/cliente (119 MB, ~1.29M registros)
    ├─ TABLAS_RUTAS.xlsx             → maestro oficial de rutas (hoja RUTA_ACTUAL)
    ├─ cartera_simplificada.xlsx     → 24.715 clientes (columna es_censo para los ~5.930 del censo)
    ├─ cuotas_ssff_abr2026.xlsx      → cuotas de abril (baseline de comparación)
    └─ cuota_oficial_ssff.xlsx       → cuotas fijadas por gerencia para cat. SSFF

[2] Proyección lineal de abril
    └─ (ventas_reales / días_transcurridos) × días_hábiles_totales

[3] Cálculo Tipo 1 global por categoría
    └─ tipo1(s): (mean(s) + max(s)) / 2  — sobre periodos seleccionados

[4] Tipo 1 por ruta
    └─ tipo1_ruta_pivot(df, col): pivot de rutas × categorías con pesos normalizados

[5] Restricciones KB y redistribución
    ├─ Rutas KB → cuota = 0 en categorías SSFF
    └─ Su porción se redistribuye a F8, M0, P0, V0

[6] Distribución y piso mínimo
    ├─ Aplicar pesos a cuota oficial por categoría
    ├─ Piso: toda ruta debe tener ≥ S/ 100.000 en total
    └─ Rutas bajo el piso: escalar categorías proporcionalmente

[7] Volumen y cobertura PROCESADOS (solo F8 + M0)
    ├─ EMBUTIDOS:  LINEA = 'EMBUTIDOS'
    └─ CONGELADOS: LINEA ∈ (ELABORADOS, SEMIELABORADOS, PRECOCIDOS)

[8] Generación Excel — 6 hojas:
    PROCESO | CUOTAS_SOLES | CUOTA_COBERTURA | CUOTA_VOL_SSFF | HISTORICO | VERSUS
```

**Funciones helper clave:**
- `tipo1(s)` — fórmula Tipo 1
- `cuota_cat_global(df, col)` — agrega cuota global por categoría
- `tipo1_ruta_pivot(df, col)` — Tipo 1 por ruta en formato pivot
- `calcular_pesos(df_t1)` — normaliza pesos como porcentaje
- `escribir_hoja_cuota()` — genera hojas de cuota con formato estándar
- `fill()`, `fnt()`, `aln()`, `brd()` — helpers de estilo para openpyxl

### `generar_bbdd.py` — Conversión a formato relacional

Lee `cuotas_ssff_may2026.xlsx` (hojas: CUOTAS_SOLES, CUOTA_VOL_SSFF, CUOTA_COBERTURA) y produce `CUOTAS_BBDD_202605.xlsx` con estructura:

```
periodo | ruta | tipo | linea | producto | soles | kilos | cobertura | temporal
```

- Filas GENERAL (79): resumen por ruta, todas las categorías sumadas
- Filas LINEA (1.580): detalle por ruta × categoría
- Validación interna: `SUM(GENERAL) == SUM(LINEA)` por ruta

### `generar_historico_procesados.py` — Análisis histórico de PROCESADOS

Genera `historico_procesados_rutas_ok.xlsx` con:
- Detalle por ruta de EMBUTIDOS y CONGELADOS (últimos 3 meses + proyección abril)
- Subtotales por FFVV
- Tabla color-coded (abril = amarillo = proyectado)

---

## Estructura de datos de entrada

El archivo `export_data_ssff.csv` (y su equivalente en BD) tiene el formato:

```
PERIODO, CATEGORIA, LINEA, PEDIDOS, VOLUMEN_KG, MONTO_SOLES
```

- **PERIODO:** formato `YYMM` (ej: `2501` = enero 2025)
- **CATEGORIA:** equivale a "marca" en la BD SQL (ej: `POLLO`, `COLGATE`)
- **LINEA:** subcategoría dentro de la categoría
- **PEDIDOS:** cantidad de pedidos (proxy de cobertura)
- **VOLUMEN_KG:** kilogramos vendidos
- **MONTO_SOLES:** ventas en soles sin IGV

`export_data_ssff_rutas.csv` tiene estructura similar pero incluye `ccod_cli` (cliente) y `ccod_ruta` (ruta), con ~1.29M registros — cargarlo con `dtype` explícito para evitar inferencia lenta.

---

## Base de datos SQL

**Conexión:** SQL Server, base `eAuren`
**Tabla principal:** `[eAuren].[dbo].[base_com]`
**Tabla de productos:** `[comercial_productos]`

Ver `consulta_historico_SSFF.sql` para la query canónica:
- Rango de periodos: 2501–2604
- Filtros: `estado = activo`, `tipo = '0003'`, rutas válidas

---

## Dominio de negocio

### Categorías del proyecto (20 marcas activas)
`ACCESORIOS, ANDINA, CERDO, COLGATE, DERMODIS, DULFINA, HIGIENE Y CUIDADO, HOMEPRO PERU, HUEVO, KIMBERLY, LA CORONA, LA PATRONA, MEDIFARMA, PAVO, POLLO, PROCESADOS, RINTI, TAMBOS PERU, VERDUM, YICHANG`

### Categorías San Fernando (SSFF) — ~45-50% de ventas totales
`CERDO, PAVO, POLLO, HUEVO, PROCESADOS`

**Líneas de PROCESADOS** (críticas para análisis de volumen):
- `EMBUTIDOS`
- `CONGELADOS` = `ELABORADOS` + `SEMIELABORADOS` + `PRECOCIDOS`

### Fuerzas de Ventas (FFVV) — 79 rutas en total
| FFVV | Rutas | Observación |
|------|-------|-------------|
| F8   | 55    | Principal; absorbe F6 y F7 desde 08/03/2026 |
| M0   | 14    | Recibe volumen PROCESADOS junto con F8 |
| K0   | 7     | Kimberly — SIN cuota en categorías SSFF |
| P0   | 1     | — |
| V0   | 2     | V001, V002 |

> Las FFVV `F6` y `F7` se fusionaron en `F8` desde el **08/03/2026** con reestructuración total de carteras.

### Reestructuración de carteras (08/03/2026)
- Se integraron ~5.000 clientes nuevos del "censo" (estudio de mercado) — marcados con `es_censo = 1` en `cartera_simplificada.xlsx`
- Las rutas de los vendedores fueron redistribuidas geográficamente
- **Impacto en cuotas:** calcular por cliente → sumar a la RUTA actual del cliente, ignorando la ruta histórica

---

## Metodologías de cuota mensual

Las cuotas se calculan sobre **VOLUMEN (kg)**, **MONTO (S/)** y **COBERTURA (clientes únicos que compraron)**.

### Tipo 1 — Promedio/Máximo de últimos N meses (implementada actualmente)
```
CUOTA = PROMEDIO( PROMEDIO(M-3, M-2, M-1*), MAX(M-3, M-2, M-1*) )
```
*M-1 puede ser proyectado si el mes aún no cerró.*

Ejemplo para mayo-2026 (febrero excluido por ser mes atípico bajo):
```
PROMEDIO( PROMEDIO(202603, 202604*), MAX(202603, 202604*) )
```

### Tipo 2 — Tendencia + estacionalidad
```
CUOTA = PROMEDIO( PROMEDIO(M-2, M-1*), PROMEDIO(M-14, M-13), MISMO_MES_AÑO_ANTERIOR )
```

Ejemplo para mayo-2026:
```
PROMEDIO( PROMEDIO(202603, 202604*), PROMEDIO(202503, 202504), 202505 )
```

Ver `ANALISIS_METODOLOGIAS_MAYO_2026.md` para comparativa detallada: Tipo 2 proyecta +5.47% (recomendado por gerencia).

### Distribución por FFVV y Ruta
1. Calcular el peso porcentual de cada FFVV sobre el total de la categoría (mes de referencia).
2. Dentro de cada FFVV, calcular el peso porcentual de cada RUTA.
3. Desde la reestructuración (08/03/2026): calcular por cliente → sumar a la RUTA actual del cliente.

### Crecimiento esperado
La gerencia espera un crecimiento aproximado del **5% mensual** (relativo al contexto). Los clientes del censo deben reflejarse con potencial de crecimiento incremental.

---

## Reglas de implementación
- Evitar duplicar lógica → usar funciones reutilizables
- Al adaptar el script a un nuevo mes: actualizar periodos de referencia, nombre del archivo de salida y cuotas oficiales
- Comentarios en español, solo lo esencial
- Código primero, explicación después
- Documentar en este archivo cualquier cambio que afecte al flujo de cálculo
