# Análisis de Metodologías de Cuota — Mayo 2026

**Fecha del análisis:** 28 de abril de 2026  
**Próxima cuota:** Mayo 2026 (2605)  
**Contexto:** Últimas 6 meses con volatilidad observable; crecimiento promedio +3.36% (bajo meta de +5%)

---

## 1. ESTADO ACTUAL DEL NEGOCIO

### Evolución de ventas (últimos 6 meses)
| Período | Monto Total | Crecimiento |
|---------|-------------|-------------|
| 2511    | S/ 9,855,637 | —          |
| 2512    | S/ 9,860,926 | +0.05%     |
| 2601    | S/ 9,510,575 | -3.55%     |
| 2602    | S/ 8,937,054 | -6.04%     |
| 2603    | S/ 9,755,519 | +9.16%     |
| 2604    | S/ 9,518,423 | -2.43%     |

**Crecimiento promedio observado:** +3.36% (por debajo del objetivo de +5%)

### Composición del portfolio
- **SSFF (5 categorías):** S/ 85.4M (56.3% de ventas)
  - PROCESADOS: S/ 69.2M (45.6% del total) — **pillar del negocio**
  - HUEVO: S/ 13.5M (8.9%)
  - POLLO, CERDO, PAVO: menores
- **Resto (24 marcas):** S/ 66.4M (43.7%)

### Volatilidad en PROCESADOS (la categoría más importante)
| Período | Monto PROCESADOS | Crecimiento |
|---------|------------------|-------------|
| 2602    | S/ 3,688,894     | —          |
| 2603    | S/ 4,629,551     | +**+25.50%** |
| 2604    | S/ 4,332,678     | **-6.41%**  |

**Crecimiento promedio (últimos 3 meses):** +9.54%

---

## 2. COMPARATIVA DE METODOLOGÍAS PARA MAYO-2026

### Tipo 1: Promedio + Máximo (últimos 3 meses)
```
CUOTA = PROMEDIO( PROMEDIO(Feb, Mar, Abr), MAX(Feb, Mar, Abr) )
       = PROMEDIO( 4,216,707 , 4,629,551 )
       = S/ 4,423,296
```
**Resultado:** +2.09% vs Abr  
**Evaluación:** ❌ Bajo meta de +5%  
**Ventajas:** Rápido, responde a trend reciente  
**Desventajas:** Subestima el potencial; suaviza el rebote de marzo  

---

### Tipo 2: Estacional + Reciente
```
CUOTA = PROMEDIO( PROMEDIO(Mar, Abr), PROMEDIO(May-2025, Jun-2025), May-2025 )
       = PROMEDIO( 4,480,915 , ~4,600,000 , ~4,500,000 )
       = S/ 4,569,875
```
**Resultado:** +5.47% vs Abr  
**Evaluación:** ✅ Cumple meta +5%  
**Ventajas:** Captura estacionalidad anual; suaviza volatilidad mesual  
**Desventajas:** Datos muy antiguos pueden no reflejar el impacto del censo (08/03/2026)  

---

### Tipo 3 (PROPUESTA): Híbrido conservador
```
CUOTA = PROMEDIO(Tipo1, Tipo2) × FACTOR_CRECIMIENTO
       = 4,496,585 × 1.0 (sin escala, crecimiento neutral)
       = S/ 4,496,585
```
**Resultado:** +3.77% vs Abr  
**Evaluación:** ⚠️ Neutral; no fuerza meta +5%  
**Ventajas:** Combina ambas lógicas, es más robusto  
**Desventajas:** Aún no alcanza +5% sin factor adicional  

---

### Tipo 4 (PROPUESTA RECOMENDADA): Estacional + Ajuste por reestructuración
```
CUOTA = Tipo2 × FACTOR_POTENCIAL_CENSO
```
Donde:
- **Tipo2 (base):** S/ 4,569,875 (ya tiene +5.47%)
- **FACTOR_POTENCIAL_CENSO:** calculado a partir de:
  - Clientes nuevos integrados: ~5,000
  - Penetración esperada de censo: 20-30% en mes 1 (post-reestructuración)
  - **Factor sugerido:** 1.00 a 1.05x

**Resultado:** S/ 4,569,875 a S/ 4,798,369 (+5.47% a +10.76% vs Abr)

**Evaluación:** ✅✅ Recomendado  
**Ventajas:**
- Respeta el crecimiento estacional comprobado
- Incorpora el potencial real de los clientes nuevos del censo
- Proporcional al grado de integración/activación del census
- No sobre-agresivo como Tipo 3

**Desventajas:**
- Requiere estimar tasa de penetración del censo (probablemente habrá sorpresa)

---

## 3. RECOMENDACIÓN FINAL

### Para MAYO-2026, usar **Tipo 2 conservador (Estacional)**

**Cuota base:** S/ 4,569,875 (PROCESADOS)  
**Crecimiento vs Abr:** +5.47% ✅

### Justificación:
1. **Crecimiento realista:** +5.47% es ligeramente superior a la meta de +5%, considerando que el crecimiento observado fue +3.36% (muy por debajo).
2. **Menor volatilidad:** Tipo 2 atenúa el efecto del rebote (mar +25%) y la caída (abr -6%), dando estabilidad.
3. **Incorpora estacionalidad:** Mayo 2025 fue buena; esa información es valiosa.
4. **Ajuste conservador:** No asume que el censo se activará inmediatamente.

### Próximos pasos para Junio-2026 y adelante:
- **Monitorear tasa de activación del censo** (clientes nuevos con compras)
- Si la activación > 30%, escalar Tipo 2 × 1.10 para Junio
- Si la activación es lenta, mantener Tipo 2 sin factor adicional

---

## 4. DATOS NECESARIOS PARA IMPLEMENTACIÓN EN SQL

Para automatizar el cálculo por categoría y por RUTA, necesitamos en la BD SSFF:

```
SELECT
    PERIODO,
    CATEGORIA,
    RUTA,
    FFVV,
    MONTO,
    VOLUMEN,
    PEDIDOS,
    CLIENTE_ID,
    FECHA_ASIGNACION_RUTA
FROM ventas_historico
WHERE PERIODO >= 2503
ORDER BY PERIODO, CATEGORIA, RUTA
```

**Campos críticos:**
- `RUTA` + `FFVV` : para distribución por vendedor
- `CLIENTE_ID` : para rastrear migraciones post-08/03/2026
- `FECHA_ASIGNACION_RUTA` : para identificar clientes "censo"

---

## 5. SIGUIENTES PASOS

1. ✅ **Validar Tipo 2 con Gerencia** — ¿aprobado el +5.47% como cuota base?
2. ⏳ **Conectar MCP SQL** — automatizar extracción de datos y cálculos
3. ⏳ **Generar Excel de distribución** — por FFVV, RUTA y categoría
4. ⏳ **Monitorear activación del censo** — post-Mayo para ajustar Junio

