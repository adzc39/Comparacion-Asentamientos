# PROMPT ARQUITECTURA EXCEL v3.0: ANÁLISIS COMPARATIVO AVANZADO
## Asentamientos Diferenciales - Tres Modelos de Cimentación

---

## ÍNDICE

1. [Resumen Ejecutivo](#resumen-ejecutivo)
2. [Estructura General del Archivo](#estructura-general-del-archivo)
3. [Hoja INPUT](#hoja-input)
4. [Tabla 1.0: Tabla de Nodos](#tabla-10-tabla-de-nodos)
5. [Tabla 2.0: Tabla de Conexiones](#tabla-20-tabla-de-conexiones)
6. [Tabla 3.0: Comparación de Asentamientos](#tabla-30-comparación-de-asentamientos)
7. [Arquitectura de Entrelazamiento](#arquitectura-de-entrelazamiento)
8. [Especificaciones de Fórmulas](#especificaciones-de-fórmulas)
9. [Flujo de Trabajo](#flujo-de-trabajo)
10. [Validaciones y Control de Calidad](#validaciones-y-control-de-calidad)

---

## RESUMEN EJECUTIVO

### Objetivo
Crear un archivo Excel que analice **asentamientos diferenciales** de cimentaciones comparando **tres modelos** (Empotrado, Winkler Corregido, Barkan Corregido) mediante:

- **Tabla 1.0:** Nodos con propiedades ETABS + asentamientos (INPUT)
- **Tabla 2.0:** Conexiones entre nodos con cálculo de distancias
- **Tabla 3.0:** Análisis avanzado con agrupaciones, rangos, diferencias y comparativas

### Principios de Diseño

✓ **Entrelazamiento Total:** Todas las fórmulas referenciadas, cero valores hardcodeados  
✓ **Modularidad:** Estructura idéntica para 3 modelos, independientes entre sí  
✓ **Trazabilidad:** Cada valor remite a su origen (INPUT o cálculo)  
✓ **Escalabilidad:** Agregar nodos/conexiones sin quebrar estructura  
✓ **Análisis Sofisticado:** Tabla 3.0 con agrupaciones, percentiles, máximos locales y globales

### Entidades Principales

| Elemento | Cantidad | Descripción |
|----------|----------|-------------|
| **Hoja INPUT** | 1 | Fuente única de asentamientos por nodo/modelo |
| **Modelos** | 3 | Empotrado, Winkler, Barkan |
| **Tablas por modelo** | 3 | 1.0 (Nodos), 2.0 (Conexiones), 3.0 (Comparación) |
| **Hojas totales** | 10+ | INPUT + 9 análisis (+ resumen comparativo opcional) |
| **Nodos** | 27 (ej.) | Columnas de estructura |
| **Conexiones** | 30 (ej.) | Vigas de cimentación |

---

## ESTRUCTURA GENERAL DEL ARCHIVO

### Organización de Hojas

```
├─ INPUT
│  └─ Datos de entrada: asentamientos por nodo y modelo
│
├─ MODELO 1: EMPOTRADO
│  ├─ EMPO_1_Nodos
│  ├─ EMPO_2_Conexiones
│  └─ EMPO_3_Comparacion
│
├─ MODELO 2: WINKLER CORREGIDO
│  ├─ WINKLER_1_Nodos
│  ├─ WINKLER_2_Conexiones
│  └─ WINKLER_3_Comparacion
│
├─ MODELO 3: BARKAN CORREGIDO
│  ├─ BARKAN_1_Nodos
│  ├─ BARKAN_2_Conexiones
│  └─ BARKAN_3_Comparacion
│
└─ RESUMEN_COMPARATIVO (opcional)
   └─ Tablas de comparación inter-modelos
```

### Convención de Nombres

- **Hojas Tabla 1.0:** `[MODELO]_1_Nodos`
  - Ejemplo: `EMPO_1_Nodos`, `WINKLER_1_Nodos`, `BARKAN_1_Nodos`

- **Hojas Tabla 2.0:** `[MODELO]_2_Conexiones`
  - Ejemplo: `EMPO_2_Conexiones`, `WINKLER_2_Conexiones`, `BARKAN_2_Conexiones`

- **Hojas Tabla 3.0:** `[MODELO]_3_Comparacion`
  - Ejemplo: `EMPO_3_Comparacion`, `WINKLER_3_Comparacion`, `BARKAN_3_Comparacion`

---

## HOJA INPUT

### Propósito
**Fuente única de verdad** para todos los asentamientos. NINGUNA otra hoja contiene valores de asentamiento; todo se extrae de INPUT mediante VLOOKUP.

### Estructura Física

| Columna | Encabezado | Tipo | Rango | Origen | Validación |
|---------|-----------|------|-------|--------|-----------|
| **A** | ID_NODO | Entero | A5:A31 | ETABS | 1–27 (único) |
| **B** | Empotrado (mm) | Decimal | B5:B31 | ETABS OUTPUT | -50 a 50 |
| **C** | Winkler Corregido (mm) | Decimal | C5:C31 | ETABS OUTPUT | -50 a 50 |
| **D** | Barkan Corregido (mm) | Decimal | D5:D31 | ETABS OUTPUT | -50 a 50 |

### Formato de Datos

```
ID_NODO | Empotrado (mm) | Winkler Corregido (mm) | Barkan Corregido (mm)
───────┼────────────────┼───────────────────────┼──────────────────────
   1   |     -2.1       |        -1.8           |        -1.9
   2   |     -2.0       |        -1.7           |        -1.8
  ...  |      ...       |         ...           |         ...
  27   |     -1.9       |        -1.6           |        -1.7
```

### Reglas de Entrada

- **Obligatorio:** Todos los 27 nodos deben tener asentamiento en las 3 columnas
- **Signo:** Negativo = hundimiento, positivo = levantamiento (convención estructural)
- **Unidad:** Milímetros (mm), con precisión a 0.1 mm
- **Actualización:** Solo se modifica cuando hay nuevo análisis ETABS
- **Celdas de entrada:** B5:D31 (amarillas, resto protegido)

---

## TABLA 1.0: TABLA DE NODOS

### Propósito
Consolidar propiedades de cada nodo (geométricas, ETABS) junto con asentamiento vinculado desde INPUT.

### Ubicación
- **Empotrado:** Hoja `EMPO_1_Nodos`
- **Winkler:** Hoja `WINKLER_1_Nodos`
- **Barkan:** Hoja `BARKAN_1_Nodos`

### Estructura (9 columnas)

| Columna | Encabezado | Tipo | Fuente | Fórmula/Referencia | Notas |
|---------|-----------|------|--------|-------------------|-------|
| **A** | ID Nodo | Entero | ETABS | Constante (1–27) | Identificador único |
| **B** | Coord X (mm) | Decimal | ETABS | Constante | Sistema de coordenadas del modelo |
| **C** | Coord Y (mm) | Decimal | ETABS | Constante | Sistema de coordenadas del modelo |
| **D** | Asentamiento (mm) | Decimal | INPUT | `=VLOOKUP(A5, INPUT!$A$5:$D$31, [col], FALSE)` | Col: 2=Emp, 3=Win, 4=Bar |
| **E** | Nombre (E2K) | Texto | ETABS | Constante | Identificador en modelo E2K |
| **F** | Columna (E2K) | Texto | ETABS | Constante | Etiqueta de elemento (ej: C1, C27) |
| **G** | Tipo Elemento | Texto | ETABS | Constante | "Columna", "Viga", "Nudo" |
| **H** | Altura (E2K) | Texto | ETABS | Constante | "Nivel 0", "Base", "Eje -1" |
| **I** | Observaciones | Texto | Manual | Entrada libre | Contexto del nodo |

### Fórmula VLOOKUP Detallada

**Para Empotrado (columna D):**
```excel
=VLOOKUP(A5, INPUT!$A$5:$D$31, 2, FALSE)
```
- Busca ID en A5
- En tabla INPUT filas 5–31
- Retorna columna B (Empotrado)
- Búsqueda exacta (FALSE)

**Para Winkler (columna D):**
```excel
=VLOOKUP(A5, INPUT!$A$5:$D$31, 3, FALSE)
```
- Mismo procedimiento, columna C (Winkler)

**Para Barkan (columna D):**
```excel
=VLOOKUP(A5, INPUT!$A$5:$D$31, 4, FALSE)
```
- Mismo procedimiento, columna D (Barkan)

### Validación Tabla 1.0

- ✓ Todos los 27 nodos presentes (A5:A31)
- ✓ Columna D (Asentamiento) sin errores #N/A o #REF!
- ✓ Valores coherentes entre modelos (Empotrado ≤ Winkler ≤ Barkan típicamente)
- ✓ Coordenadas idénticas en EMPO_1, WINKLER_1, BARKAN_1

---

## TABLA 2.0: TABLA DE CONEXIONES

### Propósito
Especificar pares de nodos conectados (vigas de cimentación) y calcular distancia euclidiana entre ellos.

### Ubicación
- **Empotrado:** Hoja `EMPO_2_Conexiones`
- **Winkler:** Hoja `WINKLER_2_Conexiones`
- **Barkan:** Hoja `BARKAN_2_Conexiones`

### Estructura (8 columnas)

| Columna | Encabezado | Tipo | Fuente | Fórmula/Referencia | Notas |
|---------|-----------|------|--------|-------------------|-------|
| **A** | Nodo i | Entero | Conexiones | Constante | Nodo inicial de viga |
| **B** | Nodo j | Entero | Conexiones | Constante | Nodo final de viga |
| **C** | Coord Xi (mm) | Decimal | Tabla 1.0 | `=VLOOKUP(A5, [Tabla1]!$A$5:$B$31, 2, FALSE)` | X de nodo i |
| **D** | Coord Yi (mm) | Decimal | Tabla 1.0 | `=VLOOKUP(A5, [Tabla1]!$A$5:$C$31, 3, FALSE)` | Y de nodo i |
| **E** | Coord Xj (mm) | Decimal | Tabla 1.0 | `=VLOOKUP(B5, [Tabla1]!$A$5:$B$31, 2, FALSE)` | X de nodo j |
| **F** | Coord Yj (mm) | Decimal | Tabla 1.0 | `=VLOOKUP(B5, [Tabla1]!$A$5:$C$31, 3, FALSE)` | Y de nodo j |
| **G** | Distancia (mm) | Decimal | Calculada | `=SQRT((E5-C5)^2 + (F5-D5)^2)` | Euclidiana 2D |
| **H** | Observaciones | Texto | Manual | Entrada libre | Tipo de viga, contexto |

### Fórmulas de Coordenadas

**Para Xi (nodo i, coordenada X):**
```excel
=VLOOKUP(A5, [TABLA1]!$A$5:$B$31, 2, FALSE)
```
- Busca Nodo i (en A5) en Tabla 1.0
- Retorna columna B (Coord X)

**Para Yi (nodo i, coordenada Y):**
```excel
=VLOOKUP(A5, [TABLA1]!$A$5:$C$31, 3, FALSE)
```
- Busca Nodo i en Tabla 1.0
- Retorna columna C (Coord Y)

**Para Xj y Yj:** Replicar con B5 (nodo j)

### Fórmula de Distancia

```excel
=SQRT((E5-C5)^2 + (F5-D5)^2)
```

**Desglose:**
- `(E5-C5)` = Diferencia de X entre nodos
- `(F5-D5)` = Diferencia de Y entre nodos
- Cada término se eleva al cuadrado
- Se suman: `(ΔX)² + (ΔY)²`
- Se extrae raíz cuadrada: √(suma)
- Resultado: Distancia euclidiana en mm

**Ejemplo:**
```
Nodo 1: X=-14600.9, Y=-9526.3
Nodo 3: X=-17776.3, Y=-4026.3

ΔX = -17776.3 - (-14600.9) = -3175.4
ΔY = -4026.3 - (-9526.3) = 5500.0

Distancia = √((-3175.4)² + (5500)²)
          = √(10,083,785 + 30,250,000)
          = √40,333,785
          = 6350.9 mm
```

### Propiedades Tabla 2.0

- **Número de filas:** 30 (conexiones) + encabezados
- **Distancias:** Todas positivas, nunca cero
- **Independencia:** Mismas distancias en EMPO_2, WINKLER_2, BARKAN_2
- **Actualización:** Automática si cambian coordenadas en Tabla 1.0

---

## TABLA 3.0: COMPARACIÓN DE ASENTAMIENTOS

### Propósito (SOFISTICADO)
Análisis detallado de diferencias de asentamiento con:
- **Agrupación por nodo:** Para cada nodo i, mostrar todos sus conectados (j)
- **Rangos de asentamiento:** Mínimo y máximo por agrupación
- **Diferencias:** Δ = Asent_i - Asent_j (con signo, luego ordenadas)
- **Máxima diferencia local:** Por cada agrupación
- **Análisis global:** Todos los nodos ordenados menor a mayor, con diferencia vs. mínimo global

### Ubicación
- **Empotrado:** Hoja `EMPO_3_Comparacion`
- **Winkler:** Hoja `WINKLER_3_Comparacion`
- **Barkan:** Hoja `BARKAN_3_Comparacion`

### Estructura: SECCIÓN 1 - Por Agrupación de Nodo

**Para cada nodo i, crear bloque:**

```
┌──────────────────────────────────────────────────────────┐
│ AGRUPACIÓN NODO [i]                                      │
├──────────────────────────────────────────────────────────┤
│ Nodo i │ Asent i │ Nodo j │ Asent j │ Δ (i-j) │ % Δ    │
│ ──────┼─────────┼────────┼─────────┼─────────┼────────│
│   1   │  -2.1   │   3    │  -2.3   │  0.2    │ 9.52%  │
│   1   │  -2.1   │   6    │  -2.0   │  -0.1   │ -4.76% │
│   1   │  -2.1   │  19    │  -2.2   │  0.1    │ 4.76%  │
│ ─────┼─────────┼────────┼─────────┼─────────┼────────│
│ STAT │Mín:-2.3 │        │Máx:-2.0 │ ΔMáx=0.2│ 9.52%  │
└──────────────────────────────────────────────────────────┘
```

#### Columnas de Agrupación

| Columna | Encabezado | Tipo | Fórmula | Descripción |
|---------|-----------|------|---------|-------------|
| **A** | Nodo i | Entero | Constante | Nodo abordado |
| **B** | Asent i (mm) | Decimal | `=VLOOKUP(A, [T1]!A:D, 4)` | De Tabla 1.0 |
| **C** | Nodo j | Entero | Constante | Nodo conectado (desde Tabla 2.0) |
| **D** | Asent j (mm) | Decimal | `=VLOOKUP(C, [T1]!A:D, 4)` | De Tabla 1.0 |
| **E** | Δ (mm) | Decimal | `=B-D` | Diferencia con signo |
| **F** | % Δ | Porcentaje | `=IF(B=0, 0, E/ABS(B)*100)` | Cambio porcentual |

#### Estadística de Agrupación (Fila Resumen)

Por cada agrupación, agregar fila de estadísticas:

| Celda | Contenido | Fórmula |
|-------|-----------|---------|
| "STAT" | Etiqueta | Constante |
| Asent i | Mín asentamiento | `=MIN([rango Asent j de agrupación])` |
| — | — | — |
| Asent j | Máx asentamiento | `=MAX([rango Asent j de agrupación])` |
| Δ Máx | Mayor diferencia | `=MAX(ABS([rango Δ de agrupación]))` |
| % Máx | % del máximo | `=MAX([rango % Δ de agrupación])` |

### Estructura: SECCIÓN 2 - Análisis Global

**Después de todas las agrupaciones, tabla resumen global:**

```
┌─────────────────────────────────────────────────────────┐
│ ANÁLISIS GLOBAL: ASENTAMIENTOS ORDENADOS               │
├─────────────────────────────────────────────────────────┤
│ Nodo │ Asent (mm) │ Δ vs Mín Global │ % Δ vs Mín      │
│ ────┼────────────┼─────────────────┼─────────────────│
│  15 │   -1.4     │      0.0        │    0.0%         │ ← Mínimo
│  16 │   -1.4     │      0.0        │    0.0%         │
│  17 │   -1.5     │     -0.1        │   -7.1%         │
│  ...│   ...      │      ...        │    ...          │
│   3 │   -2.3     │      0.9        │   64.3%         │ ← Máximo
└─────────────────────────────────────────────────────────┘
```

#### Columnas Análisis Global

| Columna | Encabezado | Tipo | Fórmula | Descripción |
|---------|-----------|------|---------|-------------|
| **A** | Nodo | Entero | Constante | ID nodo (1–27) |
| **B** | Asent (mm) | Decimal | `=VLOOKUP(A, [T1]!A:D, 4)` | De Tabla 1.0 |
| **C** | Δ vs Mín Global | Decimal | `=B - MIN([T1]!D:D)` | Diferencia respecto mínimo global |
| **D** | % Δ vs Mín | Porcentaje | `=IF(MIN([T1]!D:D)=0, 0, C/ABS(MIN([T1]!D:D))*100)` | Cambio porcentual |

**Ordenamiento:** Columna B ascendente (menor a mayor asentamiento)

### Gráficos Incluidos en Tabla 3.0

#### Gráfico 1: Barras de Diferencias por Agrupación

**Tipo:** Gráfico de barras horizontal  
**Eje X:** Diferencia Δ (mm)  
**Eje Y:** Pares (Nodo i - Nodo j)  
**Datos:** Columnas E y C (pares de nodos)

**Propósito:** Visualizar magnitud de diferencias dentro de cada agrupación

#### Gráfico 2: Análisis Global - Cascada de Asentamientos

**Tipo:** Gráfico de columnas  
**Eje X:** Nodos ordenados por asentamiento  
**Eje Y:** Asentamiento (mm) o Δ vs Mínimo global (mm)  
**Datos:** Columnas A y B (o C)

**Propósito:** Ver evolución de asentamientos globales

#### Gráfico 3: Máximas Diferencias por Agrupación

**Tipo:** Gráfico de barras o radar  
**Datos:** % Δ máximo por nodo i  
**Propósito:** Identificar nodos críticos

### Validaciones Tabla 3.0

- ✓ Cada nodo i tiene todas sus conexiones listadas
- ✓ Diferencias (Δ) tienen signo correcto: positivo = i > j
- ✓ % Δ porcentajes acorde a diferencia y magnitud
- ✓ Análisis global contiene 27 nodos únicos
- ✓ Mínimo global aparece una sola vez (primer lugar)
- ✓ Máximo global aparece una sola vez (último lugar)
- ✓ Gráficos se actualizan automáticamente

---

## ARQUITECTURA DE ENTRELAZAMIENTO

### Flujo de Datos: Cadena Completa

```
INPUT (Asentamientos por nodo/modelo)
  │
  ├─→ Columna B (Empotrado)
  │   ├─→ EMPO_1_Nodos!D (VLOOKUP col 2)
  │   │   ├─→ EMPO_2_Conexiones!C,D,E,F (VLOOKUP coords)
  │   │   └─→ EMPO_3_Comparacion!B,D (VLOOKUP)
  │   │       ├─→ Cálculo Δ (columna E)
  │   │       ├─→ Cálculo % Δ (columna F)
  │   │       ├─→ Gráfico 1 (barras)
  │   │       ├─→ Análisis Global (28–55 aprox)
  │   │       └─→ Gráfico 2 y 3
  │
  ├─→ Columna C (Winkler)
  │   └─→ WINKLER_1/2/3 (estructura idéntica)
  │
  └─→ Columna D (Barkan)
      └─→ BARKAN_1/2/3 (estructura idéntica)
```

### Regla Fundamental

**TODOS los valores están vinculados a INPUT mediante fórmulas. NO hay constantes en Tablas 1/2/3 excepto:**
- Coordenadas (extraídas de ETABS, pero no cambian)
- Propiedades ETABS (nombre, tipo, altura)
- Topología de conexiones (pares de nodos)

### Cambio en Cascada

**Ejemplo: Cambiar INPUT!B5 de -2.1 a -2.5**

```
INPUT!B5: -2.1 → -2.5 (usuario cambia)
  ↓ Automático
EMPO_1_Nodos!D5: -2.1 → -2.5 (VLOOKUP actualiza)
  ↓ Automático
EMPO_3_Comparacion!B5-B7: Actualiza diferencias
  ↓ Automático
Gráficos: Se redibujan automáticamente
  ↓ Automático
ANÁLISIS GLOBAL: Recalcula rangos y % Δ

TODO en < 1 segundo
```

---

## ESPECIFICACIONES DE FÓRMULAS

### VLOOKUP Estándar para Asentamientos

```excel
=VLOOKUP(A5, INPUT!$A$5:$D$31, 2, FALSE)
```

**Parámetros:**
- `A5`: Referencia de nodo a buscar
- `INPUT!$A$5:$D$31`: Tabla completa (rango absoluto)
- `2`: Número de columna (2=Emp, 3=Win, 4=Bar)
- `FALSE`: Búsqueda exacta

**Variantes por modelo:**

| Modelo | Columna | VLOOKUP |
|--------|---------|---------|
| Empotrado | 2 | `=VLOOKUP(A5, INPUT!$A$5:$D$31, 2, FALSE)` |
| Winkler | 3 | `=VLOOKUP(A5, INPUT!$A$5:$D$31, 3, FALSE)` |
| Barkan | 4 | `=VLOOKUP(A5, INPUT!$A$5:$D$31, 4, FALSE)` |

### VLOOKUP para Coordenadas

```excel
=VLOOKUP(A5, EMPO_1_Nodos!$A$5:$B$31, 2, FALSE)
```

**Usado en Tabla 2.0 (Conexiones) para replicar coordenadas desde Tabla 1.0**

- Busca ID nodo
- Retorna columna B (Coord X) o C (Coord Y)
- Referencia entre hojas con $

### Fórmulas de Cálculo

#### Distancia Euclidiana
```excel
=SQRT((E5-C5)^2 + (F5-D5)^2)
```

#### Diferencia de Asentamiento
```excel
=B5 - D5
```
(Con signo: positivo si i > j)

#### Diferencia Porcentual
```excel
=IF(B5=0, 0, (B5-D5)/ABS(B5)*100)
```
(Evita división por cero)

#### Diferencia vs Mínimo Global
```excel
=B5 - MIN($B$5:$B$31)
```
(Referencia absoluta al mínimo)

#### Porcentaje vs Mínimo Global
```excel
=IF(MIN($B$5:$B$31)=0, 0, (B5-MIN($B$5:$B$31))/ABS(MIN($B$5:$B$31))*100)
```

### Funciones de Agregación (en Tablas de Estadística)

```excel
Mínimo de grupo:     =MIN([rango de agrupación])
Máximo de grupo:     =MAX([rango de agrupación])
Promedio:            =AVERAGE([rango])
Desv. Estándar:      =STDEV([rango])
Máximo Absoluto:     =MAX(ABS([rango]))
```

---

## FLUJO DE TRABAJO

### Fase 1: Preparación

1. **Abrir archivo Excel:** `Asentamientos_Modelos_Comparativa_v3.xlsx`
2. **Ir a hoja INPUT**
3. **Rellenar B5:D31** con asentamientos extraídos de ETABS
   - Columna B: Modelo Empotrado
   - Columna C: Modelo Winkler
   - Columna D: Modelo Barkan

### Fase 2: Verificación Tabla 1.0

**Para cada modelo (Empotrado, Winkler, Barkan):**

1. Ir a hoja `[MODELO]_1_Nodos`
2. Verificar que columna D (Asentamiento) muestre valores sin errores
3. Confirmar que asentamientos varían entre modelos (no iguales)
4. Revisar que columnas B, C (coordenadas) son idénticas entre modelos

### Fase 3: Verificación Tabla 2.0

**Para cada modelo:**

1. Ir a hoja `[MODELO]_2_Conexiones`
2. Verificar que columna G (Distancia) tiene valores > 0
3. Confirmar que distancias son idénticas entre modelos
4. Revisar 2–3 distancias manualmente con calculadora

### Fase 4: Análisis Tabla 3.0

**Para cada modelo:**

1. Ir a hoja `[MODELO]_3_Comparacion`
2. Revisar agrupaciones: ¿Cada nodo i tiene todas sus j?
3. Analizar columna Δ: ¿Valores coherentes?
4. Ver análisis global: ¿Rango de asentamientos visible?
5. Revisar gráficos: ¿Se actualizan con cambios en INPUT?

### Fase 5: Comparación Inter-Modelos

1. Crear tabla resumen (nueva hoja RESUMEN_COMPARATIVO)
2. Extraer máximas diferencias, % máximos de cada modelo
3. Comparar patrones de concentración geográfica
4. Recomendación: ¿Cuál modelo es más conservador?

---

## VALIDACIONES Y CONTROL DE CALIDAD

### Checklist de Integridad

- [ ] INPUT tiene 27 nodos (A5:A31)
- [ ] INPUT tiene asentamientos en B5:D31 (sin celdas vacías)
- [ ] Tabla 1.0 de Empotrado tiene asentamientos ≤ Winkler ≤ Barkan (típicamente)
- [ ] Tabla 1.0 de Winkler y Barkan tienen **mismas coordenadas** que Empotrado
- [ ] Tabla 2.0: 30 conexiones (1–30 filas de datos)
- [ ] Tabla 2.0: Distancias idénticas en EMPO, WINKLER, BARKAN
- [ ] Tabla 3.0: Sin errores #N/A, #REF!, #DIV/0!
- [ ] Tabla 3.0: Agrupaciones completas (cada nodo i con todos sus j)
- [ ] Gráficos: Se actualizan al cambiar INPUT
- [ ] Análisis Global: 27 nodos únicos, ordenados menor a mayor

### Prueba de Cambio en Cascada

**Procedimiento:**

1. Abrir INPUT, celda B5 (Empotrado, Nodo 1)
2. Cambiar valor de `-2.1` a `-2.5`
3. Presionar ENTER
4. Ir a EMPO_1_Nodos!D5 → ¿Muestra -2.5? (✓ o ✗)
5. Ir a EMPO_3_Comparacion, agrupación nodo 1 → ¿Diferencias recalculadas? (✓ o ✗)
6. Ver gráficos → ¿Se redibujaron? (✓ o ✗)

**Resultado esperado:** Todas las marcas ✓

### Validación de Topología

**Verificar que Tabla 2.0 representa correctamente las conexiones:**

1. Extraer de archivo original: lista de conexiones (Nodo i, Nodo j)
2. Comparar con filas A–B de Tabla 2.0
3. Confirmar cantidad: 30 conexiones (o el número del proyecto)
4. Verificar: no hay duplicados o conexiones invertidas (1→3 ≠ 3→1 en lógica)

---

## DECISIÓN: HOJAS SEPARADAS VS. UNA SOLA HOJA

### Recomendación: **HOJAS SEPARADAS**

**Razones:**

1. **Legibilidad:** Cada tabla ocupa 30–50 filas + gráficos
2. **Performance:** Menos cálculos por hoja = recalc más rápido
3. **Organización:** Modelo Empotrado completo = 3 hojas; no mezclado
4. **Filtros/Sorts:** Más fácil en tabla aislada que compartida
5. **Gráficos:** Mejor espacio para gráficos grandes
6. **Documentación:** Cada hoja con encabezados claros

**Estructura final:**
- 1 hoja INPUT
- 3 modelos × 3 tablas = 9 hojas
- 1 hoja RESUMEN_COMPARATIVO (opcional)
- **Total: 10–11 hojas**

---

## COMPARACIÓN ENTRE MODELOS

### Tabla de Comparativa (Hoja RESUMEN_COMPARATIVO)

```
┌────────────────────────────────────────────────────────┐
│ COMPARACIÓN DE MODELOS                                 │
├────────────────────────────────────────────────────────┤
│ Métrica                 │ Empotrado │ Winkler │ Barkan │
├─────────────────────────┼───────────┼─────────┼────────│
│ Asent Mín Global (mm)   │   -1.7    │  -1.4   │  -1.5  │
│ Asent Máx Global (mm)   │   -2.3    │  -2.0   │  -2.1  │
│ Rango Asent (mm)        │   0.6     │  0.6    │  0.6   │
│ Δ Máx por Agrupación    │   0.2     │  0.25   │  0.28  │
│ % Δ Máx por Agrupación  │   9.5%    │  14.3%  │  18.7% │
│ Δ Máx Global            │   0.9     │  0.9    │  0.9   │
│ % Δ Máx Global          │   64.3%   │  64.3%  │  64.3% │
│ Nodos Críticos (Δ>0.2)  │    1      │    2    │    3   │
│ Recomendación           │ Viable    │ Viable  │ Revisar│
└────────────────────────────────────────────────────────┘
```

**Fórmulas para esta tabla:**

```excel
Asent Mín:      =MIN(EMPO_1_Nodos!D:D)
Asent Máx:      =MAX(EMPO_1_Nodos!D:D)
Δ Máx Local:    =MAX(EMPO_3_Comparacion!E:E)
Nodos Críticos: =COUNTIF(EMPO_3_Comparacion!E:E, ">0.2")
```

---

## PROPIEDADES TÉCNICAS FINALES

### Precisión de Cálculos

| Elemento | Decimales | Formato |
|----------|-----------|---------|
| Asentamientos | 1 | 0.0 mm |
| Coordenadas | 1 | 0.0 mm |
| Distancias | 1 | 0.0 mm |
| Diferencias | 1 | 0.0 mm |
| Porcentajes | 1 | 0.0% |

### Rango de Valores Esperados

| Parámetro | Mín | Típico | Máx |
|-----------|-----|--------|-----|
| Asentamientos | -50 | -2 | 0 |
| Coordenadas | -20000 | 0 | 20000 |
| Distancias | 2000 | 8000 | 25000 |
| Diferencias Δ | 0 | 0.3 | 5 |
| % Δ | 0% | 10% | 100% |

### Limitaciones Implícitas

- Análisis **2D** (X–Y), no incluye Z
- **Estático**, no dinámico
- Nodos **discretos**, sin interpolación
- Topología **predefinida**, 30 conexiones ejemplo
- Entrelazamiento **sin validación cruzada** (usuario responsable de precisión INPUT)

---

## PRÓXIMOS PASOS PARA IMPLEMENTACIÓN

### Paso 1: Validar Conexiones
Requiere compartir archivo original con lista de conexiones (Nodo i, Nodo j) en formato Excel.

### Paso 2: Crear Estructura Base
Python + openpyxl para generar hojas INPUT + Tabla 1.0/2.0/3.0 con fórmulas.

### Paso 3: Implementar Fórmulas
Insertar todas las VLOOKUP, cálculos, estadísticas.

### Paso 4: Agregar Gráficos
Gráficos dinámicos que se actualicen con datos.

### Paso 5: Validar Entrelazamiento
Prueba de cambio en cascada para confirmar propagación.

### Paso 6: Documentar
Manual de usuario con ejemplos y troubleshooting.

---

## NOTAS FINALES

### Principios de Diseño Reafirmados

✓ **Cero hardcoding:** Todos los valores vienen de fórmulas  
✓ **Modularidad:** 3 modelos independientes, misma estructura  
✓ **Trazabilidad:** Cada valor remite a INPUT o cálculo visible  
✓ **Escalabilidad:** Agregar nodos/conexiones sin quebrar  
✓ **Análisis sofisticado:** Tabla 3.0 con agrupaciones, rangos, gráficos, comparativas

### Diferencia vs. Versión Anterior

| Aspecto | v2.0 | **v3.0** |
|---------|------|---------|
| Tabla 3.0 | Simple (3 columnas) | **Expandida (6+ columnas, agrupaciones, análisis global)** |
| Agrupaciones | No | **Sí (por nodo i)** |
| Rangos locales | No | **Sí (Mín/Máx por agrupación)** |
| Análisis global | No | **Sí (todos los nodos ordenados)** |
| Gráficos | Ninguno | **3 gráficos integrados** |
| Comparativa inter-modelos | No | **Sí (hoja RESUMEN_COMPARATIVO)** |

---

**Versión:** 3.0  
**Fecha:** Julio 2025  
**Propósito:** Especificación detallada de arquitectura Excel relacional avanzada  
**Estado:** Listo para implementación

**Próximo paso:** Compartir archivo de conexiones para validar topología.
