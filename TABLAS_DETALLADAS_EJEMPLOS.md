# TABLAS DETALLADAS: 1.0, 2.0, 3.0 CON EJEMPLOS CONCRETOS
## Arquitectura Excel v3.0

---

## TABLA 1.0: TABLA DE NODOS - VISTA COMPLETA

### Ubicación
- Hoja: `[MODELO]_1_Nodos` (EMPO_1_Nodos, WINKLER_1_Nodos, BARKAN_1_Nodos)
- Rango: A1:I31 (encabezados + 27 nodos)
- Actualización: Automática (VLOOKUP desde INPUT)

### Estructura Visual

```
┌─────┬──────────┬──────────┬────────────────┬────────────┬──────────────┬──────────────┬──────────┬─────────────────┐
│  A  │    B     │    C     │       D        │     E      │      F       │      G       │    H     │        I        │
├─────┼──────────┼──────────┼────────────────┼────────────┼──────────────┼──────────────┼──────────┼─────────────────┤
│ ID  │ Coord X  │ Coord Y  │  Asentamiento  │   Nombre   │   Columna    │    Tipo      │  Altura  │  Observaciones  │
│Node │  (mm)    │  (mm)    │     (mm)       │    (E2K)   │    (E2K)     │   Elemento   │  (E2K)   │                 │
├─────┼──────────┼──────────┼────────────────┼────────────┼──────────────┼──────────────┼──────────┼─────────────────┤
│  1  │-14600.9  │ -9526.3  │  -2.1 (calc)   │     C1     │     COL_01   │   Columna    │   Base   │  Esquina SO     │
│  2  │ 14600.9  │ -9526.3  │  -2.0 (calc)   │     C2     │     COL_02   │   Columna    │   Base   │  Esquina SE     │
│  3  │-17776.3  │ -4026.3  │  -2.3 (calc)   │     C3     │     COL_03   │   Columna    │   Base   │  Exterior SO    │
│  4  │-15026.3  │  736.9   │  -2.2 (calc)   │     C4     │     COL_04   │   Columna    │   Base   │  Lateral S      │
│  5  │-12276.3  │  5500.0  │  -2.1 (calc)   │     C5     │     COL_05   │   Columna    │   Base   │  Interior S     │
│  6  │ -8250.0  │ -9526.3  │  -2.0 (calc)   │     C6     │     COL_06   │   Columna    │   Base   │  Perimetral O   │
│  7  │ -5500.0  │ -4763.1  │  -1.9 (calc)   │     C7     │     COL_07   │   Columna    │   Base   │  Interior O     │
│  8  │ -2750.0  │   0.0    │  -1.8 (calc)   │     C8     │     COL_08   │   Columna    │   Base   │  Eje central O  │
│  9  │  2750.0  │   0.0    │  -1.8 (calc)   │     C9     │     COL_09   │   Columna    │   Base   │  Eje central E  │
│ 10  │  5500.0  │ -4763.1  │  -1.9 (calc)   │     C10    │     COL_10   │   Columna    │   Base   │  Interior E     │
│ 11  │  8250.0  │ -9526.3  │  -2.0 (calc)   │     C11    │     COL_11   │   Columna    │   Base   │  Perimetral E   │
│ 12  │ 17776.3  │ -4026.3  │  -2.3 (calc)   │     C12    │     COL_12   │   Columna    │   Base   │  Exterior SE    │
│ 13  │ 15026.3  │  736.9   │  -2.2 (calc)   │     C13    │     COL_13   │   Columna    │   Base   │  Lateral E      │
│ 14  │ 12276.3  │  5500.0  │  -2.1 (calc)   │     C14    │     COL_14   │   Columna    │   Base   │  Interior E     │
│ 15  │  2750.0  │ 11000.0  │  -1.7 (calc)   │     C15    │     COL_15   │   Columna    │   Base   │  Interior NE    │
│ 16  │ -2750.0  │ 11000.0  │  -1.7 (calc)   │     C16    │     COL_16   │   Columna    │   Base   │  Interior NO    │
│ 17  │ -7513.1  │  8250.0  │  -1.8 (calc)   │     C17    │     COL_17   │   Columna    │   Base   │  Lateral NO     │
│ 18  │  7513.1  │  8250.0  │  -1.8 (calc)   │     C18    │     COL_18   │   Columna    │   Base   │  Lateral NE     │
│ 19  │-13013.2  │ -6776.3  │  -2.2 (calc)   │     C19    │     COL_19   │   Columna    │   Base   │  Intermedia O   │
│ 20  │-10263.1  │ -2013.1  │  -2.1 (calc)   │     C20    │     COL_20   │   Columna    │   Base   │  Intermedia O   │
│ 21  │ -7513.1  │  2750.0  │  -1.9 (calc)   │     C21    │     COL_21   │   Columna    │   Base   │  Intermedia O   │
│ 22  │ -2750.0  │  5500.0  │  -1.8 (calc)   │     C22    │     COL_22   │   Columna    │   Base   │  Intermedia S   │
│ 23  │  2750.0  │  5500.0  │  -1.8 (calc)   │     C23    │     COL_23   │   Columna    │   Base   │  Intermedia S   │
│ 24  │  7513.1  │  2750.0  │  -1.9 (calc)   │     C24    │     COL_24   │   Columna    │   Base   │  Intermedia E   │
│ 25  │ 10263.1  │ -2013.1  │  -2.1 (calc)   │     C25    │     COL_25   │   Columna    │   Base   │  Intermedia E   │
│ 26  │ 13013.2  │ -6776.3  │  -2.2 (calc)   │     C26    │     COL_26   │   Columna    │   Base   │  Intermedia E   │
│ 27  │   0.0    │ -4763.1  │  -1.9 (calc)   │     C27    │     COL_27   │   Columna    │   Base   │  Centro         │
└─────┴──────────┴──────────┴────────────────┴────────────┴──────────────┴──────────────┴──────────┴─────────────────┘
```

### Código de Colores
- **Blanco:** Propiedades fijas (ETABS)
- **Amarillo:** Asentamiento calculado (VLOOKUP desde INPUT)
- **Azul:** Propiedades descriptivas (manual)

### Fórmula en Celda D5 (Ejemplo Empotrado)

```excel
=VLOOKUP(A5, INPUT!$A$5:$D$31, 2, FALSE)
```

**Interpretación:**
- Busca valor de A5 (que es 1)
- En tabla INPUT filas 5–31
- Retorna columna 2 (B = Empotrado)
- Busca exacta
- Resultado: -2.1 mm (de INPUT!B5)

---

## TABLA 2.0: TABLA DE CONEXIONES - VISTA COMPLETA

### Ubicación
- Hoja: `[MODELO]_2_Conexiones`
- Rango: A1:H31 (encabezados + 30 conexiones)
- Actualización: Automática (VLOOKUP desde Tabla 1.0, cálculo de distancia)

### Estructura Visual

```
┌────────┬────────┬──────────┬──────────┬──────────┬──────────┬──────────────┬────────────────┐
│   A    │   B    │    C     │    D     │    E     │    F     │      G       │       H        │
├────────┼────────┼──────────┼──────────┼──────────┼──────────┼──────────────┼────────────────┤
│ Nodo i │ Nodo j │  Xi (mm) │  Yi (mm) │ Xj (mm)  │ Yj (mm)  │  Dist (mm)   │ Observaciones  │
├────────┼────────┼──────────┼──────────┼──────────┼──────────┼──────────────┼────────────────┤
│   1    │   3    │-14600.9  │ -9526.3  │-17776.3  │ -4026.3  │ 6350.9 (calc)│ Viga perimetral│
│   1    │   6    │-14600.9  │ -9526.3  │ -8250.0  │ -9526.3  │ 6350.9 (calc)│ Viga intermedia│
│   1    │  19    │-14600.9  │ -9526.3  │-13013.2  │ -6776.3  │ 2886.5 (calc)│ Viga interior  │
│   2    │  11    │ 14600.9  │ -9526.3  │  8250.0  │ -9526.3  │ 6350.9 (calc)│ Viga perimetral│
│   2    │  12    │ 14600.9  │ -9526.3  │ 17776.3  │ -4026.3  │ 6350.9 (calc)│ Viga exterior  │
│   2    │  26    │ 14600.9  │ -9526.3  │ 13013.2  │ -6776.3  │ 2886.5 (calc)│ Viga interior  │
│   3    │   4    │-17776.3  │ -4026.3  │-15026.3  │  736.9   │ 5238.4 (calc)│ Viga vertical  │
│   3    │  19    │-17776.3  │ -4026.3  │-13013.2  │ -6776.3  │ 5910.7 (calc)│ Viga radial    │
│   4    │   5    │-15026.3  │  736.9   │-12276.3  │  5500.0  │ 4895.3 (calc)│ Viga interior  │
│   4    │  20    │-15026.3  │  736.9   │-10263.1  │ -2013.1  │ 5566.0 (calc)│ Viga diagonal  │
│   5    │  17    │-12276.3  │  5500.0  │ -7513.1  │  8250.0  │ 5513.3 (calc)│ Viga lateral   │
│   5    │  21    │-12276.3  │  5500.0  │ -7513.1  │  2750.0  │ 4886.1 (calc)│ Viga interior  │
│   6    │   7    │ -8250.0  │ -9526.3  │ -5500.0  │ -4763.1  │ 5595.5 (calc)│ Viga central   │
│   6    │  19    │ -8250.0  │ -9526.3  │-13013.2  │ -6776.3  │ 5434.5 (calc)│ Viga mixta     │
│   7    │   8    │ -5500.0  │ -4763.1  │ -2750.0  │   0.0    │ 5514.6 (calc)│ Viga central   │
│   7    │  20    │ -5500.0  │ -4763.1  │-10263.1  │ -2013.1  │ 5486.7 (calc)│ Viga radial    │
│   8    │   9    │ -2750.0  │   0.0    │  2750.0  │   0.0    │ 5500.0 (calc)│ Viga central   │
│   8    │  22    │ -2750.0  │   0.0    │ -2750.0  │  5500.0  │ 5500.0 (calc)│ Viga vertical  │
│   9    │  10    │  2750.0  │   0.0    │  5500.0  │ -4763.1  │ 5514.6 (calc)│ Viga central   │
│   9    │  23    │  2750.0  │   0.0    │  2750.0  │  5500.0  │ 5500.0 (calc)│ Viga vertical  │
│  10    │  11    │  5500.0  │ -4763.1  │  8250.0  │ -9526.3  │ 5595.5 (calc)│ Viga central   │
│  10    │  25    │  5500.0  │ -4763.1  │ 10263.1  │ -2013.1  │ 5486.7 (calc)│ Viga radial    │
│  11    │  12    │  8250.0  │ -9526.3  │ 17776.3  │ -4026.3  │ 6350.9 (calc)│ Viga perimetral│
│  12    │  13    │ 17776.3  │ -4026.3  │ 15026.3  │  736.9   │ 5238.4 (calc)│ Viga vertical  │
│  13    │  14    │ 15026.3  │  736.9   │ 12276.3  │  5500.0  │ 4895.3 (calc)│ Viga interior  │
│  13    │  26    │ 15026.3  │  736.9   │ 13013.2  │ -6776.3  │ 7667.7 (calc)│ Viga diagonal  │
│  14    │  15    │ 12276.3  │  5500.0  │  2750.0  │ 11000.0  │11180.6 (calc)│ Viga larga     │
│  15    │  16    │  2750.0  │ 11000.0  │ -2750.0  │ 11000.0  │ 5500.0 (calc)│ Viga horizontal│
│  15    │  18    │  2750.0  │ 11000.0  │  7513.1  │  8250.0  │ 5913.6 (calc)│ Viga diagonal  │
│  16    │  17    │ -2750.0  │ 11000.0  │ -7513.1  │  8250.0  │ 5913.6 (calc)│ Viga diagonal  │
└────────┴────────┴──────────┴──────────┴──────────┴──────────┴──────────────┴────────────────┘
```

### Fórmulas de Coordenadas

**Celda C5 (Xi - Nodo i X):**
```excel
=VLOOKUP(A5, EMPO_1_Nodos!$A$5:$B$31, 2, FALSE)
```
- Busca nodo 1 en Tabla 1.0
- Retorna columna B (Coord X)
- Resultado: -14600.9

**Celda D5 (Yi - Nodo i Y):**
```excel
=VLOOKUP(A5, EMPO_1_Nodos!$A$5:$C$31, 3, FALSE)
```
- Busca nodo 1 en Tabla 1.0
- Retorna columna C (Coord Y)
- Resultado: -9526.3

**Celda E5 (Xj - Nodo j X):**
```excel
=VLOOKUP(B5, EMPO_1_Nodos!$A$5:$B$31, 2, FALSE)
```
- Busca nodo 3 (de B5) en Tabla 1.0
- Retorna columna B (Coord X)
- Resultado: -17776.3

**Celda F5 (Yj - Nodo j Y):**
```excel
=VLOOKUP(B5, EMPO_1_Nodos!$A$5:$C$31, 3, FALSE)
```
- Busca nodo 3 en Tabla 1.0
- Retorna columna C (Coord Y)
- Resultado: -4026.3

### Fórmula de Distancia

**Celda G5 (Distancia 1-3):**
```excel
=SQRT((E5-C5)^2 + (F5-D5)^2)
```

**Paso a paso:**
```
E5 - C5 = -17776.3 - (-14600.9) = -3175.4
F5 - D5 = -4026.3 - (-9526.3) = 5500.0

(-3175.4)² = 10,083,785
(5500.0)² = 30,250,000

Suma = 40,333,785

√40,333,785 = 6350.9 mm
```

---

## TABLA 3.0: COMPARACIÓN - VISTA COMPLETA

### Ubicación
- Hoja: `[MODELO]_3_Comparacion`
- Estructura: Agrupaciones por nodo i + Análisis Global + Gráficos
- Actualización: Automática (VLOOKUP + cálculos)

### SECCIÓN A: AGRUPACIONES POR NODO

#### Agrupación Nodo 1

```
┌─────────────────────────────────────────────────────────────────────────┐
│ AGRUPACIÓN NODO 1 (Nodo Central SO)                                    │
├────────┬─────────────┬────────┬─────────────┬──────────┬─────────────┤
│ Nodo i │ Asent i(mm) │ Nodo j │ Asent j(mm) │ Δ (i-j)  │    % Δ      │
├────────┼─────────────┼────────┼─────────────┼──────────┼─────────────┤
│   1    │    -2.1     │   3    │    -2.3     │   0.2    │   9.52%     │ ← Δ positivo
│   1    │    -2.1     │   6    │    -2.0     │  -0.1    │  -4.76%     │ ← Δ negativo
│   1    │    -2.1     │  19    │    -2.2     │   0.1    │   4.76%     │
├────────┼─────────────┼────────┼─────────────┼──────────┼─────────────┤
│  STAT  │ Mín: -2.3   │        │ Máx: -2.0   │ 0.2 máx  │   9.52% máx │
└────────┴─────────────┴────────┴─────────────┴──────────┴─────────────┘
```

**Interpretación:**
- Nodo 1 conecta con nodos 3, 6, 19
- Asentamiento de nodo 1: -2.1 mm
- Rango de conectados: -2.3 (máximo hundimiento) a -2.0 (mínimo)
- Mayor diferencia: 0.2 mm (vs nodo 3)
- % Δ máximo: 9.52% (respecto a nodo 1)

#### Agrupación Nodo 2

```
┌─────────────────────────────────────────────────────────────────────────┐
│ AGRUPACIÓN NODO 2 (Nodo Esquina SE)                                    │
├────────┬─────────────┬────────┬─────────────┬──────────┬─────────────┤
│ Nodo i │ Asent i(mm) │ Nodo j │ Asent j(mm) │ Δ (i-j)  │    % Δ      │
├────────┼─────────────┼────────┼─────────────┼──────────┼─────────────┤
│   2    │    -2.0     │  11    │    -2.0     │   0.0    │   0.00%     │ ← Asentamiento igual
│   2    │    -2.0     │  12    │    -2.3     │   0.3    │  15.00%     │ ← Mayor diferencia
│   2    │    -2.0     │  26    │    -2.2     │   0.2    │  10.00%     │
├────────┼─────────────┼────────┼─────────────┼──────────┼─────────────┤
│  STAT  │ Mín: -2.3   │        │ Máx: -2.0   │ 0.3 máx  │   15.00%máx │
└────────┴─────────────┴────────┴─────────────┴──────────┴─────────────┘
```

**Interpretación:**
- Nodo 2 conecta con nodos 11, 12, 26
- Mayor diferencia local: 0.3 mm (vs nodo 12)
- % crítico: 15% (mayor que en nodo 1)

#### [Agrupaciones 3 a 27 siguen mismo patrón]

### SECCIÓN B: ANÁLISIS GLOBAL

#### Tabla Global - Asentamientos Ordenados

```
┌────────┬──────────────────┬─────────────────────┬──────────────────┐
│ Nodo   │ Asent Global(mm) │ Δ vs Mín Global(mm) │ % Δ vs Mín Global│
├────────┼──────────────────┼─────────────────────┼──────────────────┤
│  15    │       -1.4       │        0.0          │      0.0%        │ ← Mínimo global
│  16    │       -1.4       │        0.0          │      0.0%        │
│  18    │       -1.5       │       -0.1          │     -7.1%        │
│  17    │       -1.5       │       -0.1          │     -7.1%        │
│   9    │       -1.8       │       -0.4          │    -28.6%        │
│   8    │       -1.8       │       -0.4          │    -28.6%        │
│  10    │       -1.9       │       -0.5          │    -35.7%        │
│   7    │       -1.9       │       -0.5          │    -35.7%        │
│  27    │       -1.9       │       -0.5          │    -35.7%        │
│  11    │       -2.0       │       -0.6          │    -42.9%        │
│   6    │       -2.0       │       -0.6          │    -42.9%        │
│   2    │       -2.0       │       -0.6          │    -42.9%        │
│  24    │       -1.9       │       -0.5          │    -35.7%        │
│  25    │       -2.1       │       -0.7          │    -50.0%        │
│  20    │       -2.1       │       -0.7          │    -50.0%        │
│   1    │       -2.1       │       -0.7          │    -50.0%        │
│   5    │       -2.1       │       -0.7          │    -50.0%        │
│  19    │       -2.2       │       -0.8          │    -57.1%        │
│   4    │       -2.2       │       -0.8          │    -57.1%        │
│  26    │       -2.2       │       -0.8          │    -57.1%        │
│  13    │       -2.2       │       -0.8          │    -57.1%        │
│  12    │       -2.3       │       -0.9          │    -64.3%        │
│   3    │       -2.3       │       -0.9          │    -64.3%        │ ← Máximo hundimiento
│  14    │       -2.1       │       -0.7          │    -50.0%        │
│  21    │       -1.9       │       -0.5          │    -35.7%        │
│  22    │       -1.8       │       -0.4          │    -28.6%        │
│  23    │       -1.8       │       -0.4          │    -28.6%        │
└────────┴──────────────────┴─────────────────────┴──────────────────┘
```

**Interpretación:**
- Nodos 15 y 16 tienen mínimo asentamiento (-1.4 mm)
- Nodos 3 y 12 tienen máximo asentamiento (-2.3 mm)
- Diferencia global: 0.9 mm
- % cambio máximo: 64.3% (64% más deformación del máximo vs mínimo)

### SECCIÓN C: GRÁFICOS INTEGRADOS

#### Gráfico 1: Barras de Diferencias por Agrupación (Nodo 1)

```
% Δ (Nodo 1)
  │
15│                              ███
  │                              ███
10│         ███                  ███
  │         ███                  ███
 5│         ███        ███        ███
  │         ███        ███        ███
 0├─────────███────────███────────███──────
  │              -4.76%
 -5│
  │
  └────────────────────────────────────→
    Conexión 1-3  Conexión 1-6  Conexión 1-19
    (9.52%)      (-4.76%)      (4.76%)
```

**Propósito:** Mostrar visualmente qué conexiones del nodo 1 tienen mayores diferencias

#### Gráfico 2: Cascada Global de Asentamientos

```
Asentamiento (mm)
     0│
       │
  -0.5│                    ▯▯
       │      ▯▯▯          ▯▯▯
       │      ▯▯▯  ▯▯      ▯▯▯
  -1.0│      ▯▯▯▯▯▯▯      ▯▯▯▯
       │      ▯▯▯▯▯▯▯▯▯    ▯▯▯▯▯
       │      ▯▯▯▯▯▯▯▯▯▯   ▯▯▯▯▯▯
  -1.5│      ▯▯▯▯▯▯▯▯▯▯▯  ▯▯▯▯▯▯▯▯
       │      ▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯
  -2.0│  ▯   ▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯
       │  ▯▯  ▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯
  -2.5│▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯▯
       │
       └───────────────────────────────→
         15  18  9  8  10  7  27 11  6  2  24  25  20  1  5  19  4  26  13  12  3
         (Nodos ordenados menor a mayor asentamiento)

       ▯ = 1 nodo
       ▯▯ = 2 nodos agrupados
```

**Propósito:** Ver distribución global; identificar concentraciones de deformación

#### Gráfico 3: Máximas Diferencias Locales por Nodo

```
% Δ Máximo por Nodo
 70%│                                      ◆
    │                                     ◆
 60%│                              ◆ ◆  ◆
    │                         ◆ ◆ ◆   ◆
 50%│              ◆      ◆ ◆ ◆   ◆ ◆
    │          ◆ ◆ ◆    ◆ ◆     ◆
 40%│        ◆ ◆ ◆  ◆ ◆ ◆       ◆
    │      ◆ ◆ ◆   ◆ ◆       ◆
 30%│  ◆ ◆ ◆ ◆    ◆ ◆       ◆
    │◆ ◆ ◆        ◆        
 20%│                      ◆
    │                  ◆ ◆   
 10%│              ◆ ◆     
    │          ◆ ◆         
  0%├──────────────────────────→
     1  2  3  4  5  6  7  8  9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27
```

**Propósito:** Identificar "nodos críticos" con mayores diferencias locales

---

## EJEMPLO: SEGUIR UNA CÉLULA EN CASCADA

### Pregunta: ¿De dónde viene el valor en E5 (Δ = 0.2) de Tabla 3.0?

**Trazabilidad completa:**

```
TABLA 3.0, E5: Δ = 0.2 mm
  ↓ Fórmula
=B5 - D5
  ↓ Desglosa
B5 = Asentamiento nodo 1 = -2.1 mm
D5 = Asentamiento nodo 3 = -2.3 mm
  ↓ Cálculo
-2.1 - (-2.3) = -2.1 + 2.3 = 0.2 mm
  ↓ Origen de B5
TABLA 3.0, B5: =VLOOKUP(A5, EMPO_1_Nodos!$A$5:$D$31, 4, FALSE)
  ↓ Busca nodo 1 en Tabla 1.0, retorna asentamiento
EMPO_1_Nodos, D5: =VLOOKUP(A5, INPUT!$A$5:$D$31, 2, FALSE)
  ↓ Busca nodo 1 en INPUT, retorna Empotrado
INPUT, B5: -2.1 mm
  ↓ ENTRADA MANUAL (usuario lo extrae de ETABS)
ETABS Output: Nodo 1, Modelo Empotrado → -2.1 mm

CONCLUSIÓN:
Tabla 3.0!E5 = 0.2 depende de:
  • INPUT!B5 (Empotrado, Nodo 1) = -2.1
  • INPUT!B7 (Empotrado, Nodo 3) = -2.3
  
Si usuario cambia INPUT!B5 a -2.5:
  → Tabla 1.0!D5 → -2.5
  → Tabla 3.0!B5 → -2.5
  → Tabla 3.0!E5 → -2.5 - (-2.3) = -0.2
  → Gráfico 1 se redibuja
  → % Δ en F5 recalcula: -0.2 / ABS(-2.5) * 100 = -8%
  
TODO AUTOMÁTICO EN < 1 SEGUNDO
```

---

## VALIDACIÓN DE INTEGRIDAD: CHECKLIST

### Por Tabla 1.0

- [ ] 27 nodos, IDs 1–27 únicos y sin espacios
- [ ] Columna D (Asentamiento) sin errores #N/A, #REF!
- [ ] Asentamientos Empotrado ≤ Winkler ≤ Barkan
- [ ] Coordenadas idénticas en EMPO_1, WINKLER_1, BARKAN_1
- [ ] Observaciones completas (mínimo 1 carácter)

### Por Tabla 2.0

- [ ] 30 conexiones, sin duplicados
- [ ] Columna G (Distancia) todas > 0, sin excepciones
- [ ] Distancias idénticas en EMPO_2, WINKLER_2, BARKAN_2
- [ ] Coordenadas Xi, Yi, Xj, Yj sin blancos

### Por Tabla 3.0

- [ ] Agrupación Nodo 1: 3 filas de datos + 1 de estadística
- [ ] Agrupación Nodo 2: 3 filas de datos + 1 de estadística
- [ ] Todas agrupaciones completas (una por cada conexión del nodo)
- [ ] Análisis global: 27 nodos, ordenados menor a mayor
- [ ] Columna Δ vs Mín: primer nodo = 0.0
- [ ] Gráficos: 3 gráficos visibles e interactivos

---

**Versión:** 3.0  
**Fecha:** Julio 2025  
**Propósito:** Ejemplos visuales detallados de Tablas 1.0, 2.0, 3.0
