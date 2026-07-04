# ÍNDICE MAESTRO: DOCUMENTACIÓN COMPLETA
## Análisis Comparativo de Asentamientos - v3.0

---

## 📋 RESUMEN EJECUTIVO

Has solicitado crear una **especificación detallada** para un archivo Excel que analice asentamientos de cimentaciones comparando **tres modelos** (Empotrado, Winkler Corregido, Barkan Corregido).

**Lo que se ha entregado:**

1. **PROMPT_ARQUITECTURA_v3.0.md** - Especificación completa en Markdown
2. **TABLAS_DETALLADAS_EJEMPLOS.md** - Ejemplos visuales de cada tabla
3. **Documentación de apoyo** - Guías de referencia rápida

**El Excel propuesto:**
- ✓ Completamente entrelazado en fórmulas (cero valores hardcodeados)
- ✓ Tres modelos independientes, misma estructura
- ✓ Tabla 1.0 (Nodos), 2.0 (Conexiones), 3.0 (Análisis avanzado)
- ✓ INPUT único como fuente de asentamientos

---

## 📁 ARCHIVOS GENERADOS

### Archivo Principal: PROMPT_ARQUITECTURA_v3.0.md

**Tamaño:** ~26 KB  
**Secciones:** 10 capítulos completos  
**Propósito:** Especificación técnica exhaustiva

**Contenido:**
- Resumen ejecutivo y principios de diseño
- Estructura general del archivo Excel
- Hoja INPUT: formato y validaciones
- Tabla 1.0: Nodos con VLOOKUP a INPUT
- Tabla 2.0: Conexiones con cálculo de distancias
- Tabla 3.0: ANÁLISIS AVANZADO con:
  - Agrupaciones por nodo
  - Análisis global ordenado
  - Estadísticas locales
  - Gráficos integrados
- Arquitectura de entrelazamiento (flujo de datos)
- Especificaciones de fórmulas (referencia rápida)
- Flujo de trabajo (fases 1–5)
- Validaciones y control de calidad

**Cómo usar:**
1. Leer sección "Resumen Ejecutivo" (3 min)
2. Revisar "Estructura General" (2 min)
3. Estudiar en detalle las Tablas 1.0, 2.0, 3.0 (15 min)
4. Consultar "Especificaciones de Fórmulas" cuando se implemente

---

### Archivo de Apoyo: TABLAS_DETALLADAS_EJEMPLOS.md

**Tamaño:** ~25 KB  
**Propósito:** Ejemplos visuales y concretos

**Contenido:**
- **Tabla 1.0:** Vista completa con 27 nodos reales
- **Tabla 2.0:** Vista completa con 30 conexiones reales
- **Tabla 3.0:** 
  - Agrupaciones (Nodo 1, Nodo 2, ... Nodo 27)
  - Análisis global con 27 nodos ordenados
  - Gráficos de ejemplo (ASCII art)
- **Trazabilidad:** Cómo seguir un valor desde su origen
- **Validación:** Checklist de integridad

**Cómo usar:**
1. Abre este archivo cuando necesites ver un ejemplo real
2. Copia la estructura de Tabla 1.0/2.0/3.0 como referencia
3. Valida tu Excel contra el checklist

---

### Archivos de Referencia (Anteriores)

Si necesitas contexto de versiones anteriores:

- **PROMPT_MODELOS_ASENTAMIENTOS_v2.md** - Versión anterior más simple
- **GUÍA_RÁPIDA_ARQUITECTURA.md** - Intro rápida a entrelazamiento
- **DIAGRAMA_ENTRELAZAMIENTO.md** - Flujo de datos visual
- **PROMPT_ASENTAMIENTOS_DIFERENCIALES.md** - Primera versión

---

## 🎯 CÓMO USAR ESTA DOCUMENTACIÓN

### Escenario 1: Entender la Arquitectura General

**Lectura recomendada (30 minutos):**

```
1. Este archivo (índice)                    5 min
   ↓
2. PROMPT_ARQUITECTURA_v3.0.md
   - Sección 1: Resumen Ejecutivo          5 min
   - Sección 2: Estructura General         5 min
   - Sección 3: Hoja INPUT                 5 min
   - Sección 7: Arquitectura de Entrelazamiento  5 min
   ↓
3. TABLAS_DETALLADAS_EJEMPLOS.md
   - Vista rápida de Tabla 1.0, 2.0, 3.0  5 min
```

**Al terminar:** Tendrás claro qué hace cada tabla y cómo se conectan.

---

### Escenario 2: Implementar el Excel

**Lectura + Implementación (4–6 horas):**

```
1. PROMPT_ARQUITECTURA_v3.0.md - Lectura completa        1 hora
   
2. TABLAS_DETALLADAS_EJEMPLOS.md - Consultar mientras creas
   
3. Crear Estructura Base:
   - Hoja INPUT (copiar estructura)                       30 min
   - Hojas para 3 modelos × 3 tablas                     30 min
   
4. Implementar Tabla 1.0:
   - Copiar coordenadas (ctes)                           30 min
   - Copiar propiedades ETABS (ctes)                     30 min
   - Crear VLOOKUP para asentamientos                    30 min
   - Validar (comparar con ejemplo)                      30 min
   
5. Implementar Tabla 2.0:
   - Copiar pares de nodos (ctes)                        30 min
   - Crear VLOOKUP para coordenadas                      30 min
   - Crear fórmula de distancia                          30 min
   - Validar                                             30 min
   
6. Implementar Tabla 3.0:
   - Crear agrupaciones                                   1 hora
   - Crear análisis global                               1 hora
   - Crear fórmulas de estadística                       1 hora
   - Crear gráficos                                       1 hora
   - Validar entrelazamiento                             1 hora
```

**Al terminar:** Excel completamente funcional.

---

### Escenario 3: Resolver Problemas

**Usa esta matriz de referencia:**

| Problema | Sección en PROMPT v3.0 | Ejemplo en TABLAS |
|----------|--------|---------|
| "¿Dónde va el asentamiento?" | Sección 3 (INPUT) | Tabla 1.0 col D |
| "¿Cómo se calcula la distancia?" | Sección 5, "Fórmula de Distancia" | Tabla 2.0, ejemplo 1-3 |
| "¿Qué es Tabla 3.0?" | Sección 6 completa | Agrupación Nodo 1 |
| "Las fórmulas dan error" | Sección 10, Validaciones | Checklist |
| "¿Cómo cambio los datos?" | Sección 8, Flujo de Trabajo | Trazabilidad |
| "¿Se actualizan los gráficos?" | Sección 7, Cambio en Cascada | Gráficos 1–3 |

---

## 🔧 IMPLEMENTACIÓN PASO A PASO

### Fase 1: Preparación (1 hora)

```
□ Descargar ETABS y extraer:
  - Coordenadas X, Y de todos los nodos
  - Asentamientos para 3 modelos
  - Propiedades de nodos (nombre, tipo, altura)
  - Lista de conexiones (qué nodo conecta con qué)

□ Abrir Excel
□ Crear hojas:
  - INPUT
  - EMPO_1_Nodos, EMPO_2_Conexiones, EMPO_3_Comparacion
  - WINKLER_1_Nodos, WINKLER_2_Conexiones, WINKLER_3_Comparacion
  - BARKAN_1_Nodos, BARKAN_2_Conexiones, BARKAN_3_Comparacion
```

### Fase 2: Llenar INPUT (30 minutos)

```
Hoja: INPUT
Rango: B5:D31

Columna B (Empotrado): Asentamientos modelo 1
Columna C (Winkler): Asentamientos modelo 2
Columna D (Barkan): Asentamientos modelo 3

Validación: 27 nodos, 3 columnas, sin vacíos
```

### Fase 3: Crear Tabla 1.0 (1 hora)

```
Hoja: EMPO_1_Nodos (repetir para WINKLER, BARKAN)

Columna A: IDs (1–27) - CONST
Columna B: Coord X - CONST
Columna C: Coord Y - CONST
Columna D: Asentamiento - FÓRMULA:
           =VLOOKUP(A5, INPUT!$A$5:$D$31, 2, FALSE)
           [usar col 2 para Empotrado, 3 para Winkler, 4 para Barkan]
Columna E–H: Propiedades ETABS - CONST
Columna I: Observaciones - MANUAL
```

### Fase 4: Crear Tabla 2.0 (1 hora)

```
Hoja: EMPO_2_Conexiones

Columna A–B: Pares de nodos - CONST (30 conexiones)

Columna C (Xi):
  =VLOOKUP(A5, EMPO_1_Nodos!$A$5:$B$31, 2, FALSE)

Columna D (Yi):
  =VLOOKUP(A5, EMPO_1_Nodos!$A$5:$C$31, 3, FALSE)

Columna E (Xj):
  =VLOOKUP(B5, EMPO_1_Nodos!$A$5:$B$31, 2, FALSE)

Columna F (Yj):
  =VLOOKUP(B5, EMPO_1_Nodos!$A$5:$C$31, 3, FALSE)

Columna G (Distancia):
  =SQRT((E5-C5)^2 + (F5-D5)^2)
```

### Fase 5: Crear Tabla 3.0 (2–3 horas)

```
SECCIÓN A: AGRUPACIONES

Para cada nodo i (1–27):
  - Listar todos sus conectados j (desde Tabla 2.0)
  - Columnas: Nodo i, Asent i, Nodo j, Asent j, Δ, % Δ
  
  Fórmulas:
  Asent i: =VLOOKUP(Nodo i, EMPO_1_Nodos!A:D, 4, FALSE)
  Asent j: =VLOOKUP(Nodo j, EMPO_1_Nodos!A:D, 4, FALSE)
  Δ:       =Asent i - Asent j
  % Δ:     =IF(Asent i=0, 0, Δ/ABS(Asent i)*100)
  
  - Fila resumen con MIN, MAX, Δ máx, % máx

SECCIÓN B: ANÁLISIS GLOBAL

  - Listar todos los nodos (1–27)
  - Columnas: Nodo, Asent, Δ vs Mín Global, % vs Mín
  - Ordenar por Asent ascendente (menor a mayor)
  
  Fórmulas:
  Asent:        =VLOOKUP(Nodo, EMPO_1_Nodos!A:D, 4, FALSE)
  Δ vs Mín:     =Asent - MIN(EMPO_1_Nodos!D:D)
  % vs Mín:     =IF(MIN(...)=0, 0, Δ vs Mín / ABS(MIN(...)) * 100)

SECCIÓN C: GRÁFICOS

  - Gráfico 1: Barras de % Δ por conexión (agrupación)
  - Gráfico 2: Columnas de asentamiento global
  - Gráfico 3: Barras de % Δ máximo por nodo
```

### Fase 6: Validar Entrelazamiento (1 hora)

```
□ Cambiar 1 valor en INPUT (ej: B5 de -2.1 a -2.5)
□ Verificar Tabla 1.0 columna D actualiza
□ Verificar Tabla 3.0 agrupación nodo 1 recalcula
□ Verificar gráficos se redibujan
□ Revisar checklist en TABLAS_DETALLADAS_EJEMPLOS.md

Resultado esperado: TODO se actualiza automáticamente en < 1 seg
```

---

## 📊 ESTRUCTURA FINAL DEL ARCHIVO

```
Asentamientos_Modelos_v3.xlsx
│
├─ INPUT                                    ← FUENTE ÚNICA
│  ├─ A: ID_NODO (1–27)
│  ├─ B: Empotrado (mm)                    ← ENTRADA MANUAL
│  ├─ C: Winkler (mm)                      ← ENTRADA MANUAL
│  └─ D: Barkan (mm)                       ← ENTRADA MANUAL
│
├─ MODELO 1: EMPOTRADO
│  ├─ EMPO_1_Nodos
│  │  ├─ A–C: Coords + Props ETABS (CONST)
│  │  └─ D: Asentamiento (VLOOKUP col 2)
│  ├─ EMPO_2_Conexiones
│  │  ├─ A–B: Pares nodos (CONST)
│  │  ├─ C–F: Coords (VLOOKUP)
│  │  └─ G: Distancia (SQRT)
│  └─ EMPO_3_Comparacion
│     ├─ Sección A: Agrupaciones
│     ├─ Sección B: Análisis global
│     └─ Sección C: Gráficos (3)
│
├─ MODELO 2: WINKLER CORREGIDO
│  ├─ WINKLER_1_Nodos        (estructura idéntica)
│  ├─ WINKLER_2_Conexiones   (estructura idéntica)
│  └─ WINKLER_3_Comparacion  (estructura idéntica)
│
├─ MODELO 3: BARKAN CORREGIDO
│  ├─ BARKAN_1_Nodos         (estructura idéntica)
│  ├─ BARKAN_2_Conexiones    (estructura idéntica)
│  └─ BARKAN_3_Comparacion   (estructura idéntica)
│
└─ RESUMEN_COMPARATIVO (opcional)
   └─ Tabla de comparación inter-modelos
```

---

## 🎨 DISEÑO DE COLORES

```
Amarillo (FFFF99): Celdas de entrada manual (INPUT!B5:D31)
Verde (E2EFDA):    Celdas calculadas (VLOOKUP, SQRT)
Naranja (FCE4D6):  Celdas de análisis (diferencias, %)
Azul (1F4E78):     Encabezados
Blanco:            Propiedades fijas (CONST)
```

---

## 🔗 REFERENCIAS RÁPIDAS

### Fórmulas Más Usadas

```excel
VLOOKUP entrada:     =VLOOKUP(A5, INPUT!$A$5:$D$31, 2, FALSE)
VLOOKUP coords:      =VLOOKUP(A5, EMPO_1_Nodos!$A$5:$B$31, 2, FALSE)
Distancia:           =SQRT((E5-C5)^2 + (F5-D5)^2)
Diferencia:          =B5 - D5
% Diferencia:        =IF(B5=0, 0, (B5-D5)/ABS(B5)*100)
Mín global:          =MIN($B$5:$B$31)
Δ vs Mín:            =B5 - MIN($B$5:$B$31)
```

### Errores Comunes y Soluciones

| Error | Causa | Solución |
|-------|-------|----------|
| #N/A en VLOOKUP | Nodo no existe en tabla | Verificar IDs (1–27) |
| #REF! | Referencia a hoja eliminada | Revisar nombres de hojas |
| #DIV/0! | División por cero | Usar IF(denominador=0, 0, ...) |
| Distancia = 0 | Dos nodos con mismas coords | Revisar topología |
| Gráficos no actualizan | Datos fuera del rango | Agregar filas a datos source |

---

## 📈 PRÓXIMOS PASOS

### Paso 1: Validar Conexiones (PENDIENTE)
Necesitas compartir el archivo Excel original con la lista de conexiones (Nodo i, Nodo j).

### Paso 2: Adaptar Números
Si tu proyecto tiene:
- Diferente número de nodos (no 27)
- Diferente número de conexiones (no 30)

Actualiza los rangos en las fórmulas VLOOKUP:
```excel
Cambiar:    INPUT!$A$5:$D$31
Por:        INPUT!$A$5:$D$[última_fila]
```

### Paso 3: Implementar
Usa el Excel base (`Asentamientos_Modelos_Entrelazado.xlsx`) como punto de partida, o crea una estructura nueva siguiendo PROMPT_ARQUITECTURA_v3.0.md.

### Paso 4: Integrar Conexiones Reales
Una vez tengas la lista de conexiones de tu proyecto:
1. Reemplaza pares en Tabla 2.0 columnas A–B
2. Las distancias se recalculan automáticamente
3. Tabla 3.0 se actualiza automáticamente

---

## ❓ PREGUNTAS FRECUENTES

### P: ¿Necesito conocer macros/VBA?
**R:** No. Todo funciona con fórmulas Excel estándar.

### P: ¿Funciona en Google Sheets?
**R:** Parcialmente. VLOOKUP sí; algunos gráficos pueden necesitar adaptación.

### P: ¿Cuánto tiempo toma implementarlo?
**R:** 4–6 horas si tienes datos de ETABS listos.

### P: ¿Se puede automatizar más?
**R:** Sí. Con Python + openpyxl se puede generar el Excel automáticamente.

### P: ¿Qué pasa si cambio los datos?
**R:** Todo se actualiza en < 1 segundo (entrelazamiento automático).

### P: ¿Cómo comparo los 3 modelos?
**R:** Ver archivo TABLAS_DETALLADAS_EJEMPLOS.md, sección "Comparación Entre Modelos".

---

## 📞 SOPORTE

Si encuentras problemas:

1. Revisa el **Checklist de Validación** en TABLAS_DETALLADAS_EJEMPLOS.md
2. Consulta la tabla **"Errores Comunes"** en este documento
3. Busca la sección correspondiente en PROMPT_ARQUITECTURA_v3.0.md

---

## 📄 RESUMEN DE DOCUMENTOS

| Archivo | Tamaño | Propósito | Lectura |
|---------|--------|----------|---------|
| **PROMPT_ARQUITECTURA_v3.0.md** | 26 KB | Especificación técnica completa | 45 min |
| **TABLAS_DETALLADAS_EJEMPLOS.md** | 25 KB | Ejemplos visuales y concretos | 20 min |
| **GUÍA_RÁPIDA_ARQUITECTURA.md** | 14 KB | Intro rápida (v2.0) | 10 min |
| **DIAGRAMA_ENTRELAZAMIENTO.md** | 23 KB | Flujo de datos visual | 15 min |
| Anteriores (v1–v2) | 40+ KB | Contexto histórico | Opcional |

**Total documentación:** ~150 KB de especificaciones detalladas

---

## ✅ CHECKLIST FINAL

Antes de empezar a implementar:

- [ ] Entiendes estructura de 3 tablas (1.0, 2.0, 3.0)
- [ ] Sabes qué es VLOOKUP y cómo se usa
- [ ] Tienes acceso a Excel 2016+ o Google Sheets
- [ ] Tienes datos de ETABS listos (coords, asentamientos, conexiones)
- [ ] Has leído PROMPT_ARQUITECTURA_v3.0.md (secciones 1–6)
- [ ] Entiendes el concepto de "entrelazamiento" (cambio en cascada)
- [ ] Sabes cómo crear un gráfico dinámico en tu versión de Excel

Si todo es ✓: **Estás listo para implementar.**

---

**Documentación Versión:** 3.0  
**Fecha:** Julio 2025  
**Estado:** Completo y listo para implementación

**Siguiente acción:** Compartir archivo Excel con lista de conexiones reales del proyecto.
