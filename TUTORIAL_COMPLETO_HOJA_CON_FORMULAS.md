# Tutorial: Generador de Horarios MATCOM con `hoja_con_formulas.py`

> **Guía completa para generar horarios universitarios en Excel de forma automatizada**

---

## Tabla de Contenidos

1. [¿Qué es un horario?](#1-qué-es-un-horario)
2. [Elementos que intervienen](#2-elementos-que-intervienen)
3. [Cómo los representamos en Excel](#3-cómo-los-representamos-en-excel)
4. [Arquitectura del sistema](#4-arquitectura-del-sistema)
5. [Instalación y dependencias](#5-instalación-y-dependencias)
6. [Uso rápido: tu primer horario en 3 minutos](#6-uso-rápido-tu-primer-horario-en-3-minutos)
7. [API de `hoja_con_formulas.py`](#7-api-de-hoja_con_formulaspy)
8. [Uso con el tutorial (`tutorial_hoja_con_formulas.py`)](#8-uso-con-el-tutorial-tutorial_hoja_con_formulaspy)
9. [Generación desde Lisp (pipeline completo)](#9-generación-desde-lisp-pipeline-completo)
10. [Fórmulas y formato condicional](#10-fórmulas-y-formato-condicional)
11. [La hoja "Aulas"](#11-la-hoja-aulas)
12. [Ejemplos prácticos](#12-ejemplos-prácticos)
13. [Preguntas frecuentes](#13-preguntas-frecuentes)

---

## 1. ¿Qué es un horario?

En el contexto de la **Facultad de Matemática y Computación (MATCOM)** de la Universidad de la Habana, un **horario** es una asignación de **asignaturas** a **aulas** en **turnos** y **días** específicos para cada **grupo** de estudiantes.

Matemáticamente, un horario es una función:

```
horario : Grupo × Día × Turno → Asignatura × Aula
```

Es decir, para cada grupo, en cada día y en cada turno, sabemos qué asignatura se imparte y en qué aula.

El objetivo de este sistema es **generar automáticamente** archivos Excel con horarios completos, validados y visualmente organizados para los **17 grupos** de la facultad.

---

## 2. Elementos que intervienen

En el caso de MATCOM, nos interesan tres elementos fundamentales:

### 2.1. Grupos

Un **grupo** es un cohort de estudiantes que cursan juntos un conjunto de asignaturas. Cada grupo tiene un identificador con formato `{Letra}{Año}{Número}`:

| Carrera | Letra | Grupos |
|---------|-------|--------|
| **Ciencia de Datos** | D | D111, D211, D311, D411 |
| **Ciencia de la Computación** | C | C111, C121, C122, C211, C212, C311, C312, C411, C412 |
| **Matemática** | M | M111, M211, M311, M411 |

Por ejemplo, **D111** = Data Science, Año 1, Grupo 1.

### 2.2. Asignaturas

Cada grupo tiene un conjunto de **asignaturas** que cursa durante el semestre. Cada asignatura se caracteriza por:

| Campo | Descripción | Ejemplo |
|-------|-------------|---------|
| **Abreviatura** | Código corto | `"AL"`, `"L"`, `"F"` |
| **Nombre completo** | Nombre oficial | `"Álgebra Lineal"`, `"Lógica"` |
| **Frecuencia** | Sesiones semanales | `3` (Álgebra tiene 3 clases/semana) |

Ejemplo de asignaturas del grupo D111:

```python
asignaturas_d111 = [
    ("AL", "Álgebra Lineal", 3),
    ("L", "Lógica", 2),
    ("F", "Fundamentos de Programación", 3),
    ("IP", "Introducción a la Programación", 2),
    ("AM I", "Análisis Matemático I", 2),
    ("ICD", "Introducción a la Ciencia de Datos", 2),
]
```

### 2.3. Aulas

Las **aulas** son los espacios físicos donde se imparten las clases. En MATCOM disponemos de:

- **Aula 1** a **Aula 9** — aulas teóricas
- **Lab** — laboratorio de computación

---

## 3. Cómo los representamos en Excel

El convenio de representación es el siguiente:

### 3.1. Una hoja por grupo

Cada grupo tiene su **propia hoja** en el libro Excel. El nombre de la hoja es el identificador del grupo: `"D111"`, `"C111"`, `"M211"`, etc.

### 3.2. Estructura de cada hoja de grupo

Cada hoja contiene **tres secciones**:

#### Sección A: Tabla de Horario (columnas B–G)

Una matriz **Días × Turnos** donde cada celda contiene la **abreviatura** de la asignatura y la fila inferior el **aula**:

```
┌──────┬─────────┬─────────┬───────────┬─────────┬─────────┐
│      │ Lunes   │ Martes  │ Miércoles │ Jueves  │ Viernes │
├──────┼─────────┼─────────┼───────────┼─────────┼─────────┤
│T1    │   AL    │   AM I  │    AL     │   F     │   ICD   │
│      │ Aula 8  │ Aula 7  │  Aula 7*  │ Aula 1  │ Aula 11 │
├──────┼─────────┼─────────┼───────────┼─────────┼─────────┤
│T2    │   L     │   AL    │    EF     │   L     │         │
│      │ Aula 6* │ Aula 6* │  SEDER    │ Aula 6* │         │
└──────┴─────────┴─────────┴───────────┴─────────┴─────────┘
```

Cada turno ocupa **3 filas** (asignatura, aula, separador). Hay **6 turnos**, lo que da 18 filas de horario + 1 fila separadora entre el Turno 3 y el Turno 4.

#### Sección B: Tabla de Asignaturas (columnas I–M)

Una lista de todas las asignaturas del grupo con sus métricas:

```
┌──────┬──────────────────────────┬──────┬────────┬───────────┐
│Abrev │ Asignatura               │ Frec │ Faltan │ Asignadas │
├──────┼──────────────────────────┼──────┼────────┼───────────┤
│ AL   │ Álgebra Lineal           │  3   │   0    │     3     │
│ L    │ Lógica                   │  2   │   0    │     2     │
│ F    │ Fund. de Programación    │  3   │   0    │     3     │
│ ...  │ ...                      │ ...  │  ...   │    ...    │
├──────┼──────────────────────────┼──────┼────────┼───────────┤
│      │                          │      │ Total: │     6     │
│      │                          │      │ Σ Frec:│    15     │
└──────┴──────────────────────────┴──────┴────────┴───────────┘
```

- **Frec** = sesiones semanales planificadas
- **Asignadas** = sesiones ya ubicadas en el horario (fórmula `COUNTIF`)
- **Faltan** = sesiones pendientes (fórmula `Frec - Asignadas`)

#### Sección C: Lista de Aulas (columna O)

Las aulas que utiliza este grupo:

```
┌───────────┐
│  Aulas    │
├───────────┤
│ Aula 1    │
│ Aula 6*   │
│ Aula 7    │
│ Aula 8    │
│ ...       │
├───────────┤
│  Total: 5 │
└───────────┘
```

### 3.3. Hoja "Aulas" (vista global)

Además de las hojas de grupo, hay una hoja maestra llamada **"Aulas"** que muestra, para cada aula y cada turno de cada día, **qué grupos** están asignados:

```
┌──────┬────────┬────────┬────────┬────────┬────────┬────────┬─────┐
│      │ Aula 1 │ Aula 2 │ Aula 3 │ Aula 4 │ Aula 5 │ Aula 6 │ ... │
├──────┼────────┼────────┼────────┼────────┼────────┼────────┼─────┤
│Lunes │        │        │        │        │        │        │     │
│ 1ro  │ D111   │        │        │        │        │ C111   │     │
│ 2do  │        │        │        │        │        │ D111   │     │
│ 3ro  │        │        │        │        │        │        │     │
├──────┼────────┼────────┼────────┼────────┼────────┼────────┼─────┤
│Mar   │        │        │        │        │        │        │     │
│ 1ro  │        │ C112   │        │        │        │ D111   │     │
└──────┴────────┴────────┴────────┴────────┴────────┴────────┴─────┘
```

Cada celda contiene una **fórmula cruzada** que lee las 17 hojas de grupo y concatena los nombres de los grupos que están en esa aula en ese turno.

---

## 4. Arquitectura del sistema

El sistema tiene **3 pipelines** que pueden usarse de forma independiente:

### Pipeline A: Producción principal (desde Lisp)

```
horario.json → convertir_json_a_lisp.py → variables_horario.lisp
     ↓
SBCL (replicar_propuesta_ods.lisp) → generar_propuesta_desde_lisp.py
     ↓
python3 generar_propuesta_desde_lisp.py → propuesta_horarios_desde_lisp.xlsx
```

Este pipeline genera el Excel **completo** con las 18 hojas (17 grupos + Aulas) a partir de datos JSON.

### Pipeline B: Tutorial / Uso directo desde Python

```
tutorial_hoja_con_formulas.py → hoja_con_formulas.py → tutorial_fase1_grupos_vacios.xlsx
                              → hoja_con_formulas.py → tutorial_fase2_con_asignaturas.xlsx
```

Este pipeline es el **más sencillo** para empezar. No necesita Lisp ni JSON. Solo Python.

### Pipeline C: Sistema integrado con fórmulas dinámicas

```
variables_horario.lisp → SBCL (sistema_horarios_integrado.lisp) → ejecutor_aulas.py
     ↓
python3 ejecutor_aulas.py → Aulas_Con_Formulas_Fernando.xlsx
```

Genera la hoja **Aulas** con fórmulas dinámicas que se adaptan automáticamente a los grupos encontrados.

---

## 5. Instalación y dependencias

### Requisitos mínimos

| Dependencia | Versión | Para qué |
|-------------|---------|----------|
| Python | 3.8+ | Motor de generación de Excel |
| openpyxl | 3.0+ | Crear archivos `.xlsx` con formato |
| SBCL | 2.1.11+ | (Opcional) Para el pipeline Lisp |
| lxml | 4.9+ | (Opcional) Para extracción de fórmulas ODS |

### Instalación

```bash
# Solo Python (suficiente para el tutorial y Pipeline B)
pip install openpyxl

# Con Lisp (para Pipelines A y C)
# En Ubuntu/Debian:
sudo apt install sbcl

# Verificar
python3 -c "import openpyxl; print(openpyxl.__version__)"
sbcl --version
```

---

## 6. Uso rápido: tu primer horario en 3 minutos

> 💡 **Esta es la forma más fácil de empezar.** No necesita Lisp ni JSON.

### Paso 1: Ejecutar el tutorial

```bash
cd /home/jose/Proyectos/Scripting/python_excel_tesis
python3 tutorial_hoja_con_formulas.py
```

Esto genera dos archivos:

| Archivo | Contenido |
|---------|-----------|
| `tutorial_fase1_grupos_vacios.xlsx` | 3 grupos (D111, D211, C111) con tablas **vacías** pero con fórmulas y formato condicional |
| `tutorial_fase2_con_asignaturas.xlsx` | Los mismos 3 grupos **con asignaturas** rellenadas |

### Paso 2: Abrir el Excel

Abre `tutorial_fase2_con_asignaturas.xlsx` en Excel o LibreOffice Calc y verás:

- **Hoja D111**: con 4 asignaturas (Álgebra, Lógica, Programación, Análisis)
- **Hoja D211**: con 4 asignaturas (Matemática, Probabilidades, Bases de Datos, Estructura de Datos)
- **Hoja C111**: con 5 asignaturas (Álgebra I, Lógica, Programación, Análisis, Filosofía)
- **Hoja Aulas**: con fórmulas cruzadas que leen las 3 hojas de grupo

### Paso 3: Probar cambiando los grupos

Edita `tutorial_hoja_con_formulas.py` y modifica la lista de grupos:

```python
grupos = ["M111", "M211"]  # Solo grupos de Matemática
```

Y ejecuta de nuevo:

```bash
python3 tutorial_hoja_con_formulas.py
```

¡Listo! Ya tienes tu primer horario generado. 🎉

---

## 7. API de `hoja_con_formulas.py`

`hoja_con_formulas.py` es el **motor de Excel** del sistema. Es una librería Python que genera archivos `.xlsx` a partir de una configuración declarativa en diccionarios.

### 7.1. Función principal

```python
from hoja_con_formulas import generar_excel_personalizado

generar_excel_personalizado(config, "salida.xlsx")
```

- **`config`** (`dict`): Configuración declarativa del libro Excel.
- **`"salida.xlsx"`** (`str`): Ruta del archivo de salida.

### 7.2. Estructura de `config`

```python
config = {
    "sheets": [
        sheet_config_1,
        sheet_config_2,
        ...
    ]
}
```

Cada `sheet_config` es un diccionario que describe **una hoja** del Excel.

### 7.3. Configuración mínima de una hoja

```python
sheet_config = {
    "title": "MiHoja",           # Nombre de la hoja
    "data": [                    # Filas de datos
        ["", "A", "B", "C"],
        ["Fila 1", 1, 2, 3],
        ["Fila 2", 4, 5, 6],
    ],
}
```

### 7.4. Configuración completa de una hoja

```python
sheet_config = {
    # ── Identificación ──
    "title": "D111",

    # ── Datos ──
    "data": [
        ["Grupo", "D111"],
        ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes"],
        ["Turno 1", "AL", "AM I", "AL", "F", "ICD"],
        ["", "Aula 8", "Aula 7", "Aula 7*", "Aula 1", "Aula 11"],
        ...
    ],

    # ── Fórmulas simples (por coordenada) ──
    "formulas": [
        {"row": 5, "col": 13, "value": "=COUNTIF($C$4:$G$15, I5)"},
        {"row": 5, "col": 12, "value": "=K5-M5"},
        {"row": 12, "col": 12, "value": "Total:"},
        {"row": 12, "col": 13, "value": "=COUNTA(I4:I11)"},
    ],

    # ── Fórmulas complejas (cross-sheet) ──
    "fernando_formulas": [
        {
            "cell": "C4",
            "formula": '=SUBSTITUTE(TRIM(CONCAT('
                'IF(D111!$C$5=C$3,D111!$B$1 & " ",""),'
                'IF(C111!$C$5=C$3,C111!$B$1 & " ",""))'
                '), " ",",")',
        },
    ],

    # ── Columnas ──
    "column_widths": {i: 14 for i in range(1, 16)},
    "cell_size": 1.0,  # Factor de escala (1.0 = normal, 2.0 = doble)

    # ── Bordes ──
    "table_borders": True,
    "border_color": "4F81BD",   # Azul
    "border_style": "medium",    # "thin", "medium", "thick"
    "table_ranges": ["B3:G3", "B4:B15", "C4:G15", "I3:M3"],
    "table_block_sizes": [
        {"range": "B4:B15", "row_step": 3, "col_step": 1},
        {"range": "C4:G15", "row_step": 3, "col_step": 1},
    ],

    # ── Estilos por rango ──
    "range_styles": [
        {"range": "I3:I11", "style": {"bg_color": "A9D18E"}},  # Verde
        {"range": "B4:B15", "style": {"bg_color": "F4CCCC"}},  # Rojo claro
    ],

    # ── Celdas fusionadas ──
    "merge_ranges": ["B4:B6", "B7:B9", "B10:B12"],

    # ── Formato condicional ──
    "conditional_format_rules": [
        {
            "tipo": "filas_pares",
            "rango": "C4:G15",
            "formula": 'AND({celda}<>"", COUNTIF($I4:I11,{celda})=0)',
            "color": "F4A460",  # Salmón: abreviatura inválida
        },
        {
            "tipo": "rango",
            "rango": "J4:J11",
            "formula": 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
            "color": "00FF00",  # Verde: asignatura completa
        },
    ],
}
```

### 7.5. Referencia rápida de claves

| Clave | Tipo | Obligatorio | Descripción |
|-------|------|-------------|-------------|
| `title` | `str` | ✅ | Nombre de la hoja |
| `data` | `list[list]` | ✅ | Filas de datos |
| `formulas` | `list[dict]` | ❌ | Fórmulas por coordenada `(row, col)` |
| `fernando_formulas` | `list[dict]` | ❌ | Fórmulas cross-sheet por celda Excel |
| `column_widths` | `dict[int, float]` | ❌ | Ancho de columnas |
| `table_borders` | `bool` | ❌ | Activar bordes |
| `border_color` | `str` | ❌ | Color del borde (hex RGB) |
| `range_styles` | `list[dict]` | ❌ | Colores de fondo por rango |
| `merge_ranges` | `list[str]` | ❌ | Celdas a fusionar |
| `conditional_format_rules` | `list[dict]` | ❌ | Reglas de formato condicional |

---

## 8. Uso con el tutorial (`tutorial_hoja_con_formulas.py`)

El archivo `tutorial_hoja_con_formulas.py` es la **forma recomendada** de usar el sistema sin necesidad de Lisp ni JSON. Proporciona funciones de alto nivel que construyen automáticamente la configuración para `hoja_con_formulas.py`.

### 8.1. Fase 1: Solo nombres de grupos

Genera hojas de grupos **vacías** pero con toda la estructura (fórmulas, bordes, formato condicional) lista:

```python
from tutorial_hoja_con_formulas import build_config, generar_excel_personalizado

grupos = ["D111", "D211", "C111"]

config = build_config(grupos)
generar_excel_personalizado(config, "mi_horario_vacio.xlsx")
```

**Resultado**: Un Excel con 4 hojas (D111, D211, C111, Aulas). Cada hoja de grupo tiene las tablas vacías pero con fórmulas funcionales.

### 8.2. Fase 2: Grupos + Asignaturas

Añade las asignaturas de cada grupo:

```python
grupos = ["D111", "D211", "C111"]

asignaturas_por_grupo = {
    "D111": [
        ("AL", "Álgebra Lineal", 3),
        ("L", "Lógica", 2),
        ("IP", "Introducción a la Programación", 2),
        ("AM I", "Análisis Matemático I", 2),
    ],
    "D211": [
        ("MA", "Matemática y Aplicaciones", 2),
        ("Prb", "Probabilidades", 2),
        ("BD", "Bases de Datos", 2),
        ("ED", "Estructura de Datos", 2),
    ],
    "C111": [
        ("A I", "Álgebra I", 3),
        ("L", "Lógica", 2),
        ("P", "Programación", 3),
        ("AM I", "Análisis Matemático I", 2),
        ("F", "Filosofía", 2),
    ],
}

config = build_config(grupos, asignaturas_por_grupo)
generar_excel_personalizado(config, "mi_horario_completo.xlsx")
```

### 8.3. Formato alternativo: diccionarios

En lugar de tuplas, puedes usar diccionarios para las asignaturas:

```python
asignaturas_por_grupo = {
    "D111": [
        {"abrev": "AL", "asignatura": "Álgebra Lineal", "frec": 3},
        {"abrev": "L", "asignatura": "Lógica", "frec": 2},
    ],
}
```

### 8.4. Funciones disponibles

| Función | Descripción |
|---------|-------------|
| `build_config(grupos, asignaturas_por_grupo)` | Construye la config completa para `generar_excel_personalizado` |
| `build_group_sheet_config(grupo, subjects, aulas_catalogo)` | Construye la config de **una** hoja de grupo |
| `build_aulas_sheet_config(groups)` | Construye la config de la hoja **Aulas** |
| `build_aulas_fernando_formulas(groups)` | Genera las fórmulas cruzadas para la hoja Aulas |
| `build_aulas_data()` | Construye la estructura de datos de la hoja Aulas |
| `build_day_blocks()` | Calcula los bloques de días en Aulas |
| `build_row_names()` | Genera los nombres de filas del horario |
| `build_dynamic_merge_ranges()` | Genera los rangos de celdas fusionadas |
| `normalize_subject(item)` | Normaliza una asignatura a formato estándar |
| `col_letter(col_num)` | Convierte número de columna a letra Excel |

### 8.5. Parámetros globales ajustables

Al inicio de `tutorial_hoja_con_formulas.py` puedes modificar:

```python
# Número de turnos (por defecto 6)
TURNOS = 6

# Filas por turno (por defecto 3: asignatura + aula + separador)
HORARIO_ROW_STEP = 3

# Catálogo de aulas disponibles
AULAS_CATALOGO = [f"Aula {i}" for i in range(1, 10)] + ["Lab"]

# Colores
COLOR_VERDE_ASIGNATURAS = "A9D18E"
COLOR_ROJO_TURNOS = "F4CCCC"
COLOR_ROJO_AULAS = "E6B8AF"
COLOR_BORDE = "4F81BD"
```

---

## 9. Generación desde Lisp (pipeline completo)

Si tienes datos reales en `horario.json`, usa el pipeline completo.

### 9.1. Paso 1: Convertir JSON → Lisp

```bash
python3 convertir_json_a_lisp.py
```

Esto genera `variables_horario.lisp` con los 17 grupos y sus asignaturas.

### 9.2. Paso 2: Ejecutar con un comando

```bash
./ru.sh
```

Este script ejecuta:

1. `sbcl --non-interactive --load "replicar_propuesta_ods.lisp" --quit` → Genera `generar_propuesta_desde_lisp.py`
2. `python3 generar_propuesta_desde_lisp.py` → Genera `propuesta_horarios_desde_lisp.xlsx`

### 9.3. Paso 2 alternativo: Manual

```bash
# Generar el Python desde Lisp
sbcl --non-interactive --load "replicar_propuesta_ods.lisp" --quit

# Ejecutar el Python generado
python3 generar_propuesta_desde_lisp.py
```

### 9.4. Resultado

Se genera `propuesta_horarios_desde_lisp.xlsx` con **18 hojas**:

- 17 hojas de grupo (D111, D211, D311, D411, C111, C121, C122, C211, C212, C311, C312, C411, C412, M111, M211, M311, M411)
- 1 hoja "Aulas" con fórmulas cruzadas

---

## 10. Fórmulas y formato condicional

### 10.1. Fórmulas en hojas de grupo

| Fórmula | Columna | Descripción |
|---------|---------|-------------|
| `=K{n}-M{n}` | L (Faltan) | Resta Frecuencia − Asignadas |
| `=COUNTIF($C$4:$G$15, I{n})` | M (Asignadas) | Cuenta apariciones de la abreviatura en el horario |
| `=COUNTA(I4:I{n})` | M (Total) | Cuenta total de asignaturas |
| `=SUM(K4:K{n})` | M (Σ Frec) | Suma todas las frecuencias |
| `=COUNTA(O4:O{n})` | O (Total Aulas) | Cuenta aulas únicas usadas |
| `=COUNTA(C4:G{n})/3` | G (Ocupados) | Cuenta turnos ocupados |

### 10.2. Formato condicional

El sistema colorea automáticamente las celdas según su estado:

| Regla | Color | Se activa cuando |
|-------|-------|-----------------|
| **Abreviatura inválida** | 🟠 Salmón (`F4A460`) | La abreviatura no existe en la lista de asignaturas |
| **Aula inválida** | 🟡 Amarillo (`FFD700`) | El aula no está en el catálogo de aulas |
| **Asignatura sin aula** | 🔴 Rojo (`FF0000`) | Hay asignatura pero el aula está vacía |
| **Asignatura completa** | 🟢 Verde (`00FF00`) | Todas las sesiones están asignadas (Faltan = 0) |
| **Asignatura excedida** | 🔴 Rojo claro (`FF6B6B`) | Más sesiones asignadas de las necesarias (Faltan < 0) |
| **Asignatura parcial** | 🟠 Naranja (`FFA500`) | Algunas sesiones asignadas pero faltan más |

### 10.3. Placeholders en fórmulas de formato condicional

Las fórmulas de formato condicional usan placeholders que el motor reemplaza automáticamente:

| Placeholder | Se reemplaza por |
|-------------|-----------------|
| `{celda}` | Referencia de la celda actual (ej: `"C4"`) |
| `{celda_siguiente}` | Referencia de la fila siguiente (ej: `"C5"`) |
| `{fila}` | Número de fila (ej: `"4"`) |
| `{columna}` | Letra de columna (ej: `"C"`) |

Ejemplo:

```python
{
    "tipo": "rango",
    "rango": "J4:J11",
    "formula": 'AND({celda}<>"", M{fila}>0, L{fila}=0)',
    "color": "00FF00",
}
```

Para la celda `J5`, se convierte en:

```excel
=AND(J5<>"", M5>0, L5=0)
```

---

## 11. La hoja "Aulas"

### 11.1. Qué muestra

La hoja **Aulas** es una **vista global** que responde a la pregunta:

> *"¿Qué grupos están en el Aula 3 el Lunes a las 10:00?"*

### 11.2. Estructura

La hoja está organizada en **5 bloques** (uno por día), cada uno con 6 turnos:

```
┌──────┬────────┬────────┬────────┬───┬────────┐
│      │ Aula 1 │ Aula 2 │ Aula 3 │...│  Lab   │
├──────┼────────┼────────┼────────┼───┼────────┤
│Lunes │        │        │        │   │        │
│ 1ro  │ D111   │        │ C112   │   │        │
│ 2do  │        │        │ D111   │   │        │
│ 3ro  │        │        │        │   │        │
│ 4to  │        │ M211   │        │   │        │
│ 5to  │        │        │        │   │        │
│ 6to  │        │        │        │   │        │
├──────┼────────┼────────┼────────┼───┼────────┤
│Martes│        │        │        │   │        │
│ 1ro  │        │ C112   │        │   │        │
│ ...  │        │        │        │   │        │
└──────┴────────┴────────┴────────┴───┴────────┘
```

### 11.3. Fórmulas cruzadas

Cada celda contiene una fórmula que **lee las 17 hojas de grupo**:

```excel
=SUBSTITUTE(TRIM(CONCAT(
    IF(D111!$C$5=C$3, D111!$B$1 & " ", ""),
    IF(D211!$C$5=C$3, D211!$B$1 & " ", ""),
    IF(C111!$C$5=C$3, C111!$B$1 & " ", ""),
    ...  (todos los grupos)
)), " ",",")
```

**Qué hace:**

1. Para cada grupo, verifica si el aula de esa celda coincide con el aula en el horario del grupo.
2. Si coincide, concatena el nombre del grupo seguido de un espacio.
3. Al final, reemplaza los espacios por comas.

**Resultado:** `"D111,C111,M211"` — los grupos que están en esa aula en ese turno.

### 11.4. Generar fórmulas de Aulas

```python
from tutorial_hoja_con_formulas import build_aulas_fernando_formulas

grupos = ["D111", "D211", "C111"]
formulas = build_aulas_fernando_formulas(grupos)

# formulas contiene ~150 fórmulas (5 días × 6 turnos × 10 aulas)
```

---

## 12. Ejemplos prácticos

### Ejemplo 1: Horario minimalista (2 grupos, 2 asignaturas)

```python
from tutorial_hoja_con_formulas import build_config, generar_excel_personalizado

grupos = ["D111", "C111"]

asignaturas_por_grupo = {
    "D111": [
        ("AL", "Álgebra Lineal", 3),
        ("L", "Lógica", 2),
    ],
    "C111": [
        ("P", "Programación", 3),
        ("A I", "Álgebra I", 2),
    ],
}

config = build_config(grupos, asignaturas_por_grupo)
generar_excel_personalizado(config, "horario_minimal.xlsx")
```

### Ejemplo 2: Horario de una sola carrera

```python
# Solo Ciencia de Datos
grupos = ["D111", "D211", "D311", "D411"]

asignaturas_por_grupo = {
    "D111": [("AL", "Álgebra Lineal", 3), ("L", "Lógica", 2)],
    "D211": [("MA", "Matemática", 2), ("Prb", "Probabilidades", 2)],
    "D311": [("ML", "Machine Learning", 3), ("ED", "Estructura de Datos", 2)],
    "D411": [("DL", "Deep Learning", 3), ("NLP", "Procesamiento de Lenguaje", 2)],
}

config = build_config(grupos, asignaturas_por_grupo)
generar_excel_personalizado(config, "horario_data_science.xlsx")
```

### Ejemplo 3: Horario con aulas personalizadas

```python
from tutorial_hoja_con_formulas import (
    build_group_sheet_config,
    build_aulas_sheet_config,
    generar_excel_personalizado,
)

grupos = ["D111", "C111"]

# Aulas personalizadas
mis_aulas = ["A-101", "A-102", "Lab-1", "Lab-2", "Auditorio"]

sheets = [
    build_group_sheet_config("D111", [
        ("AL", "Álgebra Lineal", 3),
        ("L", "Lógica", 2),
    ], aulas_catalogo=mis_aulas),
    build_group_sheet_config("C111", [
        ("P", "Programación", 3),
        ("A I", "Álgebra I", 2),
    ], aulas_catalogo=mis_aulas),
    build_aulas_sheet_config(grupos),
]

config = {"sheets": sheets}
generar_excel_personalizado(config, "horario_aulas_personalizadas.xlsx")
```

### Ejemplo 4: Hoja personalizada desde cero

```python
from hoja_con_formulas import generar_excel_personalizado

config = {
    "sheets": [{
        "title": "MiHorario",
        "data": [
            ["", "", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes"],
            ["Turno 1", "AL", "L", "AL", "F", "IP"],
            ["", "Aula 1", "Aula 2", "Aula 1", "Aula 3", "Aula 2"],
            ["Turno 2", "L", "AL", "IP", "L", ""],
            ["", "Aula 2", "Aula 1", "Aula 3", "Aula 2", ""],
        ],
        "table_borders": True,
        "border_color": "4F81BD",
        "column_widths": {i: 14 for i in range(1, 8)},
        "range_styles": [
            {"range": "A1:G1", "style": {"bg_color": "E6B8AF"}},
        ],
    }]
}

generar_excel_personalizado(config, "mi_horario_manual.xlsx")
```

### Ejemplo 5: Integración con datos desde JSON

```python
import json
from tutorial_hoja_con_formulas import build_config, generar_excel_personalizado

# Leer datos desde horario.json
with open("horario.json", "r") as f:
    datos = json.load(f)

# Extraer grupos y asignaturas
grupos = []
asignaturas_por_grupo = {}

for carrera in datos["carreras"]:
    for grupo_data in carrera["grupos"]:
        grupo_id = grupo_data["nombre"]  # ej: "D111"
        grupos.append(grupo_id)

        asignaturas = []
        for asig in grupo_data["asignaturas"]:
            asignaturas.append((
                asig["abrev"],
                asig["nombre"],
                asig["horas_semanales"],
            ))
        asignaturas_por_grupo[grupo_id] = asignaturas

# Generar Excel
config = build_config(grupos, asignaturas_por_grupo)
generar_excel_personalizado(config, "horario_desde_json.xlsx")
```

---

## 13. Preguntas frecuentes

### ¿Puedo usar esto sin Lisp?

**Sí.** El tutorial (`tutorial_hoja_con_formulas.py`) funciona solo con Python. Lisp es necesario solo si quieres generar horarios desde `horario.json` automáticamente.

### ¿Qué pasa si no tengo `horario.json`?

No hay problema. Puedes definir los grupos y asignaturas manualmente en Python usando el tutorial.

### ¿Cómo agrego un nuevo grupo?

Simplemente añade el nombre a la lista `grupos`:

```python
grupos = ["D111", "D211", "C111", "X511"]  # Nuevo grupo X511
```

Y define sus asignaturas:

```python
asignaturas_por_grupo["X511"] = [
    ("X1", "Asignatura X1", 3),
    ("X2", "Asignatura X2", 2),
]
```

### ¿Cómo cambio el número de turnos?

Modifica `TURNOS` en `tutorial_hoja_con_formulas.py`:

```python
TURNOS = 4  # Solo 4 turnos en lugar de 6
```

### ¿Cómo cambio las aulas disponibles?

Modifica `AULAS_CATALOGO`:

```python
AULAS_CATALOGO = ["A-101", "A-102", "Lab-1", "Auditorio"]
```

### ¿Por qué las fórmulas no se recalculan en LibreOffice?

LibreOffice Calc a veces no recalcula fórmulas automáticamente al abrir un archivo generado por openpyxl. Solución:

1. Abre el archivo en LibreOffice Calc.
2. Presiona `Ctrl + Shift + F9` (recalcular todo).
3. O ve a **Datos → Calcular**.

### ¿Puedo generar archivos `.ods` en lugar de `.xlsx`?

`hoja_con_formulas.py` solo genera `.xlsx`. Para `.ods`, usa los scripts `generar_ods_python.py` o `read_ods.py` del proyecto.

### ¿Qué son las "fórmulas de Fernando"?

Son las **fórmulas cruzadas** extraídas del archivo ODS original de Fernando. Permiten que la hoja "Aulas" lea automáticamente las 17 hojas de grupo y concatene qué grupos están en cada aula/turno.

### ¿Cómo depuro un Excel generado?

1. Abre el archivo en Excel.
2. Verifica que las hojas existen y tienen el nombre correcto.
3. Revisa que las fórmulas se muestran como fórmulas (no como texto).
4. Comprueba que el formato condicional se aplica (celdas coloreadas).
5. Si algo falla, revisa la consola: `hoja_con_formulas.py` imprime errores detallados.

---

## Apéndice A: Estructura de carpetas del proyecto

```
python_excel_tesis/
├── hoja_con_formulas.py              ← Motor de Excel (librería principal)
├── tutorial_hoja_con_formulas.py     ← Tutorial y funciones de alto nivel
├── convertir_json_a_lisp.py          ← Conversor JSON → Lisp
├── generar_propuesta_desde_lisp.py   ← Python generado por Lisp (auto-generado)
├── ejecutor_aulas.py                 ← Python para Aulas (auto-generado)
├── horario.json                      ← Datos fuente de horarios
├── variables_horario.lisp            ← Datos en Lisp (auto-generado)
├── replicar_propuesta_ods.lisp       ← Orquestador Lisp → Python
├── codigo-tesis.lisp                 ← Framework base Lisp (clases, macros)
├── ru.sh                             ← Script de un comando
├── FLUJO_DEL_PROYECTO.md             ← Documentación técnica del pipeline
├── README_SISTEMA_INTEGRADO.md       ← Manual del sistema integrado
└── TUTORIAL_HOJA_CON_FORMULAS.md     ← Tutorial anterior (versión corta)
```

## Apéndice B: Glosario

| Término | Significado |
|---------|-------------|
| **Grupo** | Cohort de estudiantes (ej: D111, C111, M211) |
| **Asignatura** | Materia que cursa un grupo (ej: Álgebra, Lógica) |
| **Aula** | Espacio físico (Aula 1-9, Lab) |
| **Turno** | Franja horaria (6 turnos por día) |
| **Frecuencia** | Sesiones semanales de una asignatura |
| **Faltan** | Sesiones aún no asignadas |
| **Asignadas** | Sesiones ya ubicadas en el horario |
| **Hoja de grupo** | Hoja Excel con el horario de un grupo |
| **Hoja Aulas** | Hoja Excel con vista global de aulas por grupo |
| **Fórmula cruzada** | Fórmula que lee celdas de otras hojas |
| **Pipeline** | Cadena de generación (JSON → Lisp → Python → Excel) |

## Apéndice C: Comandos rápidos

```bash
# Tutorial (solo Python, recomendado para empezar)
python3 tutorial_hoja_con_formulas.py

# Pipeline completo (desde JSON)
./ru.sh

# O manualmente
sbcl --non-interactive --load "replicar_propuesta_ods.lisp" --quit
python3 generar_propuesta_desde_lisp.py

# Ver ayuda de hoja_con_formulas.py
python3 -c "from hoja_con_formulas import generar_excel_personalizado; help(generar_excel_personalizado)"
```

---

> **Nota:** Este tutorial está diseñado como documentación de librería. Cada función puede usarse de forma independiente o combinarse para flujos de trabajo personalizados. Para la documentación técnica completa del pipeline, consulta [`FLUJO_DEL_PROYECTO.md`](FLUJO_DEL_PROYECTO.md).
