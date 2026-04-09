# Flujo del Proyecto: Generador de Horarios MATCOM

## Descripción General

Sistema que genera archivos Excel con propuestas de horarios universitarios. Usa **Lisp como generador de código Python**, y **Python como motor de generación de Excel**. El pipeline convierte datos de horarios (JSON) en un archivo Excel formateado con fórmulas, bordes, colores y formato condicional.

---

## Arquitectura General

```
┌─────────────────────────────────────────────────────────────────────┐
│                        FLUJO PRINCIPAL (Pipeline A)                  │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  horario.json          variables_horario.lisp      SBCL REPL        │
│       │                       │                      │              │
│       ▼                       ▼                      ▼              │
│  convertir_json_      (carga automática)    replicar_propuesta_     │
│  _a_lisp.py                                    ods.lisp             │
│       │                                          │                  │
│       │                                          │ genera           │
│       │                                          ▼                  │
│       │                              generar_propuesta_             │
│       │                              _desde_lisp.py                 │
│       │                                          │                  │
│       │                                          │ ejecuta          │
│       │                                          ▼                  │
│       │                              propuesta_horarios_            │
│       │                              _desde_lisp.xlsx               │
│       │                                          ▲                  │
│       │                                          │                  │
│       │                            hoja_con_formulas.py ◄───────────┘
│       │                            (motor Excel)                     │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Ejecución Paso a Paso

### Paso 1: Conversión de datos JSON → Lisp

```
horario.json ──► convertir_json_a_lisp.py ──► variables_horario.lisp
```

**Qué hace:** Lee el JSON crudo con los horarios de cada grupo y genera un archivo Lisp con `defparameter` para cada grupo (`*horario-d111*`, `*horario-c111*`, etc.) y sus asignaturas (`*asignaturas-d111*`, etc.).

**Se ejecuta cuando:** Cambian los datos fuente en `horario.json`.

---

### Paso 2: Carga del sistema en SBCL

```lisp
(load "codigo-tesis.lisp")          ;; Clases base y macros
(load "variables_horario.lisp")     ;; Datos de horarios
(load "replicar_propuesta_ods.lisp") ;; Orquestador principal
```

**Qué pasa internamente:**

1. `codigo-tesis.lisp` define las clases fundamentales:
   - `clase-tabla` → representa una tabla con filas, columnas, contenido, estilos
   - `clase-hoja` → representa una hoja del Excel (horario + asignaturas)
   - `clase-libro` → contiene múltiples hojas
   - `defclass*` → macro que crea la clase + constructor automáticamente

2. `variables_horario.lisp` carga los datos de los 17 grupos (D111-D411, C111-C412, M111-M411).

3. `replicar_propuesta_ods.lisp` es el orquestador que:
   - Construye objetos `clase-hoja` para cada grupo
   - Construye la hoja `Aulas` (matriz de aulas por día/hora)
   - Configura fórmulas, bordes, colores, formato condicional
   - Genera el archivo Python mediante el método `generate-code`

---

### Paso 3: Generación del código Python

```
replicar_propuesta_ods.lisp ──► generar_propuesta_desde_lisp.py
         (Lisp)                        (Python generado)
```

**Qué genera:** Un archivo Python de ~1400 líneas con la configuración completa de las 18 hojas:

- **Datos** (`'data'`): matrices con el contenido de cada celda
- **Fórmulas** (`'formulas'`): referencias de celda con fórmulas Excel
- **Formato condicional** (`'conditional_format_rules'`): reglas de coloreo automático
- **Rangos de merge** (`'merge_ranges'`): celdas combinadas (ej: "Turno 1" en 2 filas)
- **Estilos** (`'range_styles'`): colores de fondo por rango
- **Bordes** (`'table_borders'`): bordes de las tablas

---

### Paso 4: Generación del Excel

```
generar_propuesta_desde_lisp.py ──► propuesta_horarios_desde_lisp.xlsx
         (Python)                         (Excel final)
              │
              └── importa ──► hoja_con_formulas.py
```

**Qué hace `hoja_con_formulas.py`:**

1. Crea el workbook con `openpyxl`
2. Por cada hoja en la configuración:
   - Escribe headers y datos
   - Inserta fórmulas en celdas específicas
   - Aplica anchos de columna
   - Aplica bordes por bloques
   - Aplica colores de fondo por rango
   - Fusiona celdas (merge)
   - Aplica reglas de formato condicional
3. Inyecta las fórmulas de Fernando en la hoja Aulas (fórmulas cruzadas entre hojas)
4. Guarda el archivo `.xlsx`

---

## Mapa de Archivos

### Archivos Fuente (Lisp)

| Archivo | Rol | Dependencias |
|---------|-----|-------------|
| `codigo-tesis.lisp` | **Framework base** - Clases (`clase-tabla`, `clase-hoja`, `clase-libro`), macros (`defclass*`), método polimórfico `generate-code` | Ninguno |
| `variables_horario.lisp` | **Capa de datos** - 17 horarios + 17 listas de asignaturas como `defparameter` | Generado por `convertir_json_a_lisp.py` |
| `replicar_propuesta_ods.lisp` | **Orquestador principal** - Construye las 18 hojas, configura fórmulas/formato, genera el Python | `codigo-tesis.lisp`, `variables_horario.lisp` |
| `generador_formulas_dinamicas.lisp` | **Generador de fórmulas cruzadas** - Crea fórmulas IF/CONCATENATE que referencian las 17 hojas de grupo | `codigo-tesis.lisp` |
| `formulas_fernando.lisp` | **Cargador de fórmulas de Fernando** - Lee JSON con fórmulas complejas extraídas del ODS original | `codigo-tesis.lisp`, JSON de fórmulas |
| `paso_a_paso_aulas_con_formulas.lisp` | **Generador alternativo de Aulas** - Genera Python solo para la hoja Aulas con fórmulas de Fernando | `codigo-tesis.lisp`, JSON de fórmulas |
| `sistema_horarios_integrado.lisp` | **Sistema integrado** - Combina todos los módulos, detecta grupos automáticamente, genera `ejecutor_aulas.py` | `codigo-tesis.lisp`, `generador_formulas_dinamicas.lisp`, `variables_horario.lisp` |

### Archivos Fuente (Python)

| Archivo | Rol | Dependencias |
|---------|-----|-------------|
| `hoja_con_formulas.py` | **Motor de Excel** - Usa openpyxl para crear workbooks con todo el formato | `openpyxl` |
| `convertir_json_a_lisp.py` | **Conversor JSON→Lisp** - Genera `variables_horario.lisp` desde `horario.json` | Ninguno |
| `extraer_y_convertir_formulas.py` | **Extractor de fórmulas ODS** - Parsea .ods (zip+XML), extrae y convierte fórmulas a sintaxis Excel | Ninguno |
| `arreglar_formulas_fernando.py` | **Adaptador de fórmulas** - Adapta fórmulas de Fernando a los nombres reales de hojas | JSON de fórmulas |
| `arreglar_aulas_fernando.py` | **Copiador de hojas** - Copia hojas faltantes entre archivos Excel | openpyxl |
| `corregir_formulas_referencias.py` | **Corrector de referencias** - Arregla notación de punto a notación de exclamación | openpyxl |
| `regenerar_json_formulas.py` | **Regenerador de JSON** - Genera `formulas_fernando_adaptadas_es.json` desde cero | Ninguno |

### Archivos de Datos

| Archivo | Rol |
|---------|-----|
| `horario.json` | Datos crudos de horarios (entrada del sistema) |
| `formulas_fernando_adaptadas_es.json` | Fórmulas cruzadas adaptadas (hoja Aulas, locale español) |
| `formulas_fernando_convertidas_es.json` | 215 fórmulas extraídas del ODS original (español) |
| `formulas_fernando_convertidas_en.json` | 215 fórmulas extraídas del ODS original (inglés) |

### Archivos Generados (output del pipeline)

| Archivo | Generado por | Descripción |
|---------|-------------|-------------|
| `generar_propuesta_desde_lisp.py` | `replicar_propuesta_ods.lisp` | Config Python de las 18 hojas |
| `ejecutor_aulas.py` | `sistema_horarios_integrado.lisp` | Python con fórmulas dinámicas para Aulas |
| `ejecutar_hoja_con_formulas_con_fernando.py` | `paso_a_paso_aulas_con_formulas.lisp` | Python solo para Aulas con fórmulas de Fernando |

### Archivos de Ejecución y Documentación

| Archivo | Rol |
|---------|-----|
| `ru.sh` | Script de un comando: ejecuta Lisp → genera Python → genera Excel |
| `README_SISTEMA_INTEGRADO.md` | Manual del sistema integrado |
| `README_FORMULAS_FERNANDO.md` | Guía de integración de fórmulas de Fernando |
| `RESUMEN_SISTEMA_COMPLETO.md` | Resumen ejecutivo del sistema |

---

## Estructura de Datos

### Cómo se representa un horario en Lisp

```lisp
;; En variables_horario.lisp
(defparameter *horario-d111*
  '(("F" "ICD" "AL" "AM I" "AL")        ;; Turno 1: asignaturas (Lun-Vie)
    ("Aula 8" "Aula 7" "Aula 7*" ...)   ;; Turno 1: aulas
    ("L" "AL" "EF" "L" "")              ;; Turno 2: asignaturas
    ("Aula 6*" "Aula 6*" "SEDER" ...)   ;; Turno 2: aulas
    ...))                                ;; 12 filas totales (6 turnos × 2)

(defparameter *asignaturas-d111*
  '(("AL" "Álgebra Lineal" 3 0 0)       ;; (abrev, nombre, frecuencia, faltan, asignadas)
    ("L" "Lógica" 2 0 0)
    ("F" "Fundamentos de Programación" 3 0 0)
    ...))
```

### Cómo se convierte en objeto Lisp

```lisp
;; En replicar_propuesta_ods.lisp
(crear-hoja-grupo "D111" *horario-d111* *asignaturas-d111*)
;; → Crea un objeto clase-hoja con:
;;    - horario: clase-tabla (matriz 12×5)
;;    - asignaturas: clase-tabla (lista de asignaturas)
```

### Cómo se convierte en configuración Python

```python
# En generar_propuesta_desde_lisp.py (generado por Lisp)
sheets_cfg.append({
    'title': "D111",
    'data': [
        ["Grupo ", "D111", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["", "Turno", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes",
         "", "Abrev", "Asignaturas", "Frec", "Faltan", "Asignadas", "", "Aulas"],
        # ... filas de datos del horario y asignaturas ...
    ],
    'formulas': [
        {'row': 11, 'col': 12, 'value': "Total:"},
        {'row': 11, 'col': 13, 'value': "=COUNTA(I4:I10)"},
        {'row': 4, 'col': 13, 'value': "=COUNTIF($C$4:$G$15,I4)"},   # Asignadas
        {'row': 4, 'col': 12, 'value': "=K4-M4"},                     # Faltan
        # ... más fórmulas ...
    ],
    'conditional_format_rules': [...],
    'merge_ranges': ["B4:B5", "B6:B7", ...],
    'range_styles': [{'range': 'I3:I10', 'style': {'bg_color': 'A9D18E'}}],
    'table_borders': True,
    'border_color': '4F81BD',
})
```

---

## Las 3 Pipelines

### Pipeline A: Producción Principal (replicar_propuesta_ods)

```
┌──────────────┐     ┌───────────────────────┐     ┌──────────────────────────┐
│ horario.json │────►│ convertir_json_a_lisp │────►│ variables_horario.lisp   │
│  (datos)     │     │      .py              │     │  (datos en Lisp)         │
└──────────────┘     └───────────────────────┘     └───────────┬──────────────┘
                                                               │
                              ┌────────────────────────────────┘
                              ▼
                    ┌─────────────────────────┐
                    │ SBCL REPL               │
                    │ (load "codigo-tesis")   │
                    │ (load "variables")      │
                    │ (load "replicar_ods")   │
                    └───────────┬─────────────┘
                                │
                                ▼
                    ┌───────────────────────────────────┐
                    │ generar_propuesta_desde_lisp.py   │
                    │ (18 hojas configuradas)           │
                    └───────────┬───────────────────────┘
                                │
                                ▼
                    ┌───────────────────────────────────┐
                    │ propuesta_horarios_desde_lisp.xlsx│
                    │ (17 grupos + Aulas)               │
                    └───────────────────────────────────┘
```

**Produce:** Excel completo con 18 hojas (17 grupos + Aulas), con fórmulas, formato condicional, bordes y colores.

### Pipeline B: Sistema Integrado (sistema_horarios_integrado)

```
┌──────────────────────────┐
│ variables_horario.lisp   │
│  (datos en Lisp)         │
└───────────┬──────────────┘
            │
            ▼
┌───────────────────────────────────┐
│ SBCL REPL                         │
│ (load "sistema_horarios_integrado")│
│                                   │
│ • Detecta 17 grupos automáticamente│
│ • Genera fórmulas cruzadas        │
│ • Crea ejecutor_aulas.py          │
└───────────┬───────────────────────┘
            │
            ▼
┌───────────────────────────────────┐
│ ejecutar_hoja_con_formulas_       │
│ con_fernando.py                   │
│ (inyecta 140+ fórmulas en Aulas)  │
└───────────┬───────────────────────┘
            │
            ▼
┌───────────────────────────────────┐
│ Aulas_Con_Formulas_Fernando.xlsx  │
│ (fórmulas que referencian las     │
│  17 hojas de grupo)               │
└───────────────────────────────────┘
```

**Produce:** Excel con fórmulas dinámicas en Aulas que concatenan qué grupos están en cada aula/hora.

### Pipeline C: Extracción de Fórmulas de Fernando

```
┌──────────────────────────────────┐
│ propuesta-de-horarios-           │
│ fernando-v3.ods                  │
│ (archivo ODS original)           │
└───────────┬──────────────────────┘
            │
            ▼
┌───────────────────────────────────┐
│ extraer_y_convertir_formulas.py   │
│                                   │
│ • Descomprime ODS (zip+XML)       │
│ • Extrae 215 fórmulas             │
│ • Convierte sintaxis ODS→Excel    │
└───────────┬───────────────────────┘
            │
            ▼
┌───────────────────────────────────┐
│ formulas_fernando_convertidas_    │
│ es.json / en.json                 │
└───────────┬───────────────────────┘
            │
            ▼
┌───────────────────────────────────┐
│ arreglar_formulas_fernando.py     │
│ (adapta a nombres reales de hojas)│
└───────────┬───────────────────────┘
            │
            ▼
┌───────────────────────────────────┐
│ formulas_fernando_adaptadas_es.json│
└───────────────────────────────────┘
```

**Produce:** JSON con fórmulas listas para inyectar en el Excel generado.

---

## Estructura de una Hoja de Grupo

```
┌─────────────────────────────────────────────────────────────────────────┐
│                          Grupo D111                                     │
├────┬────────┬─────────┬─────────┬───────────┬─────────┬─────────┬───────┤
│    │ Turno  │ Lunes   │ Martes  │ Miércoles │ Jueves  │ Viernes │       │
├────┼────────┼─────────┼─────────┼───────────┼─────────┼─────────┼───────┤
│    │Turno 1 │   AL    │   AM I  │    AL     │   F     │   ICD   │       │
│    │        │ Aula 8  │ Aula 7  │  Aula 7*  │ Aula 1  │ Aula 11 │       │
├────┼────────┼─────────┼─────────┼───────────┼─────────┼─────────┼───────┤
│    │Turno 2 │   L     │   AL    │    EF     │   L     │         │       │
│    │        │ Aula 6* │ Aula 6* │  SEDER    │ Aula 6* │         │       │
├────┼────────┼─────────┼─────────┼───────────┼─────────┼─────────┼───────┤
│    │ ...    │  ...    │  ...    │   ...     │  ...    │  ...    │       │
├────┼────────┼─────────┼─────────┼───────────┼─────────┼─────────┼───────┤
│    │        │         │         │           │         │         │       │
├────┼────────┴─────────┴─────────┴───────────┴─────────┴─────────┼───────┤
│    │ Abrev │ Asignaturas            │ Frec │ Faltan │ Asignadas│ Aulas │
├────┼───────┼────────────────────────┼──────┼────────┼──────────┼───────┤
│    │  AL   │ Álgebra Lineal         │  3   │ =K4-M4 │ =COUNTIF │ ...   │
│    │  L    │ Lógica                 │  2   │ =K5-M5 │ =COUNTIF │ ...   │
│    │  F    │ Fund. de Programación  │  3   │ =K6-M6 │ =COUNTIF │ ...   │
│    │  ...  │ ...                    │ ...  │ ...    │ ...      │ ...   │
├────┼───────┼────────────────────────┼──────┼────────┼──────────┼───────┤
│    │       │                        │      │Total:  │ =COUNTA  │       │
│    │       │                        │      │Σ Frec: │ =SUM     │       │
└────┴───────┴────────────────────────┴──────┴────────┴──────────┴───────┘

Columnas: A(empty) B(Turno) C-G(Días) H(sep) I(Abrev) J(Asignatura) K(Frec) L(Faltan) M(Asignadas) N(sep) O(Aulas)
```

---

## Estructura de la Hoja Aulas

```
┌──────────────────────────────────────────────────────────────────────────────────┐
│                                    AULAS                                         │
├──────┬──────────┬──────────┬──────────┬──────────┬──────────┬──────────┬─────────┤
│      │ Aula 1   │ Aula 2   │ Aula 3   │ Aula 4   │ Aula 5   │ Aula 6   │ Lab     │
├──────┼──────────┼──────────┼──────────┼──────────┼──────────┼──────────┼─────────┤
│ LUNES│          │          │          │          │          │          │         │
│  8:00│ D111     │          │          │          │          │ D111     │         │
│ 10:00│          │          │          │          │          │ D111     │         │
│ 12:00│          │          │          │          │          │ D111     │         │
│ 14:00│          │          │          │          │          │          │         │
│ 16:00│          │          │          │          │          │          │         │
│ 18:00│          │          │          │          │          │          │         │
├──────┼──────────┼──────────┼──────────┼──────────┼──────────┼──────────┼─────────┤
│ MARTES│         │          │          │          │          │          │         │
│  8:00│          │ C112     │          │          │          │ D111     │         │
│ ...  │          │          │          │          │          │          │         │
└──────┴──────────┴──────────┴──────────┴──────────┴──────────┴──────────┴─────────┘

Cada celda contiene una fórmula =IF(COUNTIF(Grupo!Celda, Aula), Grupo, "")
que concatena los grupos asignados a esa aula en ese horario.
```

---

## Relación entre Fórmulas

### Fórmulas en hojas de grupo

| Fórmula | Columna | Qué hace |
|---------|---------|----------|
| `=K{n}-M{n}` | L (Faltan) | Resta Frecuencia menos Asignadas |
| `=COUNTIF($C$4:$G$15,I{n})` | M (Asignadas) | Cuenta cuántas veces aparece la abreviatura en el horario |
| `=COUNTA(I4:I{n})` | M (Total) | Cuenta total de asignaturas |
| `=SUM(K4:K{n})` | M (Σ Frec) | Suma todas las frecuencias |
| `=COUNTA(O4:O{n})` | O (Total Aulas) | Cuenta aulas únicas |
| `=COUNTA(C4:G{n})/2` | G (Ocupados) | Cuenta turnos ocupados |

### Formato condicional en hojas de grupo

| Regla | Fórmula | Color | Cuándo se activa |
|-------|---------|-------|-----------------|
| Asignatura inválida | `COUNTIF(abrev_range, celda)=0` | Salmón | La abreviatura no existe en la lista |
| Aula inválida | `COUNTIF(aulas_range, celda)=0` | Amarillo | El aula no está en la lista de aulas válidas |
| Asignatura completa | `Asignadas>0 AND Faltan=0` | Verde | Todas las sesiones están asignadas |
| Asignatura excedida | `Asignadas>0 AND Faltan<0` | Rojo | Más sesiones asignadas de las necesarias |
| Asignatura parcial | `Asignadas>0 AND 0<Faltan<Frec` | Naranja | Faltan sesiones por asignar |

### Fórmulas de Fernando (hoja Aulas)

Fórmulas cruzadas que referencian las 17 hojas de grupo. Ejemplo:

```excel
=IF(COUNTIF(D111!C4, "Aula 1"), "D111", "") & IF(COUNTIF(C111!C4, "Aula 1"), "C111", "") & ...
```

Cada celda de Aulas verifica si algún grupo tiene esa aula en ese horario y concatena los nombres.

---

## Ejecución Rápida

```bash
# Un solo comando (usa ru.sh)
./ru.sh

# O manualmente:
sbcl --non-interactive --load "replicar_propuesta_ods.lisp" --quit
python3 generar_propuesta_desde_lisp.py
```

---

## Dependencias Externas

| Dependencia | Versión | Para qué se usa |
|-------------|---------|-----------------|
| SBCL | 2.1.11+ | Ejecutar código Lisp |
| Python | 3.8+ | Generar el Excel |
| openpyxl | 3.0+ | Crear archivos .xlsx con formato |
| lxml | 4.9+ | Parsear XML del ODS (solo para extracción de fórmulas) |
