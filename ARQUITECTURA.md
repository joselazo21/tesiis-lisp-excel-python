# Arquitectura del DSL de Hojas de Cálculo

## Visión general

El sistema transforma una descripción declarativa de un workbook en un archivo
Excel funcional. Para hacerlo de forma mantenible, el trabajo está dividido en
**tres capas con responsabilidades estrictamente separadas**:

```
┌─────────────────────────────────────────────────────────┐
│  ast-facultad.lisp (o cualquier archivo de workbook)    │
│  "QUÉ quiero"  —  usa macros del DSL                    │
└────────────────────────┬────────────────────────────────┘
                         │ macros construyen objetos xl-*
┌────────────────────────▼────────────────────────────────┐
│  ast-def.lisp + dsl-directo.lisp                        │
│  "CÓMO SE ESCRIBE"  —  clases AST + macros DSL          │
└────────────────────────┬────────────────────────────────┘
                         │ compile-excel-formula traduce
┌────────────────────────▼────────────────────────────────┐
│  generate-code-direct.lisp                              │
│  "A QUÉ SE TRADUCE"  —  backend Excel/Python            │
└─────────────────────────────────────────────────────────┘
```

---

## Capa 1 — AST (`ast-def.lisp`)

### Responsabilidad

Declarar la **estructura** de los nodos del árbol. Solo `defclass*`.
No hay lógica, no hay strings de backend, no hay comportamiento.

### Regla fundamental

> **Ninguna clase del AST puede contener un string, símbolo o concepto
> que sea específico de un backend** (Excel, Python, ODS, etc.).

Correcto:
```lisp
(defclass* xl-expr-cross-cell () (sheet xcol row))
```

Incorrecto — el AST no sabe de letras Excel:
```lisp
;; MAL: "B" es un detalle de Excel
(defclass* xl-expr-cross-cell () (sheet col-letter row))
```

### Qué va aquí

- Expresiones atómicas: `xl-expr-if`, `xl-expr-equals`, `xl-expr-string`...
- Expresiones de rango: `xl-range`, `xl-expr-countif`...
- Expresiones cross-sheet: `xl-expr-cross-cell`, `xl-expr-collect-over`...
- Estructura del workbook: `xl-table`, `xl-region`, `xl-sheet`, `xl-workbook`

### Qué NO va aquí

- Métodos de compilación (`compile-excel-formula`)
- Strings de Excel (`"$A$1"`, `"SUBSTITUTE"`, `"IF"`)
- Conocimiento de la estructura interna de otras tablas

---

## Capa 2 — DSL (`dsl-directo.lisp`)

### Responsabilidad

Exponer una **sintaxis limpia de dominio** que construye nodos AST.
Los macros son azúcar sintáctico puro — construyen objetos, no generan código.

### Regla fundamental

> **Los macros solo crean objetos `xl-*`. No producen strings de fórmulas.**

```lisp
;; Correcto: crea un objeto AST
(defmacro equals (a b)
  `(xl-expr-equals :a ,a :b ,b))

;; Incorrecto: genera código Excel directamente
(defmacro equals (a b)
  (format nil "~a=~a" a b))  ;; ← esto pertenece en generate-code
```

### Qué va aquí

- Un macro por cada clase de expresión del AST
- Macros de estructura: `def-table`, `tabla`, `hoja`, `hoja-v`, `libro`
- Documentación de sintaxis y ejemplos de uso
- Convenciones de quoting de argumentos

### Quoting de argumentos — la regla

| Tipo de argumento | Qué hace el macro | Cómo escribe el caller |
|---|---|---|
| Nombre fijo (columna, tabla) | Quotea: `',sym` | Sin comilla: `(col lun)` |
| Variable en scope (inst-param) | No quotea: `,sym` | Sin comilla: `:col dia` |
| Literal símbolo explícito | No quotea: `,sym` | Con comilla: `:col 'a` |
| Expresión evaluada | No quotea: `,expr` | Como expresión: `:row 1` |

### Qué NO va aquí

- Macros `defmacro` en los archivos de workbook (`ast-facultad.lisp`)
- Lógica de compilación
- Strings de Excel
- Conocimiento del número de grupos o la estructura de tablas ajenas

---

## Capa 3 — Backend (`generate-code-direct.lisp`)

### Responsabilidad

Traducir cada clase del AST a código concreto del target. Es el **único lugar**
donde puede vivir el conocimiento específico del backend.

### Regla fundamental

> **Todo lo específico de Excel vive aquí y solo aquí.**

Esto incluye:
- Letras de columna (`"B"`, `"$A$1"`)
- Nombres de funciones Excel (`IF`, `SUBSTITUTE`, `TRIM`, `COUNTIF`)
- El mapeo de columnas DSL a posiciones Excel (`*turno-dia-col-map*`)
- Estructuras JSON para el protocolo Python
- El conocimiento de que `turno-table` tiene `first-row=4, cell-height=2`

### Qué va aquí

- Métodos `compile-excel-formula` para cada clase de expresión
- Variables de contexto dinámico (`*sheet-env*`, `*turno-dia-col-map*`)
- Helpers de conversión (`col->letter`, `resolve-cross-col`)
- Métodos `generate-code` para serializar el workbook a Python

### Qué NO va aquí

- Clases del AST (van en `ast-def.lisp`)
- Macros del DSL (van en `dsl-directo.lisp`)

---

## Flujo completo — ejemplo `collect-over`

### 1. El usuario escribe el AST (`ast-facultad.lisp`)

```lisp
(def-table aulas-dia-table
  ((dia "") (aula1 "") ... (lab ""))
  :inst-params (dia)
  :computed
    ((aula1 (collect-over *grupos-all* (g)
               (_if (equals (cross-cell :sheet g :col dia :row (turno-aula-row))
                            (str "Aula 1"))
                    (concat (cross-cell :sheet g :col 'a :row 1) (str " "))
                    (str ""))))))

;; Instanciación — :dia lun se pasa como inst-param
(tabla aulas-dia-table :dia lun :data *DATOS-AULAS-LUNES*)
```

### 2. Los macros construyen el AST (`dsl-directo.lisp`)

```
collect-over  → xl-expr-collect-over { groups=("D111"...), sheet-var=G, body=... }
  _if         → xl-expr-if { test=..., then=..., else=... }
    equals    → xl-expr-equals { a=xl-expr-cross-cell, b=xl-expr-string }
    cross-cell→ xl-expr-cross-cell { sheet=G, xcol=LUN, row=xl-expr-turno-aula-row }
    concat    → xl-expr-concat { a=xl-expr-cross-cell{A,1}, b=xl-expr-string{" "} }
```

El AST contiene símbolos (`G`, `LUN`, `A`) — ningún string de Excel.

### 3. `compile-excel-formula` traduce al backend (`generate-code-direct.lisp`)

Para `row-num=2` (turno 1), `first-row=2`, iteración sobre `"D111"`:

```
xl-expr-collect-over
  → itera ["D111","D211",...,"M411"]
  → para "D111": *sheet-env* = ((G . "D111"))
    xl-expr-if
      xl-expr-equals
        xl-expr-cross-cell {G, LUN, xl-expr-turno-aula-row}
          → LUN resuelto via *turno-dia-col-map* → "B"
          → turno-aula-row: 4 + (2-2)*2 + 1 = 5
          → "D111!$B$5"
        xl-expr-string "Aula 1" → "\"Aula 1\""
      → "D111!$B$5=\"Aula 1\""
      xl-expr-concat {cross-cell{G,A,1}, str{" "}}
        → "(D111!$A$1&\" \")"
      xl-expr-string "" → "\"\""
    → "IF(D111!$B$5=\"Aula 1\",(D111!$A$1&\" \"),\"\")"
  → términos para 17 grupos concatenados con &
  → SUBSTITUTE(TRIM(IF(D111!...)&IF(D211!...)&...&IF(M411!...)), " ", ",")
```

### 4. Python genera el Excel

El JSON resultante se pasa a `hoja_con_formulas.py`, que usa `openpyxl`
para escribir la fórmula en cada celda de la hoja Aulas.

---

## Principios de implementación

### Al añadir una nueva expresión

1. **`ast-def.lisp`**: declarar la clase con sus slots (nombres de dominio).
2. **`dsl-directo.lisp`**: escribir el macro que construye el objeto.
   - Documentar la sintaxis con ejemplos.
   - Definir la política de quoting de cada argumento.
3. **`generate-code-direct.lisp`**: implementar `compile-excel-formula`.
   - Aquí sí van los strings de Excel.
   - Si la expresión necesita contexto del workbook, usar variables dinámicas.

### Al añadir un nuevo tipo de tabla

1. Un `def-table` en el archivo de workbook (ej. `ast-facultad.lisp`).
2. Si necesita parámetros que afectan las fórmulas: usar `:inst-params`.
3. No crear macros auxiliares en el archivo de workbook.
   Los macros van en `dsl-directo.lisp`; si son demasiado específicos del
   dominio, cuestionar si realmente son necesarios.

### Señales de que algo está mal ubicado

| Señal | Problema | Solución |
|---|---|---|
| String `"$B$5"` en `ast-def.lisp` | AST no es backend-agnóstico | Mover a generate-code |
| `defmacro` en `ast-facultad.lisp` | Macros fuera del DSL | Mover a dsl-directo o eliminar |
| Lógica condicional en un macro DSL | El macro hace más que construir un objeto | Separar en clase AST + método |
| Nombre de clase con concepto de dominio (`xl-expr-aula-lookup`) | Clase demasiado específica | Generalizar (`xl-expr-collect-over`) |

---

## Estructura de archivos

```
ast-def.lisp              Clases del AST — sin lógica
dsl-directo.lisp          Macros del DSL — sin strings de backend
generate-code-direct.lisp Backend Excel/Python — todo lo específico del target
ast-facultad.lisp         Workbook concreto — solo usa macros del DSL
gen-data-facultad.lisp    Genera data-facultad.lisp con los datos pre-procesados
data-facultad.lisp        Datos generados (no editar a mano)
hoja_con_formulas.py      Backend Python: recibe JSON, genera xlsx con openpyxl
```

---

## inst-params — parámetros de generación de código

Los `inst-params` son parámetros de la función de tabla que afectan la
**estructura de las fórmulas** (no son datos en celdas).

Se declaran en `def-table` con `:inst-params` y se pasan al instanciar con `tabla`:

```lisp
;; Declaración — `dia` es una variable disponible en :computed
(def-table aulas-dia-table
  (...)
  :inst-params (dia)
  :computed ((aula1 (collect-over *grupos* (g)
                      (_if (equals (cross-cell :sheet g :col dia ...) ...)
                           ...)))))

;; Instanciación — distintas tablas, un solo def-table
(tabla aulas-dia-table :dia lun :data *DATOS-AULAS-LUNES*)
(tabla aulas-dia-table :dia mar :data *DATOS-AULAS-MARTES*)
```

**No usar para** valores que van en celdas (eso es `:params`).
**Usar para** cualquier cosa que cambie la forma de la fórmula generada.
