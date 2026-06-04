;; =============================================================================
;; ast-def.lisp — Definiciones de todas las clases del AST
;;
;; RESPONSABILIDAD: declarar la ESTRUCTURA de los nodos del árbol de expresiones
;; y del árbol del workbook. Este archivo no contiene lógica de ningún tipo:
;; ni generación de código, ni interpretación, ni referencias a backends.
;;
;; PRINCIPIO CLAVE: backend-agnóstico.
;;   - No hay strings de Excel ("$A$1", "SUBSTITUTE", "B", "F"...).
;;   - No hay strings de Python ni de ningún otro target.
;;   - Solo nombres de dominio: columnas, hojas, expresiones, tablas.
;;
;; Las clases se definen con defclass* (de codigo-tesis.lisp), que genera:
;;   - clase-xl-<name>  como clase CLOS
;;   - xl-<name>        como función constructora (&key slots...)
;; =============================================================================

(load "codigo-tesis.lisp")

;; =============================================================================
;; EXPRESIONES — nodos que representan valores computados celda a celda.
;; Cada expresión es compilada por compile-excel-formula en generate-code-direct.
;; =============================================================================

;; --- Flujo de control ---

;; Condicional: IF(test, then, else)
(defclass* xl-expr-if () (test then else))

;; Verdadero si la expresión no es vacía/cero
(defclass* xl-expr-non-empty () (expr))

;; True en la primera fila de datos de la tabla (ROW()=first-row)
(defclass* xl-expr-first-row () ())

;; Referencia a la celda de la misma columna en la fila anterior/siguiente
(defclass* xl-expr-previous-row () (expr))
(defclass* xl-expr-next-row () (expr))

;; Produce "" (celda visualmente vacía)
(defclass* xl-expr-show-nothing () ())

;; --- Referencia a columnas de la tabla actual ---

;; Referencia a una columna por su nombre DSL, con fila actual o desfasada.
;; context: nil (fila actual), xl-expr-previous-row, xl-expr-next-row.
(defclass* xl-expr-column-ref () (name context))

;; Referencia a un parámetro de instancia (valor almacenado en celda auxiliar)
(defclass* xl-expr-param-ref () (name))

;; --- Comparaciones ---

(defclass* xl-expr-equals () (a b))
(defclass* xl-expr-different () (a b))
(defclass* xl-expr-and () (a b))
(defclass* xl-expr-or () (a b))

;; --- Aritmética ---

(defclass* xl-expr-add () (a b))
(defclass* xl-expr-subtract () (a b))
(defclass* xl-expr-multiply () (a b))
(defclass* xl-expr-divide () (a b))

;; --- Strings ---

;; Literal de texto: "valor"
(defclass* xl-expr-string () (value))

;; Concatenación: (a & b)
(defclass* xl-expr-concat () (a b))

;; --- Tiempo ---

;; Suma de tiempos como texto: TEXT(TIMEVALUE(a)+(b/1440),"hh:mm")
(defclass* xl-expr-time-add () (a b))

;; --- Rangos y agregados ---

;; Rango de columnas en la tabla actual (de from-col a to-col).
;; from-col y to-col son xl-expr-column-ref. to-col puede ser nil (rango único).
(defclass* xl-range () (from-col to-col))

;; COUNTIF sobre un rango de la tabla actual contra un criterio
(defclass* xl-expr-countif () (count-range criteria))

;; COUNTA sobre un rango de la tabla actual
(defclass* xl-expr-counta () (count-range))

;; SUM sobre un rango de la tabla actual
(defclass* xl-expr-sum () (count-range))

;; --- Lookups ---

;; VLOOKUP en la tabla actual: busca key-expr y devuelve value-field
(defclass* xl-expr-lookup () (value-field key-expr))

;; --- Referencias cross-sheet simples (hoja fija) ---

;; Referencia a una celda en otra hoja con plantilla fija: "Hoja!$C$5"
;; cell-template puede incluir ${row-num} para fila dinámica.
(defclass* xl-expr-cross-sheet-ref () (sheet cell-template))

;; =============================================================================
;; EXPRESIONES CROSS-SHEET DINÁMICAS — para búsquedas sobre múltiples hojas.
;;
;; A diferencia de xl-expr-cross-sheet-ref (hoja fija), estas expresiones
;; trabajan con una variable de hoja resuelta en tiempo de compilación
;; por collect-over a través del entorno dinámico *sheet-env*.
;; =============================================================================

;; Referencia a una celda en la hoja ligada a sheet en *sheet-env*.
;;   sheet : símbolo de la variable de hoja (ej. G), declarado en collect-over
;;   xcol  : símbolo de columna — resuelto por resolve-cross-col en generate-code:
;;             si está en *turno-dia-col-map* → letra Excel; si no → symbol-name
;;   row   : entero fijo O expresión (ej. xl-expr-turno-aula-row)
(defclass* xl-expr-cross-cell () (sheet xcol row))

;; Fila del aula correspondiente al turno actual dentro de la turno-table
;; de un grupo. Valor concreto calculado en compile-excel-formula:
;;   first-row-grupo + (row-num - first-row-tabla) * cell-height + aula-offset
;; Encapsula el conocimiento de la estructura interna de turno-table.
(defclass* xl-expr-turno-aula-row () ())

;; Combinator: aplica body sobre cada hoja de groups, ligando sheet-var
;; al nombre de cada hoja en *sheet-env*, y concatena los resultados en:
;;   SUBSTITUTE(TRIM(t1 & t2 & ... & tN), " ", ",")
;; El resultado es una lista de grupos que cumplen la condición en body.
;;   groups    : lista de strings (nombres de hojas)
;;   sheet-var : símbolo que actúa como variable de iteración
;;   body      : expresión AST que puede referenciar sheet-var vía cross-cell
(defclass* xl-expr-collect-over () (groups sheet-var body))

;; =============================================================================
;; ESTRUCTURA DE TABLA
;; =============================================================================

;; Definición de columna (no usado directamente en el DSL moderno)
(defclass* xl-col-def () (name display-name))

;; Referencia a celda absoluta en el workbook: fila, columna, hoja opcional
(defclass* xl-cell-ref () (row col sheet))

;; Fórmula en una celda fija (fuera del área de datos de la tabla)
(defclass* xl-fixed-formula () (cell-ref expr))

;; Regla de estilo condicional por columnas
(defclass* xl-style-rule () (rule-condition target-columns))

;; Tabla de datos con sus columnas, fórmulas computed, parámetros y formato.
;;   contenido    : lista de filas (data pura)
;;   headers      : lista de strings (cabecera visual de cada columna)
;;   col-names    : lista de símbolos (nombres DSL de columnas)
;;   computed     : lista de (col-sym . expr) — fórmulas por columna
;;   fixed-formulas: lista de xl-fixed-formula
;;   params       : lista de (sym . value) — parámetros de instancia
;;   first-row    : fila Excel donde empieza el área de datos (nil = 2)
;;   cell-height  : filas Excel por fila lógica (default 1)
;;   cell-width   : columnas Excel por columna lógica (default 1)
;;   paired-columns: columnas que se expanden en pares (asig/aula)
(defclass* xl-table () (id cols rows contenido headers computed col-names
                        params style-rules fixed-formulas
                        first-row cell-height cell-width paired-columns))

;; =============================================================================
;; ESTRUCTURA DEL WORKBOOK
;; =============================================================================

;; Región: agrupa tablas colocadas horizontalmente en la hoja
(defclass* xl-region () (tables))

;; Hoja: nombre + lista de regiones
(defclass* xl-sheet () (name regions))

;; Workbook: nombre del archivo + lista de hojas
(defclass* xl-workbook () (name sheets))

;; =============================================================================
;; BACKEND: objeto singleton que identifica el target de generación
;; =============================================================================

;; xl-out: target Excel/Python (el único backend implementado actualmente)
(defclass xl-out () ())
(defparameter xl-py (make-instance 'xl-out))

(provide "ast-def")
