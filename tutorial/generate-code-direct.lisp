;; =============================================================================
;; generate-code-direct.lisp — Backend Excel/Python del DSL
;;
;; RESPONSABILIDAD: traducir nodos del AST (definidos en ast-def.lisp) a
;; fórmulas y estructuras concretas del target (Excel/Python/openpyxl).
;; Este archivo es el ÚNICO lugar donde pueden vivir:
;;   - Letras de columna Excel ("B", "C", "$A$1"...)
;;   - Nombres de funciones Excel (IF, SUBSTITUTE, TRIM, COUNTIF...)
;;   - Estructuras JSON del protocolo hoja_con_formulas.py
;;   - Conocimiento de la estructura interna de las tablas del workbook
;;     (ej: turno-table tiene first-row=4, cell-height=2)
;;
;; PRINCIPIO CLAVE: separación de responsabilidades.
;;   - ast-def.lisp              → QUÉ existe (clases, estructura)
;;   - dsl-directo               → CÓMO se escribe (sintaxis, macros)
;;   - code-generation-utils.lisp → utilidades y variables de configuración
;;   - aquí                      → A QUÉ se traduce en el backend concreto
;;
;; El método central es compile-excel-formula, que recibe un nodo AST
;; y devuelve un string con la fórmula Excel correspondiente.
;;
;; Firma: (compile-excel-formula expr col-map data-names row-num first-row last-row)
;;   col-map    : alist (col-sym → "LetraExcel") de la región actual
;;   data-names : lista de símbolos de columnas de la tabla actual
;;   row-num    : fila Excel absoluta que se está compilando
;;   first-row  : primera fila de datos de la tabla
;;   last-row   : última fila de datos de la tabla
;; =============================================================================
(load (merge-pathnames "code-generation-utils.lisp" *load-truename*))

;; xl-sheet-fixed-expr nodes from the parent sheet; consumed by xl-region generator.
(defvar *sheet-fixed-expressions* nil)

;; Alist (table-id . (first-row . last-row)) for all tables in the current region.
;; Allows trange to use per-table row bounds instead of the main-table's bounds.
(defvar *table-data-rows* nil)

; =====================================================================
; COMPILE-EXCEL-FORMULA — expresión → string de fórmula Excel
; =====================================================================

(defgeneric compile-excel-formula (expr col-map data-names row-num first-row last-row)
  (:documentation "Compila un árbol de expresión a string de fórmula Excel"))

(defmethod compile-excel-formula ((e clase-xl-expr-if) col-map data-names row-num first-row last-row)
  (format nil "IF(~a,~a,~a)"
          (compile-excel-formula (test e) col-map data-names row-num first-row last-row)
          (compile-excel-formula (then e) col-map data-names row-num first-row last-row)
          (compile-excel-formula (else e) col-map data-names row-num first-row last-row)))

(defmethod compile-excel-formula ((e clase-xl-expr-non-empty) col-map data-names row-num first-row last-row)
  (format nil "~a<>\"\""
          (compile-excel-formula (expr e) col-map data-names row-num first-row last-row)))

(defmethod compile-excel-formula ((e clase-xl-expr-first-row) col-map data-names row-num first-row last-row)
  (format nil "ROW()=~a" first-row))

(defmethod compile-excel-formula ((e clase-xl-expr-column-ref) col-map data-names row-num first-row last-row)
  (let* ((letter (cdr (assoc (name e) col-map)))
         (ctx (context e)))
    (if (typep ctx 'clase-xl-expr-previous-row)
        (format nil "~a~a" letter (1- row-num))
        (format nil "~a~a" letter row-num))))

;; xl-expr-table-col-ref: resuelve la columna en el sub-mapa de la tabla indicada,
;; ignorando el resto de la región. Evita colisiones entre tablas vecinas.
(defmethod compile-excel-formula ((e clase-xl-expr-table-col-ref) col-map data-names row-num first-row last-row)
  (declare (ignore col-map data-names))
  (let* ((tmap   (cdr (assoc (table-id e) *table-col-maps*)))
         (letter (cdr (assoc (name e) tmap)))
         (ctx    (context e)))
    (if (typep ctx 'clase-xl-expr-previous-row)
        (format nil "~a~a" letter (1- row-num))
        (format nil "~a~a" letter row-num))))

;; xl-table-range: rango absoluto de columnas de una tabla específica.
;; Usa *table-data-rows* para obtener los límites exactos de la tabla indicada;
;; cae en first-row/last-row del contexto cuando no hay entrada en *table-data-rows*.
(defmethod compile-excel-formula ((e clase-xl-table-range) col-map data-names row-num first-row last-row)
  (declare (ignore col-map data-names row-num))
  (let* ((tmap         (cdr (assoc (table-id e) *table-col-maps*)))
         (from-letter  (cdr (assoc (name (from-col e)) tmap)))
         (to-letter    (if (to-col e)
                           (cdr (assoc (name (to-col e)) tmap))
                           from-letter))
         (tbl-rows     (cdr (assoc (table-id e) *table-data-rows*)))
         (actual-first (if tbl-rows (car tbl-rows) first-row))
         (actual-last  (if tbl-rows (cdr tbl-rows) last-row)))
    (format nil "$~a$~a:$~a$~a" from-letter actual-first to-letter actual-last)))

(defmethod compile-excel-formula ((e clase-xl-expr-param-ref) col-map data-names row-num first-row last-row)
  (declare (ignore col-map data-names row-num first-row last-row))
  (let* ((name (name e))
         (cell (cdr (assoc name *param-cells*))))
    (if cell
        cell
        (format nil "$~a" (string-upcase (symbol-name name))))))

(defmethod compile-excel-formula ((e clase-xl-expr-previous-row) col-map data-names row-num first-row last-row)
  (compile-excel-formula (expr e) col-map data-names (1- row-num) first-row last-row))

(defmethod compile-excel-formula ((e clase-xl-expr-next-row) col-map data-names row-num first-row last-row)
  (compile-excel-formula (expr e) col-map data-names (1+ row-num) first-row last-row))

(defmethod compile-excel-formula ((e clase-xl-expr-equals) col-map data-names row-num first-row last-row)
  (format nil "~a=~a"
          (compile-excel-formula (a e) col-map data-names row-num first-row last-row)
          (compile-excel-formula (b e) col-map data-names row-num first-row last-row)))

(defmethod compile-excel-formula ((e clase-xl-expr-different) col-map data-names row-num first-row last-row)
  (format nil "~a<>~a"
          (compile-excel-formula (a e) col-map data-names row-num first-row last-row)
          (compile-excel-formula (b e) col-map data-names row-num first-row last-row)))

(defmethod compile-excel-formula ((e clase-xl-expr-gt) col-map data-names row-num first-row last-row)
  (format nil "~a>~a"
          (compile-excel-formula (a e) col-map data-names row-num first-row last-row)
          (compile-excel-formula (b e) col-map data-names row-num first-row last-row)))

(defmethod compile-excel-formula ((e clase-xl-expr-lt) col-map data-names row-num first-row last-row)
  (format nil "~a<~a"
          (compile-excel-formula (a e) col-map data-names row-num first-row last-row)
          (compile-excel-formula (b e) col-map data-names row-num first-row last-row)))

(defmethod compile-excel-formula ((e clase-xl-expr-gte) col-map data-names row-num first-row last-row)
  (format nil "~a>=~a"
          (compile-excel-formula (a e) col-map data-names row-num first-row last-row)
          (compile-excel-formula (b e) col-map data-names row-num first-row last-row)))

(defmethod compile-excel-formula ((e clase-xl-expr-lte) col-map data-names row-num first-row last-row)
  (format nil "~a<=~a"
          (compile-excel-formula (a e) col-map data-names row-num first-row last-row)
          (compile-excel-formula (b e) col-map data-names row-num first-row last-row)))

(defmethod compile-excel-formula ((e clase-xl-expr-and) col-map data-names row-num first-row last-row)
  (format nil "AND(~a,~a)"
          (compile-excel-formula (a e) col-map data-names row-num first-row last-row)
          (compile-excel-formula (b e) col-map data-names row-num first-row last-row)))

(defmethod compile-excel-formula ((e clase-xl-expr-or) col-map data-names row-num first-row last-row)
  (format nil "OR(~a,~a)"
          (compile-excel-formula (a e) col-map data-names row-num first-row last-row)
          (compile-excel-formula (b e) col-map data-names row-num first-row last-row)))

;; =====================================================================
; COMPILE-EXCEL-FORMULA — ARITMÉTICA
; =====================================================================

(defmethod compile-excel-formula ((e clase-xl-expr-add) col-map data-names row-num first-row last-row)
  (format nil "(~a+~a)"
          (compile-excel-formula (a e) col-map data-names row-num first-row last-row)
          (compile-excel-formula (b e) col-map data-names row-num first-row last-row)))

(defmethod compile-excel-formula ((e clase-xl-expr-subtract) col-map data-names row-num first-row last-row)
  (format nil "(~a-~a)"
          (compile-excel-formula (a e) col-map data-names row-num first-row last-row)
          (compile-excel-formula (b e) col-map data-names row-num first-row last-row)))

(defmethod compile-excel-formula ((e clase-xl-expr-multiply) col-map data-names row-num first-row last-row)
  (format nil "(~a*~a)"
          (compile-excel-formula (a e) col-map data-names row-num first-row last-row)
          (compile-excel-formula (b e) col-map data-names row-num first-row last-row)))

(defmethod compile-excel-formula ((e clase-xl-expr-divide) col-map data-names row-num first-row last-row)
  (format nil "(~a/~a)"
          (compile-excel-formula (a e) col-map data-names row-num first-row last-row)
          (compile-excel-formula (b e) col-map data-names row-num first-row last-row)))

(defmethod compile-excel-formula ((e clase-xl-expr-promedio) col-map data-names row-num first-row last-row)
  (let* ((vals (cols e))
         (n (length vals))
         (sum-str (format nil "(~{~a~^+~})"
                         (loop for v in vals
                               collect (compile-excel-formula v col-map data-names row-num first-row last-row)))))
    (format nil "(~a/~a)" sum-str n)))

;; =====================================================================
; COMPILE-EXCEL-FORMULA — AGREGADOS
; =====================================================================

(defmethod compile-excel-formula ((e clase-xl-expr-countif) col-map data-names row-num first-row last-row)
  (let ((range-str (compile-excel-formula (count-range e) col-map data-names row-num first-row last-row))
        (criteria (compile-excel-formula (criteria e) col-map data-names row-num first-row last-row)))
    (format nil "COUNTIF(~a,~a)" range-str criteria)))

(defmethod compile-excel-formula ((e clase-xl-expr-counta) col-map data-names row-num first-row last-row)
  (let ((range-str (compile-excel-formula (count-range e) col-map data-names row-num first-row last-row)))
    (format nil "COUNTA(~a)" range-str)))

(defmethod compile-excel-formula ((e clase-xl-expr-sum) col-map data-names row-num first-row last-row)
  (let ((range-str (compile-excel-formula (count-range e) col-map data-names row-num first-row last-row)))
    (format nil "SUM(~a)" range-str)))

;;; xl-range: compila un rango de columnas a string Excel "$C$4:$L$9"
(defmethod compile-excel-formula ((e clase-xl-range) col-map data-names row-num first-row last-row)
  (declare (ignore data-names row-num))
  (let ((from-name (name (from-col e)))
        (to-name (when (to-col e) (name (to-col e)))))
    (let ((from-letter (cdr (assoc from-name col-map :test #'eq))))
      (let ((to-letter (if to-name
                           (cdr (assoc to-name col-map :test #'eq))
                           from-letter)))
        (format nil "$~a$~a:$~a$~a" from-letter first-row to-letter last-row)))))

;; =====================================================================
; COMPILE-EXCEL-FORMULA — CROSS-SHEET / STRING / CONCAT
; =====================================================================

(defmethod compile-excel-formula ((e clase-xl-expr-cross-sheet-ref) col-map data-names row-num first-row last-row)
  (let ((cell (resolve-range-template (cell-template e) first-row last-row row-num col-map)))
    (format nil "~a!~a" (sheet e) cell)))

(defmethod compile-excel-formula ((e clase-xl-expr-string) col-map data-names row-num first-row last-row)
  (declare (ignore col-map data-names row-num first-row last-row))
  (format nil "~s" (value e)))

(defmethod compile-excel-formula ((e clase-xl-expr-concat) col-map data-names row-num first-row last-row)
  (format nil "(~a&~a)"
          (compile-excel-formula (a e) col-map data-names row-num first-row last-row)
          (compile-excel-formula (b e) col-map data-names row-num first-row last-row)))

;; aula-lookup: genera SUBSTITUTE(TRIM(IF(G1!$COL$ROW=AULA,G1!$A$1&" ","")&...), " ", ",")
;; El ROW en la hoja de grupo se calcula desde row-num y first-row de la tabla Aulas.
;; Asume: grupo first-row=4, cell-height=2, aula es la 2da fila del par (offset=1).
;; =============================================================================
;; compile-excel-formula — CROSS-SHEET DINÁMICO
;; =============================================================================

;; xl-expr-source-row: fila física en la tabla identificada por table-id,
;; correspondiente a la fila lógica actual más el offset de subcelda.
;; Fórmula: src-first-row + (row-num - first-row) * cell-height + offset
;; Funciona para cualquier instancia de xl-expr-source-row (ej. turno-aula-row).
(defmethod compile-excel-formula ((e clase-xl-expr-source-row) col-map data-names row-num first-row last-row)
  (declare (ignore col-map data-names last-row))
  (let* ((schema        (cdr (assoc (table-id e) *source-table-schemas*)))
         (src-first-row (getf schema :first-row))
         (ch            (getf schema :cell-height)))
    (format nil "~a" (+ src-first-row (* (- row-num first-row) ch) (offset e)))))

;; xl-expr-sheet-id: identificador canónico de la hoja.
;; Resuelve la variable de hoja desde *sheet-env* y genera la referencia
;; a la celda A1 de esa hoja, que por convención almacena el nombre del grupo.
(defmethod compile-excel-formula ((e clase-xl-expr-sheet-id) col-map data-names row-num first-row last-row)
  (declare (ignore col-map data-names row-num first-row last-row))
  (let* ((sheet-sym    (sheet e))
         (actual-sheet (cdr (assoc sheet-sym *sheet-env* :test #'string-equal))))
    (format nil "~a!$A$1" actual-sheet)))

;; xl-expr-cross-cell: referencia a celda en la hoja ligada a sheet en *sheet-env*.
;; 1. Resuelve la hoja: busca (sheet e) en *sheet-env* → string "D111"
;; 2. Resuelve la columna: resolve-cross-col sobre (xcol e)
;; 3. Resuelve la fila: entero → directo; expresión → compila recursivamente
;; Genera: "D111!$B$5"
(defmethod compile-excel-formula ((e clase-xl-expr-cross-cell) col-map data-names row-num first-row last-row)
  (let* ((sheet-sym    (sheet e))
         (actual-sheet (cdr (assoc sheet-sym *sheet-env* :test #'string-equal)))
         (col-letter   (resolve-cross-col (xcol e)))
         (row-val      (let ((r (row e)))
                         (if (integerp r)
                             (format nil "~a" r)
                             (compile-excel-formula r col-map data-names row-num first-row last-row)))))
    (format nil "~a!$~a$~a" actual-sheet col-letter row-val)))

;; xl-expr-collect-over: para cada hoja en groups, liga sheet-var→hoja en *sheet-env*
;; y compila body. Los N términos se concatenan y se envuelven en:
;;   SUBSTITUTE(TRIM(t1&t2&...&tN), " ", ",")
;;
;; El resultado es una fórmula Excel que devuelve los grupos que cumplen
;; la condición del body, separados por comas. Vacío si ninguno cumple.
(defmethod compile-excel-formula ((e clase-xl-expr-collect-over) col-map data-names row-num first-row last-row)
  (let* ((groups  (groups e))
         (sh-var  (sheet-var e))
         (body    (body e))
         (terms   (mapcar (lambda (g)
                            ;; Liga SH-VAR → g solo durante la compilación de este término
                            (let ((*sheet-env* (acons sh-var g *sheet-env*)))
                              (compile-excel-formula body col-map data-names row-num first-row last-row)))
                          groups))
         (inner   (format nil "~{~a~^&~}" terms)))
    (format nil "SUBSTITUTE(TRIM(~a),\" \",\",\")" inner)))

(defmethod compile-excel-formula ((e clase-xl-expr-time-add) col-map data-names row-num first-row last-row)
  (let* ((a-str (compile-excel-formula (a e) col-map data-names row-num first-row last-row))
         (b-str (compile-excel-formula (b e) col-map data-names row-num first-row last-row)))
    (format nil "TEXT(TIMEVALUE(~a)+(~a/1440),\"hh:mm\")" a-str b-str)))

(defmethod compile-excel-formula ((e clase-xl-expr-show-nothing) col-map data-names row-num first-row last-row)
  (declare (ignore col-map data-names row-num first-row last-row))
  "\"\"")

(defmethod compile-excel-formula ((e clase-xl-expr-empty) col-map data-names row-num first-row last-row)
  (declare (ignore col-map data-names row-num first-row last-row))
  "\"\"")

(defmethod compile-excel-formula ((e clase-xl-expr-lookup) col-map data-names row-num first-row last-row)
  (let* ((key (compile-excel-formula (key-expr e) col-map data-names row-num first-row last-row))
         (field-name (value-field e))
         (field-col (1+ (position field-name data-names)))
         (range-end-row last-row)
         (range-start-col (col->letter 1))
         (range-end-col (col->letter (length data-names))))
    (format nil "IFERROR(VLOOKUP(~a,$~a~a:$~a~a,~a,FALSE),0)"
            key range-start-col first-row range-end-col range-end-row field-col)))

; =====================================================================
; CF COLOR PALETTE — colores para conditional_formats de comparación
; Asignados por índice de regla (0, 1, 2 …). El backend garantiza que
; nunca habrá más reglas que entradas disponibles en la paleta.
; =====================================================================

(defparameter *cf-comparison-palette*
  '("FF4444"   ;; rojo  — regla 0: asignadas > frec (exceso)
    "44BB44"   ;; verde — regla 1: asignadas = frec (exacto)
    "4488FF"   ;; azul  — regla 2: asignadas < frec (déficit)
    "FF9900"   ;; naranja — regla 3
    "AA44FF"   ;; violeta — regla 4
    "FF44BB")) ;; rosa    — regla 5

; =====================================================================
; GENERATE-CODE — EXPRESIONES (JSON para compatibilidad)
; =====================================================================

(defmethod generate-code ((e clase-xl-expr-if) (lang xl-out) (stream t))
  (format stream "{\"type\": \"if\", \"condition\": ")
  (generate-code (test e) lang stream)
  (format stream ", \"then\": ")
  (generate-code (then e) lang stream)
  (format stream ", \"else\": ")
  (generate-code (else e) lang stream)
  (format stream "}"))

(defmethod generate-code ((e clase-xl-expr-non-empty) (lang xl-out) (stream t))
  (format stream "{\"type\": \"non-empty\", \"expr\": ")
  (generate-code (expr e) lang stream)
  (format stream "}"))

(defmethod generate-code ((e clase-xl-expr-first-row) (lang xl-out) (stream t))
  (format stream "{\"type\": \"first-row\"}"))

(defmethod generate-code ((e clase-xl-expr-column-ref) (lang xl-out) (stream t))
  (let ((ctx (context e)))
    (if ctx
        (progn
          (format stream "{\"type\": \"column-ref\", \"name\": \"~a\", \"context\": " (name e))
          (generate-code ctx lang stream)
          (format stream "}"))
        (format stream "{\"type\": \"column-ref\", \"name\": \"~a\"}" (name e)))))

(defmethod generate-code ((e clase-xl-expr-param-ref) (lang xl-out) (stream t))
  (format stream "{\"type\": \"param-ref\", \"name\": \"~a\"}" (name e)))

(defmethod generate-code ((e clase-xl-expr-previous-row) (lang xl-out) (stream t))
  (format stream "{\"type\": \"previous-row\", \"expr\": ")
  (generate-code (expr e) lang stream)
  (format stream "}"))

(defmethod generate-code ((e clase-xl-expr-time-add) (lang xl-out) (stream t))
  (format stream "{\"type\": \"time-add\", \"a\": ")
  (generate-code (a e) lang stream)
  (format stream ", \"b\": ")
  (generate-code (b e) lang stream)
  (format stream "}"))

(defmethod generate-code ((e clase-xl-expr-show-nothing) (lang xl-out) (stream t))
  (format stream "{\"type\": \"show-nothing\"}"))

(defmethod generate-code ((e clase-xl-expr-empty) (lang xl-out) (stream t))
  (format stream "{\"type\": \"empty\"}"))

(defmethod generate-code ((e clase-xl-expr-lookup) (lang xl-out) (stream t))
  (format stream "{\"type\": \"lookup\", \"field\": \"~a\", \"key\": " (value-field e))
  (generate-code (key-expr e) lang stream)
  (format stream "}"))

(defmethod generate-code ((e clase-xl-expr-next-row) (lang xl-out) (stream t))
  (format stream "{\"type\": \"next-row\", \"expr\": ")
  (generate-code (expr e) lang stream)
  (format stream "}"))

(defmethod generate-code ((e clase-xl-expr-equals) (lang xl-out) (stream t))
  (format stream "{\"type\": \"equals\", \"a\": ")
  (generate-code (a e) lang stream)
  (format stream ", \"b\": ")
  (generate-code (b e) lang stream)
  (format stream "}"))

(defmethod generate-code ((e clase-xl-expr-different) (lang xl-out) (stream t))
  (format stream "{\"type\": \"different\", \"a\": ")
  (generate-code (a e) lang stream)
  (format stream ", \"b\": ")
  (generate-code (b e) lang stream)
  (format stream "}"))

(defmethod generate-code ((e clase-xl-expr-and) (lang xl-out) (stream t))
  (format stream "{\"type\": \"and\", \"a\": ")
  (generate-code (a e) lang stream)
  (format stream ", \"b\": ")
  (generate-code (b e) lang stream)
  (format stream "}"))

(defmethod generate-code ((e clase-xl-expr-or) (lang xl-out) (stream t))
  (format stream "{\"type\": \"or\", \"a\": ")
  (generate-code (a e) lang stream)
  (format stream ", \"b\": ")
  (generate-code (b e) lang stream)
  (format stream "}"))

;; =====================================================================
; GENERATE-CODE — ARITMÉTICA
; =====================================================================

(defmethod generate-code ((e clase-xl-expr-add) (lang xl-out) (stream t))
  (format stream "{\"type\": \"add\", \"a\": ")
  (generate-code (a e) lang stream)
  (format stream ", \"b\": ")
  (generate-code (b e) lang stream)
  (format stream "}"))

(defmethod generate-code ((e clase-xl-expr-subtract) (lang xl-out) (stream t))
  (format stream "{\"type\": \"subtract\", \"a\": ")
  (generate-code (a e) lang stream)
  (format stream ", \"b\": ")
  (generate-code (b e) lang stream)
  (format stream "}"))

(defmethod generate-code ((e clase-xl-expr-multiply) (lang xl-out) (stream t))
  (format stream "{\"type\": \"multiply\", \"a\": ")
  (generate-code (a e) lang stream)
  (format stream ", \"b\": ")
  (generate-code (b e) lang stream)
  (format stream "}"))

(defmethod generate-code ((e clase-xl-expr-divide) (lang xl-out) (stream t))
  (format stream "{\"type\": \"divide\", \"a\": ")
  (generate-code (a e) lang stream)
  (format stream ", \"b\": ")
  (generate-code (b e) lang stream)
  (format stream "}"))

(defmethod generate-code ((e clase-xl-expr-promedio) (lang xl-out) (stream t))
  (format stream "{\"type\": \"promedio\", \"values\": [")
  (loop for v in (cols e) for i from 0
        do (when (> i 0) (format stream ", "))
           (generate-code v lang stream))
  (format stream "]}"))

;; =====================================================================
; GENERATE-CODE — AGREGADOS
; =====================================================================

(defmethod generate-code ((e clase-xl-expr-countif) (lang xl-out) (stream t))
  (format stream "{\"type\": \"countif\", \"range\": ")
  (generate-code (count-range e) lang stream)
  (format stream ", \"criteria\": ")
  (generate-code (criteria e) lang stream)
  (format stream "}"))

(defmethod generate-code ((e clase-xl-expr-counta) (lang xl-out) (stream t))
  (format stream "{\"type\": \"counta\", \"range\": ")
  (generate-code (count-range e) lang stream)
  (format stream "}"))

(defmethod generate-code ((e clase-xl-expr-sum) (lang xl-out) (stream t))
  (format stream "{\"type\": \"sum\", \"range\": ")
  (generate-code (count-range e) lang stream)
  (format stream "}"))

(defmethod generate-code ((e clase-xl-range) (lang xl-out) (stream t))
  (format stream "{\"type\": \"range\", \"from\": ~s" (name (from-col e)))
  (when (to-col e)
    (format stream ", \"to\": ~s" (name (to-col e))))
  (format stream "}"))

;; =====================================================================
; GENERATE-CODE — CROSS-SHEET, STRING, CONCAT
; =====================================================================

(defmethod generate-code ((e clase-xl-expr-cross-sheet-ref) (lang xl-out) (stream t))
  (format stream "{\"type\": \"cross-sheet\", \"sheet\": ~s, \"cell\": ~s}"
          (sheet e) (cell-template e)))

(defmethod generate-code ((e clase-xl-expr-string) (lang xl-out) (stream t))
  (format stream "{\"type\": \"string\", \"value\": ~s}" (value e)))

(defmethod generate-code ((e clase-xl-expr-concat) (lang xl-out) (stream t))
  (format stream "{\"type\": \"concat\", \"a\": ")
  (generate-code (a e) lang stream)
  (format stream ", \"b\": ")
  (generate-code (b e) lang stream)
  (format stream "}"))

; =====================================================================
; GENERATE-CODE — TABLAS
; =====================================================================

(defmethod generate-code ((tbl clase-xl-table) (lang xl-out) (stream t))
  (let* ((logical-con (contenido tbl))
         (con (expand-table-content tbl))
         (hdrs (headers tbl))
         (comp (computed tbl))
         (col-map (build-col-map tbl))
         (dnames (data-col-names tbl))
         (prms (params tbl))
         (cn (col-names tbl))
         (ffs (fixed-formulas tbl))
         (num-cols (length (col-names tbl)))
         (num-params (length (params tbl)))
         (cell-height (max 1 (or (cell-height tbl) 1)))
         (cell-width (max 1 (or (cell-width tbl) 1)))
         (key-count 0)
         (explicit-first (or (first-row tbl) 4)))
    (flet ((emit-sep ()
              (when (> key-count 0) (format stream ",~%"))))
      ;; ── data ──
      (when con
        (emit-sep)
        (format stream "        \"data\": ")
        (xl-write con stream)
        (incf key-count))
      ;; ── params ──
      (when prms
        (emit-sep)
        (format stream "        \"params\": {")
        (loop for (n . v) in prms for idx from 1
              for col-num = (+ num-cols idx)
              for cell = (format nil "~a1" (col->letter col-num))
              for i from 0
              do (when (> i 0) (format stream ", "))
                 (format stream "~s: ~s" cell v))
        (format stream "}")
        (incf key-count))
      ;; ── headers ──
      (when hdrs
        (emit-sep)
        (format stream "        \"headers\": ")
        (xl-write hdrs stream)
        (incf key-count))
      ;; ── formulas ──
      (let* ((first-row explicit-first)
             (prefix-rows (max 0 (1- first-row)))
             (logical-data-rows (max 0 (- (length logical-con) prefix-rows)))
             (physical-data-rows (* logical-data-rows cell-height))
             (last-row (if (> physical-data-rows 0)
                           (1- (+ first-row physical-data-rows))
                           (1- first-row)))
             (param-cells
                (when prms
                  (loop for (n . v) in prms for idx from 1
                        for col-num = (+ num-cols idx)
                        collect (cons n (format nil "$~a$1" (col->letter col-num))))))
        (when (or comp ffs)
          (emit-sep)
          (format stream "        \"formulas\": [")
          (let ((formula-count 0))
            ;; Computed formulas (per-row)
                  (loop for (col . expr) in comp
                   for col-index = (1+ (position col cn :test #'string-equal))
                   do
                      (loop for i from 0 below logical-data-rows
                             for row = (+ first-row (* i cell-height))
                             for formula = (let ((*param-cells* param-cells))
                                             (compile-excel-formula expr col-map dnames row first-row last-row))
                            do
                              (when (> formula-count 0) (format stream ", "))
                              (format stream "{\"row\": ~a, \"col\": ~a, \"value\": \"=~a\"}"
                                      row col-index (escape-python-string formula))
                              (incf formula-count)))
            ;; Fixed formulas (single cells)
            (loop for ff in ffs
                  for cr = (cell-ref ff)
                  for formula = (let ((*param-cells* param-cells))
                                  (compile-excel-formula (expr ff) col-map dnames (row cr) first-row last-row))
                  do
                     (when (> formula-count 0) (format stream ", "))
(format stream "{\"row\": ~a, \"col\": ~a, \"value\": \"=~a\"}"
        (row cr) (slot-value cr 'col) (escape-python-string formula))
                       (incf formula-count)))
           (format stream "]")
           (incf key-count)))
      ;; ── table_ranges / table_block_sizes (layout por backend) ──
      (let* ((first-row explicit-first)
             (prefix-rows (max 0 (1- first-row)))
             (logical-data-rows (max 0 (- (length logical-con) prefix-rows)))
             (physical-data-rows (* logical-data-rows cell-height))
             (last-row (if (> physical-data-rows 0)
                           (1- (+ first-row physical-data-rows))
                           nil))
             (last-col-letter (col->letter num-cols))
             (table-range (when last-row
                            (format nil "A~a:~a~a" first-row last-col-letter last-row))))
        (when table-range
          (emit-sep)
          (format stream "        \"table_ranges\": [~s]" table-range)
          (incf key-count)
          (when (or (> cell-height 1) (> cell-width 1))
            (emit-sep)
            (format stream "        \"table_block_sizes\": [{\"range\": ~s, \"row_step\": ~a, \"col_step\": ~a}]"
                    table-range cell-height cell-width)
            (incf key-count))))
      ;; ── column_widths ──
      (progn
        (emit-sep)
        (format stream "        \"column_widths\": {")
        (loop for (name . letter) in (build-col-map tbl) for idx from 1
              for width = (case (if (symbolp name) name (intern (string-upcase name)))
                            ((programa-calc) 30)
                            ((duracion) 8)
                            ((tipo) 10)
                            ((hora-inicio) 10)
                            ((hora-terminacion) 10)
                            ((tipo-calc) 8)
                            (otherwise 10))
              do (when (> idx 1) (format stream ", "))
                 (format stream "~a: ~a" idx width))
        (loop for (n . v) in prms for idx from (+ 1 num-cols)
              do (format stream ", ~a: 8" idx))
        (format stream "}")
        (incf key-count))
      ;; ── border_color ──
      (progn
        (emit-sep)
        (format stream "        \"border_color\": \"4F81BD\"")
        (incf key-count)
        (emit-sep)
        (format stream "        \"border_style\": \"thick\"")
        (incf key-count))
      ;; ── range_styles ──
      (let* ((first-row explicit-first)
             (data-rows (max 0 (- (length con) (1- first-row))))
             (last-row (1- (+ first-row data-rows)))
             (items ()))
        (when comp
          (setf items (nconc items (collect-range-styles-from-tipo tbl first-row last-row))))
        (let ((static-rules (remove-if (lambda (r)
                                          (typep (rule-condition r)
                                                 '(or clase-xl-expr-exists
                                                      clase-xl-expr-gt  clase-xl-expr-lt
                                                      clase-xl-expr-gte clase-xl-expr-lte)))
                                       (or (style-rules tbl) nil))))
          (when static-rules
            (setf items (nconc items (collect-range-styles-from-rules tbl static-rules first-row last-row)))))
        (when items
          (emit-sep)
          (format stream "        \"range_styles\": [")
          (loop for (range . style) in items for i from 0
                do (when (> i 0) (format stream ", "))
                   (format stream "{\"range\": ~s, \"style\": {\"bg_color\": ~s}}"
                           range style))
          (format stream "]")
           (incf key-count)))))))

; =====================================================================
; GENERATE-CODE — REGIONES
; =====================================================================

(defmethod generate-code ((region clase-xl-region) (lang xl-out) (stream t))
  (let ((tables (tables region)))
    (unless tables (return-from generate-code (format stream "{}")))
    (labels ((col-count (tbl) (length (col-names tbl)))
             (effective-first-row (tbl)
               (or (first-row tbl) 4))
             (logical-rows (tbl)
               (let* ((con (contenido tbl))
                      (ef (first-row tbl))
                      (pr (if ef (max 0 (1- ef)) 0)))
                 (max 0 (- (length con) pr))))
             (physical-rows (tbl)
               (* (logical-rows tbl) (max 1 (or (cell-height tbl) 1))))
             (tbl-last-row (tbl)
               (let* ((fr (effective-first-row tbl))
                      (pr (physical-rows tbl)))
                 (if (> pr 0) (1- (+ fr pr)) (1- fr))))
             (build-global-col-map (tbls)
               (let ((cur 1) (map nil))
                 (dolist (tb tbls (nreverse map))
                   (dolist (name (col-names tb))
                     (push (cons name (col->letter cur)) map)
                     (incf cur))
                   (incf cur)))))  ;; skip gap column
      (let* ((offsets (loop with cur = 0
                             for tbl in tables
                             for start = cur
                             do (incf cur (+ 1 (col-count tbl) (length (params tbl))))
                             collect start))
             (main-tbl (first tables))
             (ref-first-row (effective-first-row main-tbl))
             (ref-physical-rows (physical-rows main-tbl))
             (last-row (if (> ref-physical-rows 0)
                           (1- (+ ref-first-row ref-physical-rows))
                           (1- ref-first-row)))
              (total-cols (loop for tbl in tables sum (+ 1 (col-count tbl) (length (params tbl)))))
             (last-col-letter (col->letter total-cols))
             (global-col-map (build-global-col-map tables))
             (*table-col-maps* (build-table-col-maps tables offsets))
             (*table-data-rows* (loop for tbl in tables
                                       collect (cons (id tbl)
                                                     (cons (effective-first-row tbl)
                                                           (tbl-last-row tbl)))))
             (global-dnames (loop for tbl in tables append (data-col-names tbl)))
             (all-expanded (loop for tbl in tables collect (expand-table-content tbl)))
             (max-rows (loop for exp in all-expanded maximize (length exp))))
       (flet ((resolve-nav (nav-node)
                (let* ((anchor  (anchor nav-node))
                       (steps   (steps nav-node))
                       (tid     (table-id anchor))
                       (cname   (col-name anchor))
                       (atype   (anchor-type anchor))
                       (pair    (loop for tbl in tables for o in offsets
                                      when (eq (id tbl) tid) return (cons tbl o)))
                       (tbl     (car pair))
                       (offset  (cdr pair)))
                  (when tbl
                    (let* ((cidx  (position cname (col-names tbl) :test #'string-equal))
                           (pcol  (+ offset cidx 1))
                           (trows (cdr (assoc (id tbl) *table-data-rows*)))
                           (fr    (car trows))
                           (lr    (cdr trows))
                           (arow  (ecase atype
                                    (:ultima-fila  lr)
                                    (:primera-fila fr))))
                      (let ((row arow) (col pcol))
                        (loop for (dir . n) in steps do
                          (case dir
                            (abajo     (incf row n))
                            (arriba    (decf row n))
                            (derecha   (incf col n))
                            (izquierda (decf col n))))
                        (list row col)))))))
        (format stream "{~%")
        ;; ── data ──
        ;; For tables with params: prepend a param row so $G2 (param cell) has
        ;; the initial value. The separator column doubles as the param column,
        ;; so r=0 writes the param value there instead of the usual "".
        (format stream "        \"data\": ")
        (xl-write
          (loop for r from 0 below max-rows
                collect (loop for tbl in tables for exp in all-expanded
                              for nc = (col-count tbl)
                              for nprms = (length (params tbl))
                              for row = (if (< r (length exp))
                                            (nth r exp)
                                            (make-list nc :initial-element ""))
                 append (subseq row 0 (min (length row) nc))
                 append (make-list (1+ nprms) :initial-element "")))
          stream)
        (format stream ",~%")
        ;; ── headers ──
        (format stream "        \"headers\": ")
        (xl-write (loop for tbl in tables append (append (headers tbl) (list ""))) stream)
        (format stream ",~%")
        ;; ── params ──
        (let ((param-entries nil))
          (loop for tbl in tables
                for offset in offsets
                for nc = (col-count tbl)
                for prms = (params tbl)
                do (loop for (n . v) in prms for idx from 1
                         for col-num = (+ offset nc 1 idx)
                         do (push (cons (format nil "~a1" (col->letter col-num)) v) param-entries)))
          (when param-entries
            (format stream "        \"params\": {")
            (loop for (cell . value) in (nreverse param-entries) for i from 0
                  do (when (> i 0) (format stream ", "))
                     (format stream "~s: ~s" cell value))
            (format stream "},~%")))
        ;; ── formulas ──
        (let ((formulas nil))
          (loop for tbl in tables
                for offset in offsets
                for cell-h = (max 1 (or (cell-height tbl) 1))
                for first-row = (effective-first-row tbl)
                for tbl-lr = (tbl-last-row tbl)
                for ldr = (logical-rows tbl)
                for prms = (params tbl)
                for param-cells = (when prms
                                    (loop for (n . v) in prms for idx from 1
                                          for col-num = (+ (col-count tbl) 1 idx)
                                          collect (cons n (format nil "$~a$1" (col->letter (+ offset col-num))))))
                do
                ;; computed formulas (per-row)
                (loop for (col . expr) in (computed tbl)
                      for tbl-col-idx = (position col (col-names tbl) :test #'string-equal)
                      for abs-col = (+ offset (or tbl-col-idx -1) 1)
                      do
                (loop for i from 0 below ldr
                      for row = (+ first-row (* i cell-h))
                            for formula = (let ((*param-cells* param-cells))
                                            (compile-excel-formula expr global-col-map global-dnames row first-row tbl-lr))
                            do (push (list row abs-col formula) formulas))))
          ;; sheet-level fixed expressions (resolved via nav anchors)
          (let ((region-ids (mapcar #'id tables)))
            (dolist (fe *sheet-fixed-expressions*)
              (let* ((nav-node (pos fe))
                     (atid     (table-id (anchor nav-node))))
                (when (member atid region-ids)
                  (let* ((cell-pos  (resolve-nav nav-node))
                         (fe-rows   (cdr (assoc atid *table-data-rows*)))
                         (fe-first  (car fe-rows))
                         (fe-last   (cdr fe-rows))
                         (formula   (let ((*param-cells* nil))
                                      (compile-excel-formula
                                        (expr fe) global-col-map global-dnames
                                        (first cell-pos) fe-first fe-last))))
                    (when cell-pos
                      (push (list (first cell-pos) (second cell-pos) formula)
                            formulas)))))))
          (setf formulas (sort formulas (lambda (a b)
                                          (or (< (first a) (first b))
                                              (and (= (first a) (first b))
                                                   (< (second a) (second b)))))))
          (when formulas
            (format stream "        \"formulas\": [")
            (loop for (row col formula) in formulas for i from 0
                  do (when (> i 0) (format stream ", "))
                     (format stream "{\"row\": ~a, \"col\": ~a, \"value\": \"=~a\"}"
                             row col (escape-python-string formula)))
            (format stream "],~%")))
        ;; ── range_styles / conditional_formats ──
        ;; El backend decide el mecanismo según el tipo del nodo condición:
        ;;   xl-expr-exists     → FormulaRule SUMPRODUCT (recalcula en Excel)
        ;;   xl-expr-gt/lt/etc. → FormulaRule row-relative; color from *cf-comparison-palette*
        ;;   cualquier otro     → colores estáticos baked en tiempo de generación
        (let ((all-styles nil)
              (all-cf nil)
              (cf-cmp-idx 0))
          (loop for tbl in tables
                for offset in offsets
                for tbl-first = (effective-first-row tbl)
                for tbl-last  = (tbl-last-row tbl)
                for tbl-letter-map = (loop for name in (col-names tbl)
                                           for idx from (1+ offset)
                                           collect (cons name (col->letter idx)))
                for prms = (params tbl)
                for param-cells = (when prms
                                    (loop for (n . v) in prms for idx from 1
                                          for col-num = (+ (col-count tbl) 1 idx)
                                          collect (cons n (format nil "$~a$1" (col->letter (+ offset col-num))))))
                do (dolist (rule (or (style-rules tbl) nil))
                     (let ((cond-node (rule-condition rule)))
                       (cond
                         ((typep cond-node 'clase-xl-expr-exists)
                          ;; CF path — SUMPRODUCT FormulaRule
                          (let* ((exists-node  cond-node)
                                 (domain-tid   (table-id (domain exists-node)))
                                 (targets      (target-columns rule))
                                 (target-ltrs  (remove nil
                                                (loop for col in targets
                                                      collect (cdr (assoc col tbl-letter-map
                                                                          :test #'string-equal)))))
                                 (domain-pair  (when domain-tid
                                                (loop for t2 in tables for o2 in offsets
                                                      when (eq (id t2) domain-tid)
                                                      return (cons t2 o2))))
                                 (domain-lmap  (when domain-pair
                                                (loop for name in (col-names (car domain-pair))
                                                      for idx from (1+ (cdr domain-pair))
                                                      collect (cons name (col->letter idx)))))
                                 (domain-rows  (when domain-tid
                                                (cdr (assoc domain-tid *table-data-rows*))))
                                 (formula      (compile-exists-to-cf-formula
                                                exists-node tbl-letter-map tbl-first tbl-last
                                                :domain-letter-map domain-lmap
                                                :domain-first-row  (when domain-rows (car domain-rows))
                                                :domain-last-row   (when domain-rows (cdr domain-rows))
                                                :self-col-letter   (first target-ltrs)))
                                 (range-str    (when target-ltrs
                                                (format nil "~a~a:~a~a"
                                                        (first target-ltrs) tbl-first
                                                        (car (last target-ltrs)) tbl-last))))
                            (when (and formula range-str)
                              (push (list :range range-str :formula formula
                                          :style "\"font_color\": \"#FF0000\"")
                                    all-cf))))
                         ((typep cond-node '(or clase-xl-expr-gt  clase-xl-expr-lt
                                                 clase-xl-expr-gte clase-xl-expr-lte
                                                 clase-xl-expr-equals clase-xl-expr-different))
                          ;; CF path — row-relative FormulaRule; color from palette.
                          ;; One entry per target column so Excel doesn't shift the
                          ;; formula's column references across a multi-column range.
                          (let* ((targets      (target-columns rule))
                                 (target-ltrs  (remove nil
                                                (loop for col in targets
                                                      collect (cdr (assoc col tbl-letter-map
                                                                          :test #'string-equal)))))
                                 (formula      (let ((*param-cells* param-cells))
                                                 (compile-excel-formula
                                                  cond-node tbl-letter-map nil
                                                  tbl-first tbl-first tbl-last)))
                                 (color        (nth cf-cmp-idx *cf-comparison-palette*)))
                            (when (and formula target-ltrs)
                              (dolist (ltr target-ltrs)
                                (push (list :range  (format nil "~a~a:~a~a" ltr tbl-first ltr tbl-last)
                                            :formula formula
                                            :style   (format nil "\"bg_color\": \"~a\"" color))
                                      all-cf))
                              (incf cf-cmp-idx))))
                         (t
                          ;; static path — colores generados row a row
                          (let ((raw-styles (collect-range-styles-from-rules
                                              tbl (list rule) tbl-first tbl-last)))
                            (dolist (rs raw-styles)
                              (let* ((range (car rs))
                                     (color (cdr rs))
                                     (col-end (position-if (lambda (c) (digit-char-p c)) range))
                                     (tbl-letter (subseq range 0 col-end))
                                     (rest-str  (subseq range col-end))
                                     (tbl-col-num (loop for i from 0 below (length tbl-letter)
                                                        sum (* (- (char-code (char tbl-letter i)) 64)
                                                               (expt 26 (- (length tbl-letter) i 1)))))
                                     (abs-col-num (+ tbl-col-num offset))
                                     (abs-letter (col->letter abs-col-num))
                                     (adjusted (let* ((colon (position #\: rest-str))
                                                      (row1  (subseq rest-str 0 colon))
                                                      (tail  (subseq rest-str (1+ colon)))
                                                      (col2-end (position-if #'digit-char-p tail))
                                                      (row2 (subseq tail col2-end)))
                                                 (format nil "~a~a:~a~a"
                                                         abs-letter row1 abs-letter row2))))
                                (push (cons adjusted color) all-styles)))))))))
          (when all-styles
            (format stream "        \"range_styles\": [")
            (loop for (range . color) in (nreverse all-styles) for i from 0
                  do (when (> i 0) (format stream ", "))
                     (format stream "{\"range\": ~s, \"style\": {\"font_color\": ~s}}"
                             range color))
            (format stream "],~%"))
          (when all-cf
            (format stream "        \"conditional_formats\": [")
            (loop for entry in (nreverse all-cf) for i from 0
                  do (when (> i 0) (format stream ", "))
                     (format stream "{\"range\": ~s, \"formula\": ~s, \"style\": {~a}}"
                             (getf entry :range) (getf entry :formula) (getf entry :style)))
            (format stream "],~%")))
        ;; ── table_ranges ──
        (let ((global-last-row (loop for tbl in tables maximize (tbl-last-row tbl))))
          (format stream "        \"table_ranges\": [~s],~%"
                  (format nil "A~a:~a~a" ref-first-row last-col-letter global-last-row)))
        ;; ── table_block_sizes ──
        (format stream "        \"table_block_sizes\": [")
        (loop for tbl in tables
              for offset in offsets
              for cell-h = (max 1 (or (cell-height tbl) 1))
              for cell-w = (max 1 (or (cell-width tbl) 1))
              for from-letter = (col->letter (1+ offset))
              for to-letter = (col->letter (+ offset (col-count tbl)))
              for tbl-fr = (effective-first-row tbl)
              for tbl-lr = (tbl-last-row tbl)
              for range-str = (format nil "~a~a:~a~a" from-letter tbl-fr to-letter tbl-lr)
              for i from 0
              do (when (> i 0) (format stream ", "))
                 (format stream "{\"range\": ~s, \"row_step\": ~a, \"col_step\": ~a}"
                         range-str cell-h cell-w))
        (format stream "],~%")
        ;; ── column_widths ──
        (format stream "        \"column_widths\": {")
        (let ((idx 1) (first t))
          (dolist (tbl tables)
            (dolist (name (col-names tbl))
              (unless first (format stream ", "))
              (setf first nil)
              (format stream "~a: ~a" idx
                      (case (if (symbolp name) name (intern (string-upcase name)))
                        ((programa-calc) 30) ((duracion) 8) ((tipo) 10)
                        ((hora-inicio) 10) ((hora-terminacion) 10) ((tipo-calc) 8)
                        (otherwise 10)))
              (incf idx))
            (unless first (format stream ", "))
            (format stream "~a: 0" idx)
            (incf idx))
          (loop for tbl in tables do
            (loop for (n . v) in (params tbl) do
                  (format stream ", ~a: 8" idx)
                  (incf idx))))
        (format stream "},~%")
        ;; ── border_color ──
        (format stream "        \"border_color\": \"4F81BD\",~%")
        (format stream "        \"border_style\": \"thick\"~%")
        (format stream "    }"))))))  ;; flet / let* / labels / let ((tables)) / defmethod

; =====================================================================
; GENERATE-CODE — HOJAS
; =====================================================================

(defmethod generate-code ((sh clase-xl-sheet) (lang xl-out) (stream t))
  (let ((*sheet-fixed-expressions* (or (fixed-expressions sh) nil)))
    (format stream "        {~%")
    (format stream "            \"title\": ~s,~%" (name sh))
    (format stream "            \"regions\": [~%")
    (let ((regs (regions sh)))
      (loop for i from 0 for region in regs do
           (when (> i 0) (format stream ",~%"))
           (generate-code region lang stream)))
    (format stream "            ]~%")
    (format stream "        }")))

; =====================================================================
; GENERATE-CODE — WORKBOOK
; =====================================================================

(defmethod generate-code ((wb clase-xl-workbook) (lang xl-out) (stream t))
  (format stream "#!/usr/bin/env python3~%")
  (format stream "from hoja_con_formulas import generar_excel_personalizado~2%")
  (format stream "config = {~%")
  (format stream "    \"sheets\": [~%")
  (when (sheets wb)
    (loop for i from 0 for sh in (sheets wb) do
      (generate-code sh lang stream)
      (when (< i (1- (length (sheets wb)))) (format stream ",~%"))))
  (format stream "    ]~%")
  (format stream "}~2%")
  (format stream "generar_excel_personalizado(config, ~s)~%" (name wb))
  (format stream "~%if __name__=='__main__':~%")
  (format stream "    print('OK: ~a')~%" (name wb)))

(provide "generate-code-direct")
