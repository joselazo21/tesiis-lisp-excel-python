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
(load "code-generation-utils.lisp")

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
         (explicit-first (table-first-row tbl num-params)))
    (flet ((emit-sep ()
              (when (> key-count 0) (format stream ",~%"))))
      ;; ── data ──
      (when con
        (emit-sep)
        (format stream "        \"data\": ")
        (if prms
            (let ((padded
                   (append
                     (list (append (make-list num-cols :initial-element "")
                                   (loop for (n . v) in prms collect v)))
                     (loop for row in con
                           collect (append row (make-list num-params :initial-element ""))))))
              (xl-write padded stream))
            (xl-write con stream))
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
                        collect (cons n (format nil "$~a~a" (col->letter col-num) 2))))))
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
        (let ((rules (style-rules tbl)))
          (when rules
            (setf items (nconc items (collect-range-styles-from-rules tbl rules first-row last-row)))))
        (when items
          (emit-sep)
          (format stream "        \"range_styles\": [")
          (loop for (range . style) in items for i from 0
                do (when (> i 0) (format stream ", "))
                   (format stream "{\"range\": ~s, \"style\": {\"bg_color\": ~s}}"
                           range style))
          (format stream "]")
          (incf key-count))))))

; =====================================================================
; GENERATE-CODE — REGIONES
; =====================================================================

(defmethod generate-code ((region clase-xl-region) (lang xl-out) (stream t))
  (let ((tables (tables region)))
    (unless tables (return-from generate-code (format stream "{}")))
    (labels ((col-count (tbl) (length (col-names tbl)))
             (effective-first-row (tbl)
               (or (first-row tbl) (+ 2 (if (params tbl) 1 0))))
             (logical-rows (tbl)
               (let* ((con (contenido tbl))
                      (ef (first-row tbl))
                      (pr (if ef (max 0 (1- ef)) 0)))
                 (max 0 (- (length con) pr))))
             (physical-rows (tbl)
               (* (logical-rows tbl) (max 1 (or (cell-height tbl) 1))))
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
                            do (incf cur (1+ (col-count tbl)))
                            collect start))
             (main-tbl (first tables))
             (ref-first-row (effective-first-row main-tbl))
             (ref-physical-rows (physical-rows main-tbl))
             (last-row (if (> ref-physical-rows 0)
                           (1- (+ ref-first-row ref-physical-rows))
                           (1- ref-first-row)))
             (total-cols (loop for tbl in tables sum (1+ (col-count tbl))))
             (last-col-letter (col->letter total-cols))
             (global-col-map (build-global-col-map tables))
             (global-dnames (loop for tbl in tables append (data-col-names tbl)))
             (all-expanded (loop for tbl in tables collect (expand-table-content tbl)))
             (max-rows (loop for exp in all-expanded maximize (length exp))))
        (format stream "{~%")
        ;; ── data ──
        ;; For tables with params: prepend a param row so $G2 (param cell) has
        ;; the initial value. The separator column doubles as the param column,
        ;; so r=0 writes the param value there instead of the usual "".
        (let* ((any-params (some (lambda (tbl) (params tbl)) tables))
               (padded-max (if any-params (1+ max-rows) max-rows)))
          (format stream "        \"data\": ")
          (xl-write
            (loop for r from 0 below padded-max
                  collect (loop for tbl in tables for exp in all-expanded
                                for prms = (params tbl)
                                for nc = (col-count tbl)
                                for data-r = (if prms (1- r) r)
                                for row = (if (and prms (zerop r))
                                              (make-list nc :initial-element "")
                                              (if (< data-r (length exp))
                                                  (nth data-r exp)
                                                  (make-list nc :initial-element "")))
                                append (subseq row 0 (min (length row) nc))
                                collect (if (and prms (zerop r))
                                            (cdar prms)
                                            "")))
            stream))
        (format stream ",~%")
        ;; ── headers ──
        (format stream "        \"headers\": ")
        (xl-write (loop for tbl in tables append (append (headers tbl) (list ""))) stream)
        (format stream ",~%")
        ;; ── formulas ──
        (let ((formulas nil))
          (loop for tbl in tables
                for offset in offsets
                for cell-h = (max 1 (or (cell-height tbl) 1))
                for first-row = (effective-first-row tbl)
                for ldr = (logical-rows tbl)
                for prms = (params tbl)
                for param-cells = (when prms
                                    (loop for (n . v) in prms for idx from 1
                                          for col-num = (+ (col-count tbl) idx)
                                          collect (cons n (format nil "$~a~a" (col->letter (+ offset col-num)) 2))))
                do
                ;; computed formulas
                (loop for (col . expr) in (computed tbl)
                      for tbl-col-idx = (position col (col-names tbl) :test #'string-equal)
                      for abs-col = (+ offset (or tbl-col-idx -1) 1)
                      do
                      (loop for i from 0 below ldr
                            for row = (+ first-row (* i cell-h))
                            for formula = (let ((*param-cells* param-cells))
                                            (compile-excel-formula expr global-col-map global-dnames row first-row last-row))
                            do (push (list row abs-col formula) formulas)))
                ;; fixed formulas
                (loop for ff in (fixed-formulas tbl)
                      for cr = (cell-ref ff)
                      for abs-col = (+ offset (slot-value cr 'col))
                      for formula = (let ((*param-cells* param-cells))
                                      (compile-excel-formula (expr ff) global-col-map global-dnames (row cr) first-row last-row))
                      do (push (list (row cr) abs-col formula) formulas)))
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
        ;; ── table_ranges ──
        (format stream "        \"table_ranges\": [~s],~%"
                (format nil "A~a:~a~a" ref-first-row last-col-letter last-row))
        ;; ── table_block_sizes ──
        (format stream "        \"table_block_sizes\": [")
        (loop for tbl in tables
              for offset in offsets
              for cell-h = (max 1 (or (cell-height tbl) 1))
              for cell-w = (max 1 (or (cell-width tbl) 1))
              for from-letter = (col->letter (1+ offset))
              for to-letter = (col->letter (+ offset (col-count tbl)))
              for range-str = (format nil "~a~a:~a~a" from-letter ref-first-row to-letter last-row)
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
        (format stream "    }")))))

; =====================================================================
; GENERATE-CODE — HOJAS
; =====================================================================

(defmethod generate-code ((sh clase-xl-sheet) (lang xl-out) (stream t))
  (format stream "        {~%")
  (format stream "            \"title\": ~s,~%" (name sh))
  (format stream "            \"regions\": [~%")
  (let ((regs (regions sh)))
    (loop for i from 0 for region in regs do
         (when (> i 0) (format stream ",~%"))
         (generate-code region lang stream)))
  (format stream "            ]~%")
  (format stream "        }"))

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
