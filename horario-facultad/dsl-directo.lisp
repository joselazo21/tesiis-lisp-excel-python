;; dsl-directo.lisp
;; Macros que instancian nodos AST directamente (sin generar funciones).
;; Backend-agnóstico. Solo produce xl-* objects.

(load "codigo-tesis.lisp")
(load "generate-code-direct.lisp")

;; =====================================================================
;; EXPRESIONES: un macro por cada tipo de expresión
;; =====================================================================

(defmacro _if (test then else)
  `(xl-expr-if :test ,test :then ,then :else ,else))

(defmacro non-empty (expr)
  `(xl-expr-non-empty :expr ,expr))

(defmacro it-is-the-first-row ()
  `(xl-expr-first-row))

(defmacro previous-row (expr)
  `(xl-expr-previous-row :expr ,expr))

(defmacro time-add (a b)
  `(xl-expr-time-add :a ,a :b ,b))

(defmacro show-nothing ()
  `(xl-expr-show-nothing))

(defmacro lookup (key-expr field)
  `(xl-expr-lookup :value-field ',field :key-expr ,key-expr))

(defmacro col (name &optional context)
  (if context
      `(xl-expr-column-ref :name ',name :context ,context)
      `(xl-expr-column-ref :name ',name :context nil)))

(defmacro param (name)
  `(xl-expr-param-ref :name ',name))

;; =====================================================================
;; ARITMÉTICA
;; =====================================================================

(defmacro add (a b)
  `(xl-expr-add :a ,a :b ,b))

(defmacro subtract (a b)
  `(xl-expr-subtract :a ,a :b ,b))

(defmacro multiply (a b)
  `(xl-expr-multiply :a ,a :b ,b))

(defmacro divide (a b)
  `(xl-expr-divide :a ,a :b ,b))

;; =====================================================================
;; RANGOS — expresan un bloque de columnas sin letras Excel
;; =====================================================================
;;
;; (range lun)         → columna única (todas las filas)
;; (range lun vie)     → desde lun hasta vie (todas las filas)
;;
;; El rango siempre abarca de first-row a last-row (se resuelve
;; en compile-excel-formula).

(defmacro range (from &optional to)
  `(xl-range :from-col (xl-expr-column-ref :name ',from :context nil)
             :to-col ,(if to
                          `(xl-expr-column-ref :name ',to :context nil)
                          nil)))

;; =====================================================================
;; AGREGADOS
;; =====================================================================
;;
;; (countif (range lun vie) (col abrev))
;;   → COUNTIF($C$4:$L$9, N4)

(defmacro countif (range-expr criteria)
  `(xl-expr-countif :count-range ,range-expr :criteria ,criteria))

(defmacro counta (range-expr)
  `(xl-expr-counta :count-range ,range-expr))

(defmacro sum-range (range-expr)
  `(xl-expr-sum :count-range ,range-expr))

;; =====================================================================
;; STRING
;; =====================================================================

(defmacro str (value)
  `(xl-expr-string :value ,value))

(defmacro concat (a b)
  `(xl-expr-concat :a ,a :b ,b))

;; =====================================================================
;; CROSS-SHEET REF
;; =====================================================================
;;
;; (sheet-ref "C111" "$C$5") → C111!$C$5
;; (sheet-ref "C111" "$C${row-num}") → C111!$C5

(defmacro sheet-ref (sheet cell-template)
  `(xl-expr-cross-sheet-ref :sheet ,sheet :cell-template ,cell-template))

;; =====================================================================
;; CELL-REF — referencia absoluta a celda (fila, columna, hoja opcional)
;; =====================================================================
;;
;; (cell 10 14)            → fila 10, columna 14 (hoja actual)
;; (cell 10 14 "D111")     → fila 10, columna 14 en hoja D111 (futuro)

(defmacro cell (row column &optional sheet)
  `(xl-cell-ref :row ,row :col ,column :sheet ,sheet))

;; =====================================================================
;; FIXED-FORMULA — fórmula en celda fija
;; =====================================================================
;;
;; (fixed-formula (cell 10 14) (counta (range abrev)))
;;   → {"row": 10, "col": 14, "value": "=COUNTA(N4:N9)"}

(defmacro fixed-formula (cell-ref expr)
  `(xl-fixed-formula :cell-ref ,cell-ref :expr ,expr))

;; =====================================================================
;; REFERENCIAS RELATIVAS
;; =====================================================================

(defmacro previous-of (base)
  `(xl-expr-previous-row :expr ,base))

(defmacro next-of (base)
  `(xl-expr-next-row :expr ,base))

;; =====================================================================
;; CONDICIONES
;; =====================================================================

(defmacro equals (a b)
  `(xl-expr-equals :a ,a :b ,b))

(defmacro different (a b)
  `(xl-expr-different :a ,a :b ,b))

(defmacro _and (a b)
  `(xl-expr-and :a ,a :b ,b))

(defmacro _or (a b)
  `(xl-expr-or :a ,a :b ,b))

;; =====================================================================
;; RENDERIZADO CONDICIONAL
;; =====================================================================

(defmacro conditional-rendering (&key condition target-columns)
  `(xl-style-rule :rule-condition ,condition
                  :target-columns ',target-columns))

;; =====================================================================
;; DEF-TABLE: macro que DEFINE una función-clase de tabla
;; =====================================================================
;;
;; (def-table program-table
;;   ((programa "Programa") (duracion "Duración") (tipo "Tipo"))
;;   :computed ((hora-inicio ...) ...))
;;
;; (def-table program-table
;;   ...columnas...
;;   :render (conditional-rendering
;;             :condition (_or (equals (previous-of (col tipo))
;;                                    (col tipo))
;;                            (equals (col tipo)
;;                                    (next-of (col tipo))))
;;             :target-columns (hora-inicio hora-terminacion tipo-calc)))
;;
;; Define una función que acepta :data y :params y devuelve un xl-table.

(defmacro def-table (name columns &body body)
  (let ((col-names (mapcar #'first columns))
        (headers (mapcar #'second columns))
        (computed (getf body :computed))
        (render (getf body :render))
        (fixed (getf body :fixed-formulas))
        (first-row (getf body :first-row))
        (cell-height (or (getf body :cell-height) 1))
        (cell-width (or (getf body :cell-width) 1))
        (paired-columns (getf body :paired-columns)))
    `(defun ,name (&key data params)
       (xl-table :contenido (or data '())
                 :headers ',headers
                 :col-names ',col-names
                 :computed (list ,@(loop for (col expr) in computed
                                         collect `(cons ',col ,expr)))
                  :fixed-formulas (list ,@(loop for (cell-form expr-form) in fixed
                                                 collect `(xl-fixed-formula :cell-ref ,cell-form :expr ,expr-form)))
                  :style-rules (list ,@(when render (list render)))
                  :first-row ,(or first-row nil)
                  :cell-height ,cell-height
                  :cell-width ,cell-width
                  :paired-columns ',paired-columns
                  :params params))))

;; =====================================================================
;; TABLA: macro que INSTANCIA una tabla (llama a la función definida)
;; =====================================================================
;;
;; (tabla program-table
;;   :data *lunes-progs*
;;   :params ((hora-inicio-param *lunes-hour*)))

(defmacro tabla (first &rest args)
  (if (listp first)
      ;; forma inline: (tabla ((col "Header") ...) :data ... :computed ... :params ...)
      (let* ((col-names (mapcar #'first first))
             (headers (mapcar #'second first))
             (data (getf args :data))
             (computed (getf args :computed))
             (params (getf args :params))
             (fixed (getf args :fixed-formulas))
             (first-row (getf args :first-row))
             (cell-height (or (getf args :cell-height) 1))
             (cell-width (or (getf args :cell-width) 1))
             (paired-columns (getf args :paired-columns)))
    `(xl-table :contenido (or ,data '()) :headers ',headers
               :col-names ',col-names
               :computed (list ,@(loop for (col expr) in computed
                                       collect `(cons ',col ,expr)))
               :fixed-formulas (list ,@(loop for (cell-form expr-form) in fixed
                                              collect `(xl-fixed-formula :cell-ref ,cell-form :expr ,expr-form)))
               :style-rules (list ,@(when (getf args :render) (list (getf args :render))))
               :first-row ,(or first-row nil)
               :cell-height ,cell-height
               :cell-width ,cell-width
               :paired-columns ',paired-columns
               :params (list ,@(loop for (name val) in params
                                     collect `(cons ',name ,val)))))
      ;; forma instancia: (tabla program-table :data ... :params ...)
      (let ((data-arg (getf args :data))
            (params-arg (getf args :params)))
        `(,first :data ,data-arg
                 :params (list ,@(loop for (name val) in params-arg
                                       collect `(cons ',name ,val)))))))

;; =====================================================================
;; HOJA: expande directamente a xl-sheet con region + tablas
;; =====================================================================
;;
;; (hoja "Lunes"
;;   (tabla ...)
;;   (tabla ...))

(defmacro hoja (name &body tables)
  `(xl-sheet
     :name ,name
     :regions (list
       (xl-region
         :tables (list ,@tables)))))

;; =====================================================================
;; LIBRO: workbook + defparameter + generar
;; =====================================================================
;;
;; (libro horario-tv
;;   :filename "Horario.xlsx"
;;   :hojas (list (hoja ...) (hoja ...)))

(defmacro libro (name &body body)
  (let* ((filename (or (getf body :filename)
                       (concatenate 'string (string-downcase (symbol-name name)) ".xlsx")))
         (hojas-form (getf body :hojas)))
    `(let ((wb (xl-workbook :name ,filename :sheets ,hojas-form)))
       (defparameter ,name wb)
       (format t "Libro ~a creado con ~a hojas~%" ',name (length ,hojas-form))
       wb)))

(provide "dsl-directo")
