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
;; TABLE: expande directamente a xl-table
;; =====================================================================
;;
;; (tabla (programa "Programa") (duracion "Duración") (tipo "Tipo")
;;   :data ((prog1 30 tipo1) (prog2 45 tipo2)))

(defmacro tabla (columns &body body)
  (let* ((col-names (mapcar #'first columns))
         (display-names (mapcar #'second columns))
         (data (getf body :data))
         (computed (getf body :computed))
         (params (getf body :params)))
    `(xl-table :contenido (or ,data '()) :headers ',display-names
               :col-names ',col-names
               :computed (list ,@(loop for (col expr) in computed
                                       collect `(cons ',col ,expr)))
               :params (list ,@(loop for (name val) in params
                                     collect `(cons ',name ,val))))))

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
