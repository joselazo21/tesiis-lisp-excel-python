;; ast-def.lisp — Definiciones de todas las clases del AST
;; Backend-agnóstico. Sin generate-code, sin comportamiento.
(load "codigo-tesis.lisp")

(defclass* xl-col-def () (name display-name))

(defclass* xl-expr-if () (test then else))
(defclass* xl-expr-non-empty () (expr))
(defclass* xl-expr-first-row () ())
(defclass* xl-expr-column-ref () (name context))
(defclass* xl-expr-param-ref () (name))
(defclass* xl-expr-previous-row () (expr))
(defclass* xl-expr-next-row () (expr))
(defclass* xl-expr-time-add () (a b))
(defclass* xl-expr-show-nothing () ())
(defclass* xl-expr-lookup () (value-field key-expr))

(defclass* xl-expr-equals () (a b))
(defclass* xl-expr-different () (a b))
(defclass* xl-expr-and () (a b))
(defclass* xl-expr-or () (a b))

(defclass* xl-expr-add () (a b))
(defclass* xl-expr-subtract () (a b))
(defclass* xl-expr-multiply () (a b))
(defclass* xl-expr-divide () (a b))

(defclass* xl-range () (from-col to-col))

(defclass* xl-expr-countif () (count-range criteria))
(defclass* xl-expr-counta () (count-range))
(defclass* xl-expr-sum () (count-range))

(defclass* xl-expr-cross-sheet-ref () (sheet cell-template))

(defclass* xl-expr-concat () (a b))
(defclass* xl-expr-string () (value))

(defclass* xl-cell-ref () (row col sheet))

(defclass* xl-fixed-formula () (cell-ref expr))

(defclass* xl-style-rule () (rule-condition target-columns))

(defclass* xl-table () (id cols rows contenido headers computed col-names params style-rules fixed-formulas first-row cell-height cell-width paired-columns))

(defclass* xl-region () (tables))

(defclass* xl-sheet () (name regions))

(defclass* xl-workbook () (name sheets))

(defclass xl-out () ())
(defparameter xl-py (make-instance 'xl-out))

(provide "ast-def")
