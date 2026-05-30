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
(defclass* xl-expr-time-add () (a b))
(defclass* xl-expr-show-nothing () ())
(defclass* xl-expr-lookup () (value-field key-expr))

(defclass* xl-table () (id cols rows contenido headers computed col-names params))

(defclass* xl-region () (tables))

(defclass* xl-sheet () (name regions))

(defclass* xl-workbook () (name sheets))

(defclass xl-out () ())
(defparameter xl-py (make-instance 'xl-out))

(provide "ast-def")
