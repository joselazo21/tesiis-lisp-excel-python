;; dsl-lenguaje-visual.lisp
;; Lenguaje de alto nivel para describir estructuras de datos.
;; Macros: def-expresion, def-tabla, def-hoja, def-horario
;;
;; BACKEND-AGNOSTICO. Solo produce nodos AST (xl-*).
;; NO contiene: Excel-ismos, ni compilación de fórmulas.

(load "codigo-tesis.lisp")
(load "modelo-excel.lisp")

;; =====================================================================
;; UTILIDADES
;; =====================================================================

(defun parse-hoja-body (body)
  (let ((plist '())
        (table-calls '())
        (len (length body))
        (i 0))
    (loop while (< i len)
          do (let ((form (nth i body)))
               (cond
                 ((keywordp form)
                  (push (cons form (nth (1+ i) body)) plist)
                  (incf i 2))
                 ((and (listp form) (symbolp (first form))
                       (not (keywordp (first form))))
                  (push form table-calls)
                  (incf i 1))
                 (t
                  (incf i 1)))))
    (values (nreverse plist) (nreverse table-calls))))

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
;; DEF-TABLA
;; =====================================================================

(defmacro def-tabla (name params &body body)
  (let* ((columns (getf body :columns))
         (column-names (mapcar #'first columns))
         (column-display-names (mapcar #'second columns)))
    `(defun ,name (,@params &key (data nil))
       (let ((col-defs
               (loop for name in ',column-names
                     for display in ',column-display-names
                     collect (xl-col-def :name name :display-name display)))
             (table (xl-table :contenido (or data '())
                              :headers ',column-display-names)))
         (values table col-defs)))))

;; =====================================================================
;; DEF-HOJA
;; =====================================================================

(defmacro def-hoja (name params &body body)
  "Define un constructor de hoja.

   Sintaxis:
     (def-hoja nombre (param1 param2 ...)
       ;; REQUERIDO: un constructor de tabla
       (nombre-tabla arg1 arg2 ...)

       ;; Opcionales:
       ;; :data  — matriz de datos
       ;; :name  — string nombre de hoja (default: nombre del primer parámetro)
       )"
  (multiple-value-bind (plist table-calls)
      (parse-hoja-body body)
    (let* ((table-call (first table-calls))
           (table-name (first table-call))
           (table-args (rest table-call))
           (name-expr (or (cdr (assoc :name plist))
                          `(format nil "~a" ,(first params))))
           (data-expr (cdr (assoc :data plist))))
      `(defun ,name (,@params &key (data nil))
         (multiple-value-bind (table col-defs)
             (,table-name ,@table-args)
           (declare (ignore col-defs))
           (let ((sheet-data (or ,data-expr data (contenido table))))
             (xl-sheet
               :name ,name-expr
               :regions (list
                 (xl-region
                   :tables (list (xl-table :contenido sheet-data)))))))))))

;; =====================================================================
;; DEF-HORARIO: define el libro completo
;; =====================================================================
;;
;; Sintaxis:
;;   (def-horario nombre
;;     [:nombre "archivo.xlsx"]
;;     :hojas (lista-de-hojas ...))

(defmacro def-horario (name &body body)
  (let* ((hojas-form (getf body :hojas))
         (filename (or (getf body :nombre)
                        (concatenate 'string (string-downcase (symbol-name name))
                                     ".xlsx"))))
    `(let* ((sheets ,hojas-form)
            (wb (xl-workbook :name ,filename :sheets sheets)))
       (defparameter ,name wb)
       (format t "Libro ~a creado con ~a hojas~%" ',name (length sheets))
       wb)))

(provide "dsl-lenguaje-visual")
