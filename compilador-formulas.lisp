;; compilador-formulas.lisp
;; Compila expresiones de fórmulas simbólicas a cadenas de fórmula Excel.
;; El AST de la fórmula es una expresión simbólica que describe QUÉ calcular,
;; no CÓMO. El compilador traduce eso a sintaxis Excel.

(load "codigo-tesis.lisp")

;; =====================================================================
;; COLORES ABSTRACTOS
;; Los colores en el DSL se escriben como keywords (:rojo, :naranja)
;; y se mapean a códigos hex aquí. Solo este lugar sabe de códigos hex.
;; =====================================================================

(defparameter *colores*
  '((:rojo     . "FF0000")
    (:naranja  . "FFA500")
    (:verde    . "00FF00")
    (:azul     . "0000FF")
    (:amarillo . "FFFF00")
    (:gris     . "B7B7B7")
    (:blanco   . "FFFFFF")
    (:negro    . "000000")
    (:celeste  . "CFE2F3")
    (:verde-claro . "D9EAD3")
    (:naranja-claro . "FCE5CD")
    (:azul-oscuro  . "4A90E2")
    (:fondo-titulo . "FFF2CC")
    (:fondo-total  . "FFF2CC")
    (:fondo-header . "F4CCCC"))
  "Mapa de nombres de colores abstractos a códigos hex.")
  
(defun color->hex (color)
  "Convierte un color abstracto a hex.
   Acepta: :rojo → 'FF0000', \"FF0000\" → 'FF0000' (passthrough)"
  (typecase color
    (keyword (let ((hex (cdr (assoc color *colores*))))
               (or hex (string-upcase (symbol-name color)))))
    (string color)
    (t (format nil "~a" color))))

;; =====================================================================
;; UTILIDADES
;; =====================================================================

(defun column-letter (n)
  "Convierte número de columna 1-based a letra Excel (A, B, ..., Z, AA, ...)"
  (with-output-to-string (s)
    (loop while (> n 0)
          do (multiple-value-bind (q r) (floor (1- n) 26)
               (setf n q)
               (format s "~c" (code-char (+ 65 r)))))))

(defun cell-ref (col row &optional absolute)
  "Genera referencia de celda Excel: $LETRA$fila (absoluta) o LETRAfila (relativa)"
  (if absolute
      (format nil "$~a$~a" (column-letter col) row)
      (format nil "~a~a" (column-letter col) row)))

;; =====================================================================
;; CONTEXTO DE COMPILACIÓN
;; =====================================================================

(defstruct compile-context
  "Contexto para compilar expresiones de fórmula a Excel.
col-map:     alist (nombre-col . (num-col . letra-col))
start-row:   fila donde empiezan los datos (entero)
param-cells: alist (nombre-param . referencia-celda-Ej: (hora-inicio-param . \"$I$2\")
data-tables: alist (nombre-tabla . plist)
             (:start-col :start-row :end-row :column-names)
current-row: fila actual (para expresiones 0-arg, opcional)"
  col-map
  start-row
  param-cells
  data-tables
  current-row)

;; =====================================================================
;; COMPILACIÓN: expresión → string de fórmula Excel
;; =====================================================================

(defun compile-formula (expr context)
  "Compila expresión simbólica a string de fórmula Excel (sin el = inicial)."
  (cond
    ((null expr) "\"\"")
    ((eq expr 'show-nothing) "\"\"")
    ((symbolp expr) (compile-symbol expr context))
    ((atom expr) (format nil "~a" expr))
    (t (compile-list expr context))))

(defun compile-symbol (sym context)
  "Compila un símbolo: parámetro conocido → celda, otro → literal."
  (let ((param (assoc sym (compile-context-param-cells context))))
    (if param
        (cdr param)
        (symbol-name sym))))

(defun compile-list (expr context)
  "Compila una expresión en forma de lista (operación o llamada)."
  (let ((op (first expr))
        (args (rest expr)))
    (case op
      (_if           (compile-if-body args context))
      (non-empty     (compile-non-empty args context))
      (it-is-the-first-row (compile-first-row args context))
      (previous-row  (compile-previous-row args context))
      (show-nothing  "\"\"")
      (time-add      (compile-time-add args context))
      (is-not-defined (compile-is-not-defined args context))
      (is-the-same   (compile-is-the-same args context))
      (>  (compile-binary-infix ">" args context))
      (<  (compile-binary-infix "<" args context))
      (>= (compile-binary-infix ">=" args context))
      (<= (compile-binary-infix "<=" args context))
      (+  (compile-binary-infix "+" args context))
      (-  (compile-binary-infix "-" args context))
      (*  (compile-binary-infix "*" args context))
      (/  (compile-binary-infix "/" args context))
      (t
       (let ((col-info (assoc op (compile-context-col-map context))))
         (cond
           ;; (nombre-col) → cell at current-row (expression style, 0 args)
           ((and col-info (null args))
            (let ((cr (compile-context-current-row context)))
              (if cr
                  (cell-ref (cadr col-info) cr)
                  (error "compile-list: ~a necesita current-row en contexto" op))))
           ;; (nombre-col row) → cell at row (old style, 1 arg)
           ((and col-info (= (length args) 1))
            (compile-column-ref col-info (first args) context))
           ;; lookup en tabla de datos
           (t
            (let ((table-col (find-data-table-column op context)))
              (if table-col
                  (compile-data-table-lookup op args table-col context)
                  (format nil "~a(~{~a~^,~})" op
                          (mapcar (lambda (a) (compile-formula a context)) args)))))))))))

;; =====================================================================
;; COMPILADORES POR OPERACIÓN
;; =====================================================================

(defun compile-column-ref (col-info row-expr context)
  "Compila referencia a columna: (hora-inicio 4) → E4 (relativa)"
  (declare (ignore context))
  (let ((col-num (car (cdr col-info))))
    (cell-ref col-num row-expr)))

(defun compile-if-body (args context)
  "(_if cond then &optional else)"
  (let ((cond-str (compile-formula (first args) context))
        (then-str (compile-formula (second args) context))
        (else-str (if (third args)
                      (compile-formula (third args) context)
                      "\"\"")))
    (format nil "IF(~a,~a,~a)" cond-str then-str else-str)))

(defun compile-non-empty (args context)
  "(non-empty expr) → expr<>\"\""
  (format nil "~a<>\"\"" (compile-formula (first args) context)))

(defun compile-first-row (args context)
  "(it-is-the-first-row row) → ROW()=start-row"
  (declare (ignore args))
  (format nil "ROW()=~a" (compile-context-start-row context)))

(defun compile-previous-row (args context)
  "(previous-row x)
   Si x es número: devuelve x-1 (old style, post subst-row-var).
   Si x es expresión: la evalúa con current-row-1 (expression style)."
  (let ((inner (first args)))
    (if (numberp inner)
        (1- inner)
        (let ((current-row (compile-context-current-row context)))
          (unless current-row
            (error "previous-row necesita current-row en el contexto"))
          (let ((ctx (copy-compile-context context)))
            (setf (compile-context-current-row ctx) (1- current-row))
            (compile-formula inner ctx))))))

(defun compile-time-add (args context)
  "(time-add time-expr dur-expr) → TEXT(TIMEVALUE(t)+d/1440,\"hh:mm\")"
  (format nil "TEXT(TIMEVALUE(~a)+~a/1440,\"hh:mm\")"
          (compile-formula (first args) context)
          (compile-formula (second args) context)))

(defun compile-is-not-defined (args context)
  "(is-not-defined valor :in tabla) → ISNA(MATCH(valor, rango-col1, 0))
   Soporta: (is-not-defined valor tabla) o (is-not-defined valor :in tabla)"
  (let* ((value-expr (compile-formula (first args) context))
         (table-name (cond
                      ((eq (second args) :in) (third args))
                      (t (second args))))
         (table-info (cdr (assoc table-name
                                  (compile-context-data-tables context)))))
    (if (null table-info)
        (error "Tabla de datos ~a no encontrada en contexto" table-name)
        (let ((start-col (getf table-info :start-col))
              (start-row (getf table-info :start-row))
              (end-col (getf table-info :end-col))
              (end-row (getf table-info :end-row)))
          (format nil "ISNA(MATCH(~a,$~a$~a:$~a$~a,0))"
                  value-expr
                  (column-letter start-col) start-row
                  (column-letter end-col) end-row)))))

(defun compile-is-the-same (args context)
  "(is-the-same a b) → a=b"
  (format nil "~a=~a"
          (compile-formula (first args) context)
          (compile-formula (second args) context)))

(defun compile-binary-infix (op-str args context)
  "(> a b) → a>b. Múltiples args: a>b>c."
  (with-output-to-string (s)
    (loop for a in args
          for first = t then nil
          do (unless first (write-string op-str s))
             (write-string (compile-formula a context) s))))

;; =====================================================================
;; LOOKUP EN TABLAS DE DATOS
;; =====================================================================

(defun find-data-table-column (col-name context)
  "Busca col-name en las tablas de datos del contexto.
   Devuelve (table-name . col-index) o nil.
   col-index es 1-based (para VLOOKUP)."
  (dolist (table-entry (compile-context-data-tables context))
    (let* ((table-name (car table-entry))
           (table-info (cdr table-entry))
           (col-names (getf table-info :column-names))
           (pos (position col-name col-names)))
      (when pos
        (return-from find-data-table-column
          (cons table-name (1+ pos))))))
  nil)

(defun compile-data-table-lookup (op args table-col context)
  "(duracion expr) → IFERROR(VLOOKUP(valor, rango, col, FALSE),0)"
  (let* ((lookup-expr (compile-formula (first args) context))
         (table-name (car table-col))
         (return-col-idx (cdr table-col))
         (table-info (cdr (assoc table-name
                                  (compile-context-data-tables context))))
         (start-col (getf table-info :start-col))
         (start-row (getf table-info :start-row))
         (end-col (getf table-info :end-col))
         (end-row (getf table-info :end-row)))
    (format nil "IFERROR(VLOOKUP(~a,$~a$~a:$~a$~a,~a,FALSE),0)"
            lookup-expr
            (column-letter start-col) start-row
            (column-letter end-col) end-row
            return-col-idx)))

;; =====================================================================
;; SUSTITUCIÓN DE VARIABLE DE FILA
;; =====================================================================

(defun subst-row-var (expr row-var current-row)
  "Reemplaza row-var por current-row. (previous-row row-var) → current-row-1.
   Devuelve expresión con números concretos."
  (cond
    ((eq expr row-var) current-row)
    ((atom expr) expr)
    ((listp expr)
     (let ((op (first expr))
           (args (rest expr)))
       (cond
         ;; (previous-row row-var) → número
         ((and (eq op 'previous-row) (= (length args) 1)
               (eq (first args) row-var))
          (1- current-row))
         ;; (previous-row (previous-row row-var)) → número
         ((and (eq op 'previous-row) (= (length args) 1)
               (listp (first args)))
          (let ((inner (subst-row-var (first args) row-var current-row)))
            (if (numberp inner) (1- inner) `(previous-row ,inner))))
         ;; otro: recursión
         (t (cons op (mapcar (lambda (e)
                               (subst-row-var e row-var current-row))
                             args))))))))

;; =====================================================================
;; COMPILACIÓN DE COLUMNA COMPLETA
;; =====================================================================

(defun compile-column-formula-rows (target-col-name row-var body
                                      context num-rows
                                      &key (sheet-start-row 4))
  "Compila una fórmula para toda una columna.
   Genera lista de (fila . fórmula-string) para cada fila de datos.
   Si row-var es nil, compila en expression-style (current-row implícito)."
  (let ((col-info (cdr (assoc target-col-name
                                (compile-context-col-map context))))
        (results '()))
    (dotimes (i num-rows (nreverse results))
      (let* ((current-row (+ sheet-start-row i))
             (subst-expr (if row-var
                           (subst-row-var body row-var current-row)
                           body))
             (ctx (if row-var
                    context
                    (let ((new (copy-compile-context context)))
                      (setf (compile-context-current-row new) current-row)
                      new)))
             (formula (compile-formula subst-expr ctx)))
        (push (cons current-row (concatenate 'string "=" formula)) results)))))

;; =====================================================================
;; COMPILACIÓN DE FORMATO CONDICIONAL
;; =====================================================================

(defun parse-cond-format-body (_if-expr)
  "Parsea el body de un _if en contexto de cond-format.
   Extrae condición y acción (colores abstractos :rojo → hex).
   (_if cond (set-bg-color :rojo)) → (cond . \"FF0000\")"
  (let ((cond-expr (second _if-expr))
        (action (third _if-expr)))
    (if (and (listp action)
             (or (eq (first action) 'set-bg-color)
                 (eq (first action) 'set-background-color)))
        (cons cond-expr (color->hex (second action)))
        (cons cond-expr nil))))

(defun compile-cond-format-rule (target-col-name row-var
                                  _if-expr context num-rows
                                  &key (apply-from 1) (sheet-start-row 4))
  "Compila una regla de formato condicional.
   _if-expr es la expresión _if completa.
   Devuelve lista de xl-conditional-rule."
  (let* ((col-info (cdr (assoc target-col-name
                                (compile-context-col-map context))))
         (col-letter (cdr col-info))
         (parsed (parse-cond-format-body _if-expr))
         (cond-expr (car parsed))
         (color (cdr parsed))
         (results '()))
    (unless color
      (return-from compile-cond-format-rule results))
    (dotimes (i num-rows (nreverse results))
      (when (>= (1+ i) apply-from)
        (let* ((current-row (+ sheet-start-row i))
               (subst-expr (if row-var
                             (subst-row-var cond-expr row-var current-row)
                             cond-expr))
               (ctx (if row-var
                      context
                      (let ((new (copy-compile-context context)))
                        (setf (compile-context-current-row new) current-row)
                        new)))
               (cond-str (compile-formula subst-expr ctx))
               (cell-range (format nil "~a~a" col-letter current-row)))
          (push (make-instance 'clase-xl-conditional-rule
                  :tipo "rango"
                  :rango cell-range
                  :formula cond-str
                  :color color)
                results))))))

(provide "compilador-formulas")
