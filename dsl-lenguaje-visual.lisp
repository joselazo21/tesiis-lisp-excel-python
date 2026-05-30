;; dsl-lenguaje-visual.lisp
;; Lenguaje de alto nivel para describir visualización de horarios.
;; Macros: def-tabla, def-hoja, def-horario
;;
;; Este archivo es BACKEND-AGNOSTICO. Solo produce nodos AST (xl-*).
;; NO contiene: column-letters, cell-refs, rangos Excel, códigos hex,
;;            ni ninguna expresión de Excel.
;; Todo eso se compila en los métodos generate-code de modelo-excel.lisp.
;;
;; El flujo:
;;   DSL macro → produce xl-* AST con datos simbólicos
;;   xl-generate → generate-code (en modelo-excel.lisp) → compila a Excel

(load "codigo-tesis.lisp")
(load "modelo-excel.lisp")  ;; ← carga compilador-formulas.lisp internamente

;; =====================================================================
;; UTILIDADES
;; =====================================================================

(defun collect-plist-key (plist key)
  "Colecciona todos los valores para KEY en PLIST (soporta claves duplicadas).
   PLIST es una lista plana (:key1 val1 :key2 val2 ...)"
  (loop for (k v) on plist by #'cddr
        when (eq k key)
        collect v))

(defun parse-cond-format-def (cfdef)
  "Parsea una definición de :cond-format.
   Retorna (values target-col row-var body apply-from).
   Soporta:
     (col :rule (var) _if-body)           — inline (old style)
     (col :rule nombre-expresion)         — expresión (new style)
   :apply-from es opcional, default 1."
  (let* ((target-col (first cfdef))
         (rest-spec (rest cfdef))
         (rule-pos (position :rule rest-spec))
         (rule-val (if rule-pos (nth (1+ rule-pos) rest-spec) nil))
         (apply-from-pos (position :apply-from rest-spec))
         (apply-from (if apply-from-pos
                         (nth (1+ apply-from-pos) rest-spec)
                         1)))
    (if (symbolp rule-val)
        ;; (col :rule nombre-expresion)
        (let ((expr (gethash rule-val *expresiones*)))
          (unless expr
            (error "Expresión ~a no encontrada (def-expresion?)" rule-val))
          (values target-col nil (body expr) apply-from))
        ;; (col :rule (var) _if-body)
        (let ((row-var (first rule-val))
              (body (find-if (lambda (x)
                                (and (listp x) (eq (first x) '_if)))
                              (cddr (member :rule rest-spec)))))
          (values target-col row-var body apply-from)))))

(defun parse-hoja-body (body)
  "Parsea el body de def-hoja.
   Retorna (values plist table-calls)
   plist: alist de (key . value) para pares keyword/valor
   table-calls: lista de llamadas a tablas"
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
;; DEF-EXPRESION: define una expresión reutilizable
;; =====================================================================
;;
;; Sintaxis:
;;   (def-expresion nombre ()
;;     cuerpo-de-la-expresion)
;;
;; El cuerpo usa símbolos como llamadas a función (0 args):
;;   (nombre-del-programa)  → columna del contexto
;;   (hora-inicio)          → otra columna
;;   (duracion expr)        → lookup en tabla de datos
;;   (previous-row expr)    → evalúa expr en fila anterior
;;
;; generate-code produce código específico del backend:
;;   Excel:  G4, E4, VLOOKUP(G4,...)
;;   TS:     row.programName, row.startTime, getDuration(...)

(defmacro def-expresion (name params &body body)
  "Define una expresión pura, sin contexto de tabla.
   name:   símbolo que nombra la expresión
   params: lista de parámetros formales (vacía si no tiene)
   body:   un solo S-expression (el cuerpo de la expresión)
   La expresión se registra en *expresiones* para referencia posterior."
  (declare (ignore params))
  (let ((expr-body (first body)))
    `(setf (gethash ',name *expresiones*)
           (xl-expresion :body ',expr-body))))

;; =====================================================================
;; DEF-TABLA: define una plantilla de tabla
;; =====================================================================
;;
;; Sintaxis:
;;   (def-tabla nombre (param1 param2 ...)
;;     :columns ((col-name "Display Name") ...)
;;     [:height n]
;;     :formula (col-target :compute (var) cuerpo)   ;; inline (old style)
;;     :formula (col-target :compute nombre-expresion) ;; por nombre (new style)
;;     :formula ...                                   ;; repetible
;;     :cond-format (col-target :rule (var) cuerpo [:apply-from n])
;;     :cond-format ...)                              ;; repetible
;;
;; Genera: función constructora que retorna
;;   (values xl-table col-defs formula-defs cond-format-defs)
;;   Todos son objetos AST agnósticos, sin Excel-ismos.

(defun parse-formula-def (fdef)
  "Parsea una definición de :formula.
   Retorna (values target-col row-var body).
   Si es estilo expresión: row-var es nil, body es el cuerpo de la expresión.
   Si es estilo inline: row-var es el símbolo de fila, body es el cuerpo."
  (let* ((target-col (first fdef))
         (compute-part (third fdef))
         (rest-part (cddr (member :compute fdef))))
    (if (symbolp compute-part)
        ;; (col :compute nombre-expresion) → lookup en *expresiones*
        (let ((expr (gethash compute-part *expresiones*)))
          (unless expr
            (error "Expresión ~a no encontrada (def-expresion?)" compute-part))
          (values target-col nil (body expr)))
        ;; (col :compute (var) cuerpo) → inline
        (let ((row-var (first compute-part))
              (body (first rest-part)))
          (values target-col row-var body)))))

(defmacro def-tabla (name params &body body)
  (let* ((columns (getf body :columns))
         (ncols (length columns))
         (column-names (mapcar #'first columns))
         (column-display-names (mapcar #'second columns))
         (height (or (getf body :height) 1))
         (formula-defs (collect-plist-key body :formula))
         (cond-format-defs (collect-plist-key body :cond-format)))
    `(defun ,name (,@params &key (data nil))
       "Constructor generado por def-tabla.
        Retorna (values xl-table col-defs formula-defs cond-format-defs)
        Todos son objetos AST agnósticos (sin Excel-ismos)."
       (let* ((ncols ,ncols)
              (column-names ',column-names)
              (column-display-names ',column-display-names)
              ;; Crear xl-col-def objects (agnósticos)
              (col-defs
                (loop for name in column-names
                      for display in column-display-names
                      collect (xl-col-def :name name :display-name display)))
              ;; Crear xl-formula-def objects (simbólicos, sin compilar)
              (formula-def-objs
                (loop for fdef in ',formula-defs
                      collect
                      (multiple-value-bind (target-col row-var body)
                          (parse-formula-def fdef)
                        (xl-formula-def
                          :target-col target-col
                          :row-var row-var
                          :body body))))
              ;; Crear xl-cond-format-def objects (simbólicos)
              (cond-format-objs
                (loop for cfdef in ',cond-format-defs
                      collect
                      (multiple-value-bind (target-col row-var body apply-from)
                          (parse-cond-format-def cfdef)
                        (xl-cond-format-def
                          :target-col target-col
                          :row-var row-var
                          :body body
                          :apply-from apply-from)))))
         (declare (ignore ncols column-names column-display-names))
         ;; Construir xl-table (solo datos)
         (let ((table (xl-table :contenido (or data '())
                                 :headers column-display-names)))
           (values table
                   col-defs
                   formula-def-objs
                   cond-format-objs))))))

;; =====================================================================
;; DEF-HOJA: define una plantilla de hoja
;; =====================================================================
;;
;; Sintaxis:
;;   (def-hoja nombre (param1 ...)
;;     :key value ...
;;     (nombre-tabla arg1 arg2 ... :key value ...))
;;
;; La macro genera un constructor que:
;;   1. Recibe datos abstractos
;;   2. Construye la matriz de datos (layout) sin Excel-ismos
;;   3. Almacena las definiciones DSL en slots dsl-* del xl-sheet
;;   4. generate-code (en modelo-excel.lisp) compila todo a Excel

(defmacro def-hoja (name params &body body)
  "Define un constructor de hoja genérico.

   Sintaxis:
     (def-hoja nombre (param1 param2 ...)
       ;; REQUERIDO: una llamada a un constructor de tabla
       (nombre-tabla arg1 arg2 ...)

       ;; Opcionales (keyword/value):
       ;; :data       — matriz de datos completa
       ;; :layout     — lista de elementos de layout abstracto (solo :border, :merge, :col-width)
       ;; :params     — alist (sym . (col . row))
       ;; :ref-table  — plist (:sc :sr :ec :er :cn)
       ;; :table-pos  — (start-col . start-row)
       ;; :name          — string nombre de hoja (default: (format nil \"~a\" primer-param))
       ;; :fernando-formulas — lista
       ;; :conditional-format-rules — lista
       )

   NOTA: La macro NO hardcodea ningún layout ni columna específica.
   Todo lo específico del dominio debe pasarse como keyword.
   El estilo es responsabilidad del backend — no se aceptan parámetros de estilo aquí."
  (multiple-value-bind (plist table-calls)
      (parse-hoja-body body)
    (let* ((table-call (first table-calls))
           (table-name (first table-call))
           (table-args (rest table-call))
           (name-expr (or (cdr (assoc :name plist))
                          `(format nil "~a" ,(first params))))
           (data-expr (cdr (assoc :data plist)))
           (layout-expr (cdr (assoc :layout plist)))
           (params-expr (cdr (assoc :params plist)))
           (ref-table-expr (cdr (assoc :ref-table plist)))
           (table-pos-expr (cdr (assoc :table-pos plist)))
           (fernando-expr (cdr (assoc :fernando-formulas plist)))
           (cfr-expr (cdr (assoc :conditional-format-rules plist))))
      `(defun ,name (,@params &key (data nil))
         "Constructor de hoja generado por def-hoja.
          Retorna un xl-sheet con una región DSL."
         (multiple-value-bind (table t-col-defs t-formula-defs t-cond-format-defs)
             (,table-name ,@table-args)
           (let ((sheet-data (or ,data-expr data (contenido table))))
             (xl-sheet
               :name ,name-expr
               :regions (list
                 (xl-region
                   :tables (list (xl-table :contenido sheet-data))
                   :dsl-layout ,layout-expr
                   :dsl-formula-defs t-formula-defs
                   :dsl-cond-format-defs t-cond-format-defs
                   :dsl-col-defs t-col-defs
                   :dsl-params ,params-expr
                   :dsl-ref-table ,ref-table-expr
                   :dsl-table-pos ,table-pos-expr
                   :fernando-formulas ,fernando-expr
                   :conditional-format-rules ,cfr-expr)))))))))

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
