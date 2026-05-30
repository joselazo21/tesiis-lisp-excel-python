; modelo-excel.lisp - Core AST model and code generation
; Backend-agnostic AST nodes + Excel-specific generate-code methods
(load "codigo-tesis.lisp")

; =====================================================================
; BACKEND-AGNOSTIC DSL AST NODES
; These classes store symbolic/abstract definitions from the DSL.
; They contain NO Excel specifics - no column letters, no cell refs.
; The generate-code methods compile them to the target backend.
; =====================================================================

; Column definition in a DSL table (e.g., hora-inicio, "Hora Inicio")
(defclass* xl-col-def () (name display-name))

; Pure expression — backend-agnostic, no binding context.
; Used inline in formulas or stored and referenced by name.
(defclass* xl-expresion () (body))

; Registry of named expressions (defined with def-expresion)
(defparameter *expresiones* (make-hash-table)
  "Registro de expresiones DSL, accesibles por nombre.")

; Symbolic formula definition for a computed column
; target-col: column to apply formula to (e.g., 'hora-final)
; row-var: the row variable symbol (e.g., 'row), nil if expression-style
; body: symbolic expression like (_if (non-empty ...) ...)
(defclass* xl-formula-def () (target-col row-var body))

; Symbolic conditional format definition
; apply-from: first row to apply (1-based within data, default 1)
(defclass* xl-cond-format-def () (target-col row-var body apply-from))

; =====================================================================
; BACKEND-AWARE CLASSES (used by generate-code methods)
; =====================================================================

(defclass* xl-table () (id cols rows contenido headers column-widths formulas))

; A region groups tables that should be drawn together, with shared
; layout, formulas, and reference data.
(defclass* xl-region ()
  (tables formulas fernando-formulas
   table-ranges range-styles merge-ranges
   conditional-format-rules column-widths
   table-block-sizes
   ;; DSL abstract definitions (set by def-hoja, consumed by generate-code)
   dsl-formula-defs dsl-cond-format-defs dsl-col-defs
   dsl-params dsl-ref-table dsl-table-pos dsl-layout
   ;; Compiled formula results (set during generate-code)
   compiled-formulas compiled-cond-rules))

; A sheet contains one or more regions.
; No style information — that's the backend's responsibility.
(defclass* xl-sheet () (name regions))

(defclass* xl-workbook () (name sheets))

; Style classes — kept for backend use; DSL never references them.
(defclass* xl-style () (bold color bg-color align))
(defclass* xl-range-style () (range style))
(defclass* xl-header-style () (bold color bg-color align))
(defclass* xl-formula () (row col value))
(defclass* xl-formula-list () (items))
(defclass* xl-fernando-formula () (cell formula))
(defclass* xl-conditional-rule () (tipo rango formula color))
(defclass* xl-merge-range () (range))
(defclass* xl-column-widths () (pairs))

; Output class
(defclass xl-out () ())
(defparameter xl-py (make-instance 'xl-out))

; =====================================================================
; UTILITIES
; =====================================================================

(defun xl-write (val s)
  (cond ((null val) (format s "None"))
        ((stringp val) (format s "~s" val))
        ((numberp val) (format s "~a" val))
        ((listp val) (progn (format s "[") (loop for i from 0 for x in val do (when (> i 0) (format s ", ")) (xl-write x s)) (format s "]")))
        (t (format s "~s" val))))

(defun xl-write-dict (plist s)
  (format s "{")
  (loop for i from 0 for (k v) on plist by #'cddr when v do (when (> i 0) (format s ", ")) (format s "~s: " k) (xl-write v s))
  (format s "}"))

; =====================================================================
; Cargar compilador de fórmulas (necesita las clases definidas arriba,
; antes de los métodos generate-code que lo usan)
; =====================================================================
(load "compilador-formulas.lisp")

; =====================================================================
; GENERATE-CODE METHODS (Excel backend)
; Todos los colores abstractos se convierten a hex aquí.
; =====================================================================

(defun resolve-color (color)
  "Convierte color (keyword o string) a hex. Passthrough para strings."
  (if (keywordp color)
      (color->hex color)
      color))

(defmethod generate-code ((st clase-xl-style) (lang xl-out) (stream t))
  (format stream "{")
  (let ((first t))
    (when (bold st) (format stream "~a\"bold\": True" (if first "" ", ")) (setf first nil))
    (when (color st) (format stream "~a\"color\": ~s" (if first "" ", ") (resolve-color (color st))) (setf first nil))
    (when (bg-color st) (format stream "~a\"bg_color\": ~s" (if first "" ", ") (resolve-color (bg-color st))) (setf first nil))
    (when (align st) (format stream "~a\"align\": ~s" (if first "" ", ") (align st)) (setf first nil))
    (format stream "}")))

(defmethod generate-code ((st clase-xl-header-style) (lang xl-out) (stream t))
  (format stream "{")
  (let ((first t))
    (when (bold st) (format stream "~a\"bold\": True" (if first "" ", ")) (setf first nil))
    (when (color st) (format stream "~a\"color\": ~s" (if first "" ", ") (resolve-color (color st))) (setf first nil))
    (when (bg-color st) (format stream "~a\"bg_color\": ~s" (if first "" ", ") (resolve-color (bg-color st))) (setf first nil))
    (when (align st) (format stream "~a\"align\": ~s" (if first "" ", ") (align st)) (setf first nil))
    (format stream "}")))

(defmethod generate-code ((rs clase-xl-range-style) (lang xl-out) (stream t))
  (format stream "{\"range\": ~s, \"style\": " (range rs))
  (generate-code (style rs) lang stream)
  (format stream "}"))

(defmethod generate-code ((rs list) (lang xl-out) (stream t))
  (let ((rng (getf rs :range)) (st (getf rs :style)))
    (when (and rng st)
      (format stream "{\"range\": ~s, \"style\": " rng)
      (let ((bg (getf st :bg-color))
            (fg (getf st :color)))
        (cond (bg (format stream "{\"bg_color\": ~s}" (resolve-color bg)))
              (fg (format stream "{\"color\": ~s}" (resolve-color fg)))
              ((getf st :bold) (format stream "{\"bold\": True}"))
              (t (format stream "{}"))))
      (format stream "}"))))

(defmethod generate-code ((f clase-xl-formula) (lang xl-out) (stream t))
  (format stream "{\"row\": ~a, \"col\": ~a, \"value\": ~s}" (row f) (col f) (value f)))

(defmethod generate-code ((f clase-xl-fernando-formula) (lang xl-out) (stream t))
  (format stream "{\"cell\": ~s, \"formula\": ~s}" (cell f) (formula f)))

(defmethod generate-code ((r clase-xl-conditional-rule) (lang xl-out) (stream t))
  (let ((color-val (color r)))
    (format stream "{\"tipo\": ~s, \"rango\": ~s, \"formula\": ~s, \"color\": ~s}"
            (tipo r) (rango r) (formula r) (resolve-color color-val))))

(defmethod generate-code ((m clase-xl-merge-range) (lang xl-out) (stream t))
  (format stream "~s" (range m)))

(defmethod generate-code ((cw clase-xl-column-widths) (lang xl-out) (stream t))
  (format stream "            \"column_widths\": ")
  (xl-write-dict (pairs cw) stream)
  (format stream ",~%"))

(defmethod generate-code ((lst clase-xl-formula-list) (lang xl-out) (stream t))
  (format stream "            \"formulas\": [")
  (loop for i from 0 for f in (items lst) do
       (when (> i 0) (format stream ", "))
       (generate-code f lang stream))
  (format stream "],~%"))

(defmethod generate-code ((tbl clase-xl-table) (lang xl-out) (stream t))
  (let ((con (contenido tbl)) (hdrs (headers tbl)) (cwidths (column-widths tbl)) (frms (formulas tbl)))
    (when con (format stream "        \"data\": ") (xl-write con stream) (format stream ",~%"))
    (when hdrs (format stream "        \"headers\": ") (xl-write hdrs stream) (format stream ",~%"))
    (when cwidths
      (if (typep cwidths 'clase-xl-column-widths)
          (generate-code cwidths lang stream)
          (progn
            (format stream "        \"column_widths\": ")
            (xl-write-dict cwidths stream)
            (format stream ",~%"))))
    (when frms
      (if (typep frms 'clase-xl-formula-list)
          (generate-code frms lang stream)
          (progn
            (format stream "        \"formulas\": [")
            (loop for i from 0 for f in frms do (when (> i 0) (format stream ", ")) (generate-code f lang stream))
            (format stream "],~%"))))))

(defun compile-dsl-formulas-for-region (region)
  "Compila las definiciones DSL simbólicas de una región en objetos Excel concretos.
   Retorna (values formulas cond-rules) donde cada uno es lista de
   objetos xl-formula / xl-conditional-rule listos para generar código."
  (let* ((col-defs (dsl-col-defs region))
         (table-pos (dsl-table-pos region))
         (start-col (or (car table-pos) 1))
         (start-row (or (cdr table-pos) 1))
         (params (dsl-params region))
         (ref-table (dsl-ref-table region))
         (tables-list (tables region))
         (data (when tables-list (contenido (first tables-list))))
         (nrows (length data))
         (col-map
           (loop for col-def in col-defs
                 for i from 0
                 for col-num = (+ start-col i)
                 collect (cons (name col-def)
                               (cons col-num (column-letter col-num)))))
         (data-tables
            (when ref-table
              (destructuring-bind (&key tn sc sr ec er cn) ref-table
                (let ((table-name (or tn 'programas)))
                  (list (cons table-name
                              (list :start-col sc :start-row sr
                                    :end-col ec :end-row er
                                    :column-names cn)))))))
         (param-cells
           (loop for (param-sym . cell-ref) in params
                 collect (cons param-sym
                               (let ((col (car cell-ref))
                                     (row (cdr cell-ref)))
                                 (cell-ref col row t)))))
         (ctx (make-compile-context
               :col-map col-map
               :start-row start-row
               :param-cells param-cells
               :data-tables data-tables))
         (all-formulas '())
         (all-cond-rules '()))
    (dolist (fdef (dsl-formula-defs region))
      (let* ((target-col (target-col fdef))
             (row-var (row-var fdef))
             (body (body fdef))
             (formulas (compile-column-formula-rows
                         target-col row-var body ctx nrows
                         :sheet-start-row start-row)))
        (dolist (f formulas)
          (let* ((col-info (cdr (assoc target-col col-map)))
                 (col-num (car col-info))
                 (row-num (car f))
                 (formula-str (cdr f)))
            (push (make-instance 'clase-xl-formula
                    :row row-num :col col-num :value formula-str)
                  all-formulas)))))
    (dolist (cfdef (dsl-cond-format-defs region))
      (let* ((target-col (target-col cfdef))
             (row-var (row-var cfdef))
             (body (body cfdef))
             (apply-from (apply-from cfdef))
             (rules (compile-cond-format-rule
                      target-col row-var body ctx nrows
                      :apply-from (or apply-from 1)
                      :sheet-start-row start-row)))
        (setf all-cond-rules (append rules all-cond-rules))))
    (values (nreverse all-formulas) (nreverse all-cond-rules))))

(defun compile-dsl-layout (layout-elts)
  "Compila elementos de layout abstracto a rangos Excel concretos.
   Entrada: lista de plists con :type.
   Retorna (values table-ranges merge-ranges col-width-plist table-block-sizes)
   Ninguna posición Excel aparece en la entrada, solo en la salida.
   NOTE: :type :style ha sido eliminado — el estilo es responsabilidad del backend."
  (let ((table-ranges '())
        (merge-ranges '())
        (col-width-pairs '())
        (table-block-sizes '()))
    (dolist (elt layout-elts)
      (let ((type (getf elt :type)))
        (case type
          (:border
           (let ((range (format nil "~a~a:~a~a"
                                (column-letter (getf elt :col-start))
                                (getf elt :row-start)
                                (column-letter (getf elt :col-end))
                                (getf elt :row-end))))
             (push range table-ranges)
             (let ((row-step (getf elt :row-step))
                   (col-step (getf elt :col-step)))
               (when (or row-step col-step)
                 (let ((bs (list :range range)))
                   (when row-step
                     (setf (getf bs :row-step) row-step
                           (getf bs :skip-first-row) (or (getf elt :skip-first-row) 0)))
                   (when col-step
                     (setf (getf bs :col-step) col-step
                           (getf bs :skip-first-col) (or (getf elt :skip-first-col) 0)))
                   (push bs table-block-sizes))))))
          (:merge
           (push (format nil "~a~a:~a~a"
                          (column-letter (getf elt :col-start))
                          (getf elt :row-start)
                          (column-letter (getf elt :col-end))
                          (getf elt :row-end))
                  merge-ranges))
          (:col-width
           (push (getf elt :col) col-width-pairs)
           (push (getf elt :width) col-width-pairs)))))
    (values (nreverse table-ranges)
            (nreverse merge-ranges)
            (nreverse col-width-pairs)
            (nreverse table-block-sizes))))

(defun apply-dsl-layout-to-region (region)
  (let ((layout (dsl-layout region)))
    (when layout
      (multiple-value-bind (tr mr cw tbs)
          (compile-dsl-layout layout)
        (unless (table-ranges region) (setf (slot-value region 'table-ranges) tr))
        (unless (merge-ranges region) (setf (slot-value region 'merge-ranges) mr))
        (unless (column-widths region)
          (setf (slot-value region 'column-widths)
                (make-instance 'clase-xl-column-widths :pairs cw)))
        (unless (table-block-sizes region)
          (setf (slot-value region 'table-block-sizes) tbs))))))

(defmethod generate-code ((region clase-xl-region) (lang xl-out) (stream t))
  ;; Compilar layout abstracto
  (apply-dsl-layout-to-region region)

  (let ((compiled-formulas (compiled-formulas region))
        (compiled-cond-rules (compiled-cond-rules region)))
    (when (or (dsl-formula-defs region) (dsl-cond-format-defs region))
      (multiple-value-bind (f c)
          (compile-dsl-formulas-for-region region)
        (setf compiled-formulas f
              compiled-cond-rules c)))
    (format stream "{")
    (dolist (tbl (tables region)) (generate-code tbl lang stream))
    (let ((frms (or compiled-formulas (formulas region))))
      (when frms
        (if (typep frms 'clase-xl-formula-list)
            (generate-code frms lang stream)
            (progn
              (format stream "            \"formulas\": [")
              (loop for i from 0 for f in frms do
                   (when (> i 0) (format stream ", "))
                   (generate-code f lang stream))
              (format stream "],~%")))))
    (let ((ffrms (fernando-formulas region)))
      (when ffrms
        (format stream "            \"fernando_formulas\": [")
        (loop for i from 0 for f in ffrms do
             (when (> i 0) (format stream ", "))
             (generate-code f lang stream))
        (format stream "],~%")))
    (let ((tbs (table-block-sizes region)))
      (when tbs
        (format stream "            \"table_block_sizes\": [")
        (loop for i from 0 for bs in tbs do
             (when (> i 0) (format stream ", "))
             (format stream "{")
             (let ((first t))
               (labels ((emit (key kw)
                          (let ((v (getf bs kw)))
                            (when v
                              (unless first (format stream ", "))
                              (setf first nil)
                              (format stream "~s: " key)
                              (cond ((stringp v) (format stream "~s" v))
                                    ((numberp v) (format stream "~a" v))
                                    (t (format stream "~s" v)))))))
                 (emit "range" :range)
                 (emit "row_step" :row-step)
                 (emit "col_step" :col-step)
                 (emit "skip_first_row" :skip-first-row)
                 (emit "skip_first_col" :skip-first-col)))
             (format stream "}"))
        (format stream "],~%")))
    (let ((tr (table-ranges region))) (when tr (format stream "            \"table_ranges\": ") (xl-write tr stream) (format stream ",~%")))
    (let ((rs (range-styles region)))
      (when rs
        (format stream "            \"range_styles\": [")
        (loop for i from 0 for r in rs do
             (when (> i 0) (format stream ", "))
             (generate-code r lang stream))
        (format stream "],~%")))
    (let ((mr (merge-ranges region)))
      (when mr
        (format stream "            \"merge_ranges\": ")
        (if (and (listp mr) (typep (first mr) 'clase-xl-merge-range))
            (xl-write (mapcar #'range mr) stream)
            (xl-write mr stream))
        (format stream ",~%")))
    (let ((cr compiled-cond-rules))
      (when cr
        (format stream "            \"conditional_format_rules\": [")
        (loop for i from 0 for r in cr do
             (when (> i 0) (format stream ", "))
             (generate-code r lang stream))
        (format stream "],~%")))
    (let ((cw (column-widths region)))
      (when cw
        (if (typep cw 'clase-xl-column-widths)
            (generate-code cw lang stream)
            (progn
              (format stream "            \"column_widths\": ")
              (xl-write-dict cw stream)
              (format stream ",~%")))))
    (format stream "}")))

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

(defmethod generate-code ((wb clase-xl-workbook) (lang xl-out) (stream t))
  (format stream "#!/usr/bin/env python3~%")
  (format stream "from hoja_con_formulas import generar_excel_personalizado~2%")
  (format stream "config = {~%")
  (format stream "    \"sheets\": [~%")
  (when (sheets wb) (loop for i from 0 for sh in (sheets wb) do (generate-code sh lang stream) (when (< i (1- (length (sheets wb)))) (format stream ",~%"))))
  (format stream "    ]~%")
  (format stream "}~2%")
  (format stream "generar_excel_personalizado(config, ~s)~%" (name wb))
  (format stream "~%if __name__=='__main__':~%")
  (format stream "    print('OK: ~a')~%" (name wb)))

(defun xl-generate (wb file)
  (with-open-file (s file :direction :output :if-exists :supersede) (generate-code wb xl-py s))
  (format t "Generado: ~a~%" file))

(defun xl-run-generated (python-file)
  "Ejecuta el archivo Python generado (específico del backend Excel)."
  (format t "Ejecutando python3 ~a...~%" python-file)
  (sb-ext:run-program "/bin/sh" (list "-c" (format nil "python3 ~a" python-file))
                      :output *standard-output*
                      :error *error-output*))

(defun xl-header-style (&key bold color bg-color align)
  (make-instance 'clase-xl-header-style :bold bold :color color :bg-color bg-color :align align))
