; generate-code-direct.lisp — Generate-code methods para el DSL directo
; Backend Excel: serializa AST a Python (openpyxl)
; Cargar en lugar de modelo-excel.lisp para el pipeline directo.
(load "ast-def.lisp")

; =====================================================================
; UTILITIES
; =====================================================================

(defun xl-write (val s)
  (cond ((null val) (format s "None"))
        ((stringp val) (format s "~s" val))
        ((numberp val) (format s "~a" val))
        ((listp val) (progn (format s "[") (loop for i from 0 for x in val do (when (> i 0) (format s ", ")) (xl-write x s)) (format s "]")))
        (t (format s "~s" val))))

(defun escape-python-string (s)
  (with-output-to-string (out)
    (loop for c across s
          do (if (char= c #\") (format out "\\\"") (write-char c out)))))

(defun col->letter (n)
  (with-output-to-string (s)
    (loop while (> n 0)
          do (multiple-value-bind (q r) (floor (1- n) 26)
               (setf n q)
               (write-char (code-char (+ 65 r)) s)))))

(defvar *param-cells* nil
  "Alist of (param-name . cell-ref) for compile-excel-formula")

(defun build-col-map (tbl)
  (loop for name in (col-names tbl) for i from 1
        collect (cons name (col->letter i))))

(defun data-col-count (tbl)
  (let ((computed-names (mapcar #'car (computed tbl))))
    (count-if-not (lambda (n) (member n computed-names)) (col-names tbl))))

(defun data-col-names (tbl)
  (let ((computed-names (mapcar #'car (computed tbl))))
    (remove-if (lambda (n) (member n computed-names)) (col-names tbl))))

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

; =====================================================================
; HELPER — Colores por tipo-programa
; =====================================================================

(defun color-for-tipo (tipo-val)
  (let ((s (and (stringp tipo-val) (string-upcase tipo-val))))
    (cond
      ((null s)                              "#D9D9D9")
      ((search "INFORM" s :test #'char=)     "#E6B8AF")
      ((search "MUSIC" s :test #'char=)      "#A9D18E")
      ((search "CULTUR" s :test #'char=)     "#B4C7E7")
      ((search "DEPOR" s :test #'char=)      "#FFD966")
      ((search "ENTRET" s :test #'char=)     "#D5A6BD")
      ((search "EDUCA" s :test #'char=)      "#C5E0B4")
      (t                                     "#D9D9D9"))))

; =====================================================================
; HELPER — Range styles por tipo
; =====================================================================

(defun emit-range-styles-from-tipo (tbl first-row last-row stream)
  (let* ((dnames (data-col-names tbl))
         (tipo-idx (position 'tipo dnames :test #'string-equal))
         (raw-data (contenido tbl))
         current-color current-start
         (style-count 0))
    (unless tipo-idx (return-from emit-range-styles-from-tipo nil))
    (let ((stylable-cols
           (loop for col-name in '("hora-inicio" "hora-terminacion" "tipo-calc")
                 for pos = (position col-name (col-names tbl) :test #'string-equal)
                 when pos collect (1+ pos))))
      (unless stylable-cols (return-from emit-range-styles-from-tipo nil))
      (format stream "[")
      (loop for i from 0 below (length raw-data)
            for row-num = (+ first-row i)
            for tipo-val = (let ((row (nth i raw-data)))
                            (if (and (listp row) (< tipo-idx (length row)))
                                (nth tipo-idx row)
                                nil))
            for color = (color-for-tipo tipo-val)
            do
               (cond
                 ((null current-color)
                  (setf current-color color current-start row-num))
                 ((string/= color current-color)
                  (loop for col-idx in stylable-cols
                        for col-letter = (col->letter col-idx)
                        do (when (> style-count 0) (format stream ", "))
                           (format stream "{\"range\": ~s, \"style\": {\"bg_color\": ~s}}"
                                   (format nil "~a~a:~a~a" col-letter current-start col-letter (1- row-num))
                                   current-color)
                           (incf style-count))
                  (setf current-color color current-start row-num))))
      (when current-color
        (loop for col-idx in stylable-cols
              for col-letter = (col->letter col-idx)
              do (when (> style-count 0) (format stream ", "))
                 (format stream "{\"range\": ~s, \"style\": {\"bg_color\": ~s}}"
                         (format nil "~a~a:~a~a" col-letter current-start col-letter last-row)
                         current-color)
                 (incf style-count)))
      (format stream "]")
      style-count)))

; =====================================================================
; GENERATE-CODE — TABLAS
; =====================================================================

(defmethod generate-code ((tbl clase-xl-table) (lang xl-out) (stream t))
  (let ((con (contenido tbl)) (hdrs (headers tbl)) (comp (computed tbl))
        (col-map (build-col-map tbl)) (dnames (data-col-names tbl))
        (prms (params tbl)) (cn (col-names tbl))
        (num-cols (length (col-names tbl)))
        (num-params (length (params tbl)))
        (key-count 0))
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
      (when comp
        (let* ((data-rows (length con))
               (first-row (+ 2 (if (> num-params 0) 1 0)))
               (last-row (1- (+ first-row data-rows)))
               (param-cells
                 (when prms
                   (loop for (n . v) in prms for idx from 1
                         for col-num = (+ num-cols idx)
                         collect (cons n (format nil "$~a~a" (col->letter col-num) 2))))))
          (emit-sep)
          (format stream "        \"formulas\": [")
          (let ((formula-count 0))
            (loop for (col . expr) in comp
                  for col-index = (1+ (position col cn :test #'string-equal))
                  do
                     (loop for i from 0 below data-rows
                           for row = (+ first-row i)
                           for formula = (let ((*param-cells* param-cells))
                                           (compile-excel-formula expr col-map dnames row first-row last-row))
                           do
                              (when (> formula-count 0) (format stream ", "))
                              (format stream "{\"row\": ~a, \"col\": ~a, \"value\": \"=~a\"}"
                                      row col-index (escape-python-string formula))
                              (incf formula-count))))
          (format stream "]")
          (incf key-count)))
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
        (format stream "        \"border_color\": \"B7B7B7\"")
        (incf key-count))
      ;; ── range_styles ──
      (when comp
        (let* ((data-rows (length con))
               (first-row (+ 2 (if (> num-params 0) 1 0)))
               (last-row (1- (+ first-row data-rows)))
               (buf (make-string-output-stream))
               (count (emit-range-styles-from-tipo tbl first-row last-row buf)))
          (when (> count 0)
            (emit-sep)
            (format stream "        \"range_styles\": ")
            (format stream (get-output-stream-string buf))
            (incf key-count)))))))

; =====================================================================
; GENERATE-CODE — REGIONES
; =====================================================================

(defmethod generate-code ((region clase-xl-region) (lang xl-out) (stream t))
  (format stream "{")
  (loop for tbl in (tables region) for i from 0
        do (when (> i 0) (format stream ","))
           (generate-code tbl lang stream))
  (format stream "}"))

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

; =====================================================================
; FUNCIONES DE GENERACIÓN
; =====================================================================

(defun xl-generate (wb file)
  (with-open-file (s file :direction :output :if-exists :supersede)
    (generate-code wb xl-py s))
  (format t "Generado: ~a~%" file))

(defun xl-run-generated (python-file)
  (format t "Ejecutando python3 ~a...~%" python-file)
  #+sbcl
  (sb-ext:run-program "/bin/sh" (list "-c" (format nil "python3 ~a" python-file))
                      :output *standard-output* :error *error-output*)
  #+clisp
  (shell (format nil "python3 ~a" python-file)))

(provide "generate-code-direct")
