; modelo-excel.lisp — Generate-code methods y utilidades
; Las clases del AST están en ast-def.lisp
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

; =====================================================================
; GENERATE-CODE METHODS — EXPRESIONES
; Cada clase de expresión sabe serializarse como dict anidado.
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
; GENERATE-CODE METHODS — TABLAS, REGIONES, HOJAS, LIBROS
; =====================================================================

(defmethod generate-code ((tbl clase-xl-table) (lang xl-out) (stream t))
  (let ((con (contenido tbl)) (hdrs (headers tbl)) (comp (computed tbl)))
    (when con (format stream "        \"data\": ") (xl-write con stream) (format stream ",~%"))
    (when hdrs (format stream "        \"headers\": ") (xl-write hdrs stream) (format stream ",~%"))
    (when comp
      (format stream "        \"computed\": [")
      (loop for (col . expr) in comp
            for i from 0
            do (when (> i 0) (format stream ", "))
               (format stream "{\"column\": \"~a\", \"expr\": " col)
               (generate-code expr lang stream)
               (format stream "}"))
      (format stream "],~%"))))

(defmethod generate-code ((region clase-xl-region) (lang xl-out) (stream t))
  (format stream "{")
  (dolist (tbl (tables region)) (generate-code tbl lang stream))
  (format stream "}"))

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
  (format t "Ejecutando python3 ~a...~%" python-file)
  #+sbcl
  (sb-ext:run-program "/bin/sh" (list "-c" (format nil "python3 ~a" python-file))
                      :output *standard-output* :error *error-output*)
  #+clisp
  (shell (format nil "python3 ~a" python-file)))
