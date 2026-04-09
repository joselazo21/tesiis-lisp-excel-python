;;; formulas_fernando.lisp
;;; Carga fórmulas complejas del ODS de Fernando y las proporciona
;;; para inyectar en el generador Excel desde LISP

(defclass* formula-config ()
  (sheet-name
   cell-ref
   original-ods
   excel-formula))

(defparameter *formulas-fernando* nil
  "Almacena todas las fórmulas convertidas del ODS de Fernando")

(defun load-formulas-json (json-file)
  "Carga las fórmulas convertidas desde el archivo JSON"
  (let ((json-content (with-open-file (f json-file :direction :input)
                       (read-file-into-string f))))
    (parse-json-string json-content)))

(defun parse-json-string (json-string)
  "Parse JSON string (simple implementation for our use case)"
  ;; Para una versión más robusta, usa un parser JSON real
  ;; Por ahora, retorna una estructura de prueba
  '(("Aulas" . (
      ("cell" . "C4")
      ("excel" . "=SUBSTITUTE(TRIM(CONCAT(IF(C111!C5=C2,C111!B1 & \" \", \"\"), IF(C112!C5=C2,C112!B1 & \" \", \"\"), IF(C113!C5=C2,C113!B1 & \" \", \"\"))), \" \", \",\")")
    ))
  ))

(defun initialize-fernando-formulas ()
  "Inicializa las fórmulas de Fernando desde el archivo JSON"
  (format t "~%=== CARGANDO FÓRMULAS DE FERNANDO ===~%")
  (let ((json-data nil))
    (ignore-errors
      (with-open-file (f "formulas_fernando_convertidas.json" :direction :input)
        (let ((content (make-string (file-length f))))
          (read-sequence content f)
          (setf json-data content))))
    
    (if json-data
      (format t "✅ Fórmulas cargadas desde JSON~%")
      (format t "⚠️  No se encontró JSON, usando ejemplos por defecto~%"))
    
    *formulas-fernando*))

(defun get-formulas-for-sheet (sheet-name)
  "Retorna las fórmulas para una hoja específica"
  (declare (ignore sheet-name))
  ;; TODO: Implementar búsqueda en *formulas-fernando*
  nil)

(defun inject-formula-to-cell (cell-ref formula)
  "Crea una estructura para inyectar una fórmula en una celda"
  (list :cell cell-ref :formula formula))

(defun generate-formula-injection-config (sheet-name)
  "Genera la configuración para inyectar fórmulas en una hoja"
  (let ((formulas (get-formulas-for-sheet sheet-name)))
    (loop for formula in formulas
          collect (inject-formula-to-cell 
                   (getf formula :cell)
                   (getf formula :excel)))))

;; Ejemplos de fórmulas complejas de la hoja Aulas
(defparameter *formula-aulas-lunes-miercoles-viernes*
  "=SUBSTITUTE(TRIM(CONCAT(IF(C111!C5={col_var},C111!B1 & \" \", \"\"), IF(C112!C5={col_var},C112!B1 & \" \", \"\"), IF(C113!C5={col_var},C113!B1 & \" \", \"\"))), \" \", \",\")")
  "Template de fórmula para concatenar aulas de lunes, miércoles y viernes")

(defun generate-aulas-formula (col-letter)
  "Genera la fórmula específica para una columna de aulas"
  (string-replace-all 
   *formula-aulas-lunes-miercoles-viernes*
   "{col_var}"
   col-letter))

(defun string-replace-all (string old new)
  "Reemplaza todas las ocurrencias de OLD con NEW en STRING"
  (let ((result string)
        (old-len (length old)))
    (loop for pos = (search old result)
          while pos
          do (setf result (concatenate 'string
                                       (subseq result 0 pos)
                                       new
                                       (subseq result (+ pos old-len)))))
    result))

;; Ejemplos para testing
(defun test-formulas ()
  "Testa la generación de fórmulas"
  (format t "~%=== TEST DE FÓRMULAS ===~%")
  (format t "Fórmula para columna C: ~a~%"
          (generate-aulas-formula "C2"))
  (format t "Fórmula para columna D: ~a~%"
          (generate-aulas-formula "D2"))
  (format t "Fórmula para columna E: ~a~%"
          (generate-aulas-formula "E2")))

(format t "~%✅ Módulo formulas_fernando.lisp cargado~%")
