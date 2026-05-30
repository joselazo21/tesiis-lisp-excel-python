; generar_tv_modelo.lisp - Original working version
(load "codigo-tesis.lisp")
(load "modelo-excel.lisp")

; TV Parsing
(defun tv-safe (value &optional (default ""))
  (if value value default))

(defun tv-dia-display (dia)
  (let ((v (string-downcase (tv-safe dia "dia"))))
    (cond ((string= v "lunes") "Lunes") ((string= v "martes") "Martes")
          ((or (string= v "miercoles") (string= v "miércoles")) "Miércoles")
          ((string= v "jueves") "Jueves") ((string= v "viernes") "Viernes")
          ((or (string= v "sabado") (string= v "sábado")) "Sábado")
          ((string= v "domingo") "Domingo") (t (string-capitalize v)))))

(defun programa-tv-a-fila (p)
  (list (tv-safe (getf p :hora-inicio) "") 
        "" (or (getf p :duracion) 0)
        (tv-safe (getf p :nombre) "")
        (tv-safe (getf p :tipo-programa) "") (tv-safe (getf p :tipo-publico) "")))

; Colors
(defparameter *tv-colores* '(("informativo" . "D9EAD3") ("revista" . "E6B8AF")
    ("musical" . "DDCDE4") ("infantil" . "FFF2CC") ("cine" . "D9EAD3")
    ("cultural" . "CFE2F3") ("entrevista" . "FCE5CD") ("ficción" . "E6B8AF")
    ("salud" . "C6EFCE") ("documental" . "FFF2CC") ("animacion" . "E6B8AF")
    ("deporte" . "C6EFCE") (t . "F2F2F2")))

(defun tv-color-tipo (tipo)
  (or (cdr (assoc (string-downcase (tv-safe tipo "")) *tv-colores* :test #'string=)) "F2F2F2"))

; Create Table (NO formulas here - formulas at sheet level)
(defun crear-tabla (cfg)
  (let* ((prog (getf cfg :programas))
         (data (mapcar #'programa-tv-a-fila prog)))
    (xl-table :contenido data :headers '("Inicio" "Fin" "Duración" "Programa" "Tipo" "Público"))))

; Create Sheet WITH FORMULAS (using region)
(defun crear-hoja (cfg idx)
  (xl-sheet :name (format nil "~02d-~a" idx (tv-dia-display (getf cfg :dia)))
    :regions (list
      (xl-region
        :tables (list (crear-tabla cfg))))))

; Totals
(defun calcular-totales (plan)
  (let ((prog 0) (min 0))
    (dolist (d plan)
      (dolist (p (getf d :programas))
        (incf prog) (incf min (or (getf p :duracion) 0))))
    (values prog min)))

; Resumen (using region)
(defun crear-resumen (plan nombre)
  (multiple-value-bind (prog min) (calcular-totales plan)
    (let ((data (list (list "TOTAL" prog min))))
      (xl-sheet :name "Resumen"
        :regions (list
          (xl-region
            :tables (list (xl-table :contenido data :headers '("" "Programas" "Minutos")))
            :range-styles (list (list :range "A2:C2" :style (list :bold t :bg-color "FFF2CC")))))))))

; Workbook
(defun crear-libro (nombre plan &optional (con-resumen t))
  (let ((hojas (loop for i from 1 for d in plan collect (crear-hoja d i))))
    (when con-resumen (push (crear-resumen plan nombre) hojas))
    (xl-workbook :name (format nil "~a.xlsx" nombre) :sheets hojas)))

; Main
(defparameter *tv-nombre* "Canal Habana")

(defun generar-horario-tv (&key (planificacion nil) (nombre *tv-nombre*) (output "horario_tv.py") (incluir-resumen t))
  (unless planificacion (error "No hay planificacion"))
  (xl-generate (crear-libro nombre planificacion incluir-resumen) output)
  (format t "~%Generado: ~a~%" output))